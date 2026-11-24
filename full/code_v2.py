#!/usr/bin/env python3
"""
Storage arbitrage + ancillary services -- full rolling-horizon backtest.

code_v2.py: the same LP and the same 5-minute rolling loop as code.py, run
faster. code.py is left as it was; this copy reads the Project Friday Model
workbook and writes its own *_v2 output files, so the two never overwrite
each other.
  1. Solver: one HiGHS model per hour-phase is built once and kept; each solve
     only changes costs and bounds (scipy's linprog rebuilt it every time).
     Each solve still starts clean, so its answer depends only on its inputs.
  2. Parallel: the backtest is cut into chunks run on all CPU cores at once.
     See run_backtest_parallel() for why the stitched result is the same as
     one long sequential run, and how that is checked at every seam.

Normally started by the workbook's Run Python button (see README.md). To run
it by double-click, close the workbook first. The line above makes Windows' py
launcher use the "python3" on PATH, which needs the packages in
requirements.txt (py -m pip install -r requirements.txt).

Every 5 minutes, solve a 48-hour linear program over the ACTUAL historical
price series (perfect foresight of that 48-hour window -- there is no separate
forecast here, just the real price data), execute only that first 5-minute
action, advance the clock by 5 minutes, and re-solve. Repeat across the whole
2023-07 .. 2026-06 dataset (315,048 steps).

Three revenue streams are co-optimized in the same LP: energy arbitrage,
ancillary services (Regulation Up/Down and Spinning Reserve, paid the CAISO
day-ahead clearing price for committed capacity), and Resource Adequacy
capacity, which is a fixed monthly payment whose *constraints* -- a declining
SOC floor across each commitment block -- are what show up here.

AS awards are held constant across each clock hour, the way a day-ahead award
works; see RollingLP and run_backtest below.
"""

import os
import sys
import time
import traceback


def _report_crash(exc_type, exc, tb):
    """
    Any uncaught error, from the first import on: print it, save it to
    code_v2_crash_log.txt beside this script, and hold a double-clicked window
    open instead of letting it vanish with the message.
    """
    msg = "".join(traceback.format_exception(exc_type, exc, tb))
    if issubclass(exc_type, ImportError):
        msg += ("\nThis Python is missing a package code_v2.py needs. From this folder run:\n"
                f'  "{sys.executable}" -m pip install -r requirements.txt\n')
    sys.stderr.write(msg)
    try:
        log = os.path.join(os.path.dirname(os.path.abspath(__file__)), "code_v2_crash_log.txt")
        with open(log, "w") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "\n" + msg)
        sys.stderr.write(f"\nSaved to {log}\n")
    except Exception:
        pass
    _pause("\nCrashed -- press Enter to close...")


def _pause(prompt):
    """Hold a double-clicked console open; skip quietly when there is no keyboard."""
    try:
        if sys.stdin and sys.stdin.isatty():
            input(prompt)
    except EOFError:
        pass


sys.excepthook = _report_crash

import numpy as np
import pandas as pd
import openpyxl
from scipy.sparse import lil_matrix, csr_matrix, vstack
# HiGHS itself, as bundled inside scipy (the same solver linprog calls). It is
# scipy's private module: if a scipy upgrade moves it, install the public
# package with `pip install highspy` and change this to `import highspy as hc`.
from scipy.optimize._highspy import _core as hc
import multiprocessing as mp


# ============================================================
# RUN SETTINGS (defined first -- RA_FILE has to exist before we can load
# the physical specs out of it, below)
# ============================================================

# Every file is found relative to this script's own folder, so the set works
# wherever it is downloaded or copied to.
HERE = os.path.dirname(os.path.abspath(__file__))

FILE = os.path.join(HERE, "caiso_energy_and_as_5min.csv")
PRICE_COL = "RT_5min_LMP ($/MWH)"

def newest_workbook():
    """
    For a double-click run: the most recently saved Project Friday workbook in
    this folder, so a new version (1.0.4, 1.0.5, ...) is picked up without
    editing this file. Excel's ~$ lock files are skipped.
    """
    books = [f for f in os.listdir(HERE)
             if f.lower().startswith("project friday model") and f.lower().endswith(".xlsm")
             and not f.startswith("~$")]
    if not books:
        return os.path.join(HERE, "Project Friday Model.xlsm")   # reported as missing when read
    return os.path.join(HERE, max(books, key=lambda f: os.path.getmtime(os.path.join(HERE, f))))


RA_FILE = newest_workbook()
# The workbook's Run Python button passes a snapshot copy of itself (Excel,
# with OneDrive AutoSave on, locks the open file so nothing else can read it)
# and the workbook's own folder for the outputs.
OUT_DIR = None
if len(sys.argv) > 1 and sys.argv[1].lower().endswith((".xlsm", ".xlsx")):
    RA_FILE = os.path.abspath(sys.argv[1])
    if len(sys.argv) > 2:
        OUT_DIR = os.path.abspath(sys.argv[2])
# Project/physical scalars (labels in B, values in C). The tab was renamed from
# "Summary" to "Inputs" in 1.0.5; either name is accepted, newest first.
INPUTS_SHEETS = ("Inputs", "Summary")
RA_SHEET = "RA Discharge strat"   # matched by prefix; an "-Active"/"-Alt" pair
                                  # resolves to Active -- see resolve_sheet()
# Scalar inputs are located by LABEL, not cell address -- see find_input().

# Outputs go in the workbook's folder (beside the workbook itself when run by
# double-click): its Import Python Run button looks for them there.
OUT_DIR = OUT_DIR or os.path.dirname(RA_FILE)
OUT_5MIN = os.path.join(OUT_DIR, "storage_arbitrage_backtest_5min_v2.csv")
# Only what the workbook's "Revenue Data 5 Minute" tab holds (Charts reads it
# for the typical-day dispatch); the tab adds its own hour column. Every other
# 5-minute column is still in the _last_run_backup_v2.pkl backup.
FIVE_MIN_COLUMNS = ["datetime", "price", "charge_market_mwh", "discharge_market_mwh", "soc_mwh", "month"]
OUT_MONTHLY = os.path.join(OUT_DIR, "storage_arbitrage_monthly_summary_v2.csv")
# Record of every workbook input this run read, plus its output totals. The
# workbook's Import Python Run button loads it with the monthly results, and
# the Python Run tab then flags any input changed since.
OUT_RUN_INPUTS = os.path.join(OUT_DIR, "dispatch_run_inputs_v2.csv")

# A full run is ~315,000 LP solves. Validate on a short slice first, then set
# QUICK_TEST = False for the real thing. QUICK_TEST runs on one core.
QUICK_TEST = False
QUICK_TEST_STEPS = 500

# Parallel run (see run_backtest_parallel). Chunks start on a clock hour, so
# both lengths are whole hours; WARMUP is how far before its own start each
# chunk begins, to converge onto the true path before its results count.
WORKERS = max(1, (os.cpu_count() or 2) - 2)
CHUNK_STEPS = 14 * 24 * 12      # 14 days
WARMUP_STEPS = 3 * 24 * 12      # 3 days
SEAM_SOC_TOL = 1e-6             # MWh: two paths count as converged within this

ACTION_EPS = 1e-6   # below this, treat a charge/discharge/reserve as zero


# ============================================================
# STORAGE PARAMETERS -- read from the model workbook (30x Tesla
# Megapack 2 XL), not hardcoded here, so the physical model stays editable
# without touching this script.
# ============================================================

def resolve_sheet(wb, wanted):
    """
    Find a sheet by name, tolerating the suffixes this workbook keeps growing.
    An exact match wins; otherwise any sheet whose name starts with `wanted`,
    preferring one marked "-Active" when an Active/Alt pair exists. Raises
    rather than guessing when the choice is ambiguous.
    """
    if wanted in wb.sheetnames:
        return wb[wanted]
    key = wanted.strip().lower()
    hits = [n for n in wb.sheetnames if n.strip().lower().startswith(key)]
    if not hits:
        raise ValueError(f"No sheet starting with '{wanted}' in {RA_FILE}. "
                         f"Sheets present: {wb.sheetnames}")
    active = [n for n in hits if "active" in n.lower()]
    if len(active) == 1:
        return wb[active[0]]
    if len(hits) == 1:
        return wb[hits[0]]
    raise ValueError(f"'{wanted}' is ambiguous in {RA_FILE}: {hits}. "
                     f"Mark one '-Active' or set RA_SHEET to the exact name.")


def find_input(ws, label, max_col=8):
    """
    Look an input up by its LABEL rather than a fixed cell address.

    The workbook's layout keeps moving -- inputs have shifted column (C -> D ->
    C) and row more than once -- and a hardcoded address silently reads the
    wrong cell or a units string when that happens. So: find the row whose
    label matches, then take the first numeric cell to its right.

    Matching ignores case, surrounding whitespace and a trailing ":".
    """
    target = label.strip().lower().rstrip(":")
    for row in ws.iter_rows(max_col=max_col):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().lower().rstrip(":") == target:
                for probe in row[cell.column:]:
                    if isinstance(probe.value, (int, float)) and not isinstance(probe.value, bool):
                        return float(probe.value)
                raise ValueError(
                    f"'{label}' found at {ws.title}!{cell.coordinate} in {RA_FILE}, but no number "
                    f"to its right. If it's a formula, open the workbook in Excel and save it so "
                    f"the cached result is written.")
    raise ValueError(f"No input labelled '{label}' on sheet '{ws.title}' of {RA_FILE}.")


def load_model_workbook():
    """
    Read-only mode: openpyxl then parses only the tabs actually touched. The
    model's 5-minute revenue tab makes the whole file ~40 MB, and a full load
    took ~45 s and ~3 GB before anything printed.
    """
    try:
        return openpyxl.load_workbook(RA_FILE, data_only=True, read_only=True)
    except PermissionError:
        raise PermissionError(
            f"Can't read {RA_FILE}: it is open in Excel, which (with OneDrive AutoSave on) "
            f"locks it. Start the run with the workbook's Run Python button instead, or "
            f"close the workbook and double-click this script again.") from None


def load_system_specs():
    """
    Physical/operating inputs, looked up by label across three tabs:

      "Summary Tab"         Nameplate Power (MW), Nameplate Storage (MWh),
                            RTE -- converted to one-way as sqrt(RTE), since the
                            LP applies efficiency separately on charge and on
                            discharge -- and RA Capacity Revenue Growth (%/yr)
      "RA Discharge strat"  Arbitrage Floor, Arbitrage Ceiling, RA Sustain Hours
                            (these sit beside the RA grid until they're linked
                            onto the Summary Tab)
    Starting SOC is NOT read from the workbook -- see INITIAL_SOC_FRAC below.
    """
    print(f"Reading inputs from {RA_FILE}...")
    wb = load_model_workbook()
    names = [n for n in INPUTS_SHEETS if n in wb.sheetnames]
    if not names:
        raise ValueError(f"No inputs tab ({' or '.join(INPUTS_SHEETS)}) in {RA_FILE}. "
                         f"Sheets present: {wb.sheetnames}")
    summary, ra = wb[names[0]], resolve_sheet(wb, RA_SHEET)
    nameplate_mw = find_input(summary, "Nameplate Power")
    nameplate_mwh = find_input(summary, "Nameplate Storage")
    one_way_eff = find_input(summary, "RTE") ** 0.5
    ra_growth_per_year = find_input(summary, "RA Capacity Revenue Growth")
    arb_floor_frac = find_input(ra, "Arbitrage Floor")
    arb_ceiling_frac = find_input(ra, "Arbitrage Ceiling")
    return (nameplate_mw, nameplate_mwh, one_way_eff, arb_floor_frac,
            arb_ceiling_frac, ra_growth_per_year)


# Parallel workers are fresh Python processes that re-run this file's top
# level. They take the specs the main process already read, passed down in an
# environment variable, rather than re-reading the workbook -- which would be
# 20-odd reads, and could pick up an edit saved mid-run.
_SPECS_ENV = "BESS_V2_SPECS"
if os.environ.get(_SPECS_ENV):
    (NAMEPLATE_MW, NAMEPLATE_MWH, ONE_WAY_EFF, ARB_FLOOR_FRAC,
     ARB_CEILING_FRAC, RA_GROWTH_PER_YEAR) = map(float, os.environ[_SPECS_ENV].split(","))
else:
    (NAMEPLATE_MW, NAMEPLATE_MWH, ONE_WAY_EFF, ARB_FLOOR_FRAC,
     ARB_CEILING_FRAC, RA_GROWTH_PER_YEAR) = load_system_specs()

# Hardcoded on purpose: the workbook's "Inputs" tab was retired (nothing else
# referenced it), and this is a run-start condition rather than a project
# assumption -- it only affects the first few hours of a 3-year backtest.
INITIAL_SOC_FRAC = 0.20

CHARGE_EFF = ONE_WAY_EFF
DISCHARGE_EFF = ONE_WAY_EFF

CHARGE_LIMIT_MW = NAMEPLATE_MW
DISCHARGE_LIMIT_MW = NAMEPLATE_MW

# Two floors stack on SOC (and, mirrored, two ceilings): a flat operating
# limit that always applies (ARB_FLOOR/CEILING, from the 10%/90% arbitrage
# band), plus the AS/RA reserve-energy constraints on top, which can push the
# EFFECTIVE floor/ceiling tighter when active. Multiple "<=" rows on the same
# LP variable just stack automatically -- whichever is tightest wins -- so no
# explicit max()/min() logic is needed; see RollingLP below.
ARB_FLOOR_MWH = NAMEPLATE_MWH * ARB_FLOOR_FRAC
ARB_CEILING_MWH = NAMEPLATE_MWH * ARB_CEILING_FRAC

# State of charge is always expressed against true nameplate, not the
# arbitrage-restricted band.
INITIAL_SOC = NAMEPLATE_MWH * INITIAL_SOC_FRAC

INTERVAL_HOURS = 5 / 60
HORIZON = 48 * 12   # 576 five-minute intervals = 48 hours

# How long a reserve commitment must be sustainable if called, for sizing the
# SOC headroom needed to actually back it. CAISO's real requirements vary by
# product; 30 minutes is a common conservative assumption for spin/regulation
# -- not a real spec, tune this if you have the actual requirement.
RESERVE_DURATION_HOURS = 0.5


# ============================================================
# LOAD DATA
# ============================================================

def load_prices():
    """
    One file, five columns: RT/DA energy price plus RegUp/RegDown/Spin, all
    already aligned to the same 5-minute grid (energy is native 5-min; AS is
    hourly, repeated across that hour's twelve 5-min rows -- both built when
    caiso_energy_and_as_5min.csv was assembled).

    A handful of hours are genuinely missing (DST spring-forward days, where
    that wall-clock hour never happened) -- bridged here with a short
    forward/back-fill capped at one hour (12 rows), so it patches that one
    unavoidable gap without silently papering over anything bigger. If a
    real gap ever exceeds that, these rows are left as NaN on purpose --
    linprog will error out immediately rather than quietly solving on a
    made-up price, which is the point.
    """
    df = pd.read_csv(FILE)
    df["datetime"] = pd.to_datetime(df["datetime"])
    df = df.sort_values("datetime").reset_index(drop=True)

    fill_cols = [PRICE_COL, "RegUp_Price", "RegDown_Price", "Spin_Price"]
    df[fill_cols] = df[fill_cols].ffill(limit=12).bfill(limit=12)

    return (
        df["datetime"].values,
        df[PRICE_COL].to_numpy(),
        df["RegUp_Price"].to_numpy(),
        df["RegDown_Price"].to_numpy(),
        df["Spin_Price"].to_numpy(),
    )


def find_hour_rows(ws, max_scan=80):
    """
    Locate the row holding each hour 0-23 in the RA grid by reading column A,
    instead of assuming a fixed first row. The grid has already shifted once
    (a units header row was inserted above it), and a hardcoded offset reads
    the wrong hour silently -- every RA block would land an hour early.

    Column A holds Excel times (00:00, 01:00, ...); plain integers 0-23 are
    accepted too, so the sheet can be typed either way.
    """
    rows = {}
    for r, (v,) in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, max_col=1, values_only=True), start=1):
        hour = None
        if hasattr(v, "hour") and not hasattr(v, "year"):        # datetime.time
            hour = v.hour
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            if 0 <= v <= 23 and float(v).is_integer():
                hour = int(v)
        if hour is not None and hour not in rows:
            rows[hour] = r
    missing = [h for h in range(24) if h not in rows]
    if missing:
        raise ValueError(f"'{ws.title}' in {RA_FILE}: no row found for hour(s) {missing} "
                         f"in column A -- the RA grid layout has changed.")
    return rows


def load_ra_schedule():
    """
    Hourly x monthly RA quantity (MW) and price ($/kW-mo), read directly from
    the "RA Discharge strat" sheet -- both capacity AND whether RA is active
    now vary by hour-of-day and month-of-year, not a single flat window plus
    one constant MW value.

    Returns:
      qty_by_hour_month:  {(hour, month): MW}
      price_by_month:     {month: $/kW-mo}, averaged across that month's active hours
      sustain_hours:      float -- the one RA parameter that's still flat, Sheet1!C24
      price_by_hour_month, sheet name: kept only for the run record
    """
    wb = load_model_workbook()
    ws = resolve_sheet(wb, RA_SHEET)
    hour_rows = find_hour_rows(ws)
    # read-only sheets have no random cell access, so take the grid rows once
    last_row = max(hour_rows.values())
    grid = list(ws.iter_rows(min_row=1, max_row=last_row, max_col=25, values_only=True))

    qty_by_hour_month = {}
    price_by_hour_month = {}
    price_samples = {m: [] for m in range(1, 13)}
    for h in range(24):
        row = hour_rows[h]
        for m in range(1, 13):
            qty_col = 2 + (m - 1) * 2
            price_col = qty_col + 1
            line = grid[row - 1]
            qty = (line[qty_col - 1] if qty_col <= len(line) else None) or 0.0
            price = (line[price_col - 1] if price_col <= len(line) else None) or 0.0
            qty_by_hour_month[(h, m)] = qty
            price_by_hour_month[(h, m)] = price
            if qty > 0:
                price_samples[m].append(price)
    price_by_month = {m: (sum(v) / len(v) if v else 0.0) for m, v in price_samples.items()}

    # A month can have more than one distinct commitment (e.g. a morning block
    # AND an evening block, at different MW/price) -- identify each as its own
    # contiguous run of matching, non-zero hours so revenue can be summed per
    # block rather than averaged across blocks (averaging would silently blend
    # two different commitments into one wrong number).
    ra_blocks_by_month = {m: [] for m in range(1, 13)}
    for m in range(1, 13):
        h = 0
        while h < 24:
            qty = qty_by_hour_month[(h, m)]
            if qty > 0:
                price = price_by_hour_month[(h, m)]
                run_end = h
                while (run_end + 1 < 24 and qty_by_hour_month[(run_end + 1, m)] == qty
                       and price_by_hour_month[(run_end + 1, m)] == price):
                    run_end += 1
                ra_blocks_by_month[m].append({"start_hour": h, "end_hour": run_end + 1,
                                               "mw": qty, "price_per_kw_mo": price})
                h = run_end + 1
            else:
                h += 1

    sustain_hours = find_input(ws, "RA Sustain Hours")

    return (qty_by_hour_month, price_by_month, ra_blocks_by_month, sustain_hours,
            price_by_hour_month, ws.title)


def compute_ra_arrays(datetimes, qty_by_hour_month):
    """
    Per-5-min-step RA capacity (MW), built from the hourly x monthly grid.
    "Active" is just capacity > 0 -- no separate window parameter needed
    anymore, since the grid itself defines which hours/months are committed.
    """
    idx = pd.DatetimeIndex(datetimes)
    capacity = np.array([qty_by_hour_month.get((h, m), 0.0) for h, m in zip(idx.hour, idx.month)])
    return capacity


def compute_ra_floor(ra_capacity_mw, sustain_hours):
    """
    Per-5-min-step minimum SOC (MWh) needed to back the RA commitment.

    This is a DECLINING floor, not a flat one. The obligation is to deliver
    ra_capacity MW for the remainder of the commitment window (capped at the
    product's sustain duration), so an hour into a 4-hour block you still owe
    3 hours of energy, not 4:

        floor[t] = ra_capacity[t] * min(hours left in this block, sustain_hours)

    A flat `capacity * sustain_hours` floor -- what this model used before --
    demands the full 4 hours of energy at every instant including the last,
    so the battery exits the window still holding four hours it never used.
    That left ~2.6 MWh of play inside a 4-hour window (about one 5-minute
    interval of discharge) and locked the battery out of the evening peak,
    which is where the ~25% arbitrage drop came from.

    The floor is also self-policing: discharging faster than the committed MW
    runs SOC below the next step's requirement, so the LP naturally paces
    discharge at roughly the committed rate across the window -- which is
    exactly the obligation that was sold.
    """
    n = len(ra_capacity_mw)
    steps_remaining = np.zeros(n)
    run = 0
    for i in range(n - 1, -1, -1):
        if ra_capacity_mw[i] <= 0:
            run = 0                                        # not committed this step
        elif i == n - 1 or ra_capacity_mw[i + 1] != ra_capacity_mw[i]:
            run = 1                                        # last step of its block
        else:
            run += 1                                       # mid-block: one more than the step after
        steps_remaining[i] = run

    hours_remaining = np.minimum(steps_remaining * INTERVAL_HOURS, sustain_hours)
    return ra_capacity_mw * hours_remaining


# ============================================================
# OPTIMIZER
#
# Every variable except SOC is a power (MW), not an energy -- this is what
# lets energy dispatch and AS capacity commitments share the same headroom
# constraints:
#   store_charge[0:n]     MW physically charging the battery
#   store_discharge[0:n]  MW physically discharging the battery
#   reg_up[0:n]            MW committed to Regulation Up
#   reg_down[0:n]          MW committed to Regulation Down
#   spin[0:n]              MW committed to Spinning Reserve
#   soc[0:n]                MWh storage state after interval t
#
# Storage balance (plain energy conservation -- efficiency only applies when
# converting to/from the market):
#   soc[t] = soc[t-1] + store_charge[t]*dt - store_discharge[t]*dt
#
# Power headroom -- if everything committed were called simultaneously, the
# battery still has to stay inside its physical rating:
#   store_discharge[t] + reg_up[t] + spin[t]  <=  DISCHARGE_LIMIT_MW
#   store_charge[t] + reg_down[t]              <=  CHARGE_LIMIT_MW
#
# NOTE: ra_capacity[t] is deliberately NOT subtracted from discharge headroom.
# RA is not a reserve product like Spin -- you don't withhold the MW, you carry
# a must-offer obligation and get dispatched. Being dispatched at your RA
# quantity IS delivering RA, so the same MW cannot be double-counted as
# "reserved" and thereby blocked from earning energy revenue. Reserving it
# modelled the obligation backwards: it paid the RA capacity payment while
# denying the peak energy revenue a real RA resource collects on top. RA's
# real operational bite is the ENERGY (SOC) requirement below, not power.
#
# Energy headroom -- there has to be enough stored energy (or empty capacity)
# to actually deliver a commitment if called:
#   (reg_up[t] + spin[t]) * RESERVE_DURATION_HOURS  <=  soc[t]
#   reg_down[t] * RESERVE_DURATION_HOURS             <=  ARB_CEILING_MWH - soc[t]
#   ra_floor[t]                                       <=  soc[t]
# where ra_floor[t] DECLINES across each commitment block as the obligation is
# worked off (see compute_ra_floor) rather than sitting flat at 4h * MW for the
# whole window. ra_capacity varies by hour-of-day AND month, read straight from
# the "RA Discharge strat" sheet -- a fixed commitment, not something the
# optimizer chooses. Zero wherever the sheet has no RA commitment that hour.
#
# Two floors, one ceiling, stacked (whichever is tighter wins automatically):
#   - The "general arb floor" row enforces soc[t] >= ARB_FLOOR_MWH at ALL
#     times, RA-active or not. The 10% floor is a cycle-life policy, and the
#     model runs its RA windows every single day -- draining to true empty
#     daily is not something a real operator does. (It also keeps the floor
#     continuous: with a declining RA floor, releasing the 10% line would let
#     SOC hit ~0 on the window's last step and then demand 11.748 MWh on the
#     next one, which is physically unreachable in 5 minutes -- an infeasible
#     LP. Set "Arbitrage soc floor" to 0 in the workbook if you ever do want
#     true-empty RA drawdown.)
#   - The RA floor row stacks on top, binding only while it exceeds the 10%
#     line -- i.e. effective floor = max(ra_floor[t], ARB_FLOOR_MWH).
#   - The ceiling stays at ARB_CEILING_MWH (90%) unconditionally -- RA is a
#     discharge/floor product, there's no equivalent reason to relax the top.
#
# Objective: maximize energy arbitrage revenue (as in r2) plus AS capacity
# payments (price[$/MW-h] * MW * dt = $ for the interval); no efficiency loss
# on capacity payments since you're paid to reserve capacity, not to deliver
# energy. RA's own capacity payment isn't in this per-interval objective at
# all -- it's a flat monthly payment added in build_monthly_summary(), since
# it doesn't depend on dispatch. Only RA's *constraints* (above) show up here.
#
# The constraint matrix's sparsity structure never changes between solves --
# only the price vectors, starting SOC, and the RA capacity for this
# particular 48h window do -- so the matrix is built once and reused, with
# just the RA-affected right-hand sides recomputed per solve.
# ============================================================

class RollingLP:
    """
    One prebuilt LP for a 48-hour horizon.

    `phase` is where the horizon's first 5-minute step sits inside its clock
    hour (0 = on the hour, 11 = the last 5-minute slot). It only affects where
    the hour boundaries fall, so build_lps() builds the 12 possible variants
    once and reuses them -- the constraint structure never changes mid-run.
    """

    STEPS_PER_HOUR = 12

    def __init__(self, n, phase=0):
        self.n = n
        self.phase = phase
        self.charge_idx = np.arange(0 * n, 1 * n)
        self.discharge_idx = np.arange(1 * n, 2 * n)
        self.regup_idx = np.arange(2 * n, 3 * n)
        self.regdown_idx = np.arange(3 * n, 4 * n)
        self.spin_idx = np.arange(4 * n, 5 * n)
        self.soc_idx = np.arange(5 * n, 6 * n)
        total_vars = 6 * n
        dt = INTERVAL_HOURS

        # --- Equality: SOC balance, then hourly AS awards ---
        # CAISO awards AS hourly in the day-ahead market; it is not re-chosen
        # every 5 minutes. So within a clock hour each AS product is pinned to
        # one value: var[t] - var[t-1] = 0 for every step that is not the first
        # of its hour. (Energy dispatch stays free at 5-minute resolution --
        # that part really is a 5-minute market.)
        as_blocks = [self.regup_idx, self.regdown_idx, self.spin_idx]
        link_steps = [t for t in range(1, n) if (phase + t) % self.STEPS_PER_HOUR != 0]
        n_link = len(link_steps) * len(as_blocks)

        A_eq = lil_matrix((n + n_link, total_vars))
        for t in range(n):
            A_eq[t, self.soc_idx[t]] = 1
            A_eq[t, self.charge_idx[t]] = -dt
            A_eq[t, self.discharge_idx[t]] = dt
            if t > 0:
                A_eq[t, self.soc_idx[t - 1]] = -1
        r = n
        for idx in as_blocks:
            for t in link_steps:
                A_eq[r, idx[t]] = 1
                A_eq[r, idx[t - 1]] = -1
                r += 1
        self.A_eq = csr_matrix(A_eq)
        self.n_eq = n + n_link

        # --- Inequality: power + energy headroom, 6 rows per step (all "<=") ---
        A_ub = lil_matrix((6 * n, total_vars))
        b_ub_base = np.zeros(6 * n)
        for t in range(n):
            r = 6 * t
            # discharge power headroom (RA is NOT carved out of this -- see above)
            A_ub[r, self.discharge_idx[t]] = 1
            A_ub[r, self.regup_idx[t]] = 1
            A_ub[r, self.spin_idx[t]] = 1
            b_ub_base[r] = DISCHARGE_LIMIT_MW
            # charge power headroom
            A_ub[r + 1, self.charge_idx[t]] = 1
            A_ub[r + 1, self.regdown_idx[t]] = 1
            b_ub_base[r + 1] = CHARGE_LIMIT_MW
            # energy headroom to deliver reg_up + spin if called
            A_ub[r + 2, self.regup_idx[t]] = RESERVE_DURATION_HOURS
            A_ub[r + 2, self.spin_idx[t]] = RESERVE_DURATION_HOURS
            A_ub[r + 2, self.soc_idx[t]] = -1
            b_ub_base[r + 2] = 0
            # energy headroom to absorb reg_down if called (against the 90% ceiling,
            # not full nameplate -- arbitrage/AS never push SOC above ARB_CEILING_MWH)
            A_ub[r + 3, self.regdown_idx[t]] = RESERVE_DURATION_HOURS
            A_ub[r + 3, self.soc_idx[t]] = 1
            b_ub_base[r + 3] = ARB_CEILING_MWH
            # RA floor: -soc[t] <= -ra_floor[t] (i.e. soc[t] >= ra_floor[t]).
            # Declining across each block; 0 (no restriction) off-window.
            # RHS set per-solve from the precomputed floor array.
            A_ub[r + 4, self.soc_idx[t]] = -1
            b_ub_base[r + 4] = 0
            # general arb floor: soc[t] >= ARB_FLOOR_MWH, unconditionally. Stacks
            # with the row above, so the effective floor is whichever is tighter.
            A_ub[r + 5, self.soc_idx[t]] = -1
            b_ub_base[r + 5] = -ARB_FLOOR_MWH
        self.A_ub = csr_matrix(A_ub)
        self.b_ub_base = b_ub_base
        # the only RHS that still changes per solve -- everything else is static now
        self.ra_floor_row = 6 * np.arange(n) + 4    # row r+4 per step

        # bounds as arrays, so a solve can pin the already-committed AS award
        # for the rest of the current hour without rebuilding a 3,456-tuple list
        self.lb = np.zeros(total_vars)
        self.ub = np.concatenate([
            np.full(n, CHARGE_LIMIT_MW),
            np.full(n, DISCHARGE_LIMIT_MW),
            np.full(n, DISCHARGE_LIMIT_MW),   # reg_up capped by discharge-direction rating
            np.full(n, CHARGE_LIMIT_MW),      # reg_down capped by charge-direction rating
            np.full(n, DISCHARGE_LIMIT_MW),   # spin capped by discharge-direction rating
            np.full(n, ARB_CEILING_MWH),      # SOC: true-empty floor, 90% ceiling
        ])

        # steps belonging to the horizon's first (possibly partial) clock hour
        self.steps_left_in_hour = (self.STEPS_PER_HOUR - phase) % self.STEPS_PER_HOUR

        # --- One persistent HiGHS model holding the matrices above ---
        # The two SOC-floor rows per step (RA floor, general arb floor) only
        # ever bound soc[t] from below, so they go in as soc's lower bound,
        # max(ra_floor[t], ARB_FLOOR_MWH) -- the same feasible set, set per
        # solve with one bounds call instead of 1,152 row changes.
        keep = np.array([i for i in range(6 * n) if i % 6 not in (4, 5)])
        A = vstack([self.A_eq, self.A_ub[keep]]).tocsc()
        lp = hc.HighsLp()
        lp.num_col_, lp.num_row_ = A.shape[1], A.shape[0]
        lp.col_cost_ = np.zeros(A.shape[1])
        lp.col_lower_, lp.col_upper_ = self.lb, self.ub
        lp.row_lower_ = np.concatenate([np.zeros(self.n_eq), np.full(len(keep), -hc.kHighsInf)])
        lp.row_upper_ = np.concatenate([np.zeros(self.n_eq), self.b_ub_base[keep]])
        lp.a_matrix_.format_ = hc.MatrixFormat.kColwise
        lp.a_matrix_.start_, lp.a_matrix_.index_, lp.a_matrix_.value_ = A.indptr, A.indices, A.data
        self.highs = hc._Highs()
        self.highs.setOptionValue("output_flag", False)
        self.highs.setOptionValue("threads", 1)     # parallelism comes from the chunks
        self.highs.passModel(lp)
        self.all_cols = np.arange(total_vars, dtype=np.int32)

    def solve(self, price_forecast, reg_up_price, reg_down_price, spin_price,
              initial_soc, ra_floor_mwh, committed_as=None):
        """
        committed_as: (reg_up, reg_down, spin) MW already awarded for the clock
        hour this horizon starts part-way through, or None on the hour. Those
        steps get pinned to the award instead of being re-chosen -- you cannot
        walk away from a day-ahead award mid-hour.
        """
        n, dt = self.n, INTERVAL_HOURS
        c = np.zeros(6 * n)
        c[self.charge_idx] = (price_forecast / CHARGE_EFF) * dt
        c[self.discharge_idx] = -(price_forecast * DISCHARGE_EFF) * dt
        c[self.regup_idx] = -reg_up_price * dt
        c[self.regdown_idx] = -reg_down_price * dt
        c[self.spin_idx] = -spin_price * dt

        lb, ub = self.lb.copy(), self.ub.copy()
        lb[self.soc_idx] = np.maximum(ra_floor_mwh, ARB_FLOOR_MWH)   # both floor rows, as a bound
        if committed_as is not None and self.steps_left_in_hour:
            k = self.steps_left_in_hour
            for idx, mw in zip((self.regup_idx, self.regdown_idx, self.spin_idx), committed_as):
                lb[idx[:k]] = mw
                ub[idx[:k]] = mw

        h = self.highs
        # Start every solve clean. Keeping the last basis would save little here,
        # and it would let a tie between equally good answers (often RegUp vs
        # Spin at the same price) depend on what was solved before -- which the
        # parallel seam check relies on never happening.
        h.clearSolver()
        h.changeColsCost(len(c), self.all_cols, c)
        h.changeColsBounds(len(c), self.all_cols, lb, ub)
        h.changeRowBounds(0, initial_soc, initial_soc)    # SOC balance row 0 carries the start SOC
        h.run()
        if h.getModelStatus() != hc.HighsModelStatus.kOptimal:
            raise RuntimeError("HiGHS: " + h.modelStatusToString(h.getModelStatus()))
        x = np.asarray(h.getSolution().col_value)

        return (
            x[self.charge_idx],
            x[self.discharge_idx],
            x[self.regup_idx],
            x[self.regdown_idx],
            x[self.spin_idx],
        )


# ============================================================
# ROLLING BACKTEST LOOP
# ============================================================

RECORD_COLUMNS = [
    "datetime", "price", "action",
    "charge_market_mwh", "discharge_market_mwh",
    "storage_charge_mwh", "storage_discharge_mwh",
    "soc_mwh",
    "reg_up_mw", "reg_down_mw", "spin_mw",
    "ra_capacity_mw", "ra_active", "ra_floor_mwh",
    "energy_revenue", "reg_up_revenue", "reg_down_revenue", "spin_revenue",
    "cashflow",
]


def build_lps():
    # One prebuilt LP per position-within-the-hour; only the hour boundaries
    # differ, so 12 variants cover every step of the run.
    return [RollingLP(HORIZON, phase) for phase in range(RollingLP.STEPS_PER_HOUR)]


def simulate(data, lps, t_from, t_to, soc, committed_as=None, progress=None):
    """
    Run the 5-minute loop over steps t_from .. t_to-1, starting from `soc` and
    the hour's `committed_as` award (None on the hour). Returns one record per
    step and soc_path: the SOC at the start of every step plus the end SOC
    (t_to - t_from + 1 values) -- the state the seam check compares.
    """
    datetimes, prices, reg_up_price, reg_down_price, spin_price, ra_capacity_full, ra_floor_full, phases = data
    dt = INTERVAL_HOURS
    records = []
    soc_path = np.empty(t_to - t_from + 1)

    for i, t in enumerate(range(t_from, t_to)):
        soc_path[i] = soc
        price_fc = prices[t:t + HORIZON]
        ru_fc = reg_up_price[t:t + HORIZON]
        rd_fc = reg_down_price[t:t + HORIZON]
        sp_fc = spin_price[t:t + HORIZON]
        ra_floor_fc = ra_floor_full[t:t + HORIZON]

        phase = phases[t]
        if phase == 0:
            committed_as = None        # new clock hour -- the award is up for grabs
        charge_mw, discharge_mw, regup_mw, regdown_mw, spin_mw = lps[phase].solve(
            price_fc, ru_fc, rd_fc, sp_fc, soc, ra_floor_fc, committed_as)

        c0, d0, ru0, rd0, sp0 = charge_mw[0], discharge_mw[0], regup_mw[0], regdown_mw[0], spin_mw[0]
        c0 = 0.0 if c0 < ACTION_EPS else c0
        d0 = 0.0 if d0 < ACTION_EPS else d0
        ru0 = 0.0 if ru0 < ACTION_EPS else ru0
        rd0 = 0.0 if rd0 < ACTION_EPS else rd0
        sp0 = 0.0 if sp0 < ACTION_EPS else sp0

        if phase == 0:
            committed_as = (ru0, rd0, sp0)   # hold this award for the rest of the hour

        sc0, sd0 = c0 * dt, d0 * dt   # physical MWh moved this interval
        charge_market = sc0 / CHARGE_EFF
        discharge_market = sd0 * DISCHARGE_EFF
        soc = soc + sc0 - sd0

        price_now = prices[t]
        energy_revenue = discharge_market * price_now - charge_market * price_now
        reg_up_revenue = ru0 * reg_up_price[t] * dt
        reg_down_revenue = rd0 * reg_down_price[t] * dt
        spin_revenue = sp0 * spin_price[t] * dt
        cashflow = energy_revenue + reg_up_revenue + reg_down_revenue + spin_revenue

        if c0 > 0:
            action = "CHARGE"
        elif d0 > 0:
            action = "DISCHARGE"
        else:
            action = "HOLD"

        records.append((
            datetimes[t], price_now, action,
            charge_market, discharge_market, sc0, sd0, soc,
            ru0, rd0, sp0,
            ra_capacity_full[t], bool(ra_capacity_full[t] > 0), ra_floor_full[t],
            energy_revenue, reg_up_revenue, reg_down_revenue, spin_revenue,
            cashflow,
        ))
        if progress:
            progress(i + 1)

    soc_path[-1] = soc
    return records, soc_path


def records_to_frame(records):
    out = pd.DataFrame(records, columns=RECORD_COLUMNS)
    out["cumulative_profit"] = out["cashflow"].cumsum()
    out["year"] = out["datetime"].dt.year
    out["month"] = out["datetime"].dt.month
    return out


def backtest_data(datetimes, prices, reg_up_price, reg_down_price, spin_price,
                  ra_capacity_full, ra_floor_full):
    phases = (pd.DatetimeIndex(datetimes).minute.to_numpy() // 5) % RollingLP.STEPS_PER_HOUR
    return (datetimes, prices, reg_up_price, reg_down_price, spin_price,
            ra_capacity_full, ra_floor_full, phases)


def run_backtest(data, n_steps, verbose_every=1000):
    """One core, start to finish: QUICK_TEST, and the reference the parallel run is checked against."""
    n_iters = min(n_steps, len(data[1]) - HORIZON)
    t0 = time.time()

    def progress(done):
        if verbose_every and done % verbose_every == 0:
            elapsed = time.time() - t0
            remaining = (n_iters - done) / (done / elapsed)
            print(f"  {done:,}/{n_iters:,} steps  ({elapsed:,.0f}s elapsed, ~{remaining:,.0f}s left)")

    records, _ = simulate(data, build_lps(), 0, n_iters, INITIAL_SOC, None, progress)
    return records_to_frame(records)


# --- parallel run -----------------------------------------------------------
#
# The loop is sequential only through its state: the SOC carried from one step
# to the next, plus the AS award held within a clock hour (which resets every
# hour). Two runs that differ only in starting SOC converge onto the same path
# once both are at the same SOC at an hour start -- typically within a day,
# when both drain to the floor or fill to the ceiling. From there every solve
# sees identical inputs, so every later step is identical.
#
# So each chunk starts WARMUP_STEPS early from a guessed SOC, runs on its own
# core, and only its steps from its own start count. At every seam the check
# looks, inside the overlap, for an hour start where this chunk's SOC matches
# the previous chunk's (already accepted) SOC. Found: the chunk was on the true
# path before its results start. Not found: it is re-run on one core from the
# previous chunk's true end state, which is exact by construction. Either way
# the stitched result equals one long sequential run.

_worker_data = None
_worker_lps = None


def _worker_init(data):
    global _worker_data, _worker_lps
    _worker_data, _worker_lps = data, build_lps()


def warmup_start_soc(data, t):
    """
    Starting SOC guess for a warm-up beginning at step t. The guess itself
    doesn't matter (the seam check decides what counts), but it has to be one
    the LP can solve from: just before an RA block the floor can climb faster
    than the battery can charge, and from INITIAL_SOC that is infeasible. So
    take the lowest SOC that can still reach every floor in the horizon at
    full charging power, if that is above INITIAL_SOC.
    """
    floor = np.maximum(data[6][t:t + HORIZON], ARB_FLOOR_MWH)
    reach = CHARGE_LIMIT_MW * INTERVAL_HOURS * np.arange(1, len(floor) + 1)
    return float(min(max(INITIAL_SOC, ARB_FLOOR_MWH, (floor - reach).max() + 1e-6), ARB_CEILING_MWH))


def _worker_chunk(job):
    k, t_from, t_to = job
    soc0 = INITIAL_SOC if t_from == 0 else warmup_start_soc(_worker_data, t_from)
    records, soc_path = simulate(_worker_data, _worker_lps, t_from, t_to, soc0)
    return k, t_from, records, soc_path


def run_backtest_parallel(data, n_steps):
    phases = data[7]
    n_iters = min(n_steps, len(data[1]) - HORIZON)
    starts = []
    for s in range(0, n_iters, CHUNK_STEPS):
        while s < n_iters and phases[s] != 0:    # chunks begin on a clock hour
            s += 1
        if s < n_iters and (not starts or s > starts[-1]):
            starts.append(s)
    ends = starts[1:] + [n_iters]
    jobs = []
    for k, s in enumerate(starts):
        w = 0 if k == 0 else max(0, s - WARMUP_STEPS)
        while phases[w] != 0:
            w += 1
        jobs.append((k, min(w, s), ends[k]))

    print(f"  {len(jobs)} chunks of ~{CHUNK_STEPS // 288} days (+{WARMUP_STEPS // 288} days warm-up) "
          f"on {WORKERS} cores")
    t0 = time.time()
    results = {}
    with mp.Pool(WORKERS, initializer=_worker_init, initargs=(data,)) as pool:
        for n_done, (k, t_from, records, soc_path) in enumerate(pool.imap_unordered(_worker_chunk, jobs), 1):
            results[k] = (t_from, records, soc_path)
            elapsed = time.time() - t0
            print(f"  {n_done}/{len(jobs)} chunks  ({elapsed:,.0f}s elapsed, "
                  f"~{elapsed / n_done * (len(jobs) - n_done):,.0f}s left)")

    # Stitch, checking each seam against the chunk before it.
    lps = None
    stitched, reruns = [], 0
    for k, s in enumerate(starts):
        t_from, records, soc_path = results[k]
        if k > 0:
            p_from, _, p_path = results[k - 1]
            ok = False
            for j in range(t_from, s + 1):
                if phases[j] == 0 and abs(soc_path[j - t_from] - p_path[j - p_from]) <= SEAM_SOC_TOL:
                    ok = True
                    break
            if not ok:
                reruns += 1
                print(f"  chunk {k + 1} had not converged by its start; re-running it from the true state")
                lps = lps or build_lps()
                records, soc_path = simulate(data, lps, s, ends[k], p_path[s - p_from])
                t_from = s
                results[k] = (t_from, records, soc_path)
        stitched.extend(records[s - t_from:])
    print(f"  seams checked: {len(starts) - 1 - reruns} converged, {reruns} re-run")
    return records_to_frame(stitched)


def build_monthly_summary(backtest_df, ra_blocks_by_month):
    """
    Monthly export laid out like the "Revenue Data" tab of model 2.0.xlsx, so
    it can be pasted straight across. Columns that tab computes itself are
    left out on purpose: P/L Hold (folded into RTE), SOH and Retrofit CapEx
    (degradation and augmentation are handled in Excel).
    """
    df = backtest_df.copy()
    df["charge_pl"] = -df["charge_market_mwh"] * df["price"]        # cost of energy bought (negative)
    df["discharge_pl"] = df["discharge_market_mwh"] * df["price"]   # revenue from energy sold
    df["is_charge"] = df["action"] == "CHARGE"
    df["is_discharge"] = df["action"] == "DISCHARGE"
    df["is_hold"] = df["action"] == "HOLD"

    s = df.groupby(["year", "month"]).agg(
        pl_total=("energy_revenue", "sum"),
        pl_charge=("charge_pl", "sum"),
        pl_discharge=("discharge_pl", "sum"),
        charge_steps=("is_charge", "sum"),
        discharge_steps=("is_discharge", "sum"),
        hold_steps=("is_hold", "sum"),
        mwh_charged=("charge_market_mwh", "sum"),
        mwh_discharged=("discharge_market_mwh", "sum"),
        reg_up=("reg_up_revenue", "sum"),
        reg_down=("reg_down_revenue", "sum"),
        spin=("spin_revenue", "sum"),
    ).reset_index()

    # RA is a flat monthly capacity payment, independent of dispatch -- not
    # part of the per-interval cashflow, added here instead. A month can have
    # more than one commitment block (morning + evening, at different MW and
    # price); each is billed separately and summed.
    s["ra"] = s["month"].map(
        lambda m: sum(b["mw"] * b["price_per_kw_mo"] * 1000 for b in ra_blocks_by_month.get(m, [])))

    # RA contract rates step up once a year ("RA Capacity Revenue Growth",
    # Summary Tab D77). Year 1 of the backtest pays the rates as entered in the
    # grid; every 12 months after that compounds one more year of growth. Note
    # the final year already carries two years of escalation -- the Excel model
    # should escalate forward from THAT base, not from the year-1 rate.
    contract_year = np.arange(len(s)) // 12
    s["ra"] *= (1.0 + RA_GROWTH_PER_YEAR) ** contract_year

    out = pd.DataFrame({
        "Year": s["year"],
        "Month": s["month"],
        "Model Month": np.arange(1, len(s) + 1),
        "P/L Total ($)": s["pl_total"],
        "P/L Charge ($)": s["pl_charge"],
        "P/L Discharge ($)": s["pl_discharge"],
        "Charge Hours": s["charge_steps"] * INTERVAL_HOURS,
        "Discharge Hours": s["discharge_steps"] * INTERVAL_HOURS,
        "Hold Hours": s["hold_steps"] * INTERVAL_HOURS,
        "MWh Charged": s["mwh_charged"],
        "MWh Discharged": s["mwh_discharged"],
        "AS Revenue - RA ($)": s["ra"],
        "AS Revenue - Reg Up ($)": s["reg_up"],
        "AS Revenue - Reg Down ($)": s["reg_down"],
        "AS Revenue - Spinning ($)": s["spin"],
        "AS Revenue - Non-Spin ($)": 0.0,   # not modelled; kept so the layout matches
    })
    out["Total AS Revenue ($)"] = out[["AS Revenue - RA ($)", "AS Revenue - Reg Up ($)", "AS Revenue - Reg Down ($)",
                                       "AS Revenue - Spinning ($)", "AS Revenue - Non-Spin ($)"]].sum(axis=1)
    out["Grand Total Revenue ($)"] = out["P/L Total ($)"] + out["Total AS Revenue ($)"]
    # market-side $/MWh; blank rather than divide-by-zero in a month with no trades
    out["AVG purchase price"] = (-out["P/L Charge ($)"] / out["MWh Charged"].where(out["MWh Charged"] > 0))
    out["AVG sell price"] = out["P/L Discharge ($)"] / out["MWh Discharged"].where(out["MWh Discharged"] > 0)
    out["AVG spread"] = out["AVG sell price"] - out["AVG purchase price"]
    return out


def build_run_record(ra_sheet_name, qty_by_hour_month, price_by_hour_month,
                     sustain_hours, monthly, n_steps):
    """
    One row per input this run read from the workbook, then the output totals,
    as item,hour,month,value. The workbook's Python Run tab compares these
    against its live inputs and its Revenue Data tab, so editing an input that
    only takes effect through a re-run (size, RTE, the RA grid...) gets flagged
    instead of silently leaving revenue stale.
    """
    import datetime, os
    rows = [
        ("run_finished", "", "", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("workbook", "", "", os.path.basename(RA_FILE)),
        ("ra_sheet", "", "", ra_sheet_name),
        ("price_file", "", "", os.path.basename(FILE)),
        ("quick_test", "", "", int(QUICK_TEST)),
        ("steps", "", "", n_steps),
        ("Nameplate Power", "", "", NAMEPLATE_MW),
        ("Nameplate Storage", "", "", NAMEPLATE_MWH),
        ("RTE", "", "", ONE_WAY_EFF ** 2),
        ("RA Capacity Revenue Growth", "", "", RA_GROWTH_PER_YEAR),
        ("Arbitrage Floor", "", "", ARB_FLOOR_FRAC),
        ("Arbitrage Ceiling", "", "", ARB_CEILING_FRAC),
        ("RA Sustain Hours", "", "", sustain_hours),
    ]
    for h in range(24):
        for m in range(1, 13):
            rows.append(("ra_qty", h, m, qty_by_hour_month[(h, m)]))
            rows.append(("ra_price", h, m, price_by_hour_month[(h, m)]))
    rows += [
        ("out_months", "", "", len(monthly)),
        ("out_pl_total", "", "", monthly["P/L Total ($)"].sum()),
        ("out_mwh_discharged", "", "", monthly["MWh Discharged"].sum()),
        ("out_ra", "", "", monthly["AS Revenue - RA ($)"].sum()),
        ("out_reg_up", "", "", monthly["AS Revenue - Reg Up ($)"].sum()),
        ("out_reg_down", "", "", monthly["AS Revenue - Reg Down ($)"].sum()),
        ("out_spin", "", "", monthly["AS Revenue - Spinning ($)"].sum()),
    ]
    return pd.DataFrame(rows, columns=["item", "hour", "month", "value"])


# ============================================================
# MAIN
# ============================================================

def safe_to_csv(df, path, retries=20, wait_seconds=15):
    """
    A multi-hour backtest is expensive to redo. If the target file is open
    elsewhere (Excel, OneDrive sync, etc.) this waits and retries instead of
    crashing and throwing away everything that was just computed.
    """
    for attempt in range(1, retries + 1):
        try:
            df.to_csv(path, index=False)
            return
        except PermissionError:
            print(f"  '{path}' is open elsewhere (close it -- likely Excel/OneDrive). "
                  f"Retrying in {wait_seconds}s... (attempt {attempt}/{retries})")
            time.sleep(wait_seconds)
    raise PermissionError(f"Still couldn't write '{path}' after {retries} attempts -- "
                           f"the backup pickle in the same folder still has your results.")


def main():
    print("Loading price data...")
    datetimes, prices, reg_up_price, reg_down_price, spin_price = load_prices()
    n_total = len(prices)
    print(f"  {n_total:,} 5-min rows, {datetimes[0]} to {datetimes[-1]}")
    print(f"  RegUp avg ${reg_up_price.mean():.2f}/MW-h, RegDown avg ${reg_down_price.mean():.2f}/MW-h, "
          f"Spin avg ${spin_price.mean():.2f}/MW-h")

    print("Loading RA contract schedule...")
    (qty_by_hour_month, price_by_month, ra_blocks_by_month, ra_sustain_hours,
     price_by_hour_month, ra_sheet_name) = load_ra_schedule()
    ra_capacity_full = compute_ra_arrays(datetimes, qty_by_hour_month)
    ra_floor_full = compute_ra_floor(ra_capacity_full, ra_sustain_hours)
    active_qtys = sorted(set(qty_by_hour_month.values()) - {0.0})
    total_blocks = sum(len(b) for b in ra_blocks_by_month.values())
    print(f"  RA schedule loaded: sustain {ra_sustain_hours}h, {total_blocks} commitment blocks across the year, "
          f"quantities range {min(active_qtys):.2f}-{max(active_qtys):.2f} MW, "
          f"prices range ${min(price_by_month.values()):.2f}-${max(price_by_month.values()):.2f}/kW-mo")

    n_steps = n_total - HORIZON
    if QUICK_TEST:
        n_steps = min(n_steps, QUICK_TEST_STEPS)
        print(f"QUICK_TEST=True -> running {n_steps:,} steps only. "
              f"Set QUICK_TEST=False for the full {n_total - HORIZON:,}-step backtest.")

    print(f"\nRunning rolling backtest ({n_steps:,} steps, 48h horizon, re-solving every 5 min)...")
    data = backtest_data(datetimes, prices, reg_up_price, reg_down_price, spin_price,
                         ra_capacity_full, ra_floor_full)
    run_t0 = time.time()
    if QUICK_TEST or WORKERS == 1:
        backtest_df = run_backtest(data, n_steps)
    else:
        # hand the workers the specs read above (see _SPECS_ENV)
        os.environ[_SPECS_ENV] = ",".join(repr(float(v)) for v in (
            NAMEPLATE_MW, NAMEPLATE_MWH, ONE_WAY_EFF, ARB_FLOOR_FRAC, ARB_CEILING_FRAC, RA_GROWTH_PER_YEAR))
        backtest_df = run_backtest_parallel(data, n_steps)
    print(f"  backtest finished in {(time.time() - run_t0) / 60:,.1f} min")
    monthly = build_monthly_summary(backtest_df, ra_blocks_by_month)

    # Save an unconditional backup FIRST, before risking the (possibly locked)
    # named output files -- a long run should never be lost to a file lock.
    backup_path = os.path.join(os.path.dirname(OUT_5MIN), "_last_run_backup_v2.pkl")
    pd.to_pickle({"backtest": backtest_df, "monthly": monthly}, backup_path)
    print(f"\nBackup saved: {backup_path} (safe even if the CSV writes below fail)")

    safe_to_csv(backtest_df[FIVE_MIN_COLUMNS], OUT_5MIN)
    safe_to_csv(monthly, OUT_MONTHLY)
    # Written last, so it only ever describes outputs that were actually saved.
    safe_to_csv(build_run_record(ra_sheet_name, qty_by_hour_month, price_by_hour_month,
                                 ra_sustain_hours, monthly, n_steps), OUT_RUN_INPUTS)

    print(f"\nSaved: {OUT_5MIN} ({len(backtest_df):,} rows)")
    print(f"Saved: {OUT_MONTHLY} ({len(monthly):,} rows)")
    print(f"Saved: {OUT_RUN_INPUTS}")
    print("  -> in the workbook, press Import Python Run on the Python Run tab")
    print(f"\nRevenue by source:")
    print(f"  Energy arbitrage: ${monthly['P/L Total ($)'].sum():>14,.2f}")
    print(f"  RegUp:            ${monthly['AS Revenue - Reg Up ($)'].sum():>14,.2f}")
    print(f"  RegDown:          ${monthly['AS Revenue - Reg Down ($)'].sum():>14,.2f}")
    print(f"  Spin:             ${monthly['AS Revenue - Spinning ($)'].sum():>14,.2f}")
    print(f"  RA capacity:      ${monthly['AS Revenue - RA ($)'].sum():>14,.2f}")
    print(f"  TOTAL:            ${monthly['Grand Total Revenue ($)'].sum():>14,.2f}")
    print(f"\nEnding SOC: {backtest_df['soc_mwh'].iloc[-1]:.3f} MWh")


if __name__ == "__main__":
    main()
    # a double-clicked console closes the moment the script ends, taking the
    # results with it (errors are held open by _report_crash)
    _pause("\nFinished -- press Enter to close...")

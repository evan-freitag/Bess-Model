"""
BESS Revenue Simulation Model  v3.0 — Time-of-Day Dispatch
Tesla Megapack 2 XL — 28-year, 30-min timestep simulation.

New in v3.0:
  - Time-of-day dispatch replaces spread-based lookup table.
  - Dispatch schedule lives in tod_schedule.xlsx (auto-created with CAISO defaults).
  - Dynamic discharge floor: (avg_charge_cost / RTE) + profit_adder.
  - 48-hour lookahead charge gate: only charge if current price < (next_peak - adder) * RTE.
  - SOC targeting: boosts charge rate when approaching discharge window below target SOC.
  - Inactivity fallback: lowers discharge floor after extended idle period.
  - Falls back to spread-table dispatch (v2.4 behaviour) if tod_schedule.xlsx is absent.

Usage:  python bess_model_v2.py
  - Same inputs3.xlsx as v2.4 (all AS/battery parameters unchanged).
  - Edit tod_schedule.xlsx to change dispatch windows, price floors, % of max C-rate.
"""

import os, time, collections
import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────
# FILE PATHS
# ─────────────────────────────────────────────────────────────
BASE_DIR          = os.path.dirname(os.path.abspath(__file__))
DATA_EXCEL        = r"C:\Users\evanf\OneDrive\Desktop\bess model 1.2.xlsx"
INPUTS_EXCEL      = os.environ.get("BESS_INPUTS",  os.path.join(BASE_DIR, "inputs3.xlsx"))
OUTPUT_PATH       = os.environ.get("BESS_OUTPUT",  os.path.join(BASE_DIR, "bess_monthly_output_v2.xlsx"))
TOD_SCHEDULE_FILE = os.environ.get("BESS_TOD",     os.path.join(BASE_DIR, "tod_schedule.xlsx"))
PRICE_TEMP_DATA   = os.path.join(BASE_DIR, "as_data", "Price_Temp_Apr2023_Mar2026_30min.xlsx")

# ─────────────────────────────────────────────────────────────
# STYLE HELPERS  (identical to v2.4)
# ─────────────────────────────────────────────────────────────
_DARK_BLUE  = PatternFill("solid", fgColor="1F4E79")
_MED_BLUE   = PatternFill("solid", fgColor="2E75B6")
_LIGHT_BLUE = PatternFill("solid", fgColor="D6E4F0")
_LIGHT_GREY = PatternFill("solid", fgColor="F2F2F2")
_YELLOW     = PatternFill("solid", fgColor="FFFACD")
_TEAL_FILL  = PatternFill("solid", fgColor="1F6B55")
_WHITE_FONT = Font(bold=True, color="FFFFFF", size=11)
_BOLD       = Font(bold=True, size=11)
_NORMAL     = Font(size=11)
_ITALIC     = Font(italic=True, size=10, color="444444")
_THIN_SIDE  = Side(style="thin", color="BBBBBB")
_THIN_BDR   = Border(left=_THIN_SIDE, right=_THIN_SIDE,
                     top=_THIN_SIDE,  bottom=_THIN_SIDE)
_CENTER     = Alignment(horizontal="center", vertical="center")
_LEFT       = Alignment(horizontal="left",   vertical="center")

def _fmt(cell, font=None, fill=None, align=None, border=None):
    if font:   cell.font      = font
    if fill:   cell.fill      = fill
    if align:  cell.alignment = align
    if border: cell.border    = border

def _title_band(ws, row, text, ncols, fill=_DARK_BLUE, font=_WHITE_FONT):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1, text)
    _fmt(c, font=font, fill=fill, align=_CENTER)
    ws.row_dimensions[row].height = 20

def _note_band(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1, text)
    _fmt(c, font=_ITALIC, fill=_LIGHT_BLUE, align=_LEFT)
    ws.row_dimensions[row].height = 14

def _dump_df_at(ws, df, start_row):
    for r_off, row_vals in enumerate(df.itertuples(index=False, name=None)):
        for c_off, v in enumerate(row_vals):
            ws.cell(start_row + r_off, c_off + 1,
                    None if (isinstance(v, float) and np.isnan(v)) else v)

def _sec_band(ws, row, text, ncols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1, text)
    _fmt(c, font=_BOLD, fill=_LIGHT_BLUE, align=_LEFT)


# ─────────────────────────────────────────────────────────────
# TOD SCHEDULE  — default generator + file creator
# ─────────────────────────────────────────────────────────────
_MONTH_NAMES = ["January","February","March","April","May","June",
                "July","August","September","October","November","December"]
_MONTH_NUM   = {name: i+1 for i, name in enumerate(_MONTH_NAMES)}

def _default_tod():
    """Return calibrated (schedule_rows, targets, params) for CAISO duck curve.

    Key design change: overnight and solar-midday charge windows now carry
    SEPARATE ceilings, calibrated to actual CAISO seasonal price distributions.

    Problem with a single ceiling per month:
      Winter overnight (0–6h) prices = $25–55; a $30 cap blocked most hours.
      Summer midday solar prices often <$0 to $20; the same $30 cap was loose.
      Result: almost no winter charging → 0.20 cycles/day vs real BESS ~0.5–1.1.

    Fix: overnight_ceil reflects median winter overnight price ($44–48),
    midday_ceil reflects median solar-glut price ($16–28).
    Real CAISO BESS benchmark: 0.5–1.1 cycles/day (CAISO 2023–2024 reports).
    """
    # month → (overnight_hrs, midday_hrs, dis_hrs,
    #           overnight_ceil, midday_ceil, dis_floor, c_rate, soc_target)
    #
    # Ceilings set high ($999) so the 48-hour lookahead gate is the ONLY charge
    # limiter.  The gate already calculates whether charging is profitable:
    #   gate = (next_peak_48h - profit_adder) × RTE
    # A low hard ceiling on top blocks hours where the spread IS profitable,
    # which was why cycles stayed at ~0.2/day despite good summer/winter spreads.
    #
    # dis_floor: fallback sell floor used only when book_mwh < 0.001 (first period
    # or after long idle).  Dynamic floor (book_cost/RTE + adder) normally takes over.
    # Ceilings derived from actual CAISO 2023-2025 data:
    #   ceiling = dis_P50 × RTE × (1.02)^4   [break-even at median peak, yr-1 escalated]
    # Overnight and midday share the same ceiling (midday solar is always cheaper so the
    # ceiling is non-binding there; it catches expensive overnight hours in winter).
    # The 48-h lookahead gate adds a second filter: don't charge if the forward spread
    # is too thin regardless of absolute price.
    MONTH_CONFIG = {
        # ── Winter (Jan/Dec): ON P50=$46-47, ceil=$50 → ~55% ON + 70% MID below
        1:  (set(range(0, 7)), set(range(10, 14)), set(range(17, 22)), 50, 50, 45, 0.90, 0.85),
        2:  (set(range(0, 7)), set(range(10, 14)), set(range(17, 22)), 41, 41, 45, 0.90, 0.85),
        # ── Early spring (Mar): solar starting, ON P50=$31, MID P50=$15, dis P50=$34
        3:  (set(range(0, 6)), set(range(9, 14)),  set(range(16, 21)), 31, 31, 43, 0.90, 0.87),
        # ── Mid-spring (Apr/May): solar glut.  MID often negative, peaks $41-35.
        4:  (set(range(0, 5)), set(range(9, 15)),  set(range(16, 21)), 38, 38, 40, 0.90, 0.88),
        5:  (set(range(0, 5)), set(range(9, 15)),  set(range(16, 21)), 32, 32, 38, 0.90, 0.90),
        # ── Summer (Jun-Aug): strong solar. MID often negative, peaks $39-51.
        6:  (set(range(0, 6)), set(range(9, 15)),  set(range(16, 22)), 36, 36, 38, 0.90, 0.95),
        7:  (set(range(0, 6)), set(range(9, 15)),  set(range(16, 22)), 47, 47, 38, 0.90, 0.95),
        8:  (set(range(0, 6)), set(range(9, 15)),  set(range(16, 22)), 45, 45, 38, 0.90, 0.95),
        # ── Early fall (Sep): solar declining, ON P50=$39, dis P50=$48
        9:  (set(range(0, 5)), set(range(9, 14)),  set(range(16, 21)), 45, 45, 38, 0.90, 0.90),
        # ── Late fall (Oct/Nov): little solar, ON P50=$46-48, peaks $56-57
        10: (set(range(0, 5)), set(range(10, 14)), set(range(16, 21)), 52, 52, 40, 0.90, 0.87),
        # Nov overnight P50 $46-48 (expensive — charging there inflates book cost above discharge floor).
        # No overnight charging; midday ceiling $35 keeps book cost low so dynamic floor stays beatable.
        11: (set(),            set(range(10, 14)), set(range(17, 21)),  0, 35, 38, 0.90, 0.85),
        12: (set(range(0, 7)), set(range(10, 14)), set(range(17, 21)), 46, 46, 45, 0.90, 0.85),
    }

    rows = []
    targets = {}
    for m in range(1, 13):
        mname = _MONTH_NAMES[m - 1]
        onhrs, mnhrs, dhrs, on_cf, mn_cf, df_, cr, tgt = MONTH_CONFIG[m]
        targets[m] = tgt
        for h in range(24):
            if h in dhrs:
                rows.append((mname, h, "D", df_, cr))
            elif h in onhrs:
                rows.append((mname, h, "C", on_cf, cr))
            elif h in mnhrs:
                rows.append((mname, h, "C", mn_cf, cr))
            else:
                rows.append((mname, h, "H", 0.0, 0.0))

    params = {
        "profit_adder":           0.10,  # $/MWh above RTE breakeven — minimal margin, maximise cycles
        "inactivity_threshold":  24.0,   # hours idle before floor relaxes
        "inactivity_discount":    0.25,  # 25% floor reduction after idle period
        "lookahead_hours":       48.0,   # hours ahead to scan for peak price
        "charge_avg_window_days": 7.0,   # rolling window for book-cost tracking
        "rte":                    0.86,  # round-trip efficiency for floor calc
    }
    return rows, targets, params


def _create_tod_schedule_file():
    """Write tod_schedule.xlsx with default CAISO dispatch parameters."""
    rows, targets, params = _default_tod()

    wb = Workbook()
    ws = wb.active
    ws.title = "TOD Schedule"
    NC = 5  # columns

    _title_band(ws, 1, "BESS Time-of-Day Dispatch Schedule  (v3.0)", NC)
    _note_band(ws, 2,
        "Edit charge/discharge windows, price floors, and % of max C-rate below.  "
        "Mode: C = Charge  |  D = Discharge  |  H = Hold.  "
        "Price Floor for Charge = max $/MWh to buy at (charge ceiling).  "
        "For Discharge it is a fallback floor — the dynamic floor (book_cost/RTE + profit_adder) "
        "usually overrides it.  % of Max C-Rate: fraction of rated MW to use (1.0 = full power).", NC)

    r = 4
    _sec_band(ws, r, "GLOBAL PARAMETERS", NC); r += 1

    hdr_fill = PatternFill("solid", fgColor="4472C4")
    for col, hdr in enumerate(["Parameter", "Value", "Notes"], start=1):
        c = ws.cell(r, col, hdr)
        _fmt(c, font=_WHITE_FONT, fill=hdr_fill, align=_CENTER, border=_THIN_BDR)
    r += 1

    param_rows = [
        ("Profit Adder ($/MWh)",          params["profit_adder"],
         "Required $/MWh margin above breakeven (book_cost/RTE). Raise to be more selective."),
        ("Inactivity Threshold (hours)",   params["inactivity_threshold"],
         "Hours since last discharge before the floor is discounted. Default: 72 (3 days)."),
        ("Inactivity Discount (%)",        params["inactivity_discount"],
         "Fraction to reduce discharge floor after idle period. e.g. 0.20 = 20% reduction."),
        ("Lookahead Hours",                params["lookahead_hours"],
         "Hours ahead to scan for peak price. Charge only if price < (peak - adder) * RTE."),
        ("Charge Avg Window (days)",       params["charge_avg_window_days"],
         "Rolling window (days) for tracking average charge cost (book cost)."),
        ("Round-Trip Efficiency",          params["rte"],
         "Used to compute discharge breakeven: floor = book_cost / RTE + profit_adder."),
    ]
    for label, val, note in param_rows:
        ws.cell(r, 1, label); ws.cell(r, 2, val); ws.cell(r, 3, note)
        _fmt(ws.cell(r, 1), font=_BOLD, fill=_LIGHT_GREY, align=_LEFT, border=_THIN_BDR)
        _fmt(ws.cell(r, 2), fill=_YELLOW, align=_CENTER, border=_THIN_BDR)
        _fmt(ws.cell(r, 3), font=_ITALIC, align=_LEFT)
        r += 1

    r += 1
    _sec_band(ws, r, "MONTHLY SOC TARGETS  (target SOC % when discharge window opens)", NC); r += 1
    for col, hdr in enumerate(["Month", "Target SOC at Discharge Start (%)"], start=1):
        c = ws.cell(r, col, hdr)
        _fmt(c, font=_WHITE_FONT, fill=hdr_fill, align=_CENTER, border=_THIN_BDR)
    r += 1
    for m in range(1, 13):
        ws.cell(r, 1, _MONTH_NAMES[m-1])
        ws.cell(r, 2, targets[m])
        _fmt(ws.cell(r, 1), font=_BOLD, fill=_LIGHT_GREY, align=_LEFT, border=_THIN_BDR)
        _fmt(ws.cell(r, 2), fill=_YELLOW, align=_CENTER, border=_THIN_BDR)
        r += 1

    r += 1
    _sec_band(ws, r, "HOURLY DISPATCH SCHEDULE  (288 rows = 12 months × 24 hours)", NC); r += 1
    _note_band(ws, r,
        "Month column: only the first hour row per month needs the month name — "
        "subsequent rows inherit it. Price Floor is ignored for Hold periods.", NC); r += 1

    for col, hdr in enumerate(["Month", "Hour (0-23)", "Mode", "Price Floor ($/MWh)", "% of Max C-Rate"], start=1):
        c = ws.cell(r, col, hdr)
        _fmt(c, font=_WHITE_FONT, fill=hdr_fill, align=_CENTER, border=_THIN_BDR)
    r += 1

    alt = PatternFill("solid", fgColor="EBF3FB")
    prev_month = None
    for row_data in rows:
        mname, hour, mode, floor, crate = row_data
        show_month = mname if mname != prev_month else ""
        prev_month = mname
        fill = alt if (r % 2 == 0) else None
        ws.cell(r, 1, show_month)
        ws.cell(r, 2, hour)
        ws.cell(r, 3, mode)
        ws.cell(r, 4, floor if mode != "H" else None)
        ws.cell(r, 5, crate if mode != "H" else None)
        mode_fill = (PatternFill("solid", fgColor="E2EFDA") if mode == "C"
                     else PatternFill("solid", fgColor="FCE4D6") if mode == "D"
                     else fill)
        _fmt(ws.cell(r, 1), font=Font(bold=(show_month != ""), size=10), align=_LEFT)
        _fmt(ws.cell(r, 2), align=_CENTER)
        _fmt(ws.cell(r, 3), fill=mode_fill, align=_CENTER, font=Font(bold=True, size=10))
        if fill:
            _fmt(ws.cell(r, 4), fill=fill, align=_CENTER)
            _fmt(ws.cell(r, 5), fill=fill, align=_CENTER)
        else:
            _fmt(ws.cell(r, 4), align=_CENTER)
            _fmt(ws.cell(r, 5), align=_CENTER)
        r += 1

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 16
    ws.freeze_panes = "A2"

    wb.save(TOD_SCHEDULE_FILE)
    print(f"  Created {TOD_SCHEDULE_FILE}  (default CAISO seasonal schedule)")


def _load_tod_schedule():
    """
    Read tod_schedule.xlsx.  Returns (params_dict, targets_dict, schedule_dict).

    schedule_dict: {month_int: {hour_int: {"mode":str, "price_floor":float, "c_rate":float}}}
    targets_dict:  {month_int: soc_target_float}
    params_dict:   {"profit_adder": float, "inactivity_threshold": float, ...}
    """
    xl = pd.ExcelFile(TOD_SCHEDULE_FILE)
    raw = pd.read_excel(xl, "TOD Schedule", header=None)

    # ── Locate section headers ─────────────────────────────────
    def _find(marker):
        for idx, row in raw.iterrows():
            for v in row:
                if pd.notna(v) and str(v).strip().startswith(marker):
                    return idx
        raise ValueError(f"TOD Schedule: section '{marker}' not found")

    params_sec   = _find("GLOBAL PARAMETERS")
    targets_sec  = _find("MONTHLY SOC TARGETS")
    schedule_sec = _find("HOURLY DISPATCH SCHEDULE")

    # ── Global parameters ──────────────────────────────────────
    KEY_MAP = {
        "Profit Adder":          "profit_adder",
        "Inactivity Threshold":  "inactivity_threshold",
        "Inactivity Discount":   "inactivity_discount",
        "Lookahead Hours":       "lookahead_hours",
        "Charge Avg Window":     "charge_avg_window_days",
        "Round-Trip Efficiency": "rte",
    }
    params = {}
    for i in range(params_sec + 2, targets_sec):
        label = str(raw.iloc[i, 0]).strip() if pd.notna(raw.iloc[i, 0]) else ""
        val   = raw.iloc[i, 1]
        for key, pykey in KEY_MAP.items():
            if label.startswith(key) and pd.notna(val):
                params[pykey] = float(val)
    # Fill any missing keys with defaults
    _, _, _dp = _default_tod()
    for k, v in _dp.items():
        params.setdefault(k, v)
    params["lookahead_periods"]      = int(params["lookahead_hours"] * 2)
    params["charge_window_periods"]  = int(params["charge_avg_window_days"] * 48)
    params["inactivity_threshold_hp"]= int(params["inactivity_threshold"] * 2)

    # ── Monthly targets ────────────────────────────────────────
    targets = {}
    for i in range(targets_sec + 2, schedule_sec):
        mname = str(raw.iloc[i, 0]).strip() if pd.notna(raw.iloc[i, 0]) else ""
        val   = raw.iloc[i, 1]
        if mname in _MONTH_NUM and pd.notna(val):
            targets[_MONTH_NUM[mname]] = float(val)
    for m in range(1, 13):
        targets.setdefault(m, 0.85)

    # ── Hourly schedule ────────────────────────────────────────
    schedule = {m: {h: {"mode": "H", "price_floor": 0.0, "c_rate": 0.0}
                    for h in range(24)}
                for m in range(1, 13)}

    cur_month = None
    for i in range(schedule_sec + 3, len(raw)):
        row = raw.iloc[i]
        mval = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if mval in _MONTH_NUM:
            cur_month = _MONTH_NUM[mval]
        hour_v  = row.iloc[1]
        mode_v  = str(row.iloc[2]).strip().upper() if pd.notna(row.iloc[2]) else "H"
        floor_v = row.iloc[3]
        crate_v = row.iloc[4]
        if cur_month and pd.notna(hour_v):
            h = int(hour_v)
            if 0 <= h <= 23:
                schedule[cur_month][h] = {
                    "mode":        mode_v if mode_v in ("C", "D", "H") else "H",
                    "price_floor": float(floor_v) if pd.notna(floor_v) else 0.0,
                    "c_rate":      float(crate_v) if pd.notna(crate_v) else 0.0,
                }

    # ── Precompute: hours until next discharge window per (month, hour) ──
    # Used for SOC-targeting boost in charge windows.
    next_dis_hours = {}
    for m in range(1, 13):
        next_dis_hours[m] = {}
        for h in range(24):
            for offset in range(1, 25):
                if schedule[m][(h + offset) % 24]["mode"] == "D":
                    next_dis_hours[m][h] = offset
                    break
            else:
                next_dis_hours[m][h] = 24
    params["next_dis_hours"] = next_dis_hours

    return params, targets, schedule


def _forward_max(arr, window):
    """
    O(N) monotonic-deque forward rolling max.
    result[i] = max(arr[i+1 : i+1+window])  (looks forward, not backward)
    Falls back to arr[i] at the tail where no forward data exists.
    """
    n = len(arr)
    result = np.empty(n, dtype=float)
    dq = collections.deque()  # decreasing deque of indices into arr
    for i in range(n - 1, -1, -1):
        j = i + 1  # first index in the forward window
        if j < n:
            while dq and arr[dq[-1]] <= arr[j]:
                dq.pop()
            dq.append(j)
        while dq and dq[0] >= i + 1 + window:
            dq.popleft()
        result[i] = arr[dq[0]] if dq else arr[i]
    return result


# ─────────────────────────────────────────────────────────────
# STARTUP: create inputs.xlsx if needed (identical logic to v2.4)
# ─────────────────────────────────────────────────────────────
_needs_rebuild = False
if not os.path.exists(INPUTS_EXCEL):
    _needs_rebuild = True
else:
    try:
        _existing_sheets = pd.ExcelFile(INPUTS_EXCEL).sheet_names
        if "Price & Temp Data" not in _existing_sheets:
            print("inputs3.xlsx is missing the Price & Temp Data tab — please rebuild from v2.4 model.")
            _needs_rebuild = False  # don't auto-rebuild; v2.4 handles that
    except PermissionError:
        print("WARNING: inputs3.xlsx appears to be open in Excel. Close it and re-run.")

# ─────────────────────────────────────────────────────────────
# TOD SCHEDULE: load or create
# ─────────────────────────────────────────────────────────────
print("Checking TOD schedule...")
if not os.path.exists(TOD_SCHEDULE_FILE):
    print(f"  tod_schedule.xlsx not found — creating with CAISO defaults...")
    _create_tod_schedule_file()

try:
    _tod_params, _tod_targets, _tod_schedule = _load_tod_schedule()
    USE_TOD = True
    print(f"  TOD dispatch ACTIVE  |  profit_adder=${_tod_params['profit_adder']:.1f}/MWh"
          f"  |  lookahead={_tod_params['lookahead_hours']:.0f}h"
          f"  |  RTE={_tod_params['rte']:.0%}")
except Exception as _e:
    USE_TOD = False
    print(f"  WARNING: could not load TOD schedule ({_e}) — falling back to spread table")


# ─────────────────────────────────────────────────────────────
# 1.  LOAD INPUTS  (from "Inputs" tab of inputs3.xlsx)
# ─────────────────────────────────────────────────────────────
print("Loading inputs...")
_raw = pd.read_excel(INPUTS_EXCEL, sheet_name="Inputs", header=None, index_col=0)
inp  = {str(k).strip(): v
        for k, v in _raw.iloc[:, 0].items()
        if pd.notna(k) and str(k).strip() and pd.notna(v)}

N_PACKS             = int(inp["Number of Units"])
NAMEPLATE_MWH_INPUT = float(inp["Total MWh"])
MWH_PER_PACK        = NAMEPLATE_MWH_INPUT / N_PACKS
INITIAL_SOC         = float(inp["Initial SOC"])
SOH_FLOOR           = float(inp["SOH floor (warranty limit)"])
CABLE_EFF           = float(inp["Cable efficiency"])

RETROFIT_CAPEX           = float(inp.get("Retrofit CapEx ($)", 0))
RETROFIT_DAYS            = float(inp.get("Retrofit Downtime (days)", 0))
RETROFIT_DOWNTIME_PERIODS= int(round(RETROFIT_DAYS * 48))
SOH_AFTER_RETROFIT       = float(inp.get("SOH After Retrofit (%)", SOH_FLOOR))
COD_YEAR    = pd.to_datetime(inp["COD / model start date"]).year
MODEL_YEARS = int(inp["Model length (years)"])
PRICE_ESC   = float(inp["Price escalator (% per year)"])
TEMP_ESC    = float(inp["Temp escalator (F per year)"])

NAMEPLATE_MWH = N_PACKS * MWH_PER_PACK
MAX_MW        = NAMEPLATE_MWH / 4.0
SOC_MIN, SOC_MAX = 0.05, 0.95

RA_MW_PP = float(inp.get("RA MW per pack",       0)); RA_PRC = float(inp.get("RA Price Year 1", 0));  RA_ESC = float(inp.get("RA Price Escalator",  0.02))
RU_MW_PP = float(inp.get("Reg Up MW per pack",   0)); RU_PRC = float(inp.get("Reg Up Price Year 1", 0)); RU_ESC = float(inp.get("Reg Up Escalator",    0.02))
RD_MW_PP = float(inp.get("Reg Down MW per pack", 0)); RD_PRC = float(inp.get("Reg Down Price Year 1", 0)); RD_ESC = float(inp.get("Reg Down Escalator",  0.02))
SP_MW_PP = float(inp.get("Spinning MW per pack", 0)); SP_PRC = float(inp.get("Spinning Price Year 1", 0)); SP_ESC = float(inp.get("Spinning Escalator",  0.02))
NS_MW_PP = float(inp.get("Non-Spin MW per pack", 0)); NS_PRC = float(inp.get("Non-Spin Price Year 1", 0)); NS_ESC = float(inp.get("Non-Spin Escalator",  0.02))

if os.environ.get("BESS_RU_PP"): RU_MW_PP = float(os.environ["BESS_RU_PP"])
if os.environ.get("BESS_RD_PP"): RD_MW_PP = float(os.environ["BESS_RD_PP"])
if os.environ.get("BESS_SP_PP"): SP_MW_PP = float(os.environ["BESS_SP_PP"])
if os.environ.get("BESS_NS_PP"): NS_MW_PP = float(os.environ["BESS_NS_PP"])

RU_FLOOR = float(inp.get("Reg Up Min Bid ($/MWh)",   0.50))
RD_FLOOR = float(inp.get("Reg Down Min Bid ($/MWh)", 0.50))
SP_FLOOR = float(inp.get("Spinning Min Bid ($/MWh)", 0.25))
NS_FLOOR = float(inp.get("Non-Spin Min Bid ($/MWh)", 0.10))

RA_MW = RA_MW_PP * N_PACKS
RU_MW = RU_MW_PP * N_PACKS
RD_MW = RD_MW_PP * N_PACKS
SP_MW = SP_MW_PP * N_PACKS
NS_MW = NS_MW_PP * N_PACKS

_dis_reserved = RU_MW + SP_MW + NS_MW
_chg_reserved = RD_MW
MAX_CRATE_DIS = max(0.0, (MAX_MW - _dis_reserved) / MAX_MW)
MAX_CRATE_CHG = max(0.0, (MAX_MW - _chg_reserved) / MAX_MW)
SOC_MIN_EFF   = min(SOC_MAX, SOC_MIN + _dis_reserved * 0.5 / NAMEPLATE_MWH)
SOC_MAX_EFF   = max(SOC_MIN_EFF, SOC_MAX - _chg_reserved * 0.5 / NAMEPLATE_MWH)

print(f"  {N_PACKS} packs x {MWH_PER_PACK} MWh = {NAMEPLATE_MWH:.0f} MWh  |  {MAX_MW:.1f} MW max")
print(f"  {MODEL_YEARS} years from {COD_YEAR}  |  price esc {PRICE_ESC*100:.1f}%/yr")
if _dis_reserved > 0 or _chg_reserved > 0:
    print(f"  AS discharge cap: {MAX_CRATE_DIS:.1%}  |  charge cap: {MAX_CRATE_CHG:.1%}")


# ─────────────────────────────────────────────────────────────
# 2.  LOAD PRICE & TEMPERATURE DATA
# ─────────────────────────────────────────────────────────────
print("Loading price/temp data...")
xl_inp = pd.ExcelFile(INPUTS_EXCEL)
raw = pd.read_excel(xl_inp, "Price & Temp Data", header=0, skiprows=2, usecols=range(9))
raw.columns = ["op_hour", "datetime", "price", "temp", "month",
               "ns_clr", "rd_clr", "ru_clr", "sp_clr"]
raw["datetime"] = pd.to_datetime(raw["datetime"])

prices_base = raw["price"].values.astype(float)
temps_base  = raw["temp"].values.astype(float)
ns_base     = raw["ns_clr"].values.astype(float)
rd_base     = raw["rd_clr"].values.astype(float)
ru_base     = raw["ru_clr"].values.astype(float)
sp_base     = raw["sp_clr"].values.astype(float)

# Hour-of-day and calendar month from actual timestamps (used by TOD dispatch)
hours_of_day_base = raw["datetime"].dt.hour.values.astype(np.int8)
data_months_base  = raw["datetime"].dt.month.values.astype(np.int8)

def _fill_nan(arr):
    if np.isnan(arr).any():
        return pd.Series(arr).interpolate(method="linear", limit_direction="both").values
    return arr

prices_base = _fill_nan(prices_base)
temps_base  = _fill_nan(temps_base)
n_nan = np.isnan(raw["price"].values).sum()
if n_nan:
    print(f"  Interpolated {n_nan} missing price rows")

N_PERIODS    = 17520
N_BASE_YEARS = max(1, len(prices_base) // N_PERIODS)
HIST_START_YEAR = int(raw["datetime"].iloc[0].year)
print(f"  {N_BASE_YEARS} historical years detected (starting {HIST_START_YEAR})")


# ─────────────────────────────────────────────────────────────
# 3.  EFFICIENCY TABLE
# ─────────────────────────────────────────────────────────────
eff_raw   = pd.read_excel(xl_inp, "Efficiency Table", header=None, skiprows=2)
EFF_TEMPS = np.array([32., 59., 77., 95., 105., 150.], dtype=float)
EFF_H     = eff_raw.iloc[1, 2:8].values.astype(float)
EFF_D     = eff_raw.iloc[2, 2:8].values.astype(float)
EFF_C     = eff_raw.iloc[3, 2:8].values.astype(float)


# ─────────────────────────────────────────────────────────────
# 4.  DISPATCH TABLE  (kept as fallback when TOD is disabled)
# ─────────────────────────────────────────────────────────────
print("Building spread-table dispatch (fallback)...")
disp_raw    = pd.read_excel(xl_inp, "Dispatch Table", header=None, skiprows=2)
SPREAD_VALS = pd.to_numeric(disp_raw.iloc[2:, 0], errors="coerce").values.astype(float)
n_spread    = len(SPREAD_VALS)
TEMP_BANDS  = np.array([80., 90., 100., 110., 115.])
TEMP_COL_START = [1, 7, 13, 19, 25, 31]

DISPATCH = np.zeros((n_spread, 6, 5), dtype=float)
for t_i, t_col in enumerate(TEMP_COL_START):
    for s_i in range(5):
        col_vals = pd.to_numeric(disp_raw.iloc[2:, t_col + s_i], errors="coerce").fillna(0).values
        DISPATCH[:, t_i, s_i] = col_vals[:n_spread]

_disp_mods_path = os.environ.get("BESS_DISPATCH_MODS")
if _disp_mods_path and os.path.isfile(_disp_mods_path):
    import json
    with open(_disp_mods_path) as _f:
        for _sp, _ti, _si, _cr in json.load(_f):
            DISPATCH[_sp, _ti, _si] = _cr


# ─────────────────────────────────────────────────────────────
# 5.  DEGRADATION TABLES
# ─────────────────────────────────────────────────────────────
deg_raw   = pd.read_excel(xl_inp, "Degradation Tables", header=None, skiprows=2)
_tf       = deg_raw.iloc[3:, [2, 3]].apply(pd.to_numeric, errors="coerce").dropna()
DEG_T_VAL = _tf.iloc[:, 0].values.astype(float)
DEG_T_FAC = _tf.iloc[:, 1].values.astype(float)

_DEG_TF_LUT = np.ones(260, dtype=float)
for _t in range(260):
    _i = int(np.searchsorted(DEG_T_VAL, float(_t), side="right")) - 1
    _DEG_TF_LUT[_t] = DEG_T_FAC[max(0, min(_i, len(DEG_T_FAC) - 1))]

def _parse_deg(start_row):
    tbl = np.zeros((5, 6, 2), dtype=float)
    for si in range(5):
        row = deg_raw.iloc[start_row + si, :]
        for ci in range(6):
            c_v = pd.to_numeric(row.iloc[11 + ci * 2],     errors="coerce")
            d_v = pd.to_numeric(row.iloc[11 + ci * 2 + 1], errors="coerce")
            tbl[si, ci, 0] = 0.0 if pd.isna(c_v) else float(c_v)
            tbl[si, ci, 1] = 0.0 if pd.isna(d_v) else float(d_v)
    return tbl

DEG_LO = _parse_deg(4)
DEG_HI = _parse_deg(14)


# ─────────────────────────────────────────────────────────────
# 6.  SIMULATION
# ─────────────────────────────────────────────────────────────
print(f"\nRunning simulation ({MODEL_YEARS} years × {N_PERIODS:,} = "
      f"{MODEL_YEARS * N_PERIODS:,} timesteps)  [dispatch: {'TOD' if USE_TOD else 'spread table'}]...")
t0 = time.time()

total = MODEL_YEARS * N_PERIODS

_MONTH_DAYS = np.array([31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31])
_MONTH_CUM  = np.concatenate(([0], np.cumsum(_MONTH_DAYS)))
_cod_ts     = pd.to_datetime(inp["COD / model start date"])
_cod_doy0   = int((_cod_ts - pd.Timestamp(year=_cod_ts.year, month=1, day=1)).days)

_elapsed_days = np.arange(total) // 48 + _cod_doy0
_year_off     = _elapsed_days // 365
_day_in_yr    = _elapsed_days % 365
PROJ_YEAR     = (COD_YEAR + _year_off).astype(np.int16)
PROJ_MONTH    = np.searchsorted(_MONTH_CUM, _day_in_yr, side="right").astype(np.int8)

out_year   = np.empty(total, dtype=np.int16)
out_month  = np.empty(total, dtype=np.int8)
out_pl     = np.empty(total, dtype=float)
out_action = np.empty(total, dtype=np.int8)
out_dmwh   = np.empty(total, dtype=float)
out_adelta = np.empty(total, dtype=float)   # actual SOC change (includes AS-forced)
out_meter  = np.empty(total, dtype=float)
out_soh    = np.empty(total, dtype=float)
out_ru_rev = np.zeros(total, dtype=float)
out_rd_rev = np.zeros(total, dtype=float)
out_sp_rev = np.zeros(total, dtype=float)
out_ns_rev = np.zeros(total, dtype=float)
out_ra_rev = np.zeros(total, dtype=float)
out_capex  = np.zeros(total, dtype=float)

soh = 1.0
soc = INITIAL_SOC
cap = NAMEPLATE_MWH
ptr = 0
retrofit_pause_left = 0
retrofit_log = []

# ── TOD persistent state (survives year boundaries) ───────────
if USE_TOD:
    _tod_rte        = _tod_params["rte"]
    _tod_adder      = _tod_params["profit_adder"]
    _tod_inact_hp   = _tod_params["inactivity_threshold_hp"]
    _tod_inact_disc = _tod_params["inactivity_discount"]
    _tod_lp         = _tod_params["lookahead_periods"]
    _tod_cw         = _tod_params["charge_window_periods"]
    _tod_ndh        = _tod_params["next_dis_hours"]

    _book_q    = collections.deque()  # (cost_entry, mwh_entry) per period in rolling window
    _book_cost = 0.0   # running sum of (price * mwh_charged) in window
    _book_mwh  = 0.0   # running sum of mwh_charged in window
    _idle_hp   = 0     # half-periods since last discharge

# ── Diagnostics ───────────────────────────────────────────────
if USE_TOD:
    _diag_charge_gated  = 0   # charge periods where lookahead gate blocked charging
    _diag_idle_relaxed  = 0   # discharge periods where inactivity discount was applied
    _diag_soc_boost     = 0   # charge periods where SOC target boost was applied

for year in range(MODEL_YEARS):
    data_yr   = year % N_BASE_YEARS
    cycle_num = year // N_BASE_YEARS
    yrs_esc   = max(0, COD_YEAR - HIST_START_YEAR) + cycle_num * N_BASE_YEARS

    s, e    = data_yr * N_PERIODS, (data_yr + 1) * N_PERIODS
    prices  = prices_base[s:e] * (1.0 + PRICE_ESC) ** yrs_esc
    temps   = temps_base[s:e]  +  TEMP_ESC * yrs_esc

    ru_mask = ru_base[s:e] >= RU_FLOOR
    rd_mask = rd_base[s:e] >= RD_FLOOR
    sp_mask = sp_base[s:e] >= SP_FLOOR
    ns_mask = ns_base[s:e] >= NS_FLOOR

    ru_esc_f    = (1.0 + RU_ESC) ** yrs_esc
    rd_esc_f    = (1.0 + RD_ESC) ** yrs_esc
    sp_esc_f    = (1.0 + SP_ESC) ** yrs_esc
    ns_esc_f    = (1.0 + NS_ESC) ** yrs_esc
    price_esc_f = (1.0 + PRICE_ESC) ** yrs_esc   # escalate TOD ceilings with prices

    ra_esc_f      = (1.0 + RA_ESC) ** year
    RA_PERIOD_REV = RA_MW * 1000.0 * RA_PRC * ra_esc_f * 12.0 / N_PERIODS

    # ── Pre-computation (vectorised) ──────────────────────────
    ei_arr  = np.clip(np.searchsorted(EFF_TEMPS, temps, side="right") - 1, 0, 5)
    t_idx   = np.minimum(np.searchsorted(np.array([80., 90., 100., 110., 115.]), temps, side="left"), 5)

    if USE_TOD:
        hours_arr  = hours_of_day_base[s:e]    # hour 0-23 per period
        months_arr = data_months_base[s:e]     # calendar month 1-12 per period

        # 48-hour forward rolling max (O(N) monotonic deque)
        next_peak_arr = _forward_max(prices, _tod_lp)

        # 7-day backward rolling average (for optional informational use)
        # (Not currently used in dispatch — book cost drives charge decisions)
    else:
        # Spread + index arrays for legacy dispatch
        cs    = np.empty(N_PERIODS + 1); cs[0] = 0.0
        np.cumsum(prices, out=cs[1:])
        iv    = np.arange(N_PERIODS)
        sts   = np.maximum(0, iv - 48)
        cnts  = np.where(iv > 0, iv - sts, 1)
        avg_a = np.where(iv > 0, (cs[iv] - cs[sts]) / cnts, prices[0])
        denom      = np.maximum(np.abs(avg_a), 10.0)
        spread_arr = np.clip((prices - avg_a) / denom, -25.0, 3.0)
        sp_idx     = np.clip(np.searchsorted(SPREAD_VALS, spread_arr, side="right") - 1,
                             0, n_spread - 1)

    # ── Inner loop ────────────────────────────────────────────
    for i in range(N_PERIODS):
        if retrofit_pause_left > 0:
            out_year[ptr]=PROJ_YEAR[ptr]; out_month[ptr]=PROJ_MONTH[ptr]
            out_pl[ptr]=0.0; out_action[ptr]=0; out_dmwh[ptr]=0.0; out_adelta[ptr]=0.0
            out_meter[ptr]=0.0; out_soh[ptr]=soh
            out_ru_rev[ptr]=0.0; out_rd_rev[ptr]=0.0
            out_sp_rev[ptr]=0.0; out_ns_rev[ptr]=0.0
            out_ra_rev[ptr]=0.0; out_capex[ptr]=0.0
            retrofit_pause_left -= 1
            if retrofit_pause_left == 0:
                soh = SOH_AFTER_RETROFIT
                cap = NAMEPLATE_MWH * soh
            if USE_TOD: _idle_hp += 1
            ptr += 1
            continue

        temp  = temps[i];  price = prices[i]
        ei    = int(ei_arr[i])

        # Period-specific AS carve-outs
        p_dis_res  = ((RU_MW if ru_mask[i] else 0.0)
                     + (SP_MW if sp_mask[i] else 0.0)
                     + (NS_MW if ns_mask[i] else 0.0))
        p_chg_res  = RD_MW if rd_mask[i] else 0.0
        p_cdis_max = max(0.0, (MAX_MW - p_dis_res) / MAX_MW) if MAX_MW > 0 else 0.0
        p_cchg_max = max(0.0, (MAX_MW - p_chg_res) / MAX_MW) if MAX_MW > 0 else 0.0
        p_soc_lo   = min(SOC_MAX, SOC_MIN + p_dis_res * 0.5 / NAMEPLATE_MWH) if NAMEPLATE_MWH > 0 else SOC_MIN
        p_soc_hi   = max(p_soc_lo, SOC_MAX - p_chg_res * 0.5 / NAMEPLATE_MWH) if NAMEPLATE_MWH > 0 else SOC_MAX

        # SOC band (for degradation table; also used by legacy dispatch)
        if   soc >= 0.95: s_i = 4
        elif soc >= 0.80: s_i = 3
        elif soc >= 0.20: s_i = 2
        elif soc >= 0.05: s_i = 1
        else:             s_i = 0

        # ── DISPATCH DECISION ──────────────────────────────────
        if USE_TOD:
            _hr  = int(hours_arr[i])
            _mon = int(months_arr[i])
            _sched = _tod_schedule[_mon][_hr]
            _mode  = _sched["mode"]
            _cr    = _sched["c_rate"]

            if _mode == "H":
                crate = 0.0

            elif _mode == "D":
                # Dynamic floor: average book cost / RTE + profit adder
                if _book_mwh > 1e-3:
                    _rte_floor = (_book_cost / _book_mwh) / _tod_rte
                    _dfloor    = _rte_floor + _tod_adder
                else:
                    _rte_floor = _sched["price_floor"] / _tod_rte
                    _dfloor    = _sched["price_floor"]

                # Inactivity discount: only reduce the profit adder, never below RTE break-even
                if _idle_hp >= _tod_inact_hp:
                    _dfloor = _rte_floor + _tod_adder * (1.0 - _tod_inact_disc)
                    _diag_idle_relaxed += 1

                crate = -_cr if price > _dfloor else 0.0

            elif _mode == "C":
                _gate = min((next_peak_arr[i] - _tod_adder) * _tod_rte,
                            _sched["price_floor"])
                if price < _gate:
                    # SOC targeting boost: if discharge window is soon and SOC is behind
                    _ndh = _tod_ndh[_mon][_hr]
                    _tgt = _tod_targets[_mon]
                    _effective_cr = _cr
                    if _ndh <= 6 and soc < _tgt:
                        _effective_cr = min(1.0, _cr + 0.25)
                        _diag_soc_boost += 1
                    crate = _effective_cr
                else:
                    crate = 0.0
                    _diag_charge_gated += 1
            else:
                crate = 0.0

        else:
            # Legacy spread-table dispatch
            crate = DISPATCH[int(sp_idx[i]), int(t_idx[i]), s_i]

        # SOC bounds (p_soc_lo / p_soc_hi) protect AS reserves — no separate C-rate cap needed

        para = EFF_H[ei] * N_PACKS

        if crate == 0.0:
            action = 0;  abs_crate = 0.0
            battery_mw = 0.0
            meter_mw   = para
        elif crate > 0.0:
            action = 1;  abs_crate = crate
            battery_mw =  abs_crate * MAX_MW * soh - para
            meter_mw   = (abs_crate * MAX_MW * soh + EFF_C[ei] * N_PACKS * abs_crate + para) / CABLE_EFF
        else:
            action = -1; abs_crate = -crate
            battery_mw = -(abs_crate * MAX_MW * soh)
            meter_mw   =  (battery_mw + EFF_D[ei] * N_PACKS * abs_crate + para) * CABLE_EFF

        pl        = -meter_mw * price * 0.5
        delta_mwh =  battery_mw * 0.5

        _ru_i = (RU_MW * ru_base[s+i] * ru_esc_f * 0.5) if ru_mask[i] else 0.0
        _rd_i = (RD_MW * rd_base[s+i] * rd_esc_f * 0.5) if rd_mask[i] else 0.0
        _sp_i = (SP_MW * sp_base[s+i] * sp_esc_f * 0.5) if sp_mask[i] else 0.0
        _ns_i = (NS_MW * ns_base[s+i] * ns_esc_f * 0.5) if ns_mask[i] else 0.0

        # SOC update
        prev_mwh = soc * cap
        new_mwh  = prev_mwh + delta_mwh
        if   new_mwh < p_soc_lo * cap: new_mwh = p_soc_lo * cap
        elif new_mwh > p_soc_hi * cap: new_mwh = p_soc_hi * cap
        soc = new_mwh / cap if cap > 0.0 else SOC_MIN

        actual_delta = new_mwh - prev_mwh
        if abs(actual_delta - delta_mwh) > 1e-9:
            # Sign-flip: SOC clamp reversed direction (e.g. tried to discharge but SOC was
            # already below p_soc_lo, so battery was forced up instead). The simple scale
            # formula produces a sign-flipped meter_mw and garbles the P&L.
            _sign_flipped = (abs(delta_mwh) > 1e-9) and (
                (actual_delta > 1e-9) != (delta_mwh > 1e-9)
            )
            if action == 0 or _sign_flipped:
                # AS-forced correction: bill energy at spot, HVAC still running
                meter_mw  = para + actual_delta * 2.0
                pl        = -meter_mw * price * 0.5
            else:
                scale     = actual_delta / delta_mwh if abs(delta_mwh) > 1e-9 else 0.0
                meter_mw  = meter_mw * scale
                pl        = -meter_mw * price * 0.5
            delta_mwh = actual_delta

        # ── TOD book-cost tracking ─────────────────────────────
        if USE_TOD:
            # Cost entry: $ paid (positive when charging, 0 otherwise)
            if action == 1 and actual_delta > 1e-9:
                _entry_cost = -pl             # pl < 0 when charging → cost = -pl > 0
                _entry_mwh  = actual_delta    # MWh stored in battery
            else:
                _entry_cost = 0.0
                _entry_mwh  = 0.0

            _book_q.append((_entry_cost, _entry_mwh))
            _book_cost += _entry_cost
            _book_mwh  += _entry_mwh
            if len(_book_q) > _tod_cw:
                _oc, _om = _book_q.popleft()
                _book_cost -= _oc
                _book_mwh  -= _om

            # Idle counter
            if action == -1:
                _idle_hp = 0
            else:
                _idle_hp += 1

        # Degradation
        throughput = abs(delta_mwh)
        if throughput > 0.0 and abs_crate > 0.0:
            if   soc >= 0.9: di = 4
            elif soc >= 0.8: di = 3
            elif soc >= 0.2: di = 2
            elif soc >= 0.1: di = 1
            else:            di = 0
            if   abs_crate >= 1.0: ci = 5
            elif abs_crate >= 0.8: ci = 4
            elif abs_crate >= 0.6: ci = 3
            elif abs_crate >= 0.4: ci = 2
            elif abs_crate >= 0.2: ci = 1
            else:                  ci = 0
            tbl      = DEG_HI if temp > 85.0 else DEG_LO
            base_deg = tbl[di, ci, 1 if action < 0 else 0]
            tf       = _DEG_TF_LUT[min(259, max(0, int(temp)))]
            soh      = soh - base_deg * tf * throughput
            cap      = NAMEPLATE_MWH * soh

        capex_this_period = 0.0
        if soh <= SOH_FLOOR:
            soh = SOH_FLOOR
            cap = NAMEPLATE_MWH * soh
            capex_this_period = RETROFIT_CAPEX
            retrofit_log.append((PROJ_YEAR[ptr], PROJ_MONTH[ptr], RETROFIT_CAPEX))
            if RETROFIT_DOWNTIME_PERIODS > 0:
                retrofit_pause_left = RETROFIT_DOWNTIME_PERIODS
            else:
                soh = SOH_AFTER_RETROFIT
                cap = NAMEPLATE_MWH * soh

        out_year[ptr]=PROJ_YEAR[ptr]; out_month[ptr]=PROJ_MONTH[ptr]
        out_pl[ptr]=pl;       out_action[ptr]=action
        out_dmwh[ptr]=delta_mwh; out_adelta[ptr]=actual_delta; out_meter[ptr]=meter_mw
        out_soh[ptr]=soh
        out_ru_rev[ptr]=_ru_i; out_rd_rev[ptr]=_rd_i
        out_sp_rev[ptr]=_sp_i; out_ns_rev[ptr]=_ns_i
        out_ra_rev[ptr]=RA_PERIOD_REV; out_capex[ptr]=capex_this_period
        ptr += 1

print(f"Simulation done in {time.time() - t0:.1f} s")
if USE_TOD:
    print(f"  Charge gated by lookahead: {_diag_charge_gated:,} periods ({_diag_charge_gated/total:.1%})")
    print(f"  SOC boost events:          {_diag_soc_boost:,} periods")
    print(f"  Inactivity discount:       {_diag_idle_relaxed:,} periods")


# ─────────────────────────────────────────────────────────────
# 7.  MONTHLY AGGREGATION  (identical to v2.4)
# ─────────────────────────────────────────────────────────────
print("Aggregating to monthly summary...")

df = pd.DataFrame({
    "year":   out_year,  "month":  out_month,
    "pl":     out_pl,    "action": out_action,
    "dmwh":   out_dmwh,  "adelta": out_adelta,  "meter":  out_meter,
    "soh":    out_soh,
    "ru_rev": out_ru_rev,"rd_rev": out_rd_rev,
    "sp_rev": out_sp_rev,"ns_rev": out_ns_rev,
    "ra_rev": out_ra_rev,"capex":  out_capex,
})

# Arb P&L: only genuine arb moves (adelta same sign as intended direction).
# Sign-flip periods (SOC clamp reversed direction) fall through to pl_h with hold P&L.
df["pl_c"] = np.where((df["action"] ==  1) & (df["adelta"] >  1e-6), df["pl"], 0.0)
df["pl_d"] = np.where((df["action"] == -1) & (df["adelta"] < -1e-6), df["pl"], 0.0)
df["pl_h"] = df["pl"] - df["pl_c"] - df["pl_d"]
df["n_c"]  = (df["action"] ==  1).astype(np.int32)
df["n_d"]  = (df["action"] == -1).astype(np.int32)
df["n_h"]  = (df["action"] ==  0).astype(np.int32)
# Arbitrage = voluntary dispatch; AS = action==0 but SOC actually moved (forced by reserves)
df["mwh_c"]    = np.where((df["action"] ==  1) & (df["adelta"] >  1e-6),  df["adelta"], 0.0)
df["mwh_d"]    = np.where((df["action"] == -1) & (df["adelta"] < -1e-6), -df["adelta"], 0.0)
# Sign-flip periods (arb action but SOC reversed direction) count as AS-forced corrections
df["mwh_c_as"] = np.where(
    ((df["action"] ==  0) | (df["action"] == -1)) & (df["adelta"] >  1e-6),
    df["adelta"], 0.0)
df["mwh_d_as"] = np.where(
    ((df["action"] ==  0) | (df["action"] ==  1)) & (df["adelta"] < -1e-6),
    -df["adelta"], 0.0)
df["seq"]   = np.arange(total)

monthly = (
    df.groupby(["year", "month"], sort=False)
      .agg(pl_total      =("pl",     "sum"),
           pl_charge     =("pl_c",   "sum"),
           pl_discharge  =("pl_d",   "sum"),
           pl_hold       =("pl_h",   "sum"),
           soh_end       =("soh",    "last"),
           n_charge      =("n_c",    "sum"),
           n_discharge   =("n_d",    "sum"),
           n_hold        =("n_h",    "sum"),
           mwh_charged   =("mwh_c",    "sum"),
           mwh_discharged=("mwh_d",    "sum"),
           mwh_c_as      =("mwh_c_as", "sum"),
           mwh_d_as      =("mwh_d_as", "sum"),
           _ru_rev       =("ru_rev", "sum"),
           _rd_rev       =("rd_rev", "sum"),
           _sp_rev       =("sp_rev", "sum"),
           _ns_rev       =("ns_rev", "sum"),
           _ra_rev       =("ra_rev", "sum"),
           _capex        =("capex",  "sum"),
           _first_seq    =("seq",    "min"))
      .reset_index()
      .sort_values("_first_seq")
      .reset_index(drop=True)
)
monthly.drop(columns=["_first_seq"], inplace=True)
monthly.insert(2, "model_month", np.arange(1, len(monthly) + 1))
monthly.columns = [
    "Year", "Month", "Model Month",
    "P/L Total ($)", "P/L Charge ($)", "P/L Discharge ($)", "P/L Hold ($)",
    "SOH (%)", "Charge Periods", "Discharge Periods", "Hold Periods",
    "MWh Charged (Arb)", "MWh Discharged (Arb)",
    "MWh Charged (AS)", "MWh Discharged (AS)",
    "_ru", "_rd", "_sp", "_ns", "_ra", "Retrofit CapEx ($)",
]

monthly["AS Revenue - RA ($)"]       = monthly["_ra"]
monthly["AS Revenue - Reg Up ($)"]   = monthly["_ru"]
monthly["AS Revenue - Reg Down ($)"] = monthly["_rd"]
monthly["AS Revenue - Spinning ($)"] = monthly["_sp"]
monthly["AS Revenue - Non-Spin ($)"] = monthly["_ns"]
monthly.drop(columns=["_ru","_rd","_sp","_ns","_ra"], inplace=True)
monthly["Total AS Revenue ($)"]    = (monthly["AS Revenue - RA ($)"]
                                     + monthly["AS Revenue - Reg Up ($)"]
                                     + monthly["AS Revenue - Reg Down ($)"]
                                     + monthly["AS Revenue - Spinning ($)"]
                                     + monthly["AS Revenue - Non-Spin ($)"])
monthly["Grand Total Revenue ($)"] = (monthly["P/L Total ($)"]
                                     + monthly["Total AS Revenue ($)"]
                                     - monthly["Retrofit CapEx ($)"])

import calendar as _cal
monthly["Days in Month"] = monthly.apply(
    lambda r: _cal.monthrange(int(r["Year"]), int(r["Month"]))[1], axis=1)
monthly["Cycles per Day"] = (monthly["MWh Discharged (Arb)"] + monthly["MWh Discharged (AS)"]) / NAMEPLATE_MWH / monthly["Days in Month"]
monthly.drop(columns=["Days in Month"], inplace=True)

monthly = monthly[[
    "Year", "Month", "Model Month",
    "P/L Total ($)", "P/L Charge ($)", "P/L Discharge ($)", "P/L Hold ($)",
    "SOH (%)", "Charge Periods", "Discharge Periods", "Hold Periods",
    "MWh Charged (Arb)", "MWh Discharged (Arb)",
    "MWh Charged (AS)", "MWh Discharged (AS)",
    "AS Revenue - RA ($)",
    "AS Revenue - Reg Up ($)", "AS Revenue - Reg Down ($)",
    "AS Revenue - Spinning ($)", "AS Revenue - Non-Spin ($)",
    "Total AS Revenue ($)", "Retrofit CapEx ($)", "Grand Total Revenue ($)",
    "Cycles per Day",
]]


# ─────────────────────────────────────────────────────────────
# 8.  WRITE OUTPUT EXCEL
# ─────────────────────────────────────────────────────────────
print(f"Writing {OUTPUT_PATH}...")

COL_SPECS = [
    ("Year",                      "0",       8),
    ("Month",                     "0",       8),
    ("Model Month",               "0",      13),
    ("P/L Total ($)",             '#,##0',  16),
    ("P/L Charge ($)",            '#,##0',  16),
    ("P/L Discharge ($)",         '#,##0',  18),
    ("P/L Hold ($)",              '#,##0',  14),
    ("SOH (%)",                   '0.00%',  10),
    ("Charge Periods",            "0",      15),
    ("Discharge Periods",         "0",      18),
    ("Hold Periods",              "0",      13),
    ("MWh Charged (Arb)",         '#,##0.0',16),
    ("MWh Discharged (Arb)",      '#,##0.0',18),
    ("MWh Charged (AS)",          '#,##0.0',16),
    ("MWh Discharged (AS)",       '#,##0.0',18),
    ("AS Revenue - RA ($)",       '#,##0',  18),
    ("AS Revenue - Reg Up ($)",   '#,##0',  18),
    ("AS Revenue - Reg Down ($)", '#,##0',  20),
    ("AS Revenue - Spinning ($)", '#,##0',  20),
    ("AS Revenue - Non-Spin ($)", '#,##0',  20),
    ("Total AS Revenue ($)",      '#,##0',  18),
    ("Retrofit CapEx ($)",        '#,##0',  18),
    ("Grand Total Revenue ($)",   '#,##0',  20),
    ("Cycles per Day",            '0.000',  14),
]

wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Monthly Summary"

_GRN  = PatternFill("solid", fgColor="375623")
_TEAL = PatternFill("solid", fgColor="1F6B55")
_STBL = PatternFill("solid", fgColor="4472C4")
COL_GROUPS = [
    ("Period",                        1,  3,  _DARK_BLUE),
    ("Arbitrage P&L",                 4,  7,  _MED_BLUE),
    ("Battery Operations",            8, 15,  _STBL),
    ("Ancillary Services Revenue",   16, 20,  _GRN),
    ("Revenue Totals",               21, 23,  _TEAL),
    ("Utilization",                  24, 24,  _GRN),
]
for label, c1, c2, fill in COL_GROUPS:
    ws_out.merge_cells(start_row=1, start_column=c1, end_row=1, end_column=c2)
    c = ws_out.cell(1, c1, label)
    _fmt(c, font=_WHITE_FONT, fill=fill, align=Alignment(horizontal="center", vertical="center"))
ws_out.row_dimensions[1].height = 18

def _group_fill(col_i):
    for _, c1, c2, fill in COL_GROUPS:
        if c1 <= col_i <= c2: return fill
    return _DARK_BLUE

for col_i, (hdr, fmt, width) in enumerate(COL_SPECS, start=1):
    c = ws_out.cell(2, col_i, hdr)
    _fmt(c, font=_WHITE_FONT, fill=_group_fill(col_i),
         align=Alignment(horizontal="center", vertical="center", wrap_text=True),
         border=_THIN_BDR)
    ws_out.column_dimensions[get_column_letter(col_i)].width = width
ws_out.row_dimensions[2].height = 30

alt_fill = PatternFill("solid", fgColor="EBF3FB")
for row_i, row_data in enumerate(monthly.itertuples(index=False), start=3):
    fill = alt_fill if row_i % 2 == 1 else None
    for col_i, (_, fmt, _w) in enumerate(COL_SPECS, start=1):
        val = row_data[col_i - 1]
        c   = ws_out.cell(row_i, col_i, float(val) if col_i == 8 else val)
        c.number_format = fmt
        c.font = _NORMAL
        c.alignment = Alignment(horizontal="right" if col_i > 3 else "center")
        c.border = _THIN_BDR
        if fill: c.fill = fill

ws_out.freeze_panes = "A3"
ws_out.auto_filter.ref = f"A2:{get_column_letter(len(COL_SPECS))}2"
wb_out.save(OUTPUT_PATH)


# ─────────────────────────────────────────────────────────────
# 9.  CONSOLE SUMMARY
# ─────────────────────────────────────────────────────────────
total_pl    = monthly["P/L Total ($)"].sum()
total_as    = monthly["Total AS Revenue ($)"].sum()
total_grand = monthly["Grand Total Revenue ($)"].sum()
avg_soh     = monthly["SOH (%)"].mean()
final_soh   = monthly["SOH (%)"].iloc[-1]

print(f"\n{'-'*57}")
print(f"  Dispatch mode  : {'TOD (time-of-day)' if USE_TOD else 'Spread table (legacy)'}")
print(f"  Output         : {OUTPUT_PATH}")
print(f"  Months         : {len(monthly)}")
total_charge = monthly["P/L Charge ($)"].sum()
total_dis    = monthly["P/L Discharge ($)"].sum()
total_hold   = monthly["P/L Hold ($)"].sum()
print(f"  Arbitrage P/L  : ${total_pl:>14,.0f}")
print(f"    Charge       : ${total_charge:>14,.0f}")
print(f"    Discharge    : ${total_dis:>14,.0f}")
print(f"    Hold losses  : ${total_hold:>14,.0f}")
if total_as != 0:
    print(f"  AS Revenue     : ${total_as:>14,.0f}")
    print(f"    RA           : ${monthly['AS Revenue - RA ($)'].sum():>14,.0f}")
    print(f"    Reg Up       : ${monthly['AS Revenue - Reg Up ($)'].sum():>14,.0f}")
    print(f"    Reg Down     : ${monthly['AS Revenue - Reg Down ($)'].sum():>14,.0f}")
    print(f"    Spinning     : ${monthly['AS Revenue - Spinning ($)'].sum():>14,.0f}")
    print(f"    Non-Spin     : ${monthly['AS Revenue - Non-Spin ($)'].sum():>14,.0f}")
total_capex = monthly["Retrofit CapEx ($)"].sum()
if total_capex != 0:
    print(f"  Retrofit CapEx : ${-total_capex:>14,.0f}")
if total_as != 0 or total_capex != 0:
    print(f"  Grand Total    : ${total_grand:>14,.0f}")
print(f"  Avg SOH        : {avg_soh:.1%}")
print(f"  Final SOH      : {final_soh:.1%}")
if retrofit_log:
    print(f"  Retrofits      : {len(retrofit_log)} event(s)")
    for ry, rm, rcx in retrofit_log:
        print(f"    {ry}-{rm:02d}: ${rcx:,.0f}, offline {RETROFIT_DAYS:.0f} days, SOH -> {SOH_AFTER_RETROFIT:.0%}")
total_mwh_chg = monthly["MWh Charged (Arb)"].sum() + monthly["MWh Charged (AS)"].sum()
total_mwh_dis = monthly["MWh Discharged (Arb)"].sum() + monthly["MWh Discharged (AS)"].sum()
avg_cpd = monthly["Cycles per Day"].mean()
print(f"  MWh Charged    : {total_mwh_chg:>12,.0f}  (Arb: {monthly['MWh Charged (Arb)'].sum():,.0f}  AS: {monthly['MWh Charged (AS)'].sum():,.0f})")
print(f"  MWh Discharged : {total_mwh_dis:>12,.0f}  (Arb: {monthly['MWh Discharged (Arb)'].sum():,.0f}  AS: {monthly['MWh Discharged (AS)'].sum():,.0f})")
print(f"  Avg Cycles/Day : {avg_cpd:.3f}")
print(f"  Wall-clock     : {time.time() - t0:.1f} s")
print(f"{'-'*57}")
print("\nFirst 6 months:")
pd.set_option("display.float_format", "{:,.0f}".format)
print(monthly.head(6)[
    ["Year", "Month", "P/L Total ($)", "P/L Discharge ($)",
     "SOH (%)", "Charge Periods", "Discharge Periods",
     "MWh Charged (Arb)", "MWh Discharged (Arb)", "MWh Charged (AS)", "MWh Discharged (AS)",
     "Cycles per Day"]
].to_string(index=False))

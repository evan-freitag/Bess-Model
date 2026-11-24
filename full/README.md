# Project Friday: BESS revenue and project finance model

A 30 × Tesla Megapack 2 XL battery (29.37 MW / 117.48 MWh) in CAISO, node ADCC_2_N001. There are two parts:

- **`Project Friday Model 1.0.5.xlsm`** is the Excel project finance model. It holds the inputs, scenarios, cash flow, debt, tax and returns.
- **`code_v2.py`** is the Python dispatch backtest. Every 5 minutes it solves a 48-hour optimisation over real CAISO prices, choosing energy arbitrage, Regulation Up/Down and Spin around the Resource Adequacy commitments. It covers July 2023 to June 2026, which is 315,048 steps. Its monthly results feed the workbook's revenue.

This is a demonstration model, not an investment recommendation. Placeholder assumptions are noted in the workbook.

## Files you need

Keep these four in the **same folder**:

| File | What it is |
|---|---|
| `Project Friday Model 1.0.5.xlsm` | The model, with the macros that run and import the Python results |
| `code_v2.py` | The dispatch backtest |
| `caiso_energy_and_as_5min.csv` | CAISO 5-minute energy prices and hourly AS prices, 2023-07 to 2026-06 |
| `requirements.txt` | Python packages |

`code_v2.py` writes three `*_v2` files when it finishes, and Import Python Run reads all three:

- `storage_arbitrage_monthly_summary_v2.csv`: monthly results, into `Revenue Data`
- `storage_arbitrage_backtest_5min_v2.csv`: 5-minute price, charge, discharge and state of charge, into `Revenue Data 5 Minute`, which the Charts tab reads
- `dispatch_run_inputs_v2.csv`: the inputs the run used and its totals

Double-clicking `code_v2.py` reads the most recently saved `Project Friday Model *.xlsm` in its folder, so a new version needs no code change.

## Setup (once)

You need **Windows** and **Excel for Windows**. The macros don't run on a Mac.

1. **Install Python 3.11 or later** from [python.org](https://www.python.org/downloads/windows/). The installer adds the `py` launcher.
2. **Install the packages.** Open a Command Prompt in this folder and run:
   ```
   py -m pip install -r requirements.txt
   ```
3. **Unblock the workbook.** Windows blocks macros in downloaded files: right-click `Project Friday Model 1.0.5.xlsm`, choose **Properties**, tick **Unblock**, then OK. Without this, Excel shows a red "Microsoft has blocked macros" bar and the buttons do nothing.
4. Open the workbook and click **Enable Content** if Excel asks.

## Running it

On the workbook's **Python Run** tab:

1. **Run Python.** It finds a Python that has the packages, saves a snapshot of the workbook to your temp folder, and starts `code_v2.py` in its own window. That window shows progress and says **Finished** when done.
2. **Import Python Run.** It loads the monthly results into `Revenue Data` and the 5-minute results into `Revenue Data 5 Minute`, then records which inputs the run used. Before overwriting, it checks all three files belong to one complete run and shows the before/after totals. It takes about 10 seconds.

The status line at the top of the tab says when an input that Python reads has changed since the last run:

- plant size
- RTE
- RA growth
- the arbitrage floor and ceiling
- RA sustain hours
- the RA grid

When that happens, run again. Every other input, including all the scenario cases, works in Excel without a re-run.

You can also double-click `code_v2.py`, but **close the workbook first**. Excel with OneDrive AutoSave locks an open workbook so nothing else can read it, which is why Run Python works from a snapshot.

## Run time and results

- The backtest splits the 3 years into 79 two-week chunks and runs them in parallel on all but two CPU cores. That took about **10 minutes on 24 cores**. Expect 25–35 minutes on a typical 8-core laptop, and about 70 minutes on 1–3 cores.
- **Results don't depend on the number of cores.** Each chunk starts 3 days early from a guessed state of charge, and at every seam its path is checked against the chunk before. Any chunk that hasn't converged is re-run from the true state, so the stitched result equals one long sequential run.
- Each solve starts clean, so the same inputs give the same results on every run. A different scipy version can break exact ties (equal-value actions) differently. That can move individual 5-minute steps without changing total revenue.

## Troubleshooting

| What you see | Cause and fix |
|---|---|
| Red "blocked macros" bar, buttons do nothing | The file came from the internet: do the Unblock step in Setup. |
| "No Python with the packages … was found" | Do Setup steps 1–2. Run Python tries `py -3`, `python`, `py` and `python3` in turn. |
| Python window shows an error | The message stays on screen and is saved to `code_v2_crash_log.txt` in this folder. A missing package prints the exact `pip install` command for that Python. |
| `PermissionError` on the workbook | The workbook was open while `code_v2.py` was double-clicked. Close it, or use Run Python. |
| Run Python can't find the files on a SharePoint or Teams folder | The macros map personal and work OneDrive web addresses to the synced folder. A SharePoint library synced elsewhere isn't covered, so copy the set to a normal folder. |

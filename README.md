# Project Friday: BESS revenue and project finance model

A 30 × Tesla Megapack 2 XL battery (29.37 MW / 117.48 MWh) in CAISO, node ADCC_2_N001. Revenue comes from a Python dispatch backtest:

- every 5 minutes over July 2023 to June 2026
- energy arbitrage, Regulation Up/Down and Spin
- around Resource Adequacy commitments

The Excel model turns that revenue into a 24-year project finance case: monthly cash flow, quarterly senior debt, tax and ITC sale, tax equity, eight scenario cases and integrity checks.

This is a demonstration model, not an investment recommendation.

## Pick a version

| Folder | What you get | Needs |
|---|---|---|
| [`excel-only/`](excel-only) | The full model as a plain `.xlsx`: every tab, results for the saved case | Excel only, no macros |
| [`excel-with-macros/`](excel-with-macros) | The same model with Load buttons to switch between the eight scenario cases | Excel with macros enabled |
| [`full/`](full) | Everything: the model, `code_v2.py` and the CAISO price data, so you can re-run the dispatch backtest from Excel | Windows, Excel with macros, Python 3.11+ (see [`full/README.md`](full/README.md)) |

All three hold the same model and give the same results. They differ only in what can be changed:

- **excel-only:** a fixed snapshot.
- **excel-with-macros:** can switch between the eight scenario cases.
- **full:** can also change dispatch inputs (plant size, RTE, the RA schedule) and re-run Python.

## Macro versions

Windows blocks macros in downloaded files. Right-click the `.xlsm`, choose **Properties**, tick **Unblock**, then OK, and click **Enable Content** when Excel opens it.

## Version

1.0.5. The workbook's **Change Log** tab lists changes from 1.0 onwards.

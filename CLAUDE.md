# pfanal

Reports on a Fidelity portfolio-positions export: lists each position by maturity date with days to maturity, present value and value at maturity, plus totals.

Public repo: https://github.com/agoldhammer/pfanal (branch `master`).

## Commands

- `uv run pfanal` — run the report (reads from `data/`, so run from the repo root)
- `uv run pytest` — run tests

Python >= 3.15, managed with uv. Only runtime dependency: openpyxl.

## How it works

All code is in `src/pfanal/__init__.py`.

- **Input:** the newest `.xlsx` in `data/`. "Newest" means the date in the filename (e.g. `Portfolio_Positions_Oct-17-2026.xlsx`), or the file's modification time if the name has no date. Excel lock files (`~$...`) are skipped.
- **Columns used:** B account name, C symbol, D description, E quantity, H current value, N cost basis total.
- **Maturity date:** the trailing `mm/dd/yyyy` at the end of the description. Rows without one are skipped, which drops blank rows and Fidelity's disclaimer footer.
- **Cost basis:** the "Cost basis total" column.
- **Present value:** the "Current value" column.
- **Gain to date:** (present value − cost basis) as a percentage of cost basis. The Total row uses the summed columns, so it is weighted by cost basis.
- **Value at maturity:** the "Quantity" column (face amount). This is correct only for zero-coupon instruments such as T-bills, which every current holding is. Coupon-paying notes or bonds would need their final coupon added.
- **Output:** rows sorted by maturity, then a separator line and a Total row summing cost basis, present value and value at maturity, plus the overall gain percentage.

## Data privacy

The repo is public. `data/`, `*.xlsx` and `*.csv` are gitignored and must stay that way. Never commit real account numbers, account names or holdings, including in tests. Tests build synthetic spreadsheets in `tmp_path`.

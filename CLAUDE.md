# pfanal

Reports on a Fidelity portfolio-positions export: lists each position by maturity date with days to maturity, present value and value at maturity, plus totals.

Public repo: https://github.com/agoldhammer/pfanal (branch `master`).

## Commands

- `pfanal` — run the report from any directory (installed as an editable uv tool, see below)
- `uv run pfanal` — run the report from the repo without the tool install
- `uv run pytest` — run tests

Installed with `uv tool install -e ~/Prog/pfanal`, which puts `pfanal` in `~/.local/bin`. Editable mode runs the repo's code directly, so code edits take effect immediately and `data/` resolves to the repo. After changing dependencies in `pyproject.toml`, reinstall with `uv tool install -e --reinstall ~/Prog/pfanal`. A non-editable install would break the `data/` fallback, since `DATA_DIR` is derived from the package location.

Python >= 3.15, managed with uv. Only runtime dependency: openpyxl.

## How it works

All code is in `src/pfanal/__init__.py`.

- **Input:** the newest Fidelity export (`Portfolio_Positions*.xlsx` or `.csv`) in the Windows Downloads directory, found via `cmd.exe` and `wslpath`. Other files in Downloads are ignored. If there is none, or this isn't WSL, it uses the newest `.xlsx` or `.csv` of any name in the repo's `data/`, an absolute path derived from the package location, so the working directory doesn't matter. Fidelity's CSV export is read directly, so no conversion is needed. "Newest" means the date in the filename (e.g. `Portfolio_Positions_Oct-17-2026.csv`), or the file's modification time if the name has no date, which also breaks ties. Excel lock files (`~$...`) are skipped. CSV numbers are strings such as `$99594.00` or `+$0.011`, and `parse_number` converts them. If no rows have a maturity date, the program exits with an error.
- **Columns used:** B account name, C symbol, D description, E quantity, H current value, N cost basis total.
- **Maturity date:** the trailing `mm/dd/yyyy` at the end of the description. Rows without one are skipped, which drops blank rows and Fidelity's disclaimer footer.
- **Cost basis:** the "Cost basis total" column. Fidelity sometimes reports `--` (no cost basis). Those positions show `--` for cost basis and gain, and are left out of the Total cost basis and the total gain percentage.
- **Present value:** face amount (Quantity) × the CUSIP's price from TreasuryDirect's FedInvest price page (`PRICE_URL`), using the latest day on or before today that has prices published. A day's page goes up during that day with end-of-day prices of 0, and they stay 0 until the next business day (so Friday's are still 0 over the weekend). For those, the sell price is used instead, and the `Prices:` header line says `sell` rather than `end of day`. A row with both prices 0 is ignored, and a page with no usable prices is skipped. The page is a CSRF-protected form, so `treasurydirect_page` does a GET for the token and cookie, then a POST. Positions with no quoted CUSIP keep the spreadsheet's "Current value" and are marked `*`. If the fetch fails, every position falls back to the spreadsheet with a warning. The whole day's price list is downloaded, so no CUSIPs are sent to the site.
- **Gain to date:** (present value − cost basis) as a percentage of cost basis. The Total row uses the summed columns, so it is weighted by cost basis.
- **Value at maturity:** the "Quantity" column (face amount). This is correct only for zero-coupon instruments such as T-bills, which every current holding is. Coupon-paying notes or bonds would need their final coupon added.
- **Output:** rows sorted by maturity, then a separator line and a Total row summing cost basis, present value and value at maturity, plus the overall gain percentage.

## Maturity emails

`maturing.py` (repo root, outside the package) reads the newest `Portfolio_Positions*.csv` in Windows Downloads (CSV only, unlike `pfanal`), falling back to the newest `.csv` in `data/` with a warning, and emails one message listing each position maturing on the day it runs: maturity date, amount (face value), account name and CUSIP. If nothing matures, it sends nothing. It sends by running the Gmail API helper `~/Prog/scripts/gmail/gmail_send.py` (OAuth `token.json` there, no password; shared with the poll-update jobs) through `uv run`, which keeps the Google libraries out of pfanal's dependencies. The recipient is `MAIL_TO`, which `--to` overrides. `--dry-run` prints the message without sending it, and `--date YYYY-MM-DD` checks another day. The user's crontab runs it at 9:00 on weekdays via `uv run --project`, logging to `~/.local/state/pfanal/maturing.log`. The crontab sets `PATH` to include `/mnt/c/Windows/system32`, since cron's default `PATH` has no `cmd.exe` and the Downloads lookup would fail. Cron runs only while WSL is up.

## Data privacy

The repo is public. `data/`, `*.xlsx` and `*.csv` are gitignored and must stay that way. Never commit real account numbers, account names or holdings, including in tests. Tests build synthetic spreadsheets in `tmp_path`.

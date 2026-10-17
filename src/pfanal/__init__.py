"""List positions from the newest portfolio spreadsheet, sorted by maturity date."""

import re
import sys
from datetime import date, datetime
from pathlib import Path
from typing import NamedTuple

from openpyxl import load_workbook

DATA_DIR = Path("data")

# A mm/dd/yyyy date at the very end of the description (column D).
MATURITY_RE = re.compile(r"(\d{2}/\d{2}/\d{4})\s*$")
# Date embedded in export filenames, e.g. Portfolio_Positions_Oct-17-2026.xlsx
FILENAME_DATE_RE = re.compile(r"([A-Z][a-z]{2}-\d{2}-\d{4})")


class Position(NamedTuple):
    maturity: date
    account_name: str
    symbol: str


def file_date(path: Path) -> datetime:
    """Date a spreadsheet was exported: from its filename if present, else its mtime."""
    if m := FILENAME_DATE_RE.search(path.stem):
        try:
            return datetime.strptime(m.group(1), "%b-%d-%Y")
        except ValueError:
            pass
    return datetime.fromtimestamp(path.stat().st_mtime)


def latest_spreadsheet(data_dir: Path = DATA_DIR) -> Path:
    files = [p for p in data_dir.glob("*.xlsx") if not p.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"no .xlsx files found in {data_dir}/")
    return max(files, key=lambda p: (file_date(p), p.stat().st_mtime))


def parse_maturity(description: object) -> date | None:
    """Return the trailing mm/dd/yyyy date in a description, or None if absent."""
    if not isinstance(description, str):
        return None
    if m := MATURITY_RE.search(description):
        return datetime.strptime(m.group(1), "%m/%d/%Y").date()
    return None


def read_positions(path: Path) -> list[Position]:
    ws = load_workbook(path, read_only=True, data_only=True).active
    positions = []
    # Columns: A=account number, B=account name, C=symbol, D=description.
    for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        _, account_name, symbol, description = row
        # Skips blank rows and the disclaimer footer, which have no maturity date.
        if (maturity := parse_maturity(description)) is None:
            continue
        positions.append(Position(maturity, str(account_name), str(symbol)))
    return sorted(positions, key=lambda p: p.maturity)


def days_until(maturity: date, today: date | None = None) -> int:
    """Calendar days from today until maturity (negative if already matured)."""
    return (maturity - (today or date.today())).days


def main() -> None:
    try:
        path = latest_spreadsheet()
    except FileNotFoundError as e:
        sys.exit(f"error: {e}")
    print(f"Source: {path}\n")
    positions = read_positions(path)
    width = max((len(p.account_name) for p in positions), default=12)
    today = date.today()
    print(f"{'Maturity':<10}  {'Days':>4}  {'Account name':<{width}}  Symbol")
    for p in positions:
        print(
            f"{p.maturity:%m/%d/%Y}  {days_until(p.maturity, today):>4}  "
            f"{p.account_name:<{width}}  {p.symbol}"
        )

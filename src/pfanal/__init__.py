"""List positions from the newest portfolio spreadsheet, sorted by maturity date."""

import http.cookiejar
import re
import sys
import urllib.parse
import urllib.request
from collections.abc import Callable
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import NamedTuple

from openpyxl import load_workbook

DATA_DIR = Path("data")

# A mm/dd/yyyy date at the very end of the description (column D).
MATURITY_RE = re.compile(r"(\d{2}/\d{2}/\d{4})\s*$")
# Date embedded in export filenames, e.g. Portfolio_Positions_Oct-17-2026.xlsx
FILENAME_DATE_RE = re.compile(r"([A-Z][a-z]{2}-\d{2}-\d{4})")

# TreasuryDirect's daily prices for all marketable Treasury securities, by CUSIP.
PRICE_URL = "https://www.treasurydirect.gov/GA-FI/FedInvest/selectSecurityPriceDate"
# Prices are published once per business day, so look back far enough to
# cover a weekend plus a holiday.
PRICE_LOOKBACK_DAYS = 7
PRICE_ROW_RE = re.compile(r"<tr>(.*?)</tr>", re.S)
PRICE_CELL_RE = re.compile(r"<td[^>]*>(.*?)</td>", re.S)


class Position(NamedTuple):
    maturity: date
    account_name: str
    symbol: str
    cost_basis: float
    present_value: float
    maturity_value: float
    priced_online: bool = False

    @property
    def gain(self) -> float:
        return self.present_value - self.cost_basis

    @property
    def gain_pct(self) -> float:
        """Gain to date as a percentage of cost basis."""
        return self.gain / self.cost_basis * 100


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
    # Columns: A=account number, B=account name, C=symbol, D=description,
    # E=quantity, F=last price, G=last price change, H=current value,
    # I-M=gain/loss and percent of account, N=cost basis total.
    for row in ws.iter_rows(min_row=2, max_col=14, values_only=True):
        _, account_name, symbol, description, quantity, _, _, current_value = row[:8]
        cost_basis = row[13]
        # Skips blank rows and the disclaimer footer, which have no maturity date.
        if (maturity := parse_maturity(description)) is None:
            continue
        # Bills are zero-coupon, so face value (quantity) is what's paid at maturity.
        positions.append(
            Position(
                maturity,
                str(account_name),
                str(symbol),
                float(cost_basis),
                float(current_value),
                float(quantity),
            )
        )
    return sorted(positions, key=lambda p: p.maturity)


def parse_price_table(html: str) -> dict[str, float]:
    """Map CUSIP to end-of-day price (per 100 face) from a TreasuryDirect price page."""
    prices = {}
    # Cells: CUSIP, security type, rate, maturity, call date, buy, sell, end of day.
    for row in PRICE_ROW_RE.findall(html):
        cells = [c.strip() for c in PRICE_CELL_RE.findall(row)]
        if len(cells) == 8:
            prices[cells[0]] = float(cells[7])
    return prices


def treasurydirect_page(day: date) -> str:
    """Fetch TreasuryDirect's price page for a day (its table is empty if unpublished)."""
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar())
    )
    opener.addheaders = [("User-Agent", "Mozilla/5.0")]
    # The form needs the CSRF token and session cookie from a prior GET.
    form = opener.open(PRICE_URL, timeout=30).read().decode()
    token = re.search(r'name="_csrf" value="([^"]*)"', form).group(1)
    data = urllib.parse.urlencode(
        {"priceDate": day.isoformat(), "_csrf": token, "submit": "Show Prices"}
    ).encode()
    return opener.open(PRICE_URL, data, timeout=30).read().decode()


def fetch_prices(
    today: date, get_page: Callable[[date], str] = treasurydirect_page
) -> tuple[date, dict[str, float]]:
    """Prices from the most recent day on or before today that has them published."""
    for back in range(PRICE_LOOKBACK_DAYS):
        day = today - timedelta(days=back)
        if prices := parse_price_table(get_page(day)):
            return day, prices
    raise LookupError(f"no Treasury prices published in the {PRICE_LOOKBACK_DAYS} days to {today}")


def reprice(positions: list[Position], prices: dict[str, float]) -> list[Position]:
    """Value each position at its quoted price; keep the spreadsheet value if unquoted."""
    return [
        p._replace(present_value=p.maturity_value * prices[p.symbol] / 100, priced_online=True)
        if p.symbol in prices
        else p
        for p in positions
    ]


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
    today = date.today()
    try:
        price_date, prices = fetch_prices(today)
    except (OSError, LookupError, AttributeError) as e:
        print(f"warning: could not fetch Treasury prices ({e}); "
              "using spreadsheet values\n", file=sys.stderr)
    else:
        print(f"Prices: TreasuryDirect end of day, {price_date:%m/%d/%Y}\n")
        positions = reprice(positions, prices)
    width = max((len(p.account_name) for p in positions), default=12)
    print(
        f"{'Maturity':<10}  {'Days':>4}  {'Account name':<{width}}  {'Symbol':<9}  "
        f"{'Cost basis':>15}  {'Present value':>15}  {'Gain to date':>12}  "
        f"{'Value at maturity':>17}"
    )
    for p in positions:
        print(
            f"{p.maturity:%m/%d/%Y}  {days_until(p.maturity, today):>4}  "
            f"{p.account_name:<{width}}  {p.symbol:<9}  "
            f"{p.cost_basis:>15,.2f}  {p.present_value:>15,.2f}  {p.gain_pct:>11.3f}%  "
            f"{p.maturity_value:>17,.2f}{'' if p.priced_online else '  *'}"
        )
    total_cb = sum(p.cost_basis for p in positions)
    total_pv = sum(p.present_value for p in positions)
    total_mv = sum(p.maturity_value for p in positions)
    label_width = 10 + 2 + 4 + 2 + width + 2 + 9
    print("-" * (label_width + 2 + 15 + 2 + 15 + 2 + 12 + 2 + 17))
    print(
        f"{'Total':<{label_width}}  {total_cb:>15,.2f}  {total_pv:>15,.2f}  "
        f"{(total_pv - total_cb) / total_cb * 100:>11.3f}%  {total_mv:>17,.2f}"
    )
    if not all(p.priced_online for p in positions):
        print("\n* present value from the spreadsheet (no online price found)")

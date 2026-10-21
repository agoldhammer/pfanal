from datetime import date

import pytest

from pfanal import days_until, parse_maturity


def test_parses_trailing_date():
    desc = "UNITED STATES TREAS BILLS ZERO CPN 0.00000% 11/27/2026"
    assert parse_maturity(desc) == date(2026, 11, 27)


def test_ignores_non_trailing_date():
    assert parse_maturity("Issued 01/02/2026 SOMETHING") is None


def test_ignores_non_strings():
    assert parse_maturity(None) is None
    assert parse_maturity(42) is None


def test_days_until():
    today = date(2026, 10, 17)
    assert days_until(date(2026, 11, 10), today) == 24
    assert days_until(date(2027, 1, 7), today) == 82
    assert days_until(today, today) == 0
    assert days_until(date(2026, 10, 1), today) == -16


def test_read_positions_values(tmp_path):
    from openpyxl import Workbook

    from pfanal import read_positions

    wb = Workbook()
    ws = wb.active
    ws.append(["Account number", "Account name", "Symbol", "Description",
               "Quantity", "Last price", "Last price change", "Current value",
               "Today's gain/loss dollar", "Today's gain/loss percent",
               "Total gain/loss dollar", "Total gain/loss percent",
               "Percent of account", "Cost basis total"])
    ws.append([1, "IRA", "912797WS2", "TREAS BILLS 0.00000% 02/02/2027",
               100000, 98.79, 0.041, 98790, 41, 0.04, 790, 0.81, 10, 98000])
    ws.append([1, "IRA", "912797SU2", "TREAS BILLS 0.00000% 11/27/2026",
               50000, 99.571, 0.035, 49785.5, 17.5, 0.04, 85.5, 0.17, 5, 49700])
    ws.append(["Disclaimer footer text"])
    path = tmp_path / "p.xlsx"
    wb.save(path)

    positions = read_positions(path)
    assert [p.symbol for p in positions] == ["912797SU2", "912797WS2"]
    assert positions[0].cost_basis == 49700
    assert positions[0].gain == 85.5
    assert positions[1].cost_basis == 98000
    assert positions[1].gain == 790
    assert positions[0].gain_pct == pytest.approx(85.5 / 49700 * 100)
    assert positions[1].gain_pct == pytest.approx(790 / 98000 * 100)
    assert positions[0].present_value == 49785.5
    assert positions[0].maturity_value == 50000
    assert positions[1].present_value == 98790
    assert positions[1].maturity_value == 100000


PRICE_PAGE = """
<p>Prices For: October 19, 2026</p>
<table>
\t<tr><th>CUSIP</th><th>SECURITY TYPE</th><th>RATE</th><th>MATURITY DATE</th><th>CALL DATE</th><th>BUY</th><th>SELL</th><th>END OF DAY</th></tr>
\t<tr>
\t\t<td>912797SU2</td>
\t\t<td style="white-space: nowrap; text-align: left">MARKET BASED BILL</td>
\t\t<td>0.000%</td>
\t\t<td>11/27/2026</td>
\t\t<td></td>
\t\t<td>99.573167</td>
\t\t<td>99.572083</td>
\t\t<td>99.583056</td>
\t</tr>
\t<tr>
\t\t<td>912797WS2</td>
\t\t<td style="white-space: nowrap; text-align: left">MARKET BASED BILL</td>
\t\t<td>0.000%</td>
\t\t<td>02/02/2027</td>
\t\t<td></td>
\t\t<td>0.000000</td>
\t\t<td>98.800000</td>
\t\t<td>98.810000</td>
\t</tr>
</table>
"""


def test_parse_price_table():
    from pfanal import parse_price_table

    assert parse_price_table(PRICE_PAGE) == {
        "912797SU2": 99.583056,
        "912797WS2": 98.81,
    }


def test_parse_price_table_empty():
    from pfanal import parse_price_table

    assert parse_price_table("<table><tr><th>CUSIP</th></tr></table>") == {}


def test_fetch_prices_steps_back_to_latest_published_day():
    from pfanal import fetch_prices

    requested = []

    def get_page(day):
        requested.append(day)
        return PRICE_PAGE if day == date(2026, 10, 16) else ""

    price_date, prices = fetch_prices(date(2026, 10, 18), get_page)
    assert price_date == date(2026, 10, 16)
    assert prices["912797SU2"] == 99.583056
    assert requested == [date(2026, 10, 18), date(2026, 10, 17), date(2026, 10, 16)]


# Today's page during the trading day: buy/sell quoted, end of day not yet set.
INTRADAY_PAGE = PRICE_PAGE.replace("99.583056", "0.000000").replace("98.810000", "0.000000")


def test_parse_price_table_skips_zero_end_of_day():
    from pfanal import parse_price_table

    assert parse_price_table(INTRADAY_PAGE) == {}


def test_fetch_prices_skips_day_without_end_of_day_prices():
    from pfanal import fetch_prices

    def get_page(day):
        return INTRADAY_PAGE if day == date(2026, 10, 20) else PRICE_PAGE

    price_date, prices = fetch_prices(date(2026, 10, 20), get_page)
    assert price_date == date(2026, 10, 19)
    assert prices["912797SU2"] == 99.583056


def test_fetch_prices_gives_up():
    from pfanal import fetch_prices

    with pytest.raises(LookupError):
        fetch_prices(date(2026, 10, 18), lambda day: "")


def test_reprice_uses_quoted_price_and_keeps_unquoted():
    from pfanal import Position, reprice

    quoted = Position(date(2026, 11, 27), "IRA", "912797SU2", 49700, 49785.5, 50000)
    unquoted = Position(date(2027, 3, 1), "IRA", "912797ZZ9", 9800, 9850, 10000)
    repriced = reprice([quoted, unquoted], {"912797SU2": 99.583056})

    assert repriced[0].present_value == pytest.approx(50000 * 99.583056 / 100)
    assert repriced[0].priced_online
    assert repriced[1].present_value == 9850
    assert not repriced[1].priced_online


CSV_EXPORT = (
    "﻿Account number,Account name,Symbol,Description,Quantity,Last price,"
    "Last price change,Current value,Today's gain/loss dollar,"
    "Today's gain/loss percent,Total gain/loss dollar,Total gain/loss percent,"
    "Percent of account,Cost basis total,Average cost basis,Type\n"
    "1,IRA,912797WS2,TREAS BILLS 0.00000% 02/02/2027,100000,$98.79,+$0.041,"
    "$98790.00,$41.00,0.04%,+$790.00,+0.81%,10.00%,$98000.00,--,Cash,\n"
    "2,Joint,912797SU2,TREAS BILLS 0.00000% 11/27/2026,10000,$99.594,+$0.011,"
    "$9959.40,$0.00,0.00%,--,--,5.00%,--,--,Cash,\n"
    "\n"
    '"Disclaimer footer text"\n'
)


def test_read_positions_from_csv(tmp_path):
    from pfanal import read_positions

    path = tmp_path / "p.csv"
    path.write_text(CSV_EXPORT, encoding="utf-8")

    positions = read_positions(path)
    assert [p.symbol for p in positions] == ["912797SU2", "912797WS2"]
    assert positions[1].account_name == "IRA"
    assert positions[1].cost_basis == 98000
    assert positions[1].present_value == 98790
    assert positions[1].maturity_value == 100000
    assert positions[1].gain == 790


def test_missing_cost_basis_has_no_gain(tmp_path):
    from pfanal import read_positions

    path = tmp_path / "p.csv"
    path.write_text(CSV_EXPORT, encoding="utf-8")

    no_basis = read_positions(path)[0]
    assert no_basis.cost_basis is None
    assert no_basis.gain is None
    assert no_basis.gain_pct is None
    assert no_basis.present_value == 9959.40


def test_parse_number():
    from pfanal import parse_number

    assert parse_number("$99,594.00") == 99594
    assert parse_number("+$0.011") == 0.011
    assert parse_number("-$1.50") == -1.5
    assert parse_number("--") is None
    assert parse_number("") is None
    assert parse_number(None) is None
    assert parse_number(42) == 42


def test_latest_spreadsheet_includes_csv(tmp_path):
    from pfanal import latest_spreadsheet

    (tmp_path / "Portfolio_Positions_Oct-17-2026.xlsx").touch()
    (tmp_path / "Portfolio_Positions_Oct-21-2026.csv").touch()
    assert latest_spreadsheet(tmp_path).name == "Portfolio_Positions_Oct-21-2026.csv"


def test_main_exits_cleanly_when_no_positions(tmp_path, monkeypatch):
    import pfanal

    (tmp_path / "empty.csv").write_text("Account number,Account name\n")
    monkeypatch.setattr(pfanal, "windows_downloads", lambda: None)
    monkeypatch.setattr(pfanal.latest_spreadsheet, "__defaults__", (tmp_path,))
    with pytest.raises(SystemExit, match="no positions"):
        pfanal.main()


def test_find_spreadsheet_prefers_downloads_export(tmp_path, monkeypatch):
    import pfanal

    downloads, data = tmp_path / "Downloads", tmp_path / "data"
    downloads.mkdir()
    data.mkdir()
    (data / "Portfolio_Positions_Oct-21-2026.csv").touch()
    (downloads / "bank_statement.csv").touch()
    monkeypatch.setattr(pfanal, "windows_downloads", lambda: downloads)
    monkeypatch.setattr(pfanal.latest_spreadsheet, "__defaults__", (data,))
    # Only Fidelity exports count in Downloads, so this falls back to data/.
    assert pfanal.find_spreadsheet() == data / "Portfolio_Positions_Oct-21-2026.csv"

    (downloads / "Portfolio_Positions_Oct-17-2026.csv").touch()
    assert pfanal.find_spreadsheet() == downloads / "Portfolio_Positions_Oct-17-2026.csv"


def test_find_spreadsheet_falls_back_without_wsl(tmp_path, monkeypatch):
    import pfanal

    (tmp_path / "positions.xlsx").touch()
    monkeypatch.setattr(pfanal, "windows_downloads", lambda: None)
    monkeypatch.setattr(pfanal.latest_spreadsheet, "__defaults__", (tmp_path,))
    assert pfanal.find_spreadsheet() == tmp_path / "positions.xlsx"


def test_data_dir_is_absolute():
    from pfanal import DATA_DIR

    assert DATA_DIR.is_absolute() and DATA_DIR.name == "data"

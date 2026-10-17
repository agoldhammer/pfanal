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

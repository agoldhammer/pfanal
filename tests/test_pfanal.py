from datetime import date

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

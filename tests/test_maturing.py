import subprocess
from datetime import date

import pytest

import maturing

CSV_EXPORT = (
    "﻿Account number,Account name,Symbol,Description,Quantity,Last price,"
    "Last price change,Current value,Today's gain/loss dollar,"
    "Today's gain/loss percent,Total gain/loss dollar,Total gain/loss percent,"
    "Percent of account,Cost basis total,Average cost basis,Type\n"
    "1,IRA,912797WS2,TREAS BILLS 0.00000% 02/02/2027,100000,$98.79,+$0.041,"
    "$98790.00,$41.00,0.04%,+$790.00,+0.81%,10.00%,$98000.00,--,Cash,\n"
    "2,Joint,912797SU2,TREAS BILLS 0.00000% 11/27/2026,10000,$99.594,+$0.011,"
    "$9959.40,$0.00,0.00%,--,--,5.00%,--,--,Cash,\n"
    "1,IRA,912797SU2,TREAS BILLS 0.00000% 11/27/2026,50000,$99.594,+$0.011,"
    "$49797.00,$0.00,0.00%,--,--,5.00%,$49700.00,--,Cash,\n"
    "\n"
    '"Disclaimer footer text"\n'
)

@pytest.fixture
def setup(tmp_path, monkeypatch):
    downloads = tmp_path / "Downloads"
    downloads.mkdir()
    (downloads / "Portfolio_Positions_Oct-21-2026.csv").write_text(CSV_EXPORT, encoding="utf-8")
    monkeypatch.setattr(maturing, "windows_downloads", lambda: downloads)
    # data/ is empty, so these tests only pass if Downloads is read.
    data = tmp_path / "data"
    data.mkdir()
    sent = []
    return data, sent


def run(setup, today, *args):
    data, sent = setup
    maturing.main(
        [*args, "--to", "me@example.com", "--date", today.isoformat()],
        data_dir=data, send=lambda *msg: sent.append(msg),
    )
    return sent


def test_latest_csv_prefers_downloads_export(tmp_path):
    downloads, data = tmp_path / "Downloads", tmp_path / "data"
    downloads.mkdir(), data.mkdir()
    (downloads / "Portfolio_Positions_Oct-17-2026.csv").touch()
    (downloads / "Portfolio_Positions_Oct-21-2026.xlsx").touch()
    (downloads / "unrelated_Oct-28-2026.csv").touch()
    (data / "Portfolio_Positions_Oct-28-2026.csv").touch()
    assert maturing.latest_csv(downloads, data).name == "Portfolio_Positions_Oct-17-2026.csv"


def test_latest_csv_falls_back_to_data_with_warning(tmp_path, capsys):
    (tmp_path / "old.csv").touch()
    assert maturing.latest_csv(None, tmp_path).name == "old.csv"
    assert "warning" in capsys.readouterr().err


def test_latest_csv_missing(tmp_path):
    with pytest.raises(FileNotFoundError):
        maturing.latest_csv(None, tmp_path)


def test_emails_bonds_maturing_today(setup):
    sent = run(setup, date(2026, 11, 27))
    assert len(sent) == 1
    to, subject, body = sent[0]
    assert to == "me@example.com"
    assert "11/27/2026" in subject
    assert "912797SU2" in body
    assert "Joint" in body and "IRA" in body
    assert "$10,000.00" in body and "$50,000.00" in body
    assert "11/27/2026" in body
    # The bond maturing on another day is left out.
    assert "912797WS2" not in body


def test_no_email_when_nothing_matures(setup, capsys):
    assert run(setup, date(2026, 11, 26)) == []
    assert "no bonds mature" in capsys.readouterr().out


def test_dry_run_prints_instead_of_sending(setup, capsys):
    assert run(setup, date(2026, 11, 27), "--dry-run") == []
    assert "912797SU2" in capsys.readouterr().out


def test_send_failure_exits(setup):
    data, _ = setup

    def fail(*msg):
        raise subprocess.CalledProcessError(1, "gmail_send.py")

    with pytest.raises(SystemExit, match="failed"):
        maturing.main(["--date", "2026-11-27"], data_dir=data, send=fail)

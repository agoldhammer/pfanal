"""Email a notice for each bond in the newest Fidelity CSV that matures today.

The CSV is the newest Portfolio_Positions*.csv in Windows Downloads, or the
newest .csv in data/ if Downloads has none or this isn't WSL.

Meant to run from cron every weekday. Mail goes out through the Gmail API
helper in ~/Prog/scripts/gmail (gmail_send.py), which authenticates with its
OAuth token.json, so no password is stored here.

Options: --dry-run prints the message instead of sending it, --date
YYYY-MM-DD checks a day other than today, and --to changes the recipient.
"""

import argparse
import os
import subprocess
import sys
from collections.abc import Callable
from datetime import date
from pathlib import Path

from pfanal import DATA_DIR, EXPORT_GLOB, Position, newest, read_positions, windows_downloads

GMAIL_SEND = Path.home() / "Prog" / "scripts" / "gmail" / "gmail_send.py"
MAIL_TO = "art.goldhammer@gmail.com"


def latest_csv(downloads: Path | None, data_dir: Path = DATA_DIR) -> Path:
    """The newest Fidelity CSV in Downloads, else the newest .csv in data/."""
    if downloads and (files := list(downloads.glob(f"{EXPORT_GLOB}.csv"))):
        return newest(files)
    # Under cron this usually means cmd.exe isn't on PATH, so say so in the log.
    print(f"warning: no Fidelity CSV in Downloads ({downloads}); using {data_dir}/",
          file=sys.stderr)
    if not (files := list(data_dir.glob("*.csv"))):
        raise FileNotFoundError(f"no .csv files found in {data_dir}/")
    return newest(files)


def compose(maturing: list[Position], day: date) -> tuple[str, str]:
    """Subject and body of the notice."""
    lines = [f"{len(maturing)} bond(s) mature today, {day:%m/%d/%Y}:", ""]
    for p in maturing:
        lines += [
            f"Maturity date: {p.maturity:%m/%d/%Y}",
            f"Amount:        ${p.maturity_value:,.2f}",
            f"Account name:  {p.account_name}",
            f"CUSIP:         {p.symbol}",
            "",
        ]
    if len(maturing) > 1:
        lines.append(f"Total: ${sum(p.maturity_value for p in maturing):,.2f}")
    return f"Bonds maturing today, {day:%m/%d/%Y}", "\n".join(lines)


def send_gmail(to: str, subject: str, body: str) -> None:
    # `uv run` sets UV to its own path; cron's PATH may not include uv.
    uv = os.environ.get("UV", "uv")
    subprocess.run(
        [uv, "run", "--quiet", str(GMAIL_SEND), "--to", to, "--subject", subject, "--body", body],
        check=True,
    )


def main(
    argv: list[str] | None = None,
    data_dir: Path = DATA_DIR,
    send: Callable[[str, str, str], None] = send_gmail,
) -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--dry-run", action="store_true", help="print instead of sending")
    parser.add_argument("--date", type=date.fromisoformat, default=date.today(),
                        help="day to check (YYYY-MM-DD), default today")
    parser.add_argument("--to", default=MAIL_TO, help=f"recipient, default {MAIL_TO}")
    args = parser.parse_args(argv)
    try:
        path = latest_csv(windows_downloads(), data_dir)
    except FileNotFoundError as e:
        sys.exit(f"error: {e}")
    maturing = [p for p in read_positions(path) if p.maturity == args.date]
    if not maturing:
        print(f"{args.date}: no bonds mature today ({path})")
        return
    subject, body = compose(maturing, args.date)
    if args.dry_run:
        print(f"To: {args.to}\nSubject: {subject}\n\n{body}")
        return
    try:
        send(args.to, subject, body)
    except (OSError, subprocess.CalledProcessError) as e:
        sys.exit(f"error: sending to {args.to} failed: {e}")
    print(f"{args.date}: emailed {len(maturing)} maturing bond(s) to {args.to} ({path})")


if __name__ == "__main__":
    main()

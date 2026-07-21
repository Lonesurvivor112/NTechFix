#!/usr/bin/env python3
"""
Snipe-IT Monitor Assignment Audit
---------------------------------
Cross-references a monitor report (computer name -> monitor serial) against
Snipe-IT to answer, per user: "Is this monitor actually checked out to the
person using this computer?"

For each row in the report:
    1. Look up the COMPUTER in Snipe-IT -> find who it's checked out to.
    2. Look up the MONITOR by serial in Snipe-IT -> find who IT's checked out to.
    3. Compare. A row counts as a match if the monitor is assigned to the
       same user, OR assigned to the computer asset itself (asset-to-asset
       checkout).

CONFIGURATION (environment variables — recommended):

    PowerShell:
        $env:SNIPEIT_URL   = "https://snipe.domain.com"
        $env:SNIPEIT_TOKEN = "your-api-token-here"
        python snipeit_monitor_audit.py

    To set them PERMANENTLY on Windows (survives closing the terminal):
        setx SNIPEIT_URL "https://snipe.domain.com"
        setx SNIPEIT_TOKEN "your-api-token-here"
        (then open a NEW terminal window)

INPUT FILE (--input, default: monitor_report.csv):
    A CSV where one column is the computer name and one is the monitor
    serial. Headers containing "pc"/"computer"/"name" and "serial" are
    auto-detected; override with --pc-col / --serial-col if needed.
    A computer with multiple monitors just appears on multiple rows.

USAGE:
    python snipeit_monitor_audit.py
    python snipeit_monitor_audit.py -i monitor_report_clean.csv -o monitor_audit.csv



OUTPUT (CSV) columns:
    computer              computer name from the report
    computer_assigned_to  who the computer is checked out to in Snipe-IT
    monitor_serial        monitor serial from the report
    monitor_asset_tag     the monitor's asset tag in Snipe-IT (if found)
    match                 YES / YES (via computer) / NO / blank if unknowable
    monitor_assigned_to   who the monitor is ACTUALLY checked out to
    detail                explanation, especially for non-matches:
                              MATCH                    same user
                              MATCH (via computer)     assigned to the PC asset
                              MISMATCH                 assigned to someone else
                              MONITOR NOT ASSIGNED     in Snipe-IT, not checked out
                              MONITOR NOT IN SNIPE-IT  serial not found
                              COMPUTER NOT FOUND       computer not in Snipe-IT
                              COMPUTER NOT ASSIGNED    computer not checked out
"""

import argparse
import csv
import os
import sys
import time
from pathlib import Path

import requests

# ---------------------------------------------------------------------------
# Configuration — env vars take priority.
# ---------------------------------------------------------------------------
SNIPEIT_URL = os.environ.get("SNIPEIT_URL", "https://snipe.domain.com").strip().rstrip("/")
SNIPEIT_TOKEN = os.environ.get("SNIPEIT_TOKEN", "").strip()

DEFAULT_INPUT = Path("monitor_report.csv")
DEFAULT_OUTPUT = Path("monitor_audit.csv")

REQUEST_DELAY = 0.15  # seconds between API calls
TIMEOUT = 15

# Header keywords used to auto-detect columns in the report
PC_HEADER_HINTS = ("pc name", "pcname", "computer", "device", "hostname", "pc", "name")
SERIAL_HEADER_HINTS = ("monitor sn", "monitor serial", "serial number", "serialnumber", "serial", "sn")


# ---------------------------------------------------------------------------
# API helpers
# ---------------------------------------------------------------------------
def build_session(token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    })
    return s


def check_credentials(session: requests.Session) -> None:
    """Fail fast with a clear message if the URL or token is bad."""
    if not SNIPEIT_TOKEN:
        sys.exit(
            "ERROR: SNIPEIT_TOKEN is not set.\n"
            "Set it as an environment variable (see instructions at the top "
            "of this script), then re-run."
        )
    try:
        r = session.get(f"{SNIPEIT_URL}/api/v1/hardware", params={"limit": 1}, timeout=TIMEOUT)
    except requests.RequestException as e:
        sys.exit(f"ERROR: Could not reach {SNIPEIT_URL} — {e}")

    if r.status_code == 401:
        sys.exit(
            "ERROR: Snipe-IT returned 401 Unauthorized.\n"
            "Your SNIPEIT_TOKEN is invalid, expired, or truncated.\n"
            "Generate a new one in Snipe-IT: Account -> Manage API Keys."
        )
    if r.status_code != 200:
        sys.exit(f"ERROR: Unexpected response from Snipe-IT: HTTP {r.status_code}\n{r.text[:300]}")


def find_computer(session: requests.Session, name: str) -> dict | None:
    """Find a computer asset by tag first, then by name search."""
    try:
        r = session.get(f"{SNIPEIT_URL}/api/v1/hardware/bytag/{name}", timeout=TIMEOUT)
        if r.status_code == 200:
            data = r.json()
            if isinstance(data, dict) and data.get("id"):
                return data

        r = session.get(
            f"{SNIPEIT_URL}/api/v1/hardware",
            params={"search": name, "limit": 50},
            timeout=TIMEOUT,
        )
        if r.status_code != 200:
            return None
        rows = r.json().get("rows", []) or []
        for row in rows:
            if (row.get("asset_tag") or "").lower() == name.lower() \
               or (row.get("name") or "").lower() == name.lower():
                return row
        return rows[0] if rows else None
    except requests.RequestException:
        return None


def find_monitor_by_serial(session: requests.Session, serial: str) -> dict | None:
    """Find an asset by serial number. Returns the asset dict or None."""
    try:
        # byserial returns {"rows": [...]} even for a single hit
        r = session.get(f"{SNIPEIT_URL}/api/v1/hardware/byserial/{serial}", timeout=TIMEOUT)
        if r.status_code == 200:
            data = r.json()
            rows = data.get("rows") if isinstance(data, dict) else None
            if rows:
                for row in rows:
                    if (row.get("serial") or "").lower() == serial.lower():
                        return row
                return rows[0]

        # Fallback: general search
        r = session.get(
            f"{SNIPEIT_URL}/api/v1/hardware",
            params={"search": serial, "limit": 20},
            timeout=TIMEOUT,
        )
        if r.status_code == 200:
            for row in (r.json().get("rows") or []):
                if (row.get("serial") or "").lower() == serial.lower():
                    return row
        return None
    except requests.RequestException:
        return None


def assigned_summary(asset: dict) -> tuple[str, str, int | None]:
    """
    Return (type, display_name, id) of whoever an asset is assigned to.
    type is 'user', 'asset', 'location', or '' if unassigned.
    """
    assigned = asset.get("assigned_to")
    if not assigned:
        return "", "", None
    a_type = assigned.get("type", "user")
    a_name = assigned.get("name") or assigned.get("username") or ""
    return a_type, a_name, assigned.get("id")


# ---------------------------------------------------------------------------
# Report parsing
# ---------------------------------------------------------------------------
def detect_column(fieldnames: list[str], hints: tuple[str, ...], label: str) -> str:
    lowered = {fn.lower().strip(): fn for fn in fieldnames}
    for hint in hints:
        for low, original in lowered.items():
            if hint == low or hint in low:
                return original
    sys.exit(
        f"ERROR: Could not find a {label} column in the report.\n"
        f"Headers found: {fieldnames}\n"
        f"Specify it explicitly, e.g. --pc-col \"PC Name\" --serial-col \"Serial\""
    )


def read_report(path: Path, pc_col: str | None, serial_col: str | None) -> list[tuple[str, str]]:
    if not path.exists():
        sys.exit(f"ERROR: Input file not found: {path}")

    with path.open(newline="", encoding="utf-8-sig") as f:
        sample = f.read(4096)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(f, dialect=dialect)

        if not reader.fieldnames:
            sys.exit(f"ERROR: {path} appears to be empty.")

        pc_col = pc_col or detect_column(reader.fieldnames, PC_HEADER_HINTS, "computer-name")
        serial_col = serial_col or detect_column(reader.fieldnames, SERIAL_HEADER_HINTS, "monitor-serial")
        print(f"Using columns: computer = '{pc_col}', monitor serial = '{serial_col}'\n")

        rows = []
        for row in reader:
            pc = (row.get(pc_col) or "").strip()
            serial = (row.get(serial_col) or "").strip()
            if pc and serial:
                rows.append((pc, serial))

    if not rows:
        sys.exit(f"ERROR: No usable rows found in {path}")
    return rows


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main() -> None:
    parser = argparse.ArgumentParser(description="Audit monitor assignments in Snipe-IT.")
    parser.add_argument("--input", "-i", type=Path, default=DEFAULT_INPUT,
                        help=f"Monitor report CSV (default: {DEFAULT_INPUT})")
    parser.add_argument("--output", "-o", type=Path, default=DEFAULT_OUTPUT,
                        help=f"Output CSV (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--pc-col", help="Header of the computer-name column (auto-detected if omitted)")
    parser.add_argument("--serial-col", help="Header of the monitor-serial column (auto-detected if omitted)")
    args = parser.parse_args()

    session = build_session(SNIPEIT_TOKEN)
    check_credentials(session)

    report = read_report(args.input, args.pc_col, args.serial_col)
    print(f"Auditing {len(report)} monitor row(s) against {SNIPEIT_URL} ...\n")

    computer_cache: dict[str, dict | None] = {}
    results = []

    for i, (pc, serial) in enumerate(report, 1):
        out = {
            "computer": pc,
            "computer_assigned_to": "",
            "monitor_serial": serial,
            "monitor_asset_tag": "",
            "match": "",
            "monitor_assigned_to": "",
            "detail": "",
        }

        # --- Step 1: computer -> user ---
        key = pc.lower()
        if key not in computer_cache:
            computer_cache[key] = find_computer(session, pc)
            time.sleep(REQUEST_DELAY)
        computer = computer_cache[key]

        c_type = c_name = ""
        c_assigned_id = None
        if computer is None:
            out["detail"] = "COMPUTER NOT FOUND"
        else:
            c_type, c_name, c_assigned_id = assigned_summary(computer)
            if c_type == "user":
                out["computer_assigned_to"] = c_name
            elif c_type:
                out["computer_assigned_to"] = f"{c_name} ({c_type})"

        # --- Step 2: monitor by serial ---
        monitor = find_monitor_by_serial(session, serial)
        time.sleep(REQUEST_DELAY)

        m_type = m_name = ""
        m_assigned_id = None
        if monitor is None:
            out["detail"] = out["detail"] or "MONITOR NOT IN SNIPE-IT"
        else:
            out["monitor_asset_tag"] = monitor.get("asset_tag") or ""
            m_type, m_name, m_assigned_id = assigned_summary(monitor)
            if m_type == "user":
                out["monitor_assigned_to"] = m_name
            elif m_type:
                out["monitor_assigned_to"] = f"{m_name} ({m_type})"

        # --- Step 3: compare (only when both lookups say enough) ---
        if computer is not None and monitor is not None and not out["detail"]:
            if not m_type:
                out["match"] = "NO"
                out["detail"] = "MONITOR NOT ASSIGNED"
            elif not c_type:
                out["match"] = "NO"
                out["detail"] = "COMPUTER NOT ASSIGNED"
            elif m_type == "user" and c_type == "user" and m_assigned_id == c_assigned_id:
                out["match"] = "YES"
                out["detail"] = "MATCH"
            elif m_type == "asset" and m_assigned_id == computer.get("id"):
                out["match"] = "YES (via computer)"
                out["detail"] = "MATCH (via computer)"
            else:
                out["match"] = "NO"
                out["detail"] = "MISMATCH"

        results.append(out)

        # Console line: computer's user, match verdict, and (on mismatch) who has the monitor
        pc_owner = out["computer_assigned_to"] or "(unassigned)"
        line = f"[{i:>3}/{len(report)}] {pc} [{pc_owner}]  /  {serial}  ->  {out['detail']}"
        if out["detail"] == "MISMATCH":
            line += f"  (monitor is with: {out['monitor_assigned_to'] or 'unknown'})"
        print(line)

    # Write CSV
    with args.output.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(
            f,
            fieldnames=["computer", "computer_assigned_to", "monitor_serial",
                        "monitor_asset_tag", "match", "monitor_assigned_to", "detail"],
        )
        w.writeheader()
        w.writerows(results)

    # Summary
    counts: dict[str, int] = {}
    for r in results:
        counts[r["detail"]] = counts.get(r["detail"], 0) + 1
    print(f"\nDone. Results saved to: {args.output.resolve()}\nSummary:")
    for status, n in sorted(counts.items(), key=lambda kv: -kv[1]):
        print(f"  {n:>4}  {status}")


if __name__ == "__main__":
    main()

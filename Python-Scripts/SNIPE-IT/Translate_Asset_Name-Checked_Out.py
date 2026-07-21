#!/usr/bin/env python3
"""
Snipe-IT Checkout Lookup
------------------------
Reads a list of device/asset names and prints who each one is checked out to.

HOW TO USE
----------
1. Generate an API key in Snipe-IT:
      Account -> Manage API Keys

2. Set your Snipe-IT URL and API token as environment variables.

   PowerShell:
      $env:SNIPEIT_URL="https://snipe.domain.com"
      $env:SNIPEIT_TOKEN="your-api-token-here"

   CMD:
      set SNIPEIT_URL=https://snipe.domain.com
      set SNIPEIT_TOKEN=your-api-token-here

3. Create a text file containing one asset name or asset tag per line:

      PC001
      PC002
      LT-100
      LT-101

4. Run the script:

      python snipeit_checkout_lookup.py

   Or specify custom paths:

      python snipeit_checkout_lookup.py --input devices.txt --output results.csv

5. Review the results:
   - Progress is displayed in the console.
   - A CSV report is generated containing:
       • Device Name
       • Asset Tag
       • Assigned User
       • Email Address
       • Checkout Status
       • Serial Number

Example Output:
   [  1/4] PC001    -> John Smith
   [  2/4] PC002    -> Available (not checked out)
   [  3/4] LT-100   -> Jane Doe
   [  4/4] LT-101   -> NOT FOUND

The resulting CSV file can be opened in Excel for reporting,
auditing, or asset management purposes.

CONFIGURATION (environment variables — recommended):

    PowerShell:
        $env:SNIPEIT_URL="https://snipe.domain.com"
        $env:SNIPEIT_TOKEN="your-api-token-here"
        python snipeit_checkout_lookup.py
"""

"""

Snipe-IT Checkout Lookup
------------------------
Reads a list of device/asset names and prints who each one is checked out to.

CONFIGURATION (environment variables — recommended):

    PowerShell:
        $env:SNIPEIT_URL   = "https://snipe.domain.com"
        $env:SNIPEIT_TOKEN = "your-api-token-here"
        python snipeit_checkout_lookup.py

    CMD:
        set SNIPEIT_URL=https://snipe.domain.com
        set SNIPEIT_TOKEN=your-api-token-here
        python snipeit_checkout_lookup.py

    To set them PERMANENTLY on Windows (survives closing the terminal):
        setx SNIPEIT_URL "https://snipe.domain.com"
        setx SNIPEIT_TOKEN "your-api-token-here"
        (then open a NEW terminal window)

    Generate a token in Snipe-IT under: Account -> Manage API Keys.
    Paste the token by itself — no "Bearer " prefix, no quotes inside it.

USAGE:
    python snipeit_checkout_lookup.py
    python snipeit_checkout_lookup.py --input devices.txt --output results.csv

    The input file has one asset name per line. A header like "PC Name" is
    skipped automatically.
"""

import argparse
import csv
import os
import sys
import time
from pathlib import Path

import requests

# ---------------------------------------------------------------------------
# Configuration — env vars take priority; the fallbacks below are only used
# when the env var is not set.
# ---------------------------------------------------------------------------
SNIPEIT_URL = os.environ.get("SNIPEIT_URL", "https://snipe.domain.com").strip().rstrip("/")
SNIPEIT_TOKEN = os.environ.get("SNIPEIT_TOKEN", "").strip()

DEFAULT_INPUT = Path(
    r"c:/Pathtoinput/devices.txt"
)
DEFAULT_OUTPUT = Path("checkout_results.csv")

REQUEST_DELAY = 0.15  # seconds between requests, to be nice to the API
TIMEOUT = 15          # per-request timeout in seconds

HEADER_ALIASES = {"pc name", "pcname", "name", "asset", "asset_tag", "device"}


def build_session(token: str) -> requests.Session:
    """One session = connection reuse = faster, and headers set once."""
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
            "Tokens are very long (1000+ chars) — make sure the whole thing "
            "copied, with no 'Bearer ' prefix and no surrounding quotes.\n"
            "Generate a new one in Snipe-IT: Account -> Manage API Keys."
        )
    if r.status_code != 200:
        sys.exit(f"ERROR: Unexpected response from Snipe-IT: HTTP {r.status_code}\n{r.text[:300]}")


def lookup_asset(session: requests.Session, name: str) -> dict:
    """
    Look up an asset by its name (asset_tag first, then name search).
    Returns a dict with: device, asset_tag, assigned_to, email, status, serial.
    """
    result = {
        "device": name,
        "asset_tag": "",
        "assigned_to": "",
        "email": "",
        "status": "",
        "serial": "",
    }

    asset = None

    # First try: exact asset_tag lookup (fastest, most reliable)
    try:
        r = session.get(f"{SNIPEIT_URL}/api/v1/hardware/bytag/{name}", timeout=TIMEOUT)
    except requests.RequestException as e:
        result["status"] = f"ERROR: {e}"
        return result

    if r.status_code == 200:
        data = r.json()
        # bytag returns the asset directly on success, or {"status":"error",...} on miss
        if isinstance(data, dict) and data.get("id"):
            asset = data

    # Fallback: search by name
    if asset is None:
        try:
            r = session.get(
                f"{SNIPEIT_URL}/api/v1/hardware",
                params={"search": name, "limit": 50},
                timeout=TIMEOUT,
            )
        except requests.RequestException as e:
            result["status"] = f"ERROR: {e}"
            return result

        if r.status_code != 200:
            result["status"] = f"HTTP {r.status_code}"
            return result

        rows = r.json().get("rows", []) or []
        # Prefer exact match on asset_tag or name
        for row in rows:
            if (row.get("asset_tag") or "").lower() == name.lower() \
               or (row.get("name") or "").lower() == name.lower():
                asset = row
                break
        if asset is None and rows:
            asset = rows[0]  # best effort
            result["status"] = "(inexact match) "

    if asset is None:
        result["status"] = "NOT FOUND"
        return result

    result["asset_tag"] = asset.get("asset_tag") or ""
    result["serial"] = asset.get("serial") or ""

    assigned = asset.get("assigned_to")
    if not assigned:
        result["status"] += "Available (not checked out)"
        return result

    # assigned_to can be a user, asset, or location
    a_type = assigned.get("type", "user")
    a_name = assigned.get("name") or ""
    a_user = assigned.get("username") or ""
    a_email = assigned.get("email") or ""

    if a_type == "user":
        result["assigned_to"] = a_name or a_user
        result["email"] = a_email
        result["status"] += "Checked out to user"
    else:
        result["assigned_to"] = a_name
        result["status"] += f"Checked out to {a_type}"

    return result


def read_devices(path: Path) -> list[str]:
    if not path.exists():
        sys.exit(f"ERROR: Input file not found: {path}")

    names = []
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line = line.strip()
        if not line or line.lower() in HEADER_ALIASES:
            continue
        names.append(line)

    if not names:
        sys.exit(f"ERROR: No device names found in {path}")
    return names


def main() -> None:
    parser = argparse.ArgumentParser(description="Look up Snipe-IT asset checkouts.")
    parser.add_argument("--input", "-i", type=Path, default=DEFAULT_INPUT,
                        help=f"Text file with one asset name per line (default: {DEFAULT_INPUT})")
    parser.add_argument("--output", "-o", type=Path, default=DEFAULT_OUTPUT,
                        help=f"Output CSV path (default: {DEFAULT_OUTPUT})")
    args = parser.parse_args()

    session = build_session(SNIPEIT_TOKEN)
    check_credentials(session)

    devices = read_devices(args.input)
    print(f"Looking up {len(devices)} device(s) against {SNIPEIT_URL} ...\n")

    results = []
    width = max(len(d) for d in devices)
    for i, name in enumerate(devices, 1):
        res = lookup_asset(session, name)
        results.append(res)
        who = res["assigned_to"] or res["status"]
        print(f"[{i:>3}/{len(devices)}] {name:<{width}}  ->  {who}")
        time.sleep(REQUEST_DELAY)

    with args.output.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(
            f,
            fieldnames=["device", "asset_tag", "assigned_to", "email", "status", "serial"],
        )
        w.writeheader()
        w.writerows(results)

    found = sum(1 for r in results if r["status"] != "NOT FOUND" and not r["status"].startswith("ERROR"))
    print(f"\nDone. {found}/{len(results)} found. Results saved to: {args.output.resolve()}")


if __name__ == "__main__":
    main()

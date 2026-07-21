#!/usr/bin/env python3
"""
Pull Snipe-IT asset history by asset tag OR asset name.
- --mode tag  (default): look up by asset tag, e.g. B0211
- --mode name           : search by asset name, e.g. L157
- Optional CSV export

TOKEN HANDLING
==============
The API token is read from the SNIPE_TOKEN environment variable if set,
otherwise from the _HARDCODED_TOKEN fallback below. Whichever is used,
it is stripped of surrounding whitespace/newlines automatically.

If you get HTTP 401 (Unauthorized or unauthenticated):
  * The token was probably regenerated or revoked
    (Snipe-IT → your profile → Manage API Keys → create a new one).
  * The token's user may be deactivated or lack asset-view permission.
  * Watch for stray spaces, newlines, or smart-quotes in the token value.
"""
import os
import requests
import sys
import urllib.parse
import argparse
import csv
from datetime import datetime

# ============================
# CONFIG (EDIT THESE)
# ============================
SNIPE_URL = os.environ.get("SNIPE_URL", "https://snipe.domain.com").strip()

# Paste a token here as a fallback, OR leave it "" and set the SNIPE_TOKEN
# environment variable instead. Either way it gets .strip()'d below so a
# trailing newline/space can't cause a phantom 401.
_HARDCODED_TOKEN = ""

SNIPE_TOKEN = (os.environ.get("SNIPE_TOKEN") or _HARDCODED_TOKEN).strip()
# ============================

# ============================
# COLORS / FORMATTING
# ============================
# Try to enable ANSI colors on Windows via colorama. Fall back gracefully.
try:
    import colorama
    colorama.just_fix_windows_console()
except ImportError:
    # On Windows 10+ ANSI usually works in the new terminal anyway.
    pass

class C:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    DIM     = "\033[2m"

    # Foreground colors
    RED     = "\033[31m"
    GREEN   = "\033[32m"
    YELLOW  = "\033[33m"
    BLUE    = "\033[34m"
    MAGENTA = "\033[35m"
    CYAN    = "\033[36m"
    WHITE   = "\033[37m"

    # Bright foreground colors
    BRED     = "\033[91m"
    BGREEN   = "\033[92m"
    BYELLOW  = "\033[93m"
    BBLUE    = "\033[94m"
    BMAGENTA = "\033[95m"
    BCYAN    = "\033[96m"
    BWHITE   = "\033[97m"

    # Background
    BG_BLUE = "\033[44m"


def print_banner(query: str, mode: str):
    """Print a bold, colored banner for the search."""
    label = "ASSET TAG" if mode == "tag" else "ASSET NAME"
    title = f"  NEW ASSET SEARCH FOR {query.upper()}  ({label})  "
    width = max(len(title) + 4, 60)
    bar = "═" * width
    pad = " " * ((width - len(title)) // 2)
    centered = (pad + title + pad)[:width]

    print()
    print(f"{C.BCYAN}{C.BOLD}╔{bar}╗{C.RESET}")
    print(f"{C.BCYAN}{C.BOLD}║{C.BYELLOW}{centered}{C.BCYAN}║{C.RESET}")
    print(f"{C.BCYAN}{C.BOLD}╚{bar}╝{C.RESET}")
    print()


def info(msg: str):
    print(f"{C.BGREEN}{C.BOLD}[+]{C.RESET} {msg}")

def warn(msg: str):
    print(f"{C.BYELLOW}{C.BOLD}[!]{C.RESET} {msg}")

def err(msg: str):
    print(f"{C.BRED}{C.BOLD}[-]{C.RESET} {msg}")

def ok(msg: str):
    print(f"{C.BGREEN}{C.BOLD}[ok]{C.RESET} {msg}")


# Color per action type so checkouts/checkins/etc. pop visually
ACTION_COLORS = {
    "checkout":      C.BMAGENTA,
    "checkin from":  C.BCYAN,
    "checkin":       C.BCYAN,
    "update":        C.BYELLOW,
    "create new":    C.BGREEN,
    "delete":        C.BRED,
    "audit":         C.BBLUE,
}

def color_action(action: str) -> str:
    color = ACTION_COLORS.get((action or "").lower(), C.BWHITE)
    return f"{color}{C.BOLD}{action or '-'}{C.RESET}"


def format_datetime(date_field):
    """
    Snipe-IT returns created_at either as a dict {'datetime': '...', 'formatted': '...'}
    or as a raw string. Normalize to a friendly string.
    """
    if isinstance(date_field, dict):
        formatted = date_field.get("formatted")
        raw       = date_field.get("datetime")
        if formatted and raw:
            return f"{raw}  ({formatted})"
        return formatted or raw or "-"
    if isinstance(date_field, str):
        # Try to prettify a raw timestamp
        try:
            dt = datetime.strptime(date_field, "%Y-%m-%d %H:%M:%S")
            return f"{date_field}  ({dt.strftime('%Y-%m-%d %I:%M %p')})"
        except Exception:
            return date_field
    return "-"


# ============================
# ORIGINAL HELPERS
# ============================
def dict_get(obj, key, default=None):
    """Safe dict get for possibly-None objects."""
    if isinstance(obj, dict):
        return obj.get(key, default)
    return default


def make_headers():
    if not SNIPE_TOKEN:
        raise RuntimeError(
            "SNIPE_TOKEN is not set.\n"
            "  Set the SNIPE_TOKEN environment variable, or paste a token into\n"
            "  the _HARDCODED_TOKEN value near the top of this script."
        )
    return {"Authorization": f"Bearer {SNIPE_TOKEN}", "Accept": "application/json"}


def snipe_get(url: str):
    """
    Perform a GET against the Snipe-IT API and return parsed JSON.

    Centralizes auth + error handling so every endpoint reports problems
    the same way. Raises RuntimeError with a friendly message on failure.
    """
    r = requests.get(url, headers=make_headers(), timeout=30)

    try:
        data = r.json()
    except Exception:
        raise RuntimeError(f"Non-JSON response from Snipe-IT (HTTP {r.status_code}): {r.text}")

    # Snipe-IT returns 401 with {'error': 'Unauthorized or unauthenticated.'}
    if r.status_code in (401, 403):
        raise RuntimeError(
            f"HTTP {r.status_code} — Snipe-IT rejected the API token.\n"
            "  Likely causes:\n"
            "    • The token was regenerated or revoked "
            "(profile → Manage API Keys → create a new one).\n"
            "    • The token's user is deactivated or lacks asset-view permission.\n"
            "    • A stray space / newline / smart-quote got into the token value.\n"
            f"  Server said: {data}"
        )

    return r, data


def fetch_asset_by_tag(tag: str):
    """Fetch a single asset by its asset tag (exact match)."""
    encoded = urllib.parse.quote(tag)
    url = f"{SNIPE_URL.rstrip('/')}/api/v1/hardware/bytag/{encoded}?deleted=false"
    r, data = snipe_get(url)
    if r.status_code >= 400 or "id" not in data:
        raise RuntimeError(f"Error fetching asset by tag: HTTP {r.status_code} → {data}")
    return data


def fetch_asset_by_name(name: str):
    """
    Search for assets whose name matches the given string.
    Uses /api/v1/hardware?search=<name> and filters for an exact name match.
    If multiple assets share the same name, prompts the user to pick one.
    """
    encoded = urllib.parse.quote(name)
    url = f"{SNIPE_URL.rstrip('/')}/api/v1/hardware?search={encoded}&limit=50&sort=name&order=asc"
    r, data = snipe_get(url)
    if r.status_code >= 400 or "rows" not in data:
        raise RuntimeError(f"Error searching assets: HTTP {r.status_code} → {data}")

    # Prefer exact name matches; fall back to all results if none match exactly
    rows = data["rows"]
    exact = [a for a in rows if (dict_get(a, "name") or "").lower() == name.lower()]
    candidates = exact if exact else rows

    if not candidates:
        raise RuntimeError(f"No assets found matching name: {name!r}")

    if len(candidates) == 1:
        return candidates[0]

    # Multiple matches — let the user choose
    warn(f"Multiple assets found matching {name!r}:\n")
    for i, asset in enumerate(candidates):
        print(
            f"  {C.BYELLOW}[{i + 1}]{C.RESET} "
            f"ID={C.BCYAN}{dict_get(asset, 'id')}{C.RESET} | "
            f"Tag={C.BCYAN}{dict_get(asset, 'asset_tag', '-')}{C.RESET} | "
            f"Name={C.BCYAN}{dict_get(asset, 'name', '-')}{C.RESET} | "
            f"Model={C.BCYAN}{dict_get(dict_get(asset, 'model'), 'name', '-')}{C.RESET}"
        )
    while True:
        try:
            choice = int(input(f"\n{C.BYELLOW}Enter number to select asset:{C.RESET} "))
            if 1 <= choice <= len(candidates):
                return candidates[choice - 1]
        except (ValueError, KeyboardInterrupt):
            pass
        err("Invalid choice, try again.")


def fetch_history(asset_id: int):
    """Fetch activity history for an asset."""
    url = (
        f"{SNIPE_URL.rstrip('/')}/api/v1/reports/activity"
        f"?item_type=asset&item_id={asset_id}&order=desc&sort=created_at&limit=500"
    )
    r, data = snipe_get(url)
    if r.status_code >= 400 or "rows" not in data:
        raise RuntimeError(f"Error fetching history: HTTP {r.status_code} → {data}")
    return data["rows"]


def main():
    parser = argparse.ArgumentParser(
        description="Get Snipe-IT asset history by tag or name"
    )
    parser.add_argument("QUERY", help="Asset tag (e.g. B0211) or name (e.g. L157)")
    parser.add_argument(
        "--mode",
        choices=["tag", "name"],
        default="tag",
        help="Search mode: 'tag' (default) or 'name'",
    )
    parser.add_argument(
        "--csv", dest="csv_path", default=None,
        help="Optional path to export CSV (e.g. C:\\Reports\\history.csv)"
    )
    parser.add_argument(
        "--no-color", action="store_true",
        help="Disable colored output"
    )
    args = parser.parse_args()
    query = args.QUERY.strip()

    # Strip color codes if requested
    if args.no_color:
        for attr in dir(C):
            if not attr.startswith("_") and isinstance(getattr(C, attr), str):
                setattr(C, attr, "")

    print_banner(query, args.mode)

    if args.mode == "name":
        info(f"Searching by asset name: {C.BCYAN}{query}{C.RESET}")
        asset = fetch_asset_by_name(query)
    else:
        info(f"Looking up asset tag: {C.BCYAN}{query}{C.RESET}")
        asset = fetch_asset_by_tag(query)

    asset_id = asset["id"]
    info(
        f"Found asset: "
        f"ID={C.BCYAN}{asset_id}{C.RESET} | "
        f"Name={C.BCYAN}{dict_get(asset, 'name', '-')}{C.RESET} | "
        f"Tag={C.BCYAN}{dict_get(asset, 'asset_tag', query)}{C.RESET} | "
        f"Model={C.BCYAN}{dict_get(dict_get(asset, 'model'), 'name', '-')}{C.RESET}"
    )

    info("Fetching history...")
    history = fetch_history(asset_id)

    if not history:
        err("No history entries found.")
        return

    info(f"{C.BOLD}{len(history)}{C.RESET} history entries:\n")

    writer = None
    csv_file = None
    if args.csv_path:
        csv_file = open(args.csv_path, "w", newline="", encoding="utf-8")
        writer = csv.writer(csv_file)
        writer.writerow(["created_at", "action_type", "admin_name", "target_name", "note"])

    # Separator line between entries
    separator = f"{C.DIM}{'─' * 80}{C.RESET}"

    for idx, h in enumerate(history):
        date_raw = dict_get(h, "created_at", "")
        date_str = format_datetime(date_raw)
        action   = dict_get(h, "action_type", "")
        admin    = dict_get(dict_get(h, "admin"), "name", "")
        target   = dict_get(dict_get(h, "target"), "name", "")
        note     = dict_get(h, "note", "")

        # Highlight notes that have actual content vs "-"
        note_display = note if note else "-"
        note_colored = (
            f"{C.BWHITE}{note_display}{C.RESET}" if note
            else f"{C.DIM}-{C.RESET}"
        )

        print(
            f"{C.BBLUE}●{C.RESET} "
            f"{C.BWHITE}{date_str}{C.RESET}\n"
            f"   {C.DIM}Action:{C.RESET} {color_action(action)}\n"
            f"   {C.DIM}By:    {C.RESET} {C.BGREEN}{admin or '-'}{C.RESET}\n"
            f"   {C.DIM}Target:{C.RESET} {C.BMAGENTA}{target or '-'}{C.RESET}\n"
            f"   {C.DIM}Note:  {C.RESET} {note_colored}"
        )

        # Add break line between entries (but not after the last one)
        if idx < len(history) - 1:
            print(separator)

        if writer:
            # Keep CSV plain — write the raw datetime if available
            csv_date = date_raw.get("datetime") if isinstance(date_raw, dict) else date_raw
            writer.writerow([csv_date, action, admin, target, note])

    if csv_file:
        csv_file.close()
        print()
        ok(f"CSV exported to: {C.BCYAN}{args.csv_path}{C.RESET}")


if __name__ == "__main__":
    try:
        main()
    except RuntimeError as e:
        # Clean, readable failure instead of a raw traceback for expected errors
        # (bad token, asset not found, API errors, etc.)
        err(str(e))
        sys.exit(1)
    except KeyboardInterrupt:
        print()
        err("Cancelled.")
        sys.exit(130)

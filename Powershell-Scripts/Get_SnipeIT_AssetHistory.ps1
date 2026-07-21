
#!/usr/bin/env python3
"""
Pull Snipe-IT asset history by asset tag.
- Matches your label script's bytag lookup (URL-encoded tag, ?deleted=false)
- Prints history safely even when admin/target are null
- Optional CSV export
"""

import requests
import sys
import urllib.parse
import argparse
import csv

# ============================
# CONFIG (EDIT THESE)
# ============================
SNIPE_URL = "https://snipe.contoso.com"
SNIPE_TOKEN = ""
# ============================


def dict_get(obj, key, default=None):
    """Safe dict get for possibly-None objects."""
    if isinstance(obj, dict):
        return obj.get(key, default)
    return default


def fetch_asset(tag: str):
    """Fetch asset data by tag, using EXACT logic from your label script."""
    if not SNIPE_TOKEN:
        raise RuntimeError("SNIPE_TOKEN is not set")

    encoded = urllib.parse.quote(tag)
    url = f"{SNIPE_URL.rstrip('/')}/api/v1/hardware/bytag/{encoded}?deleted=false"

    headers = {"Authorization": f"Bearer {SNIPE_TOKEN}", "Accept": "application/json"}
    r = requests.get(url, headers=headers, timeout=30)

    try:
        data = r.json()
    except Exception:
        raise RuntimeError(f"Non‑JSON error from Snipe‑IT: {r.text}")

    if r.status_code >= 400 or "id" not in data:
        raise RuntimeError(f"Error fetching asset: HTTP {r.status_code} → {data}")

    return data


def fetch_history(asset_id: int):
    """Fetch history (activity) for an asset via Snipe-IT's official activity endpoint."""
    url = (
        f"{SNIPE_URL.rstrip('/')}/api/v1/reports/activity"
        f"?item_type=asset&item_id={asset_id}&order=desc&sort=created_at&limit=500"
    )

    headers = {"Authorization": f"Bearer {SNIPE_TOKEN}", "Accept": "application/json"}
    r = requests.get(url, headers=headers, timeout=30)

    try:
        data = r.json()
    except Exception:
        raise RuntimeError(f"Non‑JSON error from Snipe‑IT: {r.text}")

    if r.status_code >= 400 or "rows" not in data:
        raise RuntimeError(f"Error fetching history: HTTP {r.status_code} → {data}")

    return data["rows"]


def main():
    parser = argparse.ArgumentParser(description="Get Snipe‑IT asset history by tag")
    parser.add_argument("ASSET_TAG", help="Snipe‑IT asset tag (e.g., B0211)")
    parser.add_argument(
        "--csv", dest="csv_path", default=None,
        help="Optional path to export CSV (e.g., C:\\Reports\\asset_B0211_history.csv)"
    )
    args = parser.parse_args()

    tag = args.ASSET_TAG.strip()

    print(f"[+] Looking up asset tag: {tag}")
    asset = fetch_asset(tag)
    asset_id = asset["id"]
    print(f"[+] Found asset: ID={asset_id}, Name={dict_get(asset,'name')}, Tag={dict_get(asset,'asset_tag', tag)}")

    print("[+] Fetching history...")
    history = fetch_history(asset_id)

    if not history:
        print("[-] No history entries found.")
        return

    print(f"[+] {len(history)} history entries:\n")

    # Optional CSV setup
    writer = None
    csv_file = None
    if args.csv_path:
        csv_file = open(args.csv_path, "w", newline="", encoding="utf-8")
        writer = csv.writer(csv_file)
        writer.writerow(["created_at", "action_type", "admin_name", "target_name", "note"])

    for h in history:
        date   = dict_get(h, "created_at", "")
        action = dict_get(h, "action_type", "")
        admin  = dict_get(dict_get(h, "admin"), "name", "")
        target = dict_get(dict_get(h, "target"), "name", "")
        note   = dict_get(h, "note", "")

        print(f"- {date} | {action} | By: {admin or '-'} | Target: {target or '-'} | Note: {note or '-'}")

        if writer:
            writer.writerow([date, action, admin, target, note])

    if csv_file:
        csv_file.close()
        print(f"\n[ok] CSV exported to: {args.csv_path}")


if __name__ == "__main__":
    main()

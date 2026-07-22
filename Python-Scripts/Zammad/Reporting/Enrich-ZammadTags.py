import json
import time
import requests
import pandas as pd
from pathlib import Path

# =====================================================
# CONFIG
# =====================================================

CONFIG_FILE = "config/settings.json"

with open(
    CONFIG_FILE,
    "r",
    encoding="utf-8"
) as f:
    config = json.load(f)

BASE_URL = config["base_url"].rstrip("/")
TOKEN = config["token"]

REQUEST_TIMEOUT = int(config.get("timeout", 30))
MAX_RETRIES = int(config.get("max_retries", 3))
RETRY_DELAY = int(config.get("retry_delay", 5))

HEADERS = {
    "Authorization": f"Token token={TOKEN}",
    "Content-Type": "application/json"
}

OUTPUT_DIR = Path("output")
INPUT_FILE = OUTPUT_DIR / "Raw_Tickets.xlsx"
BACKUP_FILE = OUTPUT_DIR / "Raw_Tickets_pre_tags.xlsx"

# =====================================================
# FUNCTIONS
# =====================================================

def api_get(session, url):
    """
    HTTP GET with retry logic.
    """

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(
                url,
                timeout=REQUEST_TIMEOUT
            )

            response.raise_for_status()

            return response

        except requests.exceptions.RequestException as ex:
            last_error = ex

            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY)

    raise last_error


def get_ticket_tags(session, ticket_id):
    """
    Fetch tags for a single ticket.
    Returns a list of tag strings.
    """

    url = (
        f"{BASE_URL}/api/v1/tags"
        f"?object=Ticket&o_id={ticket_id}"
    )

    response = api_get(session, url)

    try:
        data = response.json()
    except ValueError:
        return []

    # Zammad returns {"tags": ["a","b"]}
    if isinstance(data, dict):
        tags = data.get("tags", [])

    elif isinstance(data, list):
        tags = data

    else:
        tags = []

    # Normalize
    cleaned = []

    for tag in tags:
        if isinstance(tag, dict):
            name = tag.get("name") or tag.get("tag") or ""
        else:
            name = str(tag)

        name = name.strip()

        if name:
            cleaned.append(name)

    return cleaned


# =====================================================
# LOAD DATA
# =====================================================

print()
print("=" * 60)
print("Zammad Tag Enrichment")
print("=" * 60)

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Missing input file: {INPUT_FILE}. "
        f"Run Generate-ZammadReport.py first."
    )

df = pd.read_excel(INPUT_FILE)

print(f"Loaded {len(df):,} tickets")

if "id" not in df.columns:
    raise ValueError(
        "Raw_Tickets.xlsx is missing an 'id' column."
    )

# =====================================================
# BACKUP
# =====================================================

if not BACKUP_FILE.exists():
    df.to_excel(BACKUP_FILE, index=False)
    print(f"Backup saved: {BACKUP_FILE}")

# =====================================================
# SANITY CHECK ON ONE TICKET
# =====================================================

session = requests.Session()
session.headers.update(HEADERS)

first_id = int(df["id"].iloc[0])

print()
print(f"Sanity check — testing ticket ID {first_id}...")

try:
    sample_tags = get_ticket_tags(session, first_id)

    print(
        f"  API returned {len(sample_tags)} tag(s): "
        f"{sample_tags}"
    )

except Exception as ex:
    print(f"  ERROR: {ex}")
    print()
    print(
        "The Zammad tag endpoint is not reachable. "
        "Check base_url, token, and token permissions."
    )

    raise SystemExit(1)

# =====================================================
# ENRICH
# =====================================================

print()
print("Fetching tags for all tickets...")

total = len(df)
success = 0
errors = 0
with_tags = 0

df["tags"] = ""

start_time = time.time()

for pos, (idx, row) in enumerate(df.iterrows(), start=1):
    ticket_id = row["id"]

    try:
        ticket_id_int = int(ticket_id)
    except (ValueError, TypeError):
        errors += 1
        continue

    try:
        tags = get_ticket_tags(session, ticket_id_int)

        df.at[idx, "tags"] = ",".join(tags)

        success += 1

        if tags:
            with_tags += 1

    except Exception as ex:
        errors += 1

    # Progress every 100
    if pos % 100 == 0 or pos == total:
        elapsed = time.time() - start_time
        rate = pos / elapsed if elapsed > 0 else 0
        remaining = (total - pos) / rate if rate > 0 else 0

        print(
            f"  {pos:>5,} / {total:,} "
            f"| tagged: {with_tags:,} "
            f"| errors: {errors} "
            f"| rate: {rate:.1f}/s "
            f"| eta: {remaining/60:.1f} min"
        )

# =====================================================
# SAVE
# =====================================================

print()
print("Saving Raw_Tickets.xlsx with tags column...")

df.to_excel(INPUT_FILE, index=False)

# =====================================================
# SUMMARY
# =====================================================

print()
print("=" * 60)
print("ENRICHMENT COMPLETE")
print("=" * 60)

print(f"Total tickets:        {total:,}")
print(f"API calls succeeded:  {success:,}")
print(f"API calls errored:    {errors:,}")
print(f"Tickets with tags:    {with_tags:,}")

if total > 0:
    pct = round(with_tags / total * 100, 1)
    print(f"Tag coverage:         {pct}%")

print()

if with_tags == 0:
    print(
        "WARNING: No tickets received any tags. "
        "Check that your token has permission to read tags "
        "and that tags actually exist in Zammad."
    )

else:
    print("Next step: run Analyze-ZammadTickets.py")

print()
print("Done.")
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

PER_PAGE = int(config.get("per_page", 100))
REQUEST_TIMEOUT = int(config.get("timeout", 30))
MAX_RETRIES = int(config.get("max_retries", 3))
RETRY_DELAY = int(config.get("retry_delay", 5))

HEADERS = {
    "Authorization": f"Token token={TOKEN}",
    "Content-Type": "application/json"
}

# =====================================================
# OUTPUT FOLDER
# =====================================================

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# =====================================================
# HELPER FUNCTIONS
# =====================================================

def api_get(url):
    """
    HTTP GET with retry logic.
    """

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(
                url,
                headers=HEADERS,
                timeout=REQUEST_TIMEOUT
            )

            response.raise_for_status()

            return response

        except requests.exceptions.RequestException as ex:
            print(
                f"Attempt {attempt} failed: {ex}"
            )

            if attempt < MAX_RETRIES:
                print(f"Retrying in {RETRY_DELAY}s...")
                time.sleep(RETRY_DELAY)

            else:
                raise


def get_paged_endpoint(endpoint):
    """
    Pull all records from a paged Zammad endpoint.
    """

    results = []
    page = 1

    while True:
        url = (
            f"{BASE_URL}/api/v1/{endpoint}"
            f"?page={page}&per_page={PER_PAGE}&expand=false"
        )

        response = api_get(url)

        try:
            data = response.json()
        except ValueError:
            print(
                f"Warning: {endpoint} page {page} "
                f"returned non-JSON. Stopping."
            )
            break

        if not data:
            break

        # Some Zammad endpoints wrap results in dicts
        if isinstance(data, dict):
            if "records" in data:
                data = data["records"]

            elif "assets" in data:
                # /users?expand style responses
                data = list(
                    data.get("assets", {})
                    .get(endpoint.capitalize(), {})
                    .values()
                )

            else:
                data = list(data.values())

        if not isinstance(data, list):
            print(
                f"Warning: {endpoint} page {page} "
                f"returned unexpected format."
            )
            break

        results.extend(data)

        print(
            f"{endpoint}: Page {page} "
            f"({len(data)} records, total {len(results)})"
        )

        if len(data) < PER_PAGE:
            break

        page += 1

    return results


def save_excel(data, filename):
    """
    Save a list of dictionaries to Excel.
    """

    df = pd.DataFrame(data)

    file_path = OUTPUT_DIR / filename

    df.to_excel(
        file_path,
        index=False
    )

    print(f"Saved: {file_path}")

    return df


# =====================================================
# TEST CONNECTION
# =====================================================

print()
print("------------------------------------")
print("Testing Zammad Connection")
print("------------------------------------")

test_url = f"{BASE_URL}/api/v1/tickets?page=1&per_page=1"

response = api_get(test_url)

print("Connection Successful")
print(f"HTTP Status: {response.status_code}")

# =====================================================
# PULL TICKETS
# =====================================================

print()
print("------------------------------------")
print("Downloading Tickets")
print("------------------------------------")

tickets = get_paged_endpoint("tickets")

tickets_df = save_excel(
    tickets,
    "Raw_Tickets.xlsx"
)

print(f"Total Tickets: {len(tickets_df):,}")

# =====================================================
# ENRICH TICKETS WITH TAGS
# =====================================================

print()
print("------------------------------------")
print("Enriching Tickets with Tags")
print("------------------------------------")

for idx, ticket in enumerate(tickets, start=1):

    ticket_id = ticket.get("id")

    if not ticket_id:
        continue

    tag_url = (
        f"{BASE_URL}/api/v1/tags"
        f"?object=Ticket&o_id={ticket_id}"
    )

    try:
        tag_response = api_get(tag_url)
        tag_data = tag_response.json()

        tags = tag_data.get("tags", [])

        ticket["tags"] = ",".join(tags)

    except Exception as ex:
        ticket["tags"] = ""

    if idx % 100 == 0:
        print(
            f"  Tagged {idx:,} / {len(tickets):,} tickets"
        )

# Rewrite Raw_Tickets with tags column

tickets_df = save_excel(
    tickets,
    "Raw_Tickets.xlsx"
)

# =====================================================
# PULL USERS
# =====================================================

print()
print("------------------------------------")
print("Downloading Users")
print("------------------------------------")

try:
    users = get_paged_endpoint("users")

    users_df = save_excel(
        users,
        "Users.xlsx"
    )

    print(f"Total Users: {len(users_df):,}")

except Exception as ex:
    print(f"Users skipped: {ex}")

# =====================================================
# PULL GROUPS
# =====================================================

print()
print("------------------------------------")
print("Downloading Groups")
print("------------------------------------")

try:
    groups = get_paged_endpoint("groups")

    groups_df = save_excel(
        groups,
        "Groups.xlsx"
    )

    print(f"Total Groups: {len(groups_df):,}")

except Exception as ex:
    print(f"Groups skipped: {ex}")

# =====================================================
# PULL ORGANIZATIONS
# =====================================================

print()
print("------------------------------------")
print("Downloading Organizations")
print("------------------------------------")

try:
    organizations = get_paged_endpoint("organizations")

    organizations_df = save_excel(
        organizations,
        "Organizations.xlsx"
    )

    print(
        f"Total Organizations: "
        f"{len(organizations_df):,}"
    )

except Exception as ex:
    print(f"Organizations skipped: {ex}")

# =====================================================
# PULL TICKET STATES
# =====================================================

print()
print("------------------------------------")
print("Downloading Ticket States")
print("------------------------------------")

try:
    states = get_paged_endpoint("ticket_states")

    save_excel(states, "Ticket_States.xlsx")

    print(f"Total States: {len(states):,}")

except Exception as ex:
    print(f"States skipped: {ex}")

# =====================================================
# PULL TICKET PRIORITIES
# =====================================================

print()
print("------------------------------------")
print("Downloading Ticket Priorities")
print("------------------------------------")

try:
    priorities = get_paged_endpoint("ticket_priorities")

    save_excel(priorities, "Ticket_Priorities.xlsx")

    print(f"Total Priorities: {len(priorities):,}")

except Exception as ex:
    print(f"Priorities skipped: {ex}")

# =====================================================
# CONSOLE SUMMARY
# =====================================================

print()
print("=" * 60)
print("ZAMMAD EXPORT COMPLETE")
print("=" * 60)

print(f"Total Tickets Pulled: {len(tickets_df):,}")

print()
print("Generated Files:")

for file in sorted(
    OUTPUT_DIR.glob("*.xlsx")
):
    print(f" - {file.name}")

print()
print("Done.")
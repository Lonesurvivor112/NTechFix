import json
import re
import pandas as pd
from pathlib import Path
from collections import Counter

# =====================================================
# CONFIG
# =====================================================

INPUT_FILE = "output/Raw_Tickets.xlsx"

KEYWORD_CATEGORY_FILE = "config/categories.json"
TAG_CATEGORY_FILE = "config/tag_categories.json"

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# Zammad state IDs
STATE_OPEN = 2
STATE_CLOSED = 4
STATE_MERGED = 5

# =====================================================
# LOAD CATEGORY DEFINITIONS
# =====================================================

with open(
    KEYWORD_CATEGORY_FILE,
    "r",
    encoding="utf-8"
) as f:
    KEYWORD_CATEGORIES = json.load(f)

with open(
    TAG_CATEGORY_FILE,
    "r",
    encoding="utf-8"
) as f:
    TAG_CATEGORIES = json.load(f)

# =====================================================
# REPEAT ISSUE CLEANUP
# =====================================================

EXCLUDED_TITLES = {
    "",
    "-",
    "test",
    "hello",
    "question",
    "notification test",
    "test ticket"
}

EXCLUDED_PATTERNS = [
    "zammad getting started",
    "your daily briefing",
    "your digest email",
    "new response for i.t. new employee checklist",
    "a new hire response has been received",
    "a role change for existing employee response has been received",
    "fw: cancellation request for you"
]

NORMALIZED_TITLE_RULES = [
    {
        "name": "User Offboarding",
        "patterns": [
            r"^user offboarding.*",
            r".*offboarding.*"
        ]
    },
    {
        "name": "User Onboarding",
        "patterns": [
            r"^user onboarding.*",
            r".*new employee checklist.*",
            r".*new hire.*",
            r".*setup new hire.*",
            r".*create new user.*",
            r".*create account.*"
        ]
    },
    {
        "name": "UKG",
        "patterns": [
            r"^ukg$",
            r".*ukg.*",
            r".*hrim.*"
        ]
    },
    {
        "name": "Sandata / SAM / SPOC",
        "patterns": [
            r"^sam$",
            r"^spoc$",
            r".*sandata.*"
        ]
    },
    {
        "name": "Phones",
        "patterns": [
            r"^phone$",
            r"^phone issues$",
            r"^new phone$",
            r"^desk phone$",
            r".*phone not working.*",
            r".*voicemail.*",
            r".*clearly.*",
            r".*pbx.*",
            r".*ata.*",
            r".*extension.*"
        ]
    },
    {
        "name": "Printers",
        "patterns": [
            r"^printer$",
            r"^printers$",
            r".*printer not working.*",
            r".*printer setup.*",
            r".*moriah printer.*",
            r".*scanner.*",
            r".*fax.*",
            r".*copier.*"
        ]
    },
    {
        "name": "Password / MFA",
        "patterns": [
            r".*password reset.*",
            r".*authenticator.*",
            r".*mfa.*",
            r".*locked out.*",
            r".*cannot login.*",
            r".*log in.*",
            r".*login.*"
        ]
    },
    {
        "name": "Microsoft 365",
        "patterns": [
            r"^email$",
            r".*outlook.*",
            r".*teams.*",
            r".*sharepoint.*",
            r".*onedrive.*",
            r".*office 365.*",
            r".*microsoft.*"
        ]
    },
    {
        "name": "Hardware",
        "patterns": [
            r"^laptop$",
            r".*laptop.*",
            r".*computer.*",
            r".*monitor.*",
            r".*keyboard.*",
            r".*mouse.*",
            r".*dock.*",
            r".*webcam.*",
            r".*headset.*"
        ]
    },
    {
        "name": "Network",
        "patterns": [
            r"^vpn$",
            r".*vpn.*",
            r".*wifi.*",
            r".*wi-fi.*",
            r".*internet.*",
            r".*network.*",
            r".*unifi.*",
            r".*switch.*"
        ]
    },
    {
        "name": "Sage",
        "patterns": [
            r"^sage$",
            r".*sage.*"
        ]
    },
    {
        "name": "QuickMAR",
        "patterns": [
            r".*quickmar.*",
            r".*quick mar.*"
        ]
    },
    {
        "name": "Office Move",
        "patterns": [
            r"^office move$",
            r".*office move.*",
            r".*move office.*"
        ]
    }
]

# =====================================================
# FUNCTIONS
# =====================================================

def parse_tags(tag_field):
    if pd.isna(tag_field):
        return []

    if isinstance(tag_field, list):
        raw = tag_field

    else:
        raw = str(tag_field).split(",")

    return [
        t.strip()
        for t in raw
        if t and str(t).strip()
    ]


def categorize_by_tags(tag_list):
    """
    Assign a single category based on the ticket's tags.
    First matching category (by TAG_CATEGORIES order) wins.
    """

    tag_set = {
        t.strip().lower()
        for t in tag_list
    }

    for category, tags in TAG_CATEGORIES.items():
        for tag in tags:
            if tag.strip().lower() in tag_set:
                return category

    return None


def categorize_by_keywords(title):
    title = str(title).lower().strip()

    best_category = "Other"
    best_match_len = 0

    for category, keywords in KEYWORD_CATEGORIES.items():
        for keyword in keywords:
            keyword_text = str(keyword).lower().strip()

            if keyword_text and keyword_text in title:
                if len(keyword_text) > best_match_len:
                    best_match_len = len(keyword_text)
                    best_category = category

    return best_category


def categorize_row(row):
    tag_list = parse_tags(row.get("tags"))

    if tag_list:
        tag_category = categorize_by_tags(tag_list)

        if tag_category:
            return tag_category

    return categorize_by_keywords(row.get("title", ""))


def aging_bucket(days):
    if pd.isna(days):
        return "Unknown"

    if days <= 2:
        return "0-2 Days"

    elif days <= 7:
        return "3-7 Days"

    elif days <= 14:
        return "8-14 Days"

    elif days <= 30:
        return "15-30 Days"

    else:
        return "30+ Days"


def should_exclude_repeat_title(title):
    title_clean = str(title).strip().lower()

    if title_clean in EXCLUDED_TITLES:
        return True

    for pattern in EXCLUDED_PATTERNS:
        if pattern in title_clean:
            return True

    return False


def normalize_repeat_title(title):
    title_clean = str(title).strip()
    title_lower = title_clean.lower()

    if should_exclude_repeat_title(title_lower):
        return None

    title_lower = re.sub(
        r"^(fw:|fwd:|re:)\s*",
        "",
        title_lower
    ).strip()

    title_lower = title_lower.replace(
        "[external]",
        ""
    ).strip()

    for rule in NORMALIZED_TITLE_RULES:
        for pattern in rule["patterns"]:
            if re.match(pattern, title_lower):
                return rule["name"]

    return title_clean.strip()


def ordered_aging_summary(aging_df):
    order = [
        "0-2 Days",
        "3-7 Days",
        "8-14 Days",
        "15-30 Days",
        "30+ Days",
        "Unknown"
    ]

    aging_df["Bucket"] = pd.Categorical(
        aging_df["Bucket"],
        categories=order,
        ordered=True
    )

    return aging_df.sort_values("Bucket")


def build_tag_counts(dataframe):
    counter = Counter()

    for tags in dataframe["ParsedTags"]:
        for tag in tags:
            counter[tag] += 1

    df = pd.DataFrame(
        counter.most_common(),
        columns=[
            "Tag",
            "Tickets"
        ]
    )

    return df


# =====================================================
# LOAD DATA
# =====================================================

print("Loading ticket data...")

df = pd.read_excel(INPUT_FILE)

print(f"Loaded {len(df):,} tickets")

# =====================================================
# CLEAN REQUIRED FIELDS
# =====================================================

if "title" not in df.columns:
    df["title"] = ""

if "id" not in df.columns:
    df["id"] = range(1, len(df) + 1)

if "article_count" not in df.columns:
    df["article_count"] = 0

if "tags" not in df.columns:
    print(
        "WARNING: 'tags' column not found — falling back to title-only categorization."
    )
    df["tags"] = ""

if "time_unit" not in df.columns:
    print(
        "WARNING: 'time_unit' column not found — time-tracked metrics will be zero."
    )
    df["time_unit"] = 0

df["article_count"] = pd.to_numeric(
    df["article_count"],
    errors="coerce"
).fillna(0)

df["time_unit"] = pd.to_numeric(
    df["time_unit"],
    errors="coerce"
).fillna(0)

# =====================================================
# STATE ID NORMALIZATION + MERGED FILTER
# =====================================================

if "state_id" in df.columns:
    df["state_id"] = pd.to_numeric(
        df["state_id"],
        errors="coerce"
    )

    merged_count = len(
        df[df["state_id"] == STATE_MERGED]
    )

    df = df[df["state_id"] != STATE_MERGED].copy()

    print(
        f"Excluded {merged_count:,} merged tickets (state_id={STATE_MERGED})"
    )

else:
    print(
        "WARNING: 'state_id' column not found — "
        "falling back to close_at logic for open/closed."
    )

# =====================================================
# DATE CONVERSION
# =====================================================

for col in [
    "created_at",
    "updated_at",
    "close_at"
]:
    if col in df.columns:
        df[col] = pd.to_datetime(
            df[col],
            errors="coerce",
            utc=True
        )

        df[col] = (
            df[col]
            .dt
            .tz_localize(None)
        )

# =====================================================
# PARSE TAGS
# =====================================================

print("Parsing tags...")

df["ParsedTags"] = df["tags"].apply(parse_tags)

df["Tag Count"] = df["ParsedTags"].apply(len)

# =====================================================
# CATEGORIZATION
# =====================================================

print("Categorizing tickets...")

df["Category"] = df.apply(categorize_row, axis=1)

# =====================================================
# TAGS VIEW
# =====================================================

print("Building Tag view...")

tag_counts_all = build_tag_counts(df)

tag_counts_all.to_excel(
    OUTPUT_DIR / "Top_Tags.xlsx",
    index=False
)

# =====================================================
# CATEGORY COUNTS
# =====================================================

print("Building Category view...")

category_counts = (
    df["Category"]
    .value_counts()
    .reset_index()
)

category_counts.columns = [
    "Category",
    "Tickets"
]

category_counts.to_excel(
    OUTPUT_DIR / "Top_Categories.xlsx",
    index=False
)

# =====================================================
# MONTHLY TREND
# =====================================================

print("Building monthly trend...")

df["Month"] = (
    df["created_at"]
    .dt.to_period("M")
    .astype(str)
)

created_per_month = (
    df.groupby("Month")
    .size()
    .reset_index(name="Created")
)

# Use state_id for closed if available
if "state_id" in df.columns:
    closed_df = df[
        (df["state_id"] == STATE_CLOSED) &
        (df["close_at"].notna())
    ].copy()

else:
    closed_df = df[
        df["close_at"].notna()
    ].copy()

closed_df["ClosedMonth"] = (
    closed_df["close_at"]
    .dt.to_period("M")
    .astype(str)
)

closed_per_month = (
    closed_df.groupby("ClosedMonth")
    .size()
    .reset_index(name="Closed")
)

monthly = created_per_month.merge(
    closed_per_month,
    left_on="Month",
    right_on="ClosedMonth",
    how="left"
)

monthly["Closed"] = (
    monthly["Closed"]
    .fillna(0)
    .astype(int)
)

monthly["Net"] = (
    monthly["Created"]
    - monthly["Closed"]
)

monthly.drop(
    columns=["ClosedMonth"],
    errors="ignore",
    inplace=True
)

monthly.to_excel(
    OUTPUT_DIR / "Monthly_Trend.xlsx",
    index=False
)

# =====================================================
# MONTHLY TIME SPENT
# =====================================================

monthly_time = (
    df.groupby("Month")
    .agg(
        Tickets=("id", "count"),
        TotalMinutes=("time_unit", "sum")
    )
    .reset_index()
)

monthly_time["TotalHours"] = (
    (monthly_time["TotalMinutes"] / 60).round(1)
)

monthly_time.to_excel(
    OUTPUT_DIR / "Monthly_TimeSpent.xlsx",
    index=False
)

# =====================================================
# OPEN TICKETS
# =====================================================

print("Calculating backlog...")

if "state_id" in df.columns:
    open_df = df[
        df["state_id"] == STATE_OPEN
    ].copy()

else:
    open_df = df[
        df["close_at"].isna()
    ].copy()

now_utc_naive = pd.Timestamp.utcnow().tz_localize(None)

open_df["Days Open"] = (
    (
        now_utc_naive -
        open_df["created_at"]
    )
    .dt
    .total_seconds()
    .div(86400)
    .fillna(0)
    .astype(int)
)

open_df["Days Open"] = (
    open_df["Days Open"]
    .clip(lower=0)
)

open_df["Aging Bucket"] = (
    open_df["Days Open"]
    .apply(aging_bucket)
)

# Serialize tag list to comma string for Excel
open_df["Tags"] = open_df["ParsedTags"].apply(
    lambda tags: ", ".join(tags)
)

open_df_for_export = open_df.drop(
    columns=["ParsedTags"],
    errors="ignore"
)

open_df_for_export.to_excel(
    OUTPUT_DIR / "Open_Tickets.xlsx",
    index=False
)

# =====================================================
# BACKLOG AGING
# =====================================================

aging_summary = (
    open_df["Aging Bucket"]
    .value_counts()
    .reset_index()
)

aging_summary.columns = [
    "Bucket",
    "Open Tickets"
]

aging_summary = ordered_aging_summary(aging_summary)

aging_summary.to_excel(
    OUTPUT_DIR / "Backlog_Aging.xlsx",
    index=False
)

# =====================================================
# HIGH RISK BACKLOG
# =====================================================

print("Building high-risk backlog...")

high_risk_backlog = (
    open_df[
        open_df["Days Open"] > 30
    ]
    .copy()
    .sort_values(
        "Days Open",
        ascending=False
    )
)

high_risk_for_export = high_risk_backlog.drop(
    columns=["ParsedTags"],
    errors="ignore"
)

high_risk_for_export.to_excel(
    OUTPUT_DIR / "High_Risk_Backlog.xlsx",
    index=False
)

# =====================================================
# HIGH RISK BY CATEGORY
# =====================================================

if not high_risk_backlog.empty:
    high_risk_by_category = (
        high_risk_backlog
        .groupby("Category")
        .agg(
            Tickets=("id", "count"),
            TotalMinutes=("time_unit", "sum")
        )
        .reset_index()
    )

    high_risk_by_category["TotalHours"] = (
        (high_risk_by_category["TotalMinutes"] / 60).round(1)
    )

    high_risk_by_category = high_risk_by_category.rename(
        columns={"Tickets": "30+ Day Tickets"}
    )

    high_risk_by_category = high_risk_by_category.sort_values(
        "30+ Day Tickets",
        ascending=False
    )

else:
    high_risk_by_category = pd.DataFrame(
        columns=[
            "Category",
            "30+ Day Tickets",
            "TotalMinutes",
            "TotalHours"
        ]
    )

high_risk_by_category.to_excel(
    OUTPUT_DIR / "High_Risk_By_Category.xlsx",
    index=False
)

# =====================================================
# OPEN CATEGORY COUNTS
# =====================================================

open_category_counts = (
    open_df["Category"]
    .value_counts()
    .reset_index()
)

open_category_counts.columns = [
    "Category",
    "Open Tickets"
]

open_category_counts.to_excel(
    OUTPUT_DIR / "Open_Categories.xlsx",
    index=False
)

# =====================================================
# OPEN TAG COUNTS
# =====================================================

if not open_df.empty:
    open_tag_counts = build_tag_counts(open_df)

else:
    open_tag_counts = pd.DataFrame(
        columns=[
            "Tag",
            "Tickets"
        ]
    )

open_tag_counts.rename(
    columns={
        "Tickets": "Open Tickets"
    },
    inplace=True
)

open_tag_counts.to_excel(
    OUTPUT_DIR / "Open_Tags.xlsx",
    index=False
)

# =====================================================
# OPEN TIME SPENT BY CATEGORY
# =====================================================

print("Building open time-spent report...")

if not open_df.empty:
    open_time = (
        open_df
        .groupby("Category")
        .agg(
            OpenTickets=("id", "count"),
            TicketsWithTime=(
                "time_unit",
                lambda s: int((s > 0).sum())
            ),
            TotalMinutes=("time_unit", "sum"),
            AvgMinutesPerTicket=("time_unit", "mean")
        )
        .reset_index()
    )

    open_time["AvgMinutesPerTicket"] = (
        open_time["AvgMinutesPerTicket"].round(1)
    )

    open_time["TotalHours"] = (
        (open_time["TotalMinutes"] / 60).round(1)
    )

    open_time = open_time.sort_values(
        "TotalMinutes",
        ascending=False
    )

else:
    open_time = pd.DataFrame(
        columns=[
            "Category",
            "OpenTickets",
            "TicketsWithTime",
            "TotalMinutes",
            "AvgMinutesPerTicket",
            "TotalHours"
        ]
    )

open_time.to_excel(
    OUTPUT_DIR / "Open_TimeSpent.xlsx",
    index=False
)

# =====================================================
# REPEAT PROBLEMS
# =====================================================

print("Finding recurring issues...")

normalized_titles = (
    df["title"]
    .fillna("")
    .apply(normalize_repeat_title)
)

normalized_titles = normalized_titles[
    normalized_titles.notna()
]

repeat_issues = (
    normalized_titles
    .value_counts()
    .reset_index()
)

repeat_issues.columns = [
    "Issue",
    "Count"
]

repeat_issues = repeat_issues[
    repeat_issues["Count"] >= 3
]

repeat_issues.to_excel(
    OUTPUT_DIR / "Repeat_Problems.xlsx",
    index=False
)

# =====================================================
# CATEGORY SUMMARY (time_unit based)
# =====================================================

print("Building category time-spent summary...")

category_summary = (
    df.groupby("Category")
    .agg(
        Tickets=("id", "count"),
        TicketsWithTime=(
            "time_unit",
            lambda s: int((s > 0).sum())
        ),
        TotalMinutes=("time_unit", "sum"),
        AvgMinutesPerTicket=("time_unit", "mean")
    )
    .reset_index()
)

category_summary["AvgMinutesPerTicket"] = (
    category_summary["AvgMinutesPerTicket"].round(1)
)

category_summary["TotalHours"] = (
    (category_summary["TotalMinutes"] / 60).round(1)
)

category_summary["TimeCoverage %"] = (
    (
        category_summary["TicketsWithTime"] /
        category_summary["Tickets"] * 100
    ).round(1)
)

category_summary = category_summary.sort_values(
    "TotalMinutes",
    ascending=False
)

category_summary.to_excel(
    OUTPUT_DIR / "Category_Summary.xlsx",
    index=False
)

# =====================================================
# TAG TIME SPENT SUMMARY (time_unit based)
# =====================================================

print("Building tag time-spent summary...")

tag_time_rows = []

for _, ticket in df.iterrows():
    minutes = ticket.get("time_unit", 0) or 0

    for tag in ticket["ParsedTags"]:
        tag_time_rows.append({
            "Tag": tag,
            "Minutes": minutes
        })

tag_time_df = pd.DataFrame(tag_time_rows)

if not tag_time_df.empty:
    tag_summary = (
        tag_time_df
        .groupby("Tag")
        .agg(
            Tickets=("Minutes", "count"),
            TicketsWithTime=(
                "Minutes",
                lambda s: int((s > 0).sum())
            ),
            TotalMinutes=("Minutes", "sum"),
            AvgMinutesPerTicket=("Minutes", "mean")
        )
        .reset_index()
    )

    tag_summary["AvgMinutesPerTicket"] = (
        tag_summary["AvgMinutesPerTicket"].round(1)
    )

    tag_summary["TotalHours"] = (
        (tag_summary["TotalMinutes"] / 60).round(1)
    )

    tag_summary["TimeCoverage %"] = (
        (
            tag_summary["TicketsWithTime"] /
            tag_summary["Tickets"] * 100
        ).round(1)
    )

    tag_summary = tag_summary.sort_values(
        "TotalMinutes",
        ascending=False
    )

else:
    tag_summary = pd.DataFrame(
        columns=[
            "Tag",
            "Tickets",
            "TicketsWithTime",
            "TotalMinutes",
            "AvgMinutesPerTicket",
            "TotalHours",
            "TimeCoverage %"
        ]
    )

tag_summary.to_excel(
    OUTPUT_DIR / "Tag_Summary.xlsx",
    index=False
)

# =====================================================
# KPI SUMMARY
# =====================================================

kpis = {}

tickets_created = len(df)

if "state_id" in df.columns:
    tickets_closed = len(
        df[df["state_id"] == STATE_CLOSED]
    )

else:
    tickets_closed = len(
        df[df["close_at"].notna()]
    )

open_backlog = len(open_df)
aged_backlog = len(
    open_df[
        open_df["Days Open"] > 30
    ]
)

# Time-spent metrics
total_minutes = int(df["time_unit"].sum())
total_hours = round(total_minutes / 60, 1)
tickets_with_time = int((df["time_unit"] > 0).sum())

if tickets_with_time > 0:
    avg_minutes_per_tracked = round(
        df.loc[df["time_unit"] > 0, "time_unit"].mean(),
        1
    )
else:
    avg_minutes_per_tracked = 0

time_coverage_pct = (
    round(tickets_with_time / tickets_created * 100, 1)
    if tickets_created > 0 else 0
)

kpis["Tickets Created"] = tickets_created
kpis["Tickets Closed"] = tickets_closed
kpis["Open Backlog"] = open_backlog
kpis["30+ Day Backlog"] = aged_backlog

if open_backlog > 0:
    kpis["30+ Day Backlog %"] = round(
        aged_backlog / open_backlog * 100,
        1
    )

    kpis["Oldest Open Ticket Days"] = int(
        open_df["Days Open"].max()
    )

kpis["Total Time Logged (min)"] = total_minutes
kpis["Total Time Logged (hrs)"] = total_hours
kpis["Tickets With Time Logged"] = tickets_with_time
kpis["Time Logging Coverage %"] = time_coverage_pct
kpis["Avg Minutes per Tracked Ticket"] = avg_minutes_per_tracked

if not open_category_counts.empty:
    top_open_category = open_category_counts.iloc[0]

    kpis["Top Open Category"] = (
        f"{top_open_category['Category']} "
        f"({top_open_category['Open Tickets']})"
    )

if not category_counts.empty:
    top_category = category_counts.iloc[0]

    kpis["Top Historical Category"] = (
        f"{top_category['Category']} "
        f"({top_category['Tickets']})"
    )

if not tag_counts_all.empty:
    top_tag = tag_counts_all.iloc[0]

    kpis["Top Tag"] = (
        f"{top_tag['Tag']} "
        f"({top_tag['Tickets']})"
    )

# Ticket tagging coverage
tickets_with_tags = int(
    (df["Tag Count"] > 0).sum()
)

if tickets_created > 0:
    kpis["Ticket Tagging Coverage %"] = round(
        tickets_with_tags / tickets_created * 100,
        1
    )

# Top category / tag by TIME instead of count
if not category_summary.empty:
    top_time_cat = category_summary.iloc[0]

    kpis["Top Category by Time"] = (
        f"{top_time_cat['Category']} "
        f"({top_time_cat['TotalHours']} hrs)"
    )

if not tag_summary.empty:
    top_time_tag = tag_summary.iloc[0]

    kpis["Top Tag by Time"] = (
        f"{top_time_tag['Tag']} "
        f"({top_time_tag['TotalHours']} hrs)"
    )

kpi_df = pd.DataFrame(
    list(kpis.items()),
    columns=[
        "Metric",
        "Value"
    ]
)

kpi_df.to_excel(
    OUTPUT_DIR / "KPI_Summary.xlsx",
    index=False
)

# =====================================================
# CONSOLE SUMMARY
# =====================================================

print()
print("=" * 60)
print("ZAMMAD ANALYSIS COMPLETE")
print("=" * 60)

for key, value in kpis.items():
    print(f"{key}: {value}")

print()
print("Generated Files:")

for file in sorted(
    OUTPUT_DIR.glob("*.xlsx")
):
    print(f" - {file.name}")

print()
print("Done.")
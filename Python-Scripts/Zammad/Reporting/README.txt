# Zammad Reports Suite

A collection of Python scripts used to extract ticket data from Zammad, enrich tickets with tag information, categorize tickets, and generate executive, monthly, and tag-based reporting.

---

# Features

- Export all tickets from Zammad
- Enrich tickets with Zammad tags
- Categorize tickets using keywords and tags
- Analyze ticket trends and workload
- Generate Excel datasets for reporting
- Generate Executive PowerPoint reports
- Generate Monthly PowerPoint reports
- Generate Tag Analysis PowerPoint reports

---

# Directory Structure

```text
ZammadReports/
│
├── config/
│   ├── settings.json
│   ├── categories.json
│   └── tag_categories.json
│
├── output/
│
├── reports/
│
├── Generate-ZammadReport.py
├── Enrich-ZammadTags.py
├── Analyze-ZammadTicket.py
├── Build-ExecutiveReport.py
├── Build-MonthlyReport.py
└── Build-TagReport.py
```

---

# Configuration

## settings.json

Create a file called:

```text
config/settings.json
```

Example:

```json
{
    "base_url": "https://your-zammad-instance.com",
    "token": "YOUR_ZAMMAD_API_TOKEN"
}
```

### Required API Permissions

The API token should have access to:

- Tickets
- Users
- Organizations
- Groups
- Ticket States
- Ticket Tags

---

# Workflow

The reporting process follows this order:

```text
Generate-ZammadReport.py
            ↓
Enrich-ZammadTags.py
            ↓
Analyze-ZammadTicket.py
            ↓
Build Reports
```

---

# Step 1 - Export Data From Zammad

Run:

```powershell
python .\Generate-ZammadReport.py
```

This script:

- Connects to Zammad
- Downloads all tickets
- Downloads supporting reference data
- Creates the master data file

Generated file:

```text
output\Raw_Tickets.xlsx
```

---

# Step 2 - Configure Categories

The analyzer uses two category mapping files:

```text
config\categories.json
config\tag_categories.json
```

These files determine how tickets are grouped into management-friendly categories.

---

## categories.json

Used when categorizing tickets by title keywords.

Example:

```json
{
    "Microsoft 365": [
        "outlook",
        "teams",
        "sharepoint",
        "onedrive"
    ]
}
```

The analyzer scans ticket titles and attempts to match keywords.

---

## tag_categories.json

Used when categorizing tickets from existing Zammad tags.

Example:

```json
{
    "Microsoft 365": [
        "outlook",
        "teams",
        "sharepoint"
    ]
}
```

Tag-based categorization is usually more accurate than keyword matching.

---

# Building Category Lists

There is no single correct way to build category mappings.

---

## Method 1 - Build From Ticket Titles

1. Run the ticket export.
2. Open `Raw_Tickets.xlsx`.
3. Review common ticket titles.
4. Create categories based on recurring issue types.

Example:

```text
Printer not working
Scanner issue
Copier offline
```

becomes:

```json
{
    "Printers": [
        "printer",
        "scanner",
        "copier"
    ]
}
```

---

## Method 2 - Build From Zammad Tags

After tags have been enriched:

```powershell
python .\Enrich-ZammadTags.py
```

Open:

```text
output\Raw_Tickets.xlsx
```

Review the `tags` column and group similar tags together.

Example:

```text
outlook
teams
sharepoint
```

becomes:

```json
{
    "Microsoft 365": [
        "outlook",
        "teams",
        "sharepoint"
    ]
}
```

---

## Method 3 - Use AI

You can export ticket titles or tag lists and ask an AI model:

```text
Group these ticket titles into IT support categories and provide keyword lists for each category.
```

or

```text
Group these Zammad tags into management reporting categories.
```

Always review AI-generated categories before using them in production.

---

# Step 3 - Enrich Tickets With Tags

Run:

```powershell
python .\Enrich-ZammadTags.py
```

This script:

- Reads `Raw_Tickets.xlsx`
- Pulls tag information from Zammad
- Adds a `tags` column
- Updates the original Excel file

Result:

```text
output\Raw_Tickets.xlsx
```

now contains Zammad tag information.

---

# Step 4 - Analyze Ticket Data

Run:

```powershell
python .\Analyze-ZammadTicket.py
```

This script:

- Categorizes tickets
- Builds KPIs
- Creates trend analysis
- Creates backlog reports
- Creates category summaries
- Creates tag summaries

---

## Generated Analysis Files

```text
KPI_Summary.xlsx
Top_Categories.xlsx
Top_Tags.xlsx
Open_Tickets.xlsx
Open_Categories.xlsx
Open_Tags.xlsx

Monthly_Trend.xlsx
Monthly_TimeSpent.xlsx

Category_Summary.xlsx
Tag_Summary.xlsx

Open_TimeSpent.xlsx

Repeat_Problems.xlsx

Backlog_Aging.xlsx

High_Risk_Backlog.xlsx
High_Risk_By_Category.xlsx
```

---

# Step 5 - Generate Executive Report

Run:

```powershell
python .\Build-ExecutiveReport.py
```

The Executive Report provides:

- KPI Dashboard
- Ticket Trends
- Backlog Analysis
- Open Ticket Analysis
- Historical Demand
- Time Tracking Metrics
- Recurring Issue Tracking
- Management Recommendations

---

# Step 6 - Generate Monthly Report

Run:

```powershell
python .\Build-MonthlyReport.py
```

The Monthly Report provides:

- Previous 30-Day Summary
- Daily Ticket Trends
- Top Categories
- Top Tags
- Open Backlog Analysis
- Time Tracking Metrics
- Aged Ticket Review
- Operational Recommendations

---

# Step 7 - Generate Tag Report

Run:

```powershell
python .\Build-TagReport.py
```

The Tag Report provides:

- Tag Coverage Statistics
- Top Tags (All Time)
- Top Tags (30 Days)
- Open Tag Analysis
- Time Logged By Tag
- Tag Trending
- Per-Tag Ticket Breakdowns

---

# Example categories.json

```json
{
    "Microsoft 365": [
        "outlook",
        "teams",
        "sharepoint",
        "onedrive"
    ],
    "Hardware": [
        "computer",
        "laptop",
        "monitor",
        "keyboard"
    ],
    "Network": [
        "network",
        "wifi",
        "internet",
        "vpn"
    ],
    "Printers": [
        "printer",
        "scanner",
        "copier"
    ],
    "Security": [
        "password",
        "mfa",
        "phishing",
        "security"
    ],
    "Phones": [
        "phone",
        "voicemail",
        "cellular",
        "extension"
    ],
    "Onboarding": [
        "new hire",
        "new user",
        "account creation"
    ],
    "Offboarding": [
        "termination",
        "offboarding",
        "account removal"
    ],
    "Applications": [
        "software",
        "application",
        "program"
    ]
}
```

---

# Example tag_categories.json

```json
{
    "Microsoft 365": [
        "outlook",
        "teams",
        "sharepoint",
        "onedrive"
    ],
    "Hardware": [
        "hardware",
        "laptop",
        "desktop",
        "monitor"
    ],
    "Network": [
        "network",
        "wifi",
        "internet",
        "vpn"
    ],
    "Printers": [
        "printer",
        "scanner",
        "copier"
    ],
    "Security": [
        "security",
        "password",
        "mfa",
        "phishing"
    ],
    "Phones": [
        "phone",
        "voip",
        "extension",
        "voicemail"
    ],
    "User Lifecycle": [
        "onboarding",
        "offboarding",
        "new-user",
        "termination"
    ],
    "Applications": [
        "software",
        "application",
        "licensing"
    ],
    "Projects": [
        "project",
        "implementation",
        "upgrade"
    ]
}
```

---

# Typical Reporting Run

```powershell
python .\Generate-ZammadReport.py

python .\Enrich-ZammadTags.py

python .\Analyze-ZammadTicket.py

python .\Build-ExecutiveReport.py

python .\Build-MonthlyReport.py

python .\Build-TagReport.py
```

---

# Notes

- Tag-based categorization always takes priority over keyword categorization.
- Better Zammad tagging results in more accurate reporting.
- AI can be used to help build category definitions, but human review is recommended.
- Historical tickets may not contain tags and may fall back to keyword categorization.
- Consistent ticket tagging and time logging significantly improve report quality.

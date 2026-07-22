# SharePoint Site Usage Reporting Guide

End-to-end process for collecting SharePoint site usage data across multiple sites and generating a consolidated executive reporting dashboard.

**Estimated Runtime:** 30-60 minutes depending on the number of sites, MFA prompts, and SharePoint responsiveness.

---

# Overview

This workflow consists of three primary steps:

1. Grant temporary administrative access to all SharePoint sites.
2. Download Site Analytics reports using Playwright.
3. Aggregate all reports into a single Excel dashboard.

After reporting is complete, remove the temporary administrative permissions.

---

# Prerequisites

Verify the following software is installed and functioning.

| Tool | Verify | Install |
|--------|--------|--------|
| PowerShell 5.1+ | `$PSVersionTable.PSVersion` | Included with Windows |
| Node.js 18+ | `node --version` | https://nodejs.org |
| Python 3.10+ | `python --version` | https://python.org |
| Playwright | `npx playwright --version` | See setup section |
| Pandas/OpenPyXL | `python -c "import pandas, openpyxl"` | `pip install pandas openpyxl` |

---

# Required Permissions

You should have:

- SharePoint Administrator permissions
- Microsoft 365 account with MFA enabled
- Permission to temporarily elevate site permissions

---

# Initial Setup

Run these commands one time when setting up a new workstation:

```powershell
npm install -D @playwright/test
npx playwright install chromium

pip install pandas openpyxl
```

---

# Step 1 - Grant Temporary Site Admin Access

The Site Analytics page only renders correctly for site owners or site collection administrators.

Open:

```powershell
SP-Site-CollectAdmin-Switch.ps1
```

Verify:

```powershell
$mode = "Add"
```

Run:

```powershell
.\SP-Site-CollectAdmin-Switch.ps1
```

The script will iterate through all target SharePoint sites and temporarily grant administrative access.

> Important: You should remove these permissions at the end of the process.

---

# Step 2 - Download Site Usage Reports

Run the Playwright data collection script:

```powershell
node .\SharePoint-Usage-Reports\SP-Usage-Report-Playwright.js
```

---

## What Happens

The automation performs the following steps:

1. Opens Chromium.
2. Navigates to SharePoint.
3. Prompts for interactive Microsoft 365 sign-in.
4. Prompts for MFA.
5. Iterates through each site in the configured list.
6. Opens Site Analytics.
7. Downloads the usage report.
8. Saves the workbook locally.

---

## Login Process

When prompted:

```text
>>> Log in + complete MFA in the browser, then press Enter here...
```

Complete authentication in the browser window.

Return to the terminal and press:

```text
Enter
```

to begin processing.

---

## Download Location

Reports will be stored in:

```text
downloads\
```

Example:

```text
downloads\
├── HumanResources.xlsx
├── Finance.xlsx
├── Operations.xlsx
├── Leadership.xlsx
└── _failed.txt
```

The `_failed.txt` file is only created when one or more sites cannot be processed.

---

## Possible Skip Reasons

A site may be skipped if:

- Access is denied
- Site Analytics is unavailable
- The site contains no usage data
- SharePoint returns an error page
- The download button cannot be located

Failed sites are logged for review.

---

## Retrying Failed Sites

Open:

```javascript
const SITES = [
    ...
]
```

Remove successful sites and leave only those listed in:

```text
downloads\_failed.txt
```

Rerun the Playwright script.

---

# Step 3 - Generate the Executive Dashboard

Run the Python aggregation script.

Example:

```powershell
python .\SharePoint-Usage-Reports\SP-Usage-Aggregate.py `
  ".\downloads" `
  ".\Reports\2026-05-01-SP-Aggregate.xlsx"
```

---

## Parameters

### Parameter 1

Input folder containing downloaded reports.

Example:

```text
downloads\
```

---

### Parameter 2

Destination Excel workbook.

Example:

```text
Reports\2026-05-01-SP-Aggregate.xlsx
```

Using the current date in the filename is recommended.

---

# Dashboard Contents

The generated workbook contains:

| Tab | Purpose |
|-------|-------|
| Executive Summary | KPIs and site rankings |
| Daily Trends | Aggregate visitors and visits |
| Devices | Device usage breakdown |
| Popular Content | Most-viewed pages and files |
| Site Tabs | Individual site reporting |

---

# Manual Review Recommendations

Before distributing the report to leadership, review the following items.

## Site Names

Some SharePoint display names may not match official department names.

Rename as necessary.

---

## Archived Sites

Sites with little or no activity may not be relevant.

Remove or hide them if appropriate.

---

## Date Ranges

The report generation date is not necessarily the same as the reporting period.

The exported SharePoint reports typically represent the previous:

```text
90 Days
```

---

## Top Sites Chart

Large sites can dominate executive charts.

Consider:

- Top 5 instead of Top 10
- Excluding company-wide portal sites
- Separating operational and administrative sites

---

## Empty Site Tabs

Low-usage sites may generate nearly empty worksheets.

Hide or remove them before distributing reports.

---

## SharePoint "Not Supported" Values

Some SharePoint reports export certain metrics as:

```text
Not supported
```

The aggregation process treats these values as blank.

This is a SharePoint limitation rather than a script issue.

---

# Step 4 - Security Cleanup

After report generation is complete:

Open:

```powershell
SP-Site-CollectAdmin-Switch.ps1
```

Change:

```powershell
$mode = "Remove"
```

Run:

```powershell
.\SP-Site-CollectAdmin-Switch.ps1
```

This removes the temporary site collection administrator permissions.

---

# Optional Archiving

To maintain historical comparisons:

```text
archives\
└── 2026-05-01\
    ├── downloads\
    └── 2026-05-01-SP-Aggregate.xlsx
```

Store:

- Raw downloads
- Final dashboard
- Notes or executive summaries

---

# Troubleshooting

## Playwright Module Missing

Error:

```text
Cannot find module '@playwright/test'
```

Fix:

```powershell
npm install -D @playwright/test
```

---

## Browser Login Issues

Symptoms:

- Continuous sign-in loop
- Authentication timeout
- MFA failures

Recommended actions:

- Close Chromium completely
- Rerun the script
- Complete MFA quickly
- Verify browser cookies are not blocked

---

## Multiple Access Denied Errors

Possible causes:

- Site admin script was not executed
- Permission propagation delay
- Unsupported site types

Recommendations:

1. Verify `$mode = "Add"`
2. Re-run the admin script
3. Wait 5-10 minutes
4. Retry downloads

---

## Aggregation Errors

Symptoms:

```text
KeyError
Parse Error
```

Possible cause:

One or more downloaded files are not valid SharePoint exports.

Validate that each workbook contains:

- Overall traffic
- Popular content
- Usage by device
- Usage by time

Delete invalid files and rerun aggregation.

---

## Charts Look Incorrect

Open the report in:

```text
Microsoft Excel Desktop
```

rather than Excel Online.

Certain chart types do not fully render in the browser.

---

# File Reference

| File | Purpose |
|--------|--------|
| `SP-Site-CollectAdmin-Switch.ps1` | Adds or removes temporary site collection administrator permissions |
| `SP-Usage-Report-Playwright.js` | Downloads SharePoint Site Analytics reports |
| `SP-Usage-Aggregate.py` | Builds consolidated reporting dashboard |
``
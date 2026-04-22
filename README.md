# Zenoti Sales Accrual Automation — n8n + Power BI 📊

> **Automated MedSpa sales reporting pipeline** — ingesting Zenoti sales accrual reports from email, parsing and cleaning the data, and syncing it into a live SharePoint Excel workbook connected to Power BI.

---

## About

This project was built to eliminate a daily manual reporting bottleneck for a MedSpa client using Zenoti as their practice management system. Because Zenoti delivers reports via authenticated email links rather than an open API, the solution required a custom middleware layer to bridge the gap between Zenoti's delivery mechanism and a structured BI pipeline.

The result is a fully automated, zero-touch data pipeline that keeps a live Power BI dashboard current — giving the client real-time visibility into sales performance without any daily manual effort.

This project reflects hands-on experience building automation workflows against systems with no native API, designing ETL pipelines that handle real-world data messiness, and connecting operational data sources to BI tooling in a production environment.

---

## Overview

This n8n workflow replaces a fully manual reporting process for a MedSpa client operating on the Zenoti practice management platform. Previously, Zenoti emailed a download link for the Sales-Accrual report to the client each day — requiring someone to manually open the link, download the file, clean the data, and paste it into a spreadsheet.

The automation handles the entire pipeline end-to-end: it detects the email, extracts the report link, downloads and parses the file via a custom server, and appends the clean data directly into a SharePoint Excel workbook that feeds a live Power BI dashboard. No human intervention required.

**Status:** Live and running in production — polling the client's Outlook every minute.

---

## The Problem

Zenoti does not offer a native API for scheduled report exports. Their reporting system sends a time-limited download link via email — one that requires an authenticated session to access. This means:

- Reports could only be retrieved manually, one day at a time
- Data lived in email inboxes, not in a structured, queryable format
- The client's Power BI dashboard had no automated data source to connect to
- Delays in manual entry meant the dashboard was frequently stale
- Human copy-paste introduced data formatting inconsistencies

The client needed a hands-off pipeline that captured every report the moment it arrived and made the data immediately available for BI analysis.

---

## Project Structure

```
zenoti-sales-accrual-automation/
├── n8n/
│   └── zenoti_sales_accrual_workflow.json     # Importable n8n workflow
├── server/
│   └── app.py                                 # Custom download + parse microserver
├── powerbi/
│   └── ZenotiSalesAccrual.pbix                # Power BI report file
└── README.md
```

---

## Architecture Highlights

| Concern | Approach |
|---|---|
| **Zenoti auth bypass** | Custom microserver handles the authenticated download that n8n cannot natively perform |
| **Data integrity** | Clear-before-append pattern ensures no duplicate rows accumulate over time |
| **Historical preservation** | The clear range is hard-coded to start below all pre-existing historical data, protecting it on every run |
| **CSV robustness** | Custom parser handles quoted fields, escaped quotes, and CRLF line endings — not just `split(",")` |
| **Date validation** | `yesterdayExists` flag enables downstream alerting if Zenoti skips a day's report |
| **Email noise filtering** | IF node ensures only Zenoti Sales-Accrual emails activate the pipeline |
| **BI connectivity** | SharePoint-hosted workbook provides a stable, always-refreshed data source for Power BI |

---

## How It Works

```
Zenoti emails Sales-Accrual report link to client's Outlook
                        │
                        ▼
        ┌───────────────────────────────┐
        │  n8n — Outlook Trigger        │  ← Polls every 60 seconds
        │  Checks for new emails        │
        └───────────────┬───────────────┘
                        │
                        ▼
        ┌───────────────────────────────┐
        │  IF: Subject contains         │  ← Gates on "Sales-Accrual"
        │  "Sales-Accrual"?             │    Ignores all other email
        └───────┬───────────────────────┘
              true
                │
                ▼
        ┌───────────────────────────────┐
        │  Extract URL (Code Node)      │  ← JS: regex parse Zenoti link
        │  from email body              │    from bodyPreview field
        └───────────────┬───────────────┘
                        │
                        ▼
        ┌───────────────────────────────┐
        │  Custom Server :5000/download │  ← Authenticated download bot
        │  POST with Zenoti URL         │    Returns internal download_url
        └───────────────┬───────────────┘
                        │
                        ▼
        ┌───────────────────────────────┐
        │  Custom Server :5000/parse    │  ← Parses Excel → CSV
        │  POST with download_url       │    Returns raw CSV string
        └───────────────┬───────────────┘
                        │
                        ▼
        ┌───────────────────────────────┐
        │  Clear Sheet (Excel)          │  ← Wipes append zone in SharePoint
        │                               │    Preserves historical data above
        └───────────────┬───────────────┘
                        │
                        ▼
        ┌───────────────────────────────┐
        │  Process CSV (Code Node)      │  ← Full CSV parse, date sort,
        │  Clean + validate data        │    yesterday existence check
        └───────────────┬───────────────┘
                        │
                        ▼
        ┌───────────────────────────────┐
        │  Append to Sheet (Excel)      │  ← Writes all rows to SharePoint
        │                               │    30+ columns mapped field-by-field
        └───────────────────────────────┘
                        │
                        ▼
        ┌───────────────────────────────┐
        │  Power BI Dashboard           │  ← Connected live to SharePoint
        │  Live Sales Reporting         │    Auto-refreshes from workbook
        └───────────────────────────────┘
```

---

## Tech Stack

### 🔄 n8n — Workflow Automation

| Setting | Value |
|---|---|
| Workflow Name | `Zenoti Sales Accrual Automation v1.1` |
| Trigger | Microsoft Outlook (polls every minute) |
| Execution Order | v1 |
| Binary Mode | Separate |
| Status | Active |

**Nodes summary:**

| Node | Type | Role |
|---|---|---|
| `Microsoft Outlook Trigger` | Trigger | Detects incoming Zenoti report emails |
| `IF` | Filter | Gates on `"Sales-Accrual"` in subject |
| `Extracting link from the email` | Code (JS) | Parses Zenoti download URL from body |
| `Sending bot to download the link` | HTTP Request | Sends URL to custom download server |
| `Downloading url / sheet` | HTTP Request | Retrieves parsed CSV from custom server |
| `Clear sheet` | Microsoft Excel | Wipes the append zone before rewrite |
| `Processing csvs and cleaning the data` | Code (JS) | Full CSV parse, clean, sort, date-check |
| `Append data to sheet` | Microsoft Excel | Writes all rows to SharePoint workbook |

---

### 🤖 Custom Microserver (Python — `[server-ip]:5000`)

A self-hosted VPS server that handles the two operations n8n cannot do natively against Zenoti:

| Endpoint | Method | Purpose |
|---|---|---|
| `/download` | POST | Accepts a Zenoti report URL, authenticates, downloads the file, returns an internal `download_url` |
| `/parse-excel` | POST | Accepts a `download_url`, reads the Excel file, returns raw CSV string |

This server bridges the gap between Zenoti's email-gated report system and n8n's HTTP-native world.

> ⚠️ **Security note:** The server currently operates over plain HTTP (`http://`). For production use with client data, this endpoint should be secured with TLS (HTTPS) and API key authentication.

---

### 📊 Microsoft Excel (SharePoint)

| Setting | Value |
|---|---|
| Workbook | `[client workbook name].xlsx` |
| Sheet | `Sheet1` |
| Host | SharePoint Online (`[client SharePoint domain]`) |
| Auth | Microsoft Excel OAuth2 (`Client Excel` connection) |
| Preserved range | Rows 1–N (historical data, never touched) |
| Append zone | Rows N+1 onwards (cleared and rewritten on each run) |
| Column count | 30+ mapped fields across columns A–AV |

---

### 📈 Power BI

| Setting | Value |
|---|---|
| Data source | SharePoint Excel workbook (live connection) |
| Refresh | Automatic — triggered by SharePoint file updates |
| Report scope | MedSpa sales accrual — daily, item-level granularity |
| Key dimensions | Sale Date, Item Type, Item Category, Employee, Center, Brand Name |
| Key metrics | Gross Sale, Net Sale, Discount, Tax, Collection, Sales (Exc. Redemption) |

---

## Node-by-Node Breakdown

### 1. Microsoft Outlook Trigger
Polls the client's Outlook account (`Client Outlook` OAuth2 connection) every minute using the `microsoftOutlookTrigger`. Set to `alwaysOutputData: true` — it fires on every poll cycle regardless of whether a new email arrived, relying on the downstream IF node to filter out irrelevant runs.

---

### 2. IF — Subject Filter
The pipeline gatekeeper. Checks whether the incoming email subject contains `"Sales-Accrual"` (case-sensitive, strict type validation). Only the `true` branch is wired forward — the `false` branch terminates silently. This ensures the workflow only activates on Zenoti report emails, never on general inbox traffic.

Example subject that triggers the flow:
```
Sales-Accrual report requested on [date] [time]
```

---

### 3. Extract URL (Code — JavaScript)
Parses the email `bodyPreview` field using a regex pattern `/(https?:\/\/[^\s]+)/gi` to extract all URLs. Prioritizes URLs containing `"DownloadEmailReport"` or `"zenoti.com"` — the identifying markers of a Zenoti report link. Falls back to the first URL found if no Zenoti-specific match exists.

**Output:**
```json
{
  "extractedUrl": "https://[client].zenoti.com/Admin/Reports/DownloadEmailReport.aspx?Details=...",
  "urlFound": true
}
```

---

### 4. Download Bot — POST `/download`
Sends the extracted Zenoti URL to a self-hosted microservice at `[server-ip]:5000`. This server runs an authenticated session capable of accessing Zenoti's authenticated report download links — something a plain HTTP request cannot do. Timeout is set to 10 minutes to accommodate slow Zenoti server responses.

**Request:**
```json
{ "url": "{{ extractedUrl }}" }
```

**Response includes:** an internal `download_url` pointing to the now-downloaded file on the custom server.

---

### 5. Parse Excel — POST `/parse-excel`
Sends the `download_url` from the previous step back to the same microservice, which reads the downloaded Excel file and returns its contents as a CSV string. The two-step design (download then parse) is intentional — the file is saved server-side first, then processed separately, which gives the server flexibility to handle large or slow-loading files without blocking.

---

### 6. Clear Sheet (Microsoft Excel)
Before any new data is written, the workflow clears a defined range in the SharePoint workbook (Sheet1). The range is hard-coded to start below the last row of preserved historical data — and clears only the "fresh" append zone before rewriting it.

**Workbook:** `[client workbook name].xlsx`
**SharePoint location:** `[client SharePoint domain]`

---

### 7. Process CSV (Code — JavaScript)
The most complex node. Performs the following in sequence:

- Strips the UTF-8 BOM character (`\uFEFF`) that Excel sometimes prepends
- Runs a full, spec-compliant CSV parser — handles quoted fields, escaped double-quotes, and Windows-style `\r\n` line endings
- Maps each row to a JSON object keyed by the CSV header row
- Sorts all rows ascending by `Sale Date`
- Checks whether yesterday's date exists in the dataset and flags every row with a `yesterdayExists` boolean

This flag can be used downstream to detect missing daily data or trigger alerts if Zenoti didn't send a report for the previous day.

---

### 8. Append to Sheet (Microsoft Excel)
Writes every cleaned row to the SharePoint workbook, mapping 30+ fields from the parsed CSV to named Excel columns including:

`Item Type` · `Sale Date` · `Invoice No` · `Employee` · `Center` · `Item Name` · `Item Category` · `Quantity` · `Unit Price` · `Gross Sale` · `Discount` · `Net Sale` · `Tax` · `Collection` · `Brand Name` · `Sales (Exc. Redemption)` · `Invoice Source` · `Non Taxable Redemption` · and more.

---

### 9. Power BI Dashboard *(downstream)*
The SharePoint workbook serves as the live data source for a Power BI report. Power BI connects directly to the Excel file via SharePoint, and the dataset refreshes automatically whenever new rows are appended — giving the client an always-current view of MedSpa sales performance without any manual data work.

---

## Setup / Replication

> ⚠️ This workflow integrates with Microsoft 365, a self-hosted Python server, and Zenoti. You will need active credentials for each.

### Prerequisites
- n8n instance (self-hosted or cloud)
- Microsoft 365 account with Outlook and SharePoint/OneDrive access
- Self-hosted server (VPS or equivalent) for the download + parse microservice
- Zenoti account with Sales-Accrual email reports enabled
- Power BI Desktop or Power BI Service with SharePoint connector

### Steps
1. Import `n8n/zenoti_sales_accrual_workflow.json` into your n8n instance
2. Configure the `Microsoft Outlook` OAuth2 credential in n8n (target inbox)
3. Configure the `Microsoft Excel` OAuth2 credential in n8n (SharePoint workbook access)
4. Deploy `server/app.py` to your VPS — update the IP/port in the HTTP Request nodes
5. Update the SharePoint workbook ID and worksheet ID in the Excel nodes to match your file
6. Activate the workflow — it will begin polling on the next minute cycle
7. In Power BI, connect to the SharePoint Excel file as a data source and publish the report

---

## Current Capabilities

- [x] Automated email detection — polls Outlook every 60 seconds
- [x] Subject-line filtering — only processes Zenoti Sales-Accrual reports
- [x] Authenticated Zenoti report download via custom microserver
- [x] Excel parsing → structured CSV conversion
- [x] Full CSV cleaning — BOM strip, quoted field handling, date normalization
- [x] Safe append pattern — historical data preserved, fresh zone rewritten on each run
- [x] 30+ column field mapping to SharePoint Excel
- [x] Yesterday date existence check for data completeness validation
- [x] Live Power BI dashboard connected to SharePoint data source
- [x] Automated Power BI dataset refresh — scheduled daily via Power BI Apps

---

## Roadmap

- [ ] HTTPS + API key auth on the custom microserver
- [ ] Slack/email alert if yesterday's data is missing (`yesterdayExists = false`)
- [ ] Multi-location support (currently single MedSpa center)
- [ ] Historical backfill utility for gap recovery
- [ ] Expand to additional Zenoti report types (e.g. Inventory, Payroll)

---

*Built with n8n · Microsoft Excel (SharePoint) · Python · Power BI*

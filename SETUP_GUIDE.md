# Kodely Hours Verification v3 - Setup Guide

## What This Does

Automatically calculates the **maximum allowed hours** for each leader based on:
- **Check-ins sheet** → who actually showed up (Status = Leader, Co-Lead, SUB, SCOOT)
- **Ops Hub** → workshop start/end times (duration)
- **Rule**: Allowed hours = Workshop Duration + 30 minutes

### Features:
1. **Auto-Build Verification Sheet** — pulls Check-ins + Ops Hub and generates the report
2. **Gusto CSV Import + Auto-Compare** — paste Gusto export, auto-match names, flag overages
3. **Slack Alerts** — daily reports at 8 PM + weekly payroll summaries on Fridays
4. **Hours Tracker** — auto-generated, color-coded tabs grouped by region (leaders + SCOOT separate)

## Setup (10 minutes)

### Step 1: Open Apps Script
1. Open your **Hours Verification** Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code in the editor

### Step 2: Paste the Script
1. Open the file `HoursVerification.gs` from this folder
2. Copy the entire contents
3. Paste it into the Apps Script editor
4. Click **Save**

### Step 3: Verify Spreadsheet IDs
The script already has your spreadsheet IDs pre-configured:
- Ops Hub: `17hnG_MZs81GFoz_lyJNVyMYocrZLweGkZ6fTxn5UCTs`
- Check-ins: `1rHQ1YNUUkTcmUWwy1tXksM6tWG0GUoOOG3pRrsv9DKs`
- Hours Verification: `1mEavc208nXVyiOoW2ISiRm_Zs7ipbsqau5shw8aQ7YU`

If any of these have changed, update the IDs at the top of the script.

### Step 4: Authorize
1. Click **Run** at the top
2. Select `run02_05to02_18` from the dropdown
3. Google will ask you to authorize — click **Review Permissions**
4. Select your Google account
5. Click **Advanced → Go to Untitled Project (unsafe)** (this is normal for custom scripts)
6. Click **Allow**

### Step 5: Configure Slack (Phase 3)
1. Go to api.slack.com → **Your Apps** → **Create New App**
2. Enable **Incoming Webhooks**
3. Add the webhook to your HR/ops channel
4. Copy the webhook URL
5. In the script, replace `YOUR_SLACK_WEBHOOK_URL_HERE` with your webhook URL

## How to Use

### Menu Options
After refreshing the sheet, you'll see the **Hours Verification** menu with:

| Menu Item | What It Does |
|-----------|-------------|
| Run 02/05 - 02/18 | Generate report for current pay period |
| Run Custom Date Range... | Generate report for any date range |
| Setup Gusto Import Tab | Creates the "Gusto Import" tab with instructions |
| Compare Gusto Hours | Reads Gusto data, matches names, writes Discrepancies report |
| Generate Hours Tracker 02/05 - 02/18 | Generate region-grouped "Hours Tracker" + "SCOOT Hours" tabs |
| Generate Hours Tracker (Custom Range)... | Same as above for any date range |
| Send Daily Slack Alert | Manually trigger the daily Slack report |
| Send Weekly Slack Summary | Manually trigger the weekly summary |
| Setup Auto Triggers | Enable daily (8 PM) + weekly (Friday 6 PM) auto-alerts |
| Remove All Triggers | Disable all auto-alerts |

### Phase 1: Generate the Report
- Click **Hours Verification → Run 02/05 - 02/18** (or custom range)
- Check the **"Auto Verification"** tab for the full report

### Phase 2: Gusto Comparison
1. Click **Hours Verification → Setup Gusto Import Tab** (first time only)
2. Export hours from Gusto: **People → Time Tools → Export**
3. Paste the exported data into the **"Gusto Import"** tab (Column A = names, Column B = hours)
4. Click **Hours Verification → Compare Gusto Hours**
5. Check the **"Discrepancies"** tab for:
   - **Flagged Leaders** (red) — over their allowed hours, sorted by overage
   - **All Comparisons** — every Gusto entry matched against allowed hours
   - **Unmatched Names** — Gusto names that couldn't be matched to check-ins

### Phase 3: Slack Alerts
1. Configure your webhook URL (Step 5 above)
2. Click **Hours Verification → Setup Auto Triggers** to enable:
   - **Daily at 8 PM** — summary of leaders + flagged overages
   - **Weekly on Friday at 6 PM** — full payroll verification digest
3. Or send alerts manually from the menu

### Hours Tracker (Region-Grouped Tabs)

The Hours Tracker generates two color-coded, region-grouped tabs from the same check-in data:

- **"Hours Tracker"** tab — all leaders EXCEPT SCOOT, grouped by region
- **"SCOOT Hours"** tab — SCOOT people only (for invoice matching)

#### How It Works
1. Click **Hours Verification → Generate Hours Tracker 02/05 - 02/18** (or custom range)
2. The script loads check-ins and Ops Hub data (same sources as the main report)
3. Sessions are split by status: SCOOT sessions go to the SCOOT tab, everything else goes to Hours Tracker
4. A person with both leader AND scoot sessions appears in **both** tabs (only their relevant sessions)

#### Tab Layout

Each tab has two sections:

**Section A: Summary** (top)
- One row per leader, grouped by region with color-coded tint backgrounds
- Columns: Region, Leader Name, Sessions, Total Hours, Formatted, Unmatched, Status

**Section B: Detailed Session Log** (below)
- Color-coded region header rows (rotating 8-color palette)
- One row per workshop session, sorted by date within each leader
- Columns: Leader Name, Date, Workshop, School, Status, Duration, Allowed, Source
- Bold subtotal rows for each leader with total hours
- Unmatched workshops highlighted yellow

#### No Setup Required
The Hours Tracker uses the same spreadsheet IDs and data sources already configured for the main report. No additional setup needed.

## Troubleshooting

**"I don't see the Hours Verification menu"**
→ Refresh the sheet. The menu loads when the sheet opens.

**"Workshop duration shows as DEFAULT (1hr)"**
→ The script couldn't match the check-in's school/workshop name to the Ops Hub. Check for spelling differences between the two sheets.

**"No data in the report"**
→ Make sure the Check-ins sheet has data for the selected date range.

**"Gusto names not matching"**
→ The script handles "Last, First" vs "First Last" formats and fuzzy matching. If a name still doesn't match, check for typos or nicknames. Unmatched names appear at the bottom of the Discrepancies tab.

**"Slack alerts not sending"**
→ Make sure the webhook URL is correctly set in the script (line 17). Test manually first via the menu before setting up triggers.

**"Hours Tracker shows 'Unknown' region"**
→ The check-in records are missing a region value. Make sure the Check-ins sheet has a "Region" column populated for each record.

**"Authorization error"**
→ You need edit access to all 3 spreadsheets. Ask the sheet owners to grant you access.

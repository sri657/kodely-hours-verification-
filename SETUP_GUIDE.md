# Kodely Hours Verification v3 - Setup Guide

## What This Does

Automatically calculates the **maximum allowed hours** for each leader based on:
- **Check-ins sheet** → who actually showed up (Status = Leader, Co-Lead, SUB, SCOOT)
- **Ops Hub** → workshop start/end times (duration)
- **Rule**: Allowed hours = Workshop Duration + 30 minutes

### Four Phases:
1. **Auto-Build Verification Sheet** — pulls Check-ins + Ops Hub and generates the report
2. **Gusto CSV Import + Auto-Compare** — paste Gusto export, auto-match names, flag overages
3. **Slack Alerts** — daily reports at 8 PM + weekly payroll summaries on Fridays
4. **Gusto API Integration** — pull hours directly from Gusto with one click (no CSV export needed)

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
| Authorize Gusto | Connect to Gusto via OAuth (one-time setup) |
| Sync Hours from Gusto | Pull hours directly from Gusto API into "Gusto Import" tab |
| Gusto Connection Status | Check if Gusto is connected and show config details |
| Disconnect Gusto | Clear OAuth tokens (for re-auth or switching accounts) |
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

### Phase 4: Gusto API Integration (Optional — replaces manual CSV export)

#### Prerequisites
1. **Register as a Gusto developer** at [developer.gusto.com](https://developer.gusto.com)
2. **Create a Gusto application** to get your `client_id` and `client_secret`
3. **Add the OAuth2 library** in the Apps Script editor:
   - Go to **Libraries** (+ icon in left sidebar)
   - Enter Script ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`
   - Click **Look up** → select latest version → **Add**
4. **Deploy as a web app** (needed for OAuth callback):
   - In Apps Script, click **Deploy → New deployment**
   - Select type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone** (required for OAuth callback)
   - Click **Deploy** and copy the deployment URL
5. **Configure Gusto app redirect URI**:
   - In your Gusto developer dashboard, set the redirect URI to:
     `https://script.google.com/macros/d/{YOUR_SCRIPT_ID}/usercallback`
   - You can find the exact URI by running **Gusto Connection Status** from the menu
6. **Set Script Properties** (in Apps Script: **Project Settings → Script Properties**):
   - `GUSTO_CLIENT_ID` — from your Gusto app
   - `GUSTO_CLIENT_SECRET` — from your Gusto app
   - `GUSTO_COMPANY_UUID` — your Gusto company UUID (found in Gusto admin URL or API)

#### Sandbox vs Production
The script defaults to **sandbox mode** (`GUSTO_USE_SANDBOX = true`) for safe testing. To switch to production:
1. Open `HoursVerification.gs`
2. Change `GUSTO_USE_SANDBOX` from `true` to `false`
3. Run **Disconnect Gusto** (tokens are environment-specific)
4. Run **Authorize Gusto** again with your production Gusto credentials

#### Authorization Flow
1. Click **Hours Verification → Authorize Gusto**
2. A dialog opens with an authorization link — click it
3. Log in to Gusto and grant access
4. You'll see a success page — close the tab and return to your spreadsheet
5. Click **Gusto Connection Status** to verify it says "Connected"

#### Syncing Hours
1. Click **Hours Verification → Sync Hours from Gusto**
2. Enter the start and end dates for the pay period
3. The script fetches hours from Gusto and writes them to the "Gusto Import" tab
4. You'll be prompted to auto-run "Compare Gusto Hours" — click Yes to generate the Discrepancies report

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

**"Gusto credentials not configured"**
→ Set `GUSTO_CLIENT_ID`, `GUSTO_CLIENT_SECRET`, and `GUSTO_COMPANY_UUID` in Script Properties (Project Settings → Script Properties).

**"Not authorized with Gusto"**
→ Run **Authorize Gusto** from the menu. If it was previously connected, tokens may have expired — disconnect and re-authorize.

**"Gusto authorization expired"**
→ OAuth tokens expire. Run **Disconnect Gusto** then **Authorize Gusto** to get fresh tokens.

**"No time sheet data found"**
→ Check that the date range matches hours logged in Gusto. If using sandbox mode, make sure you have test data in the Gusto demo environment.

**"OAuth2 is not defined" error**
→ The OAuth2 library is not added. In Apps Script, go to Libraries → Add → paste `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF` → Add.

**"Authorization error"**
→ You need edit access to all 3 spreadsheets. Ask the sheet owners to grant you access.

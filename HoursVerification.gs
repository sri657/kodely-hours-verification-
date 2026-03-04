// =============================================================
// KODELY HOURS VERIFICATION v3
// Phase 1: Auto-Build Verification Sheet (DONE)
// Phase 2: Gusto CSV Import + Auto-Compare
// Phase 3: Slack Daily/Weekly Alerts
// =============================================================

const OPS_HUB_ID = '17hnG_MZs81GFoz_lyJNVyMYocrZLweGkZ6fTxn5UCTs';
const CHECKINS_ID = '1rHQ1YNUUkTcmUWwy1tXksM6tWG0GUoOOG3pRrsv9DKs';
const HOURS_VERIFICATION_ID = '1mEavc208nXVyiOoW2ISiRm_Zs7ipbsqau5shw8aQ7YU';
const BUFFER_MINUTES = 30;
const WORKED_STATUSES = ['leader', 'co-lead', 'sub', 'scoot', 'coordinator'];

// --- Phase 3: Slack Configuration ---
// Replace with your actual Slack Incoming Webhook URL
// Setup: https://api.slack.com/messaging/webhooks
const SLACK_WEBHOOK_URL = 'YOUR_SLACK_WEBHOOK_URL_HERE';
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/' + HOURS_VERIFICATION_ID;

// --- Gusto Import Tab Configuration ---
const GUSTO_IMPORT_TAB = 'Gusto Import';
const DISCREPANCIES_TAB = 'Discrepancies';

// --- Hours Tracker Configuration ---
const HOURS_TRACKER_TAB = 'Hours Tracker';
const SCOOT_HOURS_TAB = 'SCOOT Hours';

// --- Daily Dashboard ---
const DASHBOARD_PAY_PERIOD_START = new Date(2026, 1, 5);       // 02/05/2026
const DASHBOARD_PAY_PERIOD_END = new Date(2026, 1, 18, 23, 59, 59); // 02/18/2026

// SCOOT detection uses the "status" column in check-ins sheet.
// Status = "scoot" → SCOOT tab. Anything else → leader tab.
// To fix a misclassification, update the status in the check-ins sheet.

// Region header colors (bold, saturated — for region header rows)
const REGION_COLORS = [
  '#4285f4', // Blue
  '#34a853', // Green
  '#ea4335', // Red
  '#fbbc04', // Yellow
  '#ff6d01', // Orange
  '#46bdc6', // Teal
  '#9334e6', // Purple
  '#e91e63'  // Pink
];

// Region tint colors (light — for data rows under each region)
const REGION_TINTS = [
  '#d0e0fc', // Light Blue
  '#ceead6', // Light Green
  '#fad2cf', // Light Red
  '#fef7cd', // Light Yellow
  '#ffe0cc', // Light Orange
  '#d4f0f3', // Light Teal
  '#e9d5fb', // Light Purple
  '#fce4ec'  // Light Pink
];

function onOpen(e) {
  try {
    SpreadsheetApp.getUi().createMenu('Hours Verification')
      .addItem('Generate Hours Tracker 02/05 - 02/18', 'generateHoursTracker02_05to02_18')
      .addItem('Generate Hours Tracker (Custom Range)...', 'promptHoursTrackerDateRange')
      .addSeparator()
      .addItem('Setup SCOOT Invoice Tab (paste invoices here)', 'setupScootInvoiceImport')
      .addItem('Verify SCOOT Invoices', 'verifyScootInvoices')
      .addSeparator()
      .addItem('Enable Auto-Refresh (every hour)', 'enableAutoRefresh')
      .addItem('Disable Auto-Refresh', 'disableAutoRefresh')
      .addSeparator()
      .addItem('Send Dashboard Email...', 'sendDashboardPrompt')
      .addSeparator()
      .addItem('Create "How It Works" Guide', 'createHowItWorksTab')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen: no UI context — ' + err.message);
  }
}

// ---- MAIN ENTRY: 02/05-02/18 ----
function run02_05to02_18() {
  var startDate = new Date(2026, 1, 5);
  var endDate = new Date(2026, 1, 18, 23, 59, 59);
  runReport(startDate, endDate);
}

function promptDateRange() {
  var ui = SpreadsheetApp.getUi();
  var s = ui.prompt('Start date (MM/DD/YYYY):');
  if (s.getSelectedButton() !== ui.Button.OK) return;
  var e = ui.prompt('End date (MM/DD/YYYY):');
  if (e.getSelectedButton() !== ui.Button.OK) return;
  var sd = new Date(s.getResponseText());
  var ed = new Date(e.getResponseText());
  ed.setHours(23, 59, 59);
  runReport(sd, ed);
}

// =====================================================
// MAIN REPORT FUNCTION
// =====================================================
function runReport(startDate, endDate) {
  var log = [];
  log.push('=== HOURS VERIFICATION RUN ===');
  log.push('Period: ' + startDate.toDateString() + ' to ' + endDate.toDateString());

  // STEP 1: Load Ops Hub
  var opsHub = {};
  try {
    var opsSS = SpreadsheetApp.openById(OPS_HUB_ID);
    var opsSheets = opsSS.getSheets();
    log.push('\nOps Hub: ' + opsSS.getName() + ' (' + opsSheets.length + ' tabs)');

    for (var s = 0; s < opsSheets.length; s++) {
      var sheet = opsSheets[s];
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      var headers = [];
      for (var h = 0; h < data[0].length; h++) {
        headers.push(String(data[0][h]).toLowerCase().trim());
      }

      var cSite = findCol(headers, ['site']);
      var cLesson = findCol(headers, ['lesson', 'workshop']);
      var cStart = findCol(headers, ['start time']);
      var cEnd = findCol(headers, ['end time']);
      var cSetup = findCol(headers, ['setup']);

      if (cSite === -1 || cStart === -1 || cEnd === -1) continue;

      var sheetCount = 0;
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var site = String(row[cSite] || '').trim();
        var lesson = cLesson !== -1 ? String(row[cLesson] || '').trim() : '';
        var startT = row[cStart];
        var endT = row[cEnd];

        if (cSetup !== -1) {
          var setup = String(row[cSetup] || '').toLowerCase();
          if (setup.indexOf('cancelled') >= 0 || setup.indexOf('cancel') >= 0) continue;
        }
        if (!site || !startT || !endT) continue;

        var dur = getDurationMin(startT, endT);
        if (dur <= 0) continue;

        var key = normalize(site + '|' + lesson);
        opsHub[key] = { site: site, lesson: lesson, dur: dur, allowed: dur + BUFFER_MINUTES };
        sheetCount++;
      }
      if (sheetCount > 0) log.push('  Tab "' + sheet.getName() + '": ' + sheetCount + ' workshops');
    }
    log.push('Total Ops Hub workshops: ' + Object.keys(opsHub).length);
  } catch (err) {
    log.push('ERROR loading Ops Hub: ' + err.message);
  }

  // STEP 2: Load ALL check-in tabs
  var allCheckIns = [];
  try {
    var ciSS = SpreadsheetApp.openById(CHECKINS_ID);
    var ciSheets = ciSS.getSheets();
    log.push('\nCheck-ins: ' + ciSS.getName() + ' (' + ciSheets.length + ' tabs)');

    for (var s = 0; s < ciSheets.length; s++) {
      var sheet = ciSheets[s];
      var sheetName = sheet.getName();
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) { continue; }

      // Find header row — scan first 20 rows
      var hdrRow = -1;
      var hdr = [];
      for (var h = 0; h < Math.min(20, data.length); h++) {
        var rowStr = data[h].map(function(c) { return String(c).toLowerCase().trim(); });
        for (var cc = 0; cc < rowStr.length; cc++) {
          if (rowStr[cc].indexOf('leader name') >= 0) {
            hdrRow = h;
            hdr = rowStr;
            break;
          }
        }
        if (hdrRow >= 0) break;
      }
      if (hdrRow === -1) {
        log.push('  Tab "' + sheetName + '": SKIPPED (no "leader name" header found)');
        continue;
      }

      var cRegion = findCol(hdr, ['region']);
      var cWorkshop = findCol(hdr, ['workshop']);
      var cSchool = findCol(hdr, ['school']);
      var cLeader = findCol(hdr, ['leader name']);
      var cDate = findCol(hdr, ['date']);
      var cStatus = findCol(hdr, ['status']);

      if (cLeader === -1 || cStatus === -1) {
        log.push('  Tab "' + sheetName + '": SKIPPED (missing leader/status columns)');
        continue;
      }

      var tabCount = 0;
      var tabSkipped = 0;
      for (var i = hdrRow + 1; i < data.length; i++) {
        var row = data[i];
        var leader = getVal(row, cLeader);
        var status = getVal(row, cStatus).toLowerCase();
        var workshop = getVal(row, cWorkshop);
        var school = getVal(row, cSchool);
        var region = getVal(row, cRegion);

        if (!leader || !status) continue;
        if (!workshop && !school) continue;

        // Parse date
        var dt = null;
        if (cDate !== -1 && row[cDate]) {
          dt = parseDate(row[cDate]);
        }

        // Filter by payroll period
        if (dt) {
          if (dt < startDate || dt > endDate) {
            tabSkipped++;
            continue;
          }
        }

        var worked = false;
        for (var w = 0; w < WORKED_STATUSES.length; w++) {
          if (status.indexOf(WORKED_STATUSES[w]) >= 0) { worked = true; break; }
        }

        allCheckIns.push({
          leader: leader,
          status: status,
          worked: worked,
          workshop: workshop,
          school: school,
          region: region,
          date: dt,
          dateStr: dt ? formatDt(dt) : 'N/A'
        });
        tabCount++;
      }
      log.push('  Tab "' + sheetName + '": ' + tabCount + ' records loaded, ' + tabSkipped + ' outside date range');
    }
    log.push('Total check-in records: ' + allCheckIns.length);
  } catch (err) {
    log.push('ERROR loading check-ins: ' + err.message);
  }

  // STEP 2b: Deduplicate leader names
  allCheckIns = deduplicateLeaderNames_(allCheckIns);

  // STEP 3: Calculate hours per leader
  var leaders = {};
  var matchedOps = 0;
  var unmatchedOps = 0;

  for (var i = 0; i < allCheckIns.length; i++) {
    var rec = allCheckIns[i];
    var name = rec.leader;

    if (!leaders[name]) {
      leaders[name] = { name: name, totalMin: 0, worked: 0, absent: 0, unmatched: 0, details: [] };
    }

    if (rec.worked) {
      var ops = matchOpsHub(rec.school, rec.workshop, opsHub);
      var dur, allowed, src;

      if (ops) {
        dur = ops.dur;
        allowed = ops.allowed;
        src = 'Ops Hub (' + ops.site + ')';
        matchedOps++;
      } else {
        dur = 60;
        allowed = 90;
        src = 'DEFAULT 1hr';
        unmatchedOps++;
        leaders[name].unmatched++;
      }

      leaders[name].totalMin += allowed;
      leaders[name].worked++;
      leaders[name].details.push({ date: rec.dateStr, ws: rec.workshop, school: rec.school, status: rec.status, dur: dur, allowed: allowed, src: src });
    } else {
      leaders[name].absent++;
      leaders[name].details.push({ date: rec.dateStr, ws: rec.workshop, school: rec.school, status: rec.status, dur: 0, allowed: 0, src: 'Not worked' });
    }
  }

  var leaderCount = Object.keys(leaders).length;
  var workedLeaders = 0;
  for (var k in leaders) {
    if (leaders[k].worked > 0) workedLeaders++;
  }

  log.push('\nLeaders total: ' + leaderCount);
  log.push('Leaders who worked: ' + workedLeaders);
  log.push('Ops Hub matches: ' + matchedOps + ' / Unmatched: ' + unmatchedOps);

  // STEP 4: Write the report into Hours Verification sheet
  try {
    var hvSS = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
    log.push('\nWriting to: ' + hvSS.getName());

    // --- Write "Auto Verification" tab ---
    var rpt = hvSS.getSheetByName('Auto Verification');
    if (!rpt) {
      rpt = hvSS.insertSheet('Auto Verification');
    } else {
      rpt.clear();
      rpt.clearFormats();
    }

    var dateRange = formatDt(startDate) + ' – ' + formatDt(endDate);

    // Title
    rpt.getRange(1, 1).setValue('HOURS VERIFICATION — ' + dateRange);
    rpt.getRange(1, 1, 1, 8).merge();
    rpt.getRange(1, 1).setFontSize(14).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

    rpt.getRange(2, 1).setValue('Generated: ' + new Date().toLocaleString() + ' | Rule: Class Duration + ' + BUFFER_MINUTES + ' min | Matched: ' + matchedOps + ' | Unmatched: ' + unmatchedOps);
    rpt.getRange(2, 1, 1, 8).merge().setFontColor('#666666');

    // Summary headers
    var sumHeaders = ['Leader Name', 'Sessions Worked', 'Allowed Hours', 'Formatted', 'Absent', 'Unmatched', 'Status', 'Session Details'];
    rpt.getRange(4, 1, 1, 8).setValues([sumHeaders]).setFontWeight('bold').setBackground('#e8eaed');

    // Sort leaders by total hours descending
    var sortedLeaders = [];
    for (var k in leaders) {
      if (leaders[k].worked > 0) sortedLeaders.push(leaders[k]);
    }
    sortedLeaders.sort(function(a, b) { return b.totalMin - a.totalMin; });

    // Write all leaders in batch
    var summaryData = [];
    for (var i = 0; i < sortedLeaders.length; i++) {
      var l = sortedLeaders[i];
      var hrs = Math.round((l.totalMin / 60) * 100) / 100;
      var hh = Math.floor(hrs);
      var mm = Math.round((hrs - hh) * 60);
      var fmt = hh + 'h ' + (mm < 10 ? '0' : '') + mm + 'm';

      var detailParts = [];
      for (var d = 0; d < l.details.length; d++) {
        var det = l.details[d];
        if (det.allowed > 0) {
          detailParts.push(det.date + ': ' + det.ws + ' @ ' + det.school + ' (' + det.status + ') → ' + det.allowed + 'min');
        }
      }

      summaryData.push([
        l.name,
        l.worked,
        hrs,
        fmt,
        l.absent,
        l.unmatched,
        l.unmatched > 0 ? '⚠️ Check' : '✅ OK',
        detailParts.join('\n')
      ]);
    }

    if (summaryData.length > 0) {
      rpt.getRange(5, 1, summaryData.length, 8).setValues(summaryData);
      rpt.getRange(5, 8, summaryData.length, 1).setWrap(true);
      log.push('Wrote ' + summaryData.length + ' leaders to Auto Verification summary');
    } else {
      log.push('WARNING: No leaders with worked sessions to write!');
    }

    // --- PAYROLL COMPARISON SECTION ---
    var pStart = 5 + summaryData.length + 2;
    rpt.getRange(pStart, 1).setValue('PAYROLL COMPARISON — Paste Gusto/ADP hours in Column C');
    rpt.getRange(pStart, 1, 1, 6).merge();
    rpt.getRange(pStart, 1).setFontSize(12).setFontWeight('bold').setBackground('#fbbc04');

    var pHeaders = ['Leader Name', 'Max Allowed Hours', 'Reported Hours (paste)', 'Difference', 'FLAG', 'Notes'];
    rpt.getRange(pStart + 1, 1, 1, 6).setValues([pHeaders]).setFontWeight('bold').setBackground('#e8eaed');

    var payrollData = [];
    for (var i = 0; i < sortedLeaders.length; i++) {
      var l = sortedLeaders[i];
      var hrs = Math.round((l.totalMin / 60) * 100) / 100;
      payrollData.push([l.name, hrs, '', '', '', '']);
    }

    if (payrollData.length > 0) {
      var pDataStart = pStart + 2;
      rpt.getRange(pDataStart, 1, payrollData.length, 6).setValues(payrollData);

      // Add formulas for difference and flag
      for (var i = 0; i < payrollData.length; i++) {
        var r = pDataStart + i;
        rpt.getRange(r, 4).setFormula('=IF(C' + r + '="","",C' + r + '-B' + r + ')');
        rpt.getRange(r, 5).setFormula('=IF(C' + r + '="","",IF(C' + r + '>B' + r + ',"⚠️ OVER","✅ OK"))');
      }
      log.push('Wrote ' + payrollData.length + ' leaders to payroll comparison');
    }

    // --- SESSION LOG ---
    var sStart = pStart + 2 + payrollData.length + 2;
    rpt.getRange(sStart, 1).setValue('FULL SESSION LOG');
    rpt.getRange(sStart, 1, 1, 8).merge();
    rpt.getRange(sStart, 1).setFontSize(12).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');

    var sHeaders = ['Leader', 'Date', 'Workshop', 'School', 'Status', 'Duration (min)', 'Allowed (min)', 'Source'];
    rpt.getRange(sStart + 1, 1, 1, 8).setValues([sHeaders]).setFontWeight('bold').setBackground('#e8eaed');

    var sessionData = [];
    for (var i = 0; i < sortedLeaders.length; i++) {
      var l = sortedLeaders[i];
      for (var d = 0; d < l.details.length; d++) {
        var det = l.details[d];
        sessionData.push([l.name, det.date, det.ws, det.school, det.status, det.dur, det.allowed, det.src]);
      }
    }

    if (sessionData.length > 0) {
      rpt.getRange(sStart + 2, 1, sessionData.length, 8).setValues(sessionData);
      log.push('Wrote ' + sessionData.length + ' session records to log');
    }

    // Auto-resize
    for (var c = 1; c <= 8; c++) rpt.autoResizeColumn(c);
    rpt.setFrozenRows(4);

    // --- Also update Van's Sheet ---
    writeToVansSheet(hvSS, leaders);

    log.push('\n=== COMPLETE ===');

  } catch (err) {
    log.push('ERROR writing report: ' + err.message + '\n' + err.stack);
  }

  // Show results
  var logStr = log.join('\n');
  Logger.log(logStr);

  try {
    SpreadsheetApp.getUi().alert('Report Complete!',
      'Leaders processed: ' + workedLeaders +
      '\nCheck-ins loaded: ' + allCheckIns.length +
      '\nOps Hub matches: ' + matchedOps +
      '\n\n→ Check "Auto Verification" tab for full report' +
      '\n→ Check Van\'s Sheet for new columns on the right',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {
    Logger.log('Could not show alert: ' + e.message);
  }
}

// =====================================================
// WRITE TO VAN'S SHEET
// =====================================================
function writeToVansSheet(hvSS, leaders) {
  // Find the right tab
  var vSheet = null;
  var sheets = hvSS.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName().toLowerCase();
    if (n.indexOf('van') >= 0 || n.indexOf('02_05') >= 0 || n.indexOf('02/05') >= 0) {
      vSheet = sheets[i];
      break;
    }
  }
  if (!vSheet) vSheet = sheets[0];

  var data = vSheet.getDataRange().getValues();
  var lastCol = data[0].length;
  var writeCol = lastCol + 1;

  // Write header
  vSheet.getRange(1, writeCol).setValue('Auto-Calc Hours');
  vSheet.getRange(1, writeCol + 1).setValue('Sessions');
  vSheet.getRange(1, writeCol + 2).setValue('Absent');
  vSheet.getRange(1, writeCol + 3).setValue('Status');
  vSheet.getRange(1, writeCol, 1, 4).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // Build normalized lookup
  var lookup = {};
  for (var k in leaders) {
    var normName = k.toLowerCase().replace(/[^a-z]/g, '');
    lookup[normName] = leaders[k];
  }

  var matched = 0;
  for (var i = 0; i < data.length; i++) {
    var cellVal = String(data[i][0] || '').trim();
    if (!cellVal || cellVal.length < 3) continue;

    var normCell = cellVal.toLowerCase().replace(/[^a-z]/g, '');
    var match = lookup[normCell];

    if (match && match.worked > 0) {
      var hrs = Math.round((match.totalMin / 60) * 100) / 100;
      var hh = Math.floor(hrs);
      var mm = Math.round((hrs - hh) * 60);
      var fmt = hh + 'h ' + (mm < 10 ? '0' : '') + mm + 'm';

      vSheet.getRange(i + 1, writeCol).setValue(fmt);
      vSheet.getRange(i + 1, writeCol + 1).setValue(match.worked);
      vSheet.getRange(i + 1, writeCol + 2).setValue(match.absent);
      vSheet.getRange(i + 1, writeCol + 3).setValue(match.unmatched > 0 ? '⚠️' : '✅');
      vSheet.getRange(i + 1, writeCol, 1, 4).setBackground(match.unmatched > 0 ? '#fff3cd' : '#e6f4ea');
      matched++;
    }
  }
  Logger.log('Van\'s Sheet: matched ' + matched + ' leaders');
}

// =====================================================
// PHASE 2: GUSTO CSV IMPORT + AUTO-COMPARE
// =====================================================

/**
 * Creates the "Gusto Import" tab with instructions and column headers.
 * HR pastes the Gusto CSV export data here.
 */
function setupGustoImportTab() {
  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var tab = ss.getSheetByName(GUSTO_IMPORT_TAB);
  if (!tab) {
    tab = ss.insertSheet(GUSTO_IMPORT_TAB);
  } else {
    tab.clear();
    tab.clearFormats();
  }

  // Instructions row
  tab.getRange(1, 1).setValue('GUSTO HOURS IMPORT');
  tab.getRange(1, 1, 1, 6).merge();
  tab.getRange(1, 1).setFontSize(14).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  tab.getRange(2, 1).setValue('Instructions: Export hours from Gusto (People > Time Tools > Export), then paste data starting at row 4. Columns: Employee Name, Total Hours. Other columns are optional and will be ignored.');
  tab.getRange(2, 1, 1, 6).merge().setFontColor('#666666').setWrap(true);

  // Expected headers
  var headers = ['Employee Name', 'Total Hours', 'Regular Hours', 'Overtime Hours', 'Department', 'Notes'];
  tab.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');

  // Sample row for reference
  tab.getRange(5, 1, 1, 6).setValues([['Jane Doe', 12.5, 12.5, 0, 'Leaders', '(sample — delete this row)']]);
  tab.getRange(5, 1, 1, 6).setFontColor('#999999').setFontStyle('italic');

  tab.setFrozenRows(4);
  for (var c = 1; c <= 6; c++) tab.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert('Gusto Import tab created!\n\nPaste your Gusto CSV export starting at row 5.\nMake sure Column A = Employee Name and Column B = Total Hours.');
}

/**
 * Reads the Gusto Import tab, matches against the Auto Verification data,
 * and writes a Discrepancies report.
 */
function compareGustoHours() {
  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);

  // Step 1: Read Gusto Import tab
  var gustoTab = ss.getSheetByName(GUSTO_IMPORT_TAB);
  if (!gustoTab) {
    SpreadsheetApp.getUi().alert('No "' + GUSTO_IMPORT_TAB + '" tab found.\n\nRun "Setup Gusto Import Tab" first, then paste your Gusto data.');
    return;
  }

  var gustoData = gustoTab.getDataRange().getValues();

  // Find header row (look for "employee" or "name" in first 10 rows)
  var gHeaderRow = -1;
  var gNameCol = -1;
  var gHoursCol = -1;
  for (var r = 0; r < Math.min(10, gustoData.length); r++) {
    for (var c = 0; c < gustoData[r].length; c++) {
      var cell = String(gustoData[r][c]).toLowerCase().trim();
      if (cell.indexOf('employee') >= 0 || cell.indexOf('name') >= 0) {
        if (gNameCol === -1) { gNameCol = c; gHeaderRow = r; }
      }
      if (cell.indexOf('total hours') >= 0 || cell.indexOf('total') >= 0 || cell.indexOf('hours') >= 0) {
        if (gHoursCol === -1) gHoursCol = c;
      }
    }
    if (gNameCol >= 0 && gHoursCol >= 0) break;
  }

  // Fallback: assume col A = name, col B = hours, header at row 4 (0-indexed row 3)
  if (gNameCol === -1) gNameCol = 0;
  if (gHoursCol === -1) gHoursCol = 1;
  if (gHeaderRow === -1) gHeaderRow = 3;

  // Parse Gusto entries
  var gustoEntries = [];
  for (var i = gHeaderRow + 1; i < gustoData.length; i++) {
    var name = String(gustoData[i][gNameCol] || '').trim();
    var hours = parseFloat(gustoData[i][gHoursCol]);
    if (!name || name.length < 2 || isNaN(hours)) continue;
    gustoEntries.push({ name: name, hours: hours, row: i + 1 });
  }

  if (gustoEntries.length === 0) {
    SpreadsheetApp.getUi().alert('No valid data found in Gusto Import tab.\n\nMake sure Column A has names and Column B has hours.');
    return;
  }

  // Step 2: Read Auto Verification tab to get allowed hours
  var autoTab = ss.getSheetByName('Auto Verification');
  if (!autoTab) {
    SpreadsheetApp.getUi().alert('No "Auto Verification" tab found.\n\nRun the main report first to generate allowed hours.');
    return;
  }

  var autoData = autoTab.getDataRange().getValues();

  // Build lookup: normalized name -> { name, allowedHours, sessions, details }
  var allowedLookup = {};
  for (var i = 4; i < autoData.length; i++) {
    var leaderName = String(autoData[i][0] || '').trim();
    var allowedHrs = parseFloat(autoData[i][2]);
    var sessions = parseInt(autoData[i][1]) || 0;
    var details = String(autoData[i][7] || '');

    if (!leaderName || leaderName.length < 2 || isNaN(allowedHrs)) continue;

    // Stop if we hit the payroll comparison section
    if (leaderName.toUpperCase().indexOf('PAYROLL') >= 0) break;
    if (leaderName.toUpperCase().indexOf('SESSION LOG') >= 0) break;

    var normKey = leaderName.toLowerCase().replace(/[^a-z]/g, '');
    allowedLookup[normKey] = {
      name: leaderName,
      allowedHours: allowedHrs,
      sessions: sessions,
      details: details
    };
  }

  // Step 3: Match and compare
  var results = [];
  var flagged = [];
  var matched = 0;
  var unmatched = [];

  for (var g = 0; g < gustoEntries.length; g++) {
    var entry = gustoEntries[g];
    var match = fuzzyMatchName(entry.name, allowedLookup);

    if (match) {
      matched++;
      var diff = Math.round((entry.hours - match.allowedHours) * 100) / 100;
      var status = diff > 0 ? 'OVER' : 'OK';

      var result = {
        gustoName: entry.name,
        matchedName: match.name,
        reportedHours: entry.hours,
        allowedHours: match.allowedHours,
        difference: diff,
        status: status,
        sessions: match.sessions,
        details: match.details
      };
      results.push(result);

      if (diff > 0) {
        flagged.push(result);
      }
    } else {
      unmatched.push(entry.name);
      results.push({
        gustoName: entry.name,
        matchedName: '(no match)',
        reportedHours: entry.hours,
        allowedHours: 'N/A',
        difference: 'N/A',
        status: 'UNMATCHED',
        sessions: '',
        details: ''
      });
    }
  }

  // Sort flagged by overage descending
  flagged.sort(function(a, b) { return b.difference - a.difference; });

  // Step 4: Write Discrepancies tab
  writeDiscrepanciesReport(ss, results, flagged, unmatched, matched, gustoEntries.length);

  // Step 5: Show summary
  var msg = 'Gusto Comparison Complete!\n\n' +
    'Gusto entries: ' + gustoEntries.length + '\n' +
    'Matched: ' + matched + '\n' +
    'Unmatched: ' + unmatched.length + '\n' +
    'Flagged (over allowed): ' + flagged.length + '\n\n' +
    '→ Check "' + DISCREPANCIES_TAB + '" tab for full report';

  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Fuzzy-matches a Gusto name against the allowed hours lookup.
 * Handles: different order (Last, First vs First Last), extra spaces, etc.
 */
function fuzzyMatchName(gustoName, lookup) {
  // Exact normalized match
  var normGusto = gustoName.toLowerCase().replace(/[^a-z]/g, '');
  if (lookup[normGusto]) return lookup[normGusto];

  // Try reversing "Last, First" → "FirstLast"
  if (gustoName.indexOf(',') >= 0) {
    var parts = gustoName.split(',');
    var reversed = (parts[1] || '').trim() + ' ' + (parts[0] || '').trim();
    var normReversed = reversed.toLowerCase().replace(/[^a-z]/g, '');
    if (lookup[normReversed]) return lookup[normReversed];
  }

  // Try reversing "First Last" → "LastFirst" (in case lookup has Last, First)
  var spaceParts = gustoName.trim().split(/\s+/);
  if (spaceParts.length >= 2) {
    var lastFirst = spaceParts[spaceParts.length - 1] + spaceParts.slice(0, -1).join('');
    var normLF = lastFirst.toLowerCase().replace(/[^a-z]/g, '');
    if (lookup[normLF]) return lookup[normLF];
  }

  // Partial match: find the best match by substring overlap
  var bestKey = null;
  var bestScore = 0;

  for (var key in lookup) {
    var score = 0;

    // Check if one contains the other
    if (key.indexOf(normGusto) >= 0 || normGusto.indexOf(key) >= 0) {
      score = Math.min(key.length, normGusto.length) / Math.max(key.length, normGusto.length);
      score += 0.5; // bonus for substring match
    } else {
      // Check first name + last name overlap
      var gustoWords = gustoName.toLowerCase().replace(/[^a-z\s]/g, '').split(/\s+/);
      var lookupWords = lookup[key].name.toLowerCase().replace(/[^a-z\s]/g, '').split(/\s+/);

      var wordMatches = 0;
      for (var gw = 0; gw < gustoWords.length; gw++) {
        for (var lw = 0; lw < lookupWords.length; lw++) {
          if (gustoWords[gw] === lookupWords[lw] && gustoWords[gw].length >= 2) {
            wordMatches++;
          }
        }
      }
      if (wordMatches >= 2) score = 0.8 + (wordMatches * 0.05);
      else if (wordMatches === 1 && gustoWords.length <= 2 && lookupWords.length <= 2) score = 0.5;
    }

    if (score > bestScore) {
      bestScore = score;
      bestKey = key;
    }
  }

  // Require a minimum confidence
  if (bestScore >= 0.5 && bestKey) return lookup[bestKey];
  return null;
}

/**
 * Writes the Discrepancies report tab.
 */
function writeDiscrepanciesReport(ss, results, flagged, unmatched, matchedCount, totalGusto) {
  var tab = ss.getSheetByName(DISCREPANCIES_TAB);
  if (!tab) {
    tab = ss.insertSheet(DISCREPANCIES_TAB);
  } else {
    tab.clear();
    tab.clearFormats();
  }

  // Title
  tab.getRange(1, 1).setValue('GUSTO vs ALLOWED HOURS — Discrepancies Report');
  tab.getRange(1, 1, 1, 8).merge();
  tab.getRange(1, 1).setFontSize(14).setFontWeight('bold').setBackground('#d93025').setFontColor('#ffffff');

  var now = new Date().toLocaleString();
  tab.getRange(2, 1).setValue('Generated: ' + now + ' | Gusto entries: ' + totalGusto + ' | Matched: ' + matchedCount + ' | Unmatched: ' + unmatched.length + ' | Flagged: ' + flagged.length);
  tab.getRange(2, 1, 1, 8).merge().setFontColor('#666666');

  // --- FLAGGED LEADERS (OVER ALLOWED) ---
  tab.getRange(4, 1).setValue('FLAGGED LEADERS — Over Allowed Hours (' + flagged.length + ')');
  tab.getRange(4, 1, 1, 8).merge();
  tab.getRange(4, 1).setFontSize(12).setFontWeight('bold').setBackground('#fce8e6');

  var flagHeaders = ['Leader Name (Gusto)', 'Matched Name', 'Reported Hours', 'Max Allowed Hours', 'Overage', 'Sessions', 'Workshop Details', 'Action Needed'];
  tab.getRange(5, 1, 1, flagHeaders.length).setValues([flagHeaders]).setFontWeight('bold').setBackground('#e8eaed');

  if (flagged.length > 0) {
    var flagData = [];
    for (var i = 0; i < flagged.length; i++) {
      var f = flagged[i];
      flagData.push([
        f.gustoName,
        f.matchedName,
        f.reportedHours,
        f.allowedHours,
        '+' + f.difference + 'h',
        f.sessions,
        f.details,
        'Review — reduce to ' + f.allowedHours + 'h'
      ]);
    }
    tab.getRange(6, 1, flagData.length, 8).setValues(flagData);
    // Highlight overage column
    tab.getRange(6, 5, flagData.length, 1).setBackground('#fce8e6').setFontWeight('bold').setFontColor('#d93025');
    tab.getRange(6, 8, flagData.length, 1).setBackground('#fff3cd');
  }

  // --- ALL COMPARISONS ---
  var allStart = 6 + Math.max(flagged.length, 0) + 2;
  tab.getRange(allStart, 1).setValue('ALL COMPARISONS (' + results.length + ')');
  tab.getRange(allStart, 1, 1, 8).merge();
  tab.getRange(allStart, 1).setFontSize(12).setFontWeight('bold').setBackground('#e8f0fe');

  var allHeaders = ['Leader Name (Gusto)', 'Matched Name', 'Reported Hours', 'Max Allowed Hours', 'Difference', 'Status', 'Sessions', 'Workshop Details'];
  tab.getRange(allStart + 1, 1, 1, allHeaders.length).setValues([allHeaders]).setFontWeight('bold').setBackground('#e8eaed');

  if (results.length > 0) {
    var allData = [];
    for (var i = 0; i < results.length; i++) {
      var r = results[i];
      var diffStr = (typeof r.difference === 'number') ? (r.difference > 0 ? '+' + r.difference : String(r.difference)) : r.difference;
      var statusIcon = r.status === 'OVER' ? '⚠️ OVER' : (r.status === 'UNMATCHED' ? '❓ UNMATCHED' : '✅ OK');
      allData.push([
        r.gustoName,
        r.matchedName,
        r.reportedHours,
        r.allowedHours,
        diffStr,
        statusIcon,
        r.sessions,
        r.details
      ]);
    }
    var allDataStart = allStart + 2;
    tab.getRange(allDataStart, 1, allData.length, 8).setValues(allData);

    // Conditional formatting: highlight over rows
    for (var i = 0; i < results.length; i++) {
      if (results[i].status === 'OVER') {
        tab.getRange(allDataStart + i, 1, 1, 8).setBackground('#fce8e6');
      } else if (results[i].status === 'UNMATCHED') {
        tab.getRange(allDataStart + i, 1, 1, 8).setBackground('#fff3cd');
      }
    }

    tab.getRange(allDataStart, 8, allData.length, 1).setWrap(true);
  }

  // --- UNMATCHED NAMES ---
  if (unmatched.length > 0) {
    var uStart = allStart + 2 + results.length + 2;
    tab.getRange(uStart, 1).setValue('UNMATCHED GUSTO NAMES (' + unmatched.length + ')');
    tab.getRange(uStart, 1, 1, 4).merge();
    tab.getRange(uStart, 1).setFontSize(11).setFontWeight('bold').setBackground('#fff3cd');

    tab.getRange(uStart + 1, 1).setValue('Gusto Name');
    tab.getRange(uStart + 1, 2).setValue('Possible Issue');
    tab.getRange(uStart + 1, 1, 1, 2).setFontWeight('bold').setBackground('#e8eaed');

    var uData = [];
    for (var i = 0; i < unmatched.length; i++) {
      uData.push([unmatched[i], 'Not found in check-ins — may be admin, non-leader, or name mismatch']);
    }
    tab.getRange(uStart + 2, 1, uData.length, 2).setValues(uData);
  }

  // Auto-resize
  for (var c = 1; c <= 8; c++) tab.autoResizeColumn(c);
  tab.setFrozenRows(5);
}

// =====================================================
// HOURS TRACKER — Auto-generated region-grouped tabs
// =====================================================

/**
 * Menu entry: generate Hours Tracker for the fixed 02/05–02/18 pay period.
 */
function generateHoursTracker02_05to02_18() {
  var startDate = new Date(2026, 1, 5);
  var endDate = new Date(2026, 1, 18, 23, 59, 59);
  generateHoursTracker(startDate, endDate);
}

// ---- AUTO-REFRESH TRIGGER ----

/**
 * Enable hourly auto-refresh for the current pay period (02/05–02/18).
 * Clears any existing auto-refresh triggers first to avoid duplicates.
 */
function enableAutoRefresh() {
  disableAutoRefresh(); // remove existing triggers first
  ScriptApp.newTrigger('autoRefreshHoursTracker')
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert(
    'Auto-refresh enabled!\n\n' +
    'The Hours Tracker will regenerate every hour with the latest check-in and absence data.\n\n' +
    'To stop, use Menu → Hours Verification → Disable Auto-Refresh.'
  );
  Logger.log('Auto-refresh trigger created (every 1 hour).');
}

/**
 * Disable auto-refresh by removing all autoRefreshHoursTracker triggers.
 */
function disableAutoRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoRefreshHoursTracker') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  if (removed > 0) {
    Logger.log('Removed ' + removed + ' auto-refresh trigger(s).');
    try {
      SpreadsheetApp.getUi().alert('Auto-refresh disabled. The Hours Tracker will no longer update automatically.');
    } catch (e) {
      // No UI context (called from enableAutoRefresh cleanup) — that's fine
    }
  }
}

/**
 * Called by the time-based trigger. Regenerates the Hours Tracker for 02/05–02/18.
 * Runs without UI context so no alerts/prompts — just regenerates silently.
 */
function autoRefreshHoursTracker() {
  var startDate = new Date(2026, 1, 5);
  var endDate = new Date(2026, 1, 18, 23, 59, 59);
  generateHoursTracker(startDate, endDate);
  Logger.log('Auto-refresh completed at ' + new Date().toLocaleString());
}

// ---- HOW IT WORKS GUIDE TAB ----

function createHowItWorksTab() {
  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var tabName = 'How It Works';
  var tab = ss.getSheetByName(tabName);
  if (tab) tab.clear(); else tab = ss.insertSheet(tabName);

  var rows = [];
  var styles = []; // [row, col, fontSize, bold, bg, fontColor]

  // Helper to add a section header
  function addHeader(text) {
    rows.push([text, '', '', '']);
    styles.push({ r: rows.length, size: 14, bold: true, bg: '#1a73e8', fg: '#ffffff' });
    rows.push(['', '', '', '']);
  }

  // Helper to add a sub-header
  function addSubHeader(text) {
    rows.push([text, '', '', '']);
    styles.push({ r: rows.length, size: 11, bold: true, bg: '#e8f0fe', fg: '#1a73e8' });
  }

  // Helper to add a row
  function addRow(col1, col2, col3, col4) {
    rows.push([col1 || '', col2 || '', col3 || '', col4 || '']);
  }

  // Helper to add an empty row
  function addBlank() { rows.push(['', '', '', '']); }

  // ======= TITLE =======
  rows.push(['KODELY HOURS VERIFICATION — HOW IT WORKS', '', '', '']);
  styles.push({ r: 1, size: 18, bold: true, bg: '#1a73e8', fg: '#ffffff' });
  rows.push(['This guide explains how the Hours Verification system works, what each tab shows, and how to use it.', '', '', '']);
  styles.push({ r: 2, size: 10, bold: false, bg: '#e8f0fe', fg: '#444444' });
  addBlank();

  // ======= OVERVIEW =======
  addHeader('OVERVIEW');
  addRow('This tool reads from two data sources and generates payroll-ready reports:');
  addBlank();
  addRow('  1.', 'Check-Ins Sheet', 'Leaders log when they arrive/leave workshops. Each row = one session.');
  addRow('  2.', 'Ops Hub Sheet', 'The master schedule of workshops with expected durations and school info.');
  addBlank();
  addRow('The system cross-references check-ins against the Ops Hub to calculate allowed hours,');
  addRow('flags absences, detects SCOOT contractors, deduplicates names, and produces clean output.');
  addBlank();

  // ======= DATA SOURCES =======
  addHeader('DATA SOURCES');
  addSubHeader('Check-Ins Sheet (live data)');
  addRow('  What:', 'Real-time log of leader check-ins from the field');
  addRow('  Key columns:', 'Leader Name, Date, Workshop, School, Status, Check-in Time, Check-out Time');
  addRow('  Status values:', '"leader", "co-lead", "sub", "coordinator" → Leader tab  |  "scoot" → SCOOT tab');
  addRow('  Absences:', 'If a leader is scheduled but has no check-in (or status = absent), they show as 0hr');
  addBlank();
  addSubHeader('Ops Hub Sheet (schedule)');
  addRow('  What:', 'Master workshop schedule with expected durations');
  addRow('  Used for:', 'Matching check-ins to workshops and calculating "allowed" hours');
  addRow('  Matching:', 'School name + workshop name fuzzy-matched to find the right schedule entry');
  addBlank();

  // ======= TABS EXPLAINED =======
  addHeader('TABS EXPLAINED');
  addSubHeader('Hours Tracker [date range]');
  addRow('  The main payroll verification tab. Contains 4 sections:');
  addBlank();
  addRow('  Section 1:', 'EXPECTED PAYROLL HOURS', 'One row per leader. Shows total sessions, absences, total hours.');
  addRow('', '', 'Sorted alphabetically. Grand total at bottom.');
  addRow('  Section 2:', 'SCOOT HOURS', 'Same format but only SCOOT contractors (status = "scoot").');
  addRow('', '', 'Billed at flat 3hr per session.');
  addRow('  Section 3:', 'DETAILED LEADER BREAKDOWN', 'One row per leader with ALL their workshops listed in a single cell.');
  addRow('', '', 'Shows dates, schools, durations — everything in one view.');
  addRow('  Section 4:', 'ALL UNIQUE WORKSHOPS', 'Every distinct workshop name that appeared in the date range.');
  addBlank();

  addSubHeader('SCOOT Hours [date range]');
  addRow('  Dedicated SCOOT contractor tab showing each session with school, date, and flat 3hr billing.');
  addRow('  Used to cross-reference against SCOOT invoices.');
  addBlank();

  addSubHeader('SCOOT Invoice Verify');
  addRow('  Cross-references SCOOT invoices against check-in data.');
  addRow('  Green = matched (invoice matches a check-in).  Red = unmatched (invoice with no matching check-in).');
  addRow('  Also shows "Unbilled Sessions" — check-ins with no corresponding invoice.');
  addBlank();

  addSubHeader('SCOOT Invoice Paste');
  addRow('  Paste raw SCOOT invoice data here. Columns: Instructor, School, Date, Hours, Amount.');
  addRow('  Then run "Verify SCOOT Invoices" from the menu.');
  addBlank();

  addSubHeader('Auto Verification / Discrepancies');
  addRow('  Legacy tabs from the original verification system. Auto Verification shows raw matched data.');
  addRow('  Discrepancies shows mismatches between check-ins and Ops Hub schedule.');
  addBlank();

  // ======= HOW TO USE =======
  addHeader('HOW TO USE');
  addSubHeader('Step 1: Generate the Hours Tracker');
  addRow('  Menu → Hours Verification → Generate Hours Tracker 02/05 - 02/18');
  addRow('  Or use "Custom Range" to pick any date range.');
  addRow('  This reads the latest check-ins and absences and builds fresh output tabs.');
  addBlank();

  addSubHeader('Step 2: Review Expected Payroll Hours');
  addRow('  Open the "Hours Tracker" tab. Section 1 shows every leader and their total hours.');
  addRow('  Compare these against what you see in Gusto or your payroll system.');
  addRow('  Look for: unusually high hours, missing leaders, unexpected 0hr entries.');
  addBlank();

  addSubHeader('Step 3: Check SCOOT Billing');
  addRow('  Open the "SCOOT Hours" tab to see all SCOOT sessions.');
  addRow('  If you have invoices from SCOOT, paste them into the "SCOOT Invoice Paste" tab,');
  addRow('  then run "Verify SCOOT Invoices" to catch any overbilling or missing sessions.');
  addBlank();

  addSubHeader('Step 4: Enable Auto-Refresh (optional)');
  addRow('  Menu → Hours Verification → Enable Auto-Refresh (every hour)');
  addRow('  The tracker will regenerate hourly with the latest data.');
  addRow('  Disable when the pay period closes and you\'ve finalized hours.');
  addBlank();

  // ======= KEY FEATURES =======
  addHeader('KEY FEATURES');
  addRow('  Name Deduplication', 'Fuzzy matching merges names like "Kelly O." and "Kelly Orzuna" into one entry.');
  addRow('  SCOOT Detection', 'Based on status column in check-ins. Status = "scoot" → SCOOT tab. Fix in check-ins if wrong.');
  addRow('  Absence Tracking', 'Scheduled sessions with no check-in show as 0hr with red highlighting.');
  addRow('  Formatted Hours', 'Durations shown as both minutes and hours+minutes (e.g., 90 → "1hr 30mins").');
  addRow('  Auto-Refresh', 'Hourly trigger keeps the tracker current as new check-ins come in.');
  addRow('  Batch Processing', 'Handles 300+ leaders efficiently without hitting Google Apps Script timeouts.');
  addBlank();

  // ======= TROUBLESHOOTING =======
  addHeader('TROUBLESHOOTING');
  addRow('Problem', 'Cause', 'Fix');
  styles.push({ r: rows.length, size: 10, bold: true, bg: '#f1f3f4', fg: '#333333' });
  addRow('Leader in wrong tab (SCOOT vs Leader)', 'Status column in check-ins is incorrect', 'Fix the status in the Check-Ins sheet, then re-run');
  addRow('Same person appears twice', 'Name spelled differently in check-ins', 'Fix spelling in check-ins, or the fuzzy matcher will merge close matches');
  addRow('Hours seem wrong', 'Check-in/check-out times are off', 'Verify the actual check-in times in the source sheet');
  addRow('SCOOT invoice doesn\'t match', 'School name or date mismatch', 'Check that invoice school name closely matches check-in school name');
  addRow('"Script timed out" error', 'Too much data for one run', 'Try a shorter date range, or wait and re-run');
  addBlank();

  // ======= QUICK REFERENCE =======
  addHeader('QUICK REFERENCE — MENU OPTIONS');
  addRow('Menu Item', 'What It Does');
  styles.push({ r: rows.length, size: 10, bold: true, bg: '#f1f3f4', fg: '#333333' });
  addRow('Generate Hours Tracker 02/05 - 02/18', 'Builds Hours Tracker + SCOOT tabs for the current pay period');
  addRow('Generate Hours Tracker (Custom Range)', 'Same thing but you pick the dates');
  addRow('Setup SCOOT Invoice Tab', 'Creates the paste tab for SCOOT invoice data');
  addRow('Verify SCOOT Invoices', 'Cross-checks pasted invoices vs check-in data');
  addRow('Enable Auto-Refresh (every hour)', 'Turns on hourly automatic regeneration');
  addRow('Disable Auto-Refresh', 'Stops automatic regeneration');
  addRow('Create "How It Works" Guide', 'Creates this tab you\'re reading now');
  addBlank();
  addBlank();
  addRow('Last updated: ' + new Date().toLocaleString());

  // ======= WRITE ALL DATA =======
  tab.getRange(1, 1, rows.length, 4).setValues(rows);

  // ======= APPLY STYLES =======
  // Set default font
  tab.getRange(1, 1, rows.length, 4).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('top').setWrap(true);

  // Apply individual row styles
  for (var i = 0; i < styles.length; i++) {
    var s = styles[i];
    var range = tab.getRange(s.r, 1, 1, 4);
    if (s.size) range.setFontSize(s.size);
    if (s.bold) range.setFontWeight('bold');
    if (s.bg) range.setBackground(s.bg);
    if (s.fg) range.setFontColor(s.fg);
  }

  // Column widths
  tab.setColumnWidth(1, 320);
  tab.setColumnWidth(2, 280);
  tab.setColumnWidth(3, 320);
  tab.setColumnWidth(4, 200);

  // Freeze title row
  tab.setFrozenRows(1);

  // Move tab to first position
  ss.setActiveSheet(tab);
  ss.moveActiveSheet(1);

  SpreadsheetApp.getUi().alert('\"How It Works\" tab created! It\'s the first tab in the spreadsheet.');
  Logger.log('How It Works tab created with ' + rows.length + ' rows.');
}

// ---- DAILY DASHBOARD EMAIL ----

/**
 * Prompt for recipient email, then send the dashboard.
 * No auto-send — only fires when you manually click it.
 */
function sendDashboardPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Send Dashboard Email', 'Enter recipient email address:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var email = response.getResponseText().trim();
  if (!email || email.indexOf('@') === -1) {
    ui.alert('Invalid email address.');
    return;
  }
  sendDailyDashboard_(email);
  ui.alert('Dashboard sent to ' + email + '!');
}

/**
 * Build and send the hours verification dashboard email to the given address.
 */
function sendDailyDashboard_(recipientEmail) {
  var startDate = DASHBOARD_PAY_PERIOD_START;
  var endDate = DASHBOARD_PAY_PERIOD_END;
  var today = new Date();
  var dayOfPeriod = Math.ceil((today - startDate) / 86400000);
  var totalDays = Math.ceil((endDate - startDate) / 86400000);

  // --- Load check-ins (same as generateHoursTracker Steps 1-3) ---
  var opsHub = {};
  try {
    var opsSS = SpreadsheetApp.openById(OPS_HUB_ID);
    var opsSheets = opsSS.getSheets();
    for (var s = 0; s < opsSheets.length; s++) {
      var sheet = opsSheets[s];
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var headers = [];
      for (var h = 0; h < data[0].length; h++) headers.push(String(data[0][h]).toLowerCase().trim());
      var cSite = findCol(headers, ['site']);
      var cLesson = findCol(headers, ['lesson', 'workshop']);
      var cStart = findCol(headers, ['start time']);
      var cEnd = findCol(headers, ['end time']);
      var cSetup = findCol(headers, ['setup']);
      if (cSite === -1 || cStart === -1 || cEnd === -1) continue;
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var site = String(row[cSite] || '').trim();
        var lesson = cLesson !== -1 ? String(row[cLesson] || '').trim() : '';
        var startT = row[cStart], endT = row[cEnd];
        if (cSetup !== -1) {
          var setup = String(row[cSetup] || '').toLowerCase();
          if (setup.indexOf('cancelled') >= 0 || setup.indexOf('cancel') >= 0) continue;
        }
        if (!site || !startT || !endT) continue;
        var dur = getDurationMin(startT, endT);
        if (dur <= 0) continue;
        var key = normalize(site + '|' + lesson);
        opsHub[key] = { site: site, lesson: lesson, dur: dur, allowed: dur + BUFFER_MINUTES };
      }
    }
  } catch (err) { Logger.log('Dashboard: Ops Hub error: ' + err.message); }

  var allCheckIns = [];
  try {
    var ciSS = SpreadsheetApp.openById(CHECKINS_ID);
    var ciSheets = ciSS.getSheets();
    for (var s = 0; s < ciSheets.length; s++) {
      var sheet = ciSheets[s];
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var hdrRow = -1, hdr = [];
      for (var h = 0; h < Math.min(20, data.length); h++) {
        var rowStr = data[h].map(function(c) { return String(c).toLowerCase().trim(); });
        for (var cc = 0; cc < rowStr.length; cc++) {
          if (rowStr[cc].indexOf('leader name') >= 0) { hdrRow = h; hdr = rowStr; break; }
        }
        if (hdrRow >= 0) break;
      }
      if (hdrRow === -1) continue;
      var cRegion = findCol(hdr, ['region']);
      var cWorkshop = findCol(hdr, ['workshop']);
      var cSchool = findCol(hdr, ['school']);
      var cLeader = findCol(hdr, ['leader name']);
      var cDate = findCol(hdr, ['date']);
      var cStatus = findCol(hdr, ['status']);
      if (cLeader === -1 || cStatus === -1) continue;
      for (var i = hdrRow + 1; i < data.length; i++) {
        var row = data[i];
        var leader = getVal(row, cLeader);
        var status = getVal(row, cStatus).toLowerCase();
        var workshop = getVal(row, cWorkshop);
        var school = getVal(row, cSchool);
        var region = getVal(row, cRegion);
        if (!leader || !status) continue;
        if (!workshop && !school) continue;
        var dt = null;
        if (cDate !== -1 && row[cDate]) dt = parseDate(row[cDate]);
        if (dt) { if (dt < startDate || dt > endDate) continue; }
        var worked = false;
        for (var w = 0; w < WORKED_STATUSES.length; w++) {
          if (status.indexOf(WORKED_STATUSES[w]) >= 0) { worked = true; break; }
        }
        allCheckIns.push({
          leader: leader, status: status, worked: worked,
          workshop: workshop, school: school, region: region,
          date: dt, dateStr: dt ? formatDt(dt) : 'N/A'
        });
      }
    }
  } catch (err) { Logger.log('Dashboard: Check-in error: ' + err.message); }

  allCheckIns = deduplicateLeaderNames_(allCheckIns);

  // Split into leader vs scoot, calculate hours
  var leaderSessions = [], scootSessions = [], absences = [];
  var leaderSet = {}, scootSet = {}, regionSet = {}, schoolSet = {};
  var totalLeaderMin = 0, totalScootSessions = 0;
  var unmatchedCount = 0;
  var dailyCounts = {}; // dateStr → count

  for (var i = 0; i < allCheckIns.length; i++) {
    var rec = allCheckIns[i];
    var statusTrimmed = rec.status.trim().toLowerCase();
    var isScoot = (statusTrimmed === 'scoot');
    var dur, allowed, ops;

    if (rec.worked) {
      ops = matchOpsHub(rec.school, rec.workshop, opsHub);
      if (ops) { dur = ops.dur; allowed = ops.allowed; }
      else { dur = 60; allowed = 90; unmatchedCount++; }
    } else {
      dur = 0; allowed = 0;
      absences.push({ leader: rec.leader, date: rec.dateStr, school: rec.school, workshop: rec.workshop });
    }

    if (isScoot) {
      scootSessions.push(rec);
      scootSet[rec.leader] = true;
      totalScootSessions++;
    } else {
      leaderSessions.push(rec);
      leaderSet[rec.leader] = (leaderSet[rec.leader] || 0) + allowed;
      totalLeaderMin += allowed;
    }

    if (rec.region) regionSet[rec.region] = (regionSet[rec.region] || 0) + 1;
    if (rec.school) schoolSet[rec.school] = true;
    if (rec.dateStr && rec.dateStr !== 'N/A') {
      dailyCounts[rec.dateStr] = (dailyCounts[rec.dateStr] || 0) + 1;
    }
  }

  // --- Compute metrics ---
  var totalLeaders = Object.keys(leaderSet).length;
  var totalScoot = Object.keys(scootSet).length;
  var totalHrs = Math.round((totalLeaderMin / 60) * 100) / 100;
  var avgHrsPerLeader = totalLeaders > 0 ? Math.round((totalHrs / totalLeaders) * 100) / 100 : 0;
  var totalSchools = Object.keys(schoolSet).length;
  var totalAbsences = absences.length;

  // Top 5 leaders by hours
  var leaderArr = [];
  for (var name in leaderSet) leaderArr.push({ name: name, min: leaderSet[name] });
  leaderArr.sort(function(a, b) { return b.min - a.min; });
  var top5 = leaderArr.slice(0, 5);

  // Bottom 5 (lowest hours, excluding 0)
  var bottom5 = leaderArr.filter(function(l) { return l.min > 0; });
  bottom5.sort(function(a, b) { return a.min - b.min; });
  bottom5 = bottom5.slice(0, 5);

  // Sessions by region
  var regionArr = [];
  for (var reg in regionSet) regionArr.push({ name: reg, count: regionSet[reg] });
  regionArr.sort(function(a, b) { return b.count - a.count; });

  // Sessions by day
  var dailyArr = [];
  for (var d in dailyCounts) dailyArr.push({ date: d, count: dailyCounts[d] });
  dailyArr.sort(function(a, b) { return new Date(a.date) - new Date(b.date); });

  // Recent absences (last 3 days)
  var recentAbsences = absences.filter(function(a) {
    var d = new Date(a.date);
    return (today - d) < 3 * 86400000;
  });

  // --- Build HTML email ---
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/' + HOURS_VERIFICATION_ID;
  var dateRange = formatDt(startDate) + ' – ' + formatDt(endDate);

  var html = '';
  html += '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;">';

  // Header
  html += '<div style="background:#1a73e8;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0;">';
  html += '<h1 style="margin:0;font-size:22px;">Kodely Hours Verification Dashboard</h1>';
  html += '<p style="margin:4px 0 0;opacity:0.9;font-size:14px;">Pay Period: ' + dateRange + ' &nbsp;|&nbsp; Day ' + dayOfPeriod + ' of ' + totalDays + '</p>';
  html += '</div>';

  // Key Metrics bar
  html += '<div style="display:flex;background:#f8f9fa;border:1px solid #e0e0e0;border-top:none;">';
  html += metricBox_('Total Leaders', totalLeaders, '#1a73e8');
  html += metricBox_('Total Hours', formatMinutes_(totalLeaderMin), '#34a853');
  html += metricBox_('SCOOT Staff', totalScoot, '#ff6d01');
  html += metricBox_('Absences', totalAbsences, totalAbsences > 0 ? '#ea4335' : '#34a853');
  html += metricBox_('Schools', totalSchools, '#9334e6');
  html += '</div>';

  // Summary stats
  html += '<div style="background:#fff;border:1px solid #e0e0e0;border-top:none;padding:16px 24px;">';
  html += '<p style="margin:0 0 4px;font-size:13px;color:#666;">Avg hours/leader: <strong>' + formatMinutes_(avgHrsPerLeader * 60) + '</strong> &nbsp;|&nbsp; ';
  html += 'Leader sessions: <strong>' + leaderSessions.length + '</strong> &nbsp;|&nbsp; ';
  html += 'SCOOT sessions: <strong>' + totalScootSessions + '</strong> (3hr flat each = ' + formatMinutes_(totalScootSessions * 180) + ') &nbsp;|&nbsp; ';
  html += 'Unmatched to Ops Hub: <strong>' + unmatchedCount + '</strong></p>';
  html += '</div>';

  // Top 5 Leaders
  html += sectionHeader_('Top 5 Leaders by Hours');
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
  html += '<tr style="background:#e8f0fe;"><th style="' + thStyle_() + '">#</th><th style="' + thStyle_() + '">Leader</th><th style="' + thStyle_() + '">Hours</th><th style="' + thStyle_() + '">Sessions</th></tr>';
  for (var i = 0; i < top5.length; i++) {
    var bg = i % 2 === 0 ? '#fff' : '#f8f9fa';
    var sessionCount = 0;
    for (var j = 0; j < leaderSessions.length; j++) { if (leaderSessions[j].leader === top5[i].name && leaderSessions[j].worked) sessionCount++; }
    html += '<tr style="background:' + bg + ';">';
    html += '<td style="' + tdStyle_() + '">' + (i + 1) + '</td>';
    html += '<td style="' + tdStyle_() + '"><strong>' + top5[i].name + '</strong></td>';
    html += '<td style="' + tdStyle_() + '">' + formatMinutes_(top5[i].min) + '</td>';
    html += '<td style="' + tdStyle_() + '">' + sessionCount + '</td>';
    html += '</tr>';
  }
  html += '</table>';

  // Bottom 5 Leaders (review for underpayment)
  if (bottom5.length > 0) {
    html += sectionHeader_('Lowest Hours (verify no missing check-ins)');
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#fce8e6;"><th style="' + thStyle_() + '">#</th><th style="' + thStyle_() + '">Leader</th><th style="' + thStyle_() + '">Hours</th></tr>';
    for (var i = 0; i < bottom5.length; i++) {
      var bg = i % 2 === 0 ? '#fff' : '#f8f9fa';
      html += '<tr style="background:' + bg + ';"><td style="' + tdStyle_() + '">' + (i + 1) + '</td>';
      html += '<td style="' + tdStyle_() + '">' + bottom5[i].name + '</td>';
      html += '<td style="' + tdStyle_() + '">' + formatMinutes_(bottom5[i].min) + '</td></tr>';
    }
    html += '</table>';
  }

  // Sessions by Day
  if (dailyArr.length > 0) {
    html += sectionHeader_('Sessions by Day');
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#e8f0fe;"><th style="' + thStyle_() + '">Date</th><th style="' + thStyle_() + '">Sessions</th><th style="' + thStyle_() + '">Visual</th></tr>';
    var maxDay = 1;
    for (var i = 0; i < dailyArr.length; i++) { if (dailyArr[i].count > maxDay) maxDay = dailyArr[i].count; }
    for (var i = 0; i < dailyArr.length; i++) {
      var bg = i % 2 === 0 ? '#fff' : '#f8f9fa';
      var barWidth = Math.round((dailyArr[i].count / maxDay) * 200);
      html += '<tr style="background:' + bg + ';">';
      html += '<td style="' + tdStyle_() + '">' + dailyArr[i].date + '</td>';
      html += '<td style="' + tdStyle_() + '">' + dailyArr[i].count + '</td>';
      html += '<td style="' + tdStyle_() + '"><div style="background:#4285f4;height:14px;width:' + barWidth + 'px;border-radius:3px;"></div></td>';
      html += '</tr>';
    }
    html += '</table>';
  }

  // Sessions by Region
  if (regionArr.length > 0) {
    html += sectionHeader_('Sessions by Region');
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#e8f0fe;"><th style="' + thStyle_() + '">Region</th><th style="' + thStyle_() + '">Sessions</th></tr>';
    for (var i = 0; i < regionArr.length; i++) {
      var bg = i % 2 === 0 ? '#fff' : '#f8f9fa';
      html += '<tr style="background:' + bg + ';"><td style="' + tdStyle_() + '">' + regionArr[i].name + '</td>';
      html += '<td style="' + tdStyle_() + '">' + regionArr[i].count + '</td></tr>';
    }
    html += '</table>';
  }

  // Recent Absences
  if (recentAbsences.length > 0) {
    html += sectionHeader_('Recent Absences (last 3 days)');
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#fce8e6;"><th style="' + thStyle_() + '">Leader</th><th style="' + thStyle_() + '">Date</th><th style="' + thStyle_() + '">School</th><th style="' + thStyle_() + '">Workshop</th></tr>';
    for (var i = 0; i < recentAbsences.length; i++) {
      var a = recentAbsences[i];
      var bg = i % 2 === 0 ? '#fff' : '#f8f9fa';
      html += '<tr style="background:' + bg + ';">';
      html += '<td style="' + tdStyle_() + '">' + a.leader + '</td>';
      html += '<td style="' + tdStyle_() + '">' + a.date + '</td>';
      html += '<td style="' + tdStyle_() + '">' + (a.school || '') + '</td>';
      html += '<td style="' + tdStyle_() + '">' + (a.workshop || '') + '</td>';
      html += '</tr>';
    }
    html += '</table>';
  } else {
    html += sectionHeader_('Recent Absences (last 3 days)');
    html += '<p style="padding:8px 24px;font-size:13px;color:#34a853;">No absences in the last 3 days.</p>';
  }

  // SCOOT Summary
  if (totalScootSessions > 0) {
    html += sectionHeader_('SCOOT Contractor Summary');
    html += '<table style="width:100%;border-collapse:collapse;font-size:13px;">';
    html += '<tr style="background:#fff3e0;"><th style="' + thStyle_() + '">Contractor</th><th style="' + thStyle_() + '">Sessions</th><th style="' + thStyle_() + '">Expected Bill (3hr × sessions)</th></tr>';
    var scootByName = {};
    for (var i = 0; i < scootSessions.length; i++) {
      var n = scootSessions[i].leader;
      scootByName[n] = (scootByName[n] || 0) + 1;
    }
    var scootArr = [];
    for (var n in scootByName) scootArr.push({ name: n, count: scootByName[n] });
    scootArr.sort(function(a, b) { return b.count - a.count; });
    for (var i = 0; i < scootArr.length; i++) {
      var bg = i % 2 === 0 ? '#fff' : '#f8f9fa';
      html += '<tr style="background:' + bg + ';">';
      html += '<td style="' + tdStyle_() + '">' + scootArr[i].name + '</td>';
      html += '<td style="' + tdStyle_() + '">' + scootArr[i].count + '</td>';
      html += '<td style="' + tdStyle_() + '">' + formatMinutes_(scootArr[i].count * 180) + '</td>';
      html += '</tr>';
    }
    html += '</table>';
  }

  // Alerts / Flags
  var flags = [];
  if (unmatchedCount > 0) flags.push('&#9888;&#65039; ' + unmatchedCount + ' session(s) could not be matched to Ops Hub — using default 1hr.');
  if (totalAbsences > 5) flags.push('&#9888;&#65039; High absence count (' + totalAbsences + ') this pay period.');
  for (var i = 0; i < top5.length; i++) {
    if (top5[i].min > 2400) flags.push('&#9888;&#65039; ' + top5[i].name + ' has ' + formatMinutes_(top5[i].min) + ' — verify this is correct.');
  }

  if (flags.length > 0) {
    html += sectionHeader_('Flags & Alerts');
    html += '<div style="padding:12px 24px;background:#fff8e1;border:1px solid #e0e0e0;border-top:none;">';
    for (var i = 0; i < flags.length; i++) {
      html += '<p style="margin:4px 0;font-size:13px;color:#e65100;">' + flags[i] + '</p>';
    }
    html += '</div>';
  }

  // Footer
  html += '<div style="padding:16px 24px;background:#f8f9fa;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;font-size:12px;color:#888;">';
  html += '<p style="margin:0;">Generated ' + today.toLocaleString() + ' &nbsp;|&nbsp; <a href="' + sheetUrl + '" style="color:#1a73e8;">Open Hours Verification Sheet</a></p>';
  html += '<p style="margin:4px 0 0;">This is an automated daily report. To stop, go to the sheet menu → Hours Verification → Disable Daily Dashboard Email.</p>';
  html += '</div>';
  html += '</div>';

  // Send
  MailApp.sendEmail({
    to: recipientEmail,
    subject: 'Kodely Hours Dashboard — ' + formatDt(today) + ' (Pay Period ' + dateRange + ')',
    htmlBody: html
  });

  Logger.log('Dashboard sent to ' + recipientEmail + ' at ' + today.toLocaleString());
}

// --- Dashboard HTML helpers ---
function metricBox_(label, value, color) {
  return '<div style="flex:1;text-align:center;padding:14px 8px;border-right:1px solid #e0e0e0;">' +
    '<div style="font-size:22px;font-weight:bold;color:' + color + ';">' + value + '</div>' +
    '<div style="font-size:11px;color:#666;margin-top:2px;">' + label + '</div></div>';
}
function sectionHeader_(text) {
  return '<div style="background:#1a73e8;color:#fff;padding:8px 24px;font-size:14px;font-weight:bold;border:1px solid #e0e0e0;border-top:none;">' + text + '</div>';
}
function thStyle_() { return 'padding:8px 12px;text-align:left;border-bottom:2px solid #ddd;'; }
function tdStyle_() { return 'padding:6px 12px;border-bottom:1px solid #eee;'; }

/**
 * Menu entry: prompt user for a custom date range, then generate Hours Tracker.
 */
function promptHoursTrackerDateRange() {
  var ui = SpreadsheetApp.getUi();
  var s = ui.prompt('Hours Tracker — Start date (MM/DD/YYYY):');
  if (s.getSelectedButton() !== ui.Button.OK) return;
  var e = ui.prompt('Hours Tracker — End date (MM/DD/YYYY):');
  if (e.getSelectedButton() !== ui.Button.OK) return;
  var sd = new Date(s.getResponseText());
  var ed = new Date(e.getResponseText());
  if (isNaN(sd.getTime()) || isNaN(ed.getTime())) {
    ui.alert('Invalid date format. Please use MM/DD/YYYY.');
    return;
  }
  ed.setHours(23, 59, 59);
  generateHoursTracker(sd, ed);
}

/**
 * Diagnostic: scans all check-in sheets and reports every unique status
 * that contains "scoot", plus sample rows. Shows results in an alert.
 */
function debugScootDetection() {
  var startDate = new Date(2026, 1, 5);
  var endDate = new Date(2026, 1, 18, 23, 59, 59);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var scootRows = [];
  var allStatuses = {};

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var name = sheet.getName();
    if (name === HOURS_TRACKER_TAB || name === SCOOT_HOURS_TAB || name === 'Auto Verification') continue;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    var hdr = data[0];
    var cStatus = -1, cLeader = -1, cSchool = -1, cDate = -1;
    for (var h = 0; h < hdr.length; h++) {
      var col = String(hdr[h]).toLowerCase().trim();
      if (col.indexOf('status') >= 0 && cStatus === -1) cStatus = h;
      if ((col.indexOf('leader') >= 0 || col.indexOf('name') >= 0) && cLeader === -1) cLeader = h;
      if (col.indexOf('school') >= 0 && cSchool === -1) cSchool = h;
      if (col.indexOf('date') >= 0 && cDate === -1) cDate = h;
    }
    if (cStatus === -1 || cLeader === -1) continue;

    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][cStatus] || '').toLowerCase().trim();
      if (!status) continue;
      allStatuses[status] = (allStatuses[status] || 0) + 1;
      if (status.indexOf('scoot') >= 0) {
        var leader = String(data[i][cLeader] || '').trim();
        var school = cSchool >= 0 ? String(data[i][cSchool] || '').trim() : '?';
        var dt = cDate >= 0 ? String(data[i][cDate] || '') : '?';
        scootRows.push(name + ': ' + leader + ' | ' + school + ' | ' + status + ' | ' + dt);
      }
    }
  }

  var msg = '=== ALL UNIQUE STATUSES ===\n';
  var keys = Object.keys(allStatuses).sort();
  for (var k = 0; k < keys.length; k++) {
    msg += keys[k] + ' (' + allStatuses[keys[k]] + ')\n';
  }
  msg += '\n=== SCOOT ROWS FOUND: ' + scootRows.length + ' ===\n';
  for (var r = 0; r < Math.min(scootRows.length, 20); r++) {
    msg += scootRows[r] + '\n';
  }

  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert('SCOOT Debug', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {
    Logger.log('Could not show alert.');
  }
}

/**
 * Main Hours Tracker function.
 * Loads check-in + Ops Hub data, splits into leaders vs scoot,
 * and writes two tabs: "Hours Tracker" and "SCOOT Hours".
 */
function generateHoursTracker(startDate, endDate) {
  var log = [];
  log.push('=== HOURS TRACKER ===');
  log.push('Period: ' + startDate.toDateString() + ' to ' + endDate.toDateString());

  // STEP 1: Load Ops Hub (reuse same logic as runReport)
  var opsHub = {};
  try {
    var opsSS = SpreadsheetApp.openById(OPS_HUB_ID);
    var opsSheets = opsSS.getSheets();
    for (var s = 0; s < opsSheets.length; s++) {
      var sheet = opsSheets[s];
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var headers = [];
      for (var h = 0; h < data[0].length; h++) {
        headers.push(String(data[0][h]).toLowerCase().trim());
      }
      var cSite = findCol(headers, ['site']);
      var cLesson = findCol(headers, ['lesson', 'workshop']);
      var cStart = findCol(headers, ['start time']);
      var cEnd = findCol(headers, ['end time']);
      var cSetup = findCol(headers, ['setup']);
      if (cSite === -1 || cStart === -1 || cEnd === -1) continue;
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var site = String(row[cSite] || '').trim();
        var lesson = cLesson !== -1 ? String(row[cLesson] || '').trim() : '';
        var startT = row[cStart];
        var endT = row[cEnd];
        if (cSetup !== -1) {
          var setup = String(row[cSetup] || '').toLowerCase();
          if (setup.indexOf('cancelled') >= 0 || setup.indexOf('cancel') >= 0) continue;
        }
        if (!site || !startT || !endT) continue;
        var dur = getDurationMin(startT, endT);
        if (dur <= 0) continue;
        var key = normalize(site + '|' + lesson);
        opsHub[key] = { site: site, lesson: lesson, dur: dur, allowed: dur + BUFFER_MINUTES };
      }
    }
    log.push('Ops Hub workshops: ' + Object.keys(opsHub).length);
  } catch (err) {
    log.push('ERROR loading Ops Hub: ' + err.message);
  }

  // STEP 2: Load check-ins (all tabs)
  var allCheckIns = [];
  try {
    var ciSS = SpreadsheetApp.openById(CHECKINS_ID);
    var ciSheets = ciSS.getSheets();
    for (var s = 0; s < ciSheets.length; s++) {
      var sheet = ciSheets[s];
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;
      var hdrRow = -1;
      var hdr = [];
      for (var h = 0; h < Math.min(20, data.length); h++) {
        var rowStr = data[h].map(function(c) { return String(c).toLowerCase().trim(); });
        for (var cc = 0; cc < rowStr.length; cc++) {
          if (rowStr[cc].indexOf('leader name') >= 0) { hdrRow = h; hdr = rowStr; break; }
        }
        if (hdrRow >= 0) break;
      }
      if (hdrRow === -1) continue;
      var cRegion = findCol(hdr, ['region']);
      var cWorkshop = findCol(hdr, ['workshop']);
      var cSchool = findCol(hdr, ['school']);
      var cLeader = findCol(hdr, ['leader name']);
      var cDate = findCol(hdr, ['date']);
      var cStatus = findCol(hdr, ['status']);
      if (cLeader === -1 || cStatus === -1) continue;
      for (var i = hdrRow + 1; i < data.length; i++) {
        var row = data[i];
        var leader = getVal(row, cLeader);
        var status = getVal(row, cStatus).toLowerCase();
        var workshop = getVal(row, cWorkshop);
        var school = getVal(row, cSchool);
        var region = getVal(row, cRegion);
        if (!leader || !status) continue;
        if (!workshop && !school) continue;
        var dt = null;
        if (cDate !== -1 && row[cDate]) dt = parseDate(row[cDate]);
        if (dt) {
          if (dt < startDate || dt > endDate) continue;
        }
        var worked = false;
        for (var w = 0; w < WORKED_STATUSES.length; w++) {
          if (status.indexOf(WORKED_STATUSES[w]) >= 0) { worked = true; break; }
        }
        allCheckIns.push({
          leader: leader, status: status, worked: worked,
          workshop: workshop, school: school, region: region,
          date: dt, dateStr: dt ? formatDt(dt) : 'N/A'
        });
      }
    }
    log.push('Check-in records: ' + allCheckIns.length);
  } catch (err) {
    log.push('ERROR loading check-ins: ' + err.message);
  }

  // STEP 2b: Deduplicate leader names (e.g., "Kelly O." → "Kelly Orzuna")
  allCheckIns = deduplicateLeaderNames_(allCheckIns);

  // STEP 3: Split into scoot vs non-scoot, calculate hours
  var leaderData = [];  // non-scoot sessions
  var scootData = [];   // scoot sessions

  for (var i = 0; i < allCheckIns.length; i++) {
    var rec = allCheckIns[i];

    // Only exact "scoot" status → SCOOT tab. All other statuses → leader tab.
    var statusTrimmed = rec.status.trim().toLowerCase();
    var isScoot = (statusTrimmed === 'scoot');
    var dur, allowed, src, ops;

    if (rec.worked) {
      ops = matchOpsHub(rec.school, rec.workshop, opsHub);
      if (ops) {
        dur = ops.dur;
        allowed = ops.allowed;
        src = 'Ops Hub (' + ops.site + ')';
      } else {
        dur = 60;
        allowed = 90;
        src = 'DEFAULT 1hr';
      }
    } else {
      ops = null;
      dur = 0;
      allowed = 0;
      src = 'ABSENT — 0hr';
    }

    var entry = {
      leader: rec.leader,
      region: rec.region || 'Unknown',
      date: rec.date,
      dateStr: rec.dateStr,
      workshop: rec.workshop,
      school: rec.school,
      status: rec.status,
      dur: dur,
      allowed: allowed,
      src: src,
      unmatched: rec.worked && !ops,
      absent: !rec.worked
    };

    if (isScoot) {
      scootData.push(entry);
    } else {
      leaderData.push(entry);
    }
  }

  log.push('Leader sessions: ' + leaderData.length);
  log.push('SCOOT sessions: ' + scootData.length);
  Logger.log('DEBUG — leaderData: ' + leaderData.length + ', scootData: ' + scootData.length);
  if (scootData.length > 0) {
    Logger.log('DEBUG — First scoot entry: ' + scootData[0].leader + ' / status=' + scootData[0].status);
  }

  // STEP 4: Write tabs with date-specific names
  var hvSS = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var dateRange = formatDt(startDate) + ' – ' + formatDt(endDate);
  var shortRange = (startDate.getMonth() + 1) + '/' + startDate.getDate() + '-' + (endDate.getMonth() + 1) + '/' + endDate.getDate();
  var trackerTabName = 'Hours Tracker ' + shortRange;
  var scootTabName = 'SCOOT Hours ' + shortRange;

  writeHoursTrackerTab_(hvSS, leaderData, scootData, dateRange, trackerTabName);
  SpreadsheetApp.flush();
  writeScootInvoiceTab2_(hvSS, scootData, dateRange, scootTabName);

  log.push('\n=== COMPLETE ===');
  Logger.log(log.join('\n'));

  try {
    SpreadsheetApp.getUi().alert('Hours Tracker Complete!',
      'Leader sessions: ' + leaderData.length +
      '\nSCOOT sessions: ' + scootData.length +
      '\n\n→ Check "' + trackerTabName + '" tab' +
      '\n→ Check "' + scootTabName + '" tab',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {
    Logger.log('Could not show alert: ' + e.message);
  }
}

/**
 * Groups session entries by region and aggregates per-leader summaries.
 */
function groupByRegion_(data) {
  var regions = {};
  for (var i = 0; i < data.length; i++) {
    var entry = data[i];
    var region = entry.region || 'Unknown';
    if (!regions[region]) regions[region] = { leaders: {} };
    var leaders = regions[region].leaders;
    if (!leaders[entry.leader]) leaders[entry.leader] = { name: entry.leader, sessions: 0, absences: 0, totalMin: 0, unmatched: 0, details: [] };
    var ldr = leaders[entry.leader];
    if (entry.absent) { ldr.absences++; } else { ldr.sessions++; ldr.totalMin += entry.allowed; if (entry.unmatched) ldr.unmatched++; }
    ldr.details.push(entry);
  }
  return regions;
}

/**
 * Groups session entries by LEADER (one entry per person).
 * Picks best region per leader (most frequent non-Unknown).
 * Returns flat array of leader objects sorted alphabetically.
 */
function groupByLeader_(data) {
  var leaders = {};

  for (var i = 0; i < data.length; i++) {
    var entry = data[i];
    var name = entry.leader;

    if (!leaders[name]) {
      leaders[name] = { name: name, sessions: 0, absences: 0, totalMin: 0, unmatched: 0, details: [], regionCounts: {} };
    }

    var ldr = leaders[name];
    if (entry.absent) {
      ldr.absences++;
    } else {
      ldr.sessions++;
      ldr.totalMin += entry.allowed;
      if (entry.unmatched) ldr.unmatched++;
    }
    ldr.details.push(entry);

    // Track region frequency to pick the best one
    var reg = entry.region || 'Unknown';
    ldr.regionCounts[reg] = (ldr.regionCounts[reg] || 0) + 1;
  }

  // Assign best region to each leader and build sorted array
  var result = [];
  for (var name in leaders) {
    var ldr = leaders[name];
    // Pick most frequent non-Unknown region
    var bestRegion = 'Unknown';
    var bestCount = 0;
    for (var reg in ldr.regionCounts) {
      if (reg !== 'Unknown' && ldr.regionCounts[reg] > bestCount) {
        bestCount = ldr.regionCounts[reg];
        bestRegion = reg;
      }
    }
    // If all are Unknown, keep Unknown
    if (bestRegion === 'Unknown' && ldr.regionCounts['Unknown']) bestRegion = 'Unknown';
    ldr.region = bestRegion;
    ldr.hrs = Math.round((ldr.totalMin / 60) * 100) / 100;
    delete ldr.regionCounts;
    result.push(ldr);
  }

  return result;
}

/**
 * Writes the main Hours Tracker tab using BATCH writes for performance.
 * Section 1: Expected Payroll Hours — leaders ranked most → least
 * Section 2: SCOOT Hours — separate billing
 * Section 3: Detailed Leader Breakdown — every session/absence
 * Section 4: All Unique Workshops
 */
function writeHoursTrackerTab_(hvSS, leaderData, scootData, dateRange, tabName) {
  var tab = hvSS.getSheetByName(tabName);
  if (!tab) {
    tab = hvSS.insertSheet(tabName);
  } else {
    tab.clear();
    tab.clearFormats();
    SpreadsheetApp.flush();
  }

  if (leaderData.length === 0 && scootData.length === 0) {
    tab.getRange(1, 1).setValue('No sessions found for ' + dateRange);
    tab.getRange(1, 1).setFontSize(12).setFontWeight('bold');
    return;
  }

  // Group by leader (one entry per person, best region picked automatically)
  var allLeaders = groupByLeader_(leaderData);

  // Grand totals
  var grandTotalMin = 0, grandSessions = 0, grandAbsences = 0;
  for (var i = 0; i < allLeaders.length; i++) {
    grandTotalMin += allLeaders[i].totalMin;
    grandSessions += allLeaders[i].sessions;
    grandAbsences += allLeaders[i].absences;
  }

  // Sort by hours descending
  var ranked = allLeaders.slice().sort(function(a, b) { return b.totalMin - a.totalMin; });

  // =====================================================================
  // SECTION 1: EXPECTED PAYROLL HOURS (batch write)
  // =====================================================================
  tab.getRange(1, 1).setValue('EXPECTED PAYROLL HOURS — ' + dateRange);
  tab.getRange(1, 1, 1, 8).merge().setFontSize(14).setFontWeight('bold')
    .setBackground('#1a73e8').setFontColor('#ffffff');

  tab.getRange(2, 1).setValue('Generated: ' + new Date().toLocaleString() + ' | ' + allLeaders.length + ' leaders | ' + grandSessions + ' sessions | ' + grandAbsences + ' absences');
  tab.getRange(2, 1, 1, 8).merge().setFontColor('#666666');

  var sumHeaders = ['#', 'Leader Name', 'Region', 'Sessions Worked', 'Absences', 'Expected Hours', 'Expected Time', 'Notes'];
  tab.getRange(4, 1, 1, 8).setValues([sumHeaders]).setFontWeight('bold').setBackground('#e8eaed');

  // Build batch data for Section 1
  var s1Data = [];
  for (var i = 0; i < ranked.length; i++) {
    var ldr = ranked[i];
    var notes = '';
    if (ldr.absences > 0) notes = ldr.absences + ' absence(s)';
    if (ldr.unmatched > 0) notes += (notes ? ', ' : '') + ldr.unmatched + ' unmatched';
    s1Data.push([i + 1, ldr.name, ldr.region, ldr.sessions, ldr.absences, ldr.hrs, formatMinutes_(ldr.totalMin), notes]);
  }
  // Grand total row
  var grandHrs = Math.round((grandTotalMin / 60) * 100) / 100;
  s1Data.push(['', 'GRAND TOTAL', '', grandSessions, grandAbsences, grandHrs, formatMinutes_(grandTotalMin), allLeaders.length + ' leaders']);

  // Batch write Section 1
  if (s1Data.length > 0) {
    tab.getRange(5, 1, s1Data.length, 8).setValues(s1Data);
    // Format: alternating rows + grand total
    for (var i = 0; i < ranked.length; i++) {
      if (i % 2 === 1) tab.getRange(5 + i, 1, 1, 8).setBackground('#f8f9fa');
      if (ranked[i].absences > 0) tab.getRange(5 + i, 5).setFontColor('#d93025').setFontWeight('bold');
    }
    // Grand total row styling
    var gtRow = 5 + ranked.length;
    tab.getRange(gtRow, 1, 1, 8).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  }

  var row = 5 + s1Data.length;

  // =====================================================================
  // SECTION 2: SCOOT HOURS (batch write)
  // =====================================================================
  var scootStart = row + 2;
  tab.getRange(scootStart, 1).setValue('SCOOT HOURS — Billed Separately (3hr flat per session)');
  tab.getRange(scootStart, 1, 1, 8).merge().setFontSize(12).setFontWeight('bold')
    .setBackground('#ff6d01').setFontColor('#ffffff');

  if (scootData.length === 0) {
    tab.getRange(scootStart + 1, 1).setValue('No SCOOT sessions found in this period.');
    tab.getRange(scootStart + 1, 1, 1, 8).merge().setFontColor('#666666');
    row = scootStart + 2;
  } else {
    var scootHeaders = ['#', 'Person', 'School', 'Date', 'Workshop', 'Status', 'Billed Hours', ''];
    tab.getRange(scootStart + 1, 1, 1, 8).setValues([scootHeaders]).setFontWeight('bold').setBackground('#e8eaed');

    scootData.sort(function(a, b) {
      if (a.leader.toLowerCase() < b.leader.toLowerCase()) return -1;
      if (a.leader.toLowerCase() > b.leader.toLowerCase()) return 1;
      if (a.date && b.date) return a.date.getTime() - b.date.getTime();
      return 0;
    });

    // Build batch
    var s2Data = [];
    for (var i = 0; i < scootData.length; i++) {
      var d = scootData[i];
      s2Data.push([i + 1, d.leader, d.school, d.dateStr, d.workshop, d.status, 3, '']);
    }
    // Total + per-person subtotals
    s2Data.push(['', 'SCOOT TOTAL', '', '', '', '', scootData.length * 3, scootData.length + ' sessions']);
    var scootCounts = {};
    for (var i = 0; i < scootData.length; i++) {
      scootCounts[scootData[i].leader] = (scootCounts[scootData[i].leader] || 0) + 1;
    }
    var scootNames = Object.keys(scootCounts).sort();
    for (var n = 0; n < scootNames.length; n++) {
      s2Data.push(['', scootNames[n], '', '', '', '', scootCounts[scootNames[n]] * 3, scootCounts[scootNames[n]] + ' sessions']);
    }

    var s2Start = scootStart + 2;
    tab.getRange(s2Start, 1, s2Data.length, 8).setValues(s2Data);
    // Format SCOOT total + subtotals
    var scootTotalRow = s2Start + scootData.length;
    tab.getRange(scootTotalRow, 1, 1, 8).setFontWeight('bold').setBackground('#ff6d01').setFontColor('#ffffff');
    for (var n = 0; n < scootNames.length; n++) {
      tab.getRange(scootTotalRow + 1 + n, 1, 1, 8).setBackground('#fef7cd');
    }
    row = s2Start + s2Data.length;
  }

  // =====================================================================
  // SECTION 3: DETAILED LEADER BREAKDOWN — one row per leader (batch)
  // =====================================================================
  var detailStart = row + 2;
  tab.getRange(detailStart, 1).setValue('DETAILED LEADER BREAKDOWN (' + dateRange + ')');
  tab.getRange(detailStart, 1, 1, 7).merge().setFontSize(12).setFontWeight('bold')
    .setBackground('#34a853').setFontColor('#ffffff');

  var detHeaders = ['#', 'Leader Name', 'Region', 'Sessions', 'Absences', 'Expected Time', 'All Workshops & Dates'];
  tab.getRange(detailStart + 1, 1, 1, 7).setValues([detHeaders]).setFontWeight('bold').setBackground('#e8eaed');

  // Build one row per leader with all sessions in a text cell
  var s3Data = [];
  for (var i = 0; i < ranked.length; i++) {
    var ldr = ranked[i];

    // Sort details by date
    ldr.details.sort(function(a, b) {
      if (!a.date && !b.date) return 0;
      if (!a.date) return 1;
      if (!b.date) return -1;
      return a.date.getTime() - b.date.getTime();
    });

    // Build workshop detail text
    var detailParts = [];
    for (var d = 0; d < ldr.details.length; d++) {
      var det = ldr.details[d];
      if (det.absent) {
        detailParts.push(det.dateStr + ': ABSENT — ' + det.workshop + ' @ ' + det.school + ' (' + det.status + ')');
      } else {
        detailParts.push(det.dateStr + ': ' + det.workshop + ' @ ' + det.school + ' (' + det.status + ', ' + formatMinutes_(det.allowed) + ')');
      }
    }

    s3Data.push([
      i + 1,
      ldr.name,
      ldr.region,
      ldr.sessions,
      ldr.absences,
      formatMinutes_(ldr.totalMin),
      detailParts.join('\n')
    ]);
  }

  // Batch write
  var dStart = detailStart + 2;
  if (s3Data.length > 0) {
    tab.getRange(dStart, 1, s3Data.length, 7).setValues(s3Data);
    tab.getRange(dStart, 7, s3Data.length, 1).setWrap(true); // wrap the details column
    // Highlight absences
    for (var i = 0; i < ranked.length; i++) {
      if (ranked[i].absences > 0) {
        tab.getRange(dStart + i, 5).setFontColor('#d93025').setFontWeight('bold');
      }
      if (i % 2 === 1) tab.getRange(dStart + i, 1, 1, 7).setBackground('#f8f9fa');
    }
  }

  var dRow = dStart + s3Data.length;

  // =====================================================================
  // SECTION 4: ALL UNIQUE WORKSHOPS (batch write)
  // =====================================================================
  var wsStart = dRow + 2;
  tab.getRange(wsStart, 1).setValue('ALL UNIQUE WORKSHOPS');
  tab.getRange(wsStart, 1, 1, 4).merge().setFontSize(12).setFontWeight('bold').setBackground('#fbbc04');

  tab.getRange(wsStart + 1, 1, 1, 4).setValues([['#', 'Workshop', 'School(s)', 'Sessions']]).setFontWeight('bold').setBackground('#e8eaed');

  var allData = leaderData.concat(scootData);
  var wsMap = {};
  for (var i = 0; i < allData.length; i++) {
    if (allData[i].absent) continue;
    var ws = allData[i].workshop || '(none)';
    if (!wsMap[ws]) wsMap[ws] = { schools: {}, count: 0 };
    wsMap[ws].count++;
    var sch = allData[i].school || '';
    if (sch) wsMap[ws].schools[sch] = true;
  }
  var wsNames = Object.keys(wsMap).sort();
  if (wsNames.length > 0) {
    var wsData = [];
    for (var w = 0; w < wsNames.length; w++) {
      var ws = wsMap[wsNames[w]];
      wsData.push([w + 1, wsNames[w], Object.keys(ws.schools).sort().join(', '), ws.count]);
    }
    tab.getRange(wsStart + 2, 1, wsData.length, 4).setValues(wsData);
  }

  // Auto-resize and freeze
  for (var c = 1; c <= 10; c++) tab.autoResizeColumn(c);
  tab.setFrozenRows(4);
}

/**
 * Creates/clears the SCOOT Hours tab and writes the invoice layout.
 */
function writeScootInvoiceTab2_(hvSS, data, dateRange, tabName) {
  var tab = hvSS.getSheetByName(tabName);
  if (!tab) {
    tab = hvSS.insertSheet(tabName);
  } else {
    tab.clear();
    tab.clearFormats();
    SpreadsheetApp.flush();
  }
  if (data.length === 0) {
    tab.getRange(1, 1).setValue('No SCOOT sessions found for ' + dateRange);
    tab.getRange(1, 1).setFontSize(12).setFontWeight('bold');
    return;
  }
  writeScootInvoiceTab_(tab, data, dateRange);
}

/**
 * Writes a simple SCOOT invoice-verification tab.
 * SCOOT bills 3 hours per session, so we just need: Person, School, Date.
 */
function writeScootInvoiceTab_(tab, data, dateRange) {
  // Title
  tab.getRange(1, 1).setValue('SCOOT HOURS — ' + dateRange);
  tab.getRange(1, 1, 1, 4).merge();
  tab.getRange(1, 1).setFontSize(14).setFontWeight('bold')
    .setBackground('#ff6d01').setFontColor('#ffffff');

  tab.getRange(2, 1).setValue('Generated: ' + new Date().toLocaleString() + ' | SCOOT bills 3 hours per session | Total sessions: ' + data.length);
  tab.getRange(2, 1, 1, 4).merge().setFontColor('#666666');

  // Headers
  var headers = ['Person', 'School', 'Date', 'Workshop'];
  tab.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');

  // Sort by person, then date
  data.sort(function(a, b) {
    var nameA = a.leader.toLowerCase(), nameB = b.leader.toLowerCase();
    if (nameA < nameB) return -1;
    if (nameA > nameB) return 1;
    if (a.date && b.date) return a.date.getTime() - b.date.getTime();
    return 0;
  });

  // Write rows with daily subtotals
  var row = 5;
  var currentPerson = '';
  var currentDate = '';
  var personColor = 0;
  var daySessionCount = 0;
  var dayStartRow = row;

  for (var i = 0; i < data.length; i++) {
    var d = data[i];

    // Person changed — flush daily subtotal for previous person's last date
    if (d.leader !== currentPerson) {
      if (daySessionCount >= 2) {
        tab.getRange(row, 1, 1, 4).setValues([['', '', currentDate + ' Total', daySessionCount + ' sessions — ' + formatMinutes_(daySessionCount * 180)]]);
        tab.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
        row++;
      }
      currentPerson = d.leader;
      currentDate = d.dateStr;
      daySessionCount = 0;
      personColor = (personColor + 1) % REGION_TINTS.length;
    } else if (d.dateStr !== currentDate) {
      // Same person, date changed — flush subtotal for previous date
      if (daySessionCount >= 2) {
        tab.getRange(row, 1, 1, 4).setValues([['', '', currentDate + ' Total', daySessionCount + ' sessions — ' + formatMinutes_(daySessionCount * 180)]]);
        tab.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
        row++;
      }
      currentDate = d.dateStr;
      daySessionCount = 0;
    }

    daySessionCount++;
    tab.getRange(row, 1, 1, 4).setValues([[d.leader, d.school, d.dateStr, d.workshop]]);
    tab.getRange(row, 1, 1, 4).setBackground(REGION_TINTS[personColor]);
    row++;
  }

  // Flush final daily subtotal
  if (daySessionCount >= 2) {
    tab.getRange(row, 1, 1, 4).setValues([['', '', currentDate + ' Total', daySessionCount + ' sessions — ' + formatMinutes_(daySessionCount * 180)]]);
    tab.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#f0f0f0');
    row++;
  }

  // Summary: count per person
  row += 1;
  tab.getRange(row, 1).setValue('SUMMARY');
  tab.getRange(row, 1, 1, 4).merge();
  tab.getRange(row, 1).setFontSize(12).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  row++;
  tab.getRange(row, 1, 1, 3).setValues([['Person', 'Sessions', 'Billed Hours (3h each)']]).setFontWeight('bold').setBackground('#e8eaed');
  row++;

  var counts = {};
  for (var i = 0; i < data.length; i++) {
    counts[data[i].leader] = (counts[data[i].leader] || 0) + 1;
  }
  var names = Object.keys(counts).sort();
  for (var n = 0; n < names.length; n++) {
    tab.getRange(row, 1, 1, 3).setValues([[names[n], counts[names[n]], counts[names[n]] * 3]]);
    row++;
  }

  // Total row
  tab.getRange(row, 1, 1, 3).setValues([['TOTAL', data.length, data.length * 3]]);
  tab.getRange(row, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaed');

  // Auto-resize and freeze
  for (var c = 1; c <= 4; c++) tab.autoResizeColumn(c);
  tab.setFrozenRows(4);
}

// =====================================================
// PHASE 3: SLACK ALERTS
// =====================================================

/**
 * Sends a daily Slack alert at 8 PM with check-in summary and flagged leaders.
 * Can be run manually or via time-driven trigger.
 */
function sendDailySlackAlert() {
  if (SLACK_WEBHOOK_URL === 'YOUR_SLACK_WEBHOOK_URL_HERE') {
    try {
      SpreadsheetApp.getUi().alert('Slack webhook URL not configured.\n\nEdit the script and replace YOUR_SLACK_WEBHOOK_URL_HERE with your actual webhook URL.\n\nSetup: api.slack.com → Your Apps → Incoming Webhooks');
    } catch(e) {
      Logger.log('Slack webhook URL not configured.');
    }
    return;
  }

  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var today = new Date();
  var dateStr = formatDt(today);

  // Read Auto Verification tab
  var autoTab = ss.getSheetByName('Auto Verification');
  if (!autoTab) {
    sendSlack('⚠️ Hours Verification — No "Auto Verification" tab found. Run the report first.');
    return;
  }

  var autoData = autoTab.getDataRange().getValues();

  // Count leaders from summary section (rows 5+)
  var totalLeaders = 0;
  var leaderList = [];
  for (var i = 4; i < autoData.length; i++) {
    var name = String(autoData[i][0] || '').trim();
    var hrs = parseFloat(autoData[i][2]);
    if (!name || name.length < 2 || isNaN(hrs)) continue;
    if (name.toUpperCase().indexOf('PAYROLL') >= 0) break;
    if (name.toUpperCase().indexOf('SESSION LOG') >= 0) break;
    totalLeaders++;
    leaderList.push({ name: name, allowedHours: hrs });
  }

  // Read Discrepancies tab if it exists
  var flaggedLeaders = [];
  var discTab = ss.getSheetByName(DISCREPANCIES_TAB);
  if (discTab) {
    var discData = discTab.getDataRange().getValues();
    // Flagged section starts at row 6 (0-indexed row 5)
    for (var i = 5; i < discData.length; i++) {
      var gustoName = String(discData[i][0] || '').trim();
      var reported = parseFloat(discData[i][2]);
      var allowed = parseFloat(discData[i][3]);
      var overage = String(discData[i][4] || '').trim();

      if (!gustoName || gustoName.length < 2) continue;
      // Stop at the "ALL COMPARISONS" section
      if (gustoName.toUpperCase().indexOf('ALL COMPARISONS') >= 0) break;
      if (isNaN(reported) || isNaN(allowed)) continue;

      flaggedLeaders.push({
        name: gustoName,
        allowed: allowed,
        reported: reported,
        overage: overage
      });
    }
  }

  // Build Slack message
  var blocks = [];
  blocks.push({
    type: 'header',
    text: { type: 'plain_text', text: '📊 Hours Verification — Daily Report (' + dateStr + ')' }
  });

  var summaryText = '✅ *' + totalLeaders + ' leaders* in current verification report';

  if (flaggedLeaders.length > 0) {
    summaryText += '\n⚠️ *' + flaggedLeaders.length + ' leaders flagged:*';
    for (var i = 0; i < Math.min(flaggedLeaders.length, 10); i++) {
      var fl = flaggedLeaders[i];
      summaryText += '\n  • ' + fl.name + ' — ' + fl.allowed + 'h allowed, ' + fl.reported + 'h reported (' + fl.overage + ' over)';
    }
    if (flaggedLeaders.length > 10) {
      summaryText += '\n  _...and ' + (flaggedLeaders.length - 10) + ' more_';
    }
  } else {
    summaryText += '\n✅ No leaders flagged for overages';
    if (!discTab) {
      summaryText += '\n_Note: Run "Compare Gusto Hours" to check for discrepancies_';
    }
  }

  summaryText += '\n\n📋 <' + SHEET_URL + '|View Full Report>';

  blocks.push({
    type: 'section',
    text: { type: 'mrkdwn', text: summaryText }
  });

  sendSlack({ blocks: blocks });
  Logger.log('Daily Slack alert sent: ' + dateStr);
}

/**
 * Sends a weekly payroll summary (designed for Friday evenings).
 */
function sendWeeklySlackSummary() {
  if (SLACK_WEBHOOK_URL === 'YOUR_SLACK_WEBHOOK_URL_HERE') {
    try {
      SpreadsheetApp.getUi().alert('Slack webhook URL not configured.\n\nEdit the script and replace YOUR_SLACK_WEBHOOK_URL_HERE with your actual webhook URL.');
    } catch(e) {
      Logger.log('Slack webhook URL not configured.');
    }
    return;
  }

  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var today = new Date();

  // Calculate week start (Monday)
  var weekStart = new Date(today);
  weekStart.setDate(today.getDate() - today.getDay() + 1); // Monday
  var weekStartStr = formatDt(weekStart);

  // Read Auto Verification tab
  var autoTab = ss.getSheetByName('Auto Verification');
  if (!autoTab) {
    sendSlack('⚠️ Weekly Summary — No "Auto Verification" tab found. Run the report first.');
    return;
  }

  var autoData = autoTab.getDataRange().getValues();

  var totalLeaders = 0;
  var totalAllowedMin = 0;
  for (var i = 4; i < autoData.length; i++) {
    var name = String(autoData[i][0] || '').trim();
    var hrs = parseFloat(autoData[i][2]);
    if (!name || name.length < 2 || isNaN(hrs)) continue;
    if (name.toUpperCase().indexOf('PAYROLL') >= 0) break;
    if (name.toUpperCase().indexOf('SESSION LOG') >= 0) break;
    totalLeaders++;
    totalAllowedMin += hrs * 60;
  }

  var totalAllowedHrs = Math.round((totalAllowedMin / 60) * 10) / 10;

  // Read Discrepancies tab
  var flaggedLeaders = [];
  var discTab = ss.getSheetByName(DISCREPANCIES_TAB);
  if (discTab) {
    var discData = discTab.getDataRange().getValues();
    for (var i = 5; i < discData.length; i++) {
      var gustoName = String(discData[i][0] || '').trim();
      var reported = parseFloat(discData[i][2]);
      var allowed = parseFloat(discData[i][3]);
      var overage = String(discData[i][4] || '').trim();

      if (!gustoName || gustoName.length < 2) continue;
      if (gustoName.toUpperCase().indexOf('ALL COMPARISONS') >= 0) break;
      if (isNaN(reported) || isNaN(allowed)) continue;

      flaggedLeaders.push({
        name: gustoName,
        allowed: allowed,
        reported: reported,
        overage: overage
      });
    }
  }

  // Build weekly message
  var blocks = [];
  blocks.push({
    type: 'header',
    text: { type: 'plain_text', text: '💰 Payroll Verification Summary — Week of ' + weekStartStr }
  });

  var bodyText = '*Leaders who worked:* ' + totalLeaders +
    '\n*Total allowed hours:* ' + totalAllowedHrs + 'h' +
    '\n*Flagged for review:* ' + flaggedLeaders.length + ' leaders';

  if (flaggedLeaders.length > 0) {
    bodyText += '\n\n*Top overages:*';
    for (var i = 0; i < Math.min(flaggedLeaders.length, 5); i++) {
      var fl = flaggedLeaders[i];
      bodyText += '\n' + (i + 1) + '. ' + fl.name + ': ' + fl.overage + ' over allowed';
    }
    if (flaggedLeaders.length > 5) {
      bodyText += '\n_...and ' + (flaggedLeaders.length - 5) + ' more_';
    }
  }

  bodyText += '\n\n👉 <' + SHEET_URL + '|Review Full Report>';

  blocks.push({
    type: 'section',
    text: { type: 'mrkdwn', text: bodyText }
  });

  // Add divider and context
  blocks.push({ type: 'divider' });
  blocks.push({
    type: 'context',
    elements: [{ type: 'mrkdwn', text: '_Auto-generated by Kodely Hours Verification • ' + new Date().toLocaleString() + '_' }]
  });

  sendSlack({ blocks: blocks });
  Logger.log('Weekly Slack summary sent: week of ' + weekStartStr);
}

/**
 * Posts a message to Slack via Incoming Webhook.
 * @param {Object|string} payload - Slack message payload (object with blocks) or simple string
 */
function sendSlack(payload) {
  if (SLACK_WEBHOOK_URL === 'YOUR_SLACK_WEBHOOK_URL_HERE') {
    Logger.log('Slack not configured — would have sent: ' + JSON.stringify(payload));
    return;
  }

  var body;
  if (typeof payload === 'string') {
    body = JSON.stringify({ text: payload });
  } else {
    // Ensure there's a fallback text for notifications
    if (!payload.text && payload.blocks && payload.blocks.length > 0) {
      payload.text = 'Kodely Hours Verification Alert';
    }
    body = JSON.stringify(payload);
  }

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: body,
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
  var code = response.getResponseCode();
  if (code !== 200) {
    Logger.log('Slack webhook error (' + code + '): ' + response.getContentText());
  }
}

/**
 * Sets up time-driven triggers:
 * - Daily at 8 PM: sendDailySlackAlert
 * - Weekly on Friday at 6 PM: sendWeeklySlackSummary
 */
function setupTriggers() {
  // Remove existing triggers for these functions first
  removeTriggers();

  // Daily at 8 PM
  ScriptApp.newTrigger('sendDailySlackAlert')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .create();

  // Weekly on Friday at 6 PM
  ScriptApp.newTrigger('sendWeeklySlackSummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(18)
    .create();

  // Hours Tracker daily at 6 PM
  ScriptApp.newTrigger('generateHoursTracker02_05to02_18')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  Logger.log('Triggers created: daily at 8 PM + weekly Friday at 6 PM + hours tracker daily at 6 PM');
  try {
    SpreadsheetApp.getUi().alert('Triggers set up!\n\n• Daily alert: Every day at 8 PM\n• Weekly summary: Every Friday at 6 PM\n• Hours Tracker: Daily at 6 PM\n\nMake sure your Slack webhook URL is configured in the script.');
  } catch(e) {
    Logger.log('Triggers created successfully.');
  }
}

/**
 * Removes all triggers for the daily and weekly alert functions.
 */
function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var funcName = triggers[i].getHandlerFunction();
    if (funcName === 'sendDailySlackAlert' || funcName === 'sendWeeklySlackSummary' || funcName === 'generateHoursTracker02_05to02_18') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  Logger.log('Removed existing alert triggers.');
}

// =====================================================
// UTILITIES
// =====================================================
function normalize(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9|]/g, '');
}

function getVal(row, idx) {
  if (idx < 0 || idx >= row.length) return '';
  return String(row[idx] || '').trim();
}

function findCol(headers, names) {
  for (var i = 0; i < headers.length; i++) {
    for (var n = 0; n < names.length; n++) {
      if (headers[i].indexOf(names[n]) >= 0) return i;
    }
  }
  return -1;
}

function getDurationMin(startTime, endTime) {
  var s = toMinutes(startTime);
  var e = toMinutes(endTime);
  if (s === null || e === null) return 0;
  return e - s;
}

function toMinutes(v) {
  if (v instanceof Date) {
    return v.getHours() * 60 + v.getMinutes();
  }
  var str = String(v).trim().toUpperCase();
  var match = str.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/);
  if (!match) return null;
  var h = parseInt(match[1]);
  var m = parseInt(match[2]);
  if (match[3] === 'PM' && h !== 12) h += 12;
  if (match[3] === 'AM' && h === 12) h = 0;
  return h * 60 + m;
}

function matchOpsHub(school, workshop, hub) {
  var key = normalize(school + '|' + workshop);
  if (hub[key]) return hub[key];

  var sn = school.toLowerCase().replace(/[^a-z0-9]/g, '');
  var wn = workshop.toLowerCase().replace(/[^a-z0-9]/g, '');
  var bestMatch = null;
  var bestScore = 0;

  for (var k in hub) {
    var v = hub[k];
    var siteN = v.site.toLowerCase().replace(/[^a-z0-9]/g, '');
    var lesN = v.lesson.toLowerCase().replace(/[^a-z0-9]/g, '');
    var score = 0;

    if (siteN && sn) {
      if (siteN === sn) score += 4;
      else if (siteN.indexOf(sn) >= 0 || sn.indexOf(siteN) >= 0) score += 3;
      else if (siteN.length > 5 && sn.length > 5 && siteN.substring(0, 6) === sn.substring(0, 6)) score += 2;
    }
    if (lesN && wn) {
      if (lesN === wn) score += 4;
      else if (lesN.indexOf(wn) >= 0 || wn.indexOf(lesN) >= 0) score += 3;
      else if (lesN.length > 5 && wn.length > 5 && lesN.substring(0, 6) === wn.substring(0, 6)) score += 2;
    }

    if (score > bestScore) { bestScore = score; bestMatch = v; }
  }
  return bestScore >= 3 ? bestMatch : null;
}

function parseDate(val) {
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  var s = String(val).trim();

  // "Feb-9", "Feb 9", "Feb-10" etc
  var m = s.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[- ](\d{1,2})$/i);
  if (m) {
    var months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    return new Date(2026, months[m[1].toLowerCase()], parseInt(m[2]));
  }

  // "2/5/2026", "02/05/2026" etc
  var p = new Date(s);
  if (!isNaN(p.getTime())) return p;

  return null;
}

function formatDt(d) {
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
}

/**
 * Formats minutes as "1hr 30mins" style for timesheet readability.
 * Examples: 90 → "1hr 30mins", 60 → "1hr", 45 → "45mins", 0 → "0mins"
 */
function formatMinutes_(min) {
  var h = Math.floor(min / 60);
  var m = Math.round(min % 60);
  if (h === 0) return m + 'mins';
  if (m === 0) return h + 'hr';
  return h + 'hr ' + m + 'mins';
}

/**
 * Deduplicates leader names in check-in records using fuzzy matching.
 * When two names match (e.g., "Kelly O." and "Kelly Orzuna"),
 * keeps the longer/more complete name as canonical.
 */
function deduplicateLeaderNames_(checkIns) {
  // Collect unique names
  var nameSet = {};
  for (var i = 0; i < checkIns.length; i++) {
    nameSet[checkIns[i].leader] = true;
  }
  var uniqueNames = Object.keys(nameSet);

  // Sort by length descending so longer names become canonical first
  uniqueNames.sort(function(a, b) { return b.length - a.length; });

  // Build canonical map by checking each name against existing canonicals
  var canonicalMap = {};   // shortName → canonicalName
  var canonicalLookup = {}; // normalizedName → {name: canonicalName}

  for (var i = 0; i < uniqueNames.length; i++) {
    var name = uniqueNames[i];
    var normName = name.toLowerCase().replace(/[^a-z]/g, '');

    // Check if this name matches any existing canonical name
    var match = fuzzyMatchName(name, canonicalLookup);

    if (match) {
      // Map this (shorter) name to the existing canonical
      canonicalMap[name] = match.name;
      Logger.log('Merged name: "' + name + '" → "' + match.name + '"');
    } else {
      // No match — this becomes a new canonical name
      canonicalLookup[normName] = { name: name };
    }
  }

  // Apply mappings to all check-in records
  var mergeCount = Object.keys(canonicalMap).length;
  if (mergeCount > 0) {
    Logger.log('Name dedup: ' + mergeCount + ' name(s) merged');
    for (var i = 0; i < checkIns.length; i++) {
      if (canonicalMap[checkIns[i].leader]) {
        checkIns[i].leader = canonicalMap[checkIns[i].leader];
      }
    }
  }

  return checkIns;
}

// =====================================================
// SCOOT INVOICE VERIFICATION
// =====================================================

var SCOOT_INVOICE_TAB = 'SCOOT Invoices';

/**
 * Creates the SCOOT Invoices tab with headers and instructions.
 */
function setupScootInvoiceImport() {
  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var tab = ss.getSheetByName(SCOOT_INVOICE_TAB);
  if (!tab) {
    tab = ss.insertSheet(SCOOT_INVOICE_TAB);
  } else {
    tab.clear();
    tab.clearFormats();
  }

  tab.getRange(1, 1).setValue('SCOOT INVOICE DATA — Paste below');
  tab.getRange(1, 1, 1, 6).merge().setFontSize(14).setFontWeight('bold')
    .setBackground('#ff6d01').setFontColor('#ffffff');

  tab.getRange(2, 1).setValue('Paste your SCOOT invoice data starting at row 4. Match the columns below.');
  tab.getRange(2, 1, 1, 6).merge().setFontColor('#666666');

  var headers = ['Paid Status', 'Invoice Number', 'Invoice Date', 'Issued To (School)', 'Amount', 'Paid Date'];
  tab.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');

  // Sample row
  tab.getRange(5, 1, 1, 6).setValues([['Not Paid', '154545', '2/16/2026', 'Maryland Elementary School', '$522.00', '']]);
  tab.getRange(5, 1, 1, 6).setFontColor('#999999').setFontStyle('italic');

  tab.setFrozenRows(4);
  for (var c = 1; c <= 6; c++) tab.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert('SCOOT Invoices tab created!\n\nPaste your SCOOT invoice data starting at row 5.\nThen run "Verify SCOOT Invoices" from the menu.');
}

/**
 * Reads SCOOT invoices and check-in data, cross-references, and writes a verification report.
 */
function verifyScootInvoices() {
  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);

  // Step 1: Read SCOOT invoices
  var invTab = ss.getSheetByName(SCOOT_INVOICE_TAB);
  if (!invTab) {
    SpreadsheetApp.getUi().alert('No "' + SCOOT_INVOICE_TAB + '" tab found.\n\nRun "Setup SCOOT Invoice Tab" first, then paste your data.');
    return;
  }

  var invData = invTab.getDataRange().getValues();
  var invoices = [];
  for (var i = 4; i < invData.length; i++) {
    var row = invData[i];
    var school = String(row[3] || '').trim();
    var dateVal = row[2];
    var amount = parseFloat(String(row[4] || '').replace(/[$,]/g, ''));
    var invNum = String(row[1] || '').trim();
    var status = String(row[0] || '').trim();

    if (!school || isNaN(amount)) continue;

    var dt = parseDate(dateVal);
    var dateStr = dt ? formatDt(dt) : String(dateVal);

    invoices.push({
      school: school,
      date: dt,
      dateStr: dateStr,
      amount: amount,
      invNum: invNum,
      status: status
    });
  }

  if (invoices.length === 0) {
    SpreadsheetApp.getUi().alert('No valid invoice data found.\n\nMake sure columns match: Paid Status | Invoice # | Invoice Date | School | Amount');
    return;
  }

  // Step 2: Load SCOOT check-ins for the invoice date range
  var minDate = null, maxDate = null;
  for (var i = 0; i < invoices.length; i++) {
    if (invoices[i].date) {
      if (!minDate || invoices[i].date < minDate) minDate = invoices[i].date;
      if (!maxDate || invoices[i].date > maxDate) maxDate = invoices[i].date;
    }
  }
  // Expand range by a week on each side
  var startDate = new Date(minDate.getTime() - 7 * 86400000);
  var endDate = new Date(maxDate.getTime() + 7 * 86400000);
  endDate.setHours(23, 59, 59);

  // Load check-ins
  var allCheckIns = [];
  var ciSS = SpreadsheetApp.openById(CHECKINS_ID);
  var ciSheets = ciSS.getSheets();
  for (var s = 0; s < ciSheets.length; s++) {
    var sheet = ciSheets[s];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;
    var hdrRow = -1, hdr = [];
    for (var h = 0; h < Math.min(20, data.length); h++) {
      var rowStr = data[h].map(function(c) { return String(c).toLowerCase().trim(); });
      for (var cc = 0; cc < rowStr.length; cc++) {
        if (rowStr[cc].indexOf('leader name') >= 0) { hdrRow = h; hdr = rowStr; break; }
      }
      if (hdrRow >= 0) break;
    }
    if (hdrRow === -1) continue;
    var cLeader = findCol(hdr, ['leader name']);
    var cDate = findCol(hdr, ['date']);
    var cStatus = findCol(hdr, ['status']);
    var cSchool = findCol(hdr, ['school']);
    var cWorkshop = findCol(hdr, ['workshop']);
    if (cLeader === -1 || cStatus === -1) continue;

    for (var i = hdrRow + 1; i < data.length; i++) {
      var row = data[i];
      var leader = getVal(row, cLeader);
      var status = getVal(row, cStatus).toLowerCase().trim();
      var school = getVal(row, cSchool);
      if (!leader || status !== 'scoot') continue;

      var dt = null;
      if (cDate !== -1 && row[cDate]) dt = parseDate(row[cDate]);
      if (dt && (dt < startDate || dt > endDate)) continue;

      allCheckIns.push({
        leader: leader,
        school: school,
        workshop: getVal(row, cWorkshop),
        date: dt,
        dateStr: dt ? formatDt(dt) : 'N/A'
      });
    }
  }

  // Step 3: Build check-in lookup by date+school
  var ciByDateSchool = {};
  for (var i = 0; i < allCheckIns.length; i++) {
    var ci = allCheckIns[i];
    var key = (ci.dateStr + '|' + ci.school).toLowerCase().replace(/[^a-z0-9|]/g, '');
    if (!ciByDateSchool[key]) ciByDateSchool[key] = [];
    ciByDateSchool[key].push(ci);
  }

  // Step 4: Match invoices to check-ins
  var results = [];
  var matched = 0, unmatched = 0;

  for (var i = 0; i < invoices.length; i++) {
    var inv = invoices[i];
    var key = (inv.dateStr + '|' + inv.school).toLowerCase().replace(/[^a-z0-9|]/g, '');

    // Try exact match first
    var sessions = ciByDateSchool[key] || [];

    // If no exact match, try fuzzy school name match on same date
    if (sessions.length === 0) {
      var invSchoolNorm = inv.school.toLowerCase().replace(/[^a-z0-9]/g, '');
      for (var k in ciByDateSchool) {
        if (k.indexOf(inv.dateStr.toLowerCase().replace(/[^a-z0-9]/g, '')) === 0) {
          var ciSchoolPart = k.split('|')[1] || '';
          if (ciSchoolPart.indexOf(invSchoolNorm) >= 0 || invSchoolNorm.indexOf(ciSchoolPart) >= 0) {
            sessions = ciByDateSchool[k];
            break;
          }
        }
      }
    }

    var sessionCount = sessions.length;
    var leaders = [];
    for (var j = 0; j < sessions.length; j++) {
      if (leaders.indexOf(sessions[j].leader) === -1) leaders.push(sessions[j].leader);
    }

    var matchStatus;
    if (sessionCount === 0) {
      matchStatus = 'NO MATCH — not in our check-ins';
      unmatched++;
    } else {
      matchStatus = 'MATCHED — ' + sessionCount + ' session(s)';
      matched++;
    }

    results.push({
      invDate: inv.dateStr,
      invSchool: inv.school,
      invAmount: inv.amount,
      invNum: inv.invNum,
      invStatus: inv.status,
      sessions: sessionCount,
      leaders: leaders.join(', '),
      matchStatus: matchStatus
    });
  }

  // Step 5: Find check-ins with NO matching invoice
  var ciUsedKeys = {};
  for (var i = 0; i < invoices.length; i++) {
    var inv = invoices[i];
    var key = (inv.dateStr + '|' + inv.school).toLowerCase().replace(/[^a-z0-9|]/g, '');
    ciUsedKeys[key] = true;
  }

  var unbilledSessions = [];
  for (var k in ciByDateSchool) {
    if (!ciUsedKeys[k]) {
      var sessions = ciByDateSchool[k];
      for (var j = 0; j < sessions.length; j++) {
        unbilledSessions.push(sessions[j]);
      }
    }
  }

  // Step 6: Write verification report
  var rptTabName = 'SCOOT Invoice Verify';
  var rpt = ss.getSheetByName(rptTabName);
  if (!rpt) { rpt = ss.insertSheet(rptTabName); } else { rpt.clear(); rpt.clearFormats(); }

  rpt.getRange(1, 1).setValue('SCOOT INVOICE VERIFICATION');
  rpt.getRange(1, 1, 1, 8).merge().setFontSize(14).setFontWeight('bold')
    .setBackground('#ff6d01').setFontColor('#ffffff');

  rpt.getRange(2, 1).setValue('Generated: ' + new Date().toLocaleString() + ' | Invoices: ' + invoices.length + ' | Matched: ' + matched + ' | No Match: ' + unmatched + ' | Unbilled sessions: ' + unbilledSessions.length);
  rpt.getRange(2, 1, 1, 8).merge().setFontColor('#666666');

  // Section 1: Invoice comparison
  var headers = ['Invoice Date', 'School', 'Invoice Amount', 'Invoice #', 'Paid Status', 'Our Sessions', 'SCOOT Leader(s)', 'Verification'];
  rpt.getRange(4, 1, 1, 8).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');

  var rData = [];
  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    rData.push([r.invDate, r.invSchool, '$' + r.invAmount.toFixed(2), r.invNum, r.invStatus, r.sessions, r.leaders, r.matchStatus]);
  }

  if (rData.length > 0) {
    rpt.getRange(5, 1, rData.length, 8).setValues(rData);
    // Highlight unmatched rows red
    for (var i = 0; i < results.length; i++) {
      if (results[i].sessions === 0) {
        rpt.getRange(5 + i, 1, 1, 8).setBackground('#fce8e6').setFontColor('#d93025');
      } else {
        rpt.getRange(5 + i, 1, 1, 8).setBackground('#e6f4ea');
      }
    }
  }

  // Section 2: Unbilled sessions (in our records but no invoice)
  var ubStart = 5 + rData.length + 2;
  rpt.getRange(ubStart, 1).setValue('SCOOT SESSIONS WITH NO MATCHING INVOICE (' + unbilledSessions.length + ')');
  rpt.getRange(ubStart, 1, 1, 5).merge().setFontSize(12).setFontWeight('bold').setBackground('#fbbc04');

  if (unbilledSessions.length > 0) {
    rpt.getRange(ubStart + 1, 1, 1, 5).setValues([['Date', 'School', 'Leader', 'Workshop', 'Note']]).setFontWeight('bold').setBackground('#e8eaed');
    var ubData = [];
    for (var i = 0; i < unbilledSessions.length; i++) {
      var s = unbilledSessions[i];
      ubData.push([s.dateStr, s.school, s.leader, s.workshop, 'No invoice found']);
    }
    rpt.getRange(ubStart + 2, 1, ubData.length, 5).setValues(ubData);
  } else {
    rpt.getRange(ubStart + 1, 1).setValue('All SCOOT sessions have matching invoices.');
    rpt.getRange(ubStart + 1, 1, 1, 5).merge().setFontColor('#666666');
  }

  for (var c = 1; c <= 8; c++) rpt.autoResizeColumn(c);
  rpt.setFrozenRows(4);

  SpreadsheetApp.getUi().alert('SCOOT Invoice Verification Complete!\n\n' +
    'Invoices checked: ' + invoices.length + '\n' +
    'Matched: ' + matched + '\n' +
    'No match (potential overbilling): ' + unmatched + '\n' +
    'Unbilled sessions: ' + unbilledSessions.length + '\n\n' +
    '→ Check "SCOOT Invoice Verify" tab');
}

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
      .addItem('Run 02/05 - 02/18', 'run02_05to02_18')
      .addItem('Run Custom Date Range...', 'promptDateRange')
      .addSeparator()
      .addItem('Setup Gusto Import Tab', 'setupGustoImportTab')
      .addItem('Compare Gusto Hours', 'compareGustoHours')
      .addSeparator()
      .addItem('Generate Hours Tracker 02/05 - 02/18', 'generateHoursTracker02_05to02_18')
      .addItem('Generate Hours Tracker (Custom Range)...', 'promptHoursTrackerDateRange')
      .addItem('🔍 Debug SCOOT Detection', 'debugScootDetection')
      .addSeparator()
      .addItem('Send Daily Slack Alert', 'sendDailySlackAlert')
      .addItem('Send Weekly Slack Summary', 'sendWeeklySlackSummary')
      .addSeparator()
      .addItem('Setup Auto Triggers (Daily + Weekly)', 'setupTriggers')
      .addItem('Remove All Triggers', 'removeTriggers')
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

  // STEP 3: Split into scoot vs non-scoot, calculate hours
  var leaderData = [];  // non-scoot sessions
  var scootData = [];   // scoot sessions

  for (var i = 0; i < allCheckIns.length; i++) {
    var rec = allCheckIns[i];
    if (!rec.worked) continue;

    var isScoot = rec.status.indexOf('scoot') >= 0;
    var ops = matchOpsHub(rec.school, rec.workshop, opsHub);
    var dur, allowed, src;
    if (ops) {
      dur = ops.dur;
      allowed = ops.allowed;
      src = 'Ops Hub (' + ops.site + ')';
    } else {
      dur = 60;
      allowed = 90;
      src = 'DEFAULT 1hr';
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
      unmatched: !ops
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

  // STEP 4: Write both tabs
  var hvSS = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var dateRange = formatDt(startDate) + ' – ' + formatDt(endDate);

  writeHoursTrackerTab_(hvSS, leaderData, HOURS_TRACKER_TAB, false, dateRange);
  SpreadsheetApp.flush();  // Force pending writes before starting second tab
  writeHoursTrackerTab_(hvSS, scootData, SCOOT_HOURS_TAB, true, dateRange);

  log.push('\n=== COMPLETE ===');
  Logger.log(log.join('\n'));

  try {
    var leadersByRegion = groupByRegion_(leaderData);
    var regionCount = Object.keys(leadersByRegion).length;
    SpreadsheetApp.getUi().alert('Hours Tracker Complete!',
      'Leader sessions: ' + leaderData.length +
      '\nSCOOT sessions: ' + scootData.length +
      '\nRegions: ' + regionCount +
      '\n\n→ Check "' + HOURS_TRACKER_TAB + '" tab for all leaders' +
      '\n→ Check "' + SCOOT_HOURS_TAB + '" tab for SCOOT invoices',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {
    Logger.log('Could not show alert: ' + e.message);
  }
}

/**
 * Groups session entries by region and aggregates per-leader summaries.
 * @param {Array} data - Array of session entry objects
 * @return {Object} Map of region → { leaders: { name → { sessions, totalMin, unmatched, details[] } } }
 */
function groupByRegion_(data) {
  var regions = {};

  for (var i = 0; i < data.length; i++) {
    var entry = data[i];
    var region = entry.region || 'Unknown';

    if (!regions[region]) {
      regions[region] = { leaders: {} };
    }

    var leaders = regions[region].leaders;
    if (!leaders[entry.leader]) {
      leaders[entry.leader] = { name: entry.leader, sessions: 0, totalMin: 0, unmatched: 0, details: [] };
    }

    var ldr = leaders[entry.leader];
    ldr.sessions++;
    ldr.totalMin += entry.allowed;
    if (entry.unmatched) ldr.unmatched++;
    ldr.details.push(entry);
  }

  return regions;
}

/**
 * Writes a formatted Hours Tracker tab with two sections:
 *   Section A: Summary by region (one row per leader)
 *   Section B: Detailed session log by region (one row per session)
 *
 * @param {Spreadsheet} hvSS - Hours Verification spreadsheet
 * @param {Array} data - Session entries (already filtered to leaders-only or scoot-only)
 * @param {string} tabName - Tab name to create/overwrite
 * @param {boolean} isScoot - True if this is the SCOOT tab
 * @param {string} dateRange - Formatted date range string for the title
 */
function writeHoursTrackerTab_(hvSS, data, tabName, isScoot, dateRange) {
  var tab = hvSS.getSheetByName(tabName);
  if (!tab) {
    tab = hvSS.insertSheet(tabName);
  } else {
    tab.clear();
    tab.clearFormats();
    SpreadsheetApp.flush();  // Ensure clear completes before writing
  }

  if (data.length === 0) {
    tab.getRange(1, 1).setValue('No ' + (isScoot ? 'SCOOT' : 'leader') + ' sessions found for ' + dateRange);
    tab.getRange(1, 1).setFontSize(12).setFontWeight('bold');
    return;
  }

  // === SCOOT: simple invoice-verification layout ===
  if (isScoot) {
    writeScootInvoiceTab_(tab, data, dateRange);
    return;
  }

  var regions = groupByRegion_(data);
  var regionNames = Object.keys(regions).sort();

  // =============== SECTION A: SUMMARY ===============
  var titleLabel = 'HOURS TRACKER';
  tab.getRange(1, 1).setValue(titleLabel + ' — ' + dateRange);
  tab.getRange(1, 1, 1, 8).merge();
  tab.getRange(1, 1).setFontSize(14).setFontWeight('bold')
    .setBackground(isScoot ? '#ff6d01' : '#1a73e8').setFontColor('#ffffff');

  tab.getRange(2, 1).setValue('Generated: ' + new Date().toLocaleString() + ' | Regions: ' + regionNames.length + ' | Sessions: ' + data.length);
  tab.getRange(2, 1, 1, 8).merge().setFontColor('#666666');

  // Summary headers
  var sumHeaders = ['Region', 'Leader Name', 'Sessions', 'Total Hours', 'Formatted', 'Unmatched', 'Status'];
  tab.getRange(4, 1, 1, sumHeaders.length).setValues([sumHeaders]).setFontWeight('bold').setBackground('#e8eaed');

  var row = 5;
  for (var r = 0; r < regionNames.length; r++) {
    var regionName = regionNames[r];
    var colorIdx = r % REGION_TINTS.length;
    var tint = REGION_TINTS[colorIdx];
    var leaderMap = regions[regionName].leaders;

    // Sort leaders alphabetically within region
    var leaderNames = Object.keys(leaderMap).sort();

    for (var l = 0; l < leaderNames.length; l++) {
      var ldr = leaderMap[leaderNames[l]];
      var hrs = Math.round((ldr.totalMin / 60) * 100) / 100;
      var hh = Math.floor(hrs);
      var mm = Math.round((hrs - hh) * 60);
      var fmt = hh + 'h ' + (mm < 10 ? '0' : '') + mm + 'm';

      tab.getRange(row, 1, 1, 7).setValues([[
        regionName,
        ldr.name,
        ldr.sessions,
        hrs,
        fmt,
        ldr.unmatched,
        ldr.unmatched > 0 ? '⚠️ Check' : '✅ OK'
      ]]);
      tab.getRange(row, 1, 1, 7).setBackground(tint);
      row++;
    }
  }

  // =============== SECTION B: DETAILED SESSION LOG ===============
  var detailStart = row + 2;
  tab.getRange(detailStart, 1).setValue('DETAILED SESSION LOG');
  tab.getRange(detailStart, 1, 1, 8).merge();
  tab.getRange(detailStart, 1).setFontSize(12).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');

  var detHeaders = ['Leader Name', 'Date', 'Workshop', 'School', 'Status', 'Duration (min)', 'Allowed (min)', 'Source'];
  tab.getRange(detailStart + 1, 1, 1, detHeaders.length).setValues([detHeaders]).setFontWeight('bold').setBackground('#e8eaed');

  var dRow = detailStart + 2;
  for (var r = 0; r < regionNames.length; r++) {
    var regionName = regionNames[r];
    var colorIdx = r % REGION_COLORS.length;
    var regionColor = REGION_COLORS[colorIdx];
    var tint = REGION_TINTS[colorIdx];
    var leaderMap = regions[regionName].leaders;
    var leaderNames = Object.keys(leaderMap).sort();

    // Region header row
    tab.getRange(dRow, 1).setValue(regionName);
    tab.getRange(dRow, 1, 1, 8).merge();
    tab.getRange(dRow, 1).setFontSize(11).setFontWeight('bold')
      .setBackground(regionColor).setFontColor('#ffffff');
    dRow++;

    for (var l = 0; l < leaderNames.length; l++) {
      var ldr = leaderMap[leaderNames[l]];

      // Sort details by date
      ldr.details.sort(function(a, b) {
        if (!a.date && !b.date) return 0;
        if (!a.date) return 1;
        if (!b.date) return -1;
        return a.date.getTime() - b.date.getTime();
      });

      // Write each session row
      for (var d = 0; d < ldr.details.length; d++) {
        var det = ldr.details[d];
        tab.getRange(dRow, 1, 1, 8).setValues([[
          det.leader, det.dateStr, det.workshop, det.school,
          det.status, det.dur, det.allowed, det.src
        ]]);
        // Highlight unmatched workshops yellow
        if (det.unmatched) {
          tab.getRange(dRow, 1, 1, 8).setBackground('#fff3cd');
        } else {
          tab.getRange(dRow, 1, 1, 8).setBackground(tint);
        }
        dRow++;
      }

      // Leader subtotal row
      var totalHrs = Math.round((ldr.totalMin / 60) * 100) / 100;
      var hh = Math.floor(totalHrs);
      var mm = Math.round((totalHrs - hh) * 60);
      var fmtTotal = hh + 'h ' + (mm < 10 ? '0' : '') + mm + 'm';

      tab.getRange(dRow, 1, 1, 8).setValues([[
        '', '', '', ldr.name + ' TOTAL', ldr.sessions + ' sessions', '', ldr.totalMin, fmtTotal
      ]]);
      tab.getRange(dRow, 1, 1, 8).setFontWeight('bold').setBackground('#e8eaed');
      dRow++;
    }
  }

  // Auto-resize columns and freeze header rows
  for (var c = 1; c <= 8; c++) tab.autoResizeColumn(c);
  tab.setFrozenRows(4);
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

  // Write rows
  var row = 5;
  var currentPerson = '';
  var personColor = 0;
  for (var i = 0; i < data.length; i++) {
    var d = data[i];
    // Alternate tint per person for readability
    if (d.leader !== currentPerson) {
      currentPerson = d.leader;
      personColor = (personColor + 1) % REGION_TINTS.length;
    }
    tab.getRange(row, 1, 1, 4).setValues([[d.leader, d.school, d.dateStr, d.workshop]]);
    tab.getRange(row, 1, 1, 4).setBackground(REGION_TINTS[personColor]);
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

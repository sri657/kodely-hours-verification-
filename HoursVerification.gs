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

// --- Phase 4: Gusto API Configuration ---
const GUSTO_API_BASE = 'https://api.gusto.com';
const GUSTO_SANDBOX_BASE = 'https://api.gusto-demo.com';
const GUSTO_USE_SANDBOX = true; // Set to false for production
const GUSTO_API_VERSION = '2024-03-01';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Hours Verification')
    .addItem('Run 02/05 - 02/18', 'run02_05to02_18')
    .addItem('Run Custom Date Range...', 'promptDateRange')
    .addSeparator()
    .addItem('Setup Gusto Import Tab', 'setupGustoImportTab')
    .addItem('Compare Gusto Hours', 'compareGustoHours')
    .addSeparator()
    .addItem('Authorize Gusto', 'authorizeGusto')
    .addItem('Sync Hours from Gusto', 'syncGustoHours')
    .addItem('Gusto Connection Status', 'checkGustoStatus')
    .addItem('Disconnect Gusto', 'disconnectGusto')
    .addSeparator()
    .addItem('Send Daily Slack Alert', 'sendDailySlackAlert')
    .addItem('Send Weekly Slack Summary', 'sendWeeklySlackSummary')
    .addSeparator()
    .addItem('Setup Auto Triggers (Daily + Weekly)', 'setupTriggers')
    .addItem('Remove All Triggers', 'removeTriggers')
    .addToUi();
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
// PHASE 4: GUSTO API INTEGRATION
// =====================================================

/**
 * Creates and configures the OAuth2 service for Gusto API.
 * Requires the OAuth2 library: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 * Credentials are read from Script Properties (GUSTO_CLIENT_ID, GUSTO_CLIENT_SECRET).
 */
function getGustoService_() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('GUSTO_CLIENT_ID');
  var clientSecret = props.getProperty('GUSTO_CLIENT_SECRET');

  if (!clientId || !clientSecret) {
    throw new Error('Gusto credentials not configured. Set GUSTO_CLIENT_ID and GUSTO_CLIENT_SECRET in Script Properties.');
  }

  var baseUrl = GUSTO_USE_SANDBOX ? GUSTO_SANDBOX_BASE : GUSTO_API_BASE;

  return OAuth2.createService('gusto')
    .setAuthorizationBaseUrl(baseUrl + '/oauth/authorize')
    .setTokenUrl(baseUrl + '/oauth/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('gustoAuthCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('employees:read time_sheet:read')
    .setParam('redirect_uri', getRedirectUri_());
}

/**
 * Returns the OAuth2 redirect URI for this script's web app deployment.
 */
function getRedirectUri_() {
  return 'https://script.google.com/macros/d/' + ScriptApp.getScriptId() + '/usercallback';
}

/**
 * Opens a modal dialog with the Gusto OAuth authorization link.
 * User clicks the link to authorize in a new tab.
 */
function authorizeGusto() {
  var service = getGustoService_();

  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('Already connected to Gusto!\n\nUse "Gusto Connection Status" to check details, or "Disconnect Gusto" to re-authorize.');
    return;
  }

  var authUrl = service.getAuthorizationUrl();
  var html = HtmlService.createHtmlOutput(
    '<p>Click the link below to authorize Gusto access:</p>' +
    '<p><a href="' + authUrl + '" target="_blank" style="font-size:16px;font-weight:bold;">Authorize Gusto</a></p>' +
    '<p style="color:#666;font-size:12px;">After authorizing, you can close this dialog.</p>'
  )
    .setWidth(400)
    .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, 'Authorize Gusto');
}

/**
 * Handles the OAuth2 callback after the user authorizes in Gusto.
 * This function is called automatically by the OAuth2 library.
 */
function gustoAuthCallback(request) {
  var service = getGustoService_();
  var authorized = service.handleCallback(request);

  if (authorized) {
    return HtmlService.createHtmlOutput(
      '<h3 style="color:#34a853;">✓ Gusto Connected Successfully!</h3>' +
      '<p>You can close this tab and return to your spreadsheet.</p>' +
      '<p>Use <b>Sync Hours from Gusto</b> from the Hours Verification menu to pull hours.</p>'
    );
  } else {
    return HtmlService.createHtmlOutput(
      '<h3 style="color:#d93025;">✗ Authorization Failed</h3>' +
      '<p>Please try again from the Hours Verification menu → Authorize Gusto.</p>'
    );
  }
}

/**
 * Clears stored OAuth tokens, disconnecting from Gusto.
 */
function disconnectGusto() {
  var service = getGustoService_();
  service.reset();
  SpreadsheetApp.getUi().alert('Gusto disconnected.\n\nOAuth tokens have been cleared. You will need to re-authorize to sync hours.');
}

/**
 * Displays the current Gusto connection status — credentials, auth state, and mode.
 */
function checkGustoStatus() {
  var props = PropertiesService.getScriptProperties();
  var clientId = props.getProperty('GUSTO_CLIENT_ID');
  var clientSecret = props.getProperty('GUSTO_CLIENT_SECRET');
  var companyUuid = props.getProperty('GUSTO_COMPANY_UUID');

  var lines = [];
  lines.push('=== Gusto Connection Status ===\n');
  lines.push('Mode: ' + (GUSTO_USE_SANDBOX ? 'SANDBOX (demo)' : 'PRODUCTION'));
  lines.push('API Base: ' + (GUSTO_USE_SANDBOX ? GUSTO_SANDBOX_BASE : GUSTO_API_BASE));
  lines.push('API Version: ' + GUSTO_API_VERSION);
  lines.push('');
  lines.push('Client ID: ' + (clientId ? clientId.substring(0, 8) + '...' : '❌ NOT SET'));
  lines.push('Client Secret: ' + (clientSecret ? '••••••••' : '❌ NOT SET'));
  lines.push('Company UUID: ' + (companyUuid ? companyUuid : '❌ NOT SET'));

  if (clientId && clientSecret) {
    try {
      var service = getGustoService_();
      lines.push('');
      lines.push('Auth Status: ' + (service.hasAccess() ? '✅ Connected' : '❌ Not authorized — run "Authorize Gusto"'));
    } catch (err) {
      lines.push('');
      lines.push('Auth Status: ❌ Error — ' + err.message);
    }
  } else {
    lines.push('');
    lines.push('⚠️ Set GUSTO_CLIENT_ID and GUSTO_CLIENT_SECRET in Script Properties first.');
  }

  lines.push('');
  lines.push('Redirect URI (for Gusto app config):');
  lines.push(getRedirectUri_());

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

/**
 * Authenticated GET request to the Gusto API.
 * Handles Bearer token injection, 401 (re-auth needed), and 429 (rate limit).
 * @param {string} endpoint - API path (e.g. '/v1/companies/{uuid}/employees')
 * @param {Object} params - Optional query parameters
 * @return {Object} Parsed JSON response
 */
function gustoApiFetch_(endpoint, params) {
  var service = getGustoService_();
  if (!service.hasAccess()) {
    throw new Error('Not authorized with Gusto. Run "Authorize Gusto" from the menu first.');
  }

  var baseUrl = GUSTO_USE_SANDBOX ? GUSTO_SANDBOX_BASE : GUSTO_API_BASE;
  var url = baseUrl + endpoint;

  // Append query params
  if (params && Object.keys(params).length > 0) {
    var qs = [];
    for (var key in params) {
      qs.push(encodeURIComponent(key) + '=' + encodeURIComponent(params[key]));
    }
    url += '?' + qs.join('&');
  }

  var options = {
    method: 'get',
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'X-Gusto-API-Version': GUSTO_API_VERSION,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();

  if (code === 401) {
    service.reset();
    throw new Error('Gusto authorization expired. Please run "Authorize Gusto" again.');
  }

  if (code === 429) {
    // Rate limited — wait and retry once
    Utilities.sleep(2000);
    response = UrlFetchApp.fetch(url, options);
    code = response.getResponseCode();
    if (code === 429) {
      throw new Error('Gusto API rate limit exceeded. Please wait a few minutes and try again.');
    }
  }

  if (code !== 200) {
    throw new Error('Gusto API error (' + code + '): ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * Fetches all pages from a paginated Gusto API endpoint.
 * Uses 100 items per page with a safety cap of 50 pages (5000 items).
 * @param {string} endpoint - API path
 * @param {Object} params - Optional base query parameters
 * @return {Array} All items across all pages
 */
function gustoFetchAllPages_(endpoint, params) {
  var allItems = [];
  var page = 1;
  var maxPages = 50;
  params = params || {};
  params.per = 100;

  while (page <= maxPages) {
    params.page = page;
    var data = gustoApiFetch_(endpoint, params);

    // Handle both array responses and object responses with a data key
    var items = Array.isArray(data) ? data : (data.employees || data.time_sheets || data.data || []);

    if (items.length === 0) break;

    allItems = allItems.concat(items);

    if (items.length < 100) break; // Last page
    page++;
  }

  return allItems;
}

/**
 * Fetches all employees from the Gusto API and returns a UUID-to-name map.
 * @return {Object} Map of employee UUID → full name string
 */
function fetchGustoEmployees_() {
  var companyUuid = PropertiesService.getScriptProperties().getProperty('GUSTO_COMPANY_UUID');
  if (!companyUuid) {
    throw new Error('GUSTO_COMPANY_UUID not set in Script Properties.');
  }

  var employees = gustoFetchAllPages_('/v1/companies/' + companyUuid + '/employees');

  var nameMap = {};
  for (var i = 0; i < employees.length; i++) {
    var emp = employees[i];
    var uuid = emp.uuid || emp.id;
    var firstName = emp.first_name || emp.firstName || '';
    var lastName = emp.last_name || emp.lastName || '';
    var fullName = (firstName + ' ' + lastName).trim();
    if (uuid && fullName) {
      nameMap[uuid] = fullName;
    }
  }

  return nameMap;
}

/**
 * Fetches time sheets from Gusto for a date range, resolves employee UUIDs to names,
 * and aggregates total hours per employee.
 * @param {Date} start - Start date
 * @param {Date} end - End date
 * @return {Array} Array of {name, hours} objects sorted by name
 */
function fetchGustoTimeSheets_(start, end) {
  var companyUuid = PropertiesService.getScriptProperties().getProperty('GUSTO_COMPANY_UUID');
  if (!companyUuid) {
    throw new Error('GUSTO_COMPANY_UUID not set in Script Properties.');
  }

  // Fetch employee name map
  var nameMap = fetchGustoEmployees_();

  // Format dates as YYYY-MM-DD for API
  var startStr = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var endStr = Utilities.formatDate(end, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Fetch time sheets with date params
  var timeSheets = gustoFetchAllPages_(
    '/v1/companies/' + companyUuid + '/time_sheets',
    { start_date: startStr, end_date: endStr }
  );

  // Aggregate hours per employee
  var hoursByEmployee = {};

  for (var i = 0; i < timeSheets.length; i++) {
    var sheet = timeSheets[i];
    var empUuid = sheet.employee_uuid || sheet.employee_id || '';
    var empName = nameMap[empUuid] || empUuid;

    // Parse hours from the time sheet entry
    var totalHours = 0;

    // Try different field names the API might use
    if (typeof sheet.total_hours === 'number') {
      totalHours = sheet.total_hours;
    } else if (typeof sheet.regular_hours === 'number') {
      totalHours = (sheet.regular_hours || 0) + (sheet.overtime_hours || 0) + (sheet.double_overtime_hours || 0);
    } else if (sheet.hours_by_type) {
      // Some API versions nest hours
      var hbt = sheet.hours_by_type;
      totalHours = (hbt.regular || 0) + (hbt.overtime || 0) + (hbt.double_overtime || 0);
    } else if (typeof sheet.total_regular_hours_worked === 'number') {
      totalHours = sheet.total_regular_hours_worked + (sheet.total_overtime_hours_worked || 0);
    }

    // Client-side date filtering as fallback
    if (sheet.date || sheet.check_date) {
      var entryDate = new Date(sheet.date || sheet.check_date);
      if (entryDate < start || entryDate > end) continue;
    }

    if (!empName || totalHours === 0) continue;

    if (!hoursByEmployee[empName]) {
      hoursByEmployee[empName] = 0;
    }
    hoursByEmployee[empName] += totalHours;
  }

  // Convert to sorted array
  var result = [];
  for (var name in hoursByEmployee) {
    result.push({
      name: name,
      hours: Math.round(hoursByEmployee[name] * 100) / 100
    });
  }
  result.sort(function(a, b) { return a.name.localeCompare(b.name); });

  return result;
}

/**
 * Writes API-fetched data to the "Gusto Import" tab in the same format as a CSV paste.
 * Column A = Employee Name, Column B = Total Hours, headers at row 4.
 * @param {Array} data - Array of {name, hours} objects
 */
function writeGustoImportData_(data) {
  var ss = SpreadsheetApp.openById(HOURS_VERIFICATION_ID);
  var tab = ss.getSheetByName(GUSTO_IMPORT_TAB);
  if (!tab) {
    tab = ss.insertSheet(GUSTO_IMPORT_TAB);
  } else {
    tab.clear();
    tab.clearFormats();
  }

  // Title row
  tab.getRange(1, 1).setValue('GUSTO HOURS IMPORT (via API)');
  tab.getRange(1, 1, 1, 6).merge();
  tab.getRange(1, 1).setFontSize(14).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  // Info row
  tab.getRange(2, 1).setValue('Synced from Gusto API on ' + new Date().toLocaleString() + ' | ' + data.length + ' employees');
  tab.getRange(2, 1, 1, 6).merge().setFontColor('#666666').setWrap(true);

  // Headers at row 4 (same as setupGustoImportTab)
  var headers = ['Employee Name', 'Total Hours', 'Regular Hours', 'Overtime Hours', 'Department', 'Notes'];
  tab.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eaed');

  // Write data starting at row 5
  if (data.length > 0) {
    var rows = [];
    for (var i = 0; i < data.length; i++) {
      rows.push([data[i].name, data[i].hours, '', '', '', '']);
    }
    tab.getRange(5, 1, rows.length, 6).setValues(rows);
  }

  tab.setFrozenRows(4);
  for (var c = 1; c <= 6; c++) tab.autoResizeColumn(c);
}

/**
 * Main user-facing function — prompts for date range, fetches hours from Gusto API,
 * writes to the Gusto Import tab, and optionally auto-runs the comparison.
 */
function syncGustoHours() {
  var ui = SpreadsheetApp.getUi();

  // Check authorization
  var service;
  try {
    service = getGustoService_();
  } catch (err) {
    ui.alert('Gusto Not Configured\n\n' + err.message);
    return;
  }

  if (!service.hasAccess()) {
    ui.alert('Not authorized with Gusto.\n\nRun "Authorize Gusto" from the menu first.');
    return;
  }

  // Prompt for start date
  var startResp = ui.prompt('Sync Hours from Gusto', 'Start date (MM/DD/YYYY):', ui.ButtonSet.OK_CANCEL);
  if (startResp.getSelectedButton() !== ui.Button.OK) return;

  var endResp = ui.prompt('Sync Hours from Gusto', 'End date (MM/DD/YYYY):', ui.ButtonSet.OK_CANCEL);
  if (endResp.getSelectedButton() !== ui.Button.OK) return;

  var startDate = new Date(startResp.getResponseText());
  var endDate = new Date(endResp.getResponseText());
  endDate.setHours(23, 59, 59);

  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    ui.alert('Invalid date format. Please use MM/DD/YYYY.');
    return;
  }

  // Fetch from Gusto
  try {
    ui.alert('Fetching hours from Gusto...\n\n' +
      'Period: ' + formatDt(startDate) + ' to ' + formatDt(endDate) + '\n\n' +
      'Click OK to start. This may take a moment.');

    var data = fetchGustoTimeSheets_(startDate, endDate);

    if (data.length === 0) {
      ui.alert('No time sheet data found for ' + formatDt(startDate) + ' to ' + formatDt(endDate) + '.\n\n' +
        'Check that:\n• The date range is correct\n• Employees have logged hours in Gusto\n• Your Gusto app has time_sheet:read scope');
      return;
    }

    // Write to Gusto Import tab
    writeGustoImportData_(data);

    // Ask if they want to auto-run compare
    var compare = ui.alert('Gusto Sync Complete!',
      data.length + ' employees synced to "' + GUSTO_IMPORT_TAB + '" tab.\n\n' +
      'Run "Compare Gusto Hours" now to check for discrepancies?',
      ui.ButtonSet.YES_NO);

    if (compare === ui.Button.YES) {
      compareGustoHours();
    }
  } catch (err) {
    ui.alert('Gusto Sync Error\n\n' + err.message);
    Logger.log('Gusto sync error: ' + err.message + '\n' + err.stack);
  }
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

  Logger.log('Triggers created: daily at 8 PM + weekly Friday at 6 PM');
  try {
    SpreadsheetApp.getUi().alert('Triggers set up!\n\n• Daily alert: Every day at 8 PM\n• Weekly summary: Every Friday at 6 PM\n\nMake sure your Slack webhook URL is configured in the script.');
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
    if (funcName === 'sendDailySlackAlert' || funcName === 'sendWeeklySlackSummary') {
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

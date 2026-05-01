// ════════════════════════════════════════════════════════════════
// NAAD QUESTION BANK CMS — Google Apps Script
// ════════════════════════════════════════════════════════════════
// This script manages the Question Bank for NAAD Safety & Academic editions.
// GNFS officers, security agencies and teachers submit questions via Google Sheet.
// The script serves approved questions as JSON to the NAAD platform.
//
// SHEET STRUCTURE:
// Sheet name: "Question Bank"
// Columns:
//   A: Module (e.g. "Fire Safety", "Police", "Mathematics")
//   B: Edition (Safety / Academic)
//   C: Level (JHS / SHS / Primary / All)
//   D: Question
//   E: Option A
//   F: Option B
//   G: Option C
//   H: Option D
//   I: Correct Answer (A / B / C / D)
//   J: Explanation
//   K: Difficulty (Easy / Medium / Hard)
//   L: Submitted By (agency/teacher name)
//   M: Status (Pending / Approved / Rejected)
//   N: Reviewed By
//   O: Date Submitted
//   P: Date Reviewed
//   Q: Notes
// ════════════════════════════════════════════════════════════════

const QB_SHEET = 'Question Bank';
const QB_STATS_SHEET = 'QB Stats';

// ── Module to quiz file mapping ──
const MODULE_MAP = {
  'General Safety':        'general',
  'Student Safety':        'studentSafety',
  'Fire Safety':           'fire',
  'Police':                'police',
  'Community Safety':      'police',
  'Ambulance':             'ambulance',
  'Health & Ambulance':    'ambulance',
  'Prisons':               'prisons',
  'Prisons Service':       'prisons',
  'NADMO':                 'nadmo',
  'Disaster Preparedness': 'nadmo',
  'Immigration':           'immigration',
  'Customs':               'customs',
  'Tourism':               'tourism',
  'Forest':                'forest',
  'Forest Conservation':   'forest',
  'GAF':                   'gaf',
  'National Defence':      'gaf',
  'Narcotics':             'narcotics',
  'Mathematics':           'mathematics',
  'English':               'english',
  'Science':               'science',
  'Social Studies':        'socialStudies',
  'Computing':             'computing',
  'ICT':                   'computing',
};

// ════════════════════════════════════════
//  doGet — Route all requests
// ════════════════════════════════════════
function doGet(e) {
  var p = e.parameter || {};

  if (p.action === 'getQuestions')  return getQuestions(p);
  if (p.action === 'getModules')    return getModules(p);
  if (p.action === 'getStats')      return getStats();
  if (p.action === 'approveAll')    return approveAllPending(p);

  // Default: return all approved questions
  return getQuestions({});
}

function doPost(e) {
  var p = JSON.parse(e.postData.contents || '{}');
  if (p.action === 'submitQuestion') return jsonResponse(submitQuestion(p));
  if (p.action === 'reviewQuestion') return jsonResponse(reviewQuestion(p));
  return jsonResponse({ status: 'error', message: 'Unknown action' });
}

// ════════════════════════════════════════
//  GET QUESTIONS
//  Returns approved questions as JSON
//  Params: module, edition, level, limit
// ════════════════════════════════════════
function getQuestions(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);
  if (!sheet) return jsonResponse({ status: 'error', message: 'Question Bank sheet not found' });

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows    = [];

  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[String(headers[j]).trim()] = String(data[i][j] || '').trim();
    }

    // Only return approved questions
    if (row['Status'] !== 'Approved') continue;

    // Filter by module if specified
    if (p.module && row['Module'].toLowerCase() !== p.module.toLowerCase()) continue;

    // Filter by edition if specified
    if (p.edition && row['Edition'].toLowerCase() !== p.edition.toLowerCase()) continue;

    // Filter by level if specified
    if (p.level && row['Level'] !== 'All' && row['Level'].toLowerCase() !== p.level.toLowerCase()) continue;

    // Map correct answer letter to index
    var answerMap = { 'A': 0, 'B': 1, 'C': 2, 'D': 3 };
    var correctIndex = answerMap[row['Correct Answer'].toUpperCase()] || 0;

    rows.push({
      module:      row['Module'],
      edition:     row['Edition'],
      level:       row['Level'],
      question:    row['Question'],
      answers:     [row['Option A'], row['Option B'], row['Option C'], row['Option D']],
      correct:     correctIndex,
      explain:     row['Explanation'],
      difficulty:  row['Difficulty'],
      source:      row['Submitted By'],
    });
  }

  // Shuffle and limit if requested
  if (p.shuffle === 'true') {
    rows = shuffleArray(rows);
  }
  if (p.limit && parseInt(p.limit) > 0) {
    rows = rows.slice(0, parseInt(p.limit));
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', count: rows.length, questions: rows }))
    .setMimeType(ContentService.MimeType.TEXT);
}

// ════════════════════════════════════════
//  GET MODULES
//  Returns list of available modules with question counts
// ════════════════════════════════════════
function getModules(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);
  if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found' });

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var modules = {};

  for (var i = 1; i < data.length; i++) {
    var module  = String(data[i][0] || '').trim();
    var edition = String(data[i][1] || '').trim();
    var status  = String(data[i][12] || '').trim();
    if (!module) continue;

    var key = module + '|' + edition;
    if (!modules[key]) modules[key] = { module: module, edition: edition, total: 0, approved: 0, pending: 0 };
    modules[key].total++;
    if (status === 'Approved') modules[key].approved++;
    if (status === 'Pending')  modules[key].pending++;
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', modules: Object.values(modules) }))
    .setMimeType(ContentService.MimeType.TEXT);
}

// ════════════════════════════════════════
//  GET STATS
//  Returns overview stats for admin
// ════════════════════════════════════════
function getStats() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);
  if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found' });

  var data  = sheet.getDataRange().getValues();
  var stats = { total: 0, approved: 0, pending: 0, rejected: 0, byModule: {}, byAgency: {} };

  for (var i = 1; i < data.length; i++) {
    var module  = String(data[i][0]  || '').trim();
    var status  = String(data[i][12] || '').trim();
    var agency  = String(data[i][11] || '').trim();
    if (!module) continue;

    stats.total++;
    if (status === 'Approved') stats.approved++;
    if (status === 'Pending')  stats.pending++;
    if (status === 'Rejected') stats.rejected++;
    if (!stats.byModule[module])  stats.byModule[module]  = 0;
    if (!stats.byAgency[agency])  stats.byAgency[agency]  = 0;
    stats.byModule[module]++;
    stats.byAgency[agency]++;
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', stats: stats }))
    .setMimeType(ContentService.MimeType.TEXT);
}

// ════════════════════════════════════════
//  SUBMIT QUESTION (via POST)
//  Called when agency/teacher submits from the CMS portal
// ════════════════════════════════════════
function submitQuestion(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);
  if (!sheet) return { status: 'error', message: 'Sheet not found' };

  // Validate required fields
  if (!p.module)         return { status: 'error', message: 'Module is required' };
  if (!p.question)       return { status: 'error', message: 'Question is required' };
  if (!p.optionA)        return { status: 'error', message: 'Option A is required' };
  if (!p.optionB)        return { status: 'error', message: 'Option B is required' };
  if (!p.optionC)        return { status: 'error', message: 'Option C is required' };
  if (!p.optionD)        return { status: 'error', message: 'Option D is required' };
  if (!p.correctAnswer)  return { status: 'error', message: 'Correct answer is required' };
  if (!p.explanation)    return { status: 'error', message: 'Explanation is required' };
  if (!p.submittedBy)    return { status: 'error', message: 'Submitted by is required' };

  // Check for duplicate question
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toLowerCase() === p.question.trim().toLowerCase()) {
      return { status: 'duplicate', message: 'This question already exists in the bank' };
    }
  }

  sheet.appendRow([
    p.module        || '',
    p.edition       || 'Safety',
    p.level         || 'All',
    p.question      || '',
    p.optionA       || '',
    p.optionB       || '',
    p.optionC       || '',
    p.optionD       || '',
    p.correctAnswer || '',
    p.explanation   || '',
    p.difficulty    || 'Medium',
    p.submittedBy   || '',
    'Pending',
    '',
    new Date().toISOString(),
    '',
    p.notes || '',
  ]);

  // Colour row yellow (pending)
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 17).setBackground('#fff9c4');

  // Send notification email to admin
  try {
    var adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (adminEmail) {
      MailApp.sendEmail({
        to: adminEmail,
        subject: 'NAAD Question Bank — New Question Submitted',
        body: 'A new question has been submitted to the NAAD Question Bank.\n\n' +
              'Module: ' + p.module + '\n' +
              'Submitted by: ' + p.submittedBy + '\n' +
              'Question: ' + p.question + '\n\n' +
              'Please review and approve/reject in the Question Bank sheet.'
      });
    }
  } catch(e) {}

  return { status: 'ok', message: 'Question submitted for review' };
}

// ════════════════════════════════════════
//  REVIEW QUESTION (approve/reject)
//  Called by admin from the CMS portal
// ════════════════════════════════════════
function reviewQuestion(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);
  if (!sheet) return { status: 'error', message: 'Sheet not found' };

  var rowIndex = parseInt(p.row);
  if (!rowIndex || rowIndex < 2) return { status: 'error', message: 'Invalid row' };

  var status = p.status; // 'Approved' or 'Rejected'
  if (status !== 'Approved' && status !== 'Rejected') {
    return { status: 'error', message: 'Status must be Approved or Rejected' };
  }

  sheet.getRange(rowIndex, 13).setValue(status);
  sheet.getRange(rowIndex, 14).setValue(p.reviewedBy || 'Admin');
  sheet.getRange(rowIndex, 16).setValue(new Date().toISOString());
  if (p.notes) sheet.getRange(rowIndex, 17).setValue(p.notes);

  // Colour: green = approved, red = rejected
  var colour = status === 'Approved' ? '#c8f7c5' : '#ffcdd2';
  sheet.getRange(rowIndex, 1, 1, 17).setBackground(colour);

  return { status: 'ok', message: 'Question ' + status.toLowerCase() };
}

// ════════════════════════════════════════
//  APPROVE ALL PENDING (bulk action)
// ════════════════════════════════════════
function approveAllPending(p) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);
  if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found' });

  var data  = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][12]).trim() === 'Pending') {
      sheet.getRange(i + 1, 13).setValue('Approved');
      sheet.getRange(i + 1, 14).setValue(p.reviewedBy || 'Admin');
      sheet.getRange(i + 1, 16).setValue(new Date().toISOString());
      sheet.getRange(i + 1, 1, 1, 17).setBackground('#c8f7c5');
      count++;
    }
  }

  return jsonResponse({ status: 'ok', message: count + ' questions approved', count: count });
}

// ════════════════════════════════════════
//  SETUP — Create Question Bank sheet
//  Run this ONCE to create the sheet
// ════════════════════════════════════════
function setupQuestionBank() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(QB_SHEET);
  }

  // Headers
  var headers = [
    'Module', 'Edition', 'Level', 'Question',
    'Option A', 'Option B', 'Option C', 'Option D',
    'Correct Answer', 'Explanation', 'Difficulty',
    'Submitted By', 'Status', 'Reviewed By',
    'Date Submitted', 'Date Reviewed', 'Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a3060')
    .setFontColor('#FFD700')
    .setFontSize(11);

  // Column widths
  var widths = [150, 80, 80, 350, 120, 120, 120, 120, 80, 300, 80, 120, 80, 100, 120, 120, 150];
  widths.forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  // Freeze header row
  sheet.setFrozenRows(1);

  // Data validation for Edition column (B)
  var editionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Safety', 'Academic'], true).build();
  sheet.getRange('B2:B1000').setDataValidation(editionRule);

  // Data validation for Level column (C)
  var levelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['All', 'Primary', 'JHS', 'SHS'], true).build();
  sheet.getRange('C2:C1000').setDataValidation(levelRule);

  // Data validation for Correct Answer column (I)
  var answerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['A', 'B', 'C', 'D'], true).build();
  sheet.getRange('I2:I1000').setDataValidation(answerRule);

  // Data validation for Difficulty column (K)
  var diffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Easy', 'Medium', 'Hard'], true).build();
  sheet.getRange('K2:K1000').setDataValidation(diffRule);

  // Data validation for Status column (M)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Rejected'], true).build();
  sheet.getRange('M2:M1000').setDataValidation(statusRule);

  // Module dropdown — all known modules
  var modules = [
    'General Safety', 'Student Safety', 'Fire Safety', 'Police',
    'Ambulance', 'Prisons', 'NADMO', 'Immigration', 'Customs',
    'Tourism', 'Forest', 'GAF', 'Narcotics',
    'Mathematics', 'English', 'Science', 'Social Studies', 'Computing',
    'Religious & Moral Education', 'Creative Arts', 'Ghanaian Language'
  ];
  var moduleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(modules, true).build();
  sheet.getRange('A2:A1000').setDataValidation(moduleRule);

  // Add sample row
  sheet.appendRow([
    'Fire Safety', 'Safety', 'All',
    'What is the primary purpose of a fire extinguisher?',
    'To alert others of fire',
    'To put out small fires before they spread',
    'To cool down a burning building',
    'To create smoke signals for rescue',
    'B',
    'A fire extinguisher is designed to put out small fires before they spread and become uncontrollable. It should be used only on small, contained fires.',
    'Easy',
    'GNFS', 'Approved', 'Admin',
    new Date().toISOString(), new Date().toISOString(),
    'Sample question'
  ]);

  // Colour sample row green
  sheet.getRange(2, 1, 1, 17).setBackground('#c8f7c5');

  SpreadsheetApp.getUi().alert('✅ Question Bank sheet created successfully!\n\nYou can now start adding questions.');
}

// ════════════════════════════════════════
//  ADD CUSTOM MENU
// ════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 NAAD Question Bank')
    .addItem('Setup Question Bank Sheet', 'setupQuestionBank')
    .addItem('Approve All Pending Questions', 'approveAllPendingFromMenu')
    .addItem('View Stats', 'showStats')
    .addSeparator()
    .addItem('Export Approved Questions (JSON)', 'exportApprovedJSON')
    .addToUi();
}

function approveAllPendingFromMenu() {
  approveAllPending({ reviewedBy: 'Admin (Bulk Approve)' });
  SpreadsheetApp.getUi().alert('All pending questions have been approved.');
}

function showStats() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QB_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('Question Bank sheet not found.'); return; }

  var data    = sheet.getDataRange().getValues();
  var total   = data.length - 1;
  var approved = 0, pending = 0, rejected = 0;
  var byModule = {};

  for (var i = 1; i < data.length; i++) {
    var s = String(data[i][12] || '').trim();
    var m = String(data[i][0] || '').trim();
    if (s === 'Approved') approved++;
    if (s === 'Pending')  pending++;
    if (s === 'Rejected') rejected++;
    if (m) { byModule[m] = (byModule[m] || 0) + 1; }
  }

  var msg = '📊 NAAD Question Bank Stats\n\n' +
    '✅ Approved: ' + approved + '\n' +
    '⏳ Pending:  ' + pending + '\n' +
    '❌ Rejected: ' + rejected + '\n' +
    '📝 Total:    ' + total + '\n\n' +
    '📚 By Module:\n';

  Object.keys(byModule).sort().forEach(function(m) {
    msg += '  ' + m + ': ' + byModule[m] + '\n';
  });

  SpreadsheetApp.getUi().alert(msg);
}

function exportApprovedJSON() {
  var result = JSON.parse(getQuestions({ shuffle: 'false' }).getContent());
  var json = JSON.stringify(result.questions, null, 2);
  var file = DriveApp.createFile('NAAD_Questions_' + new Date().toISOString().slice(0,10) + '.json', json, 'application/json');
  SpreadsheetApp.getUi().alert('✅ Exported ' + result.count + ' approved questions to Google Drive.\n\nFile: ' + file.getName());
}

// ── HELPER ──
function shuffleArray(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
  }
  return arr;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.TEXT);
}

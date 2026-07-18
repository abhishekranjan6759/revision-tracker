/**
 * Revision Tracker - Google Apps Script Backend
 * ==============================================================
 * Deploy this as a Web App in Google Apps Script (Extensions > Apps Script).
 *
 * Sheet Structure (Row 1 Headers):
 * Sheet1:          entry_id | subject | title | notes | links | current_stage | last_reviewed_at | next_review_at | intervals | created_at
 * Subjects:        subject
 * Captures:        capture_id | type | title | description | priority | status | links | created_at
 * DayReflections:  reflection_id | date | rating | summary | reflection | notice
 *
 * Setup:
 * 1. Create a Google Sheet with headers above in Row 1 of "Sheet1".
 * 2. Also create a sheet named "Subjects" with subject names in column A (starting row 2). Row 1 header: "subject".
 * 3. Create a sheet named "DayReflections" with headers: reflection_id | date | rating | summary | reflection | notice.
 * 4. Open Extensions > Apps Script, paste this code.
 * 5. Run initialSetup() once.
 * 6. Deploy > New Deployment > Web App (Execute as: Me, Access: Anyone).
 * 7. Copy the Web App URL into your index.html SCRIPT_URL constant.
 */

var ENTRIES_SHEET = 'Sheet1';
var SUBJECTS_SHEET = 'Subjects';
var CAPTURES_SHEET = 'Captures';
var REFLECTIONS_SHEET = 'DayReflections';
var scriptProp = PropertiesService.getScriptProperties();

function initialSetup() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', activeSpreadsheet.getId());
}

// Handle GET requests (read data)
function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    var action = e.parameter.action;

    if (action === 'getEntries') {
      return getEntries(doc, e);
    } else if (action === 'getSubjects') {
      return getSubjects(doc);
    } else if (action === 'getTodayRevisions') {
      return getTodayRevisions(doc);
    } else if (action === 'getStats') {
      return getStats(doc);
    } else if (action === 'getCaptures') {
      return getCaptures(doc);
    } else if (action === 'getDayReflection') {
      return getDayReflection(doc, e);
    } else if (action === 'getAllReflections') {
      return getAllReflections(doc);
    }

    return jsonResponse({ result: 'error', error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ result: 'error', error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// Handle POST requests (write data)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    
    // Parse data from various POST formats
    var data;
    if (e.parameter && e.parameter.payload) {
      // FormData with payload field
      data = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      return jsonResponse({ result: 'error', error: 'No data received' });
    }
    
    var action = data.action;

    if (action === 'addEntry') {
      return addEntry(doc, data);
    } else if (action === 'markDone') {
      return markDone(doc, data);
    } else if (action === 'addSubject') {
      return addSubject(doc, data);
    } else if (action === 'deleteSubject') {
      return deleteSubject(doc, data);
    } else if (action === 'deleteEntry') {
      return deleteEntry(doc, data);
    } else if (action === 'addCapture') {
      return addCapture(doc, data);
    } else if (action === 'updateCapture') {
      return updateCapture(doc, data);
    } else if (action === 'deleteCapture') {
      return deleteCapture(doc, data);
    } else if (action === 'saveDayReflection') {
      return saveDayReflection(doc, data);
    }

    return jsonResponse({ result: 'error', error: 'Unknown action' });
  } catch (err) {
    return jsonResponse({ result: 'error', error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ==================== ENTRY OPERATIONS ====================

function addEntry(doc, data) {
  var sheet = doc.getSheetByName(ENTRIES_SHEET);
  var now = new Date();
  var intervals = data.intervals || { stage_1: 1, stage_2: 7, stage_3: 30, stage_4: 90 };
  var nextReview = new Date(now);
  nextReview.setDate(nextReview.getDate() + intervals.stage_1);

  var row = [
    data.entry_id || Utilities.getUuid(),
    data.subject,
    data.title,
    data.notes || '',
    JSON.stringify(data.links || []),
    0, // current_stage
    now.toISOString(),
    nextReview.toISOString(),
    JSON.stringify(intervals),
    now.toISOString()
  ];

  sheet.appendRow(row);
  return jsonResponse({ result: 'success', entry_id: row[0] });
}

function markDone(doc, data) {
  var sheet = doc.getSheetByName(ENTRIES_SHEET);
  var allData = sheet.getDataRange().getValues();
  var entryId = data.entry_id;

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === entryId) {
      var currentStage = parseInt(allData[i][5]) || 0;
      var intervals = JSON.parse(allData[i][8]);
      var now = new Date();

      var newStage = Math.min(currentStage + 1, 4);
      var stageKey = 'stage_' + newStage;
      var daysToAdd = intervals[stageKey] || 90;

      var nextReview = new Date(now);
      nextReview.setDate(nextReview.getDate() + daysToAdd);

      // Update current_stage, last_reviewed_at, next_review_at
      sheet.getRange(i + 1, 6).setValue(newStage);
      sheet.getRange(i + 1, 7).setValue(now.toISOString());
      sheet.getRange(i + 1, 8).setValue(nextReview.toISOString());

      return jsonResponse({
        result: 'success',
        new_stage: newStage,
        next_review_at: nextReview.toISOString(),
        mastered: newStage >= 4
      });
    }
  }

  return jsonResponse({ result: 'error', error: 'Entry not found' });
}

function deleteEntry(doc, data) {
  var sheet = doc.getSheetByName(ENTRIES_SHEET);
  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.entry_id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ result: 'success' });
    }
  }

  return jsonResponse({ result: 'error', error: 'Entry not found' });
}

function getEntries(doc, e) {
  var sheet = doc.getSheetByName(ENTRIES_SHEET);
  var allData = sheet.getDataRange().getValues();
  var entries = [];

  for (var i = 1; i < allData.length; i++) {
    entries.push({
      entry_id: allData[i][0],
      subject: allData[i][1],
      title: allData[i][2],
      notes: allData[i][3],
      links: JSON.parse(allData[i][4] || '[]'),
      current_stage: parseInt(allData[i][5]) || 0,
      last_reviewed_at: allData[i][6],
      next_review_at: allData[i][7],
      intervals: JSON.parse(allData[i][8] || '{}'),
      created_at: allData[i][9]
    });
  }

  return jsonResponse({ result: 'success', entries: entries });
}

function getTodayRevisions(doc) {
  var sheet = doc.getSheetByName(ENTRIES_SHEET);
  var allData = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var entries = [];

  for (var i = 1; i < allData.length; i++) {
    var nextReview = new Date(allData[i][7]);
    nextReview.setHours(0, 0, 0, 0);

    if (nextReview <= today && parseInt(allData[i][5]) < 4) {
      entries.push({
        entry_id: allData[i][0],
        subject: allData[i][1],
        title: allData[i][2],
        notes: allData[i][3],
        links: JSON.parse(allData[i][4] || '[]'),
        current_stage: parseInt(allData[i][5]) || 0,
        last_reviewed_at: allData[i][6],
        next_review_at: allData[i][7],
        intervals: JSON.parse(allData[i][8] || '{}'),
        created_at: allData[i][9]
      });
    }
  }

  return jsonResponse({ result: 'success', entries: entries });
}

function getStats(doc) {
  var sheet = doc.getSheetByName(ENTRIES_SHEET);
  var allData = sheet.getDataRange().getValues();
  var stats = {};

  for (var i = 1; i < allData.length; i++) {
    var subject = allData[i][1];
    var stage = parseInt(allData[i][5]) || 0;

    if (!stats[subject]) {
      stats[subject] = { total: 0, mastered: 0, stages: [0, 0, 0, 0, 0] };
    }
    stats[subject].total++;
    stats[subject].stages[stage]++;
    if (stage >= 4) stats[subject].mastered++;
  }

  return jsonResponse({ result: 'success', stats: stats });
}

// ==================== SUBJECT OPERATIONS ====================

function getSubjects(doc) {
  var sheet = doc.getSheetByName(SUBJECTS_SHEET);
  if (!sheet) return jsonResponse({ result: 'success', subjects: [] });

  var allData = sheet.getDataRange().getValues();
  var subjects = [];

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0]) subjects.push(allData[i][0]);
  }

  return jsonResponse({ result: 'success', subjects: subjects });
}

function addSubject(doc, data) {
  var sheet = doc.getSheetByName(SUBJECTS_SHEET);
  if (!sheet) {
    sheet = doc.insertSheet(SUBJECTS_SHEET);
    sheet.getRange(1, 1).setValue('subject');
  }

  sheet.appendRow([data.subject]);
  return jsonResponse({ result: 'success' });
}

function deleteSubject(doc, data) {
  var sheet = doc.getSheetByName(SUBJECTS_SHEET);
  if (!sheet) return jsonResponse({ result: 'error', error: 'Subjects sheet not found' });

  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.subject) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ result: 'success' });
    }
  }

  return jsonResponse({ result: 'error', error: 'Subject not found' });
}

// ==================== CAPTURE OPERATIONS ====================

function addCapture(doc, data) {
  var sheet = doc.getSheetByName(CAPTURES_SHEET);
  if (!sheet) {
    sheet = doc.insertSheet(CAPTURES_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([['capture_id', 'type', 'title', 'description', 'priority', 'status', 'links', 'created_at']]);
  }

  var now = new Date();
  var row = [
    data.capture_id || Utilities.getUuid(),
    data.type || 'thought',
    data.title,
    data.description || '',
    data.priority || 'medium',
    data.status || 'open',
    JSON.stringify(data.links || []),
    now.toISOString()
  ];

  sheet.appendRow(row);
  return jsonResponse({ result: 'success', capture_id: row[0] });
}

function getCaptures(doc) {
  var sheet = doc.getSheetByName(CAPTURES_SHEET);
  if (!sheet) return jsonResponse({ result: 'success', captures: [] });

  var allData = sheet.getDataRange().getValues();
  var captures = [];

  for (var i = 1; i < allData.length; i++) {
    captures.push({
      capture_id: allData[i][0],
      type: allData[i][1],
      title: allData[i][2],
      description: allData[i][3],
      priority: allData[i][4],
      status: allData[i][5],
      links: JSON.parse(allData[i][6] || '[]'),
      created_at: allData[i][7]
    });
  }

  // Return newest first
  captures.reverse();
  return jsonResponse({ result: 'success', captures: captures });
}

function updateCapture(doc, data) {
  var sheet = doc.getSheetByName(CAPTURES_SHEET);
  if (!sheet) return jsonResponse({ result: 'error', error: 'Captures sheet not found' });

  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.capture_id) {
      if (data.status) sheet.getRange(i + 1, 6).setValue(data.status); // status is col F
      return jsonResponse({ result: 'success' });
    }
  }

  return jsonResponse({ result: 'error', error: 'Capture not found' });
}

function deleteCapture(doc, data) {
  var sheet = doc.getSheetByName(CAPTURES_SHEET);
  if (!sheet) return jsonResponse({ result: 'error', error: 'Captures sheet not found' });

  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.capture_id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ result: 'success' });
    }
  }

  return jsonResponse({ result: 'error', error: 'Capture not found' });
}

// ==================== DAY REFLECTION OPERATIONS ====================

function saveDayReflection(doc, data) {
  var sheet = doc.getSheetByName(REFLECTIONS_SHEET);
  if (!sheet) {
    sheet = doc.insertSheet(REFLECTIONS_SHEET);
    sheet.getRange(1, 1, 1, 6).setValues([['reflection_id', 'date', 'rating', 'summary', 'reflection', 'notice']]);
  }

  var dateStr = data.date; // e.g., "2026-07-14"

  // Check if today's reflection already exists — update it
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][1] === dateStr) {
      // Update existing row
      sheet.getRange(i + 1, 3).setValue(data.rating);
      sheet.getRange(i + 1, 4).setValue(data.summary || '');
      sheet.getRange(i + 1, 5).setValue(data.reflection || '');
      sheet.getRange(i + 1, 6).setValue(data.notice || '');
      return jsonResponse({ result: 'success', updated: true });
    }
  }

  // Otherwise add new row
  var row = [
    data.reflection_id || Utilities.getUuid(),
    dateStr,
    data.rating,
    data.summary || '',
    data.reflection || '',
    data.notice || ''
  ];

  sheet.appendRow(row);
  return jsonResponse({ result: 'success', updated: false });
}

function getDayReflection(doc, e) {
  var sheet = doc.getSheetByName(REFLECTIONS_SHEET);
  if (!sheet) return jsonResponse({ result: 'success', reflection: null });

  var dateStr = e.parameter.date; // e.g., "2026-07-14"
  var allData = sheet.getDataRange().getValues();

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][1] === dateStr) {
      return jsonResponse({
        result: 'success',
        reflection: {
          reflection_id: allData[i][0],
          date: allData[i][1],
          rating: parseFloat(allData[i][2]),
          summary: allData[i][3],
          reflection: allData[i][4] || '',
          notice: allData[i][5] || ''
        }
      });
    }
  }

  return jsonResponse({ result: 'success', reflection: null });
}

function getAllReflections(doc) {
  var sheet = doc.getSheetByName(REFLECTIONS_SHEET);
  if (!sheet) return jsonResponse({ result: 'success', reflections: [] });

  var allData = sheet.getDataRange().getValues();
  var reflections = [];

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][0]) {
      reflections.push({
        reflection_id: allData[i][0],
        date: allData[i][1],
        rating: parseFloat(allData[i][2]),
        summary: allData[i][3] || '',
        reflection: allData[i][4] || '',
        notice: allData[i][5] || ''
      });
    }
  }

  // Return newest first
  reflections.reverse();
  return jsonResponse({ result: 'success', reflections: reflections });
}

// ==================== HELPERS ====================

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

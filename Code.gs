// ============================================================
// RecruitSync - Gmail → Google Sheets Job Application Tracker
// v2: Adds a daily auto-trigger.
//
// First-time setup:
//   1. Run scanJobApplications() once manually to populate the sheet.
//   2. Run createDailyTrigger() once to enable automatic daily scans.
//   3. That's it — never touch it again.
//
// To stop auto-scanning: run removeDailyTrigger().
// ============================================================

// --------------- CONFIGURATION ---------------

var CONFIG = {
  SHEET_NAME: 'Job Applications',

  // How many days back to search on first run (set lower once the sheet is populated)
  DAYS_BACK: 180,

  // Gmail search query to find initial application confirmation emails
  // Casts a wide net — false positives are filtered by STATUS_KEYWORDS below
  CONFIRMATION_QUERY: [
    'subject:"thank you for applying"',
    'subject:"application received"',
    'subject:"application confirmation"',
    'subject:"we received your application"',
    'subject:"thanks for applying"',
    'subject:"your application to"',
    'subject:"your application for"',
    'subject:"application submitted"',
  ].join(' OR '),
};

// Column positions (1-indexed)
var COL = {
  COMPANY:      1,
  JOB_TITLE:    2,
  DATE_APPLIED: 3,
  STATUS:       4,
  THREAD_ID:    5,  // used for deduplication; hide this column if you like
  LAST_UPDATED: 6,
};

// Status detection — order matters (more specific first)
// Each entry: { status: string, keywords: string[] }
// The email subject + body are matched (case-insensitive)
var STATUS_RULES = [
  {
    status: 'Offer',
    keywords: [
      'offer of employment', 'pleased to offer', 'job offer',
      'we would like to offer', 'congratulations.*offer',
    ],
  },
  {
    status: 'Interview',
    keywords: [
      'interview', 'schedule a call', 'schedule time', 'meet with',
      'next steps', 'move forward', 'moving forward', 'next round',
      'phone screen', 'technical screen', 'coding challenge', 'assessment',
    ],
  },
  {
    status: 'Rejected',
    keywords: [
      'unfortunately', 'not moving forward', 'not be moving forward',
      'decided to move forward with other', 'we will not', 'won\'t be moving',
      'position has been filled', 'other candidates', 'pursue other applicants',
      'not selected', 'not a fit', 'not the right fit', 'no longer considering',
    ],
  },
  {
    status: 'Applied',
    keywords: [
      'thank you for applying', 'application received', 'application confirmation',
      'we received your application', 'thanks for applying', 'application submitted',
    ],
  },
];

// ---------------  MAIN ENTRY POINT  ---------------

/**
 * Run this function manually (or via trigger) to scan Gmail and update the sheet.
 */
function scanJobApplications() {
  var sheet = getOrCreateSheet();
  var existingData = sheet.getDataRange().getValues(); // includes header row

  Logger.log('Starting scan. Existing rows (incl. header): ' + existingData.length);

  // --- Step 1: Find new confirmation emails and add them ---
  var afterDate = getDateFilter();
  var query = '(' + CONFIG.CONFIRMATION_QUERY + ') after:' + afterDate;
  var threads = GmailApp.search(query);

  Logger.log('Found ' + threads.length + ' confirmation threads to process.');

  threads.forEach(function(thread) {
    processThread(thread, sheet, existingData);
    // Refresh data so later iterations see rows added this pass
    existingData = sheet.getDataRange().getValues();
  });

  // --- Step 2: Re-scan all tracked threads for status updates ---
  updateTrackedThreads(sheet, existingData);

  Logger.log('Scan complete.');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'RecruitSync scan complete! Check the "' + CONFIG.SHEET_NAME + '" sheet.',
    'Done',
    5
  );
}

// --------------- THREAD PROCESSING ---------------

function processThread(thread, sheet, existingData) {
  var threadId = thread.getId();

  // Skip if this thread is already in the sheet
  if (findRowIndexByThreadId(existingData, threadId) !== -1) {
    return;
  }

  // Use the first message for initial parsing
  var messages = thread.getMessages();
  var firstMsg = messages[0];
  var subject   = firstMsg.getSubject();
  var body      = firstMsg.getPlainBody();
  var from      = firstMsg.getFrom();
  var date      = firstMsg.getDate();

  var company  = parseCompanyName(subject, from);
  var jobTitle = parseJobTitle(subject, body);
  var status   = detectStatus(subject, body);

  // If we can't determine any meaningful status, skip (probably not a real app email)
  if (!status) return;

  appendRow(sheet, company, jobTitle, date, status, threadId);
  Logger.log('Added: ' + company + ' | ' + jobTitle + ' | ' + status);
}

/**
 * For every thread already in the sheet, fetch the latest message and
 * update status if it has progressed.
 */
function updateTrackedThreads(sheet, existingData) {
  // Skip header row (row index 0 in the array = row 1 in sheet)
  for (var i = 1; i < existingData.length; i++) {
    var row      = existingData[i];
    var threadId = row[COL.THREAD_ID - 1];
    if (!threadId) continue;

    var currentStatus = row[COL.STATUS - 1];
    // No point re-checking if we already have a terminal status
    if (currentStatus === 'Offer' || currentStatus === 'Rejected') continue;

    try {
      var thread   = GmailApp.getThreadById(threadId);
      if (!thread) continue;

      var messages = thread.getMessages();
      // Check all messages for the best (highest-priority) status
      var bestStatus = currentStatus;
      messages.forEach(function(msg) {
        var s = detectStatus(msg.getSubject(), msg.getPlainBody());
        if (s && isHigherPriority(s, bestStatus)) {
          bestStatus = s;
        }
      });

      if (bestStatus !== currentStatus) {
        var sheetRow = i + 1; // 1-indexed, offset by header
        sheet.getRange(sheetRow, COL.STATUS).setValue(bestStatus);
        sheet.getRange(sheetRow, COL.LAST_UPDATED).setValue(new Date());
        colorRow(sheet, sheetRow, bestStatus);
        Logger.log('Updated row ' + sheetRow + ': ' + currentStatus + ' → ' + bestStatus);
      }
    } catch (e) {
      Logger.log('Could not fetch thread ' + threadId + ': ' + e.message);
    }
  }
}

// --------------- PARSING HELPERS ---------------

/**
 * Attempt to extract the company name from the email subject or sender.
 * Falls back to the sender domain if no pattern matches.
 */
function parseCompanyName(subject, from) {
  // Patterns to try against the subject line (case-insensitive)
  var subjectPatterns = [
    /thank you for applying (?:to|at|with)\s+([^!,.(–\-]+)/i,
    /thanks for applying (?:to|at|with)\s+([^!,.(–\-]+)/i,
    /your application (?:to|at|with|for a role at)\s+([^!,.(–\-]+)/i,
    /application (?:received|confirmation) (?:-\s*)?(?:at|to|from)?\s*([^!,.(–\-]+)/i,
    /we received your application (?:to|at|with)?\s*([^!,.(–\-]+)/i,
    // "[Company] - Application Received"
    /^([^|–\-]+?)\s*[-–|]\s*(?:application|your application)/i,
    // "Application to [Company]"
    /application to\s+([^!,.(–\-]+)/i,
  ];

  for (var i = 0; i < subjectPatterns.length; i++) {
    var match = subject.match(subjectPatterns[i]);
    if (match && match[1]) {
      return titleCase(match[1].trim());
    }
  }

  // Fall back: extract from sender's email domain
  // e.g. "Careers <jobs@stripe.com>" → "Stripe"
  var domainMatch = from.match(/@([\w-]+)\./);
  if (domainMatch) {
    var domain = domainMatch[1];
    // Strip common generic prefixes
    domain = domain.replace(/^(mail|noreply|no-reply|careers|jobs|hr|recruiting|talent|apply|greenhouse|lever|workday|ashby|jobvite|icims|smartrecruiters)$/i, '');
    if (domain) return titleCase(domain);
  }

  return 'Unknown';
}

/**
 * Attempt to extract the job title from the subject or first ~500 chars of body.
 */
function parseJobTitle(subject, body) {
  var text = subject + ' ' + body.substring(0, 500);

  var patterns = [
    /(?:applying|applied) (?:for(?: the)?|to(?: the)?)\s+(?:position of\s+|role of\s+)?["""]?([A-Z][^"""\n,.(]{2,60})["""]?/i,
    /(?:position|role|opening|opportunity)(?:\s*:|\s+of|\s+for)?\s+["""]?([A-Z][^"""\n,.(]{2,60})["""]?/i,
    /(?:job title|title)\s*[:\-]\s*([^\n,.(]{2,60})/i,
  ];

  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match && match[1]) {
      var title = match[1].trim();
      // Sanity-check: skip if it looks like a sentence fragment
      if (title.split(' ').length <= 8) {
        return titleCase(title);
      }
    }
  }

  return '';
}

/**
 * Determine the application status from email content.
 * Returns the highest-priority matching status, or null if none match.
 */
function detectStatus(subject, body) {
  var text = (subject + ' ' + body).toLowerCase();

  for (var i = 0; i < STATUS_RULES.length; i++) {
    var rule = STATUS_RULES[i];
    for (var j = 0; j < rule.keywords.length; j++) {
      var kw = rule.keywords[j];
      try {
        if (new RegExp(kw, 'i').test(text)) {
          return rule.status;
        }
      } catch (e) {
        // Fallback to plain string match if regex is invalid
        if (text.indexOf(kw.toLowerCase()) !== -1) {
          return rule.status;
        }
      }
    }
  }

  return null;
}

// Row background colors per status
var STATUS_COLORS = {
  'Applied':   '#ffffff',  // white
  'Interview': '#c9daf8',  // blue
  'Offer':     '#b6d7a8',  // green
  'Rejected':  '#f4cccc',  // red
};

// Status priority order (higher index = higher priority)
var STATUS_PRIORITY = ['Applied', 'Interview', 'Offer', 'Rejected'];

function isHigherPriority(newStatus, currentStatus) {
  var newIdx     = STATUS_PRIORITY.indexOf(newStatus);
  var currentIdx = STATUS_PRIORITY.indexOf(currentStatus);
  // Offer and Rejected are both "terminal" — don't overwrite either with the other
  if (currentStatus === 'Offer' || currentStatus === 'Rejected') return false;
  return newIdx > currentIdx;
}

// --------------- SHEET HELPERS ---------------

function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    var headers = ['Company', 'Job Title', 'Date Applied', 'Status', 'Thread ID', 'Last Updated'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(COL.COMPANY, 180);
    sheet.setColumnWidth(COL.JOB_TITLE, 220);
    sheet.setColumnWidth(COL.DATE_APPLIED, 120);
    sheet.setColumnWidth(COL.STATUS, 110);
    sheet.setColumnWidth(COL.THREAD_ID, 160);
    sheet.setColumnWidth(COL.LAST_UPDATED, 140);
    Logger.log('Created new sheet: ' + CONFIG.SHEET_NAME);
  }

  return sheet;
}

function appendRow(sheet, company, jobTitle, date, status, threadId) {
  sheet.appendRow([company, jobTitle, date, status, threadId, new Date()]);
  var newRow = sheet.getLastRow();
  colorRow(sheet, newRow, status);
}

/** Sets the background color of an entire row based on status. */
function colorRow(sheet, rowNum, status) {
  var color = STATUS_COLORS[status] || '#ffffff';
  sheet.getRange(rowNum, 1, 1, 6).setBackground(color);
}

/**
 * Run this ONCE to backfill colors on all existing rows.
 * Safe to re-run anytime — just recolors everything based on current status.
 */
function colorAllRows() {
  var sheet = getOrCreateSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var status = data[i][COL.STATUS - 1];
    colorRow(sheet, i + 1, status);
  }
  Logger.log('Colored ' + (data.length - 1) + ' rows.');
  SpreadsheetApp.getActiveSpreadsheet().toast('All rows colored!', 'Done', 4);
}

/**
 * Returns the array index (0-based) of the row matching threadId, or -1.
 * Skips the header (index 0).
 */
function findRowIndexByThreadId(data, threadId) {
  for (var i = 1; i < data.length; i++) {
    if (data[i][COL.THREAD_ID - 1] === threadId) return i;
  }
  return -1;
}

// --------------- UTILITY ---------------

/** Returns a Gmail date filter string N days ago, e.g. "2024/01/01" */
function getDateFilter() {
  var d = new Date();
  d.setDate(d.getDate() - CONFIG.DAYS_BACK);
  var yyyy = d.getFullYear();
  var mm   = String(d.getMonth() + 1).padStart(2, '0');
  var dd   = String(d.getDate()).padStart(2, '0');
  return yyyy + '/' + mm + '/' + dd;
}

/** Basic title-case helper */
function titleCase(str) {
  return str.toLowerCase().replace(/(?:^|\s|-)\S/g, function(c) {
    return c.toUpperCase();
  });
}

// --------------- TRIGGER SETUP ---------------

/**
 * Run this ONCE to enable daily auto-scanning.
 * Creates a trigger that runs scanJobApplications() every morning at 8am.
 * Safe to call multiple times — won't create duplicates.
 */
function createDailyTrigger() {
  // Remove any existing RecruitSync triggers first to avoid duplicates
  removeDailyTrigger();

  ScriptApp.newTrigger('scanJobApplications')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Daily trigger created — scanJobApplications() will run every morning at 8am.');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Auto-scan enabled! RecruitSync will run every morning at 8am.',
    'Trigger created',
    5
  );
}

/**
 * PAUSE — stops auto-scanning but keeps all your data in the sheet.
 * Run createDailyTrigger() again to re-enable.
 */
function pauseRecruitSync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'scanJobApplications') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  if (removed > 0) {
    Logger.log('RecruitSync paused. Your sheet data is untouched.');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Auto-scan paused. Your data is safe. Run createDailyTrigger() to resume.',
      'RecruitSync paused',
      6
    );
  } else {
    Logger.log('No active triggers found — RecruitSync was already paused.');
  }
}

/**
 * FULL RESET — removes all triggers AND deletes the Job Applications sheet.
 * ⚠️  This cannot be undone. All logged data will be lost.
 * After this, you can also revoke Gmail access at myaccount.google.com/permissions.
 */
function resetRecruitSync() {
  // Remove all triggers
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'scanJobApplications') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Delete the sheet
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) {
    ss.deleteSheet(sheet);
    Logger.log('Sheet "' + CONFIG.SHEET_NAME + '" deleted.');
  } else {
    Logger.log('Sheet not found — may have already been deleted.');
  }

  Logger.log('RecruitSync fully reset. To also revoke Gmail access, visit: myaccount.google.com/permissions');
  Browser.msgBox(
    'RecruitSync Reset',
    'All data and triggers have been removed.\\n\\nTo fully revoke Gmail access, go to:\\nmyaccount.google.com/permissions',
    Browser.Buttons.OK
  );
}

// Alias for backwards compatibility
function removeDailyTrigger() { pauseRecruitSync(); }

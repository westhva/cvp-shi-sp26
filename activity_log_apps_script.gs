/**
 * Canva x SHI SMB Q2 2026 Blitz — Activity Log Apps Script
 *
 * Paste this entire script into your Google Sheet's Apps Script editor.
 * Tools → Apps Script → replace any existing code → Save → Run setup()
 *
 * What it does:
 *   1. onFormSubmit() — fires automatically when a rep submits the Google Form.
 *      Parses the account lists, enforces the 5-call/5-email cap per account,
 *      calculates points, updates the leaderboard, and emails the right people.
 *   2. setup() — creates required sheet tabs and installs the form trigger.
 *      Run ONCE after pasting.
 *   3. recalcAll() — recalculates everything from Submissions from scratch.
 *      Use after manual edits to submissions.
 *   4. clearTestData() — wipes all data except headers from all tabs.
 *      Use after testing to reset before launch.
 */

// ── Email config ──────────────────────────────────────────────────────────────
const EMAIL_WES = 'wes@canva.com';  // always gets every submission

// District managers — each only receives submissions from their district
const DISTRICT_MANAGERS = {
  'Small Business-West District':    'aj_yedinak@shi.com',
  'Small Business-Central District': 'Amable_Viloria@SHI.com',
  'Small Business-East District':    'Chris_Porterfield@shi.com'
};

// ── Points config ─────────────────────────────────────────────────────────────
const POINTS = {
  deal:         100,
  meeting:       40,
  registration:  25,
  call:           3,
  email:          1
};
const CAP_PER_ACCOUNT = 5; // max calls OR emails that earn points per account

// ── Rep → District map ────────────────────────────────────────────────────────
// Used to route email alerts to the right district manager
const REP_DISTRICT = {
  'mathew chacko':      'Small Business-West District',
  'hayden riba':        'Small Business-West District',
  'phillip jackson':    'Small Business-West District',
  'jessica jorgensen':  'Small Business-West District',
  'cornelius sirls':    'Small Business-West District',
  'maria martins':      'Small Business-West District',
  'kat nelson':         'Small Business-West District',
  'aj yedinak':         'Small Business-West District',
  'tatiana sultzer':    'Small Business-Central District',
  'logan shaw':         'Small Business-Central District',
  'daniel cecilio':     'Small Business-Central District',
  'isaac gallegos':     'Small Business-Central District',
  'michael white':      'Small Business-Central District',
  'john salas':         'Small Business-Central District',
  'olivia lora':        'Small Business-Central District',
  'jason fowler':       'Small Business-Central District',
  'keenan womack':      'Small Business-East District',
  'zach page':          'Small Business-East District',
  'austin dodson':      'Small Business-East District',
  'scott gray':         'Small Business-East District',
  'nancy bui':          'Small Business-East District',
  'olivia campbell':    'Small Business-East District',
  'art barron':         'Small Business-East District',
  'wyatt speaks':       'Small Business-East District',
  'aisha mohammed':     'Small Business-East District',
  'cade cunov':         'Small Business-East District',
  'chris porterfield':  'Small Business-East District'
};

// ── Sheet names ───────────────────────────────────────────────────────────────
const SHEET_SUBMISSIONS  = 'Submissions';
const SHEET_ACTIVITY     = 'Activity Log';
const SHEET_LEADERBOARD  = 'Leaderboard';
const SHEET_CAPS         = 'Cap Tracker';


// ── Trigger: fires on every form submission ───────────────────────────────────
function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const response = e.namedValues;

    const repName      = cleanStr(response['Rep Name']?.[0] || '');
    const dateRange    = cleanStr(response['Date Range']?.[0] || '');
    const calledRaw    = cleanStr(response['Accounts Called']?.[0] || '');
    const emailedRaw   = cleanStr(response['Accounts Emailed']?.[0] || '');
    const notes        = cleanStr(response['Anything else to add?']?.[0] || '');
    const timestamp    = new Date();

    if (!repName) return;

    const calledAccounts  = parseAccountList(calledRaw);
    const emailedAccounts = parseAccountList(emailedRaw);

    // Write to Activity Log
    const actSheet = ss.getSheetByName(SHEET_ACTIVITY);
    calledAccounts.forEach(account => {
      actSheet.appendRow([timestamp, repName, dateRange, 'Call', account, 1, notes]);
    });
    emailedAccounts.forEach(account => {
      actSheet.appendRow([timestamp, repName, dateRange, 'Email', account, 1, notes]);
    });
    parseManualNotes(notes, repName).forEach(item => {
      actSheet.appendRow([timestamp, repName, dateRange, item.type, item.account, item.count, notes]);
    });

    // Rebuild leaderboard
    rebuildLeaderboard(ss);

    // Send email alerts
    sendAlerts(repName, dateRange, calledAccounts, emailedAccounts, notes, timestamp);

  } catch(err) {
    Logger.log('onFormSubmit error: ' + err.toString());
  }
}


// ── Email alerts ──────────────────────────────────────────────────────────────
function sendAlerts(repName, dateRange, called, emailed, notes, timestamp) {
  try {
    const repKey   = repName.toLowerCase().trim();
    const district = REP_DISTRICT[repKey] || 'Unknown District';
    const distLabel = district.replace('Small Business-', '').replace(' District', '');
    const managerEmail = DISTRICT_MANAGERS[district];

    const subject = `[SHI Blitz] Activity logged — ${repName} (${distLabel})`;

    let body = `New activity submission from the SHI x Canva Q2 Blitz.\n\n`;
    body += `Rep: ${repName}\n`;
    body += `District: ${distLabel}\n`;
    body += `Period: ${dateRange || 'Not specified'}\n`;
    body += `Submitted: ${timestamp.toLocaleString()}\n\n`;

    if (called.length > 0) {
      body += `📞 Accounts Called (${called.length}):\n`;
      called.forEach(a => body += `  • ${a}\n`);
      body += '\n';
    }
    if (emailed.length > 0) {
      body += `📧 Accounts Emailed (${emailed.length}):\n`;
      emailed.forEach(a => body += `  • ${a}\n`);
      body += '\n';
    }
    if (notes) {
      body += `📝 Notes / Meetings / Deals:\n${notes}\n\n`;
    }

    body += `---\nView leaderboard: https://westhva.github.io/cvp-shi-sp26/leaderboard.html`;

    // Always email Wes
    MailApp.sendEmail(EMAIL_WES, subject, body);

    // Email the district manager (if different from Wes and if found)
    if (managerEmail && managerEmail.toLowerCase() !== EMAIL_WES.toLowerCase()) {
      MailApp.sendEmail(managerEmail, subject, body);
    }

    Logger.log(`Alerts sent to: ${EMAIL_WES}${managerEmail ? ', ' + managerEmail : ''}`);

  } catch(err) {
    Logger.log('sendAlerts error: ' + err.toString());
  }
}


// ── Parse a comma-separated account list ─────────────────────────────────────
function parseAccountList(raw) {
  if (!raw) return [];
  return raw.split(',')
    .map(s => s.trim())
    .filter(s => s.length > 0);
}


// ── Parse manual notes for meetings / deals / registrations ──────────────────
function parseManualNotes(notes) {
  const items = [];
  if (!notes) return items;

  const patterns = [
    { regex: /meeting[s]?:\s*([^,\n]+)/gi, type: 'Meeting' },
    { regex: /deal\s*reg(?:istration)?[s]?:\s*([^,\n]+)/gi, type: 'Registration' },
    { regex: /deal[s]?\s*booked?:\s*([^,\n]+)/gi, type: 'Deal' },
    { regex: /deal[s]?:\s*([^,\n]+)/gi, type: 'Deal' }
  ];

  patterns.forEach(p => {
    let match;
    const re = new RegExp(p.regex.source, 'gi');
    while ((match = re.exec(notes)) !== null) {
      const account = match[1].trim();
      if (account) items.push({ type: p.type, account: account, count: 1 });
    }
  });

  return items;
}


// ── Rebuild Leaderboard from Activity Log ─────────────────────────────────────
function rebuildLeaderboard(ss) {
  const actSheet = ss.getSheetByName(SHEET_ACTIVITY);
  const lbSheet  = ss.getSheetByName(SHEET_LEADERBOARD);
  const capSheet = ss.getSheetByName(SHEET_CAPS);

  const actData = actSheet.getDataRange().getValues();
  if (actData.length <= 1) return;

  const repMap = {};

  for (let i = 1; i < actData.length; i++) {
    const row     = actData[i];
    const rep     = cleanStr(String(row[1]));
    const type    = cleanStr(String(row[3]));
    const account = cleanStr(String(row[4]));
    const count   = parseInt(row[5]) || 1;

    if (!rep) continue;
    if (!repMap[rep]) repMap[rep] = { calls: {}, emails: {}, meetings: 0, deals: 0, regs: 0 };

    if      (type === 'Call')         repMap[rep].calls[account]  = (repMap[rep].calls[account]  || 0) + count;
    else if (type === 'Email')        repMap[rep].emails[account] = (repMap[rep].emails[account] || 0) + count;
    else if (type === 'Meeting')      repMap[rep].meetings += count;
    else if (type === 'Deal')         repMap[rep].deals    += count;
    else if (type === 'Registration') repMap[rep].regs     += count;
  }

  // Write Cap Tracker
  capSheet.clearContents();
  capSheet.appendRow(['Rep', 'Account', 'Calls (raw)', 'Calls (capped)', 'Emails (raw)', 'Emails (capped)']);
  Object.keys(repMap).sort().forEach(rep => {
    const allAccounts = new Set([...Object.keys(repMap[rep].calls), ...Object.keys(repMap[rep].emails)]);
    allAccounts.forEach(account => {
      capSheet.appendRow([
        rep, account,
        repMap[rep].calls[account]  || 0, Math.min(repMap[rep].calls[account]  || 0, CAP_PER_ACCOUNT),
        repMap[rep].emails[account] || 0, Math.min(repMap[rep].emails[account] || 0, CAP_PER_ACCOUNT)
      ]);
    });
  });

  // Write Leaderboard
  lbSheet.clearContents();
  lbSheet.appendRow(['Rep Name','Calls Made','Emails Sent','Meetings Booked','Deal Registrations','Deals Booked','Call Points','Email Points','Meeting Points','Reg Points','Deal Points','Total Points']);

  Object.keys(repMap).sort().forEach(rep => {
    const data = repMap[rep];
    const totalCalls  = Object.values(data.calls).reduce((s,c)  => s + Math.min(c, CAP_PER_ACCOUNT), 0);
    const totalEmails = Object.values(data.emails).reduce((s,c) => s + Math.min(c, CAP_PER_ACCOUNT), 0);
    const callPts  = totalCalls  * POINTS.call;
    const emailPts = totalEmails * POINTS.email;
    const meetPts  = data.meetings * POINTS.meeting;
    const regPts   = data.regs     * POINTS.registration;
    const dealPts  = data.deals    * POINTS.deal;
    const total    = callPts + emailPts + meetPts + regPts + dealPts;
    lbSheet.appendRow([rep, totalCalls, totalEmails, data.meetings, data.regs, data.deals, callPts, emailPts, meetPts, regPts, dealPts, total]);
  });

  // Sort by total points descending
  const lbRange  = lbSheet.getDataRange();
  const lbValues = lbRange.getValues();
  if (lbValues.length > 2) {
    const header = lbValues.shift();
    lbValues.sort((a, b) => b[11] - a[11]);
    lbValues.unshift(header);
    lbRange.setValues(lbValues);
  }
}


// ── One-time setup ────────────────────────────────────────────────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  [SHEET_SUBMISSIONS, SHEET_ACTIVITY, SHEET_LEADERBOARD, SHEET_CAPS].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  const actSheet = ss.getSheetByName(SHEET_ACTIVITY);
  if (actSheet.getLastRow() === 0) {
    actSheet.appendRow(['Timestamp','Rep Name','Date Range','Activity Type','Account','Count','Notes']);
  }

  const lbSheet = ss.getSheetByName(SHEET_LEADERBOARD);
  if (lbSheet.getLastRow() === 0) {
    lbSheet.appendRow(['Rep Name','Calls Made','Emails Sent','Meetings Booked','Deal Registrations','Deals Booked','Call Points','Email Points','Meeting Points','Reg Points','Deal Points','Total Points']);
  }

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();

  Logger.log('Setup complete.');
  SpreadsheetApp.getUi().alert('Setup complete! Sheets created and form trigger installed.');
}


// ── Clear test data ───────────────────────────────────────────────────────────
// Run this after testing to wipe all submissions before launch.
// Preserves headers in every sheet.
function clearTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Clear ALL test data?', 'This will wipe Submissions, Activity Log, Cap Tracker, and Leaderboard (headers kept). This cannot be undone.', ui.ButtonSet.OK_CANCEL);
  if (confirm !== ui.Button.OK) return;

  [SHEET_SUBMISSIONS, SHEET_ACTIVITY, SHEET_CAPS].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  });

  // Reset leaderboard to headers only
  const lbSheet = ss.getSheetByName(SHEET_LEADERBOARD);
  if (lbSheet) {
    lbSheet.clearContents();
    lbSheet.appendRow(['Rep Name','Calls Made','Emails Sent','Meetings Booked','Deal Registrations','Deals Booked','Call Points','Email Points','Meeting Points','Reg Points','Deal Points','Total Points']);
  }

  Logger.log('Test data cleared.');
  ui.alert('Done! All test data cleared. Ready for launch.');
}


// ── Manual recalculation ──────────────────────────────────────────────────────
function recalcAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actSheet = ss.getSheetByName(SHEET_ACTIVITY);
  actSheet.clearContents();
  actSheet.appendRow(['Timestamp','Rep Name','Date Range','Activity Type','Account','Count','Notes']);

  const subSheet = ss.getSheetByName(SHEET_SUBMISSIONS);
  const subData  = subSheet.getDataRange().getValues();
  if (subData.length <= 1) { Logger.log('No submissions.'); return; }

  const headers   = subData[0].map(h => cleanStr(String(h)).toLowerCase());
  const colRep     = headers.findIndex(h => h.includes('rep name'));
  const colDate    = headers.findIndex(h => h.includes('date'));
  const colCalled  = headers.findIndex(h => h.includes('called'));
  const colEmailed = headers.findIndex(h => h.includes('emailed'));
  const colNotes   = headers.findIndex(h => h.includes('anything'));

  for (let i = 1; i < subData.length; i++) {
    const row       = subData[i];
    const ts        = row[0] || new Date();
    const repName   = colRep     >= 0 ? cleanStr(String(row[colRep]))    : '';
    const dateRange = colDate    >= 0 ? cleanStr(String(row[colDate]))   : '';
    const called    = colCalled  >= 0 ? cleanStr(String(row[colCalled])) : '';
    const emailed   = colEmailed >= 0 ? cleanStr(String(row[colEmailed])): '';
    const notes     = colNotes   >= 0 ? cleanStr(String(row[colNotes]))  : '';

    if (!repName) continue;

    parseAccountList(called).forEach(a  => actSheet.appendRow([ts, repName, dateRange, 'Call',  a, 1, notes]));
    parseAccountList(emailed).forEach(a => actSheet.appendRow([ts, repName, dateRange, 'Email', a, 1, notes]));
    parseManualNotes(notes).forEach(item => actSheet.appendRow([ts, repName, dateRange, item.type, item.account, item.count, notes]));
  }

  rebuildLeaderboard(ss);
  SpreadsheetApp.getUi().alert('Recalculation complete!');
}


// ── Utilities ─────────────────────────────────────────────────────────────────
function cleanStr(s) { return String(s || '').trim().replace(/\s+/g, ' '); }

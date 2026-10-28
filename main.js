/*******************************************************************
 * BASECAMP → GOOGLE SHEETS SYNC
 *
 * Paste this entire file into Extensions → Apps Script (Code.gs) of
 * this spreadsheet:
 * https://docs.google.com/spreadsheets/d/1fswZencFkJoVOluYvPK4BQmNDX1O_YBRQEy0_CYV49s
 *
 * See the accompanying setup guide for step-by-step instructions.
 *******************************************************************/

// ============================================================
// 0. FIXED CONFIG (safe to edit — none of this is a secret)
// ============================================================
const SHEET_NAME_DEFAULT = null;          // null = use the first sheet (gid=0)
const SYNC_LOG_SHEET_NAME = 'Sync Log';
const TIMEZONE = 'Asia/Kolkata';          // India (Mumbai) — used for all dates, due-date calculations, and the Activity Time column
const BASECAMP_USER_AGENT = 'GoogleSheetsBasecampSync (set-your-email@example.com)';

// Visible columns (do not change without updating the sheet header)
const COL = {
  ASSIGNED_DATE: 1,   // A
  DUE_DATE: 2,        // B
  PROJECT_NAME: 3,    // C
  BASECAMP_URL_RAW: 4,// D  — raw Basecamp task/comment URL (visible, plain link text)
  WORK_TITLE: 5,      // E
  WEBSITE_LINK: 6,    // F
  ASSIGNEE_NAME: 7,   // G  ("Added By")
  MY_STATUS: 8,       // H  — manual dropdown (Done / Pending / Working). Script only READS this, never writes it.
  ACTIVITY_TIME: 9,   // I  — date+time of original assignment, OR of the latest @mention/comment that brought this row to the bottom
  BASECAMP_TODO_ID: 10,   // J (hidden helper)
  BASECAMP_PROJECT_ID: 11,// K (hidden helper)
  STATUS: 12,             // L (hidden helper)
  LAST_UPDATED: 13,       // M (hidden helper)
  LAST_SYNCED: 14,        // N (hidden helper)
  TASK_ID: 15             // O  — clickable HYPERLINK ("Open Task")
};
// Total number of sheet columns this script manages (the widest column index).
// Use this — NOT any single column constant — whenever reading a whole row.
const COL_COUNT = Math.max.apply(null, Object.keys(COL).map(function (k) { return COL[k]; }));
const HEADER_ROW = 1;
const MY_STATUS_VALUES = ['Done', 'Pending', 'Working', 'DONE By Assigner'];
const AUTO_STATUS_LABEL = 'DONE By Assigner'; // written automatically when Basecamp itself marks the task complete

// Colours
const COLOR = {
  NEW_OR_FAR: '#D9EAF7',     // >7 days / brand new / no due date
  DUE_7: '#FCE5CD',          // 5-7 days
  DUE_4: '#F9CB9C',          // 3-4 days
  DUE_2: '#E06666',          // 1-2 days
  DUE_TODAY: '#CC0000',      // due today
  OVERDUE: '#990000',        // overdue
  COMPLETED: '#96fe0e',      // manual "Done" — bright green
  COMPLETED_BY_ASSIGNER: '#274E13', // Basecamp itself marked it complete — forest/dark green
  WORKING: '#e3ea1d'         // manual "Working" status
};
const FONT = {
  DEFAULT: '#000000',
  WHITE: '#FFFFFF',          // used on overdue/due-today red backgrounds for readability
  RED: '#CC0000'             // used on the yellow "Working" background once it's overdue
};

// Domains that are never treated as a client "website" when scanning
// task/project descriptions for a website URL.
const EXCLUDED_WEBSITE_DOMAINS = [
  'basecamp.com', '3.basecamp.com', 'basecampapi.com', 'launchpad.37signals.com',
  'docs.google.com', 'drive.google.com', 'sheets.google.com', 'calendar.google.com',
  'slides.google.com', 'forms.google.com', 'mail.google.com', 'goo.gl', 'g.page',
  'zoom.us', 'slack.com', 'dropbox.com', 'wetransfer.com', 'trello.com'
];

// ============================================================
// 1. MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Basecamp Sync')
    .addItem('Sync Now', 'syncNow')
    .addItem('Refresh Colours', 'refreshColoursOnly')
    .addItem('Jump to Latest Activity', 'jumpToLatestActivity_')
    .addSeparator()
    .addItem('Setup Basecamp (Step 1: Authorize)', 'setupBasecamp')
    .addItem('Find My Basecamp User ID', 'findMyBasecampUserId')
    .addItem('Find a Person\'s ID by Name (for TL/mentions)', 'findPersonIdByName')
    .addItem('Test Basecamp Connection', 'testBasecampConnection')
    .addItem('One-Time: Insert Activity Time Column', 'migrateAddActivityTimeColumn')
    .addItem('Fix Task Link Columns (D = URL, O = Open Task)', 'fixTaskLinkColumns')
    .addItem('Show Configuration Status', 'showConfigStatus')
    .addSeparator()
    .addItem('Create Hourly Auto-Sync Trigger', 'createAutoSyncTrigger')
    .addItem('Create 5-Minute Auto-Sync Trigger', 'create5MinuteAutoSyncTrigger')
    .addItem('Remove Auto-Sync Triggers', 'removeAutoSyncTriggers')
    .addToUi();

  // Every time the sheet is opened (or the browser tab is refreshed, which
  // re-fires this same trigger), jump straight to whichever row has the
  // most recent "Activity Time" — the newest assignment or @mention.
  try {
    jumpToLatestActivity_();
  } catch (e) {
    // Never let this block the menu from loading.
  }
}

// Parses Activity Time text such as "27/10/2026 17:34" or "27/10/2026 05:34 PM"
// into a comparable timestamp. AM/PM and seconds are OPTIONAL on purpose:
// Google Sheets converts the text we write into a real date-time and then
// DISPLAYS it in 24-hour form without AM/PM, so a parser that demands AM/PM
// finds nothing and the jump silently does nothing.
function parseActivityTimeText_(text) {
  if (!text) return null;
  const m = String(text).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (!m) return null;
  const dd = m[1], mm = m[2], yyyy = m[3], hhRaw = m[4], min = m[5], ampm = m[6];
  let hh = parseInt(hhRaw, 10);
  if (ampm) {
    if (/pm/i.test(ampm) && hh !== 12) hh += 12;
    if (/am/i.test(ampm) && hh === 12) hh = 0;
  }
  const d = new Date(Number(yyyy), Number(mm) - 1, Number(dd), hh, Number(min));
  return isNaN(d.getTime()) ? null : d.getTime();
}

// Finds the row with the latest Activity Time (column I) and selects the full
// row, which scrolls it into view. Runs from onOpen, so it fires every time
// the sheet is opened or the browser tab is refreshed.
function jumpToLatestActivity_() {
  const sheet = getMainSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < HEADER_ROW + 1) return;

  const range = sheet.getRange(HEADER_ROW + 1, COL.ACTIVITY_TIME, lastRow - HEADER_ROW, 1);
  const rawValues = range.getValues();          // real Date objects when Sheets parsed the text
  const shownValues = range.getDisplayValues(); // what you see on screen

  let latestRow = -1;
  let latestTime = -Infinity;

  for (let i = 0; i < rawValues.length; i++) {
    const raw = rawValues[i][0];
    let t = null;
    if (raw instanceof Date && !isNaN(raw.getTime())) {
      t = raw.getTime();
    } else {
      t = parseActivityTimeText_(shownValues[i][0] || raw);
    }
    if (t !== null && t >= latestTime) { // ">=" so the lower row wins a tie
      latestTime = t;
      latestRow = HEADER_ROW + 1 + i;
    }
  }

  if (latestRow === -1) return; // no valid Activity Time values yet

  SpreadsheetApp.setActiveSheet(sheet);
  sheet.setActiveRange(sheet.getRange(latestRow, 1, 1, sheet.getLastColumn()));
  SpreadsheetApp.flush();
}

// ============================================================
// 2. CONFIG / PROPERTIES HELPERS
// ============================================================
function getProps_() {
  return PropertiesService.getScriptProperties();
}

function getConfig_() {
  const p = getProps_().getProperties();
  return {
    clientId: p.BASECAMP_CLIENT_ID || '',
    clientSecret: p.BASECAMP_CLIENT_SECRET || '',
    accountId: p.BASECAMP_ACCOUNT_ID || '',
    accessToken: p.BASECAMP_ACCESS_TOKEN || '',
    refreshToken: p.BASECAMP_REFRESH_TOKEN || '',
    tokenExpiresAt: p.BASECAMP_TOKEN_EXPIRES_AT || '',
    myPersonId: p.MY_BASECAMP_PERSON_ID || '',
    sheetName: p.SHEET_NAME || SHEET_NAME_DEFAULT,
    extractWebsiteFromProjectName: (p.EXTRACT_WEBSITE_FROM_PROJECT_NAME || 'true') === 'true',
    redirectUri: p.BASECAMP_REDIRECT_URI || '',
    mySgid: p.MY_ATTACHABLE_SGID || '',
    watchedTlIds: (p.WATCHED_TL_PERSON_IDS || '').split(',').map(s => s.trim()).filter(Boolean),
    mentionScanBudgetMs: Number(p.MENTION_SCAN_TIME_BUDGET_SECONDS || 240) * 1000,
    reconcileBudgetMs: Number(p.RECONCILE_TIME_BUDGET_SECONDS || 90) * 1000
  };
}

// The OAuth redirect URI MUST be a fixed value that both setupBasecamp()
// and doGet() agree on, and MUST exactly match what's registered at
// launchpad.37signals.com/integrations.
//
// Do NOT rely on ScriptApp.getService().getUrl() for this: if this project
// has more than one Web App deployment (e.g. you created a "New deployment"
// after editing code, instead of pushing a "New version" onto the SAME
// deployment), getUrl() can resolve to a different deployment ID than the
// one you actually registered with Basecamp, causing OAuth to look "stuck"
// or callback into the wrong context.
//
// Fix: deploy ONE Web App, copy its exact /exec URL, save it as the
// BASECAMP_REDIRECT_URI script property AND as the Redirect URI on
// Basecamp's integration settings. From then on, whenever you change the
// code, update that SAME deployment via Deploy → Manage deployments →
// pencil icon → Version: "New version" → Deploy — this keeps the /exec
// URL identical while picking up your new code. Do not click
// "New deployment" for routine code changes; that mints a new URL.
function getRedirectUri_() {
  const c = getConfig_();
  if (c.redirectUri) return c.redirectUri;
  // Fallback only — prone to the mismatch described above. Prefer setting
  // BASECAMP_REDIRECT_URI explicitly.
  return ScriptApp.getService().getUrl();
}

// Run this ONCE from the menu. It inserts a new visible "Activity Time"
// column at I (shifting the hidden helper columns from I:N to J:O), safely
// preserving all existing data — Google Sheets shifts formulas, values, and
// formatting automatically on insertColumnBefore. Safe to run more than
// once: it checks whether the column is already there first.
function migrateAddActivityTimeColumn() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getMainSheet_();
  const currentColI = String(sheet.getRange(HEADER_ROW, 9).getValue() || '');
  if (currentColI.indexOf('Activity Time') !== -1) {
    ui.alert('Already migrated — column I is already "Activity Time". No changes made.');
    return;
  }
  sheet.insertColumnBefore(9);
  sheet.getRange(HEADER_ROW, 9).setValue('Activity Time');
  ui.alert('Migration complete',
    'Column I is now "Activity Time". Your Basecamp Todo ID / Project ID / Status / etc. columns have shifted one to ' +
    'the right (J onward) — Google Sheets moved that data automatically, nothing was lost. Run "Sync Now" to start ' +
    'filling in Activity Time going forward.',
    ui.ButtonSet.OK);
}

// Bring columns D and O into the intended layout, whatever state they are in:
//   D = the raw Basecamp URL as plain text (full, readable, includes the
//       "#__recording_..." part when someone re-mentioned you)
//   O = the clickable "Open Task" link pointing at that same URL
// It reads whichever of D / O currently holds a URL (plain text OR inside a
// HYPERLINK formula), keeps the more specific one (the one with
// "#__recording_"), and rewrites both. Safe to run any number of times, and it
// also runs automatically at the start of every sync.
function fixTaskLinkColumns() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getMainSheet_();
  sheet.getRange(HEADER_ROW, COL.BASECAMP_URL_RAW).setValue('Basecamp Task URL (raw)');
  sheet.getRange(HEADER_ROW, COL.TASK_ID).setValue('Task ID');
  const changed = normalizeTaskLinkColumns_(sheet);
  sheet.autoResizeColumn(COL.BASECAMP_URL_RAW);
  ui.alert('Task link columns fixed',
    (changed ? 'Corrected rows where D / O were mixed up. ' : 'Nothing needed correcting. ') +
    'Column D now shows the full raw URL and column O holds the "Open Task" link.',
    ui.ButtonSet.OK);
}

function extractUrlFromCell_(value, formula) {
  const re = /https?:\/\/[^\s"'<>]+/i;
  const f = String(formula || '');
  if (f) {
    const m = f.match(re);
    if (m) return m[0];
  }
  const v = String(value || '').trim();
  return /^https?:\/\//i.test(v) ? v : '';
}

// Returns true if it changed anything. Rows without a Basecamp Todo ID
// (your own manual rows) are left exactly as they are.
function normalizeTaskLinkColumns_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return false;
  const n = lastRow - HEADER_ROW;

  const dRange = sheet.getRange(HEADER_ROW + 1, COL.BASECAMP_URL_RAW, n, 1);
  const oRange = sheet.getRange(HEADER_ROW + 1, COL.TASK_ID, n, 1);
  const dVals = dRange.getValues(), dForm = dRange.getFormulas();
  const oVals = oRange.getValues(), oForm = oRange.getFormulas();
  const ids = sheet.getRange(HEADER_ROW + 1, COL.BASECAMP_TODO_ID, n, 1).getValues();

  const hasAnchor = function (u) { return u.indexOf('#__recording_') !== -1; };
  const newD = [], newO = [];
  let changed = false;

  for (let i = 0; i < n; i++) {
    const keepD = dForm[i][0] || dVals[i][0];
    const keepO = oForm[i][0] || oVals[i][0];
    if (!ids[i][0]) { newD.push([keepD]); newO.push([keepO]); continue; }

    const urlD = extractUrlFromCell_(dVals[i][0], dForm[i][0]);
    const urlO = extractUrlFromCell_(oVals[i][0], oForm[i][0]);
    let url = urlD || urlO;
    if (urlD && urlO && urlD !== urlO) url = hasAnchor(urlO) && !hasAnchor(urlD) ? urlO : urlD;
    if (!url) { newD.push([keepD]); newO.push([keepO]); continue; }

    const wantO = '=HYPERLINK("' + url + '","Open Task")';
    if (String(dVals[i][0]) !== url || dForm[i][0] || oForm[i][0] !== wantO) changed = true;
    newD.push([url]);
    newO.push([wantO]);
  }

  if (changed) {
    dRange.setValues(newD);
    oRange.setValues(newO);
    sheet.autoResizeColumn(COL.BASECAMP_URL_RAW);
  }
  return changed;
}

function showConfigStatus() {
  const c = getConfig_();
  const mask = v => v ? ('set (' + String(v).slice(0, 4) + '…)') : 'NOT SET';
  const msg = [
    'BASECAMP_CLIENT_ID: ' + mask(c.clientId),
    'BASECAMP_CLIENT_SECRET: ' + (c.clientSecret ? 'set' : 'NOT SET'),
    'BASECAMP_ACCOUNT_ID: ' + (c.accountId || 'NOT SET'),
    'BASECAMP_ACCESS_TOKEN: ' + (c.accessToken ? 'set' : 'NOT SET'),
    'BASECAMP_REFRESH_TOKEN: ' + (c.refreshToken ? 'set' : 'NOT SET'),
    'MY_BASECAMP_PERSON_ID: ' + (c.myPersonId || 'NOT SET'),
    'MY_ATTACHABLE_SGID: ' + (c.mySgid ? 'set' : 'NOT SET (re-run Find My Basecamp User ID — mention detection needs this)'),
    'SHEET_NAME override: ' + (c.sheetName || '(using first sheet)'),
    '',
    'BASECAMP_REDIRECT_URI (pinned): ' + (c.redirectUri || 'NOT SET — falling back to getUrl(), see below'),
    'ScriptApp.getService().getUrl() right now: ' + (ScriptApp.getService().getUrl() || '(null — not deployed as Web App)'),
    (c.redirectUri && ScriptApp.getService().getUrl() && c.redirectUri !== ScriptApp.getService().getUrl())
      ? '⚠ MISMATCH — these two URLs differ. Basecamp must be registered with the pinned one above.'
      : ''
  ].join('\n');
  SpreadsheetApp.getUi().alert('Basecamp Sync — Configuration Status', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
// 3. OAUTH — AUTHORIZATION
// ============================================================
// Basecamp OAuth lives at 37signals Launchpad:
//   Authorize: https://launchpad.37signals.com/authorization/new
//   Token:     https://launchpad.37signals.com/authorization/token
// This Apps Script project must be deployed as a Web App so its
// /exec URL can serve as the OAuth redirect_uri (doGet below
// captures the "code" parameter). See the setup guide.

function setupBasecamp() {
  const ui = SpreadsheetApp.getUi();
  const c = getConfig_();
  if (!c.clientId || !c.clientSecret) {
    ui.alert('Missing credentials',
      'Before authorizing, set BASECAMP_CLIENT_ID and BASECAMP_CLIENT_SECRET in ' +
      'Project Settings → Script Properties. See the setup guide for how to register ' +
      'an integration at launchpad.37signals.com/integrations.',
      ui.ButtonSet.OK);
    return;
  }

  const redirectUri = getRedirectUri_();
  if (!redirectUri) {
    ui.alert('Web app not deployed',
      'Deploy this project as a Web App first (Deploy → New deployment → Web app), ' +
      'then save that exact /exec URL as the BASECAMP_REDIRECT_URI script property, ' +
      'set the same URL as the Redirect URI on launchpad.37signals.com/integrations, ' +
      'and run Setup Basecamp again.',
      ui.ButtonSet.OK);
    return;
  }
  if (!getConfig_().redirectUri) {
    ui.alert('BASECAMP_REDIRECT_URI not pinned',
      'You have not set the BASECAMP_REDIRECT_URI script property, so this is falling back to ' +
      'ScriptApp.getService().getUrl(), which can silently point at the wrong deployment if this ' +
      'project has more than one Web App deployment. Strongly recommended: set ' +
      'BASECAMP_REDIRECT_URI to the exact /exec URL below and re-run this.',
      ui.ButtonSet.OK);
  }

  const authUrl = 'https://launchpad.37signals.com/authorization/new'
    + '?type=web_server'
    + '&client_id=' + encodeURIComponent(c.clientId)
    + '&redirect_uri=' + encodeURIComponent(redirectUri);

  const html = HtmlService.createHtmlOutput(
    '<p>Your Web App redirect URI is:</p><p><code>' + redirectUri + '</code></p>' +
    '<p>Make sure that exact URL is saved as the Redirect URI for your integration at ' +
    '<a href="https://launchpad.37signals.com/integrations" target="_blank">launchpad.37signals.com/integrations</a>.</p>' +
    '<p><a href="' + authUrl + '" target="_blank"><b>Click here to authorize with Basecamp</b></a></p>' +
    '<p>After you approve access, you will land on a confirmation page. Then come back here — you are done.</p>'
  ).setWidth(480).setHeight(260);
  ui.showModalDialog(html, 'Authorize Basecamp');
}

// Handles the OAuth redirect from Basecamp (the Web App's /exec URL).
function doGet(e) {
  const code = e && e.parameter && e.parameter.code;
  if (!code) {
    return HtmlService.createHtmlOutput('<p>No authorization code received. Close this tab and try Setup Basecamp again.</p>');
  }
  try {
    const c = getConfig_();
    const redirectUri = getRedirectUri_();
    const resp = UrlFetchApp.fetch('https://launchpad.37signals.com/authorization/token', {
      method: 'post',
      payload: {
        type: 'web_server',
        client_id: c.clientId,
        client_secret: c.clientSecret,
        redirect_uri: redirectUri,
        code: code
      },
      muteHttpExceptions: true
    });
    const status = resp.getResponseCode();
    const body = JSON.parse(resp.getContentText());
    if (status !== 200 || !body.access_token) {
      return HtmlService.createHtmlOutput('<p>Token exchange failed (HTTP ' + status + '). Check Client ID/Secret and try again.</p>');
    }
    storeTokens_(body.access_token, body.refresh_token);
    autoDetectAccountId_(body.access_token);
    return HtmlService.createHtmlOutput('<p><b>Basecamp authorization complete.</b> You can close this tab and return to your Google Sheet.</p>');
  } catch (err) {
    return HtmlService.createHtmlOutput('<p>Error during authorization: ' + err.message + '</p>');
  }
}

function storeTokens_(accessToken, refreshToken) {
  const props = getProps_();
  // Basecamp access tokens currently last ~2 weeks. We store an expiry
  // estimate slightly conservative (13 days) so we refresh before expiry.
  const expiresAt = new Date().getTime() + 13 * 24 * 60 * 60 * 1000;
  const toSet = {
    BASECAMP_ACCESS_TOKEN: accessToken,
    BASECAMP_TOKEN_EXPIRES_AT: String(expiresAt)
  };
  if (refreshToken) toSet.BASECAMP_REFRESH_TOKEN = refreshToken;
  props.setProperties(toSet, false);
}

function autoDetectAccountId_(accessToken) {
  const resp = UrlFetchApp.fetch('https://launchpad.37signals.com/authorization.json', {
    headers: { Authorization: 'Bearer ' + accessToken, 'User-Agent': BASECAMP_USER_AGENT },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) return;
  const data = JSON.parse(resp.getContentText());
  const accounts = data.accounts || [];
  const bc3 = accounts.find(a => a.product === 'bc3');
  if (bc3) {
    getProps_().setProperty('BASECAMP_ACCOUNT_ID', String(bc3.id));
  }
}

// Refreshes the access token using the stored refresh token.
// Returns a valid access token string, or throws if refresh fails.
function getValidAccessToken_() {
  const c = getConfig_();
  const now = new Date().getTime();
  const expiresAt = Number(c.tokenExpiresAt || 0);
  if (c.accessToken && expiresAt && now < expiresAt) {
    return c.accessToken;
  }
  if (!c.refreshToken) {
    if (c.accessToken) return c.accessToken; // best effort, may be expired
    throw new Error('Not authorized with Basecamp yet. Run "Setup Basecamp" from the menu.');
  }
  const resp = UrlFetchApp.fetch('https://launchpad.37signals.com/authorization/token', {
    method: 'post',
    payload: {
      type: 'refresh',
      refresh_token: c.refreshToken,
      client_id: c.clientId,
      client_secret: c.clientSecret
    },
    muteHttpExceptions: true
  });
  const status = resp.getResponseCode();
  const body = JSON.parse(resp.getContentText());
  if (status !== 200 || !body.access_token) {
    throw new Error('Failed to refresh Basecamp access token (HTTP ' + status + '). Re-run Setup Basecamp.');
  }
  storeTokens_(body.access_token, body.refresh_token || c.refreshToken);
  return body.access_token;
}

// ============================================================
// 4. BASECAMP API HELPERS
// ============================================================
function bcRequest_(path) {
  const c = getConfig_();
  if (!c.accountId) throw new Error('BASECAMP_ACCOUNT_ID is not set.');
  const token = getValidAccessToken_();
  const url = 'https://3.basecampapi.com/' + c.accountId + path;
  return bcFetchWithRetry_(url, token);
}

function bcFetchWithRetry_(url, token) {
  const maxAttempts = 5;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token, 'User-Agent': BASECAMP_USER_AGENT },
      muteHttpExceptions: true
    });
    const status = resp.getResponseCode();
    if (status === 200) {
      return { json: JSON.parse(resp.getContentText() || 'null'), headers: resp.getAllHeaders() };
    }
    if (status === 429) {
      const retryAfter = Number(resp.getAllHeaders()['Retry-After'] || 5);
      Utilities.sleep((retryAfter + 1) * 1000);
      continue;
    }
    if (status === 401) {
      throw new Error('Basecamp returned 401 Unauthorized. Your token may be invalid — re-run Setup Basecamp.');
    }
    throw new Error('Basecamp API error ' + status + ' for ' + url + ': ' + resp.getContentText());
  }
  throw new Error('Basecamp API rate-limited too many times for ' + url);
}

// Generic Link-header paginator, kept for future extension (e.g. listing
// all projects). The "My Assignments" endpoints used below are NOT
// paginated per Basecamp's documentation, so this is not needed for the
// core sync, but is provided so the script scales if extended later.
function bcFetchAllPages_(path) {
  const c = getConfig_();
  const token = getValidAccessToken_();
  let url = 'https://3.basecampapi.com/' + c.accountId + path;
  let results = [];
  let guard = 0;
  while (url && guard < 200) {
    guard++;
    const { json, headers } = bcFetchWithRetry_(url, token);
    results = results.concat(json);
    url = null;
    const link = headers['Link'] || headers['link'];
    if (link) {
      const match = link.split(',').map(s => s.trim()).find(s => s.endsWith('rel="next"'));
      if (match) {
        const m = match.match(/<([^>]+)>/);
        if (m) url = m[1];
      }
    }
  }
  return results;
}

// ============================================================
// 5. FIND MY PERSON ID / TEST CONNECTION
// ============================================================
function findMyBasecampUserId() {
  const ui = SpreadsheetApp.getUi();
  try {
    const { json } = bcRequest_('/my/profile.json');
    getProps_().setProperty('MY_BASECAMP_PERSON_ID', String(json.id));
    let sgidMsg = '';
    if (json.attachable_sgid) {
      getProps_().setProperty('MY_ATTACHABLE_SGID', json.attachable_sgid);
      sgidMsg = '\nMention-detection ID also saved (MY_ATTACHABLE_SGID).';
    } else {
      sgidMsg = '\nWarning: attachable_sgid not returned — mention detection (WATCHED_TL_PERSON_IDS) will not work until this is available.';
    }
    ui.alert('Basecamp User ID Found',
      'Name: ' + json.name + '\nPerson ID: ' + json.id +
      '\n\nThis has been saved as MY_BASECAMP_PERSON_ID automatically.' + sgidMsg,
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Could not fetch profile', err.message, ui.ButtonSet.OK);
  }
}

// Looks up a colleague's Basecamp person ID by (partial) name — use this to
// find your TL's ID for WATCHED_TL_PERSON_IDS.
function findPersonIdByName() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Find a Person\'s Basecamp ID',
    'Enter (part of) their name, e.g. Tushar:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const query = resp.getResponseText().trim().toLowerCase();
  if (!query) return;
  try {
    const people = bcFetchAllPages_('/people.json');
    const matches = people.filter(p => p.name && p.name.toLowerCase().indexOf(query) !== -1);
    if (matches.length === 0) {
      ui.alert('No matches found for "' + query + '".');
      return;
    }
    const list = matches.map(p => p.name + '  —  ID: ' + p.id).join('\n');
    ui.alert('Matches',
      list + '\n\nCopy the ID(s) you want into the WATCHED_TL_PERSON_IDS script property ' +
      '(comma-separated if watching more than one person).',
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Error', err.message, ui.ButtonSet.OK);
  }
}

function testBasecampConnection() {
  const ui = SpreadsheetApp.getUi();
  try {
    const { json } = bcRequest_('/my/profile.json');
    ui.alert('Connection OK', 'Connected as ' + json.name + ' (ID ' + json.id + ').', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('Connection failed', err.message, ui.ButtonSet.OK);
  }
}

// ============================================================
// 6. TRIGGERS
// ============================================================
function createAutoSyncTrigger() {
  removeAutoSyncTriggers();
  ScriptApp.newTrigger('syncNow').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('An hourly trigger for syncNow() has been created.');
}

function create5MinuteAutoSyncTrigger() {
  removeAutoSyncTriggers();
  ScriptApp.newTrigger('syncNow').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('A 5-minute trigger for syncNow() has been created. Since the mention scan is incremental (only checks what changed since the last run), this stays fast.');
}

function removeAutoSyncTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncNow') ScriptApp.deleteTrigger(t);
  });
}

// ============================================================
// 7. SHEET HELPERS
// ============================================================
function getMainSheet_() {
  const c = getConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (c.sheetName) {
    const sh = ss.getSheetByName(c.sheetName);
    if (sh) return sh;
  }
  return ss.getSheets()[0];
}

function getOrCreateSyncLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SYNC_LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SYNC_LOG_SHEET_NAME);
    sh.appendRow(['Timestamp', 'Action', 'Basecamp Task ID', 'Project', 'Result', 'Error Message']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function logSync_(action, taskId, project, result, errorMessage) {
  try {
    const sh = getOrCreateSyncLogSheet_();
    sh.appendRow([new Date(), action, taskId || '', project || '', result, errorMessage || '']);
  } catch (e) {
    // If logging itself fails, do not crash the sync.
  }
}

// Builds a map of Basecamp Todo ID -> spreadsheet row number, for
// dedupe/update matching. Rows without a Todo ID in column H are
// treated as manual/unrelated rows and are never touched.
function buildExistingRowIndex_(sheet) {
  const lastRow = sheet.getLastRow();
  const index = {};
  if (lastRow < HEADER_ROW + 1) return index;
  const ids = sheet.getRange(HEADER_ROW + 1, COL.BASECAMP_TODO_ID, lastRow - HEADER_ROW, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i][0];
    if (id) index[String(id)] = HEADER_ROW + 1 + i;
  }
  return index;
}

// ============================================================
// 8. WEBSITE / PROJECT NAME EXTRACTION
// ============================================================
const DOMAIN_LIKE_REGEX = /^(https?:\/\/)?(www\.)?([a-z0-9-]+(\.[a-z0-9-]+)+)\/?$/i;
const URL_IN_TEXT_REGEX = /https?:\/\/[^\s"'<>)]+/gi;

// "Smilestudioballarat.com.au" (bucket/project name) -> "Smilestudioballarat"
function deriveProjectDisplayName_(bucketName) {
  if (!bucketName) return '';
  const trimmed = bucketName.trim();
  if (DOMAIN_LIKE_REGEX.test(trimmed)) {
    return trimmed.split('.')[0];
  }
  return trimmed;
}

function stripHtml_(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/\s+/g, ' ')
    .trim();
}

function isExcludedDomain_(hostname) {
  hostname = hostname.toLowerCase();
  return EXCLUDED_WEBSITE_DOMAINS.some(d => hostname === d || hostname.endsWith('.' + d));
}

// Looks for a plausible client website URL in free text (task/project
// description). Returns the first non-excluded external URL found, or
// '' if none. Never invents a URL.
function findWebsiteUrlInText_(text) {
  if (!text) return '';
  const matches = text.match(URL_IN_TEXT_REGEX) || [];
  for (const raw of matches) {
    try {
      const url = raw.replace(/[).,;]+$/, ''); // trim trailing punctuation
      const host = url.replace(/^https?:\/\//i, '').split('/')[0];
      if (!isExcludedDomain_(host)) return url;
    } catch (e) { /* skip malformed */ }
  }
  return '';
}

// Combines description text + project-name-as-domain heuristic.
function resolveWebsiteLink_(descriptionText, bucketName, allowProjectNameFallback) {
  const fromText = findWebsiteUrlInText_(descriptionText);
  if (fromText) return fromText;
  if (allowProjectNameFallback && bucketName && DOMAIN_LIKE_REGEX.test(bucketName.trim())) {
    const domain = bucketName.trim().replace(/^https?:\/\//i, '').replace(/^www\./i, '');
    return 'https://' + domain + '/';
  }
  return '';
}

// ============================================================
// 9. DATE / COLOUR LOGIC
// ============================================================
function formatDateAU_(dateOrString) {
  if (!dateOrString) return '';
  const d = (dateOrString instanceof Date) ? dateOrString : new Date(dateOrString + 'T00:00:00');
  return Utilities.formatDate(d, TIMEZONE, 'dd/MM/yyyy');
}

// Full date + time (Basecamp timestamps come as ISO datetimes, e.g.
// "2026-09-15T10:02:14.000+05:30") — used for the Activity Time column so
// you can see exactly when a task was assigned or when someone tagged you.
function formatDateTimeAU_(isoDateTimeString) {
  if (!isoDateTimeString) return '';
  const d = new Date(isoDateTimeString);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, TIMEZONE, 'dd/MM/yyyy hh:mm a');
}

function daysRemaining_(dueDateStr) {
  if (!dueDateStr) return null;
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  const today = new Date(todayStr + 'T00:00:00');
  const due = new Date(dueDateStr + 'T00:00:00');
  return Math.round((due - today) / (1000 * 60 * 60 * 24));
}

function computeRowColor_(isCompleted, dueDateStr) {
  if (isCompleted) return COLOR.COMPLETED;
  const days = daysRemaining_(dueDateStr);
  if (days === null) return COLOR.NEW_OR_FAR;
  if (days < 0) return COLOR.OVERDUE;
  if (days === 0) return COLOR.DUE_TODAY;
  if (days <= 2) return COLOR.DUE_2;
  if (days <= 4) return COLOR.DUE_4;
  if (days <= 7) return COLOR.DUE_7;
  return COLOR.NEW_OR_FAR;
}

function computeStatusLabel_(isCompleted, dueDateStr) {
  if (isCompleted) return 'Completed';
  const days = daysRemaining_(dueDateStr);
  if (days === null) return 'Active';
  if (days < 0) return 'Overdue';
  if (days === 0) return 'Due Today';
  return 'Active';
}

// Combines the manual "My Status" dropdown (column H: Done / Pending /
// Working / DONE By Assigner / blank) with Basecamp's own completion + due
// date to decide the row's background AND font colour. Rules (highest
// priority first):
//   1. Basecamp itself says the to-do is complete (or the dropdown already
//      reads "DONE By Assigner") -> forest/dark green, white text. This is
//      distinct from a manual "Done" so you can tell who closed it.
//   2. My Status = Done (set by you) -> light green, regardless of due date.
//   3. My Status = Working -> yellow background; if the due date has passed
//      while still "Working", the TEXT turns red (background stays yellow).
//   4. Otherwise -> the normal due-date colour ladder; if that row is
//      overdue or due today, the TEXT turns white for readability against
//      the dark red background.
function computeRowVisuals_(myStatus, isCompletedApi, dueDateStr) {
  const status = String(myStatus || '').trim().toLowerCase();
  const days = daysRemaining_(dueDateStr);
  const isOverdueOrDueToday = days !== null && days <= 0;

  if (isCompletedApi || status === AUTO_STATUS_LABEL.toLowerCase()) {
    return { background: COLOR.COMPLETED_BY_ASSIGNER, font: FONT.WHITE };
  }
  if (status === 'done') {
    return { background: COLOR.COMPLETED, font: FONT.DEFAULT };
  }
  if (status === 'working') {
    return { background: COLOR.WORKING, font: isOverdueOrDueToday ? FONT.RED : FONT.DEFAULT };
  }
  const background = computeRowColor_(false, dueDateStr);
  const font = isOverdueOrDueToday ? FONT.WHITE : FONT.DEFAULT;
  return { background, font };
}

// ============================================================
// 10. CORE SYNC
// ============================================================
function syncNow() {
  const ui = getUiOrNull_();
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(5000);
  if (!gotLock) {
    logSync_('SYNC', '', '', 'SKIPPED', 'Another sync was already running — skipped this run to avoid overlapping writes.');
    return;
  }

  try {
    const c = getConfig_();
    if (!c.myPersonId) {
      safeAlert_(ui, 'Missing MY_BASECAMP_PERSON_ID', 'Run "Find My Basecamp User ID" first.');
      return;
    }

    const sheet = getMainSheet_();
    ensureHeaders_(sheet);
    try {
      normalizeTaskLinkColumns_(sheet); // D = raw URL, O = Open Task; keeps anchored links
    } catch (err) {
      logSync_('NORMALIZE_LINKS', '', '', 'WARN', err.message);
    }
    const rowIndex = buildExistingRowIndex_(sheet);

    let assignments, freshMentionIds;
    try {
      const result = fetchAllRelevantAssignments_(c);
      assignments = result.items;
      freshMentionIds = result.freshMentionIds;
    } catch (err) {
      logSync_('SYNC', '', '', 'FAILED', err.message);
      safeAlert_(ui, 'Sync failed', 'Could not reach Basecamp. Existing sheet data was left untouched.\n\n' + err.message);
      return;
    }

    let created = 0, updated = 0, failed = 0;

    assignments.forEach(item => {
      try {
        upsertAssignmentRow_(sheet, rowIndex, item, c, freshMentionIds);
        if (rowIndex[String(item.id)] && rowIndex.__isNew) created++; else updated++;
      } catch (err) {
        failed++;
        logSync_('SYNC_ITEM', item.id, item.bucket && item.bucket.name, 'FAILED', err.message);
      }
    });

    logSync_('SYNC', '', '', 'OK', 'Processed ' + assignments.length + ' assignment(s), ' + failed + ' failure(s).');

    // Rows that came in via a COMMENT mention (not a direct assignment, and
    // not a description mention) have no way to be re-checked unless a NEW
    // comment/tag happens on them again — a plain "mark as complete" click
    // in Basecamp doesn't create a new comment, so it would otherwise never
    // be noticed. This sweep re-verifies completion/due-date/title directly
    // against Basecamp for whichever rows weren't already touched above,
    // oldest-checked first, under its own time budget.
    try {
      const processedTodoIds = new Set(assignments.map(item => String(item.id)));
      reconcileStaleRows_(sheet, processedTodoIds, c.reconcileBudgetMs);
    } catch (err) {
      logSync_('RECONCILE', '', '', 'FAILED', err.message);
    }

    // Recompute colour/status for EVERY row from its stored Due Date, not
    // just the ones this particular sync touched. Without this, a row that
    // came from the mention-scan (which only re-checks recently-changed
    // Basecamp activity) would keep whatever colour it got at creation
    // forever, even after its due date passes.
    refreshColoursOnly();

    safeAlert_(ui, 'Sync complete', 'Processed ' + assignments.length + ' Basecamp assignment(s).\nFailures: ' + failed +
      '\n\nSee the "Sync Log" sheet for details.');
  } finally {
    lock.releaseLock();
  }
}

// Returns the UI object when run interactively (menu click), or null when
// run from a time-driven trigger — SpreadsheetApp.getUi() throws in that
// context, which would otherwise silently break the 5-minute/hourly sync.
function getUiOrNull_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

function safeAlert_(ui, title, message) {
  if (ui) {
    ui.alert(title, message, ui.ButtonSet.OK);
  } else {
    logSync_('ALERT', '', '', title, message);
  }
}

function ensureHeaders_(sheet) {
  const headers = ['Assigned Date', 'Due Date', 'Project Name', 'Basecamp Task URL (raw)', 'Work Title',
    'Website Link', 'Added By', 'My Status', 'Activity Time', 'Basecamp Todo ID', 'Basecamp Project ID', 'Status',
    'Last Updated', 'Last Synced', 'Task ID'];
  const existing = sheet.getRange(HEADER_ROW, 1, 1, headers.length).getValues()[0];
  const isEmpty = existing.every(v => v === '' || v === null);
  if (isEmpty) {
    sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    // "My Status" (H) stays visible — it's a manual dropdown, not a helper column.
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(MY_STATUS_VALUES, true).setAllowInvalid(false).build();
    sheet.getRange(2, COL.MY_STATUS, 998, 1).setDataValidation(rule);
    // Hide the true helper columns (Basecamp Todo ID onward)
    sheet.hideColumns(COL.BASECAMP_TODO_ID, COL.LAST_SYNCED - COL.BASECAMP_TODO_ID + 1);
  }
}

// Fetches active + recently-completed assignments for the current user
// via Basecamp's "My Assignments" endpoints (not paginated).
function fetchMyAssignments_() {
  const active = bcRequest_('/my/assignments.json').json || {};
  const priorities = active.priorities || [];
  const nonPriorities = active.non_priorities || [];
  let completed = [];
  try {
    completed = bcRequest_('/my/assignments/completed.json').json || [];
  } catch (e) {
    // Non-fatal: if this endpoint is unavailable, active assignments still sync.
    logSync_('FETCH_COMPLETED', '', '', 'WARN', e.message);
  }

  const all = [].concat(priorities, nonPriorities, completed);
  // De-dupe by id (an item could theoretically appear in more than one list)
  const seen = {};
  const result = [];
  all.forEach(item => {
    if (item && item.type === 'Todo' || item && item.type === 'todo' || item) {
      if (!seen[item.id]) {
        seen[item.id] = true;
        result.push(item);
      }
    }
  });
  return result;
}

// ============================================================
// 8b. ACCOUNT-WIDE MENTION SCAN
// ============================================================
// Catches tasks where ANYONE @mentions you on ANY to-do, regardless of who
// it's assigned to (Tushar, Moses, Rohit, Aaron, etc.) — not just a fixed
// "TL" list. Uses Basecamp's cross-project /projects/recordings.json feed
// (documented in bc3-api's recordings.md), sorted newest-first, so each run
// only has to look at what changed since the last successful scan — this is
// what makes a 5-minute refresh practical instead of re-walking the whole
// account every time.
//
// Two feeds are scanned:
//   - type=Comment  -> catches a mention added inside a comment on a to-do.
//                      The comment's own app_url already points straight at
//                      that comment (…/todos/123#__recording_456), which is
//                      exactly the "jump straight to it" link behaviour.
//   - type=Todo     -> catches a mention placed directly in a to-do's own
//                      description (not a comment).
// Both are matched by your attachable_sgid (MY_ATTACHABLE_SGID) — the same
// ID Basecamp embeds in a real @mention, so this cannot false-positive on
// someone merely typing your name.
function fetchMentionsAccountWide_(config) {
  if (!config.mySgid) {
    logSync_('MENTION_SCAN', '', '', 'WARN', 'MY_ATTACHABLE_SGID not set — run "Find My Basecamp User ID" again.');
    return [];
  }

  const startTime = Date.now();
  const timedOut = () => (Date.now() - startTime) > config.mentionScanBudgetMs;
  const checkpointKey = 'MENTION_SCAN_LAST_UPDATED_AT';
  const checkpoint = getProps_().getProperty(checkpointKey) || isoDaysAgo_(30); // first run: look back 30 days
  const scanStartedAt = new Date().toISOString();

  const results = [];
  const seenTodoIds = {};
  let completedFully = true;

  // --- Feed 1: comments, newest first ---
  completedFully = scanRecordingFeed_('Comment', checkpoint, timedOut, (comment) => {
    if (!comment.content || comment.content.indexOf(config.mySgid) === -1) return;
    if (!comment.parent || comment.parent.type !== 'Todo') return; // only care about to-dos
    if (seenTodoIds[comment.parent.id]) return;

    const todoDetail = fetchTodoDetail_(null, comment.parent.id);
    if (!todoDetail) return;
    seenTodoIds[comment.parent.id] = true;
    results.push(buildMentionItem_(todoDetail, comment.app_url, comment.created_at || comment.updated_at));
  }) && completedFully;

  // --- Feed 2: to-dos themselves, newest first (mention in description) ---
  // IMPORTANT: don't trust the raw item from the recordings list for fields
  // like `completed`/`due_on` — always fetch the dedicated to-do detail
  // (same as Feed 1 does for comments) so completion state is authoritative.
  // Using the list item directly was the cause of rows moving to a new
  // position on a fresh mention but keeping the old (e.g. overdue/red)
  // colour instead of turning green when the task was actually completed.
  completedFully = scanRecordingFeed_('Todo', checkpoint, timedOut, (todo) => {
    if (seenTodoIds[todo.id]) return; // already caught via a comment above
    const todoDetail = fetchTodoDetail_(null, todo.id);
    if (!todoDetail) return;
    if (!todoDetail.description || todoDetail.description.indexOf(config.mySgid) === -1) return;
    seenTodoIds[todo.id] = true;
    results.push(buildMentionItem_(todoDetail, todoDetail.app_url, todoDetail.updated_at));
  }) && completedFully;

  if (completedFully) {
    getProps_().setProperty(checkpointKey, scanStartedAt);
  } else {
    logSync_('MENTION_SCAN', '', '', 'WARN', 'Time budget reached — checkpoint not advanced, next sync will pick up where this left off.');
  }

  return results;
}

// Pages through /projects/recordings.json?type=X&sort=updated_at&direction=desc,
// calling onItem(item) for each, and stops as soon as it reaches an item
// older than the checkpoint (everything after that was already scanned).
// Returns true if it finished cleanly (reached the checkpoint or ran out of
// pages) or false if it had to bail out early due to the time budget.
function scanRecordingFeed_(type, checkpointIso, timedOut, onItem) {
  const checkpointTime = new Date(checkpointIso).getTime();
  let url = '/projects/recordings.json?type=' + type + '&sort=updated_at&direction=desc';
  const c = getConfig_();
  const token = getValidAccessToken_();
  let guard = 0;

  while (url && guard < 300) {
    if (timedOut()) return false;
    guard++;
    const fullUrl = url.indexOf('http') === 0 ? url : 'https://3.basecampapi.com/' + c.accountId + url;
    let json, headers;
    try {
      ({ json, headers } = bcFetchWithRetry_(fullUrl, token));
    } catch (e) {
      logSync_('MENTION_SCAN', '', '', 'WARN', type + ' feed: ' + e.message);
      return true; // don't block the checkpoint on a transient page failure
    }
    for (const item of json) {
      if (new Date(item.updated_at).getTime() <= checkpointTime) {
        return true; // reached already-scanned territory
      }
      try { onItem(item); } catch (e) {
        logSync_('MENTION_SCAN_ITEM', item.id, '', 'WARN', e.message);
      }
    }
    url = null;
    const link = headers['Link'] || headers['link'];
    if (link) {
      const match = link.split(',').map(s => s.trim()).find(s => s.endsWith('rel="next"'));
      if (match) {
        const m = match.match(/<([^>]+)>/);
        if (m) url = m[1];
      }
    }
  }
  return true; // ran out of pages, nothing older to check
}

function buildMentionItem_(todoDetail, linkUrl, mentionAt) {
  return {
    id: todoDetail.id,
    app_url: linkUrl || todoDetail.app_url,
    content: todoDetail.content,
    due_on: todoDetail.due_on,
    bucket: todoDetail.bucket,
    completed: !!todoDetail.completed,
    assignees: todoDetail.assignees || [],
    origin: 'mention',
    mentionAt: mentionAt || null, // exact timestamp of the comment/mention that surfaced this
    // Link straight to the comment that mentioned you (only comments have the
    // "#__recording_..." anchor; a mention in the description has none).
    mentionUrl: (linkUrl && linkUrl.indexOf('#__recording_') !== -1) ? linkUrl : null,
    _detail: todoDetail
  };
}

function isoDaysAgo_(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  return d.toISOString();
}

// Combines direct assignments with account-wide "mentioned" tasks, de-duped
// by to-do ID (a direct assignment always wins if a to-do somehow shows up
// in both — that just means you're literally assigned to it too). Also
// returns freshMentionIds: the set of to-do IDs that had a genuinely NEW
// @mention detected THIS run (regardless of which item "won" the merge) —
// used to decide which existing rows should move to the bottom of the sheet.
function fetchAllRelevantAssignments_(config) {
  const direct = fetchMyAssignments_();
  direct.forEach(item => { item.origin = 'direct'; });

  let viaMention = [];
  try {
    viaMention = fetchMentionsAccountWide_(config);
  } catch (e) {
    logSync_('MENTION_SCAN', '', '', 'WARN', e.message);
  }
  const freshMentionIds = new Set(viaMention.map(item => String(item.id)));

  // A direct assignment wins for the row's live fields, but it must still
  // carry the mention's time AND its comment link — otherwise a task that is
  // both assigned to you and re-mentioned would lose the "#__recording_" anchor.
  const merged = {};
  direct.concat(viaMention).forEach(item => {
    const key = String(item.id);
    const existing = merged[key];
    if (!existing) {
      merged[key] = item;
    } else if (item.origin === 'direct') {
      item.mentionAt = existing.mentionAt || item.mentionAt;
      item.mentionUrl = existing.mentionUrl || item.mentionUrl;
      merged[key] = item;
    } else {
      existing.mentionAt = existing.mentionAt || item.mentionAt;
      existing.mentionUrl = existing.mentionUrl || item.mentionUrl;
    }
  });
  return { items: Object.keys(merged).map(k => merged[k]), freshMentionIds };
}

function upsertAssignmentRow_(sheet, rowIndex, item, config, freshMentionIds) {
  const todoId = String(item.id);
  const bucketName = item.bucket ? item.bucket.name : '';
  const bucketId = item.bucket ? item.bucket.id : '';
  const projectDisplayName = deriveProjectDisplayName_(bucketName);
  const workTitle = item.content || '';
  const dueOn = item.due_on || '';
  const isCompleted = !!item.completed;
  const taskUrl = item.mentionUrl || item.app_url || '';
  const originSuffix = item.origin === 'mention' ? ' (Mentioned)' : '';

  let existingRow = rowIndex[todoId];
  const isFreshMention = freshMentionIds && freshMentionIds.has(todoId);
  let isMove = false;
  let preservedMyStatus = null;

  if (existingRow && isFreshMention) {
    // A NEW @mention/comment came in on a task already in the sheet — per
    // your request, this doesn't update quietly in place: the old row is
    // removed and the task is re-added at the bottom, so recently-tagged
    // tasks bubble to where you'll notice them. Your manual "My Status"
    // (Done/Pending/Working) is carried over, not reset.
    preservedMyStatus = sheet.getRange(existingRow, COL.MY_STATUS).getValue();
    // A re-mention means the assigner wants more work on it, so a manual
    // "Done" must NOT stick — clear it so the row goes back to normal
    // due-date colouring (overdue / due today / etc.). "Working" and
    // "Pending" are still carried over. The auto "DONE By Assigner" label
    // is only kept if Basecamp still has the task completed.
    const preservedNorm = String(preservedMyStatus || '').trim().toLowerCase();
    if (preservedNorm === 'done') {
      preservedMyStatus = '';
    } else if (preservedNorm === AUTO_STATUS_LABEL.toLowerCase()) {
      preservedMyStatus = isCompleted ? AUTO_STATUS_LABEL : '';
    }
    sheet.deleteRow(existingRow);
    delete rowIndex[todoId];
    Object.keys(rowIndex).forEach(key => {
      if (key !== '__isNew' && rowIndex[key] > existingRow) rowIndex[key] = rowIndex[key] - 1;
    });
    existingRow = null;
    isMove = true;
  }

  rowIndex.__isNew = !existingRow;

  if (existingRow) {
    updateExistingRow_(sheet, existingRow, {
      dueOn, workTitle, isCompleted, projectDisplayName, originSuffix, taskUrl
    });
  } else {
    // Mention-origin items already carry the full to-do detail (fetched
    // during the scan itself) — reuse it instead of calling the API again.
    const detail = item._detail || fetchTodoDetail_(bucketId, item.id);
    const assignedDateStr = detail && detail.created_at ? detail.created_at.substring(0, 10) : '';
    const description = detail ? stripHtml_(detail.description || '') : '';
    const websiteLink = resolveWebsiteLink_(description, bucketName, config.extractWebsiteFromProjectName);
    const addedByName = (detail && detail.creator && detail.creator.name) || '';

    // Activity Time: the moment you were tagged (if this is mention-driven),
    // otherwise the moment the to-do was originally created/assigned.
    const activityTimestamp = item.mentionAt || (detail && detail.created_at) || null;
    const activityTimeStr = activityTimestamp ? formatDateTimeAU_(activityTimestamp) : '';
    const initialMyStatus = (preservedMyStatus !== null && preservedMyStatus !== undefined)
      ? preservedMyStatus
      : (isCompleted ? AUTO_STATUS_LABEL : '');

    const row = sheet.getLastRow() + 1;
    writeNewRow_(sheet, row, {
      assignedDateStr, dueOn, projectDisplayName, taskUrl, workTitle,
      websiteLink, addedByName, todoId, bucketId, isCompleted, originSuffix,
      activityTimeStr, initialMyStatus, isMove
    });
    rowIndex[todoId] = row;
  }
}

function uniqueNames_(names) {
  const seen = {};
  const out = [];
  (names || []).forEach(n => {
    if (n && !seen[n]) { seen[n] = true; out.push(n); }
  });
  return out.join(', ');
}

// Fetches full to-do detail (needed once, at creation time, for
// created_at and description/notes). Cached per-execution to avoid
// duplicate calls if the same todo appears twice in one sync.
const _todoDetailCache = {};
function fetchTodoDetail_(bucketId, todoId) {
  const key = (bucketId || 'flat') + ':' + todoId;
  if (_todoDetailCache[key]) return _todoDetailCache[key];
  try {
    const path = bucketId ? ('/buckets/' + bucketId + '/todos/' + todoId + '.json') : ('/todos/' + todoId + '.json');
    const { json } = bcRequest_(path);
    _todoDetailCache[key] = json;
    return json;
  } catch (e) {
    logSync_('FETCH_DETAIL', todoId, '', 'WARN', e.message);
    return null;
  }
}

function writeNewRow_(sheet, row, d) {
  // Brand-new rows never have a manual "My Status" set yet — unless the
  // to-do was already completed in Basecamp before we ever imported it
  // (mark "DONE By Assigner"), or this is a re-mention move that carries
  // over whatever you'd already set manually.
  const initialMyStatus = (typeof d.initialMyStatus === 'string')
    ? d.initialMyStatus
    : (d.isCompleted ? AUTO_STATUS_LABEL : '');
  const visuals = computeRowVisuals_(initialMyStatus, d.isCompleted, d.dueOn);
  const status = computeStatusLabel_(d.isCompleted, d.dueOn) + (d.originSuffix || '');
  const now = new Date();

  sheet.getRange(row, COL.ASSIGNED_DATE).setValue(d.assignedDateStr ? formatDateAU_(d.assignedDateStr) : '');
  sheet.getRange(row, COL.DUE_DATE).setValue(d.dueOn ? formatDateAU_(d.dueOn) : '');
  sheet.getRange(row, COL.PROJECT_NAME).setValue(d.projectDisplayName);
  sheet.getRange(row, COL.TASK_ID).setFormula('=HYPERLINK("' + d.taskUrl + '","Open Task")');
  sheet.getRange(row, COL.WORK_TITLE).setValue(d.workTitle);
  sheet.getRange(row, COL.WEBSITE_LINK).setValue(d.websiteLink);
  sheet.getRange(row, COL.ASSIGNEE_NAME).setValue(d.addedByName); // column G = "Added By" (creator)
  sheet.getRange(row, COL.MY_STATUS).setValue(initialMyStatus);
  sheet.getRange(row, COL.ACTIVITY_TIME).setValue(d.activityTimeStr || '');
  sheet.getRange(row, COL.BASECAMP_TODO_ID).setValue(d.todoId);
  sheet.getRange(row, COL.BASECAMP_PROJECT_ID).setValue(d.bucketId);
  sheet.getRange(row, COL.STATUS).setValue(status);
  sheet.getRange(row, COL.LAST_UPDATED).setValue(now);
  sheet.getRange(row, COL.LAST_SYNCED).setValue(now);
  sheet.getRange(row, COL.BASECAMP_URL_RAW).setValue(d.taskUrl);

  sheet.getRange(row, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setBackground(visuals.background);
  sheet.getRange(row, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setFontColor(visuals.font);
  logSync_(d.isMove ? 'RE-MENTION_MOVE' : 'CREATE', d.todoId, d.projectDisplayName, 'OK', '');
}

function updateExistingRow_(sheet, row, d) {
  const current = sheet.getRange(row, 1, 1, COL_COUNT).getValues()[0];
  const currentDue = current[COL.DUE_DATE - 1];
  const currentTitle = current[COL.WORK_TITLE - 1];

  const newDueFormatted = d.dueOn ? formatDateAU_(d.dueOn) : '';
  let changed = false;

  if (newDueFormatted !== currentDue) { sheet.getRange(row, COL.DUE_DATE).setValue(newDueFormatted); changed = true; }
  if (d.workTitle !== currentTitle) { sheet.getRange(row, COL.WORK_TITLE).setValue(d.workTitle); changed = true; }
  if (d.projectDisplayName) { sheet.getRange(row, COL.PROJECT_NAME).setValue(d.projectDisplayName); }
  // Task ID link: refresh it too, in case a later mention gives us a more
  // specific (comment-anchored) link than the plain task link we had before.
  // Keep the "#__recording_..." comment link once a row has one. Syncs and the
  // reconcile sweep only know the plain task URL, so without this guard they
  // would overwrite the anchored link on the very next run.
  const currentUrl = String(current[COL.BASECAMP_URL_RAW - 1] || '').trim();
  const urlHasAnchor = function (u) { return u.indexOf('#__recording_') !== -1; };
  let urlToUse = currentUrl;
  if (d.taskUrl) {
    const currentIsUrl = /^https?:\/\//i.test(currentUrl);
    if (!currentIsUrl || urlHasAnchor(d.taskUrl) || !urlHasAnchor(currentUrl)) urlToUse = d.taskUrl;
  }
  if (urlToUse && (urlToUse !== currentUrl || String(current[COL.TASK_ID - 1]) !== 'Open Task')) {
    sheet.getRange(row, COL.BASECAMP_URL_RAW).setValue(urlToUse);
    sheet.getRange(row, COL.TASK_ID).setFormula('=HYPERLINK("' + urlToUse + '","Open Task")');
  }
  // Added By (creator) never changes after a to-do is created, so column G
  // is intentionally left untouched here — only set once, at creation.

  const status = computeStatusLabel_(d.isCompleted, d.dueOn) + (d.originSuffix || '');
  const currentStatus = sheet.getRange(row, COL.STATUS).getValue();
  const wasCompleted = String(currentStatus).indexOf('Completed') === 0;
  if (status !== currentStatus) { sheet.getRange(row, COL.STATUS).setValue(status); changed = true; }

  let myStatus = current[COL.MY_STATUS - 1];
  const myStatusNorm = String(myStatus || '').trim().toLowerCase();
  if (d.isCompleted && myStatusNorm !== AUTO_STATUS_LABEL.toLowerCase()) {
    myStatus = AUTO_STATUS_LABEL;
    sheet.getRange(row, COL.MY_STATUS).setValue(myStatus);
    changed = true;
  } else if (!d.isCompleted && myStatusNorm === AUTO_STATUS_LABEL.toLowerCase()) {
    // Basecamp task was reopened — clear our own auto-set label. A manual
    // "Done"/"Pending"/"Working" you set yourself is never touched here.
    myStatus = '';
    sheet.getRange(row, COL.MY_STATUS).setValue(myStatus);
    changed = true;
  }

  const visuals = computeRowVisuals_(myStatus, d.isCompleted, d.dueOn);
  sheet.getRange(row, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setBackground(visuals.background);
  sheet.getRange(row, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setFontColor(visuals.font);

  sheet.getRange(row, COL.LAST_SYNCED).setValue(new Date());
  if (changed || wasCompleted !== d.isCompleted) {
    sheet.getRange(row, COL.LAST_UPDATED).setValue(new Date());
    logSync_('UPDATE', current[COL.BASECAMP_TODO_ID - 1], d.projectDisplayName, 'OK', '');
  }
}

// Fires automatically whenever any cell is edited. If the edit was to the
// "My Status" dropdown (column H), recolour just that row immediately —
// no need to wait for the next 5-minute sync.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() === SYNC_LOG_SHEET_NAME) return;
    if (e.range.getColumn() !== COL.MY_STATUS || e.range.getRow() <= HEADER_ROW) return;

    const row = e.range.getRow();
    const rowValues = sheet.getRange(row, 1, 1, COL.STATUS).getValues()[0];
    const todoId = rowValues[COL.BASECAMP_TODO_ID - 1];
    if (!todoId) return; // manual/unrelated row — leave it alone

    const dueDateDisplay = rowValues[COL.DUE_DATE - 1];
    const status = rowValues[COL.STATUS - 1];
    const isCompletedApi = String(status).indexOf('Completed') === 0;
    const dueIso = auDateToIso_(dueDateDisplay);
    const myStatus = e.range.getValue();

    const visuals = computeRowVisuals_(myStatus, isCompletedApi, dueIso);
    sheet.getRange(row, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setBackground(visuals.background);
    sheet.getRange(row, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setFontColor(visuals.font);
  } catch (err) {
    // onEdit must never throw back at the user — swallow and log only.
    logSync_('ON_EDIT', '', '', 'WARN', err.message);
  }
}

// Re-verifies completion/due date/title directly against Basecamp for rows
// that this sync's direct-assignment + mention-scan didn't already touch —
// this is what catches "task was marked complete in Basecamp with no new
// comment/tag" for tasks that only ever entered the sheet via a comment
// mention. Oldest-last-checked rows first, so every row gets reconciled
// over time even on a large sheet; bounded by a time budget so it never
// blocks the rest of the sync.
function reconcileStaleRows_(sheet, processedTodoIds, timeBudgetMs) {
  const startTime = Date.now();
  const lastRow = sheet.getLastRow();
  if (lastRow < HEADER_ROW + 1) return;

  const numRows = lastRow - HEADER_ROW;
  const data = sheet.getRange(HEADER_ROW + 1, 1, numRows, COL_COUNT).getValues();

  const candidates = [];
  for (let i = 0; i < data.length; i++) {
    const todoId = data[i][COL.BASECAMP_TODO_ID - 1];
    if (!todoId) continue; // manual/unrelated row
    if (processedTodoIds.has(String(todoId))) continue; // already fresh this run
    const statusText = String(data[i][COL.STATUS - 1] || '');
    if (statusText.indexOf('Completed') === 0) continue; // already known complete
    candidates.push({
      rowNum: HEADER_ROW + 1 + i,
      todoId: String(todoId),
      lastUpdated: data[i][COL.LAST_UPDATED - 1]
    });
  }
  candidates.sort((a, b) => new Date(a.lastUpdated || 0) - new Date(b.lastUpdated || 0));

  let checked = 0;
  for (const cand of candidates) {
    if (Date.now() - startTime > timeBudgetMs) {
      logSync_('RECONCILE', '', '', 'WARN',
        'Time budget reached after checking ' + checked + ' row(s) — the rest will be checked on a later sync.');
      return;
    }
    const detail = fetchTodoDetail_(null, cand.todoId);
    checked++;
    if (!detail) continue;

    const currentRow = sheet.getRange(cand.rowNum, 1, 1, COL_COUNT).getValues()[0];
    const currentStatusText = String(currentRow[COL.STATUS - 1] || '');
    const originSuffix = currentStatusText.indexOf('(Mentioned)') !== -1 ? ' (Mentioned)' : '';
    const bucketName = detail.bucket ? detail.bucket.name : '';

    updateExistingRow_(sheet, cand.rowNum, {
      dueOn: detail.due_on || '',
      workTitle: detail.content || currentRow[COL.WORK_TITLE - 1],
      isCompleted: !!detail.completed,
      projectDisplayName: deriveProjectDisplayName_(bucketName),
      originSuffix,
      taskUrl: detail.app_url || currentRow[COL.BASECAMP_URL_RAW - 1]
    });
  }
}

// ============================================================
// 11. COLOUR-ONLY REFRESH (no API calls — cheap, can run often)
// ============================================================
function refreshColoursOnly() {
  const sheet = getMainSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < HEADER_ROW + 1) return;

  const numRows = lastRow - HEADER_ROW;
  const data = sheet.getRange(HEADER_ROW + 1, 1, numRows, COL.BASECAMP_TODO_ID).getValues();

  for (let i = 0; i < data.length; i++) {
    const rowNum = HEADER_ROW + 1 + i;
    const todoId = data[i][COL.BASECAMP_TODO_ID - 1];
    if (!todoId) continue; // skip manual/unrelated rows

    const dueDateDisplay = data[i][COL.DUE_DATE - 1];
    let myStatus = data[i][COL.MY_STATUS - 1];
    const status = sheet.getRange(rowNum, COL.STATUS).getValue();
    const isCompletedApi = String(status).indexOf('Completed') === 0;
    const dueIso = auDateToIso_(dueDateDisplay);

    const myStatusNorm = String(myStatus || '').trim().toLowerCase();
    if (isCompletedApi && myStatusNorm !== AUTO_STATUS_LABEL.toLowerCase()) {
      myStatus = AUTO_STATUS_LABEL;
      sheet.getRange(rowNum, COL.MY_STATUS).setValue(myStatus);
    } else if (!isCompletedApi && myStatusNorm === AUTO_STATUS_LABEL.toLowerCase()) {
      myStatus = '';
      sheet.getRange(rowNum, COL.MY_STATUS).setValue(myStatus);
    }

    const visuals = computeRowVisuals_(myStatus, isCompletedApi, dueIso);
    sheet.getRange(rowNum, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setBackground(visuals.background);
    sheet.getRange(rowNum, COL.ASSIGNED_DATE, 1, COL.ASSIGNEE_NAME).setFontColor(visuals.font);

    if (!isCompletedApi) {
      const originSuffix = String(status).indexOf('(Mentioned)') !== -1 ? ' (Mentioned)' : '';
      const newStatus = computeStatusLabel_(false, dueIso) + originSuffix;
      if (newStatus !== status) sheet.getRange(rowNum, COL.STATUS).setValue(newStatus);
    }
  }
}

function auDateToIso_(auDateStr) {
  if (!auDateStr) return '';
  if (auDateStr instanceof Date) {
    return Utilities.formatDate(auDateStr, TIMEZONE, 'yyyy-MM-dd');
  }
  const parts = String(auDateStr).split('/');
  if (parts.length !== 3) return '';
  return parts[2] + '-' + parts[1] + '-' + parts[0];
}
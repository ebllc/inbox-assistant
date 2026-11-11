/**
 * Inbox Assistant — CORE. This file is downloaded and run by Loader.gs (the small file in each person's project).
 * Don't paste this file into Apps Script; paste Loader.gs. Edit App.html, then run build.py.
 *
 * Inbox Assistant — a personal, free MailSynth-style Gmail helper.
 *
 * Runs entirely inside your Google account (Apps Script + Gemini free tier).
 * Rules are plain English in a Google Sheet. The script:
 *   - reads new inbox mail every few minutes
 *   - asks Gemini what to do with each message, based on your rules
 *   - labels / stars / archives (now or after N days) / alerts
 *   - emails you a morning digest
 *
 * One-time: run setup() from the Apps Script editor, click Allow,
 * then paste your Gemini API key into the Settings tab of the sheet it creates.
 */

// ---------- constants ----------
const SHEET_NAME = 'Inbox Assistant';
const LABEL_ROOT = 'AI';
const LABEL_PROCESSED = LABEL_ROOT + '/Processed';
const LABEL_PREVIEWED = LABEL_ROOT + '/Previewed'; // used only while DRY_RUN is on
const LABEL_ALERT = LABEL_ROOT + '/🔔 Alert';
const LABEL_IMPORTANT = LABEL_ROOT + '/⭐ Needs me';
const PROP_SHEET_ID = 'SHEET_ID';

// ---------- release info ----------
const CODE_VERSION = '2026.11.11-6';
const REPO = 'ebllc/inbox-assistant'; // public GitHub repo the loader downloads this file from

const DEFAULT_SETTINGS = [
  ['GEMINI_API_KEY', '', 'Paste your key from https://aistudio.google.com/apikey'],
  ['MODEL', 'gemini-3.8-flash', 'Gemini model to use (free tier)'],
  ['DRY_RUN', 'TRUE', 'TRUE = only log what it WOULD do. Set FALSE to go live.'],
  ['POLL_MINUTES', '10', 'How often to check the inbox (5, 10, 15, 30)'],
  ['DIGEST_HOUR', '7', 'Hour(s) (0-23) to send a digest. Several = comma-separated, e.g. 7,12. Changes apply within 10 min.'],
  ['ALERTS_IMMEDIATE', 'TRUE', 'Email me right away when a rule says "alert"'],
  ['MAX_PER_RUN', '30', 'Max emails to process each run'],
  ['LOOKBACK_DAYS', '2', 'Only look at inbox mail newer than this many days'],
  ['LAST_DIGEST', '', '(auto) when the last digest was sent'],
  ['LAST_CHECK', '', '(auto) what happened at the last inbox check'],
  ['ARCHIVE_READ_AFTER_DAYS', '3', 'Archive emails you\'ve read once they\'re this many days old (stars and "Needs me" are kept). Blank or 0 = off.'],
  ['UPDATE_CODE', 'FALSE', 'Set TRUE to install the latest version now. Flips back to FALSE when done.'],
  ['UPDATE_CHANNEL', 'main', 'main = stable releases (recommended). beta = newest changes first.'],
  ['AUTO_UPDATE', 'FALSE', 'TRUE = check for and install new versions once a day automatically.'],
  ['LAST_UPDATE', '', '(auto) result of the last update attempt'],
  ['SEND_DIGEST', 'TRUE', 'FALSE = this account sends no digest of its own (use when another account includes it)'],
  ['SEND_ALERTS_TO', '', '🔔 alerts go to this address. Blank = this account.'],
  ['INCLUDE_ACCOUNTS', '', 'Main account only: links to your other accounts\' Inbox Assistant sheets, comma-separated. Their mail joins this digest.'],
  ['ACCOUNT', '', '(auto) this account\'s email address'],
  ['ACCOUNT_LABEL', '', 'Short name for this account in the app and digest, e.g. Personal or Work.'],
  ['TAP_ACTION', 'app', 'When you tap an email in the app: app = show it in the app, gmail = open Gmail.'],
  ['MAIN_ACCOUNT', '', '(auto) the main account this one is linked to'],
  ['APP_URL', '', '(auto) your Inbox Assistant app link — open it on your phone and Add to Home Screen'],
];

const DEFAULT_RULES = [
  ['Calendar invites, 2FA codes, password resets, and security notices: leave alone.', 'TRUE'],
  ['If an email is clearly waiting on a reply or action from me, keep it in the inbox and mark it "Needs me".', 'TRUE'],
  ['Archive receipts, order confirmations, and shipping updates after 1 day.', 'TRUE'],
  ['Label newsletters "Newsletter" and archive them right away; summarize them in my digest.', 'TRUE'],
  ['Archive promotions, sales, and marketing emails immediately. No need to mention them in the digest.', 'TRUE'],
];

/** Check the inbox right now. (First on purpose: it's the default in the Run menu.) */
function runNow() {
  const msg = processInbox();
  try { if (boundSheet_()) SpreadsheetApp.getActiveSpreadsheet().toast(msg || 'Done.', '📬 Inbox Assistant', 8); } catch (e) {}
}

// ---------- setup ----------
function setup() {
  let ss = getSheet_(true);
  ensureTabs_(ss);
  ensureLabels_();
  installTriggers_();
  try { syncSettingsRows_(); ensureHowTo_(); ensureRuleLibrary_(); styleTabs_(); } catch (e) { Logger.log('How-to/styling: ' + e); }
  const url = ss.getUrl();
  if (boundSheet_()) {
    try { ss.setActiveSheet(ss.getSheetByName('Settings')); } catch (e) {}
    try {
      SpreadsheetApp.getUi().alert('Inbox Assistant is set up ✅',
        'Last step: in the Settings tab (open now), paste your Gemini key into the GEMINI_API_KEY row.\n\n' +
        'It starts in test mode — it only writes in the Log tab what it would do. When you like what you see, set DRY_RUN to FALSE.',
        SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) { /* not opened from the menu */ }
  }
  GmailApp.sendEmail(
    Session.getActiveUser().getEmail(),
    'Inbox Assistant is set up ✅',
    'Your rules & settings sheet: ' + url +
      '\n\nNext: open the Settings tab and paste your Gemini API key.' +
      '\nIt starts in DRY_RUN mode (logs only). Flip DRY_RUN to FALSE when you like what you see in the Log tab.'
  );
  Logger.log('Sheet: ' + url);
}

/** The sheet this script lives in, when it was installed from the template (bound script). */
function boundSheet_() {
  if (viaLibrary_()) return null;
  try { return SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { return null; }
}

function getSheet_(createIfMissing) {
  const props = PS_().getUserProperties();
  const bound = boundSheet_();
  if (bound) {
    if (props.getProperty(PROP_SHEET_ID) !== bound.getId()) props.setProperty(PROP_SHEET_ID, bound.getId());
    return bound;
  }
  const id = props.getProperty(PROP_SHEET_ID);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* fall through */ }
  }
  if (!createIfMissing) throw new Error('Run setup() first.');
  const ss = SpreadsheetApp.create(SHEET_NAME);
  props.setProperty(PROP_SHEET_ID, ss.getId());
  return ss;
}

function ensureTabs_(ss) {
  // Rules
  let rules = ss.getSheetByName('Rules');
  if (!rules) {
    rules = ss.getSheets()[0];
    rules.setName('Rules');
    rules.getRange(1, 1, 1, 2).setValues([['Rule (plain English)', 'On?']]).setFontWeight('bold');
    rules.getRange(2, 1, DEFAULT_RULES.length, 2).setValues(DEFAULT_RULES);
    rules.setColumnWidth(1, 620);
    rules.getRange('A2:A200').setWrap(true);
    rules.getRange('B2:B200').insertCheckboxes();
    rules.setFrozenRows(1);
  }
  // Settings
  let settings = ss.getSheetByName('Settings');
  if (!settings) {
    settings = ss.insertSheet('Settings');
    settings.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it does']]).setFontWeight('bold');
    settings.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    settings.setColumnWidth(1, 160); settings.setColumnWidth(2, 260); settings.setColumnWidth(3, 420);
    settings.setFrozenRows(1);
  }
  // Log
  let log = ss.getSheetByName('Log');
  if (!log) {
    log = ss.insertSheet('Log');
    log.getRange(1, 1, 1, 8).setValues([['When', 'From', 'Subject', 'Category', 'Action', 'Summary', 'Mode', 'Thread']]).setFontWeight('bold');
    log.setColumnWidth(2, 180); log.setColumnWidth(3, 280); log.setColumnWidth(5, 220); log.setColumnWidth(6, 360);
    log.setFrozenRows(1);
  }
  // Pending (deferred archives)
  let pending = ss.getSheetByName('Pending');
  if (!pending) {
    pending = ss.insertSheet('Pending');
    pending.getRange(1, 1, 1, 4).setValues([['Thread ID', 'Subject', 'Action', 'Due']]).setFontWeight('bold');
    pending.setFrozenRows(1);
  }
}

function ensureLabels_() {
  [LABEL_ROOT, LABEL_PROCESSED, LABEL_ALERT, LABEL_IMPORTANT].forEach(getOrCreateLabel_);
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function installTriggers_() {
  SA_().getProjectTriggers().forEach(t => SA_().deleteTrigger(t));
  const s = readSettings_();
  PS_().getUserProperties().setProperty('TRIGGER_SIG', triggerSig_(s));
  const poll = [5, 10, 15, 30].includes(Number(s.POLL_MINUTES)) ? Number(s.POLL_MINUTES) : 10;
  SA_().newTrigger('processInbox').timeBased().everyMinutes(poll).create();
  SA_().newTrigger('processPending').timeBased().everyHours(1).create();
  digestHours_(s).forEach(h => SA_().newTrigger('sendDigest').timeBased().everyDays(1).atHour(h).create());
  // Instant reactions to sheet edits (Rule Library ticks, UPDATE_CODE). Full copies only; library accounts poll instead.
  if (!viaLibrary_()) SA_().newTrigger('onSheetEdit').forSpreadsheet(getSheet_(false).getId()).onEdit().create();
}

/** Schedule fingerprint: when POLL_MINUTES or DIGEST_HOUR change in the sheet, triggers rebuild themselves. */
function triggerSig_(s) { return 'v2|' + String(s.POLL_MINUTES) + '|' + digestHours_(s).join(',') + '|' + (viaLibrary_() ? 'lib' : 'full'); }

function syncTriggers_() {
  const s = readSettings_();
  if (PS_().getUserProperties().getProperty('TRIGGER_SIG') !== triggerSig_(s)) installTriggers_();
}

/** DIGEST_HOUR like "7" or "7,12" → [7, 12]. */
function digestHours_(s) {
  const hrs = String(s.DIGEST_HOUR || '7').split(/[,;\s]+/).map(Number)
    .filter(h => Number.isInteger(h) && h >= 0 && h <= 23);
  return hrs.length ? [...new Set(hrs)].sort((a, b) => a - b) : [7];
}

/** Which digest this is, by time of day. */
function digestKind_() {
  const h = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'H'));
  return h < 11 ? { icon: '☕', name: 'Morning' } : h < 17 ? { icon: '🥪', name: 'Midday' } : { icon: '🌙', name: 'Evening' };
}

/** Re-run this after changing POLL_MINUTES or DIGEST_HOUR. */
function reinstallTriggers() { installTriggers_(); }

// ---------- settings & rules ----------
function readSettings_() {
  const out = {};
  DEFAULT_SETTINGS.forEach(r => out[r[0]] = r[1]);
  try {
    const sh = getSheet_(false).getSheetByName('Settings');
    sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues()
      .forEach(([k, v]) => { if (k) out[String(k).trim()] = String(v).trim(); });
  } catch (e) { /* defaults */ }
  out.DRY_RUN = String(out.DRY_RUN).toUpperCase() !== 'FALSE';
  out.ALERTS_IMMEDIATE = String(out.ALERTS_IMMEDIATE).toUpperCase() !== 'FALSE';
  out.SEND_DIGEST = String(out.SEND_DIGEST).toUpperCase() !== 'FALSE';
  out.UPDATE_CODE = String(out.UPDATE_CODE).toUpperCase() === 'TRUE';
  out.AUTO_UPDATE = String(out.AUTO_UPDATE).toUpperCase() === 'TRUE';
  return out;
}

/** Adds any settings rows that are new since this sheet was created, and records this account's address. */
function syncSettingsRows_() {
  const sh = getSheet_(false).getSheetByName('Settings');
  const keys = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim());
  DEFAULT_SETTINGS.forEach(row => { if (keys.indexOf(row[0]) < 0) sh.appendRow(row); });
  const me = Session.getActiveUser().getEmail();
  if (me && readSettings_().ACCOUNT !== me) writeSetting_('ACCOUNT', me);
  try { ensureHowTo_(); } catch (e) { Logger.log('How-to tab: ' + e); }
  try { ensureRuleLibrary_(); } catch (e) { Logger.log('Rule Library: ' + e); }
  try { styleTabs_(); } catch (e) { Logger.log('Styling: ' + e); }
}

// ---------- Rule Library ----------
// A tab of ready-made rules by account type. Ticked = the rule is in the Rules tab; unticking removes it.
// Rules you type yourself in the Rules tab are never touched.
const LIBRARY_VERSION = '2026-11-11a';
const RULE_LIBRARY = [
  ['Everyone', [
    'Calendar invites, 2FA codes, password resets, and security notices: leave alone.',
    'If an email is clearly waiting on a reply or action from me, keep it in the inbox and mark it "Needs me".',
    'Archive receipts, order confirmations, and shipping updates after 1 day.',
    'Label newsletters "Newsletter" and archive them right away; summarize them in my digest.',
    'Archive promotions, sales, and marketing emails immediately. No need to mention them in the digest.',
  ]],
  ['Personal', [
    'Emails from family and friends (real people writing to me personally): keep in the inbox and mark "Needs me" if they ask something.',
    'Label bills, statements, and payment reminders "Bills" and mark "Needs me" if a payment is due.',
    'Label travel bookings, boarding passes, and reservations "Travel" and keep them in the inbox.',
    'Label social media notifications "Social" and archive them right away.',
    'Alert me immediately about anything from my bank or credit card about fraud or a suspicious charge.',
  ]],
  ['Work', [
    'Alert me immediately about anything from my manager that mentions a deadline, a problem, or "urgent".',
    'Label meeting recaps, notes, and transcripts "Meetings" and archive them; list their action items in my digest.',
    'Label company-wide announcements "Company" and archive them after 2 days.',
    'Label tool notifications (GitHub, Jira, Slack, Google Drive shares) "Tools" and archive right away, unless it\'s an error or security issue — then alert me.',
    'Label HR, payroll, and benefits emails "HR" and keep them in the inbox.',
  ]],
  ['Freelance / consulting', [
    'Alert me immediately about anything from a client, especially about a deadline, a problem, scope, or budget.',
    'Label invoices, contracts, proposals, SOWs, and signature requests "Admin", keep in the inbox, and mark "Needs me".',
    'Label new inquiries and leads "Leads" and mark "Needs me".',
    'Label payment confirmations and payouts "Payments" and archive them.',
    'Archive cold sales pitches and agency outreach right away.',
  ]],
  ['Job search', [
    'Alert me immediately about any email from a recruiter or hiring manager writing to me personally, or about an interview.',
    'Label job alerts "Jobs", archive them, and list the roles in my digest.',
    'Label application confirmations and rejections "Applications" and archive them.',
    'Label interview scheduling emails "Interviews", keep in the inbox, and mark "Needs me".',
  ]],
  ['Parent / household', [
    'Label anything from school, teachers, or the district "School" and keep it in the inbox; alert me about closures, early dismissals, or emergencies.',
    'Label sports, activities, and camp emails "Activities" and keep them in the inbox.',
    'Label doctor, dentist, and vet appointment reminders "Appointments" and keep them in the inbox.',
    'Label utility, insurance, and household bills "Bills" and mark "Needs me" if payment is due.',
    'Archive package delivery notices after 1 day.',
  ]],
  ['Shopping & deals', [
    'Label returns, refunds, and order problems "Orders" and keep them in the inbox.',
    'Label price-drop and back-in-stock alerts "Deals" and keep them in the inbox for 2 days, then archive.',
    'Archive store loyalty and rewards emails right away.',
  ]],
  ['Small business / seller', [
    'Label new orders "Orders" and list them in my digest.',
    'Alert me immediately about customer messages, complaints, or order problems.',
    'Label payouts and sales summaries "Payouts" and archive them.',
    'Label platform policy or account notices (Etsy, Shopify, Amazon) "Platform" and mark "Needs me".',
    'Label reviews "Reviews" and archive them; summarize them in my digest.',
  ]],
  ['Student', [
    'Label emails from professors, TAs, and course systems "Classes" and mark "Needs me" if they ask something.',
    'Alert me immediately about assignment deadline changes, exam changes, or class cancellations.',
    'Label financial aid, tuition, and registrar emails "School Admin" and mark "Needs me".',
    'Label campus announcements and club emails "Campus" and archive them after 2 days.',
  ]],
];

function ensureRuleLibrary_() {
  const ss = getSheet_(false);
  const props = PS_().getUserProperties();
  let sh = ss.getSheetByName('Rule Library');
  if (sh && props.getProperty('LIBRARY_STAMP') === LIBRARY_VERSION) return;
  if (!sh) sh = ss.insertSheet('Rule Library');
  else { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); sh.clear(); sh.clearConditionalFormatRules(); }
  const rows = [['Use?', 'Ready-made rule — tick to add it to your Rules tab, untick to remove it']];
  const kinds = ['h'];
  RULE_LIBRARY.forEach(([type, list]) => {
    rows.push(['', type.toUpperCase()]); kinds.push('s');
    list.forEach(r => { rows.push([false, r]); kinds.push('r'); });
  });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 760);
  sh.setHiddenGridlines(true); sh.setFrozenRows(1);
  sh.getRange(1, 1, rows.length, 2).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('middle').setWrap(true);
  sh.getRange(1, 1, 1, 2).setBackground(C.ink).setFontColor(C.white).setFontWeight('bold');
  sh.setRowHeight(1, 34);
  kinds.forEach((k, i) => {
    if (k === 's') sh.getRange(i + 1, 1, 1, 2).setBackground('#e5e7eb').setFontWeight('bold').setFontColor(C.ink);
    if (k === 'r') sh.getRange(i + 1, 1).insertCheckboxes().setHorizontalAlignment('center');
  });
  const body = sh.getRange(2, 1, rows.length - 1, 2);
  sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2=TRUE').setBackground(C.greenBg).setRanges([body]).build()]);
  sh.setTabColor('#7c3aed');
  props.setProperty('LIBRARY_STAMP', LIBRARY_VERSION);
  markLibraryTicks_();
}

/** Ticks mirror the Rules tab: a library rule is ticked exactly when the same text is in Rules. */
function markLibraryTicks_() {
  const ss = getSheet_(false), lib = ss.getSheetByName('Rule Library'), rules = ss.getSheetByName('Rules');
  if (!lib || !rules || lib.getLastRow() < 2) return;
  const have = new Set(rules.getLastRow() < 2 ? [] :
    rules.getRange(2, 1, rules.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()));
  const rng = lib.getRange(2, 1, lib.getLastRow() - 1, 2);
  const vals = rng.getValues();
  const ticks = vals.map(([t, text]) => [typeof t === 'boolean' ? have.has(String(text).trim()) : t]);
  lib.getRange(2, 1, ticks.length, 1).setValues(ticks);
  const on = vals.filter((v, i) => ticks[i][0] === true).map(v => String(v[1]).trim());
  PS_().getUserProperties().setProperty('LIB_LAST', JSON.stringify(on));
}

/** Apply ticks/unticks from the Rule Library to the Rules tab. */
function syncRuleLibrary_() {
  const ss = getSheet_(false), lib = ss.getSheetByName('Rule Library'), rules = ss.getSheetByName('Rules');
  if (!lib || !rules || lib.getLastRow() < 2) return;
  const last = new Set(JSON.parse(PS_().getUserProperties().getProperty('LIB_LAST') || '[]'));
  // Only act on boxes that changed since the last sync, so edits made in the Rules tab win otherwise.
  const want = lib.getRange(2, 1, lib.getLastRow() - 1, 2).getValues()
    .filter(r => typeof r[0] === 'boolean' && r[0] !== last.has(String(r[1]).trim()));
  if (!want.length) return;
  const n = Math.max(rules.getLastRow() - 1, 0);
  const current = n ? rules.getRange(2, 1, n, 2).getValues() : [];
  const texts = current.map(r => String(r[0]).trim());
  // Remove unticked library rules (bottom-up so row numbers stay valid)
  const drop = new Set(want.filter(r => r[0] === false).map(r => String(r[1]).trim()));
  for (let i = texts.length - 1; i >= 0; i--) if (texts[i] && drop.has(texts[i])) rules.deleteRow(i + 2);
  // Add ticked ones that aren't there yet
  const now = new Set(texts.filter(t => !drop.has(t)));
  want.filter(r => r[0] === true && !now.has(String(r[1]).trim())).forEach(r => {
    const row = nextEmptyRuleRow_(rules);
    rules.getRange(row, 1, 1, 2).setValues([[String(r[1]).trim(), true]]);
    rules.getRange(row, 2).insertCheckboxes().check();
  });
  compactRules_();
  markLibraryTicks_();
}

/** First row in the Rules tab with no rule text (the On? checkboxes run far down, so getLastRow can't be used). */
function nextEmptyRuleRow_(rules) {
  const last = Math.max(rules.getLastRow(), 2);
  const texts = rules.getRange(2, 1, last - 1, 1).getValues();
  const i = texts.findIndex(r => String(r[0]).trim() === '');
  return i >= 0 ? i + 2 : last + 1;
}

/** Close any gaps so every rule sits together at the top of the Rules tab. */
function compactRules_() {
  const rules = getSheet_(false).getSheetByName('Rules');
  if (!rules || rules.getLastRow() < 2) return;
  const n = rules.getLastRow() - 1;
  const vals = rules.getRange(2, 1, n, 2).getValues();
  const kept = vals.filter(r => String(r[0]).trim() !== '');
  const firstGap = vals.findIndex(r => String(r[0]).trim() === '');
  if (firstGap < 0 || !vals.slice(firstGap).some(r => String(r[0]).trim() !== '')) return; // already tidy
  const out = kept.map(r => [r[0], r[1] === true || String(r[1]).toUpperCase() === 'TRUE'])
    .concat(Array.from({ length: n - kept.length }, () => ['', false]));
  rules.getRange(2, 1, n, 2).setValues(out);
  rules.getRange(2, 2, n, 1).insertCheckboxes();
}

/** Installable on-edit handler (full copies): react instantly to the Rule Library and UPDATE_CODE. */
function onSheetEdit(e) {
  try {
    const name = e && e.range ? e.range.getSheet().getName() : '';
    if (name === 'Rule Library') syncRuleLibrary_();
    if (name === 'Rules') markLibraryTicks_();
    if (name === 'Settings' && readSettings_().UPDATE_CODE) { writeSetting_('UPDATE_CODE', 'FALSE'); updateCode(); }
  } catch (err) { Logger.log('onSheetEdit: ' + err); }
}

// ---------- tab styling ----------
const STYLE_VERSION = '2026-11-11d';
const C = {
  ink: '#1f2937', white: '#ffffff', band: '#f8fafc', line: '#e5e7eb', muted: '#6b7280',
  edit: '#fff7e6', amber: '#b45309', amberBg: '#fef3c7', green: '#047857', greenBg: '#d1fae5',
  faint: '#9ca3af',
};

/** Applies the house style to every tab once per STYLE_VERSION. */
function styleTabs_(force) {
  const props = PS_().getUserProperties();
  if (!force && props.getProperty('STYLE_STAMP') === STYLE_VERSION) return;
  const ss = getSheet_(false);
  const order = [['How to', C.ink], ['Rules', '#059669'], ['Rule Library', '#7c3aed'], ['Settings', '#d97706'], ['Log', '#2563eb'], ['Pending', C.faint]];
  order.forEach(([name, color], i) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setTabColor(color);
    ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName('How to') || ss.getSheets()[0]);

  styleRules_(ss.getSheetByName('Rules'));
  styleSettings_(ss.getSheetByName('Settings'));
  styleLog_(ss.getSheetByName('Log'));
  stylePending_(ss.getSheetByName('Pending'));
  newestFirst_(ss.getSheetByName('Log'), 1);
  newestFirst_(ss.getSheetByName('Pending'), 4);
  props.setProperty('STYLE_STAMP', STYLE_VERSION);
}

/** Re-apply styling on demand (e.g. after accidental formatting changes). */
function restyleTabs() { styleTabs_(true); }

function baseTable_(sh, cols, widths) {
  if (!sh) return null;
  sh.getBandings().forEach(b => b.remove());
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(1);
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const all = sh.getRange(1, 1, sh.getMaxRows(), cols);
  all.setFontFamily('Arial').setFontSize(11).setVerticalAlignment('middle')
     .setBorder(null, null, null, null, null, null);
  const head = sh.getRange(1, 1, 1, cols);
  head.setBackground(C.ink).setFontColor(C.white).setFontWeight('bold').setFontSize(11);
  sh.setRowHeight(1, 34);
  const band = sh.getRange(1, 1, sh.getMaxRows(), cols)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor(C.ink).setFirstRowColor(C.white).setSecondRowColor(C.band);
  return all;
}

function styleRules_(sh) {
  if (!baseTable_(sh, 2, [680, 70])) return;
  sh.getRange(1, 1, 1, 2).setValues([['📝 Rule — plain English, one per row', 'On?']]);
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 2, sh.getMaxRows(), 1).setHorizontalAlignment('center');
  // Checkbox column: green when on, grey text when off
  const body = sh.getRange(2, 1, sh.getMaxRows() - 1, 2);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2<>"",$B2=FALSE)')
      .setFontColor(C.faint).setRanges([body]).build(),
  ]);
}

function styleSettings_(sh) {
  if (!baseTable_(sh, 3, [230, 300, 560])) return;
  sh.getRange(1, 1, 1, 3).setValues([['⚙️ Setting', 'Value  (edit these)', 'What it does']]);
  const n = Math.max(sh.getLastRow() - 1, 1);
  sh.getRange(2, 1, n, 1).setFontWeight('bold').setFontFamily('Roboto Mono').setFontSize(10);
  sh.getRange(2, 3, n, 1).setFontColor(C.muted).setWrap(true).setFontSize(10);
  sh.getRange(2, 2, n, 1).setBackground(C.edit).setWrap(true);
  // Auto-managed rows: grey, not highlighted as editable
  const keys = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  keys.forEach((k, i) => {
    if (['LAST_DIGEST', 'LAST_CHECK', 'LAST_UPDATE', 'ACCOUNT'].indexOf(k) >= 0) sh.getRange(i + 2, 2).setBackground(null).setFontColor(C.faint).setFontStyle('italic');
  });
  const vals = sh.getRange(2, 2, n, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2="DRY_RUN",UPPER($B2&"")="TRUE")')
      .setBackground(C.amberBg).setFontColor(C.amber).setBold(true).setRanges([vals]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($A2="DRY_RUN",UPPER($B2&"")="FALSE")')
      .setBackground(C.greenBg).setFontColor(C.green).setBold(true).setRanges([vals]).build(),
  ]);
}

function styleLog_(sh) {
  if (!baseTable_(sh, 8, [150, 200, 300, 110, 230, 380, 80, 120])) return;
  sh.getRange(1, 1, 1, 8).setValues([['🕒 When', 'From', 'Subject', 'Category', 'What it did', 'Summary', 'Mode', 'Thread']]);
  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, 1, rows, 1).setNumberFormat('ddd mmm d, h:mm am/pm').setFontColor(C.muted).setFontSize(10);
  sh.getRange(2, 2, rows, 1).setFontColor(C.muted);
  sh.getRange(2, 6, rows, 1).setFontColor(C.muted).setFontSize(10);
  sh.getRange(2, 7, rows, 1).setHorizontalAlignment('center').setFontSize(9);
  sh.getRange(2, 8, rows, 1).setFontColor(C.faint).setFontSize(9);
  sh.getRange(2, 1, rows, 8).setWrap(false).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  const act = sh.getRange(2, 5, rows, 1), mode = sh.getRange(2, 7, rows, 1), whole = sh.getRange(2, 1, rows, 8);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=REGEXMATCH($E2,"ALERT")')
      .setBackground('#fee2e2').setRanges([whole]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=REGEXMATCH($E2,"needs me")')
      .setBackground(C.amberBg).setRanges([whole]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('archive')
      .setFontColor(C.muted).setRanges([act]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('DRY RUN')
      .setFontColor(C.amber).setBold(true).setRanges([mode]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('live')
      .setFontColor(C.green).setBold(true).setRanges([mode]).build(),
  ]);
}

function stylePending_(sh) {
  if (!baseTable_(sh, 4, [150, 420, 100, 180])) return;
  sh.getRange(1, 1, 1, 4).setValues([['Thread', '⏳ Subject (waiting to be archived)', 'Action', 'Due']]);
  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, 1, rows, 1).setFontColor(C.faint).setFontSize(9);
  sh.getRange(2, 4, rows, 1).setNumberFormat('ddd mmm d, h:mm am/pm').setFontColor(C.muted);
}

// ---------- How to tab ----------
const HOWTO_VERSION = '2026-11-11m';
const SETUP_GUIDE_URL = 'https://claude.ai/artifact/MqpvLwnKzfprXRwC25EjDm';

function ensureHowTo_() {
  const ss = getSheet_(false);
  const props = PS_().getUserProperties();
  const s = readSettings_();
  const includes = String(s.INCLUDE_ACCOUNTS || '').split(',').filter(x => x.trim()).length;
  const role = !s.SEND_DIGEST ? 'SECONDARY — its mail shows up in another account\'s digest' +
      (s.SEND_ALERTS_TO ? '; alerts go to ' + s.SEND_ALERTS_TO : '')
    : includes ? 'MAIN — sends the combined digest (includes ' + includes + ' other account' + (includes > 1 ? 's' : '') + ')'
    : 'STANDALONE — sends its own digest';
  const status = s.DRY_RUN ? 'DRY RUN — only logging, not touching mail' : 'LIVE — acting on mail';
  const codeMode = viaLibrary_() ? 'Shared from your main account — always current. Nothing to install here.'
    : 'Version ' + CODE_VERSION + ' · ' + (/^beta$/i.test(String(s.UPDATE_CHANNEL)) ? 'beta' : 'main') + ' channel · automatic updates ' + (s.AUTO_UPDATE ? 'ON' : 'OFF');
  const bound = !!boundSheet_();
  const stamp = HOWTO_VERSION + '|' + role + '|' + status + '|' + codeMode + '|' + bound;
  let sh = ss.getSheetByName('How to');
  if (sh && props.getProperty('HOWTO_STAMP') === stamp) return;
  if (!sh) sh = ss.insertSheet('How to', 0);
  else { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); sh.clear(); }

  const H = 'h', S = 's', T = 't'; // header, section, table header
  const buttons = bound ? [
    [S, 'BUTTONS — the 📬 Inbox Assistant menu at the top of this sheet', ''],
    [T, 'Menu item', 'Use it when'],
    ['', 'Open the app', 'You want the phone-friendly app: Today, Rules, Settings, and all your linked accounts in one place.'],
    ['', 'Check inbox now', 'You don\'t want to wait up to 10 minutes for the next check.'],
    ['', 'Send digest now', 'You want a digest right now, e.g. to preview a change.'],
    ['', 'Install latest version', 'You want the newest version right away.'],
    ['', 'Undo last update', 'Something broke after an update.'],
    ['', 'Re-check last 2 days', 'You made big rule changes and want recent inbox mail looked at again.'],
    ['', 'Fix tab layout & How to', 'A tab\'s colors or layout got messed up.'],
    ['', 'Set up (first time only)', 'Only once, right after making your copy.'],
  ] : [
    [S, 'BUTTONS — the Run menu in the Apps Script editor (dropdown next to ▶ Run)', ''],
    [T, 'Pick this…', '…when'],
    ['', 'runNow', 'You don\'t want to wait up to 10 minutes for the next check.'],
    ['', 'sendDigestNow', 'You want a digest right now, e.g. to preview a change.'],
    ['', 'updateCode', 'You want the newest version right away.'],
    ['', 'restorePreviousCode', 'Something broke after an update.'],
    ['', 'forgetProcessedLabel', 'You made big rule changes. Run it, then runNow, to re-check recent inbox mail.'],
    ['', 'publishApp', 'Shows your app link, or how to publish the app (View → Logs after running).'],
    ['', 'refreshHowTo · restyleTabs', 'A tab\'s text, colors, or layout got messed up.'],
    ['', 'reinstallTriggers', 'Rarely needed — schedule changes apply on their own within 10 min.'],
    ['', 'setup', 'Only on first install. Don\'t re-run it.'],
    ['', 'processInbox · processPending · sendDigest · onSheetEdit', 'Automatic. You never need to pick them.'],
  ];
  const rows = [
    [H, '📬 Inbox Assistant — How to', ''],
    ['', 'This account', s.ACCOUNT || Session.getActiveUser().getEmail()],
    ['', 'Role', role],
    ['', 'Status', status],
    ['', 'Version', codeMode],
    ['', 'Setup guide', 'Open the setup guide ↗  (step-by-step for a new account or for sharing with someone)'],
    ['', '', ''],
  ].concat(buttons, [
    ['', '', ''],
    [S, 'COMMON SCENARIOS', ''],
    [T, 'I want to…', 'Do this'],
    ['', 'Use it on my phone', 'Publish it once: Extensions → Apps Script → Deploy → New deployment → gear → Web app → Execute as Me, Only myself → Deploy. Open the link on your phone, then Share → Add to Home Screen. It updates itself after that.'],
    ['', 'See all my Gmail accounts in one place', 'In the app on your main account: Settings → Link another account. On the other account\'s app: Settings → Join a main account, and enter the code.'],
    ['', 'Add ready-made rules', 'Rule Library tab → tick a rule to add it to your Rules tab, untick to remove it.'],
    ['', 'Write my own rule', 'Rules tab → type it in plain English in a new row and tick On?. Takes effect on the next check.'],
    ['', 'Turn a rule off for a while', 'Rules tab → untick its On? box.'],
    ['', 'Keep an email safe from the assistant', 'Star it in Gmail. Starred = hands off, always.'],
    ['', 'Test rule changes without risk', 'Settings → DRY_RUN = TRUE. Watch the Log tab. Set FALSE to go live again.'],
    ['', 'Pause everything', 'Settings → DRY_RUN = TRUE.'],
    ['', 'Change how fast read mail gets archived', 'Settings → ARCHIVE_READ_AFTER_DAYS (blank or 0 = never).'],
    ['', 'Change digest time', 'Settings → DIGEST_HOUR (0–23). Applies on its own within 10 min.'],
    ['', 'Get more than one digest a day', 'Settings → DIGEST_HOUR = hours separated by commas, e.g. 7,12 (7am and noon). Each digest covers everything since the last one.'],
    ['', 'See why something happened', 'Log tab — newest at the top. Red rows = alerts, yellow = needs you.'],
    ['', 'Something broke', 'Extensions → Apps Script → Executions (left sidebar) → look at the red error.'],
    ['', '', ''],

    [S, 'UPDATES', ''],
    [T, 'I want to…', 'Do this'],
    ['', 'Get the newest version now', bound ? 'Menu: 📬 Inbox Assistant → Install latest version.' : 'Settings → UPDATE_CODE = TRUE (or Run → updateCode).'],
    ['', 'Get updates automatically', 'Settings → AUTO_UPDATE = TRUE (or the switch in the app). It checks once a day.'],
    ['', 'Undo an update', bound ? 'Menu: 📬 Inbox Assistant → Undo last update.' : 'Run → restorePreviousCode.'],
    ['', 'Try new features early', 'Settings → UPDATE_CHANNEL = beta. Set back to main for stable releases only.'],
    ['', 'Know if an update is waiting', 'Your digest shows a 🆕 note when a new version is out.'],
    ['', '', ''],

    [S, 'MORE THAN ONE GMAIL ACCOUNT  (one digest for all of them)', ''],
    [T, 'Step', 'What to do'],
    ['', '1. Set up each account', 'Follow the setup guide (link at the top) while signed in to that account.'],
    ['', '2. In each EXTRA account', 'Settings: SEND_DIGEST = FALSE, SEND_ALERTS_TO = your main address. Share that sheet with your main address (Viewer).'],
    ['', '3. In your MAIN account', 'Settings: paste each extra sheet\'s link into INCLUDE_ACCOUNTS (comma-separated).'],
    ['', '4. Check', bound ? 'Main account menu → Send digest now. A red "couldn\'t open" line = a sheet isn\'t shared yet.' : 'Main account: Run → sendDigestNow. A red "couldn\'t open" line = a sheet isn\'t shared yet.'],
    ['', '', ''],

    [S, 'SETTINGS CHEAT SHEET (by account type)', ''],
    [T, 'Setting', 'Main  /  Extra  /  Only one account'],
    ['', 'SEND_DIGEST', 'TRUE  /  FALSE  /  TRUE'],
    ['', 'SEND_ALERTS_TO', 'blank  /  your main address  /  blank'],
    ['', 'INCLUDE_ACCOUNTS', 'links to extra sheets  /  blank  /  blank'],
    ['', 'Everything else', 'Same everywhere — see the Settings tab\'s "What it does" column.'],
    ['', '', ''],

    [S, 'GMAIL LABELS IT USES', ''],
    ['', 'AI/⭐ Needs me', 'Waiting on you. Clears itself once you reply.'],
    ['', 'AI/🔔 Alert', 'A rule said "alert me" — you also got an email.'],
    ['', 'AI/Processed', 'Already handled. Ignore.'],
    ['', 'AI/Previewed', 'Seen during test mode (DRY_RUN). Ignore.'],
    ['', 'AI/(anything else)', 'Labels your rules asked for, e.g. AI/Newsletter.'],
  ]);

  sh.getRange(1, 1, rows.length, 2).setValues(rows.map(r => [r[1], r[2]]));
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 640);
  sh.getRange(1, 1, rows.length, 2).setWrap(true).setVerticalAlignment('top').setFontSize(11);
  rows.forEach((r, i) => {
    const rg = sh.getRange(i + 1, 1, 1, 2);
    if (r[0] === H) { rg.merge().setFontSize(18).setFontWeight('bold'); }
    else if (r[0] === S) { rg.merge().setBackground('#1f2937').setFontColor('#ffffff').setFontWeight('bold'); }
    else if (r[0] === T) { rg.setBackground('#e5e7eb').setFontWeight('bold'); }
    else if (r[1] && r[2]) { sh.getRange(i + 1, 1).setFontWeight('bold'); }
  });
  // Make the setup guide a clickable link
  const guideRow = rows.findIndex(r => r[1] === 'Setup guide');
  if (guideRow >= 0) {
    const text = rows[guideRow][2], linkText = 'Open the setup guide ↗';
    sh.getRange(guideRow + 1, 2).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text)
      .setLinkUrl(0, linkText.length, SETUP_GUIDE_URL).build());
  }
  sh.getRange(4, 2).setFontColor(s.DRY_RUN ? '#b45309' : '#047857').setFontWeight('bold');
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(0);
  props.setProperty('HOWTO_STAMP', stamp);
}

/** Rebuild the How to tab on demand. */
function refreshHowTo() {
  PS_().getUserProperties().deleteProperty('HOWTO_STAMP');
  ensureHowTo_();
}

function writeSetting_(key, value) {
  const sh = getSheet_(false).getSheetByName('Settings');
  const keys = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).trim());
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(value);
  else sh.appendRow([key, value, '']);
}

function readRules_() {
  const sh = getSheet_(false).getSheetByName('Rules');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(([text, on]) => text && (on === true || String(on).toUpperCase() === 'TRUE'))
    .map(([text]) => String(text).trim());
}

// ---------- main loop ----------
function processInbox() {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a');
  let msg;
  try { msg = processInbox_(); }
  catch (e) { msg = '⚠ ' + e.message; writeSetting_('LAST_CHECK', stamp + ' · ' + msg); throw e; }
  if (msg) writeSetting_('LAST_CHECK', stamp + ' · ' + msg);
  return msg;
}

function processInbox_() {
  syncSettingsRows_();
  try { syncTriggers_(); } catch (e) { Logger.log('Trigger sync: ' + e); }
  try { compactRules_(); syncRuleLibrary_(); markLibraryTicks_(); } catch (e) { Logger.log('Rule Library sync: ' + e); }
  if (readSettings_().UPDATE_CODE) {
    writeSetting_('UPDATE_CODE', 'FALSE');
    updateCode(); // new code takes over from the next run
    return 'Installed the latest version — see LAST_UPDATE.';
  }
  try { checkLinkRequests_(); } catch (e) { Logger.log('Link requests: ' + e); }
  const s = readSettings_();
  if (!s.GEMINI_API_KEY) return '⚠ Waiting for your Gemini key — paste it into GEMINI_API_KEY.';
  const rules = readRules_();
  if (!rules.length) return '⚠ No rules are on — tick some in Rule Library or add your own in Rules.';
  const lookback = Number(s.LOOKBACK_DAYS) || 2;

  let query = 'in:inbox -label:"' + LABEL_PROCESSED + '" newer_than:' + (Number(s.LOOKBACK_DAYS) || 2) + 'd';
  if (s.DRY_RUN) query += ' -label:"' + LABEL_PREVIEWED + '"';
  query += ' -is:starred'; // starred = hands off, always
  let threads = GmailApp.search(query, 0, Number(s.MAX_PER_RUN) || 30);

  // Never process the assistant's own emails (digests, alerts, setup notice) — just mark them done.
  const own = threads.filter(t => isOwnEmail_(t.getFirstMessageSubject()));
  if (own.length && !s.DRY_RUN) { const p = getOrCreateLabel_(LABEL_PROCESSED); own.forEach(t => t.addLabel(p)); }
  threads = threads.filter(t => !isOwnEmail_(t.getFirstMessageSubject()));
  // You sent the latest message → nothing to decide yet; wait for their reply.
  const meLower = Session.getActiveUser().getEmail().toLowerCase();
  threads = threads.filter(t => String(t.getMessages().slice(-1)[0].getFrom()).toLowerCase().indexOf(meLower) < 0);
  if (!threads.length) return 'No new emails to sort. It looks at inbox mail from the last ' + lookback +
    ' days that hasn\'t been sorted yet' + (s.DRY_RUN ? ' or previewed in test mode' : '') + '.';

  const items = threads.map(t => {
    const msgs = t.getMessages();
    const m = msgs[msgs.length - 1];
    return {
      id: t.getId(),
      from: m.getFrom(),
      to: m.getTo(),
      subject: t.getFirstMessageSubject() || '(no subject)',
      date: m.getDate().toISOString(),
      messages_in_thread: msgs.length,
      i_replied_already: msgs.some(x => x.getFrom().indexOf(Session.getActiveUser().getEmail()) >= 0),
      body: cleanBody_(m.getPlainBody()),
      thread: t,
    };
  });

  const decisions = askGemini_(s, rules, items);
  const processed = getOrCreateLabel_(LABEL_PROCESSED);
  const previewed = getOrCreateLabel_(LABEL_PREVIEWED);
  const log = getSheet_(false).getSheetByName('Log');
  const pending = getSheet_(false).getSheetByName('Pending');
  const me = Session.getActiveUser().getEmail();

  items.forEach(it => {
    const d = decisions[it.id] || {};
    const actions = applyDecision_(it, d, s, pending, me);
    log.appendRow([new Date(), it.from, it.subject, d.category || '', actions.join('; ') || 'leave alone',
      d.summary || '', s.DRY_RUN ? 'DRY RUN' : 'live', it.id]);
    if (s.DRY_RUN) it.thread.addLabel(previewed);
    else it.thread.addLabel(processed);
  });
  newestFirst_(log, 1);
  newestFirst_(pending, 4);
  return 'Sorted ' + items.length + ' email' + (items.length === 1 ? '' : 's') +
    (s.DRY_RUN ? ' in test mode (nothing changed in Gmail) — see the Log tab.' : ' — see the Log tab.');
}

function applyDecision_(it, d, s, pending, me) {
  const done = [];
  const t = it.thread;
  const live = !s.DRY_RUN;

  (d.labels || []).forEach(name => {
    name = String(name).trim(); if (!name) return;
    const full = name.startsWith(LABEL_ROOT + '/') ? name : LABEL_ROOT + '/' + name;
    if (live) t.addLabel(getOrCreateLabel_(full));
    done.push('label ' + name);
  });
  if (d.needs_me) {
    if (live) t.addLabel(getOrCreateLabel_(LABEL_IMPORTANT));
    done.push('needs me');
  }
  if (d.star) { if (live) t.markImportant(); done.push('mark important'); }
  if (d.alert) {
    if (live) {
      t.addLabel(getOrCreateLabel_(LABEL_ALERT));
      if (s.ALERTS_IMMEDIATE) {
        const to = s.SEND_ALERTS_TO || me;
        const via = to !== me ? ' [' + me + ']' : '';
        GmailApp.sendEmail(to, '🔔 ' + it.subject + via, (d.alert_reason || 'Flagged by your rules.') +
          '\n\nFrom: ' + it.from + '\nTo: ' + me + '\n' + threadLink_(it.id, me));
      }
    }
    done.push('ALERT: ' + (d.alert_reason || ''));
  }
  if (d.archive_now) {
    if (live) t.moveToArchive();
    done.push('archive now');
  } else if (Number(d.archive_after_days) > 0) {
    const due = new Date(Date.now() + Number(d.archive_after_days) * 86400000);
    if (live) pending.appendRow([it.id, it.subject, 'archive', due]);
    done.push('archive in ' + d.archive_after_days + 'd');
  }
  return done;
}

/**
 * Hourly tidy-up:
 *  1. You replied to a "Needs me" email → remove the badge (and let a future reply get re-checked).
 *  2. Read emails older than ARCHIVE_READ_AFTER_DAYS → archive (never starred or "Needs me").
 */
function tidyInbox_() {
  const s = readSettings_();
  if (s.DRY_RUN) return;
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const needsMe = GmailApp.getUserLabelByName(LABEL_IMPORTANT);
  const processed = GmailApp.getUserLabelByName(LABEL_PROCESSED);
  const log = getSheet_(false).getSheetByName('Log');

  // 1. Replied → clear "Needs me"
  if (needsMe) {
    GmailApp.search('label:"' + LABEL_IMPORTANT + '"', 0, 100).forEach(t => {
      const msgs = t.getMessages();
      const last = msgs[msgs.length - 1];
      if (String(last.getFrom()).toLowerCase().indexOf(me) >= 0) {
        t.removeLabel(needsMe);
        if (processed) t.removeLabel(processed); // their next reply gets looked at fresh
        try { setLogStatus_(getSheet_(false), t.getId(), 'done'); } catch (e) {}
      }
    });
  }

  // 2. Read + old enough → archive
  const days = Number(s.ARCHIVE_READ_AFTER_DAYS);
  if (days > 0) {
    const q = 'in:inbox is:read -is:starred -label:"' + LABEL_IMPORTANT + '" older_than:' + Math.floor(days) + 'd';
    const threads = GmailApp.search(q, 0, 100);
    threads.forEach(t => {
      const m = t.getMessages().slice(-1)[0];
      log.appendRow([new Date(), m.getFrom(), t.getFirstMessageSubject() || '(no subject)', 'Read',
        'archive (read, ' + days + '+ days old)', '', 'live', t.getId()]);
    });
    for (let i = 0; i < threads.length; i += 100) GmailApp.moveThreadsToArchive(threads.slice(i, i + 100));
    if (threads.length) newestFirst_(log, 1);
  }
}

/** Sort a tab's rows (below the header) newest first by the given date column. */
function newestFirst_(sh, col) {
  if (!sh || sh.getLastRow() < 3) return;
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).sort({ column: col, ascending: false });
}

function processPending() {
  try { syncSettingsRows_(); tidyInbox_(); } catch (e) { Logger.log('Tidy: ' + e); }
  try { dailyUpdateCheck_(); } catch (e) { Logger.log('Update check: ' + e); }
  try { appUrl_(); } catch (e) { Logger.log('App link: ' + e); }
  const s = readSettings_();
  const sh = getSheet_(false).getSheetByName('Pending');
  if (sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const now = Date.now();
  for (let i = rows.length - 1; i >= 0; i--) {
    const [id, , action, due] = rows[i];
    if (!id || !(due instanceof Date) || due.getTime() > now) continue;
    try {
      const t = GmailApp.getThreadById(id);
      // Starred (or moved out of the inbox by you) since it was queued? Hands off.
      if (t && action === 'archive' && !s.DRY_RUN && !t.hasStarredMessages() && t.isInInbox()) t.moveToArchive();
      // Asked for in the app from your main account: always carried out.
      if (t && /^app-/.test(action)) applyEmailAction_(t, action.slice(4));
    } catch (e) { Logger.log('Pending ' + id + ': ' + e); }
    sh.deleteRow(i + 2);
  }
}

// ---------- the app (web page) ----------
// Each account serves its own private app (Deploy → Web app, execute as me, only me).
// The app shows this account plus any linked accounts (INCLUDE_ACCOUNTS), reading and writing their sheets.
const LOG_STATUS_COL = 9;
const EDITABLE_SETTINGS = ['DRY_RUN', 'DIGEST_HOUR', 'ARCHIVE_READ_AFTER_DAYS', 'TAP_ACTION', 'AUTO_UPDATE', 'ACCOUNT_LABEL'];

function doGet() {
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('Inbox Assistant')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes');
}

function defaultLabel_(email) {
  const [local, domain] = String(email || '').toLowerCase().split('@');
  if (!local) return 'Account';
  if (domain && !/^(gmail|googlemail)\.com$/.test(domain)) { const d = domain.split('.')[0]; return d.charAt(0).toUpperCase() + d.slice(1); }
  return local;
}

function settingIn_(ss, key) { return settingFromSheet_(ss, key); }

function writeSettingIn_(ss, key, value) {
  const sh = ss.getSheetByName('Settings');
  const n = Math.max(sh.getLastRow() - 1, 1);
  const keys = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim());
  const i = keys.indexOf(key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(value);
  else { const d = DEFAULT_SETTINGS.find(r => r[0] === key); sh.appendRow([key, value, d ? d[2] : '']); }
}

/** This account first, then every linked account the app can open. */
function appAccounts_() {
  const me = Session.getActiveUser().getEmail();
  const self = getSheet_(false);
  const list = [{ key: 'self', ss: self, email: me, label: settingIn_(self, 'ACCOUNT_LABEL') || defaultLabel_(me), own: true }];
  String(readSettings_().INCLUDE_ACCOUNTS || '').split(',').map(x => x.trim()).filter(Boolean).forEach(ref => {
    try {
      const ss = /^https?:/.test(ref) ? SpreadsheetApp.openByUrl(ref) : SpreadsheetApp.openById(ref);
      const email = settingIn_(ss, 'ACCOUNT') || ss.getName();
      list.push({ key: ss.getId(), ss: ss, email: email, label: settingIn_(ss, 'ACCOUNT_LABEL') || defaultLabel_(email), own: false });
    } catch (e) {
      list.push({ key: ref, ss: null, email: '', label: 'Linked account', own: false, error: 'Can\'t open this account\'s sheet. It needs to be shared with ' + me + ' as Editor.' });
    }
  });
  return list;
}

function appAccount_(key) {
  const a = appAccounts_().find(x => x.key === (key || 'self'));
  if (!a || !a.ss) throw new Error('That account isn\'t available right now.');
  return a;
}

function readLogRows_(ss, limit) {
  const log = ss.getSheetByName('Log');
  if (!log || log.getLastRow() < 2) return [];
  const n = Math.min(log.getLastRow() - 1, limit || 400);
  const cols = Math.max(log.getLastColumn(), LOG_STATUS_COL);
  return log.getRange(2, 1, n, cols).getValues().filter(r => r[0] instanceof Date).map((r, i) => ({
    when: r[0], from: String(r[1]), subject: String(r[2]), category: String(r[3]), action: String(r[4]),
    summary: String(r[5]), mode: String(r[6]), id: String(r[7]), status: String(r[8] || ''), row: i + 2,
  }));
}

function setLogStatus_(ss, threadId, status) {
  const log = ss.getSheetByName('Log');
  if (!log || log.getLastRow() < 2) return;
  if (!String(log.getRange(1, LOG_STATUS_COL).getValue())) log.getRange(1, LOG_STATUS_COL).setValue('Status');
  const ids = log.getRange(2, 8, log.getLastRow() - 1, 1).getValues();
  ids.forEach((r, i) => { if (String(r[0]) === String(threadId)) log.getRange(i + 2, LOG_STATUS_COL).setValue(status); });
}

function readRulesIn_(ss) {
  const sh = ss.getSheetByName('Rules');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ text: String(r[0]).trim(), on: r[1] === true || String(r[1]).toUpperCase() === 'TRUE' }));
}

function libraryTypeOf_(text) {
  for (const [type, list] of RULE_LIBRARY) if (list.indexOf(text) >= 0) return type;
  return '';
}

function startOfToday_() {
  const tz = Session.getScriptTimeZone();
  return new Date(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'00:00:00XXX"));
}

function iso_(d) { return d instanceof Date ? d.toISOString() : ''; }

/** Everything the app shows, for one account or 'all'. */
function appState(view) {
  view = view || 'all';
  const accounts = appAccounts_();
  const today = startOfToday_().getTime();
  const weekAgo = Date.now() - 7 * 86400000;
  const shown = view === 'all' ? accounts : accounts.filter(a => a.key === view);
  const needs = [], cleared = {}; let sorted = 0, clearedCount = 0;
  shown.forEach(a => {
    if (!a.ss) return;
    const seen = {};
    readLogRows_(a.ss, 400).forEach(r => {
      const t = r.when.getTime();
      if (t >= today) {
        sorted++;
        if (/archive/.test(r.action)) {
          clearedCount++;
          const cat = r.category || 'Other';
          (cleared[cat] = cleared[cat] || []).push({ subject: r.subject, from: shortFrom_(r.from), when: iso_(r.when),
            account: a.label, link: threadLink_(r.id, a.email) });
        }
      }
      if (t >= weekAgo && /needs me|ALERT/i.test(r.action) && r.status !== 'done' && !seen[r.id]) {
        seen[r.id] = true;
        needs.push({ id: r.id, acct: a.key, account: a.label, own: a.own, subject: r.subject, from: shortFrom_(r.from),
          summary: r.summary, category: r.category, alert: /ALERT/.test(r.action), when: iso_(r.when), link: threadLink_(r.id, a.email) });
      }
    });
  });
  needs.sort((x, y) => (y.alert - x.alert) || (y.when > x.when ? 1 : -1));

  // Rules and settings belong to one account: the one in view, or this account when viewing all.
  const edit = view === 'all' ? accounts[0] : (shown[0] || accounts[0]);
  const es = edit.ss;
  const get = k => settingIn_(es, k);
  const rules = es ? readRulesIn_(es).map(r => Object.assign(r, { type: libraryTypeOf_(r.text) })) : [];
  const inRules = new Set(rules.map(r => r.text));
  const s = readSettings_();
  return {
    me: Session.getActiveUser().getEmail(),
    view: view,
    accounts: accounts.map(a => ({ key: a.key, label: a.label, email: a.email, own: a.own, error: a.error || '' })),
    counts: { needs: needs.length, sorted: sorted, cleared: clearedCount },
    needs: needs.slice(0, 30),
    cleared: cleared,
    editing: { key: edit.key, label: edit.label, own: edit.own },
    rules: rules,
    library: RULE_LIBRARY.map(([type, list]) => ({ type: type, rules: list.map(t => ({ text: t, added: inRules.has(t) })) })),
    settings: es ? {
      live: String(get('DRY_RUN')).toUpperCase() === 'FALSE',
      digestHours: digestHours_({ DIGEST_HOUR: get('DIGEST_HOUR') || '7' }),
      archiveDays: Number(get('ARCHIVE_READ_AFTER_DAYS')) || 0,
      tapAction: (get('TAP_ACTION') || 'app') === 'gmail' ? 'gmail' : 'app',
      autoUpdate: String(get('AUTO_UPDATE')).toUpperCase() === 'TRUE',
      lastCheck: get('LAST_CHECK'),
      sendsDigest: String(get('SEND_DIGEST')).toUpperCase() !== 'FALSE',
      mainAccount: get('MAIN_ACCOUNT'),
    } : null,
    self: { version: CODE_VERSION, lastUpdate: s.LAST_UPDATE || '', sheetUrl: getSheet_(false).getUrl(),
      library: viaLibrary_() || !loader_(), linkCode: activeLinkCode_() },
  };
}

function appSetSetting(acct, key, value) {
  if (EDITABLE_SETTINGS.indexOf(key) < 0) throw new Error('That setting can\'t be changed here.');
  if (key === 'ACCOUNT_LABEL') value = String(value || '').trim().slice(0, 24);
  if (key === 'ARCHIVE_READ_AFTER_DAYS') value = Math.max(0, Math.min(60, Number(value) || 0));
  if (key === 'TAP_ACTION') value = value === 'gmail' ? 'gmail' : 'app';
  if (key === 'DIGEST_HOUR') value = digestHours_({ DIGEST_HOUR: String(value) }).join(',');
  const a = appAccount_(acct);
  writeSettingIn_(a.ss, key, value);
  return true;
}

function appRuleSet(acct, text, on) {
  const a = appAccount_(acct), sh = a.ss.getSheetByName('Rules');
  text = String(text || '').trim();
  const n = Math.max(sh.getLastRow() - 1, 0);
  const vals = n ? sh.getRange(2, 1, n, 1).getValues() : [];
  const i = vals.findIndex(r => String(r[0]).trim() === text);
  if (i >= 0) { sh.getRange(i + 2, 2).insertCheckboxes(); sh.getRange(i + 2, 2).setValue(!!on); }
  else if (on) appRuleAdd(acct, text);
  if (a.own) markLibraryTicks_();
  return true;
}

function appRuleAdd(acct, text) {
  text = String(text || '').trim().slice(0, 500);
  if (text.length < 8) throw new Error('Write the rule as a short sentence.');
  const a = appAccount_(acct), sh = a.ss.getSheetByName('Rules');
  if (readRulesIn_(a.ss).some(r => r.text === text)) return true;
  const row = nextEmptyRuleRow_(sh);
  sh.getRange(row, 1, 1, 2).setValues([[text, true]]);
  sh.getRange(row, 2).insertCheckboxes(); sh.getRange(row, 2).setValue(true);
  if (a.own) markLibraryTicks_();
  return true;
}

function appRuleRemove(acct, text) {
  const a = appAccount_(acct), sh = a.ss.getSheetByName('Rules');
  text = String(text || '').trim();
  const n = Math.max(sh.getLastRow() - 1, 0);
  const vals = n ? sh.getRange(2, 1, n, 1).getValues() : [];
  for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim() === text) sh.deleteRow(i + 2);
  if (a.own) markLibraryTicks_();
  return true;
}

/** One email for the in-app view. Full text only for this account's own mail. */
function appEmail(acct, threadId) {
  const a = appAccount_(acct);
  const r = readLogRows_(a.ss, 600).find(x => x.id === String(threadId)) || {};
  const out = { id: threadId, acct: a.key, account: a.label, own: a.own, link: threadLink_(threadId, a.email),
    subject: r.subject || '', from: r.from || '', summary: r.summary || '', category: r.category || '', action: r.action || '', body: '', date: iso_(r.when) };
  if (a.own) {
    const t = GmailApp.getThreadById(threadId);
    if (t) {
      const m = t.getMessages().slice(-1)[0];
      out.subject = t.getFirstMessageSubject() || out.subject;
      out.from = m.getFrom(); out.to = m.getTo(); out.date = iso_(m.getDate());
      out.body = String(m.getPlainBody() || '').replace(/\r/g, '').replace(/\n{3,}/g, '\n\n').trim().slice(0, 8000);
      out.starred = t.hasStarredMessages();
    }
  }
  return out;
}

function applyEmailAction_(t, action) {
  const needsMe = GmailApp.getUserLabelByName(LABEL_IMPORTANT);
  if (action === 'done' || action === 'archive') { if (needsMe) t.removeLabel(needsMe); }
  if (action === 'archive') t.moveToArchive();
  if (action === 'star') { const m = t.getMessages().slice(-1)[0]; m.isStarred() ? m.unstar() : m.star(); }
}

function appEmailAction(acct, threadId, action) {
  if (['done', 'archive', 'star'].indexOf(action) < 0) throw new Error('Unknown action.');
  const a = appAccount_(acct);
  if (action !== 'star') setLogStatus_(a.ss, threadId, 'done');
  if (a.own) {
    const t = GmailApp.getThreadById(threadId);
    if (t) applyEmailAction_(t, action);
    return action === 'star' ? 'Star changed.' : action === 'archive' ? 'Archived.' : 'Marked done.';
  }
  const p = a.ss.getSheetByName('Pending');
  const r = readLogRows_(a.ss, 600).find(x => x.id === String(threadId));
  p.appendRow([threadId, r ? r.subject : '', 'app-' + action, new Date()]);
  return (action === 'done' ? 'Marked done.' : action === 'archive' ? 'Archiving' : 'Starring') +
    (action === 'done' ? '' : ' — ' + a.label + ' will do it within the hour.');
}

function appCheckNow(acct) {
  const a = appAccount_(acct);
  if (!a.own) return a.label + ' checks its inbox every 10 minutes on its own.';
  return processInbox() || 'Checked.';
}

function appSendDigestNow() { sendDigest(); return 'Digest sent — check your inbox.'; }
function appUpdateNow() { updateCode(); return readSettings_().LAST_UPDATE || 'Done.'; }
function appUndoUpdate() { restorePreviousCode(); return readSettings_().LAST_UPDATE || 'Done.'; }

// ---------- linking accounts with a code ----------
const LINK_SUBJECT = 'Inbox Assistant link request';

function activeLinkCode_() {
  const p = PS_().getUserProperties();
  const code = p.getProperty('LINK_CODE'), exp = Number(p.getProperty('LINK_EXP') || 0);
  return code && Date.now() < exp ? { code: code, expires: new Date(exp).toISOString() } : null;
}

/** On the main account: make a one-hour code another account can use to link itself. */
function appLinkCode() {
  const abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let c = ''; for (let i = 0; i < 6; i++) c += abc.charAt(Math.floor(Math.random() * abc.length));
  const code = c.slice(0, 3) + '-' + c.slice(3);
  const p = PS_().getUserProperties();
  p.setProperty('LINK_CODE', code); p.setProperty('LINK_EXP', String(Date.now() + 3600 * 1000));
  return { code: code, email: Session.getActiveUser().getEmail(), expires: new Date(Date.now() + 3600 * 1000).toISOString() };
}

/** On the other account: share this sheet with the main account and ask it to link. */
function appJoinMain(mainEmail, code) {
  mainEmail = String(mainEmail || '').trim().toLowerCase();
  code = String(code || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!/^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/.test(mainEmail)) throw new Error('That email address doesn\'t look right.');
  if (code.length !== 6) throw new Error('The code has 6 letters and numbers, like K7P-42Q.');
  const me = Session.getActiveUser().getEmail().toLowerCase();
  if (mainEmail === me) throw new Error('Enter your MAIN account\'s address, not this one.');
  const ss = getSheet_(false);
  ss.addEditor(mainEmail);
  writeSettingIn_(ss, 'SEND_DIGEST', 'FALSE');
  writeSettingIn_(ss, 'SEND_ALERTS_TO', mainEmail);
  writeSettingIn_(ss, 'MAIN_ACCOUNT', mainEmail);
  GmailApp.sendEmail(mainEmail, LINK_SUBJECT,
    'This email links two Inbox Assistant accounts. Your main account handles it automatically — no need to do anything.\n\n' +
    JSON.stringify({ code: code.slice(0, 3) + '-' + code.slice(3), email: me, sheet: ss.getUrl() }),
    { name: 'Inbox Assistant' });
  return 'Sent. Your main account will finish linking within about 10 minutes.';
}

/** On the main account: accept link requests that carry the current code. */
function checkLinkRequests_() {
  const threads = GmailApp.search('subject:"' + LINK_SUBJECT + '" newer_than:2d -label:"' + LABEL_PROCESSED + '"', 0, 10);
  if (!threads.length) return;
  const processed = getOrCreateLabel_(LABEL_PROCESSED);
  const active = activeLinkCode_();
  threads.forEach(t => {
    t.getMessages().forEach(m => {
      const json = String(m.getPlainBody() || '').match(/\{[^{}]*"sheet"[^{}]*\}/);
      if (!json) return;
      let req; try { req = JSON.parse(json[0]); } catch (e) { return; }
      if (!active || String(req.code).toUpperCase() !== active.code) return;
      if (!/^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+/.test(String(req.sheet))) return;
      const id = String(req.sheet).match(/\/d\/([A-Za-z0-9_-]+)/)[1];
      const cur = String(readSettings_().INCLUDE_ACCOUNTS || '').split(',').map(x => x.trim()).filter(Boolean);
      if (!cur.some(x => x.indexOf(id) >= 0)) cur.push('https://docs.google.com/spreadsheets/d/' + id + '/edit');
      writeSetting_('INCLUDE_ACCOUNTS', cur.join(', '));
      const p = PS_().getUserProperties(); p.deleteProperty('LINK_CODE'); p.deleteProperty('LINK_EXP');
      writeSetting_('LAST_CHECK', 'Linked ' + (req.email || 'an account') + '.');
    });
    t.addLabel(processed); t.moveToArchive();
  });
}

/** On the main account: stop including a linked account. */
function appUnlink(key) {
  const cur = String(readSettings_().INCLUDE_ACCOUNTS || '').split(',').map(x => x.trim()).filter(Boolean);
  writeSetting_('INCLUDE_ACCOUNTS', cur.filter(x => x.indexOf(key) < 0).join(', '));
  return 'Unlinked. That account keeps sorting its own mail.';
}

// ---------- the app link ----------
/** This copy's app link, once it has been published (Deploy → New deployment → Web app). */
function appUrl_() {
  let url = '';
  try { url = SA_().getService().getUrl() || ''; } catch (e) {}
  if (url && readSettings_().APP_URL !== url) { try { writeSetting_('APP_URL', url); } catch (e) {} }
  return url;
}

const DEPLOY_STEPS = 'Publish your app once (about a minute):\n' +
  '1. In the sheet: Extensions → Apps Script.\n' +
  '2. Top right: Deploy → New deployment.\n' +
  '3. Click the gear next to "Select type" → Web app.\n' +
  '4. Execute as: Me. Who has access: Only myself.\n' +
  '5. Click Deploy (and Authorize / Allow if asked).\n' +
  '6. Copy the Web app URL and open it — on your phone, Share → Add to Home Screen.\n\n' +
  'After that, the app updates itself along with everything else.';

/** Run menu: show the app link, or how to publish it. */
function publishApp() {
  const url = appUrl_();
  Logger.log(url ? 'Your app: ' + url : DEPLOY_STEPS);
  return url;
}

/** Menu: 📬 Inbox Assistant → Open the app */
function openAppFromMenu() {
  const url = appUrl_();
  const ui = SpreadsheetApp.getUi();
  if (!url) { ui.alert('Open the app', DEPLOY_STEPS, ui.ButtonSet.OK); return; }
  const html = HtmlService.createHtmlOutput('<div style="font:15px/1.5 -apple-system,Segoe UI,Arial,sans-serif;padding:4px 2px">' +
    '<p style="margin:0 0 12px">Your private Inbox Assistant app:</p>' +
    '<p style="margin:0 0 14px"><a href="' + url + '" target="_blank" style="font-weight:600;color:#0e7c86">Open the app ↗</a></p>' +
    '<p style="margin:0;color:#55625f;font-size:13px">On your phone: open this link (it\'s also in Settings → APP_URL), then Share → Add to Home Screen.</p></div>')
    .setWidth(360).setHeight(170);
  ui.showModalDialog(html, '📬 Inbox Assistant');
}

// ---------- sheet menu (template installs) ----------
/** Adds the 📬 Inbox Assistant menu when the sheet opens. Works only when the code lives in this sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📬 Inbox Assistant')
    .addItem('Open the app', 'openAppFromMenu')
    .addSeparator()
    .addItem('Set up (first time only)', 'setup')
    .addSeparator()
    .addItem('Check inbox now', 'runNow')
    .addItem('Send digest now', 'sendDigestNow')
    .addSeparator()
    .addItem('Install latest version', 'updateNowFromMenu')
    .addItem('Undo last update', 'restorePreviousCode')
    .addSeparator()
    .addItem('Re-check last 2 days (after big rule changes)', 'recheckFromMenu')
    .addItem('Fix tab layout & How to', 'repairTabs')
    .addToUi();
}

function updateNowFromMenu() {
  updateCode();
  const msg = readSettings_().LAST_UPDATE || 'Done.';
  try { SpreadsheetApp.getUi().alert('Install latest version', String(msg), SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
}

function recheckFromMenu() { forgetProcessedLabel(); runNow(); }

/** Rebuild the How to tab and re-apply styling. */
function repairTabs() { refreshHowTo(); restyleTabs(); }

// ---------- updates (through the loader) ----------
// Loader.gs keeps the downloaded code (and the previous version) in this project's script properties
// and hands us HOST_.loader with: info(), remoteVersion(channel), install(channel), undo().
function loader_() { return HOST_ && HOST_.loader ? HOST_.loader : null; }

function channel_() { return /^beta$/i.test(String(readSettings_().UPDATE_CHANNEL || '')) ? 'beta' : 'main'; }

function stampNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a'); }

/** Install the latest version for this copy's channel. Takes effect from the next run. */
function updateCode() {
  const L = loader_();
  if (viaLibrary_()) { writeSetting_('LAST_UPDATE', stampNow_() + ' · shared code — always current, nothing to do'); return; }
  if (!L) { writeSetting_('LAST_UPDATE', stampNow_() + ' · ⚠ This copy is from before automatic updates. Replace its code with Loader.gs from the setup page.'); return; }
  try {
    const r = L.install(channel_());
    PS_().getUserProperties().deleteProperty('UPDATE_AVAILABLE');
    writeSetting_('LAST_UPDATE', stampNow_() + (r.changed ? ' · updated ' + r.from + ' → ' + r.to + ' (' + channel_() + '). If anything breaks, use Undo last update.'
      : ' · already up to date (' + r.to + ')'));
  } catch (e) {
    writeSetting_('LAST_UPDATE', stampNow_() + ' · ⚠ ' + e.message);
  }
}

/** Go back to the version from before the last update. */
function restorePreviousCode() {
  const L = loader_();
  if (!L) { writeSetting_('LAST_UPDATE', stampNow_() + ' · nothing to undo here'); return; }
  try { const r = L.undo(); writeSetting_('LAST_UPDATE', stampNow_() + ' · went back to ' + r.to + ' (from ' + r.from + ')'); }
  catch (e) { writeSetting_('LAST_UPDATE', stampNow_() + ' · ⚠ ' + e.message); }
}

/** Once a day: install updates if AUTO_UPDATE is on, otherwise just note that one is available. */
function dailyUpdateCheck_() {
  const L = loader_();
  if (!L) return;
  const props = PS_().getUserProperties();
  const last = Number(props.getProperty('LAST_UPDATE_CHECK') || 0);
  if (Date.now() - last < 23 * 3600 * 1000) return;
  props.setProperty('LAST_UPDATE_CHECK', String(Date.now()));
  if (readSettings_().AUTO_UPDATE) { updateCode(); return; }
  try {
    const v = L.remoteVersion(channel_());
    if (v && v !== CODE_VERSION) props.setProperty('UPDATE_AVAILABLE', v); else props.deleteProperty('UPDATE_AVAILABLE');
  } catch (e) { /* offline — try again tomorrow */ }
}

function updateBanner_() {
  const v = PS_().getUserProperties().getProperty('UPDATE_AVAILABLE');
  if (!v || v === CODE_VERSION) return '';
  return '<div style="background:#eef4ff;border:1px solid #c7d7fe;border-radius:8px;padding:10px 14px;margin:0 0 14px">' +
    '🆕 <b>A new version is available.</b> In the app: Settings → Check for update. Or turn on automatic updates.</div>';
}

// ---------- digest ----------
const MAIL_FONT = 'font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;max-width:640px;line-height:1.5';

function sendDigest() {
  syncSettingsRows_();
  const s = readSettings_();
  if (!s.SEND_DIGEST) return; // another account includes this one
  const since = s.LAST_DIGEST ? new Date(s.LAST_DIGEST) : new Date(Date.now() - 86400000);
  const me = Session.getActiveUser().getEmail();

  // This account first, then any included accounts
  const accounts = [{ email: me, entries: logEntriesSince_(getSheet_(false), since) }];
  String(s.INCLUDE_ACCOUNTS || '').split(',').map(x => x.trim()).filter(Boolean).forEach(ref => {
    try {
      const ss = /^https?:/.test(ref) ? SpreadsheetApp.openByUrl(ref) : SpreadsheetApp.openById(ref);
      const email = settingFromSheet_(ss, 'ACCOUNT') || ss.getName();
      accounts.push({ email: email, entries: logEntriesSince_(ss, since) });
    } catch (e) {
      accounts.push({ email: ref, error: 'Couldn\'t open this account\'s sheet — make sure it\'s shared with ' + me + '.' });
    }
  });

  const total = accounts.reduce((n, a) => n + (a.entries ? a.entries.length : 0), 0);
  if (!total && !accounts.some(a => a.error)) { writeSetting_('LAST_DIGEST', new Date().toISOString()); return; }
  const multi = accounts.length > 1;
  const kind = digestKind_();

  let cleared = 0;
  const body = accounts.map(a => {
    const head = multi ? '<h2 style="margin:28px 0 4px;padding-top:12px;border-top:1px solid #ddd;font-size:18px">📬 ' + esc_(a.email) + '</h2>' : '';
    if (a.error) return head + '<p style="color:#b00">' + esc_(a.error) + '</p>';
    if (!a.entries.length) return multi ? head + '<p style="color:#666">Nothing new.</p>' : '';
    a.entries.forEach(e => e.link = threadLink_(e.id, a.email));
    const isArchived = e => /archive/.test(e.action);
    const kept = a.entries.filter(e => !isArchived(e));
    const gone = a.entries.filter(isArchived);
    cleared += gone.length;
    const needsYou = kept.filter(e => /needs me|ALERT/i.test(e.action));
    const rest = kept.filter(e => !/needs me|ALERT/i.test(e.action));
    let html = head;
    if (needsYou.length) html += needsList_(needsYou);
    if (rest.length) html += '<h3 style="margin:24px 0 4px">📥 Also in your inbox (' + rest.length + ')</h3>' + plainDigest_(rest);
    if (gone.length) html += clearedList_(gone);
    return html;
  }).join('');

  const html = '<div style="' + MAIL_FONT + '">' +
    '<h1 style="margin:0 0 4px;font-size:22px">' + kind.icon + ' Your inbox, ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE MMM d') + '</h1>' +
    '<div style="color:#666;margin-bottom:16px">' + total + ' emails handled · ' + cleared + ' cleared out' +
    (multi ? ' · ' + accounts.length + ' accounts' : '') + '</div>' +
    updateBanner_() +
    body +
    '<p style="color:#999;font-size:12px;margin-top:24px">Inbox Assistant · <a href="' + getSheet_(false).getUrl() + '">edit rules</a></p></div>';
  GmailApp.sendEmail(me, kind.icon + ' ' + kind.name + ' digest — ' + total + ' emails handled', 'See HTML version.', { htmlBody: htmlSafe_(html) });
  writeSetting_('LAST_DIGEST', new Date().toISOString());
}

function logEntriesSince_(ss, since) {
  const log = ss.getSheetByName('Log');
  if (!log || log.getLastRow() < 2) return [];
  return log.getRange(2, 1, log.getLastRow() - 1, 8).getValues()
    .filter(r => r[0] instanceof Date && r[0] > since)
    .map(r => ({ from: r[1], subject: r[2], category: r[3], action: r[4], summary: r[5], id: r[7] }));
}

function settingFromSheet_(ss, key) {
  const sh = ss.getSheetByName('Settings');
  if (!sh || sh.getLastRow() < 2) return '';
  const row = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().find(r => String(r[0]).trim() === key);
  return row ? String(row[1]).trim() : '';
}

/** Top section: alerts first, then anything waiting on you. Highlighted box. */
function needsList_(entries) {
  entries.sort((a, b) => (/ALERT/.test(b.action) ? 1 : 0) - (/ALERT/.test(a.action) ? 1 : 0));
  return '<div style="background:#fff7e6;border:1px solid #f5d79e;border-radius:8px;padding:12px 16px;margin:8px 0 4px">' +
    '<h3 style="margin:0 0 6px">⭐ Needs you (' + entries.length + ')</h3><ul style="margin:0;padding-left:18px">' +
    entries.map(e => '<li style="margin-bottom:6px">' + (/ALERT/.test(e.action) ? '🔔 ' : '') +
      '<a href="' + e.link + '"><b>' + esc_(e.subject) + '</b></a><br><span style="color:#444">' +
      esc_(e.summary || shortFrom_(e.from)) + '</span></li>').join('') +
    '</ul></div>';
}

/** Everything else still in the inbox, grouped by category. */
function plainDigest_(entries) {
  const groups = {};
  entries.forEach(e => { (groups[e.category || 'Other'] = groups[e.category || 'Other'] || []).push(e); });
  return Object.keys(groups).sort().map(cat =>
    '<div style="margin:10px 0 2px;font-weight:600">' + esc_(cat) + ' · ' + groups[cat].length + '</div>' +
    '<ul style="margin:0;padding-left:18px">' +
    groups[cat].map(e => '<li><a href="' + e.link + '">' + esc_(e.subject) + '</a> — <span style="color:#444">' +
      esc_(e.summary || shortFrom_(e.from)) + '</span></li>').join('') +
    '</ul>').join('');
}

/** Every archived email, grouped by category, each linked. Built in code so nothing gets dropped. */
function clearedList_(entries) {
  const groups = {};
  entries.forEach(e => { (groups[e.category || 'Other'] = groups[e.category || 'Other'] || []).push(e); });
  return '<h3 style="margin:24px 0 4px">🧹 Cleared out (' + entries.length + ')</h3>' +
    '<div style="color:#666;font-size:13px;margin-bottom:8px">Archived, not deleted — click any to open it.</div>' +
    Object.keys(groups).sort().map(cat =>
      '<div style="margin:10px 0 2px;font-weight:600">' + esc_(cat) + ' · ' + groups[cat].length + '</div>' +
      '<ul style="margin:0;padding-left:18px">' +
      groups[cat].map(e => {
        const when = /archive in/.test(e.action) ? ' <span style="color:#999">(' + esc_(e.action.match(/archive in \S+/)[0]) + ')</span>' : '';
        return '<li><a href="' + e.link + '">' + esc_(e.subject) + '</a> — <span style="color:#666">' +
          esc_(shortFrom_(e.from)) + '</span>' + when + '</li>';
      }).join('') + '</ul>'
    ).join('');
}

function shortFrom_(from) {
  const m = String(from || '').match(/^\s*"?([^"<]+?)"?\s*</);
  return m ? m[1] : String(from || '');
}

function geminiDigest_(s, entries) {
  const prompt = 'You write a short, warm, scannable morning email digest for one person. ' +
    'Input is a list of emails still in their inbox (sender, subject, category, action taken, one-line summary, link). ' +
    'Output ONLY HTML fragments (no <html>/<body>), using <h3> section headings and <ul><li> items. ' +
    'Sections: 1) "Needs you" — anything that needs their attention or was an ALERT, 2) "Worth knowing" — everything else. Include every email; omit a section if empty. ' +
    'Make each item\'s subject a link using exactly its "link" value: <a href="LINK">subject</a>. One line per item. Be brief.\n\n' +
    JSON.stringify(entries.map(e => ({ from: e.from, subject: e.subject, category: e.category, action: e.action, summary: e.summary, link: e.link })));
  try {
    const text = geminiText_(s, prompt);
    return text ? text.replace(/^```html\s*|```$/g, '').trim() : null;
  } catch (e) { Logger.log('Digest fallback: ' + e); return null; }
}

// ---------- gemini ----------
function askGemini_(s, rules, items) {
  const BATCH = 15;
  const out = {};
  for (let i = 0; i < items.length; i += BATCH) {
    const slice = items.slice(i, i + BATCH).map(it => ({
      id: it.id, from: it.from, subject: it.subject, date: it.date,
      messages_in_thread: it.messages_in_thread, i_replied_already: it.i_replied_already, body: it.body,
    }));
    const prompt =
      'You are a careful personal email assistant. Apply the user\'s rules to each email and decide what to do.\n' +
      'If no rule clearly applies, do nothing (empty labels, no archive, no alert) — when in doubt, leave it in the inbox.\n' +
      'Only use archive_after_days when a rule says to archive after a delay. Only set alert when a rule says to alert/notify immediately.\n' +
      'Labels should be short Title Case words (e.g. "Newsletter", "Receipts"). needs_me = the email is waiting on a reply or action from the user.\n' +
      'category = one short word like Receipt, Newsletter, Promotion, Personal, Work, Finance, Notification, Other.\n' +
      'summary = one plain sentence of what the email is about, for a morning digest.\n\n' +
      'USER RULES:\n' + rules.map((r, n) => (n + 1) + '. ' + r).join('\n') +
      '\n\nEMAILS (JSON):\n' + JSON.stringify(slice);

    const schema = {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          id: { type: 'STRING' },
          category: { type: 'STRING' },
          labels: { type: 'ARRAY', items: { type: 'STRING' } },
          needs_me: { type: 'BOOLEAN' },
          star: { type: 'BOOLEAN' },
          archive_now: { type: 'BOOLEAN' },
          archive_after_days: { type: 'INTEGER' },
          alert: { type: 'BOOLEAN' },
          alert_reason: { type: 'STRING' },
          summary: { type: 'STRING' },
        },
        required: ['id', 'category', 'labels', 'needs_me', 'archive_now', 'archive_after_days', 'alert', 'summary'],
      },
    };
    const text = geminiText_(s, prompt, schema);
    let arr = [];
    try { arr = JSON.parse(text); } catch (e) { Logger.log('Bad JSON from Gemini: ' + text); }
    (Array.isArray(arr) ? arr : []).forEach(d => { if (d && d.id) out[d.id] = d; });
    if (i + BATCH < items.length) Utilities.sleep(1500);
  }
  return out;
}

function geminiText_(s, prompt, schema) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + (s.MODEL || 'gemini-3.8-flash') +
    ':generateContent?key=' + encodeURIComponent(s.GEMINI_API_KEY);
  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2 },
  };
  if (schema) {
    payload.generationConfig.responseMimeType = 'application/json';
    payload.generationConfig.responseSchema = schema;
  }
  let res, code;
  for (let attempt = 0; attempt < 3; attempt++) {
    res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true });
    code = res.getResponseCode();
    if (code === 200) break;
    if (code === 429 || code >= 500) { Utilities.sleep(4000 * (attempt + 1)); continue; }
    throw new Error('Gemini ' + code + ': ' + res.getContentText().slice(0, 300));
  }
  if (code !== 200) throw new Error('Gemini ' + code + ' after retries');
  const data = JSON.parse(res.getContentText());
  const parts = (((data.candidates || [])[0] || {}).content || {}).parts || [];
  return parts.map(p => p.text || '').join('');
}

// ---------- helpers ----------
function isOwnEmail_(subject) {
  return /^(☕ Morning digest|🥪 Midday digest|🌙 Evening digest|🔔 |Inbox Assistant is set up|Inbox Assistant link request|Inbox Assistant: )/.test(String(subject || ''));
}

/** Turns every non-ASCII character (emoji, accents) into an HTML entity so no mail app can garble it. */
function htmlSafe_(html) {
  return String(html).replace(/[^\x00-\x7F]/gu, c => '&#x' + c.codePointAt(0).toString(16) + ';');
}

function cleanBody_(txt) {
  return String(txt || '')
    .replace(/\r/g, '')
    .replace(/https?:\/\/\S{60,}/g, '[link]')   // drop giant tracking URLs
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim()
    .slice(0, 1500);
}
function threadLink_(id, account) {
  // authuser picks the right signed-in account, so links from a second account open in that account
  const who = account || Session.getActiveUser().getEmail();
  return 'https://mail.google.com/mail/?authuser=' + encodeURIComponent(who) + '#all/' + id;
}
function esc_(s) { return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

/** Manual helpers you can run from the editor */
function sendDigestNow() { sendDigest(); }
function forgetProcessedLabel() {
  // Lets the assistant re-evaluate everything in the inbox (useful after changing rules).
  const l = GmailApp.getUserLabelByName(LABEL_PROCESSED);
  if (l) GmailApp.search('label:"' + LABEL_PROCESSED + '"').forEach(t => t.removeLabel(l));
}

// ---------- shared-code (library) support ----------
// Other accounts can use this project as an Apps Script library: their own project holds only a short stub
// that calls run('<function>', {ScriptApp, PropertiesService}). Passing those two services in keeps each
// account's triggers and stored settings in ITS OWN project, exactly as before.
let HOST_ = null;
function SA_() { return (HOST_ && HOST_.ScriptApp) || ScriptApp; }
function PS_() { return (HOST_ && HOST_.PropertiesService) || PropertiesService; }
function viaLibrary_() { return !!(HOST_ && HOST_.ScriptApp); }

/** Entry point: Loader.gs (and the older library stub) call every function through here. */
function run(name, host, args) {
  HOST_ = host || null;
  const fn = ENTRY_[name || 'runNow'];
  if (!fn) throw new Error('Unknown function: ' + name);
  return fn.apply(null, args || []);
}
const ENTRY_ = { setup, runNow, processInbox, processPending, sendDigest, sendDigestNow, reinstallTriggers,
  forgetProcessedLabel, refreshHowTo, restyleTabs, repairTabs, recheckFromMenu, updateCode, restorePreviousCode,
  updateNowFromMenu, openAppFromMenu, publishApp, onSheetEdit, doGet,
  appState, appSetSetting, appRuleSet, appRuleAdd, appRuleRemove, appEmail, appEmailAction, appCheckNow,
  appSendDigestNow, appUpdateNow, appUndoUpdate, appLinkCode, appJoinMain, appUnlink };

// ---------- app page (generated from App.html by build.py — edit App.html, not this) ----------
const APP_HTML = "<!doctype html>\n<html lang=\"en\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap\">\n<style>\n/* Design B · glanceable dashboard. Dark header on Today, teal accent, cards on a cool grey ground. */\n:root {\n  --ground: #f3f5f5; --card: #ffffff; --line: #e1e6e6; --line-2: #edf1f1; --ink: #13201f; --muted: #55625f; --faint: #6b7a79;\n  --head: #10292b; --head-ink: #e9f4f3; --head-muted: #a9c9c6; --head-line: #2f5b5d; --head-chip: #173b3d;\n  --accent: #0e7c86; --accent-ink: #ffffff; --accent-soft: #e3eeee; --mint: #3fc1b0;\n  --alert: #d14b2f; --alert-ink: #b4441f; --needs: #e0a100; --peach: #ffb86b; --off: #c9d4d3;\n  --display: \"Space Grotesk\", system-ui, sans-serif; --body: Inter, system-ui, -apple-system, \"Segoe UI\", sans-serif;\n}\n@media (prefers-color-scheme: dark) {\n  :root {\n    --ground: #0d1717; --card: #152222; --line: #233434; --line-2: #1c2b2b; --ink: #e4efee; --muted: #9fb3b1; --faint: #86a09d;\n    --head: #0a1d1f; --head-line: #2a4c4e; --head-chip: #12302f; --accent: #3fb3bb; --accent-ink: #062224; --accent-soft: #1d3434;\n    --off: #3b4f4f; color-scheme: dark;\n  }\n}\n* { box-sizing: border-box; }\nhtml, body { margin: 0; height: 100%; }\nbody { background: var(--ground); color: var(--ink); font: 15px/1.45 var(--body); -webkit-font-smoothing: antialiased; }\nbutton, input, textarea { font: inherit; color: inherit; }\nbutton { cursor: pointer; }\nbutton:focus-visible, a:focus-visible, input:focus-visible, textarea:focus-visible { outline: 3px solid var(--mint); outline-offset: 2px; }\na { color: var(--accent); }\n.shell { max-width: 560px; margin: 0 auto; min-height: 100%; display: flex; flex-direction: column; }\n.page { flex: 1; padding: 14px 16px 96px; display: flex; flex-direction: column; gap: 14px; }\nh1, h2 { font-family: var(--display); margin: 0; text-wrap: balance; }\nh1 { font-size: 24px; font-weight: 700; } h2 { font-size: 16px; font-weight: 600; }\n.top { padding: 18px 16px 6px; display: flex; align-items: center; gap: 8px; }\n.top h1 { flex: 1; }\n.back { width: 44px; height: 44px; margin-left: -10px; display: grid; place-items: center; border: 0; background: none; color: var(--ink); border-radius: 12px; }\n/* Today header */\n.hero { background: var(--head); color: var(--head-ink); padding: 18px 16px 16px; display: flex; flex-direction: column; gap: 14px; }\n.hero .row { display: flex; align-items: center; justify-content: space-between; gap: 10px; }\n.eyebrow { font-size: 12px; letter-spacing: .08em; color: #8fb8b5; text-transform: uppercase; }\n.hero h1 { margin-top: 2px; }\n.counts { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; }\n.counts b { display: block; font: 700 30px/1.1 var(--display); font-variant-numeric: tabular-nums; }\n.counts span { font-size: 12px; color: var(--head-muted); }\n.counts .hot { color: var(--peach); }\n.livebtn { min-height: 40px; border: 1px solid var(--head-line); background: var(--head-chip); color: var(--head-ink); border-radius: 999px; padding: 0 6px 0 12px; display: flex; align-items: center; gap: 8px; font-weight: 600; font-size: 13px; }\n.chips { display: flex; gap: 6px; flex-wrap: wrap; }\n.chip { min-height: 34px; padding: 0 12px; border-radius: 999px; border: 1px solid var(--off); background: var(--card); color: var(--ink); font-size: 13px; font-weight: 500; display: inline-flex; align-items: center; gap: 6px; }\n.chip[aria-pressed=\"true\"] { background: var(--head); border-color: var(--head); color: #fff; font-weight: 600; }\n.hero .chip { background: transparent; border-color: var(--head-line); color: #cfe3e1; min-height: 32px; }\n.hero .chip[aria-pressed=\"true\"] { background: var(--head-ink); color: var(--head); border-color: var(--head-ink); }\n.chip.warn { border-style: dashed; }\n/* switch */\n.sw { width: 44px; height: 26px; border: 0; border-radius: 999px; background: var(--off); position: relative; flex: none; padding: 0; }\n.sw::after { content: \"\"; position: absolute; top: 3px; left: 3px; width: 20px; height: 20px; border-radius: 50%; background: #fff; transition: left .15s; }\n.sw[aria-checked=\"true\"] { background: var(--accent); }\n.sw[aria-checked=\"true\"]::after { left: 21px; }\n.livebtn .sw { width: 34px; height: 20px; background: #4a6466; }\n.livebtn .sw::after { width: 16px; height: 16px; top: 2px; left: 2px; }\n.livebtn .sw[aria-checked=\"true\"] { background: var(--mint); }\n.livebtn .sw[aria-checked=\"true\"]::after { left: 16px; }\n/* cards */\n.card { background: var(--card); border: 1px solid var(--line); border-radius: 14px; }\n.list { display: flex; flex-direction: column; gap: 8px; }\n.need { display: flex; gap: 12px; padding: 12px 14px 12px 12px; text-align: left; width: 100%; color: inherit; }\n.need .bar { width: 4px; border-radius: 2px; background: var(--needs); flex: none; }\n.need.alert .bar { background: var(--alert); }\n.need .meta { display: flex; justify-content: space-between; align-items: center; gap: 8px; font-size: 12px; font-weight: 600; color: var(--accent); }\n.need.alert .meta { color: var(--alert-ink); }\n.need .meta .r { display: flex; gap: 6px; align-items: center; color: var(--faint); font-weight: 500; white-space: nowrap; }\n.tag { padding: 1px 7px; border-radius: 999px; background: var(--accent-soft); color: var(--ink); font-size: 12px; font-weight: 600; }\n.need .t { display: block; font-weight: 600; font-size: 15px; margin-top: 3px; }\n.need .s { display: block; font-size: 13px; color: var(--muted); }\n.empty { padding: 16px; text-align: center; color: var(--muted); font-size: 14px; }\n.actions2 { display: flex; gap: 8px; }\n.btn { min-height: 48px; border-radius: 14px; border: 1px solid var(--off); background: var(--card); font-weight: 600; font-size: 15px; padding: 0 16px; flex: 1; display: inline-flex; align-items: center; justify-content: center; gap: 8px; text-decoration: none; color: var(--ink); }\n.btn.primary { background: var(--accent); border-color: var(--accent); color: var(--accent-ink); }\n.btn.soft { background: var(--accent-soft); border-color: var(--accent-soft); color: var(--accent); }\n.btn.small { min-height: 40px; font-size: 14px; border-radius: 10px; }\n.btn:disabled { opacity: .55; }\n.note { font-size: 12px; color: var(--muted); text-align: center; }\n/* rules */\n.rule { padding: 12px 14px; display: flex; gap: 12px; align-items: flex-start; }\n.rule .k { font-size: 12px; font-weight: 600; color: var(--accent); text-transform: uppercase; letter-spacing: .02em; }\n.rule .x { font-size: 14px; line-height: 1.4; }\n.rule.off { color: var(--faint); } .rule.off .k { color: var(--faint); }\n.linkbtn { border: 0; background: none; color: var(--accent); font-weight: 600; font-size: 13px; padding: 4px 0; }\n.addbtn { min-height: 36px; padding: 0 12px; border-radius: 10px; border: 1px solid var(--off); background: var(--card); font-weight: 600; font-size: 13px; flex: none; }\n.addbtn[aria-pressed=\"true\"] { background: var(--accent-soft); border-color: var(--accent-soft); color: var(--accent); }\n/* settings */\n.group { overflow: hidden; }\n.item { display: flex; align-items: center; gap: 12px; padding: 12px 14px; border-bottom: 1px solid var(--line-2); }\n.item:last-child { border-bottom: 0; }\n.item .lbl { flex: 1; min-width: 0; } .item .lbl b { display: block; font-weight: 600; } .item .lbl span { font-size: 13px; color: var(--muted); }\n.block { padding: 12px 14px; border-bottom: 1px solid var(--line-2); }\n.block:last-child { border-bottom: 0; }\n.seg { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); background: var(--accent-soft); border-radius: 10px; padding: 3px; margin-top: 8px; }\n.seg button { min-height: 36px; border: 0; border-radius: 8px; background: transparent; color: var(--muted); font-size: 13px; font-weight: 500; }\n.seg button[aria-checked=\"true\"] { background: var(--card); color: var(--ink); font-weight: 600; box-shadow: 0 1px 2px rgba(16,41,43,.15); }\n.stepper { display: flex; align-items: center; border: 1px solid var(--off); border-radius: 10px; }\n.stepper button { width: 40px; height: 40px; border: 0; background: none; font-weight: 600; font-size: 18px; }\n.stepper span { min-width: 58px; text-align: center; font-weight: 600; font-size: 14px; font-variant-numeric: tabular-nums; }\n.dot { width: 8px; height: 8px; border-radius: 50%; background: var(--accent); flex: none; }\n.dot.bad { background: var(--alert); }\n/* forms */\nlabel.f { font-size: 14px; font-weight: 600; display: block; margin-bottom: 6px; }\n.input, textarea.input { width: 100%; border: 1px solid var(--off); border-radius: 12px; padding: 12px; background: var(--card); font-size: 15px; }\ntextarea.input { resize: vertical; min-height: 110px; line-height: 1.5; }\n.input:focus { border-color: var(--accent); outline: none; box-shadow: 0 0 0 3px var(--accent-soft); }\n/* link code */\n.codecard { background: var(--head); color: var(--head-ink); border-radius: 16px; padding: 18px; display: flex; flex-direction: column; gap: 10px; }\n.code { font: 700 40px var(--display); letter-spacing: .12em; font-variant-numeric: tabular-nums; }\n/* email */\n.mail { padding: 14px; font-size: 15px; line-height: 1.55; white-space: pre-wrap; overflow-wrap: anywhere; }\n.avatar { width: 36px; height: 36px; border-radius: 50%; background: var(--head); color: var(--head-ink); display: grid; place-items: center; font-weight: 600; flex: none; }\n.dock { position: sticky; bottom: 0; background: var(--card); border-top: 1px solid var(--line); padding: 12px 16px calc(14px + env(safe-area-inset-bottom, 0px)); display: flex; flex-direction: column; gap: 8px; }\n.three { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }\n.three .btn { flex-direction: column; gap: 2px; font-size: 14px; min-height: 52px; }\n/* nav */\nnav.tabs { position: fixed; left: 0; right: 0; bottom: 0; background: var(--card); border-top: 1px solid var(--line); padding: 6px 0 calc(10px + env(safe-area-inset-bottom, 0px)); }\nnav.tabs .in { max-width: 560px; margin: 0 auto; display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); }\nnav.tabs button { border: 0; background: none; min-height: 48px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px; font-size: 11px; font-weight: 500; color: var(--muted); }\nnav.tabs button[aria-current=\"page\"] { color: var(--accent); font-weight: 600; }\n#toast { position: fixed; left: 50%; bottom: 84px; transform: translateX(-50%); background: var(--head); color: #fff; padding: 10px 16px; border-radius: 12px; font-size: 14px; max-width: calc(100% - 32px); box-shadow: 0 6px 20px rgba(0,0,0,.25); z-index: 20; }\n.loading { padding: 60px 16px; text-align: center; color: var(--muted); }\n@media (prefers-reduced-motion: reduce) { * { transition: none !important; } }\n</style>\n</head>\n<body>\n<div class=\"shell\" id=\"app\"><div class=\"loading\">Loading your inbox…</div></div>\n<div id=\"toast\" role=\"status\" hidden></div>\n<script>\nvar S = { view: 'all', tab: 'today', sub: null, d: null, filter: 'active', cleared: null, email: null, link: null, edit: null, joinMsg: '' };\n\nfunction call(name) {\n  var args = Array.prototype.slice.call(arguments, 1);\n  return new Promise(function (res, rej) {\n    var r = google.script.run.withSuccessHandler(res).withFailureHandler(function (e) { rej(e && e.message ? e : new Error(String(e))); });\n    r[name].apply(r, args);\n  });\n}\nfunction esc(s) { return String(s == null ? '' : s).replace(/[&<>\"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '\"': '&quot;', \"'\": '&#39;' }[c]; }); }\nfunction toast(msg) { var t = document.getElementById('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast.h); toast.h = setTimeout(function () { t.hidden = true; }, 3500); }\nfunction fail(e) { toast(e && e.message ? e.message : 'Something went wrong.'); }\nfunction when(iso) {\n  if (!iso) return '';\n  var d = new Date(iso), now = new Date(), y = new Date(now); y.setDate(now.getDate() - 1);\n  if (d.toDateString() === now.toDateString()) return d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });\n  if (d.toDateString() === y.toDateString()) return 'Yesterday';\n  return d.toLocaleDateString([], { weekday: 'short' });\n}\nfunction hourLabel(h) { return h === 0 ? '12 AM' : h < 12 ? h + ' AM' : h === 12 ? '12 PM' : (h - 12) + ' PM'; }\nvar I = {\n  today: '<svg width=\"22\" height=\"22\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" aria-hidden=\"true\"><rect x=\"3\" y=\"3\" width=\"7\" height=\"9\" rx=\"1.5\"/><rect x=\"14\" y=\"3\" width=\"7\" height=\"5\" rx=\"1.5\"/><rect x=\"14\" y=\"12\" width=\"7\" height=\"9\" rx=\"1.5\"/><rect x=\"3\" y=\"16\" width=\"7\" height=\"5\" rx=\"1.5\"/></svg>',\n  rules: '<svg width=\"22\" height=\"22\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" aria-hidden=\"true\"><path d=\"M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01\"/></svg>',\n  settings: '<svg width=\"22\" height=\"22\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" aria-hidden=\"true\"><circle cx=\"12\" cy=\"12\" r=\"3\"/><path d=\"M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9l2.1 2.1M17 17l2.1 2.1M4.9 19.1L7 17M17 7l2.1-2.1\"/></svg>',\n  back: '<svg width=\"22\" height=\"22\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.2\" aria-hidden=\"true\"><path d=\"M15 18l-6-6 6-6\"/></svg>',\n  plus: '<svg width=\"18\" height=\"18\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.4\" aria-hidden=\"true\"><path d=\"M12 5v14M5 12h14\"/></svg>',\n  check: '<svg width=\"18\" height=\"18\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.4\" aria-hidden=\"true\"><path d=\"M20 6L9 17l-5-5\"/></svg>',\n  box: '<svg width=\"18\" height=\"18\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" aria-hidden=\"true\"><rect x=\"3\" y=\"4\" width=\"18\" height=\"4\" rx=\"1\"/><path d=\"M5 8v11h14V8M10 12h4\"/></svg>',\n  star: '<svg width=\"18\" height=\"18\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" aria-hidden=\"true\"><path d=\"M12 3l2.8 5.7 6.2.9-4.5 4.4 1.1 6.2L12 17.3 6.4 20.2l1.1-6.2L3 9.6l6.2-.9z\"/></svg>',\n  out: '<svg width=\"16\" height=\"16\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2.2\" aria-hidden=\"true\"><path d=\"M7 17L17 7M8 7h9v9\"/></svg>'\n};\n\nfunction load(view) {\n  if (view) S.view = view;\n  return call('appState', S.view).then(function (d) { S.d = d; render(); }).catch(function (e) {\n    document.getElementById('app').innerHTML = '<div class=\"loading\">Couldn\\'t load: ' + esc(e.message) + '<br><br><button class=\"btn small\" onclick=\"load()\">Try again</button></div>';\n  });\n}\nfunction loadAcct(i) { var a = S.d.accounts[i]; if (a) load(a.key); }\nfunction catList() { return Object.keys(S.d.cleared).sort(); }\nfunction multi() { return S.d && S.d.accounts.length > 1; }\n\nfunction render() {\n  var d = S.d, html = '';\n  if (S.sub === 'email') html = viewEmail();\n  else if (S.sub === 'write') html = viewWrite();\n  else if (S.sub === 'cleared') html = viewCleared();\n  else if (S.sub === 'link') html = viewLink();\n  else if (S.sub === 'join') html = viewJoin();\n  else if (S.sub === 'manage') html = viewManage();\n  else if (S.tab === 'rules') html = viewRules();\n  else if (S.tab === 'settings') html = viewSettings();\n  else html = viewToday();\n  if (!S.sub || S.sub === 'cleared' || S.sub === 'link' || S.sub === 'join' || S.sub === 'manage') html += nav();\n  document.getElementById('app').innerHTML = html;\n  window.scrollTo(0, 0);\n}\nfunction nav() {\n  function b(id, label) { return '<button type=\"button\" onclick=\"go(\\'' + id + '\\')\"' + (S.tab === id && !S.sub ? ' aria-current=\"page\"' : '') + '>' + I[id] + label + '</button>'; }\n  return '<nav class=\"tabs\" aria-label=\"Main\"><div class=\"in\">' + b('today', 'Today') + b('rules', 'Rules') + b('settings', 'Settings') + '</div></nav>';\n}\nfunction go(tab) { S.tab = tab; S.sub = null; render(); }\nfunction topBar(title, backTo) { return '<div class=\"top\"><button class=\"back\" type=\"button\" aria-label=\"Back\" onclick=\"' + backTo + '\">' + I.back + '</button><h1 style=\"font-size:22px\">' + esc(title) + '</h1></div>'; }\n\n/* ---------- Today ---------- */\nfunction viewToday() {\n  var d = S.d, c = d.counts, s = d.settings || {};\n  var title = c.needs ? (c.needs + (c.needs === 1 ? ' needs you' : ' need you')) : 'Inbox is calm';\n  var date = new Date().toLocaleDateString([], { weekday: 'long', month: 'short', day: 'numeric' });\n  var h = '<header class=\"hero\"><div class=\"row\"><div><div class=\"eyebrow\">' + esc(date) + '</div><h1>' + esc(title) + '</h1></div>' +\n    '<button class=\"livebtn\" type=\"button\" onclick=\"toggleLive()\" title=\"Off = test mode\">' + (s.live ? 'Live' : 'Test') +\n    '<span class=\"sw\" role=\"switch\" aria-checked=\"' + !!s.live + '\" aria-label=\"Live mode for ' + esc(d.editing.label) + '\"></span></button></div>';\n  if (multi()) {\n    h += '<div class=\"chips\" role=\"tablist\" aria-label=\"Account\"><button class=\"chip\" role=\"tab\" aria-pressed=\"' + (S.view === 'all') + '\" onclick=\"load(\\'all\\')\">All' + (S.view === 'all' && c.needs ? ' · ' + c.needs : '') + '</button>';\n    d.accounts.forEach(function (a, i) { h += '<button class=\"chip' + (a.error ? ' warn' : '') + '\" role=\"tab\" aria-pressed=\"' + (S.view === a.key) + '\" onclick=\"loadAcct(' + i + ')\">' + esc(a.label) + '</button>'; });\n    h += '</div>';\n  }\n  h += '<div class=\"counts\"><div><b class=\"hot\">' + c.needs + '</b><span>need you</span></div><div><b>' + c.sorted + '</b><span>sorted today</span></div><div><b>' + c.cleared + '</b><span>cleared out</span></div></div></header>';\n  h += '<main class=\"page\"><section class=\"list\" aria-labelledby=\"nh\"><h2 id=\"nh\">Needs you</h2>';\n  if (!d.needs.length) h += '<div class=\"card empty\">Nothing waiting on you.</div>';\n  d.needs.forEach(function (n, i) {\n    var label = n.alert ? 'ALERT · ' + n.from : (n.category ? n.category.toUpperCase() + ' · ' : '') + n.from;\n    h += '<button type=\"button\" class=\"card need' + (n.alert ? ' alert' : '') + '\" onclick=\"openNeed(' + i + ')\"><span class=\"bar\"></span><span style=\"flex:1;min-width:0\">' +\n      '<span class=\"meta\"><span>' + esc(label) + '</span><span class=\"r\">' + (multi() ? '<span class=\"tag\">' + esc(n.account) + '</span>' : '') + esc(when(n.when)) + '</span></span>' +\n      '<span class=\"t\">' + esc(n.subject) + '</span>' + (n.summary ? '<span class=\"s\">' + esc(n.summary) + '</span>' : '') + '</span></button>';\n  });\n  h += '</section>';\n  var cats = Object.keys(d.cleared);\n  h += '<section class=\"list\" aria-labelledby=\"ch\"><h2 id=\"ch\">Cleared out today</h2>';\n  if (!cats.length) h += '<div class=\"note\" style=\"text-align:left\">Nothing archived yet today.</div>';\n  else { h += '<div class=\"chips\">'; catList().forEach(function (k, i) { h += '<button class=\"chip\" style=\"background:var(--accent-soft);border-color:var(--accent-soft)\" onclick=\"openCleared(' + i + ')\"><b>' + d.cleared[k].length + '</b> ' + esc(k) + '</button>'; }); h += '</div>'; }\n  h += '</section>';\n  var own = S.view === 'all' || d.editing.own;\n  h += '<div class=\"actions2\" style=\"margin-top:auto\"><button class=\"btn primary\" type=\"button\" onclick=\"checkNow(this)\">Check now</button>' +\n    (own ? '<button class=\"btn\" type=\"button\" onclick=\"digestNow(this)\">Digest now</button>' : '') + '</div>';\n  h += '<div class=\"note\">' + esc(s.lastCheck || 'Checks every 10 minutes') + '</div></main>';\n  return h;\n}\nfunction toggleLive() {\n  var s = S.d.settings; var to = !s.live;\n  call('appSetSetting', S.d.editing.key, 'DRY_RUN', to ? 'FALSE' : 'TRUE').then(function () { s.live = to; render(); toast(to ? S.d.editing.label + ' is live — it will act on mail.' : 'Test mode — it only logs what it would do.'); }).catch(fail);\n}\nfunction openNeed(i) {\n  var n = S.d.needs[i];\n  if ((S.d.settings || {}).tapAction === 'gmail') { window.open(n.link, '_blank'); return; }\n  S.sub = 'email'; S.email = { loading: true, n: n }; render();\n  call('appEmail', n.acct, n.id).then(function (e) { S.email = { e: e, n: n }; if (S.sub === 'email') render(); }).catch(function (e) { S.email = { n: n, err: e.message }; render(); });\n}\nfunction checkNow(btn) { btn.disabled = true; btn.textContent = 'Checking…'; call('appCheckNow', S.view === 'all' ? 'self' : S.view).then(function (m) { toast(m); return load(); }).catch(fail).then(function () { btn.disabled = false; }); }\nfunction digestNow(btn) { btn.disabled = true; call('appSendDigestNow').then(toast).catch(fail).then(function () { btn.disabled = false; }); }\n\n/* ---------- Email ---------- */\nfunction viewEmail() {\n  var x = S.email || {}, n = x.n || {}, e = x.e;\n  var h = '<div class=\"top\"><button class=\"back\" type=\"button\" aria-label=\"Back to Today\" onclick=\"go(\\'today\\')\">' + I.back + '</button><span style=\"flex:1;font-size:13px;color:var(--muted)\">Needs you</span>' + (multi() ? '<span class=\"tag\">' + esc(n.account) + '</span>' : '') + '</div><main class=\"page\" style=\"padding-bottom:20px\">';\n  h += '<div><div style=\"font-size:12px;font-weight:600;color:' + (n.alert ? 'var(--alert-ink)' : 'var(--accent)') + '\">' + (n.alert ? 'ALERT' : esc((n.category || 'NEEDS YOU').toUpperCase())) + '</div><h1 style=\"font-size:22px;line-height:1.25;margin-top:4px\">' + esc(n.subject) + '</h1></div>';\n  var from = e ? e.from : n.from;\n  h += '<div style=\"display:flex;align-items:center;gap:10px\"><span class=\"avatar\">' + esc(String(from || '?').replace(/^[\"\\s]+/, '').charAt(0).toUpperCase()) + '</span><div style=\"flex:1;min-width:0\"><div style=\"font-weight:600;font-size:14px;overflow-wrap:anywhere\">' + esc(from) + '</div><div style=\"font-size:12px;color:var(--muted)\">' + esc(when(e ? e.date : n.when)) + '</div></div></div>';\n  if (x.loading) h += '<div class=\"card mail\" style=\"color:var(--muted)\">Loading…</div>';\n  else if (x.err) h += '<div class=\"card mail\">' + esc(x.err) + '</div>';\n  else if (e && e.own && e.body) h += '<div class=\"card mail\">' + esc(e.body) + '</div>';\n  else h += '<div class=\"card mail\"><b>Summary:</b> ' + esc((e && e.summary) || n.summary || '') + '\\n\\n<span style=\"color:var(--muted)\">This email is in your ' + esc(n.account) + ' account. Open it in Gmail to read the whole thing.</span></div>';\n  h += '</main><div class=\"dock\"><div class=\"three\">' +\n    '<button class=\"btn primary\" type=\"button\" onclick=\"mailAct(\\'done\\')\">' + I.check + 'Done</button>' +\n    '<button class=\"btn\" type=\"button\" onclick=\"mailAct(\\'archive\\')\">' + I.box + 'Archive</button>' +\n    '<button class=\"btn\" type=\"button\" onclick=\"mailAct(\\'star\\')\">' + I.star + ((e && e.starred) ? 'Unstar' : 'Star') + '</button></div>' +\n    '<a class=\"btn soft small\" href=\"' + esc(n.link) + '\" target=\"_blank\" rel=\"noopener\">Reply in Gmail ' + I.out + '</a></div>';\n  return h;\n}\nfunction mailAct(a) {\n  var n = S.email.n;\n  call('appEmailAction', n.acct, n.id, a).then(function (m) { toast(m); if (a === 'star') { if (S.email.e) S.email.e.starred = !S.email.e.starred; render(); return; } S.sub = null; S.tab = 'today'; return load(); }).catch(fail);\n}\n\n/* ---------- Cleared ---------- */\nfunction openCleared(i) { S.sub = 'cleared'; S.cleared = catList()[i]; render(); }\nfunction viewCleared() {\n  var d = S.d, cats = catList(), cur = S.cleared || cats[0];\n  var h = topBar('Cleared out today', \"go('today')\") + '<main class=\"page\"><div class=\"chips\">';\n  cats.forEach(function (k, i) { h += '<button class=\"chip\" aria-pressed=\"' + (k === cur) + '\" onclick=\"openCleared(' + i + ')\">' + esc(k) + ' · ' + d.cleared[k].length + '</button>'; });\n  h += '</div><div class=\"card group\">';\n  (d.cleared[cur] || []).forEach(function (x) {\n    h += '<a class=\"item\" href=\"' + esc(x.link) + '\" target=\"_blank\" rel=\"noopener\" style=\"text-decoration:none;color:inherit\"><span class=\"lbl\"><b style=\"font-size:15px\">' + esc(x.subject) + '</b><span>' + esc(x.from) + (multi() ? ' · ' + esc(x.account) : '') + '</span></span><span style=\"font-size:12px;color:var(--faint)\">' + esc(when(x.when)) + '</span></a>';\n  });\n  h += '</div><div class=\"note\">Archived, not deleted · tap any to open it in Gmail</div>' +\n    '<div class=\"card\" style=\"border-style:dashed;padding:12px 14px;display:flex;align-items:center;gap:10px\"><span style=\"flex:1;font-size:14px\">Something here you wanted to keep?</span><button class=\"btn soft small\" style=\"flex:none\" onclick=\"S.sub=\\'write\\';render()\">Write a rule</button></div></main>';\n  return h;\n}\n\n/* ---------- Rules ---------- */\nfunction editChips(withAll) {\n  if (!multi()) return '';\n  var h = '<div class=\"chips\" aria-label=\"Account\">';\n  S.d.accounts.forEach(function (a, i) { if (a.error) return; h += '<button class=\"chip\" aria-pressed=\"' + (S.d.editing.key === a.key) + '\" onclick=\"loadAcct(' + i + ')\">' + esc(a.label) + '</button>'; });\n  return h + '</div>';\n}\nfunction viewRules() {\n  var d = S.d, on = d.rules.filter(function (r) { return r.on; }), off = d.rules.filter(function (r) { return !r.on; });\n  var h = '<div class=\"top\" style=\"flex-direction:column;align-items:stretch;gap:10px\"><div style=\"display:flex;align-items:baseline;justify-content:space-between\"><h1>Rules</h1><span style=\"font-size:13px;color:var(--muted)\">' + on.length + ' active' + (multi() ? ' · ' + esc(d.editing.label) : '') + '</span></div>' + editChips() +\n    '<div class=\"chips\" aria-label=\"Show\"><button class=\"chip\" aria-pressed=\"' + (S.filter === 'active') + '\" onclick=\"setFilter(\\'active\\')\">Active · ' + on.length + '</button>' +\n    (off.length ? '<button class=\"chip\" aria-pressed=\"' + (S.filter === 'off') + '\" onclick=\"setFilter(\\'off\\')\">Off · ' + off.length + '</button>' : '');\n  d.library.forEach(function (g, gi) { h += '<button class=\"chip\" aria-pressed=\"' + (S.filter === 'lib' + gi) + '\" onclick=\"setFilter(\\'lib' + gi + '\\')\">' + esc(g.type) + '</button>'; });\n  h += '</div></div><main class=\"page\" style=\"padding-top:6px\"><div class=\"list\">';\n  if (S.filter === 'active' || S.filter === 'off') {\n    var list = S.filter === 'active' ? on : off;\n    if (!list.length) h += '<div class=\"card empty\">' + (S.filter === 'active' ? 'No rules on yet. Pick some from the account types above.' : 'No rules are off.') + '</div>';\n    list.forEach(function (r) {\n      var idx = d.rules.indexOf(r);\n      h += '<div class=\"card rule' + (r.on ? '' : ' off') + '\"><div style=\"flex:1\"><div class=\"k\">' + esc(r.type ? r.type : 'Your rule') + (r.on ? '' : ' · off') + '</div><div class=\"x\">' + esc(r.text) + '</div>' +\n        (r.on ? '' : '<button class=\"linkbtn\" onclick=\"removeRule(' + idx + ')\">Remove</button>') + '</div>' +\n        '<button class=\"sw\" role=\"switch\" aria-checked=\"' + r.on + '\" aria-label=\"' + (r.on ? 'Turn off' : 'Turn on') + '\" onclick=\"toggleRule(' + idx + ')\"></button></div>';\n    });\n  } else {\n    var gi = Number(String(S.filter).slice(3)), g = d.library[gi];\n    (g ? g.rules : []).forEach(function (r, i) {\n      h += '<div class=\"card rule\"><div style=\"flex:1\"><div class=\"k\">' + esc(g.type) + '</div><div class=\"x\">' + esc(r.text) + '</div></div>' +\n        '<button class=\"addbtn\" aria-pressed=\"' + r.added + '\" onclick=\"libToggle(' + gi + ',' + i + ')\">' + (r.added ? 'Added ✓' : 'Add') + '</button></div>';\n    });\n  }\n  h += '</div><button class=\"btn primary\" style=\"flex:none;margin-top:4px\" onclick=\"S.sub=\\'write\\';render()\">' + I.plus + 'Write a rule</button></main>';\n  return h;\n}\nfunction setFilter(f) { S.filter = f; render(); }\nfunction toggleRule(i) { var r = S.d.rules[i]; r.on = !r.on; render(); call('appRuleSet', S.d.editing.key, r.text, r.on).catch(function (e) { r.on = !r.on; render(); fail(e); }); }\nfunction removeRule(i) { var r = S.d.rules[i]; call('appRuleRemove', S.d.editing.key, r.text).then(function () { toast('Removed.'); return load(); }).catch(fail); }\nfunction libToggle(gi, i) {\n  var g = S.d.library[gi], r = g.rules[i];\n  var p = r.added ? call('appRuleRemove', S.d.editing.key, r.text) : call('appRuleAdd', S.d.editing.key, r.text);\n  r.added = !r.added; render();\n  p.then(function () { return load(); }).catch(function (e) { r.added = !r.added; render(); fail(e); });\n}\n\n/* ---------- Write a rule ---------- */\nfunction viewWrite() {\n  return topBar('Write a rule', \"S.sub=null;render()\") + '<main class=\"page\">' +\n    '<div><label class=\"f\" for=\"rule\">In plain English, what should happen?</label><textarea class=\"input\" id=\"rule\" maxlength=\"500\" placeholder=\"e.g. Alert me about anything from my landlord, and keep it in the inbox.\"></textarea></div>' +\n    '<div class=\"chips\" aria-label=\"Ideas\"><button class=\"chip\" onclick=\"idea(\\'Label anything from … as “…” and keep it in the inbox.\\')\">Label anything from …</button><button class=\"chip\" onclick=\"idea(\\'Archive … after 2 days.\\')\">Archive … after 2 days</button><button class=\"chip\" onclick=\"idea(\\'Leave … alone.\\')\">Leave … alone</button><button class=\"chip\" onclick=\"idea(\\'Alert me right away about …\\')\">Alert me about …</button></div>' +\n    (multi() ? '<div class=\"note\" style=\"text-align:left\">Adds to ' + esc(S.d.editing.label) + '.</div>' : '') +\n    '<div style=\"margin-top:auto;display:flex;flex-direction:column;gap:8px\"><button class=\"btn primary\" style=\"flex:none\" id=\"addrule\" onclick=\"addRule()\">Add rule</button><div class=\"note\">Applies to new mail from the next check.</div></div></main>';\n}\nfunction idea(t) { var el = document.getElementById('rule'); el.value = t; el.focus(); var p = t.indexOf('…'); if (p >= 0) el.setSelectionRange(p, p + 1); }\nfunction addRule() {\n  var t = document.getElementById('rule').value.trim(), b = document.getElementById('addrule');\n  if (t.length < 8) { toast('Write the rule as a short sentence.'); return; }\n  b.disabled = true;\n  call('appRuleAdd', S.d.editing.key, t).then(function () { toast('Rule added.'); S.sub = null; S.tab = 'rules'; S.filter = 'active'; return load(); }).catch(function (e) { b.disabled = false; fail(e); });\n}\n\n/* ---------- Settings ---------- */\nfunction viewSettings() {\n  var d = S.d, s = d.settings || {}, self = d.self;\n  var hours = [6, 7, 8, 12, 17, 20]; (s.digestHours || []).forEach(function (x) { if (hours.indexOf(x) < 0) hours.push(x); }); hours.sort(function (a, b) { return a - b; });\n  var h = '<div class=\"top\" style=\"flex-direction:column;align-items:stretch;gap:10px\"><h1>Settings</h1>' + editChips() + '</div><main class=\"page\" style=\"padding-top:6px\">';\n  h += '<div class=\"card group\"><div class=\"item\"><div class=\"lbl\"><b>Live</b><span>Off = test mode: only logs what it would do.</span></div><button class=\"sw\" role=\"switch\" aria-checked=\"' + !!s.live + '\" aria-label=\"Live\" onclick=\"toggleLive()\"></button></div>';\n  if (s.sendsDigest) {\n    h += '<div class=\"block\"><b style=\"font-weight:600\">Digest times</b><div class=\"chips\" style=\"margin-top:8px\">';\n    hours.forEach(function (x) { h += '<button class=\"chip\" aria-pressed=\"' + ((s.digestHours || []).indexOf(x) >= 0) + '\" onclick=\"toggleHour(' + x + ')\">' + hourLabel(x) + '</button>'; });\n    h += '</div></div>';\n  } else {\n    h += '<div class=\"item\"><div class=\"lbl\"><b>Digest</b><span>Included in ' + esc(s.mainAccount || 'your main account') + '\\'s digest.</span></div></div>';\n  }\n  h += '<div class=\"block\"><b style=\"font-weight:600\">When I tap an email</b><div class=\"seg\" role=\"radiogroup\" aria-label=\"When I tap an email\">' +\n    '<button role=\"radio\" aria-checked=\"' + (s.tapAction !== 'gmail') + '\" onclick=\"setTap(\\'app\\')\">Show it here</button><button role=\"radio\" aria-checked=\"' + (s.tapAction === 'gmail') + '\" onclick=\"setTap(\\'gmail\\')\">Open in Gmail</button></div></div>';\n  h += '<div class=\"item\"><div class=\"lbl\"><b>Archive read mail after</b><span>Starred and “needs you” stay. 0 = never.</span></div><div class=\"stepper\"><button aria-label=\"Fewer days\" onclick=\"stepDays(-1)\">−</button><span>' + (s.archiveDays ? s.archiveDays + (s.archiveDays === 1 ? ' day' : ' days') : 'Never') + '</span><button aria-label=\"More days\" onclick=\"stepDays(1)\">+</button></div></div></div>';\n\n  h += '<section class=\"list\"><h2>Accounts</h2><div class=\"card group\">';\n  d.accounts.forEach(function (a) {\n    h += '<div class=\"item\"><span class=\"dot' + (a.error ? ' bad' : '') + '\"></span><div class=\"lbl\"><b>' + esc(a.label) + '</b><span>' + esc(a.error || (a.own ? 'This account' : a.email + ' · linked')) + '</span></div>' +\n      (a.error ? '' : '<button class=\"linkbtn\" onclick=\"manage(' + d.accounts.indexOf(a) + ')\">Manage</button>') + '</div>';\n  });\n  h += '<button class=\"item\" style=\"width:100%;border:0;background:none;color:var(--accent);font-weight:600;min-height:48px\" onclick=\"openLink()\">' + I.plus + 'Link another account</button>';\n  h += '<button class=\"item\" style=\"width:100%;border:0;border-top:1px solid var(--line-2);background:none;color:var(--accent);font-weight:600;min-height:48px\" onclick=\"S.sub=\\'join\\';render()\">Join a main account</button></div></section>';\n\n  if (!self.library) {\n    h += '<section class=\"list\"><h2>Updates</h2><div class=\"card group\"><div class=\"item\"><div class=\"lbl\"><b>Install updates automatically</b><span>Version ' + esc(self.version) + '</span></div><button class=\"sw\" role=\"switch\" aria-checked=\"' + !!s.autoUpdate + '\" aria-label=\"Automatic updates\" onclick=\"toggleAuto()\"></button></div>' +\n      (self.lastUpdate ? '<div class=\"block\" style=\"font-size:13px;color:var(--muted)\">' + esc(self.lastUpdate) + '</div>' : '') +\n      '<div class=\"block\" style=\"display:flex;gap:8px\"><button class=\"btn small\" onclick=\"updNow(this)\">Check for update</button><button class=\"btn small\" onclick=\"undoUpd(this)\">Undo last update</button></div></div></section>';\n  }\n  h += '<a class=\"btn small\" style=\"flex:none\" href=\"' + esc(self.sheetUrl) + '\" target=\"_blank\" rel=\"noopener\">Open my sheet ' + I.out + '</a></main>';\n  return h;\n}\nfunction setSetting(k, v, msg) { return call('appSetSetting', S.d.editing.key, k, v).then(function () { if (msg) toast(msg); return load(); }).catch(fail); }\nfunction toggleHour(x) {\n  var cur = (S.d.settings.digestHours || []).slice(), i = cur.indexOf(x);\n  if (i >= 0) { if (cur.length === 1) { toast('Keep at least one digest time.'); return; } cur.splice(i, 1); } else cur.push(x);\n  S.d.settings.digestHours = cur.sort(function (a, b) { return a - b; }); render();\n  setSetting('DIGEST_HOUR', cur.join(','), 'Digest times saved. They take effect within 10 minutes.');\n}\nfunction setTap(v) { S.d.settings.tapAction = v; render(); setSetting('TAP_ACTION', v); }\nfunction stepDays(n) { var v = Math.max(0, Math.min(60, (S.d.settings.archiveDays || 0) + n)); S.d.settings.archiveDays = v; render(); clearTimeout(stepDays.h); stepDays.h = setTimeout(function () { setSetting('ARCHIVE_READ_AFTER_DAYS', v); }, 600); }\nfunction toggleAuto() { var v = !S.d.settings.autoUpdate; S.d.settings.autoUpdate = v; render(); setSetting('AUTO_UPDATE', v ? 'TRUE' : 'FALSE'); }\nfunction updNow(b) { b.disabled = true; b.textContent = 'Checking…'; call('appUpdateNow').then(function (m) { toast(m); return load(); }).catch(fail); }\nfunction undoUpd(b) { b.disabled = true; call('appUndoUpdate').then(function (m) { toast(m); return load(); }).catch(fail); }\n\n/* ---------- Manage / link / join ---------- */\nfunction manage(i) { S.edit = S.d.accounts[i]; S.sub = 'manage'; render(); }\nfunction viewManage() {\n  var a = S.edit || {};\n  return topBar(a.label || 'Account', \"S.sub=null;render()\") + '<main class=\"page\"><div><label class=\"f\" for=\"lbl\">Name in the app and digest</label><input class=\"input\" id=\"lbl\" maxlength=\"24\" value=\"' + esc(a.label) + '\"></div>' +\n    '<div class=\"note\" style=\"text-align:left\">' + esc(a.email) + '</div><button class=\"btn primary\" style=\"flex:none\" onclick=\"saveLabel()\">Save name</button>' +\n    (a.own ? '' : '<button class=\"btn\" style=\"flex:none;color:var(--alert-ink)\" onclick=\"unlink()\">Unlink this account</button><div class=\"note\" style=\"text-align:left\">Unlinking only removes it from this app and your digest. It keeps sorting its own mail.</div>') + '</main>';\n}\nfunction saveLabel() { var v = document.getElementById('lbl').value.trim(); if (!v) return; call('appSetSetting', S.edit.key, 'ACCOUNT_LABEL', v).then(function () { toast('Saved.'); S.sub = null; return load(); }).catch(fail); }\nfunction unlink() { call('appUnlink', S.edit.key).then(function (m) { toast(m); S.sub = null; return load('all'); }).catch(fail); }\nfunction openLink() {\n  S.sub = 'link'; S.link = S.d.self.linkCode ? { code: S.d.self.linkCode.code, email: S.d.me } : null; render();\n  if (!S.link) call('appLinkCode').then(function (r) { S.link = r; if (S.sub === 'link') render(); }).catch(fail);\n}\nfunction viewLink() {\n  var L = S.link;\n  return topBar('Link an account', \"S.sub=null;S.tab='settings';render()\") + '<main class=\"page\">' +\n    '<p style=\"margin:0;font-size:14px;color:var(--muted)\">See every inbox here and get one digest. Each account keeps sorting its own mail.</p>' +\n    '<div class=\"codecard\"><div class=\"eyebrow\">Step 1 · on this account</div><div style=\"font-size:14px\">Your link code</div><div class=\"code\">' + (L ? esc(L.code) : '·······') + '</div>' +\n    '<div style=\"display:flex;gap:8px;align-items:center\"><button class=\"btn small\" style=\"flex:none;background:var(--mint);border-color:var(--mint);color:#0b2224\" onclick=\"copyCode()\">Copy code</button><span style=\"font-size:13px;color:var(--head-muted)\">Works for 1 hour</span></div></div>' +\n    '<div class=\"card\" style=\"padding:16px;display:flex;flex-direction:column;gap:10px\"><div class=\"eyebrow\" style=\"color:var(--muted);font-weight:600\">Step 2 · on the other account</div>' +\n    '<ol style=\"margin:0;padding-left:20px;font-size:14px;display:flex;flex-direction:column;gap:4px\"><li>Open its Inbox Assistant app.</li><li>Settings → <b>Join a main account</b>.</li><li>Enter <b>' + esc(S.d.me) + '</b> and this code.</li></ol>' +\n    '<div style=\"font-size:13px;color:var(--muted);background:var(--ground);border-radius:10px;padding:10px 12px\">It shares its sheet with this account, sends its alerts here, and turns off its own digest.</div></div>' +\n    '<button class=\"btn soft\" style=\"flex:none\" onclick=\"load()\">Check if it\\'s linked</button><div class=\"note\">Linking finishes within about 10 minutes of joining.</div></main>';\n}\nfunction copyCode() { var c = S.link && S.link.code; if (!c) return; (navigator.clipboard ? navigator.clipboard.writeText(c) : Promise.reject()).then(function () { toast('Copied ' + c); }, function () { toast('Code: ' + c); }); }\nfunction viewJoin() {\n  return topBar('Join a main account', \"S.sub=null;S.tab='settings';render()\") + '<main class=\"page\">' +\n    '<p style=\"margin:0;font-size:14px;color:var(--muted)\">Use this on an extra account to add it to your main account\\'s app and digest. Get the code from the main account: Settings → Link another account.</p>' +\n    '<div><label class=\"f\" for=\"me\">Main account\\'s email</label><input class=\"input\" id=\"me\" type=\"email\" autocomplete=\"email\" placeholder=\"you@gmail.com\"></div>' +\n    '<div><label class=\"f\" for=\"cd\">Link code</label><input class=\"input\" id=\"cd\" autocapitalize=\"characters\" placeholder=\"K7P-42Q\" style=\"font-family:var(--display);letter-spacing:.1em;font-size:20px\"></div>' +\n    '<button class=\"btn primary\" style=\"flex:none\" id=\"jb\" onclick=\"join()\">Link to main account</button><div class=\"note\">This account keeps sorting its own mail. Its digest and alerts go to the main account.</div></main>';\n}\nfunction join() {\n  var b = document.getElementById('jb'); b.disabled = true;\n  call('appJoinMain', document.getElementById('me').value, document.getElementById('cd').value).then(function (m) { toast(m); S.sub = null; S.tab = 'settings'; return load(); }).catch(function (e) { b.disabled = false; fail(e); });\n}\n\nload();\n</script>\n</body>\n</html>\n";

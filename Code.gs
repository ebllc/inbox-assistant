/**
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
const CODE_VERSION = '2026.11.11-4';
const REPO = 'ebllc/inbox-assistant'; // public GitHub repo every copy updates from
const RAW_URL = branch => 'https://raw.githubusercontent.com/' + REPO + '/' + branch + '/Code.gs';

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
  try { updateCode(); } catch (e) { /* API switch off — fine, they can update later */ }
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
  const bound = boundSheet_();
  if (bound) return bound;
  const props = PS_().getUserProperties();
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
const HOWTO_VERSION = '2026-11-11j';
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
    ['', 'Get updates automatically', 'One time: script.google.com/home/usersettings → Google Apps Script API → On. Then Settings → AUTO_UPDATE = TRUE. It checks once a day.'],
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
    } catch (e) { Logger.log('Pending ' + id + ': ' + e); }
    sh.deleteRow(i + 2);
  }
}

// ---------- sheet menu (template installs) ----------
/** Adds the 📬 Inbox Assistant menu when the sheet opens. Works only when the code lives in this sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📬 Inbox Assistant')
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

// ---------- self-update ----------
// Every copy pulls Code.gs from the public GitHub repo (branch = UPDATE_CHANNEL) and rewrites its own project
// through the Apps Script API. Needs, once per account: the Apps Script API switched on at
// script.google.com/home/usersettings, and the manifest scopes below (appsscript.json on the setup page).
const REQUIRED_SCOPES = [
  'https://mail.google.com/',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/script.external_request',
  'https://www.googleapis.com/auth/script.scriptapp',
  'https://www.googleapis.com/auth/script.projects',
  'https://www.googleapis.com/auth/script.container.ui',
  'https://www.googleapis.com/auth/userinfo.email',
];
const BACKUP_FILE = 'PreviousVersion';

function channel_() { return /^beta$/i.test(String(readSettings_().UPDATE_CHANNEL || '')) ? 'beta' : 'main'; }

/** Latest published code for this copy's channel. */
function fetchRelease_() {
  const res = UrlFetchApp.fetch(RAW_URL(channel_()), { muteHttpExceptions: true, headers: { 'Cache-Control': 'no-cache' } });
  if (res.getResponseCode() !== 200) throw new Error('Couldn\'t reach the update server (' + res.getResponseCode() + '). Try again later.');
  return [{ name: 'Code', type: 'SERVER_JS', source: res.getContentText() }];
}

function scriptApi_(method, scriptId, payload) {
  const res = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + scriptId + '/content', {
    method: method, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + SA_().getOAuthToken() },
    payload: payload ? JSON.stringify(payload) : undefined,
  });
  const code = res.getResponseCode(), text = res.getContentText();
  if (code === 200) return JSON.parse(text);
  if (/SERVICE_DISABLED|has not been used|not enabled|User has not enabled the Apps Script API/i.test(text))
    throw new Error('Turn on the Apps Script API: script.google.com/home/usersettings → Google Apps Script API → On. Then try again.');
  if (code === 403 && /scope/i.test(text))
    throw new Error('This copy is missing a permission. Replace appsscript.json with the one on the setup page, save, run updateCode once and click Allow.');
  throw new Error('Apps Script API ' + code + ': ' + text.slice(0, 200));
}

function versionOf_(files) {
  const src = files.filter(f => f.type === 'SERVER_JS').map(f => f.source).join('\n');
  const m = src.match(/const CODE_VERSION = '([^']+)'/);
  return m ? m[1] : '';
}

/** Install the latest release into this project. Keeps this project's own settings file and a backup. */
function updateCode(force) {
  const stamp = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a');
  if (viaLibrary_()) { writeSetting_('LAST_UPDATE', stamp() + ' · shared code — always current, nothing to do'); return; }
  try {
    const src = fetchRelease_();
    const newVersion = versionOf_(src);
    if (!newVersion || !src.some(f => f.type === 'SERVER_JS' && /function processInbox\(/.test(f.source)))
      throw new Error('The source code looks incomplete, so nothing was changed.');
    if (newVersion === CODE_VERSION && force !== true) {
      writeSetting_('LAST_UPDATE', stamp() + ' · already up to date (' + CODE_VERSION + ')');
      PS_().getUserProperties().deleteProperty('UPDATE_AVAILABLE');
      return;
    }
    const mine = scriptApi_('get', SA_().getScriptId()).files || [];
    const manifest = mine.find(f => f.type === 'JSON' && f.name === 'appsscript');
    const m = JSON.parse(manifest.source);
    m.oauthScopes = Array.from(new Set((m.oauthScopes || []).concat(REQUIRED_SCOPES)));
    const backup = { name: BACKUP_FILE, type: 'HTML',
      source: JSON.stringify(mine.filter(f => f.type === 'SERVER_JS').map(f => ({ name: f.name, source: f.source }))) };
    const incoming = src.filter(f => f.type === 'SERVER_JS').map(f => ({ name: f.name, type: f.type, source: f.source }));
    scriptApi_('put', SA_().getScriptId(), { files: [{ name: 'appsscript', type: 'JSON', source: JSON.stringify(m, null, 2) }].concat(incoming, [backup]) });
    PS_().getUserProperties().deleteProperty('UPDATE_AVAILABLE');
    writeSetting_('LAST_UPDATE', stamp() + ' · updated ' + CODE_VERSION + ' → ' + newVersion + ' (' + channel_() + '). If anything breaks, run restorePreviousCode.');
  } catch (e) {
    writeSetting_('LAST_UPDATE', stamp() + ' · ⚠ ' + e.message);
    Logger.log(e);
  }
}

/** Put back the version from before the last update. */
function restorePreviousCode() {
  const mine = scriptApi_('get', SA_().getScriptId()).files || [];
  const backup = mine.find(f => f.name === BACKUP_FILE);
  if (!backup) { writeSetting_('LAST_UPDATE', 'No previous version saved — nothing to restore.'); return; }
  const manifest = mine.find(f => f.type === 'JSON' && f.name === 'appsscript');
  const old = JSON.parse(backup.source).map(f => ({ name: f.name, type: 'SERVER_JS', source: f.source }));
  scriptApi_('put', SA_().getScriptId(), { files: [manifest].concat(old) });
  writeSetting_('LAST_UPDATE', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d, h:mm a') + ' · restored the previous version');
}

/** Once a day: install updates if AUTO_UPDATE is on, otherwise just note that one is available. */
function dailyUpdateCheck_() {
  if (viaLibrary_()) return;
  const props = PS_().getUserProperties();
  const last = Number(props.getProperty('LAST_UPDATE_CHECK') || 0);
  if (Date.now() - last < 23 * 3600 * 1000) return;
  props.setProperty('LAST_UPDATE_CHECK', String(Date.now()));
  if (readSettings_().AUTO_UPDATE) { updateCode(); return; }
  try {
    const v = versionOf_(fetchRelease_());
    if (v && v !== CODE_VERSION) props.setProperty('UPDATE_AVAILABLE', v);
    else props.deleteProperty('UPDATE_AVAILABLE');
  } catch (e) { /* offline — try again tomorrow */ }
}

function updateBanner_() {
  const v = PS_().getUserProperties().getProperty('UPDATE_AVAILABLE');
  if (!v || v === CODE_VERSION) return '';
  return '<div style="background:#eef4ff;border:1px solid #c7d7fe;border-radius:8px;padding:10px 14px;margin:0 0 14px">' +
    '🆕 <b>A new version is available.</b> In your sheet\'s Settings tab, set <b>UPDATE_CODE</b> to TRUE ' +
    '(or <b>AUTO_UPDATE</b> to TRUE to get them automatically).</div>';
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
  return /^(☕ Morning digest|🥪 Midday digest|🌙 Evening digest|🔔 |Inbox Assistant is set up)/.test(String(subject || ''));
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

/** Library entry point used by the stub in other accounts. */
function run(name, host) {
  HOST_ = host || null;
  if (!name) name = 'runNow'; // picked from the Run menu by hand → just check the inbox
  const fns = { setup, runNow, processInbox, processPending, sendDigest, sendDigestNow, reinstallTriggers, updateCode, restorePreviousCode, repairTabs,
    forgetProcessedLabel, refreshHowTo, restyleTabs };
  if (!fns[name]) throw new Error('Unknown function: ' + name);
  return fns[name]();
}

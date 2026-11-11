/**
 * Inbox Assistant — LOADER
 *
 * This is the only code that goes in your Apps Script project. It downloads the real assistant
 * (Core.gs) from github.com/ebllc/inbox-assistant, keeps a copy in this project, and runs it.
 * Updates: the app's Settings → Check for update (or AUTO_UPDATE = TRUE in the sheet). Undo keeps the previous version.
 */
const IA_REPO = 'ebllc/inbox-assistant';
const IA_LOADER_VERSION = '1';
const IA_CHUNK = 2000; // characters per stored piece (Apps Script caps each stored value at 9 KB)

// ---------- the 📬 menu (built without loading anything, so it always appears) ----------
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

// ---------- every entry point hands off to the downloaded code ----------
function setup() { return ia_('setup', arguments); }
function runNow() { return ia_('runNow', arguments); }
function processInbox() { return ia_('processInbox', arguments); }
function processPending() { return ia_('processPending', arguments); }
function sendDigest() { return ia_('sendDigest', arguments); }
function sendDigestNow() { return ia_('sendDigestNow', arguments); }
function reinstallTriggers() { return ia_('reinstallTriggers', arguments); }
function forgetProcessedLabel() { return ia_('forgetProcessedLabel', arguments); }
function refreshHowTo() { return ia_('refreshHowTo', arguments); }
function restyleTabs() { return ia_('restyleTabs', arguments); }
function repairTabs() { return ia_('repairTabs', arguments); }
function recheckFromMenu() { return ia_('recheckFromMenu', arguments); }
function updateCode() { return ia_('updateCode', arguments); }
function restorePreviousCode() { return ia_('restorePreviousCode', arguments); }
function updateNowFromMenu() { return ia_('updateNowFromMenu', arguments); }
function openAppFromMenu() { return ia_('openAppFromMenu', arguments); }
function publishApp() { return ia_('publishApp', arguments); }
function onSheetEdit() { return ia_('onSheetEdit', arguments); }
function doGet() { return ia_('doGet', arguments); }
// the app (google.script.run)
function appState() { return ia_('appState', arguments); }
function appSetSetting() { return ia_('appSetSetting', arguments); }
function appRuleSet() { return ia_('appRuleSet', arguments); }
function appRuleAdd() { return ia_('appRuleAdd', arguments); }
function appRuleRemove() { return ia_('appRuleRemove', arguments); }
function appEmail() { return ia_('appEmail', arguments); }
function appEmailAction() { return ia_('appEmailAction', arguments); }
function appCheckNow() { return ia_('appCheckNow', arguments); }
function appSendDigestNow() { return ia_('appSendDigestNow', arguments); }
function appUpdateNow() { return ia_('appUpdateNow', arguments); }
function appUndoUpdate() { return ia_('appUndoUpdate', arguments); }
function appLinkCode() { return ia_('appLinkCode', arguments); }
function appJoinMain() { return ia_('appJoinMain', arguments); }
function appUnlink() { return ia_('appUnlink', arguments); }

// ---------- loading, storing, updating ----------
let IA_CORE_ = null;

function ia_(name, args) {
  return ia_core_().run(name, IA_HOST_, Array.prototype.slice.call(args || []));
}

const IA_HOST_ = {
  loader: {
    version: IA_LOADER_VERSION,
    info: function () { return { current: ia_version_(ia_read_('CUR')), previous: ia_version_(ia_read_('PREV')) }; },
    remoteVersion: function (channel) { return ia_version_(ia_fetch_(channel)); },
    install: function (channel) {
      const cur = ia_read_('CUR'), src = ia_fetch_(channel);
      const from = ia_version_(cur), to = ia_version_(src);
      if (cur && from === to) return { changed: false, from: from, to: to };
      if (cur) ia_write_('PREV', cur);
      ia_write_('CUR', src);
      return { changed: true, from: from || 'none', to: to };
    },
    undo: function () {
      const cur = ia_read_('CUR'), prev = ia_read_('PREV');
      if (!prev) throw new Error('No earlier version is saved, so there is nothing to undo.');
      ia_write_('CUR', prev);
      ia_write_('PREV', cur);
      return { from: ia_version_(cur), to: ia_version_(prev) };
    },
  },
};

function ia_core_() {
  if (IA_CORE_) return IA_CORE_;
  let src = ia_read_('CUR');
  if (!src) { src = ia_fetch_('main'); ia_write_('CUR', src); }
  try {
    IA_CORE_ = ia_compile_(src)();
  } catch (e) {
    // The stored version won't start: fall back to the previous one, then to a fresh download.
    const prev = ia_read_('PREV');
    const fallback = prev && prev !== src ? prev : ia_fetch_('main');
    ia_write_('CUR', fallback);
    IA_CORE_ = ia_compile_(fallback)();
  }
  return IA_CORE_;
}

function ia_compile_(src) {
  return new Function(src + '\n;return { run: run, CODE_VERSION: CODE_VERSION };');
}

function ia_fetch_(channel) {
  const url = 'https://raw.githubusercontent.com/' + IA_REPO + '/' + (channel === 'beta' ? 'beta' : 'main') + '/Core.gs';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'Cache-Control': 'no-cache' } });
  if (res.getResponseCode() !== 200) throw new Error('Couldn\'t download Inbox Assistant (' + res.getResponseCode() + '). Try again in a few minutes.');
  const src = res.getContentText();
  if (!/const CODE_VERSION = '/.test(src) || src.indexOf('function processInbox(') < 0) throw new Error('The download looked incomplete, so nothing was changed.');
  ia_compile_(src); // throws on a syntax error, before anything is stored
  return src;
}

function ia_version_(src) {
  const m = String(src || '').match(/const CODE_VERSION = '([^']+)'/);
  return m ? m[1] : '';
}

function ia_read_(slot) {
  const p = PropertiesService.getScriptProperties().getProperties();
  const n = Number(p['IA_' + slot + '_N'] || 0);
  let s = '';
  for (let i = 0; i < n; i++) s += p['IA_' + slot + '_' + i] || '';
  return s;
}

function ia_write_(slot, src) {
  const store = PropertiesService.getScriptProperties();
  const old = Number(store.getProperty('IA_' + slot + '_N') || 0);
  const out = {};
  let n = 0;
  for (let i = 0; i < src.length; i += IA_CHUNK) out['IA_' + slot + '_' + (n++)] = src.slice(i, i + IA_CHUNK);
  out['IA_' + slot + '_N'] = String(n);
  store.setProperties(out, false);
  for (let i = n; i < old; i++) store.deleteProperty('IA_' + slot + '_' + i);
}

/** Never called. Lists the Google services the assistant uses, so Google asks for the right permissions. */
function ia_permissions_() {
  GmailApp.search; GmailApp.sendEmail; SpreadsheetApp.openById; SpreadsheetApp.create; SpreadsheetApp.getUi;
  HtmlService.createHtmlOutput; UrlFetchApp.fetch; ScriptApp.newTrigger; Session.getActiveUser; CacheService.getScriptCache;
}

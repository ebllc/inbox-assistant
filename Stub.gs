/**
 * Inbox Assistant — SHARED-CODE STUB
 * The real code lives in your main account's project (added here as the library "IA").
 * Never paste updates here: when the main project's code changes, this account picks it up automatically.
 */
const HOST = { ScriptApp: ScriptApp, PropertiesService: PropertiesService };

function runNow()               { IA.run('runNow', HOST); }
function sendDigestNow()        { IA.run('sendDigestNow', HOST); }
function reinstallTriggers()    { IA.run('reinstallTriggers', HOST); }
function forgetProcessedLabel() { IA.run('forgetProcessedLabel', HOST); }
function refreshHowTo()         { IA.run('refreshHowTo', HOST); }
function restyleTabs()          { IA.run('restyleTabs', HOST); }
function setup()                { IA.run('setup', HOST); }

// Run automatically on a schedule — no need to pick these.
function processInbox()   { IA.run('processInbox', HOST); }
function processPending() { IA.run('processPending', HOST); }
function sendDigest()     { IA.run('sendDigest', HOST); }

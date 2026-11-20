/* ==========================================================================
   QuickLinksSync.gs  |  v1.1  |  2026-11-20 2:25 PM ET (GitHub key from Firebase user_secrets)
   Apps Script for the "Quick Links 4 Scott" Google Sheet.

   WHAT IT DOES
   - Looks at the SMLC GitHub repository (Claude-Test branch) for every web page (.html).
   - Each page gets a row in this sheet: its werewolf3788.github.io link (column "Url"),
     its GitHub file link, what the last change was, and when it was last changed.
   - Rows you already made by hand (Home, Xenia, Flora, Photo Uploader...) are matched by
     their Url. Only the GitHub / What Changed / Last Updated cells are filled in; your Name,
     Description and Group are never changed.
   - New pages get a new row at the bottom. You can rename them or change the Group freely.
   - The Quick Links website sorts by "Last Updated", so the newest changes show first.

   SETUP (once)
   1. In the sheet: Extensions -> Apps Script. Paste this whole file in place of what's there.
   2. Save, then reload the sheet. A "Quick Links" menu appears.
   3. Quick Links -> "Sync from GitHub now". Google asks for permission the first time: allow it.
   4. Quick Links -> "Turn on hourly sync".

   GITHUB KEY (needed: Apps Script shares Google's servers with everyone, so GitHub's
   free no-key limit is almost always used up already)
   a. GitHub -> Settings -> Developer settings -> Personal access tokens -> Fine-grained tokens
      -> Generate new token. Name: Quick Links. Repository access: "Public repositories".
      No permissions needed (read-only). Copy the token (starts with github_pat_).
   b. Save it in Firebase (entertainment-71888) Realtime Database at
         user_secrets/raykevin71888@gmail,com/github_token
      The script reads it from there with your Google sign-in (you own that Firebase project).
   c. In Apps Script: Project Settings (gear) -> tick "Show appsscript.json manifest file".
      Open appsscript.json and make "oauthScopes" match the list at the bottom of this file.

   Pages to leave out (old/test copies) are listed in EXCLUDE below.
   ========================================================================== */

var QL_SOURCES = [
  {
    repo: 'Werewolf3788/SMLC',
    branch: 'Claude-Test',
    site: 'https://werewolf3788.github.io/SMLC/',
    group: 'SMLC Github Webpage',          // Group for NEW rows (you can change it afterwards)
    exclude: [/^full\//, /^test\//, /^wordpress\//, /^node_modules\//, /^local-news\//]
  }
  // Another repo can be added the same way, e.g.
  // { repo: 'Werewolf3788/Quick-Links', branch: 'main', site: 'https://werewolf3788.github.io/Quick-Links/',
  //   group: 'Quick Links', exclude: [] }
];

// Where the GitHub key lives (Kevin's Firebase, entertainment-71888)
var QL_SECRETS_DB = 'https://entertainment-71888-default-rtdb.firebaseio.com';
var QL_SECRET_PATHS = ['user_secrets/raykevin71888@gmail,com/github_token',
                       'user_secrets/raykevin71888@gmail_com/github_token',
                       'user_secrets/raykevin71888_gmail_com/github_token'];

var QL_COLS = { name: 'Name', desc: 'Description', group: 'Group', url: 'Url',
                github: 'GitHub', changed: 'What Changed', updated: 'Last Updated' };
var QL_DATE_FORMAT = 'yyyy-mm-dd hh:mm';    // the website reads this format

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quick Links')
    .addItem('Sync from GitHub now', 'syncQuickLinks')
    .addSeparator()
    .addItem('Turn on hourly sync', 'turnOnHourlySync')
    .addItem('Turn off hourly sync', 'turnOffHourlySync')
    .addItem('Start over (re-check every page)', 'resetQuickLinksSync')
    .addToUi();
}

function turnOnHourlySync() {
  turnOffHourlySync();
  ScriptApp.newTrigger('syncQuickLinks').timeBased().everyHours(1).create();
  toast_('Hourly sync is on.');
}
function turnOffHourlySync() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncQuickLinks') ScriptApp.deleteTrigger(t);
  });
}
function resetQuickLinksSync() {
  var props = PropertiesService.getScriptProperties();
  props.getKeys().forEach(function (k) { if (k.indexOf('QL_') === 0) props.deleteProperty(k); });
  syncQuickLinks();
}

/* ---------------------------------------------------------------------- */

function syncQuickLinks() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var col = ensureColumns_(sheet);
    var added = 0, updated = 0;

    QL_SOURCES.forEach(function (src) {
      var pages = latestChanges_(src);                    // { path: { date, message, sha } }
      var data = sheet.getDataRange().getValues();

      // Existing rows by normalized Url
      var rowByKey = {};
      for (var r = 1; r < data.length; r++) {
        var k = normUrl_(data[r][col.url - 1]);
        if (k) rowByKey[k] = r + 1;
      }

      Object.keys(pages).sort().forEach(function (path) {
        var info = pages[path];
        var pageUrl = pageUrl_(src, path);
        var ghUrl = 'https://github.com/' + src.repo + '/blob/' + src.branch + '/' + path;
        var row = rowByKey[normUrl_(pageUrl)];

        if (!row) {
          row = firstEmptyRow_(sheet, col.name);
          sheet.getRange(row, col.name).setValue(niceName_(path));
          sheet.getRange(row, col.group).setValue(src.group);
          sheet.getRange(row, col.url).setValue(pageUrl);
          rowByKey[normUrl_(pageUrl)] = row;
          added++;
        }
        var cur = sheet.getRange(row, col.updated).getValue();
        var curTime = cur instanceof Date ? cur.getTime() : 0;
        if (!curTime || Math.abs(curTime - info.date.getTime()) > 60000) {
          sheet.getRange(row, col.github).setValue(ghUrl);
          sheet.getRange(row, col.changed).setValue(info.message);
          sheet.getRange(row, col.updated).setValue(info.date).setNumberFormat(QL_DATE_FORMAT);
          updated++;
        }
      });
    });

    toast_('Synced: ' + added + ' new, ' + updated + ' updated.');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- GitHub ---------- */

// Returns the newest change for every .html page. Remembers what it already knows in Script
// Properties, so after the first run each sync only asks GitHub about new commits.
function latestChanges_(src) {
  var props = PropertiesService.getScriptProperties();
  var key = 'QL_' + src.repo.replace(/\W/g, '_') + '_' + src.branch.replace(/\W/g, '_');
  var saved = JSON.parse(props.getProperty(key) || '{"head":"","pages":{}}');

  var branch = gh_('/repos/' + src.repo + '/branches/' + encodeURIComponent(src.branch));
  var head = branch.commit.sha;

  // Current list of pages
  var tree = gh_('/repos/' + src.repo + '/git/trees/' + head + '?recursive=1');
  var paths = tree.tree.filter(function (t) {
    return t.type === 'blob' && /\.html?$/i.test(t.path) &&
           !src.exclude.some(function (re) { return re.test(t.path); });
  }).map(function (t) { return t.path; });

  var pages = saved.pages || {};
  var known = {};
  paths.forEach(function (p) { known[p] = true; });
  Object.keys(pages).forEach(function (p) { if (!known[p]) delete pages[p]; });   // page was deleted

  if (saved.head && saved.head !== head) {
    // New commits since last time, oldest to newest, so the newest change wins
    try {
      var cmp = gh_('/repos/' + src.repo + '/compare/' + saved.head + '...' + head);
      (cmp.commits || []).forEach(function (c) {
        var detail = gh_('/repos/' + src.repo + '/commits/' + c.sha);
        (detail.files || []).forEach(function (f) {
          if (known[f.filename]) pages[f.filename] = pack_(c);
        });
      });
    } catch (e) {
      // History was rewritten (or too many commits): re-check every page below
      Logger.log('Compare failed, re-checking all pages: ' + e.message);
      pages = {};
    }
  }

  // Pages we have no date for yet (first run, or brand-new pages): ask per page
  paths.forEach(function (p) {
    if (pages[p]) return;
    var list = gh_('/repos/' + src.repo + '/commits?sha=' + encodeURIComponent(src.branch) +
                   '&path=' + encodeURIComponent(p) + '&per_page=1');
    if (list && list.length) pages[p] = pack_(list[0]);
  });

  props.setProperty(key, JSON.stringify({ head: head, pages: pages }));

  var out = {};
  Object.keys(pages).forEach(function (p) {
    out[p] = { date: new Date(pages[p].d), message: pages[p].m };
  });
  return out;
}

function pack_(c) {
  var msg = String(c.commit.message || '').split('\n')[0].slice(0, 140);
  return { d: c.commit.committer.date, m: msg };
}

// Reads the GitHub key from Firebase user_secrets (cached for 6 hours).
// Owner access through your Google sign-in, so the database rules don't need to allow it.
var QL_TOKEN_ = null;
function githubToken_() {
  if (QL_TOKEN_ !== null) return QL_TOKEN_;
  var cache = CacheService.getScriptCache();
  var hit = cache.get('QL_GH_TOKEN');
  if (hit) return (QL_TOKEN_ = hit);
  var oauth = ScriptApp.getOAuthToken();
  for (var i = 0; i < QL_SECRET_PATHS.length; i++) {
    var res = UrlFetchApp.fetch(QL_SECRETS_DB + '/' + QL_SECRET_PATHS[i].split('/').map(encodeURIComponent).join('/') + '.json',
                                { headers: { Authorization: 'Bearer ' + oauth }, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      var v = JSON.parse(res.getContentText());
      if (typeof v === 'string' && v.trim()) {
        QL_TOKEN_ = v.trim();
        cache.put('QL_GH_TOKEN', QL_TOKEN_, 21600);
        return QL_TOKEN_;
      }
    } else if (i === 0) {
      Logger.log('Firebase key read: HTTP ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 150));
    }
  }
  return (QL_TOKEN_ = '');
}

function gh_(path) {
  var headers = { Accept: 'application/vnd.github+json', 'User-Agent': 'QuickLinksSync' };
  var token = githubToken_();
  if (token) headers.Authorization = 'Bearer ' + token;
  var res = UrlFetchApp.fetch('https://api.github.com' + path, { headers: headers, muteHttpExceptions: true });
  var code = res.getResponseCode();
  if (code === 401) throw new Error('GitHub rejected the key in Firebase user_secrets (expired or mistyped). Make a new one.');
  if (code === 403 || code === 429) {
    var hd = res.getHeaders();
    var reset = Number(hd['x-ratelimit-reset'] || hd['X-RateLimit-Reset'] || 0);
    var when = reset ? ' Resets at ' + Utilities.formatDate(new Date(reset * 1000), Session.getScriptTimeZone(), 'h:mm a') + '.' : '';
    throw new Error(token
      ? 'GitHub hourly limit reached even with the key.' + when
      : 'No GitHub key found in Firebase (user_secrets/raykevin71888@gmail,com/github_token), and GitHub\'s no-key limit is used up.' + when +
        ' See GITHUB KEY at the top of this script.');
  }
  if (code >= 300) throw new Error('GitHub ' + code + ' for ' + path + ': ' + res.getContentText().slice(0, 200));
  return JSON.parse(res.getContentText());
}

/* ---------- Sheet helpers ---------- */

// Finds the columns by their header names; adds GitHub / What Changed / Last Updated if missing.
function ensureColumns_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var col = {};
  Object.keys(QL_COLS).forEach(function (k) {
    var i = headers.indexOf(QL_COLS[k]);
    if (i === -1) {
      // always add new columns after the last used one (column E has no header but holds poster links)
      var pos = sheet.getLastColumn() + 1;
      sheet.getRange(1, pos).setValue(QL_COLS[k]).setFontWeight('bold');
      headers[pos - 1] = QL_COLS[k];
      i = pos - 1;
    }
    col[k] = i + 1;
  });
  return col;
}

function firstEmptyRow_(sheet, nameCol) {
  var vals = sheet.getRange(1, nameCol, Math.max(sheet.getLastRow(), 1), 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) { if (String(vals[i][0]).trim()) return i + 2; }
  return 2;
}

function pageUrl_(src, path) {
  var p = path.replace(/(^|\/)index\.html?$/i, '$1');
  return src.site + p.replace(/\.html?$/i, '');
}

// "https://werewolf3788.github.io/SMLC/flora.html?x#y" and ".../SMLC/flora" -> same key
function normUrl_(u) {
  u = String(u || '').trim().toLowerCase();
  if (!u) return '';
  return u.replace(/^https?:\/\//, '').replace(/^www\./, '').replace(/[?#].*$/, '')
          .replace(/\/index\.html?$/, '/').replace(/\.html?$/, '').replace(/\/+$/, '');
}

function niceName_(path) {
  var parts = path.replace(/\.html?$/i, '').split('/');
  var file = parts.pop();
  var name = file.replace(/[-_]+/g, ' ').replace(/\b\w/g, function (c) { return c.toUpperCase(); });
  if (/^index$/i.test(file)) name = parts.length ? parts[parts.length - 1] + ' Home' : 'Home';
  return parts.length ? name + ' (' + parts.join('/') + ')' : name;
}

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Quick Links', 5); } catch (e) { Logger.log(msg); }
}

/* ---------- appsscript.json ----------
{
  "timeZone": "America/New_York",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/firebase.database"
  ]
}
   (Keep your own "timeZone" if it's different.)
-------------------------------------- */

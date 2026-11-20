/* ==========================================================================
   QuickLinksSync.gs  |  v2.0  |  2026-11-20 2:40 PM ET  (no key needed)
   Apps Script for the "Quick Links 4 Scott" Google Sheet.

   WHAT IT DOES
   - Every push to the SMLC repo (Claude-Test) makes GitHub write a list of all the web
     pages and when each last changed: data/site-pages.json. Nothing to set up for that.
   - This script reads that list and gives each page a row in this sheet: its
     werewolf3788.github.io link ("Url"), GitHub file link, what changed, and when.
   - Rows you made by hand (Home, Xenia, Flora, Photo Uploader...) are matched by their Url.
     Only the GitHub / What Changed / Last Updated cells are filled in; your Name,
     Description and Group are never changed.
   - New pages get a new row at the bottom. Rename them or change the Group freely.
   - The Quick Links website sorts by "Last Updated", so the newest changes show first.

   SETUP (once)
   1. Extensions -> Apps Script. Paste this whole file in place of what's there. Save.
   2. Reload the sheet. A "Quick Links" menu appears.
   3. Quick Links -> "Sync now" (allow the permission prompt the first time).
   4. Quick Links -> "Turn on hourly sync".
   ========================================================================== */

var QL_SOURCES = [
  {
    list: 'https://raw.githubusercontent.com/Werewolf3788/SMLC/Claude-Test/data/site-pages.json',
    group: 'SMLC Github Webpage'          // Group for NEW rows (you can change it afterwards)
  }
];

var QL_COLS = { name: 'Name', desc: 'Description', group: 'Group', url: 'Url',
                github: 'GitHub', changed: 'What Changed', updated: 'Last Updated' };
var QL_DATE_FORMAT = 'yyyy-mm-dd hh:mm';    // the website reads this format

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Quick Links')
    .addItem('Sync now', 'syncQuickLinks')
    .addSeparator()
    .addItem('Turn on hourly sync', 'turnOnHourlySync')
    .addItem('Turn off hourly sync', 'turnOffHourlySync')
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

function syncQuickLinks() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var col = ensureColumns_(sheet);
    var added = 0, updated = 0;

    QL_SOURCES.forEach(function (src) {
      var list = fetchList_(src.list);
      var data = sheet.getDataRange().getValues();

      // Existing rows by their Url
      var rowByKey = {};
      for (var r = 1; r < data.length; r++) {
        var k = normUrl_(data[r][col.url - 1]);
        if (k) rowByKey[k] = r + 1;
      }

      (list.pages || []).forEach(function (pg) {
        var when = new Date(pg.updated);
        var row = rowByKey[normUrl_(pg.url)];

        if (!row) {
          row = firstEmptyRow_(sheet, col.name);
          sheet.getRange(row, col.name).setValue(niceName_(pg.path));
          sheet.getRange(row, col.group).setValue(src.group);
          sheet.getRange(row, col.url).setValue(pg.url);
          rowByKey[normUrl_(pg.url)] = row;
          added++;
        }
        var cur = sheet.getRange(row, col.updated).getValue();
        var curTime = cur instanceof Date ? cur.getTime() : 0;
        if (!curTime || Math.abs(curTime - when.getTime()) > 60000) {
          sheet.getRange(row, col.github).setValue(pg.github);
          sheet.getRange(row, col.changed).setValue(pg.message || '');
          sheet.getRange(row, col.updated).setValue(when).setNumberFormat(QL_DATE_FORMAT);
          updated++;
        }
      });
    });

    toast_('Synced: ' + added + ' new, ' + updated + ' updated.');
  } finally {
    lock.releaseLock();
  }
}

function fetchList_(url) {
  var res = UrlFetchApp.fetch(url + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('Could not read the page list (' + res.getResponseCode() + '): ' + url);
  }
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

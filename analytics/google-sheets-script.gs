/**
 * Mr Wakefield's Fount — anonymous usage counts.
 *
 * Paste this whole file into your Google Sheet: Extensions → Apps Script.
 * Then: Deploy → New deployment → type "Web app" → Execute as "Me" →
 * Who has access "Anyone" → Deploy. Copy the web-app URL (ends in /exec).
 *
 * The revision site only ever sends totals: games started, right/wrong per
 * mode and per card, the date and the year group. No names, no device IDs.
 */

const TABS = {
  "Daily":         ["Date", "Year", "Questions answered", "Right", "% right", "Games started"],
  "Games":         ["Game", "Times started", "Last played"],
  "Modes":         ["Mode", "Right", "Wrong", "Answers", "% right"],
  "Cards":         ["Card #", "Topic", "Question", "Right", "Wrong", "Answers", "% right"],
};

const TEXT_COLS = ["Topic", "Question", "Last played"];

/** Run once from the editor (or it runs itself on the first data). Makes the tabs. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.entries(TABS).forEach(([name, head]) => {
    let sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(head);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, head.length).setFontWeight("bold");
    }
  });
  let hard = ss.getSheetByName("Hardest cards");
  if (!hard) {
    hard = ss.insertSheet("Hardest cards");
    hard.getRange("A1").setValue("Cards answered at least 10 times, lowest % right first");
    hard.getRange("A1").setFontWeight("bold");
    hard.getRange("A3").setFormula('=IFERROR(QUERY(Cards!A:G,"select A,B,C,F,G where F >= 10 order by G asc limit 50",1),"Not enough answers yet")');
  }
  const first = ss.getSheetByName("Sheet1");
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) ss.deleteSheet(first);
}

function doGet() {
  return ContentService.createTextOutput("The Fount analytics endpoint is working.");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!p || p.v !== 1) return done_();
    setup();
    const date = safeDate_(p.d), year = ["y10", "y11"].includes(p.y) ? p.y : "";
    const num = v => Math.max(0, Math.min(500, Math.floor(Number(v) || 0)));   // ignore silly numbers

    // per mode
    let right = 0, answered = 0;
    const modes = Object.entries(p.modes || {}).slice(0, 40);
    upsert_("Modes", 1, modes.map(([m, rw]) => {
      const r = num(rw && rw[0]), w = num(rw && rw[1]); right += r; answered += r + w;
      return { key: [clean_(m, 40)], add: [r, w] };
    }), (row, a) => { row[1] += a[0]; row[2] += a[1]; row[3] = row[1] + row[2]; row[4] = pct_(row[1], row[3]); });

    // games started
    let games = 0;
    const plays = Object.entries(p.plays || {}).slice(0, 40);
    upsert_("Games", 1, plays.map(([g, n]) => { const k = num(n); games += k; return { key: [clean_(g, 40)], add: [k] }; }),
      (row, a) => { row[1] += a[0]; row[2] = date; });

    // per day and year
    upsert_("Daily", 2, [{ key: [date, year], add: [answered, right, games] }],
      (row, a) => { row[2] += a[0]; row[3] += a[1]; row[4] = pct_(row[3], row[2]); row[5] += a[2]; });

    // per card
    const cards = (Array.isArray(p.cards) ? p.cards : []).slice(0, 300);
    upsert_("Cards", 1, cards.map(c => ({ key: [Math.floor(Number(c[0]) || 0)], add: [num(c[1]), num(c[2])], topic: clean_(c[3], 40), q: clean_(c[4], 120) }))
      .filter(c => c.key[0] > 0),
      (row, a, c) => { row[1] = c.topic || row[1]; row[2] = c.q || row[2]; row[3] += a[0]; row[4] += a[1]; row[5] = row[3] + row[4]; row[6] = pct_(row[3], row[5]); });
  } catch (err) {
    // bad data is ignored
  } finally {
    lock.releaseLock();
  }
  return done_();
}

/** Add to existing rows (matched on the first `keyCols` columns) or append new ones, in one write. */
function upsert_(tab, keyCols, items, apply) {
  if (!items.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(tab), width = TABS[tab].length;
  const last = sh.getLastRow(), data = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  const index = {};
  data.forEach((row, i) => { index[keyOf_(row.slice(0, keyCols))] = i; });
  items.forEach(it => {
    let i = index[keyOf_(it.key)];
    if (i === undefined) {
      const row = new Array(width).fill(0);
      it.key.forEach((v, j) => row[j] = v);
      for (let j = keyCols; j < width; j++) if (TEXT_COLS.includes(TABS[tab][j])) row[j] = "";
      data.push(row); i = index[keyOf_(it.key)] = data.length - 1;
    }
    apply(data[i], it.add, it);
  });
  if (data.length) sh.getRange(2, 1, data.length, width).setValues(data);
}

function keyOf_(vals) {
  return vals.map(v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(v)).join("|");
}
function safeDate_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d)) ? String(d) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"); }
function pct_(r, n) { return n ? Math.round(r / n * 100) : 0; }
/** Plain text only: stops anything that looks like a formula being stored. */
function clean_(s, max) { s = String(s == null ? "" : s).replace(/[\r\n\t]+/g, " ").slice(0, max); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function done_() { return ContentService.createTextOutput("ok"); }

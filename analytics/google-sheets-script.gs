/**
 * Mr Wakefield's Fount — anonymous usage counts.
 *
 * Paste this whole file into your Google Sheet: Extensions → Apps Script.
 * Then: Deploy → New deployment → ⚙️ "Select type" → Web app → Execute as "Me" →
 * Who has access "Anyone" → Deploy. Copy the web-app URL (ends in /exec).
 *
 * Updating the script later: paste the new version, save, then
 * Deploy → Manage deployments → ✏️ Edit → Version: "New version" → Deploy.
 * That keeps the same web address. Then run "setup" once from the editor.
 *
 * The revision site only ever sends totals: games started, right/wrong per
 * mode and per card, the date, the year group, and two one-off flags
 * ("first time this device has sent anything", "first time today").
 * It also sends a list of actions (game finished, tab opened, level up...)
 * for the Events tab, tagged with two random codes: a visit code (new each
 * time the site is opened) and a device code (made up once per browser, so
 * returning visitors can be counted). Neither says who anyone is.
 * No names, emails or personal details.
 */

const TABS = {
  "Daily": ["Date", "Year", "Questions answered", "Right", "% right", "Games started", "Active students", "New students"],
  "Games": ["Game", "Times started", "Last played"],
  "Modes": ["Mode", "Right", "Wrong", "Answers", "% right"],
  "Cards": ["Card #", "Topic", "Question", "Right", "Wrong", "Answers", "% right", "AO"],
  "Events": ["Time", "Visit", "Device", "Event", "Detail", "Year", "Seconds"],
};
const TEXT_COLS = ["Topic", "Question", "Last played", "AO"];
const EVENT_ROWS_READ = 20000;   // the Summary looks at this many of the newest Events rows
const MIN_ANSWERS = 10;
const SITE_CODE = "fount-gcse";
/** Friendly names for the games in the Summary (the Games tab keeps the short codes). */
const GAME_NAMES = {"arcade:daily":"Daily 10","arcade:blitz":"Blitz","arcade:surv":"Survival","arcade:exam":"Exam Skills","arcade:maths":"RM Maths","arcade:terms":"Key Terms",
  "arcade:fix":"Fix my mistakes","arcade:boss":"Boss Battles","flaw":"Spot the Flaw","order":"Put in Order","swipe":"Swipe Sort","which-study":"Which Study?","truths":"Two Truths and a Lie",
  "odd":"Odd One Out","examiner":"Be the Examiner","tdee":"TDEE Builder","synoptic":"Synoptic Links","psychle":"Psychle","9-mark":"9-mark answers"};
const gameName_ = k => GAME_NAMES[k] || (String(k).startsWith("match:") ? "Study Matcher" : String(k).startsWith("lab:") ? "Psych Lab" : String(k));   // the revision site sends this with every batch; anything without it is ignored   // a card needs this many answers before it counts as hardest/easiest

/** Makes the tabs (and updates their headings). Runs itself whenever data arrives. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName("Summary")) ss.insertSheet("Summary", 0);
  Object.entries(TABS).forEach(([name, head]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
    sh.setFrozenRows(1);
  });
  renameGames_();
  const old = ss.getSheetByName("Hardest cards");   // replaced by the Summary tab
  if (old) ss.deleteSheet(old);
  const first = ss.getSheetByName("Sheet1");
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
}

function doGet() {
  return ContentService.createTextOutput("The Fount analytics endpoint is working.");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!p || p.v !== 1 || p.k !== SITE_CODE) return done_();   // only the revision site sends this code
    setup();
    const date = safeDate_(p.d), year = ["y10", "y11"].includes(p.y) ? p.y : "";
    const num = v => Math.max(0, Math.min(500, Math.floor(Number(v) || 0)));   // ignore silly numbers
    const u = p.u || {}, active = u.d === 1 ? 1 : 0, fresh = u.n === 1 ? 1 : 0;

    let right = 0, answered = 0;
    upsert_("Modes", 1, Object.entries(p.modes || {}).slice(0, 40).map(([m, rw]) => {
      const r = num(rw && rw[0]), w = num(rw && rw[1]); right += r; answered += r + w;
      return { key: [clean_(m, 40)], add: [r, w] };
    }), (row, a) => { row[1] += a[0]; row[2] += a[1]; row[3] = row[1] + row[2]; row[4] = pct_(row[1], row[3]); });

    let games = 0;
    upsert_("Games", 1, Object.entries(p.plays || {}).slice(0, 40).map(([g, n]) => { const k = num(n); games += k; return { key: [clean_(gameName_(g), 40)], add: [k] }; }),
      (row, a) => { row[1] += a[0]; row[2] = date; });

    upsert_("Daily", 2, [{ key: [date, year], add: [answered, right, games, active, fresh] }],
      (row, a) => { row[2] += a[0]; row[3] += a[1]; row[4] = pct_(row[3], row[2]); row[5] += a[2]; row[6] += a[3]; row[7] += a[4]; });

    const cards = (Array.isArray(p.cards) ? p.cards : []).slice(0, 300);
    upsert_("Cards", 1, cards.map(c => ({ key: [Math.floor(Number(c[0]) || 0)], add: [num(c[1]), num(c[2])], topic: clean_(c[3], 40), q: clean_(c[4], 120), ao: clean_(c[5], 20) }))
      .filter(c => c.key[0] > 0),
      (row, a, c) => { row[1] = c.topic || row[1]; row[2] = c.q || row[2]; row[3] += a[0]; row[4] += a[1]; row[5] = row[3] + row[4]; row[6] = pct_(row[3], row[5]); row[7] = c.ao || row[7]; });

    // Events: one new row per action, written in one go
    const evs = (Array.isArray(p.ev) ? p.ev : []).slice(-150), now = Date.now(), tz = Session.getScriptTimeZone(), visit = clean_(p.s, 12), device = clean_(p.dv, 16);
    const evRows = evs.filter(x => Array.isArray(x) && x[0]).map(x => {
      const t = Number(x[3]), when = t > now - 3 * 864e5 && t <= now + 6e4 ? new Date(t) : new Date();
      return [Utilities.formatDate(when, tz, "yyyy-MM-dd HH:mm"), visit, device, clean_(x[0], 40), clean_(x[1], 120), year, Math.max(0, Math.min(3600, Math.floor(Number(x[2]) || 0)))];
    });
    if (evRows.length) { const sh = SpreadsheetApp.getActive().getSheetByName("Events"), at = sh.getLastRow() + 1;
      sh.getRange(at, 1, evRows.length, 5).setNumberFormat("@"); sh.getRange(at, 1, evRows.length, 7).setValues(evRows); }

    buildSummary_();
  } catch (err) {
    // bad data is ignored
  } finally {
    lock.releaseLock();
  }
  return done_();
}

/** The Summary tab: headline numbers, hardest/easiest cards, weakest topics and AOs. Rebuilt on every update. */
function buildSummary() { setup(); buildSummary_(); }
function buildSummary_() {
  const ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName("Summary");
  const rows = name => { const s = ss.getSheetByName(name), n = s.getLastRow(); return n > 1 ? s.getRange(2, 1, n - 1, TABS[name].length).getValues() : []; };
  const daily = rows("Daily"), games = rows("Games"), cards = rows("Cards");
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const dayKey = v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(v);
  const sum = (arr, i, f) => arr.filter(f || (() => true)).reduce((t, r) => t + (Number(r[i]) || 0), 0);
  const qs = sum(daily, 2), rt = sum(daily, 3);
  const week = []; for (let i = 0; i < 7; i++) { const d = new Date(); d.setDate(d.getDate() - i); week.push(Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd")); }
  const byDay = {}; daily.forEach(r => { const k = dayKey(r[0]); byDay[k] = (byDay[k] || 0) + (Number(r[6]) || 0); });
  const busiest = Object.entries(byDay).sort((a, b) => b[1] - a[1])[0];
  const gameTotals = {}; games.forEach(r => { const n = gameName_(r[0]); gameTotals[n] = (gameTotals[n] || 0) + (Number(r[1]) || 0); });
  const gameList = Object.entries(gameTotals).sort((a, b) => b[1] - a[1]), topGame = gameList[0];

  const enough = cards.filter(r => Number(r[5]) >= MIN_ANSWERS);
  const cardRow = r => ["#" + r[0], r[1], r[2], r[5], r[6] + "%"];
  // with only a few cards answered 10+ times, split them so no card shows up as both hardest and easiest
  const byHard = enough.slice().sort((a, b) => a[6] - b[6] || b[5] - a[5]), nHard = Math.min(10, Math.ceil(byHard.length / 2));
  const hardRows = byHard.slice(0, nHard), hardIds = new Set(hardRows.map(r => r[0]));
  const hardest = hardRows.map(cardRow);
  const easiest = enough.filter(r => !hardIds.has(r[0])).sort((a, b) => b[6] - a[6] || b[5] - a[5]).slice(0, 10).map(cardRow);
  const group = i => { const g = {}; cards.forEach(r => { const k = r[i] || "(unknown)"; g[k] = g[k] || [0, 0]; g[k][0] += Number(r[3]) || 0; g[k][1] += Number(r[5]) || 0; });
    return Object.entries(g).filter(([, v]) => v[1] > 0).map(([k, v]) => [k, v[1], pct_(v[0], v[1]) + "%", pct_(v[0], v[1])]).sort((a, b) => a[3] - b[3]).map(r => r.slice(0, 3)); };
  const topics = group(1), aos = group(7);

  // Events: the newest rows only, so this stays quick as the tab grows
  const evSh = ss.getSheetByName("Events"), evLast = evSh.getLastRow(), evN = Math.min(EVENT_ROWS_READ, Math.max(0, evLast - 1));
  const evs = evN ? evSh.getRange(evLast - evN + 1, 1, evN, 7).getValues().filter(e => week.includes(String(e[0] instanceof Date ? dayKey(e[0]) : e[0]).slice(0, 10))) : [];
  const visits = new Set(evs.filter(e => e[1]).map(e => e[1])), devices = new Set(evs.filter(e => e[2]).map(e => e[2]));
  const secs = evs.filter(e => e[3] === "time").reduce((t, e) => t + (Number(e[6]) || 0), 0);
  const evCount = {}; evs.forEach(e => { if (e[3] !== "time") evCount[e[3]] = (evCount[e[3]] || 0) + 1; });
  const evList = Object.entries(evCount).sort((a, b) => b[1] - a[1]);

  const out = [];   // [row, col, values[][], bold?]
  const put = (r, c, vals, bold) => out.push([r, c, vals, bold]);
  put(1, 1, [["📊 Fount summary", "Updated " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "d MMM yyyy, HH:mm")]], true);
  put(3, 1, [["Headline numbers", ""]], true);
  put(4, 1, [
    ["Students (all time)", sum(daily, 7)],
    ["Active today", sum(daily, 6, r => dayKey(r[0]) === today)],
    ["Student-days, last 7 days", sum(daily, 6, r => week.includes(dayKey(r[0])))],
    ["Questions answered", qs],
    ["% right overall", pct_(rt, qs) + "%"],
    ["Busiest day", busiest ? busiest[0] + " (" + busiest[1] + " students)" : "–"],
    ["Most played game", topGame ? topGame[0] + " (" + topGame[1] + ")" : "–"],
    ["Devices, last 7 days", devices.size],
    ["Visits, last 7 days", visits.size],
    ["Average time per visit", visits.size ? Math.round(secs / visits.size / 60) + " min" : "–"],
  ]);
  const table = (r, c, title, head, body, empty) => { put(r, c, [[title].concat(new Array(head.length - 1).fill(""))], true); put(r + 1, c, [head], true);
    if (body.length) put(r + 2, c, body); else put(r + 2, c, [[empty].concat(new Array(head.length - 1).fill(""))]); return r + 3 + Math.max(1, body.length); };
  let r = 16;
  r = table(r, 1, "😬 Hardest cards (at least " + MIN_ANSWERS + " answers)", ["Card", "Topic", "Question", "Answers", "% right"], hardest, "Not enough answers yet");
  r = table(r, 1, "😎 Easiest cards (at least " + MIN_ANSWERS + " answers)", ["Card", "Topic", "Question", "Answers", "% right"], easiest, "Not enough answers yet");
  // right-hand tables start in column G, clear of the five-column card tables
  let r2 = table(3, 7, "📚 Topics, weakest first", ["Topic", "Answers", "% right"], topics, "No answers yet");
  r2 = table(r2, 7, "🎯 Assessment objectives, weakest first", ["AO", "Answers", "% right"], aos, "No answers yet");
  r2 = table(r2, 7, "🎮 Games, most played", ["Game", "Times started", ""], gameList.slice(0, 8).map(([g, n]) => [g, n, ""]), "No games yet");
  table(r2, 7, "⚡ Actions, last 7 days", ["Event", "Times", ""], evList.slice(0, 15).map(([k, n]) => [k, n, ""]), "No actions yet");

  sh.clear(); sh.clearFormats();
  const maxRow = Math.max(...out.map(([row, , vals]) => row + vals.length)), maxCol = 9;
  sh.getRange(1, 1, maxRow, maxCol).setNumberFormat("@").setFontWeight("normal");   // plain text, so "76%" stays as written
  out.forEach(([row, col, vals, bold]) => { const g = sh.getRange(row, col, vals.length, vals[0].length); g.setValues(vals.map(v => v.map(x => String(x)))); if (bold) g.setFontWeight("bold"); });
  sh.getRange(1, 1).setFontSize(16);
  [[1, 200], [2, 170], [3, 360], [4, 80], [5, 80], [6, 30], [7, 220], [8, 110], [9, 80]].forEach(([c, w]) => sh.setColumnWidth(c, w));
}

/** Older versions stored games by short code (e.g. "arcade:daily"): rename them and merge duplicates. */
function renameGames_() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Games"); if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues(), merged = {}, order = [];
  let changed = false;
  rows.forEach(([g, n, last]) => { const name = gameName_(g); if (name !== g) changed = true;
    if (!merged[name]) { merged[name] = [name, 0, last]; order.push(name); } else changed = true;
    merged[name][1] += Number(n) || 0; if (String(last) > String(merged[name][2])) merged[name][2] = last; });
  if (!changed) return;
  sh.getRange(2, 1, rows.length, 3).clearContent();
  sh.getRange(2, 1, order.length, 3).setValues(order.map(k => merged[k]));
}

/** Add to existing rows (matched on the first `keyCols` columns) or append new ones, in one write. */
function upsert_(tab, keyCols, items, apply) {
  if (!items.length) return;
  const sh = SpreadsheetApp.getActive().getSheetByName(tab), width = TABS[tab].length;
  const last = sh.getLastRow(), data = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  data.forEach(row => { for (let j = keyCols; j < width; j++) if (row[j] === "" && !TEXT_COLS.includes(TABS[tab][j])) row[j] = 0; });   // new columns on old rows
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

/** Run from the editor to clear all the data (keeps the tabs and headings). */
function wipeData() {
  const ss = SpreadsheetApp.getActive();
  ["Daily", "Games", "Modes", "Cards", "Events"].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });
  const summary = ss.getSheetByName("Summary");
  if (summary) summary.clear();
}

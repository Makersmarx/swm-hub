/**
 * SWM 12/13B Premier hub: backend for the public GitHub Pages site.
 * Paste into the team Google Sheet: Extensions > Apps Script, replace Code.gs, Save.
 * Then: Deploy > New deployment > Web app, Execute as "Me", Who has access "Anyone".
 * Run installTriggers() once to turn on the league refresh (daily 6 AM, plus Wed + Fri ~9 PM after games).
 *
 * Sheet layout it reads (as it exists today):
 *  - "Schedule & Attendance": player grid under the "Player Name" header (checkbox TRUE = in),
 *    game table under the "Game #" header. "As needed" in the Roster Type column marks a player
 *    who only joins the games checked on their row.
 *  - "Sheet2": extra Wed Nov 18 game, names in column A, "x" in Yes (B) or No (C).
 *  - "Replies": created by this script; holds "out" marks for regular games
 *    (a checkbox can only say in / not in).
 *  - "League": standings + fixtures, refreshed from EZLeagues by refreshLeague().
 */

const MAIN = 'Schedule & Attendance';
const REPLIES = 'Replies';
const LEAGUE = 'League';
const MIN_PLAYERS = 7;
const EXTRA_GAMES = [
  { id: 'x1', date: '2026-11-18', time: null, ha: 'TBC', opp: 'TBC', venue: 'Indoor',
    sheet: 'Sheet2', firstRow: 3, nameCol: 1, yesCol: 2, noCol: 3 },
];
const LEAGUE_URL = 'https://stjoekickers.ezleagues.ezfacility.com/leagues/481959/26-27-Indoor-1-MS-Oberliga.aspx';
const MON = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

/* ---------------- web app ---------------- */

function doGet() {
  try { return json_(readAll_()); } catch (err) { console.error(err); return json_({ error: 'unavailable' }); }
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad_request' }); }
  if (!body || typeof body !== 'object' || typeof body.pid !== 'string' || typeof body.gid !== 'string' || typeof body.v !== 'string') {
    return json_({ error: 'bad_request' });
  }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) { return json_({ error: 'busy' }); }
  try {
    const m = model_();
    const p = m.players.find(x => x.id === body.pid);
    const g = m.games.find(x => x.id === body.gid);
    const v = body.v;
    if (!p || !g || ['in', 'out', ''].indexOf(v) < 0) return json_({ error: 'invalid_argument' });
    writeRsvp_(m, p, g, v);
    SpreadsheetApp.flush();
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ error: 'unavailable' });
  } finally {
    lock.releaseLock();
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- reading ---------------- */

function model_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MAIN);
  const vals = sh.getDataRange().getDisplayValues();
  const hdrRow = vals.findIndex(r => r[1] === 'Player Name');
  const gameHdrRow = vals.findIndex(r => r[1] === 'Game #');
  if (hdrRow < 0 || gameHdrRow < 0) throw new Error('Sheet layout changed: header not found');

  const typeCol = vals[hdrRow].indexOf('Roster Type');
  const players = [];
  for (let r = hdrRow + 1; r < vals.length && vals[r][1] && vals[r][1] !== 'Confirmed Attending'; r++) {
    players.push({ id: 'p' + (r + 1), row: r + 1, full: vals[r][1].trim(), name: short_(vals[r][1]),
      asNeeded: typeCol >= 0 && /as needed/i.test(vals[r][typeCol]) });
  }

  const games = [];
  for (let r = gameHdrRow + 1; r < vals.length && /^Game \d+/.test(vals[r][1]); r++) {
    const n = Number(vals[r][1].match(/\d+/)[0]);
    const col = vals[hdrRow].findIndex(h => h.indexOf('Game ' + n + ' ') === 0) + 1;
    games.push({ id: 'g' + n, n: n, type: vals[r][7] || 'Regular', date: isoDate_(vals[r][2]), time: t24_(vals[r][3]),
      ha: vals[r][4], opp: vals[r][5], venue: vals[r][6], col: col });
  }
  EXTRA_GAMES.forEach(x => games.push(Object.assign({ n: 0, type: 'Extra' }, x)));
  return { ss: ss, sh: sh, vals: vals, players: players, games: games };
}

function readAll_() {
  const m = model_();
  const out = repliesSheet_(m.ss).getDataRange().getValues();
  const outKey = {};
  out.slice(1).forEach(r => { for (let c = 1; c < r.length; c++) if (r[c] === 'out') outKey[r[0] + '|' + out[0][c]] = true; });

  const extraVals = {};
  EXTRA_GAMES.forEach(x => { const s = m.ss.getSheetByName(x.sheet); extraVals[x.id] = s ? s.getDataRange().getDisplayValues() : []; });

  const players = m.players.map(p => {
    const rsvp = {};
    m.games.forEach(g => {
      if (g.type === 'Extra') {
        const row = extraVals[g.id].find((r, i) => i >= g.firstRow - 1 && r[g.nameCol - 1].trim() === p.full);
        if (row) rsvp[g.id] = /\S/.test(row[g.yesCol - 1]) ? 'in' : /\S/.test(row[g.noCol - 1]) ? 'out' : '';
      } else if (g.col > 0) {
        const checked = String(m.vals[p.row - 1][g.col - 1]).toUpperCase() === 'TRUE';
        rsvp[g.id] = checked ? 'in' : outKey[p.full + '|' + g.id] ? 'out' : '';
      }
    });
    return { id: p.id, name: p.name, order: p.row, asNeeded: p.asNeeded, rsvp: rsvp };
  });

  const games = m.games.map(g => ({ id: g.id, n: g.n, type: g.type, date: g.date, time: g.time, ha: g.ha, opp: g.opp, venue: g.venue }));
  return { team: 'SWM 12/13B Premier', minPlayers: MIN_PLAYERS, games: games, players: players, league: readLeague_(m.ss), at: new Date().toISOString() };
}

function readLeague_(ss) {
  const sh = ss.getSheetByName(LEAGUE);
  if (!sh) return null;
  const v = sh.getDataRange().getDisplayValues();
  if (v.length < 5) return null;
  const standings = [], fixtures = [];
  for (let r = 4; r < v.length; r++) {
    const x = v[r];
    if (x[1]) standings.push({ team: x[1], gp: +x[2], w: +x[3], t: +x[4], l: +x[5], gf: +x[6], ga: +x[7], gd: +x[8], pts: +x[9] });
    if (x[11]) fixtures.push({ date: x[11], time: x[12] || null, home: x[13], away: x[14],
      hs: x[15] === '' ? null : +x[15], as: x[16] === '' ? null : +x[16], status: x[17] });
  }
  return { league: v[0][0], scrapedAt: v[1][1], source: v[2][1], standings: standings, fixtures: fixtures };
}

/* ---------------- writing ---------------- */

function writeRsvp_(m, p, g, v) {
  if (g.type === 'Extra') {
    const s = m.ss.getSheetByName(g.sheet);
    const names = s.getRange(g.firstRow, g.nameCol, Math.max(s.getLastRow() - g.firstRow + 1, 1), 1).getDisplayValues();
    const i = names.findIndex(r => r[0].trim() === p.full);
    if (i < 0) return;
    s.getRange(g.firstRow + i, g.yesCol).setValue(v === 'in' ? 'x' : '');
    s.getRange(g.firstRow + i, g.noCol).setValue(v === 'out' ? 'x' : '');
    return;
  }
  m.sh.getRange(p.row, g.col).setValue(v === 'in');
  const rs = repliesSheet_(m.ss);
  const data = rs.getDataRange().getValues();
  let c = data[0].indexOf(g.id);
  if (c < 0) { c = data[0].length; rs.getRange(1, c + 1).setValue(g.id); }
  let r = data.findIndex((row, i) => i > 0 && row[0] === p.full);
  if (r < 0) { r = rs.getLastRow(); rs.getRange(r + 1, 1).setValue(p.full); }
  rs.getRange(r + 1, c + 1).setValue(v === 'out' ? 'out' : '');
}

function repliesSheet_(ss) {
  let s = ss.getSheetByName(REPLIES);
  if (!s) {
    s = ss.insertSheet(REPLIES);
    s.getRange(1, 1).setValue('Player (out marks for the hub; "in" lives in the checkboxes)');
  }
  return s;
}

/* ---------------- league refresh ---------------- */

function refreshLeague() {
  const page = UrlFetchApp.fetch(LEAGUE_URL, { headers: { 'User-Agent': 'Mozilla/5.0' }, muteHttpExceptions: true });
  if (page.getResponseCode() !== 200) throw new Error('EZLeagues returned ' + page.getResponseCode());
  const tables = tables_(page.getContentText());
  const st = tables.find(t => t.length && t[0].indexOf('PTS') >= 0);
  const fx = tables.find(t => t.length && t[0][0] === 'Date' && t[0][1] === 'Home');
  if (!st || !fx) throw new Error('standings or schedule table not found');

  const h = st[0], k = n => h.indexOf(n);
  const stand = st.slice(1).filter(r => r[0]).map(r => ({ team: r[0], gp: +r[k('GP')], w: +r[k('W')], t: +r[k('T')], l: +r[k('L')],
    gf: +r[k('GF')], ga: +r[k('GA')], gd: +r[k('GD')], pts: +r[k('PTS')] }));
  if (stand.length < 2) throw new Error('too few teams parsed');
  stand.sort((a, b) => b.pts - a.pts || b.gd - a.gd || b.gf - a.gf || a.team.localeCompare(b.team));

  const fix = fx.slice(1).filter(r => /\w+-\w{3} \d+/.test(r[0])).map(r => {
    const sc = r[2].match(/(\d+)\s*-\s*(\d+)/), t = t24_(r[4]);
    return [ezDate_(r[0]), t || '', r[1], r[3], sc ? +sc[1] : '', sc ? +sc[2] : '', sc ? 'final' : (t ? 'scheduled' : r[4])];
  });

  const rows = [['26-27 Indoor 1-MS Oberliga', ''], ['Updated (UTC)', new Date().toISOString()], ['Source', LEAGUE_URL],
    ['Pos', 'Team', 'GP', 'W', 'T', 'L', 'GF', 'GA', 'GD', 'Pts', '', 'Date', 'Time', 'Home', 'Away', 'Home goals', 'Away goals', 'Status']];
  for (let i = 0; i < Math.max(stand.length, fix.length); i++) {
    const r = new Array(18).fill('');
    if (i < stand.length) { const s = stand[i]; [i + 1, s.team, s.gp, s.w, s.t, s.l, s.gf, s.ga, s.gd, s.pts].forEach((x, j) => r[j] = x); }
    if (i < fix.length) fix[i].forEach((x, j) => r[11 + j] = x);
    rows.push(r);
  }
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(LEAGUE) || ss.insertSheet(LEAGUE);
  sh.getRange('A1:R200').clearContent();
  const width = 18;
  const padded = rows.map(r => r.concat(new Array(width - r.length).fill('')).slice(0, width));
  sh.getRange(1, 1, padded.length, width).setNumberFormat('@').setValues(padded);
}

function installTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'refreshLeague').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshLeague').timeBased().everyDays(1).atHour(6).create();
  // Same-night scores after Wed and Fri games (runs within the 9-10 PM hour, script time zone).
  [ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.FRIDAY].forEach(d =>
    ScriptApp.newTrigger('refreshLeague').timeBased().onWeekDay(d).atHour(21).create());
}

/* ---------------- helpers ---------------- */

function tables_(html) {
  const clean = c => c.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/\s+/g, ' ').trim();
  return (html.match(/<table[\s\S]*?<\/table>/g) || []).map(tb =>
    (tb.match(/<tr[\s\S]*?<\/tr>/g) || []).map(tr => (tr.match(/<t[hd][^>]*>[\s\S]*?<\/t[hd]>/g) || [])
      .map(c => clean(c.replace(/^<t[hd][^>]*>|<\/t[hd]>$/g, '')))).filter(r => r.some(Boolean)));
}

function ezDate_(d) { // "Wed-Nov 4" -> season runs Nov..Jan
  const m = d.match(/\w+-(\w{3}) (\d+)/);
  const mo = MON[m[1]];
  return (mo >= 8 ? 2026 : 2027) + '-' + pad_(mo) + '-' + pad_(+m[2]);
}

function isoDate_(s) { // "Fri, Nov 6, 2026"
  const m = String(s).match(/([A-Z][a-z]{2})[a-z]* (\d+), (\d{4})/);
  return m ? m[3] + '-' + pad_(MON[m[1]]) + '-' + pad_(+m[2]) : String(s);
}

function t24_(s) {
  const m = String(s || '').match(/(\d+):(\d+)\s*(AM|PM)/i);
  if (!m) return null;
  return pad_(+m[1] % 12 + (m[3].toUpperCase() === 'PM' ? 12 : 0)) + ':' + m[2];
}

function short_(full) {
  const w = String(full).trim().split(/\s+/);
  return w.length > 1 ? w[0] + ' ' + w[w.length - 1][0] + '.' : w[0];
}

function pad_(n) { return (n < 10 ? '0' : '') + n; }

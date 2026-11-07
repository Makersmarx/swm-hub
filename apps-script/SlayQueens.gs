/**
 * Slay Queens hub: backend for https://makersmarx.github.io/swm-hub/slay-queens/
 *
 * Setup (once):
 *  1. Make a new, empty Google Sheet. Extensions > Apps Script.
 *  2. Replace Code.gs with this file. Add a second file (+ > Script) named Config
 *     and paste the Config block you were given (roster link + coach names).
 *     Config is NOT committed to the public repo: the roster link exposes family contact info.
 *  3. Run setup() once and accept the permissions prompt.
 *  4. Deploy > New deployment > Web app, Execute as "Me", Who has access "Anyone". Send the /exec URL.
 *
 * Tabs it keeps (all created by setup):
 *  - "Roster":  id | roster key (Last, First) | hub name | role | active | order | source.
 *               Refreshed from the EZFacility roster page. Only names are stored, never phones or emails.
 *               Edit "hub name" for a nickname; add a row with source "manual" for a guest player.
 *  - "League":  standings + fixtures from EZLeagues. Her games are taken from these fixtures.
 *  - "Replies": in/out per player id (rows) per game id (columns).
 */

const TEAM = 'Slay Queens';
const TEAM_LEAGUE = 'Slay Queens FC'; // name as written on EZLeagues
const MIN_PLAYERS = 7;
const ROSTER = 'Roster', REPLIES = 'Replies', LEAGUE = 'League';
const LEAGUE_NAME = '26-27 Indoor 1-u12 Mittelliga';
const LEAGUE_URL = 'https://stjoekickers.ezleagues.ezfacility.com/leagues/482036/26-27-Indoor-1-u12-Mittelliga.aspx';
const MON = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };
const ROSTER_HEAD = ['id', 'Roster key (Last, First)', 'Hub name', 'Role', 'Active', 'Order', 'Source'];

/* ---------------- setup ---------------- */

function setup() {
  const ss = SpreadsheetApp.getActive();
  rosterSheet_(ss); repliesSheet_(ss);
  refreshRoster();
  refreshLeague();
  installTriggers();
}

function installTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['refreshLeague', 'refreshRoster'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshLeague').timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger('refreshLeague').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(15).create(); // after Sunday games
  ScriptApp.newTrigger('refreshRoster').timeBased().everyDays(1).atHour(5).create();
}

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
    const p = m.roster.find(x => x.id === body.pid && x.active && x.role === 'player');
    const g = m.games.find(x => x.id === body.gid);
    if (!p || !g || ['in', 'out', ''].indexOf(body.v) < 0) return json_({ error: 'invalid_argument' });
    writeRsvp_(m.ss, p.id, g.id, body.v);
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
  const league = readLeague_(ss);
  return { ss: ss, roster: readRoster_(ss), league: league, games: gamesFrom_(league) };
}

function readRoster_(ss) {
  const sh = ss.getSheetByName(ROSTER);
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues().slice(1).filter(r => r[0] && r[2]).map(r => ({
    id: r[0].trim(), name: r[2].trim(), role: (r[3] || 'player').trim().toLowerCase(),
    active: String(r[4]).toUpperCase() !== 'FALSE', order: Number(r[5]) || 999,
  }));
}

function gamesFrom_(L) {
  if (!L) return [];
  return L.fixtures
    .filter(f => f.home === TEAM_LEAGUE || f.away === TEAM_LEAGUE)
    .sort((a, b) => (a.date + (a.time || '')).localeCompare(b.date + (b.time || '')))
    .map((f, i) => {
      const home = f.home === TEAM_LEAGUE;
      return { id: 'g' + f.date.replace(/-/g, '') + (f.time || '').replace(':', ''), n: i + 1,
        type: f.type || 'Regular', date: f.date, time: f.time, ha: home ? 'Home' : 'Away',
        opp: home ? f.away : f.home, venue: f.venue || 'Indoor' };
    });
}

function readAll_() {
  const m = model_();
  const rep = repliesSheet_(m.ss).getDataRange().getDisplayValues();
  const gcol = {}; rep[0].forEach((h, c) => { if (c > 0 && h) gcol[h] = c; });
  const prow = {}; rep.forEach((r, i) => { if (i > 0 && r[0]) prow[r[0]] = r; });

  const players = m.roster.filter(p => p.active && p.role === 'player').sort((a, b) => a.order - b.order).map(p => {
    const rsvp = {}, row = prow[p.id];
    if (row) m.games.forEach(g => { const v = gcol[g.id] ? row[gcol[g.id]] : ''; if (v === 'in' || v === 'out') rsvp[g.id] = v; });
    return { id: p.id, name: p.name, order: p.order, rsvp: rsvp };
  });
  const coaches = m.roster.filter(p => p.active && p.role === 'coach').sort((a, b) => a.order - b.order).map(p => p.name);
  return { team: TEAM, minPlayers: MIN_PLAYERS, games: m.games, players: players, coaches: coaches,
    league: m.league, at: new Date().toISOString() };
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
      hs: x[15] === '' ? null : +x[15], as: x[16] === '' ? null : +x[16], status: x[17], venue: x[18] || '', type: x[19] || '' });
  }
  return { league: v[0][0], scrapedAt: v[1][1], source: v[2][1], standings: standings, fixtures: fixtures };
}

/* ---------------- writing ---------------- */

function writeRsvp_(ss, pid, gid, v) {
  const rs = repliesSheet_(ss);
  const data = rs.getDataRange().getValues();
  let c = data[0].indexOf(gid);
  if (c < 0) { c = Math.max(data[0].length, 1); rs.getRange(1, c + 1).setValue(gid); }
  let r = data.findIndex((row, i) => i > 0 && row[0] === pid);
  if (r < 0) { r = Math.max(rs.getLastRow(), 1); rs.getRange(r + 1, 1).setValue(pid); }
  rs.getRange(r + 1, c + 1).setValue(v);
}

function repliesSheet_(ss) {
  let s = ss.getSheetByName(REPLIES);
  if (!s) { s = ss.insertSheet(REPLIES); s.getRange(1, 1).setValue('player id'); s.setFrozenRows(1); }
  return s;
}

function rosterSheet_(ss) {
  let s = ss.getSheetByName(ROSTER);
  if (!s) { s = ss.insertSheet(ROSTER); s.getRange(1, 1, 1, ROSTER_HEAD.length).setValues([ROSTER_HEAD]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}

/* ---------------- roster refresh ---------------- */

function refreshRoster() {
  if (typeof CONFIG === 'undefined' || !CONFIG.rosterUrl) throw new Error('Add the Config file (rosterUrl, coaches) first');
  const res = UrlFetchApp.fetch(CONFIG.rosterUrl, { headers: { 'User-Agent': 'Mozilla/5.0' }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Roster page returned ' + res.getResponseCode());
  const rowsHtml = rows_(res.getContentText());
  const hi = rowsHtml.findIndex(r => r[0] === 'Name' && r.indexOf('Email') >= 0);
  if (hi < 0) throw new Error('roster table not found');
  const keys = [];
  for (let i = hi + 1; i < rowsHtml.length && rowsHtml[i].length >= 5 && rowsHtml[i][0].indexOf(',') > 0; i++) {
    keys.push(rowsHtml[i][0]); // only the name column is kept
  }
  if (keys.length < 3) throw new Error('too few roster names parsed');

  const coaches = (CONFIG.coaches || []).map(norm_);
  const ss = SpreadsheetApp.getActive(), sh = rosterSheet_(ss);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  let maxId = rows.reduce((mx, r) => Math.max(mx, Number(String(r[0]).replace(/\D/g, '')) || 0), 0);
  const seen = {};
  keys.forEach((k, i) => {
    const nk = norm_(k); seen[nk] = true;
    let r = rows.find(x => norm_(x[1]) === nk);
    if (!r) { r = ['p' + (++maxId), k, '', '', true, 0, 'roster']; rows.push(r); }
    r[3] = coaches.indexOf(nk) >= 0 ? 'coach' : 'player';
    r[4] = true; r[5] = i + 1; r[6] = 'roster';
  });
  rows.forEach(r => { if (r[6] === 'roster' && !seen[norm_(r[1])]) r[4] = false; });

  const taken = {}; rows.forEach(r => { if (r[2]) taken[r[2]] = true; });
  rows.forEach(r => { if (!r[2]) { r[2] = shortName_(r[1], taken); taken[r[2]] = true; } });

  sh.getRange(2, 1, Math.max(sh.getLastRow(), 2), ROSTER_HEAD.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, ROSTER_HEAD.length).setValues(rows.map(r => ROSTER_HEAD.map((_, j) => r[j] === undefined ? '' : r[j])));
}

function shortName_(key, taken) { // "Biehn, Abby" -> "Abby B."
  const parts = String(key).split(',');
  const last = (parts[0] || '').trim(), first = (parts[1] || '').trim().split(/\s+/)[0] || last;
  if (!parts[1]) return last;
  for (let n = 1; n <= last.length; n++) {
    const s = first + ' ' + last.slice(0, n) + '.';
    if (!taken[s]) return s;
  }
  return first + ' ' + last;
}

function norm_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }

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
    return [ezDate_(r[0]), t || '', r[1], r[3], sc ? +sc[1] : '', sc ? +sc[2] : '', sc ? 'final' : (t ? 'scheduled' : r[4]), r[5] || '', r[6] || ''];
  });

  const width = 20;
  const rows = [[LEAGUE_NAME], ['Updated (UTC)', new Date().toISOString()], ['Source', LEAGUE_URL],
    ['Pos', 'Team', 'GP', 'W', 'T', 'L', 'GF', 'GA', 'GD', 'Pts', '', 'Date', 'Time', 'Home', 'Away', 'Home goals', 'Away goals', 'Status', 'Venue', 'Game type']];
  for (let i = 0; i < Math.max(stand.length, fix.length); i++) {
    const r = new Array(width).fill('');
    if (i < stand.length) { const s = stand[i]; [i + 1, s.team, s.gp, s.w, s.t, s.l, s.gf, s.ga, s.gd, s.pts].forEach((x, j) => r[j] = x); }
    if (i < fix.length) fix[i].forEach((x, j) => r[11 + j] = x);
    rows.push(r);
  }
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(LEAGUE) || ss.insertSheet(LEAGUE);
  sh.getRange('A1:T300').clearContent();
  const padded = rows.map(r => r.concat(new Array(width).fill('')).slice(0, width));
  sh.getRange(1, 1, padded.length, width).setNumberFormat('@').setValues(padded);
}

/* ---------------- helpers ---------------- */

function tables_(html) {
  const clean = c => c.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/\s+/g, ' ').trim();
  return (html.match(/<table[\s\S]*?<\/table>/g) || []).map(tb =>
    (tb.match(/<tr[\s\S]*?<\/tr>/g) || []).map(tr => (tr.match(/<t[hd][^>]*>[\s\S]*?<\/t[hd]>/g) || [])
      .map(c => clean(c.replace(/^<t[hd][^>]*>|<\/t[hd]>$/g, '')))).filter(r => r.some(Boolean)));
}

function rows_(html) { // every leaf <tr> in the page as cell text, ignoring table nesting
  const clean = c => c.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/\s+/g, ' ').trim();
  return (html.match(/<tr[^>]*>(?:(?!<tr[\s>])[\s\S])*?<\/tr>/g) || []).map(tr =>
    (tr.match(/<t[hd][^>]*>[\s\S]*?<\/t[hd]>/g) || []).map(c => clean(c.replace(/^<t[hd][^>]*>|<\/t[hd]>$/g, ''))))
    .filter(r => r.some(Boolean));
}

function ezDate_(d) { // "Sun-Nov 8" -> season runs Nov..Jan
  const m = d.match(/\w+-(\w{3}) (\d+)/);
  const mo = MON[m[1]];
  return (mo >= 8 ? 2026 : 2027) + '-' + pad_(mo) + '-' + pad_(+m[2]);
}

function t24_(s) {
  const m = String(s || '').match(/(\d+):(\d+)\s*(AM|PM)/i);
  if (!m) return null;
  return pad_(+m[1] % 12 + (m[3].toUpperCase() === 'PM' ? 12 : 0)) + ':' + m[2];
}

function pad_(n) { return (n < 10 ? '0' : '') + n; }

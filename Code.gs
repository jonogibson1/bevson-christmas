/**
 * Bevson Christmas backend. Lives inside the Google Sheet (Extensions > Apps Script).
 * 1. Run setup() once. 2. Deploy > New deployment > Web app, execute as Me, access Anyone.
 * Tabs: Settings (edit to change the site), Guests (replies), Songs, Summary.
 */
const G = ['updated','name','coming','adults','kids','dogs','who','dog_names','phone','email','arrive','leave','bring_type','bringing','drinks','dietary','dietary_detail','santa','tunes','notes','token','cold_space'];
const NUM = { adults:1, kids:1, dogs:1, santa:1 };
const MAXLEN = { name:60, who:300, dog_names:120, tunes:400, phone:20, email:100, bring_type:40, bringing:200, drinks:300, dietary:200, dietary_detail:300, notes:500 };
const PRIVATE_SETTINGS = ['address'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mk = (name, head) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(head);
    sh.setFrozenRows(1); sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    return sh;
  };
  const st = mk('Settings', ['key', 'value', 'what it does']);
  if (st.getLastRow() === 1) st.getRange(2, 1, 13, 3).setValues([
    ['date', '2026-12-25', 'Event date, YYYY-MM-DD'],
    ['hosts', 'Jono & Charlotte', 'Shown at the top'],
    ['rsvp_by', '', 'RSVP deadline, YYYY-MM-DD. Blank hides it'],
    ['intro', '', 'Optional line under "The day"'],
    ['we_provide_food', '', 'One item per line (Alt+Enter for a new line in the cell)'],
    ['we_provide_drinks', '', 'One item per line'],
    ['please_bring', '', 'One item per line'],
    ['still_needed', '', 'Gaps you want filled. Blank hides it'],
    ['santa', '', 'Bastard Santa rules. Blank uses the built-in wording'],
    ['address', '', 'PRIVATE: only shown to people after they RSVP yes/maybe'],
    ['parking', '', ''],
    ['notes', '', 'Good to know: pool, pets, kids, etc. One per line'],
    ['contact', '', 'Who to ask. This is public, so first name only is safest'],
  ]);
  st.appendRow(['playlist_url', '', 'https:// link to a Spotify/Apple playlist']);
  st.getRange('A:C').setNumberFormat('@'); st.setColumnWidth(2, 420); st.getRange('B:B').setWrap(true);

  const g = mk('Guests', G);
  g.getRange(1, 1, g.getMaxRows(), G.length).setNumberFormat('@');
  G.forEach((k, i) => { if (NUM[k]) g.getRange(2, i + 1, g.getMaxRows() - 1, 1).setNumberFormat('0'); });
  const s = mk('Songs', ['id', 'title', 'artist', 'added_by', 'votes', 'voters', 'created']);
  s.getRange('A:D').setNumberFormat('@'); s.getRange('F:G').setNumberFormat('@');

  const sm = mk('Summary', ['', 'count']);
  if (sm.getLastRow() === 1) sm.getRange(2, 1, 7, 2).setValues([
    ['Adults (yes)', '=SUMIF(Guests!C:C,"yes",Guests!D:D)'],
    ['Kids (yes)', '=SUMIF(Guests!C:C,"yes",Guests!E:E)'],
    ['People (maybe)', '=SUMIF(Guests!C:C,"maybe",Guests!D:D)+SUMIF(Guests!C:C,"maybe",Guests!E:E)'],
    ["Households that can't come", '=COUNTIF(Guests!C:C,"no")'],
    ['Dogs (yes)', '=SUMIF(Guests!C:C,"yes",Guests!F:F)'],
    ['Bastard Santa players', '=SUMIF(Guests!C:C,"yes",Guests!R:R)+SUMIF(Guests!C:C,"maybe",Guests!R:R)'],
    ['Replies', '=COUNTA(Guests!B2:B)'],
  ]);
  const blank = ss.getSheetByName('Sheet1'); if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function doGet() { return out({ ok: true, hint: 'POST only' }); }

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    const token = String(b.token || '');
    if (!/^[\w-]{8,64}$/.test(token)) return out({ error: 'bad token' });
    if (b.website) return out(state(token)); // honeypot: pretend it worked
    if (b.action !== 'state') {
      const lock = LockService.getScriptLock(); lock.waitLock(20000);
      try {
        if (b.action === 'rsvp') rsvp(token, b.data || {});
        else if (b.action === 'song') song(token, b);
        else if (b.action === 'vote') vote(token, b);
        SpreadsheetApp.flush();
      } finally { lock.releaseLock(); }
    }
    return out(state(token));
  } catch (err) { return out({ error: String(err.message || err) }); }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sheet(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }
function rows(n) { const v = sheet(n).getDataRange().getValues(); const h = v.shift(); return v.map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; }); }
/** Trim, cap length, and stop a cell being read as a formula. */
function clean(v, max) { let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]+/g, ' ').trim().slice(0, max); if (/^[=+\-@]/.test(s)) s = "'" + s; return s; }
function int(v) { return Math.max(0, Math.min(20, parseInt(v, 10) || 0)); }

function rsvp(token, d) {
  const sh = sheet('Guests');
  if (!clean(d.name, 60)) throw new Error('name required');
  const coming = ['yes', 'maybe', 'no'].indexOf(d.coming) >= 0 ? d.coming : 'yes';
  const time = t => /^\d{2}:00$/.test(t) ? t : '';
  const row = G.map(k => {
    if (k === 'updated') return Utilities.formatDate(new Date(), 'Australia/Brisbane', 'd MMM HH:mm');
    if (k === 'token') return token;
    if (k === 'coming') return coming;
    if (NUM[k]) return int(d[k]);
    if (k === 'arrive' || k === 'leave') return time(d[k]);
    if (k === 'phone') return String(d.phone || '').replace(/[^\d+ ]/g, '').slice(0, 20);
    return clean(d[k], MAXLEN[k] || 100);
  });
  const tokens = sh.getRange(1, G.indexOf('token') + 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => r[0]);
  const at = tokens.indexOf(token);
  if (at < 0 && tokens.length > 300) throw new Error('guest list full');
  sh.getRange(at > 0 ? at + 1 : sh.getLastRow() + 1, 1, 1, G.length).setValues([row]);
}

function song(token, b) {
  const sh = sheet('Songs'), title = clean(b.title, 80);
  if (!title) throw new Error('title required');
  if (sh.getLastRow() > 200) throw new Error('song list full');
  const me = rows('Guests').filter(g => g.token === token)[0];
  sh.appendRow(['s' + Date.now().toString(36) + Math.floor(Math.random() * 1e4), title, clean(b.artist, 80), me ? me.name : '', 1, token, new Date().toISOString()]);
}

function vote(token, b) {
  const sh = sheet('Songs'), v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] === b.id) {
    const voters = String(v[i][5] || '').split(',').filter(String), at = voters.indexOf(token);
    if (at < 0) voters.push(token); else voters.splice(at, 1);
    sh.getRange(i + 1, 5, 1, 2).setValues([[voters.length, voters.join(',')]]);
    return;
  }
}

/** Everything the page may show. Phone, email, notes and dietary detail never leave the sheet. */
function state(token) {
  const settings = {}; let address = '';
  rows('Settings').forEach(r => { const k = String(r.key).trim(); if (!k || String(r.value) === '') return; if (PRIVATE_SETTINGS.indexOf(k) >= 0) address = String(r.value); else settings[k] = String(r.value); });
  const all = rows('Guests').filter(g => g.name);
  const mine = all.filter(g => g.token === token)[0] || null;
  const tally = k => { const t = {}; all.filter(g => g.coming !== 'no').forEach(g => String(g[k] || '').split(', ').filter(String).forEach(x => t[x] = (t[x] || 0) + 1)); return t; };
  if (mine) delete mine.token;
  return {
    settings: settings,
    mine: mine,
    address: mine && mine.coming !== 'no' ? address : '',
    guests: all.map(g => ({ name: g.name, coming: g.coming, adults: g.adults, kids: g.kids, dogs: g.dogs, arrive: g.arrive, leave: g.leave, bring_type: g.bring_type, bringing: g.bringing, santa: g.santa })),
    diets: tally('dietary'), drinks: tally('drinks'),
    songs: rows('Songs').filter(s => s.id).map(s => ({ id: s.id, title: s.title, artist: s.artist, by: s.added_by, votes: Number(s.votes) || 0, voted: String(s.voters || '').split(',').indexOf(token) >= 0 })),
  };
}

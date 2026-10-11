/* ============================================================================
   DAILY GAMES LEADERBOARD BACKEND — generic, multi-game
   Google Apps Script + Google Sheets. Zero cost, no auth, no accounts.

   One Sheet serves ALL of Anthony's daily mini-games:
     - "users" tab  -> username | display_name | phrase_hash | created_at
                       (one shared identity across every game)
     - one tab per game, tab name = game id (e.g. "daily-race")
                     -> date | username | score

   ----------------------------------------------------------------------------
   DEPLOY (one time, ~5 minutes):
   1. Create a new Google Sheet (e.g. "Daily Games Leaderboard").
   2. Extensions > Apps Script. Delete the default code, paste this whole file.
   3. Deploy > New deployment > type: Web app.
        Execute as:  Me
        Who has access:  Anyone
   4. Copy the web-app URL (ends in /exec) and paste it into the game's
      LEADERBOARD_URL constant (top of the app script in index.html).
   5. Done — tabs are created automatically on first use.
      NOTE: after editing this code, always create a NEW deployment version
      (Deploy > Manage deployments > edit > New version), or changes won't go live.

   BROWSER / CORS NOTES:
   - Clients POST JSON with Content-Type: text/plain (avoids the preflight
     OPTIONS request Apps Script can't answer). Responses are readable via
     normal fetch() because ContentService returns CORS-friendly output.
   - doGet also accepts ?callback=fn for JSONP fallback.

   USERNAME SECURITY MODEL (claimable names, no passwords on file):
   - Claim:  client hashes SHA-256(lowercase_name + "|" + phrase) and sends the
     hash. Server stores ONLY the hash — never the phrase, never plaintext.
   - Unlock: client re-sends the hash; server compares hashes.
   - Scores: client sends username + hash with every submit; server verifies
     the hash before writing. Unverified submissions are rejected.
   - "Remember me": the game keeps name+hash in localStorage and silently
     re-verifies on load. There is NO reset — if a player loses their phrase,
     the name is stuck (by design, keeps impersonation out).
   ============================================================================ */

var USERS_TAB = 'users';
var MAX_SCORE = 3600000; /* 1 hour in ms — sanity cap */

function _sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === USERS_TAB) sh.appendRow(['username', 'display_name', 'phrase_hash', 'created_at']);
    else sh.appendRow(['date', 'username', 'score']);
  }
  return sh;
}

function _out(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback && /^[A-Za-z0-9_]+$/.test(callback)) json = callback + '(' + json + ');';
  return ContentService.createTextOutput(json)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function _norm(u) { return String(u || '').trim().toLowerCase(); }
function _validUser(u) { return /^[A-Za-z0-9 ]{1,16}$/.test(String(u || '').trim()); }
function _validHash(h) { return /^[0-9a-f]{64}$/.test(String(h || '')); }

/* columns: 0=username(lower) 1=display_name 2=phrase_hash 3=created_at */
function _findUser(sh, username) {
  var key = _norm(username);
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === key)
      return { row: i + 1, display: String(vals[i][1]), hash: String(vals[i][2]) };
  }
  return null;
}

function _checkUser(username) {
  if (!_validUser(username)) return { status: 'invalid' };
  var u = _findUser(_sheet(USERS_TAB), username);
  return u ? { status: 'taken', display: u.display } : { status: 'free' };
}

function _claimUser(username, hash) {
  if (!_validUser(username)) return { ok: false, error: 'bad name' };
  if (!_validHash(hash)) return { ok: false, error: 'bad phrase' };
  var sh = _sheet(USERS_TAB);
  if (_findUser(sh, username)) return { ok: false, error: 'name taken' };
  sh.appendRow([_norm(username), String(username).trim(), String(hash), new Date().toISOString()]);
  return { ok: true };
}

function _verifyUser(username, hash) {
  if (!_validUser(username) || !_validHash(hash)) return { ok: false };
  var u = _findUser(_sheet(USERS_TAB), username);
  return { ok: !!u && u.hash === String(hash) };
}

/* one row per username per day per game — keeps the FASTER time */
function _submitScore(game, date, username, hash, score) {
  game = String(game || 'daily-race').replace(/[^\w-]/g, '').slice(0, 32) || 'daily-race';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ''))) return { ok: false, error: 'bad date' };
  score = Math.round(Number(score));
  if (!(score > 0) || score > MAX_SCORE) return { ok: false, error: 'bad score' };
  if (!_verifyUser(username, hash).ok) return { ok: false, error: 'not verified' };

  var sh = _sheet(game);
  var key = _norm(username), disp = String(username).trim();
  var vals = sh.getDataRange().getValues();
  var best = score, foundRow = -1, prev = score;
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(date) && String(vals[i][1]).toLowerCase() === key) {
      foundRow = i + 1; prev = Number(vals[i][2]) || score; break;
    }
  }
  if (foundRow > 0) {
    best = Math.min(prev, score);
    if (score < prev) sh.getRange(foundRow, 3).setValue(score);
    if (String(vals[foundRow - 1][1]) !== disp) sh.getRange(foundRow, 2).setValue(disp);
  } else {
    sh.appendRow([String(date), disp, score]);
  }
  return { ok: true, best: best };
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var a = d.action;
    if (a === 'check')  return _out(_checkUser(d.username));
    if (a === 'claim')  return _out(_claimUser(d.username, d.phrase_hash));
    if (a === 'verify') return _out(_verifyUser(d.username, d.phrase_hash));
    if (a === 'submit') return _out(_submitScore(d.game, d.date, d.username, d.phrase_hash, d.score));
    return _out({ ok: false, error: 'bad action' });
  } catch (err) {
    return _out({ ok: false, error: 'server error' });
  }
}

function doGet(e) {
  var game = String((e.parameter.game) || 'daily-race').replace(/[^\w-]/g, '').slice(0, 32) || 'daily-race';
  var date = String(e.parameter.date || '');
  var cb = e.parameter.callback;
  var out = { game: game, date: date, count: 0, entries: [] };
  try {
    var vals = _sheet(game).getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === date)
        rows.push({ username: String(vals[i][1]), score: Number(vals[i][2]) });
    }
    rows.sort(function (a, b) { return a.score - b.score; });
    out.count = rows.length;
    out.entries = rows.slice(0, 200);
  } catch (err) { /* brand-new tab: empty board */ }
  return _out(out, cb);
}

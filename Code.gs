/**
 * تحدي توت — سكريبت شيت جوجل
 * بيسجل المشتركين، وبيحسب الفايزين، وبيقفل المسابقة لوحده بعد أول ٥.
 *
 * غيّر كلمة السر دي قبل النشر (هي اللي بتفتح لوحة الفايزين في اللعبة):
 */
const ADMIN_KEY = 'غير-كلمة-السر-دي';
const MAX_WINNERS = 5;
const MIN_GAME_SECONDS = 25;   // أقل وقت ممكن لجولة حقيقية — أي فوز أسرع من كده بيترفض

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#FFB500');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}
const players_ = () => sheet_('المشتركين', ['الوقت', 'الاسم', 'الموبايل', 'العنوان', 'الكود الداخلي', 'عدد الجولات']);
const winners_ = () => sheet_('الفايزين', ['الترتيب', 'الوقت', 'الاسم', 'الموبايل', 'العنوان', 'كود الخصم', 'النقط', 'مدة الجولة (ث)']);

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function winnerCount_() { return Math.max(0, winners_().getLastRow() - 1); }
function isOpen_() { return winnerCount_() < MAX_WINNERS; }
function now_() { return Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd HH:mm:ss'); }
function clean_(s, max) { return String(s || '').replace(/[=+\-@]/, "'$&").slice(0, max); } // يمنع حقن معادلات في الشيت

function findPlayerByToken_(token) {
  const sh = players_(), n = sh.getLastRow();
  if (n < 2) return null;
  const vals = sh.getRange(2, 1, n - 1, 6).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][4] === token) return { row: i + 2, name: vals[i][1], phone: String(vals[i][2]), address: vals[i][3], plays: vals[i][5] };
  return null;
}
function findPlayerByPhone_(phone) {
  const sh = players_(), n = sh.getLastRow();
  if (n < 2) return null;
  const vals = sh.getRange(2, 1, n - 1, 6).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][2]).replace(/^'/, '') === phone) return { row: i + 2, token: vals[i][4] };
  return null;
}
function findWinner_(phone) {
  const sh = winners_(), n = sh.getLastRow();
  if (n < 2) return null;
  const vals = sh.getRange(2, 1, n - 1, 6).getValues();
  for (const v of vals) if (String(v[3]).replace(/^'/, '') === phone) return { rank: v[0], code: v[5] };
  return null;
}

function doGet(e) {
  if ((e.parameter.action || '') === 'status') return out_({ ok: true, open: isOpen_(), winners: winnerCount_(), max: MAX_WINNERS });
  return out_({ ok: false });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, reason: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cache = CacheService.getScriptCache();
    switch (body.action) {

      case 'register': {
        if (!isOpen_()) return out_({ ok: false, reason: 'closed' });
        const phone = String(body.phone || '').replace(/\D/g, '');
        if (!/^01[0125]\d{8}$/.test(phone)) return out_({ ok: false, reason: 'bad_phone' });
        const name = clean_(body.name, 60), address = clean_(body.address, 160);
        if (name.length < 3 || address.length < 6) return out_({ ok: false, reason: 'bad_data' });
        const existing = findPlayerByPhone_(phone);
        if (existing) {
          players_().getRange(existing.row, 2, 1, 3).setValues([[name, "'" + phone, address]]);
          return out_({ ok: true, token: existing.token });
        }
        const token = Utilities.getUuid();
        players_().appendRow([now_(), name, "'" + phone, address, token, 0]);
        return out_({ ok: true, token: token });
      }

      case 'start': {
        if (!isOpen_()) return out_({ ok: false, reason: 'closed' });
        const p = findPlayerByToken_(body.token);
        if (!p) return out_({ ok: false, reason: 'unknown' });
        const attempt = Utilities.getUuid();
        cache.put('a:' + attempt, JSON.stringify({ token: body.token, t: Date.now() }), 21600);
        players_().getRange(p.row, 6).setValue((Number(p.plays) || 0) + 1);
        return out_({ ok: true, attempt: attempt });
      }

      case 'win': {
        const p = findPlayerByToken_(body.token);
        if (!p) return out_({ ok: false, reason: 'unknown' });
        const prev = findWinner_(p.phone);
        if (prev) return out_({ ok: false, reason: 'already', code: prev.code });
        if (!isOpen_()) return out_({ ok: false, reason: 'closed' });
        const a = JSON.parse(cache.get('a:' + body.attempt) || 'null');
        if (!a || a.token !== body.token) return out_({ ok: false, reason: 'no_attempt' });
        const secs = (Date.now() - a.t) / 1000;
        const s = body.stats || {};
        if (secs < MIN_GAME_SECONDS || (s.caught && s.caught.letters !== 4)) return out_({ ok: false, reason: 'invalid' });
        cache.remove('a:' + body.attempt);
        const rank = winnerCount_() + 1;
        const code = 'TOOT-' + rank + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
        winners_().appendRow([rank, now_(), p.name, "'" + p.phone, p.address, code, Number(s.score) || 0, Math.round(secs)]);
        return out_({ ok: true, rank: rank, code: code });
      }

      case 'winners': {
        if (body.key !== ADMIN_KEY) return out_({ ok: false, reason: 'auth' });
        const sh = winners_(), n = sh.getLastRow();
        const rows = n < 2 ? [] : sh.getRange(2, 1, n - 1, 6).getDisplayValues();
        return out_({
          ok: true, open: isOpen_(), max: MAX_WINNERS, players: Math.max(0, players_().getLastRow() - 1),
          winners: rows.map(r => ({ rank: r[0], time: r[1], name: r[2], phone: r[3], address: r[4], code: r[5] }))
        });
      }
    }
    return out_({ ok: false, reason: 'bad_action' });
  } finally {
    lock.releaseLock();
  }
}

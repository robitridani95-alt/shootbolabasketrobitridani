/**
 * HAND HOOPS - API database Google Sheets
 *
 * Akun siswa: Nama Siswa + Kelas + NIS (NIS dipakai untuk masuk, harus unik).
 *
 * Cara pasang:
 * 1. Buat script dari Spreadsheet-nya langsung (Extensions > Apps Script),
 *    ATAU isi SPREADSHEET_ID di bawah kalau script berdiri sendiri.
 * 2. Pilih fungsi "setup" lalu klik Run (sekali saja, setujui izin akses).
 *    Sheet "Users" dan "Scores" akan otomatis dibuat.
 *    Kalau sheet lama punya format kolom berbeda (versi sebelumnya), sheet itu
 *    tidak dihapus: namanya diganti jadi "Users_lama_..." / "Scores_lama_..."
 *    dan sheet baru dibuat.
 * 3. Deploy > New deployment > Web app
 *    Execute as: Me | Who has access: Anyone
 *    (setiap ubah kode, deploy ulang dengan "New version").
 */

// Kosongkan kalau script dibuat dari dalam Spreadsheet (container-bound).
// Isi dengan ID spreadsheet (bagian antara /d/ dan /edit di URL) kalau script berdiri sendiri.
const SPREADSHEET_ID = '';

const SHEET_USERS = 'Users';
const SHEET_SCORES = 'Scores';
const USER_HEADERS = ['user_id', 'nis', 'nama', 'kelas', 'password_hash', 'salt', 'token', 'created_at', 'last_login'];
const SCORE_HEADERS = ['timestamp', 'nis', 'nama', 'kelas', 'score', 'makes', 'attempts', 'best_streak', 'accuracy'];
const TEXT_COLUMNS = ['nis', 'nama', 'kelas'];   // dipaksa teks agar "0123" atau "10-1" tidak berubah jadi angka/tanggal

const LEADERBOARD_SIZE = 5;
const MAX_SCORE = 1000;     // batas wajar skor satu game (cegah angka ngawur)
const MAX_ATTEMPTS = 50;

// Nama: 3-30 karakter (huruf, angka, spasi, titik, apostrof, - dan _)
const NAME_PATTERN = /^[A-Za-z0-9\u00C0-\u00FF][A-Za-z0-9\u00C0-\u00FF .'_-]{2,29}$/;
// NIS: 4-20 digit angka
const NIS_PATTERN = /^[0-9]{4,20}$/;
// Kelas: 1-15 karakter, mis. 7A, VII-A, X IPA 1
const KELAS_PATTERN = /^[A-Za-z0-9][A-Za-z0-9 .\/-]{0,14}$/;

/* ---------------------------- Setup ---------------------------- */

/**
 * Jalankan manual dari editor Apps Script (pilih "setup" > Run).
 * Membuat sheet Users dan Scores beserta header-nya kalau belum ada.
 */
function setup() {
  ensureSheets_();
  Logger.log('Setup selesai. Sheet "' + SHEET_USERS + '" dan "' + SHEET_SCORES + '" siap dipakai.');
}

function ensureSheets_() {
  getSheet_(SHEET_USERS, USER_HEADERS);
  getSheet_(SHEET_SCORES, SCORE_HEADERS);
}

/* ---------------------------- Endpoint ---------------------------- */

function doGet() {
  try {
    ensureSheets_(); // sheet otomatis muncul saat URL web app dibuka
    return json_({ ok: true, message: 'Hand Hoops API aktif' });
  } catch (err) {
    return json_({ ok: false, error: 'Gagal menyiapkan sheet: ' + err });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    ensureSheets_(); // pastikan sheet ada sebelum aksi apa pun
    const req = JSON.parse(e.postData.contents);
    let out;
    switch (req.action) {
      case 'register':    out = register_(req); break;
      case 'login':       out = login_(req); break;
      case 'session':     out = session_(req); break;
      case 'logout':      out = logout_(req); break;
      case 'saveScore':   out = saveScore_(req); break;
      case 'leaderboard': out = { ok: true, leaderboard: leaderboard_() }; break;
      default:            out = { ok: false, error: 'Aksi tidak dikenal.' };
    }
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: 'Server sibuk atau terjadi kesalahan: ' + err });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------------------------- Akun ---------------------------- */
// Kolom Users: 0 user_id | 1 nis | 2 nama | 3 kelas | 4 password_hash | 5 salt | 6 token | 7 created_at | 8 last_login

function register_(req) {
  const nama = cleanName_(req.nama);
  const kelas = cleanKelas_(req.kelas);
  const nis = cleanNis_(req.nis);
  const passHash = String(req.passHash || '');
  if (!NAME_PATTERN.test(nama)) return { ok: false, error: 'Nama 3-30 karakter: huruf, angka, spasi, titik, apostrof, atau - dan _ saja.' };
  if (!KELAS_PATTERN.test(kelas)) return { ok: false, error: 'Kelas 1-15 karakter, mis. 7A atau X IPA 1.' };
  if (!NIS_PATTERN.test(nis)) return { ok: false, error: 'NIS harus angka 4-20 digit.' };
  if (passHash.length < 8 || passHash.length > 128) return { ok: false, error: 'Password tidak valid.' };

  const sheet = getSheet_(SHEET_USERS, USER_HEADERS);
  if (findUser_(sheet, nis)) return { ok: false, error: 'NIS ini sudah terdaftar. Silakan masuk.' };

  const salt = Utilities.getUuid();
  const token = Utilities.getUuid();
  const now = new Date();
  sheet.appendRow([Utilities.getUuid(), nis, nama, kelas, hash_(salt, passHash), salt, token, now, now]);
  return { ok: true, nis: nis, nama: nama, kelas: kelas, token: token };
}

function login_(req) {
  const sheet = getSheet_(SHEET_USERS, USER_HEADERS);
  const user = findUser_(sheet, cleanNis_(req.nis));
  if (!user || user.values[4] !== hash_(user.values[5], String(req.passHash || ''))) {
    return { ok: false, error: 'NIS atau password salah.' };
  }
  const token = Utilities.getUuid();
  sheet.getRange(user.row, 7).setValue(token);        // token baru (kolom G)
  sheet.getRange(user.row, 9).setValue(new Date());   // last_login (kolom I)
  return { ok: true, nis: String(user.values[1]), nama: String(user.values[2]), kelas: String(user.values[3]), token: token };
}

function session_(req) {
  const user = authUser_(req);
  return user
    ? { ok: true, nis: String(user.values[1]), nama: String(user.values[2]), kelas: String(user.values[3]) }
    : { ok: false, error: 'Sesi berakhir.' };
}

function logout_(req) {
  const user = authUser_(req);
  if (user) getSheet_(SHEET_USERS, USER_HEADERS).getRange(user.row, 7).setValue('');
  return { ok: true };
}

/* ---------------------------- Skor ---------------------------- */
// Kolom Scores: 0 timestamp | 1 nis | 2 nama | 3 kelas | 4 score | 5 makes | 6 attempts | 7 best_streak | 8 accuracy

function saveScore_(req) {
  const user = authUser_(req);
  if (!user) return { ok: false, error: 'Sesi berakhir, silakan masuk lagi.' };

  const score = clampInt_(req.score, 0, MAX_SCORE);
  const attempts = clampInt_(req.attempts, 0, MAX_ATTEMPTS);
  const makes = clampInt_(req.makes, 0, attempts);
  const bestStreak = clampInt_(req.bestStreak, 0, attempts);
  const accuracy = attempts > 0 ? makes / attempts : 0;

  getSheet_(SHEET_SCORES, SCORE_HEADERS)
    .appendRow([new Date(), String(user.values[1]), String(user.values[2]), String(user.values[3]),
                score, makes, attempts, bestStreak, accuracy]);
  return { ok: true, leaderboard: leaderboard_() };
}

// Papan peringkat hanya memuat nama & kelas (NIS tidak dikirim ke publik).
function leaderboard_() {
  const sheet = getSheet_(SHEET_SCORES, SCORE_HEADERS);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const rows = sheet.getRange(2, 1, last - 1, SCORE_HEADERS.length).getValues();
  const best = {};
  rows.forEach(function (r) {
    const key = String(r[1]).trim();
    const score = Number(r[4]) || 0;
    const b = best[key] || (best[key] = { nama: '', kelas: '', best: 0, games: 0 });
    b.nama = String(r[2]);
    b.kelas = String(r[3]);
    b.games++;
    if (score > b.best) b.best = score;
  });
  return Object.keys(best).map(function (k) { return best[k]; })
    .sort(function (a, b) { return b.best - a.best || a.games - b.games; })
    .slice(0, LEADERBOARD_SIZE);
}

/* ---------------------------- Helper ---------------------------- */

function authUser_(req) {
  const token = String(req.token || '');
  if (!token) return null;
  const user = findUser_(getSheet_(SHEET_USERS, USER_HEADERS), cleanNis_(req.nis));
  return (user && String(user.values[6]) === token) ? user : null;
}

function findUser_(sheet, nis) {
  if (!nis) return null;
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === nis) return { row: i + 2, values: rows[i] };
  }
  return null;
}

function getSpreadsheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Spreadsheet tidak ditemukan. Buat script dari menu Extensions > Apps Script di dalam Spreadsheet, atau isi SPREADSHEET_ID.');
  }
  return ss;
}

function headersMatch_(sheet, headers) {
  if (sheet.getLastColumn() < headers.length) return false;
  const row = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(row[i]) !== headers[i]) return false;
  }
  return true;
}

function getSheet_(name, headers) {
  const ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);

  // Sheet lama dengan format kolom berbeda: jangan dihapus, cukup diganti namanya sebagai arsip.
  if (sheet && sheet.getLastRow() > 0 && !headersMatch_(sheet, headers)) {
    const stamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
    sheet.setName(name + '_lama_' + stamp);
    sheet = null;
  }
  if (!sheet) sheet = ss.insertSheet(name);

  // Sheet baru / kosong: tulis header dan atur kolom teks
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    TEXT_COLUMNS.forEach(function (col) {
      const idx = headers.indexOf(col);
      if (idx >= 0) sheet.getRange(1, idx + 1, sheet.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return sheet;
}

function hash_(salt, passHash) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + passHash, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function cleanName_(v) {
  return String(v || '').trim().replace(/\s+/g, ' ');
}

function cleanKelas_(v) {
  return String(v || '').trim().replace(/\s+/g, ' ').toUpperCase();
}

function cleanNis_(v) {
  return String(v || '').trim().replace(/\s+/g, '');
}

function clampInt_(v, min, max) {
  const n = Math.round(Number(v));
  if (!isFinite(n)) return min;
  return Math.max(min, Math.min(max, n));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

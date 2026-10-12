/**
 * カチカチラリー ランキング受付
 * スプレッドシートの「拡張機能 > Apps Script」に、この中身をまるごと貼り付けて使います。
 *
 * シート「records」に1プレイ1行で記録します（列：日時 / ニックネーム / 対戦相手 / ラリー数 / 週）
 * 消したい記録は、シートの行を消すだけでOKです。
 */

const SHEET_NAME = 'records';
const MAX_RALLY = 300;          // これより大きい記録は受け付けない
const NAME_MAX = 10;
const OPPONENTS = ['やさしい先輩', 'ママさん', 'エース', 'ベテランじいちゃん', 'ナマイキ後輩', 'ガチ勢くん', '監督'];
const NG_WORDS = ['しね', '死ね', 'ころす', '殺す', 'うんこ', 'ちんこ', 'まんこ', 'fuck', 'shit', 'sex'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['日時', 'ニックネーム', '対戦相手', 'ラリー数', '週']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 月曜はじまりの週（日本時間）を "2026-10-12" の形で返す
function weekKey_(d) {
  const j = new Date(d.getTime() + 9 * 3600 * 1000);
  const day = (j.getUTCDay() + 6) % 7;
  const m = new Date(Date.UTC(j.getUTCFullYear(), j.getUTCMonth(), j.getUTCDate() - day));
  return m.toISOString().slice(0, 10);
}

function cleanName_(s) {
  s = String(s || '').replace(/[\u0000-\u001f<>=]/g, '').trim().slice(0, NAME_MAX);
  // 先頭の = + - @ は数式扱いされるので取り除く
  s = s.replace(/^[=+\-@]+/, '');
  const low = s.toLowerCase();
  if (NG_WORDS.some(function (w) { return low.indexOf(w) >= 0; })) return '';
  return s;
}

// ゲームから記録を受け取る
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const name = cleanName_(body.name);
    const opp = String(body.opp || '');
    const rally = Math.floor(Number(body.rally));
    if (!name || OPPONENTS.indexOf(opp) < 0 || !(rally >= 1 && rally <= MAX_RALLY)) {
      return json_({ ok: false });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      const now = new Date();
      sheet_().appendRow([now, name, opp, rally, weekKey_(now)]);
      CacheService.getScriptCache().remove('board-' + weekKey_(now));
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false });
  }
}

// ゲームにランキングを返す（今週・歴代、ニックネームごとのベスト上位10）
function doGet(e) {
  const cache = CacheService.getScriptCache();
  const wk = weekKey_(new Date());
  const cached = cache.get('board-' + wk);
  if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);

  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, 5).getValues() : [];
  const records = rows
    .filter(function (r) { return r[1] && r[3]; })
    .map(function (r) { return { name: String(r[1]), opp: String(r[2]), rally: Number(r[3]), week: String(r[4]) }; });

  const out = JSON.stringify({
    week: aggregate_(records.filter(function (r) { return r.week === wk; })),
    all: aggregate_(records),
    totalPlays: records.length
  });
  cache.put('board-' + wk, out, 20); // 20秒だけキャッシュ
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function aggregate_(records) {
  const by = {};
  records.forEach(function (r) {
    const a = by[r.name] || (by[r.name] = { name: r.name, rally: -1, opp: '', plays: 0 });
    a.plays++;
    if (r.rally > a.rally) { a.rally = r.rally; a.opp = r.opp; }
  });
  return Object.keys(by).map(function (k) { return by[k]; })
    .sort(function (a, b) { return b.rally - a.rally || b.plays - a.plays; })
    .slice(0, 10);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

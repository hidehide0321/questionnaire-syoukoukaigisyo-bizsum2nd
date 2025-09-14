/**
 * ふなばしビジネスサミット申込フォーム backend (GAS)
 * 固定スキーマ（列名）で保存します。列の並びは任意ですが、必須の列名が必要です。
 */

// 必要に応じて設定（コンテナバインドなら空でOK）
const SHEET_ID = '';
const SHEET_NAME = '';

// 必須ヘッダー（列名）
const REQUIRED_HEADERS = [
  'タイムスタンプ',
  'お名前',
  'フリガナ',
  'メールアドレス',
  '事業所名',
  '業種',
  '所属団体など',
  '出身地',
  '電話番号',
  // PR は現表記/旧表記どちらかが存在すればOK
  '活動・自己PR'
];

/**
 * Webアプリの受け口（application/x-www-form-urlencoded）
 */
function doPost(e) {
  try {
    var params = e && e.parameter ? e.parameter : {};

    var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    var sh = (function () {
      if (SHEET_NAME) return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
      return ss.getSheets()[0] || ss.insertSheet('フォーム');
    })();

    // ヘッダー取得/初期化
    var headers = getOrInitHeaders(sh);

    // 必須列の存在チェック（PRは別処理で旧表記も許容）
    headers = ensureRequiredHeaders(sh, headers);

    // ヘッダー名→列インデックスの辞書（0-based）
    var idx = indexMap(headers);
    var prCol = idx['活動・自己PR'];
    if (prCol == null) prCol = headers.indexOf('活動自己PR'); // 旧表記対応
    if (prCol < 0) throw new Error('ヘッダーに「活動・自己PR」列がありません');

    // シートの列数に合わせた行配列を作る
    var row = new Array(headers.length).fill('');
    // 値を対応する列へ格納
    if (idx['タイムスタンプ'] >= 0) row[idx['タイムスタンプ']] = new Date();
    if (idx['お名前'] >= 0)         row[idx['お名前']] = (params.name || '');
    if (idx['フリガナ'] >= 0)       row[idx['フリガナ']] = (params.furigana || '');
    if (idx['メールアドレス'] >= 0) row[idx['メールアドレス']] = (params.email || '');
    if (idx['事業所名'] >= 0)       row[idx['事業所名']] = (params.company || '');
    if (idx['業種'] >= 0)           row[idx['業種']] = (params.industry || '');
    if (idx['所属団体など'] >= 0)   row[idx['所属団体など']] = (params.affiliation || '');
    if (idx['出身地'] >= 0)         row[idx['出身地']] = (params.hometown || '');
    if (idx['電話番号'] >= 0)       row[idx['電話番号']] = (params.tel || '');
    if (prCol >= 0)                  row[prCol] = (params.pr || '');

    // email列のフォールバック（英語系の別名列がある場合にも書き込む）
    // 既に「メールアドレス」列がある場合はそちらを優先し、
    // 無い場合のみ英語別名列へフォールバックする
    if (headers.indexOf('メールアドレス') < 0 && params && params.email) {
      var __emailAliases = ['Email', 'E-mail', 'email', 'mail', 'Mail'];
      for (var __k = 0; __k < __emailAliases.length; __k++) {
        var __col = headers.indexOf(__emailAliases[__k]);
        if (__col >= 0) { row[__col] = (params.email || '').toString().trim(); break; }
      }
    }
    // メールアドレス列を必ず特定（見出しの別名があればそれを使用、無ければ新規作成）
    try {
      var __emailIdx = ensureEmailColumn(sh, headers);
      row[__emailIdx] = (params.email || '').toString().trim();
    } catch (_e) {}

    sh.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: (err && err.message) || String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrInitHeaders(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow === 0) {
    // 空シートなら固定ヘッダーを作る（順は任意だが初期値として）
    var init = [
      'タイムスタンプ','お名前','フリガナ','メールアドレス','事業所名','業種','所属団体など','出身地','電話番号','活動・自己PR'
    ];
    sh.getRange(1, 1, 1, init.length).setValues([init]);
    return init;
  }
  var lastCol = Math.max(1, sh.getLastColumn());
  return sh.getRange(1, 1, 1, lastCol).getValues()[0];
}

function ensureRequiredHeaders(headers) {
  // PRは現表記 or 旧表記どちらかがあればOK
  var needs = REQUIRED_HEADERS.filter(function(h){ return h !== '活動・自己PR'; });
  for (var i = 0; i < needs.length; i++) {
    if (headers.indexOf(needs[i]) === -1) {
      throw new Error('ヘッダーに「' + needs[i] + '」列がありません');
    }
  }
  if (headers.indexOf('活動・自己PR') === -1 && headers.indexOf('活動自己PR') === -1) {
    throw new Error('ヘッダーに「活動・自己PR」列がありません（旧表記「活動自己PR」でも可）');
  }
}

function indexMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) map[headers[i]] = i;
  return map;
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// メールアドレス列を見出しの別名も含めて探し、無ければ1行目末尾に新規作成して列インデックス（0-based）を返す
function ensureEmailColumn(sh, headers) {
  var aliases = ['メールアドレス', 'Email', 'E-mail', 'email', 'mail', 'Mail'];
  for (var i = 0; i < aliases.length; i++) {
    var idx = headers.indexOf(aliases[i]);
    if (idx >= 0) return idx;
  }
  // 見つからなければ「メールアドレス」列を追加
  headers.push('メールアドレス');
  var col = headers.length; // 1-based列番号
  sh.getRange(1, col, 1, 1).setValue('メールアドレス');
  return col - 1; // 0-based
}

// --- 追加: 必須ヘッダー自動追加版の再定義 ---
function ensureRequiredHeaders(sh, headers) {
  var updated = headers.slice();
  // PR 列（自己紹介PR or 自己PR）の存在確認
  if (updated.indexOf('�����E����PR') === -1 && updated.indexOf('��������PR') === -1) {
    updated.push('�����E����PR');
    sh.getRange(1, updated.length, 1, 1).setValue('�����E����PR');
  }
  // その他の必須列を自動付与
  for (var i = 0; i < REQUIRED_HEADERS.length; i++) {
    var key = REQUIRED_HEADERS[i];
    if (key === '�����E����PR') continue;
    if (updated.indexOf(key) === -1) {
      updated.push(key);
      sh.getRange(1, updated.length, 1, 1).setValue(key);
    }
  }
  return updated;
}

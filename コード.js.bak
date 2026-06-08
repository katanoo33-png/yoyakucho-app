/**
 * 厚生局提出スケジューラー — GAS バックエンド
 *
 * 【デプロイ手順】
 *   1. このファイルの内容を Google Apps Script エディタに貼り付ける
 *   2. スクリプトプロパティに以下を設定:
 *        MASTER_SS_ID  = 1QmBFIvQklZLllhN5kbyR8Z1Ozy04gfrM2vG7_4XLqG8
 *   3. デプロイ → 新しいデプロイ → ウェブアプリ
 *        実行ユーザー : 自分
 *        アクセスできるユーザー : 全員
 *   4. 発行されたURLをReactアプリの「GAS URL設定」に入力
 *
 * 【スプレッドシート】
 *   患者マスター : MASTER_SS_ID で指定したGSS の「マスター」シート
 *   スケジュール保存先 : このスクリプトが紐付いている GSS
 *                        （厚生局提出スケジュール.gsheet）の「スケジュール」シート
 */

var MASTER_SS_ID    = PropertiesService.getScriptProperties().getProperty('MASTER_SS_ID')
                      || '1QmBFIvQklZLllhN5kbyR8Z1Ozy04gfrM2vG7_4XLqG8';
var MASTER_SHEET    = 'マスター';
var SCHEDULE_SHEET  = 'スケジュール';

// ──────────────────────────────────────────────
// エントリーポイント（GET）
// ──────────────────────────────────────────────
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var result;
  try {
    if (action === 'getPatients') {
      result = getHygienePatients_();
    } else if (action === 'getSchedule') {
      result = getSchedule_(parseInt(e.parameter.year), parseInt(e.parameter.month));
    } else if (action === 'listSaves') {
      result = listNamedSaves_();
    } else if (action === 'loadSave') {
      result = loadNamedSave_(e.parameter.name || '');
    } else if (action === 'getEmployees') {
      result = getEmployees_();
    } else {
      result = { ok: false, error: '不明なアクション: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return jsonOut_(result);
}

// ──────────────────────────────────────────────
// エントリーポイント（POST）
// ──────────────────────────────────────────────
function doPost(e) {
  var result;
  try {
    var action = e.parameter.action;
    if (action === 'saveSchedule') {
      result = saveSchedule_(JSON.parse(e.parameter.data || '[]'));
    } else if (action === 'saveNamed') {
      result = saveNamedSnapshot_(e.parameter.name || '', JSON.parse(e.parameter.data || '[]'));
    } else if (action === 'deleteSave') {
      result = deleteNamedSave_(e.parameter.name || '');
    } else {
      result = { ok: false, error: '不明なアクション: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return jsonOut_(result);
}

// ──────────────────────────────────────────────
// 患者マスター取得
// O列=TRUE の患者の A,B,P,Q,R,S 列を返す
// ──────────────────────────────────────────────
function getHygienePatients_() {
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName(MASTER_SHEET);
  if (!sheet) return { ok: false, error: 'マスターシートが見つかりません' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, data: [] };

  // A〜S列（19列）を一括取得
  var data     = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  var patients = [];
  data.forEach(function(row, i) {
    var flag = row[14]; // O列 (index 14)
    if (flag !== true && String(flag).toLowerCase() !== 'true') return;
    var lastName  = String(row[0] || '').trim(); // A列: 患者氏名（漢字）
    var firstName = String(row[1] || '').trim(); // B列: 患者カナ
    var fullName  = lastName; // 表示名はA列（漢字）のみ
    if (!fullName) return;
    patients.push({
      row:       i + 2,
      lastName:  lastName,
      firstName: firstName,
      name:      fullName,
      youbi:     fmtYoubi_(row[15]),              // P列: 曜日
      time:      fmtTime_(row[16]),              // Q列: 時間
      doctor:    String(row[17] || '').trim(),   // R列: 担当医師
      hygienist: String(row[18] || '').trim()    // S列: 担当衛生士
    });
  });
  return { ok: true, data: patients };
}

// ──────────────────────────────────────────────
// 月別スケジュール取得
// ──────────────────────────────────────────────
function getSchedule_(year, month) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SCHEDULE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  var ym   = year + '-' + pad_(month);
  var data = sheet.getDataRange().getValues();
  var records = [];
  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][0];
    // Sheetsがセルを Date 型で返す場合は ISO 形式に変換
    var dateStr = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(rawDate || '');
    if (!dateStr || dateStr.indexOf(ym) !== 0) continue;
    records.push({
      id:          String(data[i][6] || ''),
      date:        dateStr,
      patientName: String(data[i][1] || ''),
      time:        String(data[i][2] || ''),
      doctor:      String(data[i][3] || ''),
      hygienist:   String(data[i][4] || ''),
      isNew:       data[i][5] === true || String(data[i][5]).toLowerCase() === 'true',
      note:        String(data[i][7] || '')
    });
  }
  return { ok: true, data: records };
}

// ──────────────────────────────────────────────
// スケジュール保存（月単位で差し替え）
// ──────────────────────────────────────────────
function saveSchedule_(visits) {
  if (!visits || !visits.length) return { ok: false, error: 'データがありません' };
  var ym = String(visits[0].date).substring(0, 7); // 'YYYY-MM'

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SCHEDULE_SHEET) || ss.insertSheet(SCHEDULE_SHEET);

  // ヘッダー初期化
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, 8)
         .setValues([['日付', '患者名', '時間', '担当医師', '担当衛生士', '新規', 'ID', 'メモ']]);
    sheet.setFrozenRows(1);
  }

  // 対象月の既存行を削除（後ろから）
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = dates.length - 1; i >= 0; i--) {
      if (String(dates[i][0]).indexOf(ym) === 0) sheet.deleteRow(i + 2);
    }
  }

  // 新規行を追加
  var rows = visits.map(function(v) {
    return [v.date, v.patientName, v.time, v.doctor, v.hygienist,
            v.isNew || false, v.id || '', v.note || ''];
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  }

  SpreadsheetApp.flush();
  return { ok: true, message: rows.length + '件を保存しました' };
}

// ──────────────────────────────────────────────
// 名前付きスナップショット一覧取得
// ──────────────────────────────────────────────
function listNamedSaves_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('スナップショット');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  // 行: [保存名, 件数, 保存日時, (unused)]
  // 同名は最後の行（最新）だけを返す
  var seen = {};
  var result = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var name = String(data[i][0] || '').trim();
    if (!name || seen[name]) continue;
    seen[name] = true;
    result.unshift({
      name:      name,
      count:     Number(data[i][1]) || 0,
      createdAt: String(data[i][2] || '')
    });
  }
  return { ok: true, data: result };
}

// ──────────────────────────────────────────────
// 名前付きスナップショット読み込み
// ──────────────────────────────────────────────
function loadNamedSave_(name) {
  if (!name) return { ok: false, error: '保存名が空です' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('スナップショット');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, data: [] };

  var data = sheet.getDataRange().getValues();
  var records = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== name) continue;
    // 列: 保存名, 件数, 保存日時, JSON（訪問記録全体）
    var json = String(data[i][3] || '[]');
    try { records = JSON.parse(json); } catch(e) {}
    // 最後に見つかった行（最新）を使う
  }
  return { ok: true, data: records };
}

// ──────────────────────────────────────────────
// 名前付きスナップショット保存（同名は上書き）
// ──────────────────────────────────────────────
function saveNamedSnapshot_(name, visits) {
  if (!name) return { ok: false, error: '保存名が空です' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('スナップショット') || ss.insertSheet('スナップショット');

  // ヘッダー初期化
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, 4).setValues([['保存名', '件数', '保存日時', 'データJSON']]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, 30); // JSONカラムを細く
  }

  // 同名の既存行を削除
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = names.length - 1; i >= 0; i--) {
      if (String(names[i][0]).trim() === name) sheet.deleteRow(i + 2);
    }
  }

  // 新規行追加
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sheet.appendRow([name, visits.length, now, JSON.stringify(visits)]);
  SpreadsheetApp.flush();
  return { ok: true, message: '「' + name + '」として ' + visits.length + '件を保存しました' };
}

// ──────────────────────────────────────────────
// 名前付きスナップショット削除
// ──────────────────────────────────────────────
function deleteNamedSave_(name) {
  if (!name) return { ok: false, error: '保存名が空です' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('スナップショット');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, message: '対象なし' };

  var lastRow = sheet.getLastRow();
  var names   = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var deleted = 0;
  for (var i = names.length - 1; i >= 0; i--) {
    if (String(names[i][0]).trim() === name) { sheet.deleteRow(i + 2); deleted++; }
  }
  SpreadsheetApp.flush();
  return { ok: true, message: '「' + name + '」を削除しました（' + deleted + '行）' };
}

// ──────────────────────────────────────────────
// ユーティリティ
// ──────────────────────────────────────────────
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pad_(n) {
  return String(n).padStart(2, '0');
}

function fmtTime_(val) {
  if (!val) return '';
  if (typeof val === 'string') {
    var m = val.match(/(\d{1,2})時(\d{2})分?/);
    if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
    m = val.match(/(\d{1,2})時/);
    if (m) return ('0' + m[1]).slice(-2) + ':00';
    m = val.match(/(\d{1,2}):(\d{2})/);
    if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
    return val;
  }
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
  }
  return String(val);
}

// ──────────────────────────────────────────────
// 従業員リスト取得（医師・衛生士）
// シート「従業員リスト」: A列=名前 B列=役職（医師/衛生士）
// ──────────────────────────────────────────────
function getEmployees_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('従業員リスト');
  if (!sheet || sheet.getLastRow() < 1) {
    return { ok: true, data: { doctors: [], hygienists: [] } };
  }
  // ヘッダーなし: A列=役職（歯科医師/歯科衛生士）, B列=名前（行1からデータ）
  var data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  var doctors = [], hygienists = [];
  data.forEach(function(row) {
    var role = String(row[0] || '').trim(); // A列: 役職
    var name = String(row[1] || '').trim(); // B列: 名前
    if (!name) return;
    if (role === '歯科医師') doctors.push(name);
    else if (role === '歯科衛生士') hygienists.push(name);
  });
  return { ok: true, data: { doctors: doctors, hygienists: hygienists } };
}

function fmtYoubi_(val) {
  return String(val || '').trim().replace(/曜日$/, '').replace(/曜$/, '');
}

// ──────────────────────────────────────────────
// 動作確認用（GASエディタから手動実行）
// ──────────────────────────────────────────────

/**
 * 従業員リスト取得のデバッグ
 * GASエディタで [実行] → ログで結果を確認する
 */
function debugGetEmployees() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('従業員リスト');
  if (!sheet) {
    Logger.log('ERROR: シート「従業員リスト」が見つかりません');
    Logger.log('スプレッドシート内のシート一覧:');
    ss.getSheets().forEach(function(s) { Logger.log('  - ' + s.getName()); });
    return;
  }
  Logger.log('✓ シート「従業員リスト」を発見');
  Logger.log('  最終行: ' + sheet.getLastRow() + '  最終列: ' + sheet.getLastColumn());

  if (sheet.getLastRow() < 1) {
    Logger.log('ERROR: データが1行もありません');
    return;
  }

  var cols = Math.max(2, sheet.getLastColumn());
  var data = sheet.getRange(1, 1, sheet.getLastRow(), cols).getValues();
  Logger.log('── 全行データ ──');
  data.forEach(function(row, i) {
    Logger.log('行' + (i + 1) + ':  A=[' + row[0] + ']  B=[' + row[1] + ']');
  });

  var result = getEmployees_();
  Logger.log('── getEmployees_() 戻り値 ──');
  Logger.log(JSON.stringify(result));
  Logger.log('歯科医師: ' + (result.data ? result.data.doctors.join('、') : '取得失敗'));
  Logger.log('歯科衛生士: ' + (result.data ? result.data.hygienists.join('、') : '取得失敗'));
}

function debugGetPatients() {
  var result = getHygienePatients_();
  Logger.log('取得件数: ' + (result.data ? result.data.length : 0));
  if (result.data) {
    result.data.forEach(function(p) {
      Logger.log(p.name + ' / ' + p.youbi + ' / ' + p.time + ' / Dr.' + p.doctor + ' / Hyg.' + p.hygienist);
    });
  }
}

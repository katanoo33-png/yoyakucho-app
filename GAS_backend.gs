/**
 * 厚生局提出スケジューラー — GAS バックエンド
 *
 * 【デプロイ手順】
 *   1. このファイルの内容を Google Apps Script エディタに貼り付ける
 *   2. スクリプトプロパティに以下を設定:
 *        MASTER_SS_ID    = マスターシートのスプレッドシートID
 *        DRIVE_FOLDER_ID = 出力Excelの保存先GoogleドライブフォルダID
 *   3. デプロイ → 新しいデプロイ → ウェブアプリ
 *        実行ユーザー : 自分
 *        アクセスできるユーザー : 全員
 *   4. 発行されたURLをReactアプリの「GAS URL設定」に入力
 *
 * 【スプレッドシート】
 *   患者マスター : MASTER_SS_ID で指定したGSS の「マスター」シート
 *   スケジュール保存先 : このスクリプトが紐付いている GSS の「スケジュール」シート
 */

var MASTER_SS_ID    = PropertiesService.getScriptProperties().getProperty('MASTER_SS_ID')
                      || '1QmBFIvQklZLllhN5kbyR8Z1Ozy04gfrM2vG7_4XLqG8';
var MASTER_SHEET    = 'マスター';
var SCHEDULE_SHEET  = 'スケジュール';
var DRIVE_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID') || '';

// ──────────────────────────────────────────────
// エントリーポイント（GET）
// ──────────────────────────────────────────────
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var result;
  try {
    if (action === 'getPatients') {
      // includeAll=1 で厚生局対象フラグ(O列)に関わらず全患者を返す
      var includeAll = (e.parameter.includeAll === '1' || e.parameter.includeAll === 'true');
      result = getHygienePatients_(includeAll);
    } else if (action === 'getSchedule') {
      result = getSchedule_(parseInt(e.parameter.year), parseInt(e.parameter.month));
    } else if (action === 'listSaves') {
      result = listNamedSaves_();
    } else if (action === 'loadSave') {
      result = loadNamedSave_(e.parameter.name || '');
    } else if (action === 'getEmployees') {
      result = getEmployees_();
    } else if (action === 'getServiceTimes') {
      result = getKirokuServiceTimes_(
        e.parameter.name || '',
        parseInt(e.parameter.year),
        parseInt(e.parameter.month)
      );
    } else if (action === 'getRegisteredNames') {
      result = getRegisteredNames_(
        parseInt(e.parameter.year),
        parseInt(e.parameter.month)
      );
    } else if (action === 'getAdminLinks') {
      result = getAdminLinks_();
    } else if (action === 'saveKiroku') {
      // POST リダイレクト時のボディ消失を避けるため doGet 経由でも受け付ける
      var kirokuData;
      try { kirokuData = JSON.parse(e.parameter.data || '[]'); } catch(_) { kirokuData = []; }
      result = saveKiroku_(
        e.parameter.name  || '',
        e.parameter.kana  || '',
        e.parameter.kaigo === '1',
        parseInt(e.parameter.year  || '0', 10),
        parseInt(e.parameter.month || '0', 10),
        kirokuData
      );
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
// URLSearchParams形式とJSON形式の両方に対応
// ──────────────────────────────────────────────
function doPost(e) {
  var result;
  try {
    // JSON ボディ優先、なければ URLSearchParams パラメータを使用
    var p = e.parameter;
    var contentType = (e.postData && e.postData.type) || '';
    if (contentType.indexOf('application/json') !== -1 || contentType.indexOf('text/plain') !== -1) {
      try { p = JSON.parse(e.postData.contents); } catch (_) {}
    }

    var action = p.action;
    if (action === 'saveSchedule') {
      var data = typeof p.data === 'string' ? JSON.parse(p.data) : p.data;
      result = saveSchedule_(data || []);
    } else if (action === 'saveNamed') {
      var data2 = typeof p.data === 'string' ? JSON.parse(p.data) : p.data;
      result = saveNamedSnapshot_(p.name || '', data2 || []);
    } else if (action === 'deleteSave') {
      result = deleteNamedSave_(p.name || '');
    } else if (action === 'sendEmail') {
      result = sendExportEmail_(p.to, p.filename, p.base64);
    } else if (action === 'saveToDrive') {
      result = saveToDrive_(p.filename, p.base64);
    } else if (action === 'updatePatientTimes') {
      var times = typeof p.times === 'string' ? JSON.parse(p.times) : p.times;
      result = updatePatientTimes_(parseInt(p.row, 10), times || {});
    } else if (action === 'saveKiroku') {
      var kirokuData = typeof p.data === 'string' ? JSON.parse(p.data) : p.data;
      result = saveKiroku_(
        p.name   || '',
        p.kana   || '',
        p.kaigo,
        parseInt(p.year,  10),
        parseInt(p.month, 10),
        kirokuData || []
      );
    } else {
      result = { ok: false, error: '不明なアクション: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message };
  }
  return jsonOut_(result);
}

// ──────────────────────────────────────────────
// 患者マスター取得（P列=介護区分追加後の新列構成）
// A:患者氏名 B:カナ C:曜日 D:週パターン E:時間
// F-J:第1〜5週フラグ  O:厚生局対象フラグ
// P(15):介護区分（あり/なし） Q(16):訪問曜日
// R(17):訪問時間_開始 S(18):訪問時間_終了
// T(19):Dr居宅Ⅰ_開始 U(20):Dr居宅Ⅰ_終了
// V(21):Dr居宅Ⅱ_開始 W(22):Dr居宅Ⅱ_終了
// X(23):訪問衛生_開始 Y(24):訪問衛生_終了
// Z(25):衛士居宅Ⅰ_開始 AA(26):衛士居宅Ⅰ_終了
// AB(27):診察医師 AC(28):衛生士
// ──────────────────────────────────────────────
function getHygienePatients_(includeAll) {
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName(MASTER_SHEET);
  if (!sheet) return { ok: false, error: 'マスターシートが見つかりません' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, data: [] };

  // A〜AC列（29列）を一括取得
  var data     = sheet.getRange(2, 1, lastRow - 1, 29).getValues();
  var patients = [];
  data.forEach(function(row, i) {
    var flag = row[14]; // O(14): 厚生局対象フラグ
    // includeAll の場合は O列フィルタをスキップ（全患者）
    if (!includeAll && flag !== true && String(flag).toLowerCase() !== 'true') return;
    var lastName  = String(row[0] || '').trim(); // A: 患者氏名
    var firstName = String(row[1] || '').trim(); // B: 患者カナ
    if (!lastName) return;
    var careLevelRaw = String(row[15] || '').trim(); // P(15): 介護区分
    patients.push({
      row:       i + 2,
      lastName:  lastName,
      firstName: firstName,
      name:      lastName,
      careLevel: (careLevelRaw === 'あり') ? 'あり' : 'なし', // P(15)
      youbi:     fmtYoubi_(row[16]),             // Q(16): 厚生局_曜日
      time:      fmtTime_(row[4]),               // E(4):  時間
      weekFlags: [                               // F-J(5-9): 第1〜5週
        row[5] === true  || String(row[5]).toLowerCase()  === 'true',
        row[6] === true  || String(row[6]).toLowerCase()  === 'true',
        row[7] === true  || String(row[7]).toLowerCase()  === 'true',
        row[8] === true  || String(row[8]).toLowerCase()  === 'true',
        row[9] === true  || String(row[9]).toLowerCase()  === 'true'
      ],
      svcVisitTime:         fmtTime_(row[17]),   // R(17): 訪問時間_開始
      svcVisitTimeEnd:      fmtTime_(row[18]),   // S(18): 訪問時間_終了
      svcDrType1Start:      fmtTime_(row[19]),   // T(19): Dr居宅Ⅰ_開始
      svcDrType1End:        fmtTime_(row[20]),   // U(20): Dr居宅Ⅰ_終了
      svcDrType2Start:      fmtTime_(row[21]),   // V(21): Dr居宅Ⅱ_開始
      svcDrType2End:        fmtTime_(row[22]),   // W(22): Dr居宅Ⅱ_終了
      svcHygieneVisitStart: fmtTime_(row[23]),   // X(23): 訪問衛生指導_開始
      svcHygieneVisitEnd:   fmtTime_(row[24]),   // Y(24): 訪問衛生指導_終了
      svcHygieneType1Start: fmtTime_(row[25]),   // Z(25): 衛士居宅Ⅰ_開始
      svcHygieneType1End:   fmtTime_(row[26]),   // AA(26): 衛士居宅Ⅰ_終了
      doctor:    safeStr_(row[27]),  // AB(27): 診察医師
      hygienist: safeStr_(row[28])   // AC(28): 衛生士
    });
  });
  return { ok: true, data: patients };
}

// ──────────────────────────────────────────────
// 月別スケジュール取得
// ──────────────────────────────────────────────
function getSchedule_(year, month) {
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
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
      id:              String(data[i][6] || ''),
      date:            dateStr,
      patientName:     String(data[i][1] || ''),
      time:            String(data[i][2] || ''),
      doctor:          String(data[i][3] || ''),
      hygienist:       String(data[i][4] || ''),
      isNew:           data[i][5] === true || String(data[i][5]).toLowerCase() === 'true',
      note:            String(data[i][7] || ''),
      svcVisitTime:         String(data[i][8]  || ''),
      svcVisitTimeEnd:      String(data[i][9]  || ''),
      svcDrType1Start:      String(data[i][10] || ''),
      svcDrType1End:        String(data[i][11] || ''),
      svcDrType2Start:      String(data[i][12] || ''),
      svcDrType2End:        String(data[i][13] || ''),
      svcHygieneVisitStart: String(data[i][14] || ''),
      svcHygieneVisitEnd:   String(data[i][15] || ''),
      svcHygieneType1Start: String(data[i][16] || ''),
      svcHygieneType1End:   String(data[i][17] || '')
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

  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName(SCHEDULE_SHEET) || ss.insertSheet(SCHEDULE_SHEET);

  var NEW_HEADERS = ['日付', '患者名', '時間', '担当医師', '担当衛生士', '新規', 'ID', 'メモ',
                     '訪問時間_開始', '訪問時間_終了',
                     'Dr居宅Ⅰ_開始', 'Dr居宅Ⅰ_終了',
                     'Dr居宅Ⅱ_開始', 'Dr居宅Ⅱ_終了',
                     '訪問衛生_開始', '訪問衛生_終了',
                     '衛士居宅Ⅰ_開始', '衛士居宅Ⅰ_終了'];
  var COL_COUNT = 18;

  // ヘッダー初期化（新規シート or 旧列数シートの拡張）
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, COL_COUNT).setValues([NEW_HEADERS]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < COL_COUNT) {
    var cur = sheet.getLastColumn();
    sheet.getRange(1, cur + 1, 1, COL_COUNT - cur).setValues([NEW_HEADERS.slice(cur)]);
  }

  // 対象月の既存行を削除（後ろから）
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = dates.length - 1; i >= 0; i--) {
      if (String(dates[i][0]).indexOf(ym) === 0) sheet.deleteRow(i + 2);
    }
  }

  // 新規行を追加（18列）
  var rows = visits.map(function(v) {
    return [v.date, v.patientName, v.time, v.doctor, v.hygienist,
            v.isNew || false, v.id || '', v.note || '',
            v.svcVisitTime         || '',
            v.svcVisitTimeEnd      || '',
            v.svcDrType1Start      || '',
            v.svcDrType1End        || '',
            v.svcDrType2Start      || '',
            v.svcDrType2End        || '',
            v.svcHygieneVisitStart || '',
            v.svcHygieneVisitEnd   || '',
            v.svcHygieneType1Start || '',
            v.svcHygieneType1End   || ''];
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, COL_COUNT).setValues(rows);
  }

  SpreadsheetApp.flush();
  return { ok: true, message: rows.length + '件を保存しました' };
}

// ──────────────────────────────────────────────
// 患者マスターのサービス時間と介護区分を更新
// P(col16)介護区分  R(col18)〜AA(col27) の10列を上書き
// ──────────────────────────────────────────────
function updatePatientTimes_(rowNum, times) {
  if (!rowNum || rowNum < 2) return { ok: false, error: '無効な行番号です' };
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName(MASTER_SHEET);
  if (!sheet) return { ok: false, error: 'マスターシートが見つかりません' };

  // 介護区分を P(col16) に書き込む（渡された場合のみ）
  if (times.careLevel !== undefined) {
    sheet.getRange(rowNum, 16, 1, 1).setValue(times.careLevel);
  }

  // R(col18)〜AA(col27): 10列分のサービス時間を書き込む
  var vals = [[
    times.svcVisitTime         || '',  // R(18)
    times.svcVisitTimeEnd      || '',  // S(19)
    times.svcDrType1Start      || '',  // T(20)
    times.svcDrType1End        || '',  // U(21)
    times.svcDrType2Start      || '',  // V(22)
    times.svcDrType2End        || '',  // W(23)
    times.svcHygieneVisitStart || '',  // X(24)
    times.svcHygieneVisitEnd   || '',  // Y(25)
    times.svcHygieneType1Start || '',  // Z(26)
    times.svcHygieneType1End   || ''   // AA(27)
  ]];
  sheet.getRange(rowNum, 18, 1, 10).setValues(vals);
  SpreadsheetApp.flush();
  return { ok: true, message: '患者時間を更新しました' };
}

// ──────────────────────────────────────────────
// 名前付きスナップショット一覧取得
// ──────────────────────────────────────────────
function listNamedSaves_() {
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
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
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
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
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
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
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
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
  if (val === '' || val === null || val === undefined) return '';
  // Google Sheets が時刻を数値（1日の分数）で返す場合の変換
  if (typeof val === 'number') {
    var totalMin = Math.round(val * 24 * 60);
    var h = Math.floor(totalMin / 60);
    var m = totalMin % 60;
    return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
  }
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
  }
  if (typeof val === 'string') {
    var s = val.trim();
    var m = s.match(/(\d{1,2})時(\d{2})分?/);
    if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
    m = s.match(/(\d{1,2})時/);
    if (m) return ('0' + m[1]).slice(-2) + ':00';
    m = s.match(/(\d{1,2}):(\d{2})/);
    if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
    return s;
  }
  return String(val);
}

// ──────────────────────────────────────────────
// 従業員リスト取得（医師・衛生士）
// シート「従業員リスト」: A列=名前 B列=役職（医師/衛生士）
// ──────────────────────────────────────────────
function getEmployees_() {
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName('従業員リスト');
  if (!sheet || sheet.getLastRow() < 1) {
    return { ok: true, data: { doctors: [], hygienists: [] } };
  }
  // ヘッダーなし: A列=役職（歯科医師/歯科衛生士）, B列=名前（行1からデータ）
  var data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
  var doctors = [], hygienists = [];
  data.forEach(function(row) {
    var role = safeStr_(row[0]); // A列: 役職
    var name = safeStr_(row[1]); // B列: 名前
    if (!name) return;
    if (role === '歯科医師') doctors.push(name);
    else if (role === '歯科衛生士') hygienists.push(name);
  });
  return { ok: true, data: { doctors: doctors, hygienists: hygienists } };
}

function fmtYoubi_(val) {
  return String(val || '').trim().replace(/曜日$/, '').replace(/曜$/, '');
}

// Date オブジェクトや数値が混入した場合に空文字を返す（名前フィールド用）
function safeStr_(val) {
  if (!val) return '';
  if (val instanceof Date) return '';
  if (typeof val === 'number') return '';
  return String(val).trim();
}

// ──────────────────────────────────────────────
// メール送信（Excelファイルを添付）
// ──────────────────────────────────────────────
function sendExportEmail_(to, filename, base64) {
  if (!to || !to.trim()) return { ok: false, error: '送付先メールアドレスが空です' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(to.trim())) return { ok: false, error: 'メールアドレスの形式が正しくありません: ' + to };
  if (!base64) return { ok: false, error: 'ファイルデータが空です（Excelビルドに失敗した可能性があります）' };

  var decoded;
  try {
    decoded = Utilities.base64Decode(base64);
  } catch (e) {
    return { ok: false, error: 'ファイルのデコードに失敗しました: ' + e.message };
  }

  var blob = Utilities.newBlob(
    decoded,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    filename || 'schedule.xlsx'
  );

  var FROM_ADDRESS = 'umemura.hanzoumon.929@gmail.com';

  GmailApp.sendEmail(
    to.trim(),
    '【厚生局提出スケジュール】' + (filename || 'schedule.xlsx'),
    '添付ファイルをご確認ください。\n\n訪問歯科スケジューラーより自動送信',
    {
      from: FROM_ADDRESS,
      attachments: [blob]
    }
  );
  return { ok: true, message: FROM_ADDRESS + ' から ' + to.trim() + ' に送信しました' };
}

// ──────────────────────────────────────────────
// Google Drive 保存
// DRIVE_FOLDER_ID が未設定の場合はマイドライブのルートに保存
// ──────────────────────────────────────────────
function saveToDrive_(filename, base64) {
  if (!base64) return { ok: false, error: 'ファイルデータが空です' };

  var decoded;
  try {
    decoded = Utilities.base64Decode(base64);
  } catch (e) {
    return { ok: false, error: 'ファイルのデコードに失敗しました: ' + e.message };
  }

  var blob = Utilities.newBlob(
    decoded,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    filename || 'schedule.xlsx'
  );

  var folder;
  if (DRIVE_FOLDER_ID) {
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch (e) {
      return { ok: false, error: 'DriveフォルダIDが無効です（スクリプトプロパティ DRIVE_FOLDER_ID を確認してください）: ' + e.message };
    }
  } else {
    folder = DriveApp.getRootFolder();
  }

  // 同名ファイルがあれば上書き（削除→新規作成）
  var existing = folder.getFilesByName(filename);
  while (existing.hasNext()) { existing.next().setTrashed(true); }
  var file = folder.createFile(blob);
  return { ok: true, url: file.getUrl(), folderName: folder.getName() };
}

// ──────────────────────────────────────────────
// 動作確認用（GASエディタから手動実行）
// ──────────────────────────────────────────────

/**
 * 従業員リスト取得のデバッグ
 * GASエディタで [実行] → ログで結果を確認する
 */
function debugGetEmployees() {
  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
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

function debugSendTestEmail() {
  GmailApp.sendEmail(
    'katanoo33@gmail.com',
    'GAS権限テスト',
    'このメールはGmail権限の承認テストです。',
    { from: 'umemura.hanzoumon.929@gmail.com' }
  );
  Logger.log('送信完了');
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

// ──────────────────────────────────────────────
// その月にkiroku登録済みの患者名一覧を取得
// ──────────────────────────────────────────────
function getRegisteredNames_(year, month) {
  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName('厚生局_訪問時間記録');
  if (!sheet) return { ok: true, data: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, data: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues(); // A〜F列
  var names = {};
  data.forEach(function(row) {
    if (String(row[4]) === String(year) && String(row[5]) === String(month)) {
      var nm = String(row[1] || '').trim(); // B列: 患者氏名
      if (nm) names[nm] = true;
    }
  });
  return { ok: true, data: Object.keys(names) };
}

// ──────────────────────────────────────────────
// kiroku（厚生局_訪問時間記録）から週別サービス時間を取得
// ──────────────────────────────────────────────
function getKirokuServiceTimes_(name, year, month) {
  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName('厚生局_訪問時間記録');
  if (!sheet) return { ok: true, data: [] };

  var data = sheet.getDataRange().getValues();
  // ヘッダー行スキップ（行1がヘッダーの場合）
  var rows = data.filter(function(row, idx) {
    if (idx === 0) return false; // header
    return String(row[1]).trim() === String(name).trim()
      && String(row[4]) === String(year)
      && String(row[5]) === String(month);
  });

  if (rows.length === 0) return { ok: true, data: [] };

  rows.sort(function(a, b) { return Number(a[6]) - Number(b[6]); });

  function toTimeStr(v) {
    if (!v) return '';
    if (typeof v === 'string' && /^\d{1,2}:\d{2}/.test(v)) return v.substring(0, 5);
    if (v instanceof Date) {
      var h = String(v.getHours()).padStart(2, '0');
      var m = String(v.getMinutes()).padStart(2, '0');
      return h + ':' + m;
    }
    if (typeof v === 'number') {
      var totalMin = Math.round(v * 1440);
      var hh = Math.floor(totalMin / 60);
      var mm = totalMin % 60;
      return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
    }
    return String(v).substring(0, 5);
  }

  var result = rows.map(function(row) {
    return {
      svcVisitTime:         toTimeStr(row[7]),
      svcVisitTimeEnd:      toTimeStr(row[8]),
      svcDrType1Start:      toTimeStr(row[9]),
      svcDrType1End:        toTimeStr(row[10]),
      svcDrType2Start:      toTimeStr(row[11]),
      svcDrType2End:        toTimeStr(row[12]),
      svcHygieneVisitStart: toTimeStr(row[13]),
      svcHygieneVisitEnd:   toTimeStr(row[14]),
      svcHygieneType1Start: toTimeStr(row[15]),
      svcHygieneType1End:   toTimeStr(row[16]),
    };
  });

  return { ok: true, data: result };
}

// ──────────────────────────────────────────────
// 厚生局_訪問時間記録シートへの書き込み（kiroku_index.html から呼ばれる）
//
// 列構成（getKirokuServiceTimes_ の row[] 参照インデックスで確定済み）:
//   A(0):登録日時  B(1):患者氏名  C(2):カナ  D(3):介護区分
//   E(4):年  F(5):月  G(6):週
//   H(7):訪問開始   I(8):訪問終了
//   J(9):指導Ⅰ開始  K(10):指導Ⅰ終了
//   L(11):指導Ⅱ開始  M(12):指導Ⅱ終了
//   N(13):訪問衛生開始  O(14):訪問衛生終了
//   P(15):衛生士開始   Q(16):衛生士終了
// ──────────────────────────────────────────────
function saveKiroku_(name, kana, kaigo, year, month, data) {
  if (!name) return { ok: false, error: '患者名が空です' };
  if (!data || !data.length) return { ok: false, error: 'データがありません' };
  if (!year || !month) return { ok: false, error: '年月が不正です' };

  var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
  var sheet = ss.getSheetByName('厚生局_訪問時間記録');
  if (!sheet) sheet = ss.insertSheet('厚生局_訪問時間記録');

  var HEADERS = [
    '登録日時', '患者氏名', 'カナ', '介護区分', '年', '月', '週',
    '訪問開始', '訪問終了',
    '指導Ⅰ開始', '指導Ⅰ終了',
    '指導Ⅱ開始', '指導Ⅱ終了',
    '訪問衛生開始', '訪問衛生終了',
    '衛生士開始', '衛生士終了'
  ];
  var COL = HEADERS.length; // 17

  // ヘッダー初期化（新規または空の場合のみ）
  if (!sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, COL).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  // 同一患者・同一年月の既存行を削除（上書き登録）
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existing = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = existing.length - 1; i >= 0; i--) {
      if (
        String(existing[i][1]).trim() === String(name).trim() &&
        String(existing[i][4]) === String(year) &&
        String(existing[i][5]) === String(month)
      ) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  // 新規行追加（週ごとに1行）
  var now  = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var rows = data.map(function(w) {
    return [
      now,
      name,
      kana || '',
      kaigo ? 'あり' : 'なし',
      year,
      month,
      w.week        || '',
      w['訪問開始']   || '',
      w['訪問終了']   || '',
      w['指導I開始']  || '',
      w['指導I終了']  || '',
      w['指導II開始'] || '',
      w['指導II終了'] || '',
      w['訪問衛生開始'] || '',
      w['訪問衛生終了'] || '',
      w['衛生士開始']   || '',
      w['衛生士終了']   || ''
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, COL).setValues(rows);
  SpreadsheetApp.flush();
  return { ok: true, message: name + ' の ' + rows.length + '週分を登録しました' };
}

// ──────────────────────────────────────────────
// 管理画面用：各シートへの直リンクを返す
// ──────────────────────────────────────────────
function getAdminLinks_() {
  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  var ssUrl = 'https://docs.google.com/spreadsheets/d/' + MASTER_SS_ID;

  function sheetUrl(name) {
    var s = ss.getSheetByName(name);
    return s ? ssUrl + '/edit#gid=' + s.getSheetId() : null;
  }

  return {
    ok: true,
    spreadsheetUrl: ssUrl + '/edit',
    sheets: {
      master:   { label: '患者マスターDB',       url: sheetUrl('マスター') },
      schedule: { label: 'スケジュール（月別訪問日）', url: sheetUrl('スケジュール') },
      kiroku:   { label: '厚生局_訪問時間記録',   url: sheetUrl('厚生局_訪問時間記録') }
    }
  };
}

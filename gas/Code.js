/** ============================================================
 *  HARIBOW照明キューエディタ — バックエンド
 *  データストア：単独公演2026 SS の「照明_演目一覧」「照明_キュー」タブ
 *  ============================================================ */

const CUE_SS_ID = '1qZPr5yUliMlEP-KQT8p31i5R8P5CbcRihkpS2bt3HqM';
const SHEET_SHOWS = '照明_演目一覧';
const SHEET_CUES = '照明_キュー';

const COL_SHOW = { NAME: 1, AUDIO_FILE: 2, DURATION: 3, BPM: 4, OWNER: 5, LAST_EDITOR: 6, LAST_UPDATED: 7, NOTE: 8 };
const COL_CUE = { SHOW: 1, NO: 2, TIME: 3, TIME_END: 4, PART: 5, FIXTURES: 6, COLOR: 7, INTENSITY: 8, MEMO: 9, MEASURED: 10, LAST_EDITOR: 11, LAST_UPDATED: 12, MUST: 13 };
// MUST(13列目)=「外せない演出」フラグ。2026-09-24追加。
// 全部を細かく指示するのではなく、要所だけを必ず守ってもらうための区別。

function cueSs_() { return SpreadsheetApp.openById(CUE_SS_ID); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === 'listShows') return json_({ ok: true, shows: listShows_() });
    if (action === 'getCues') return json_({ ok: true, show: e.parameter.show, cues: getCues_(e.parameter.show), meta: getShowMeta_(e.parameter.show) });
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === 'saveCues') return json_(saveCues_(payload.show, payload.cues, payload.editor, payload.baseVersion));
    if (payload.action === 'createShow') return json_(createShow_(payload.show, payload.audioFile, payload.duration, payload.bpm, payload.owner, payload.editor));
    if (payload.action === 'updateShow') return json_(updateShow_(payload.show, payload.audioFile, payload.duration, payload.bpm, payload.note, payload.editor));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** A列のラベル一致で行番号を探す(1-based)。見つからなければ0 */
function findRowByShow_(sheet, colIdx, show) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var data = sheet.getRange(2, colIdx, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === show) return i + 2;
  }
  return 0;
}

function listShows_() {
  var sheet = cueSs_().getSheetByName(SHEET_SHOWS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, COL_SHOW.NOTE).getValues();
  return data.filter(function (r) { return String(r[0]).trim() !== ''; }).map(function (r) {
    return {
      name: r[COL_SHOW.NAME - 1], audioFile: r[COL_SHOW.AUDIO_FILE - 1], duration: r[COL_SHOW.DURATION - 1],
      bpm: r[COL_SHOW.BPM - 1], owner: r[COL_SHOW.OWNER - 1], lastEditor: r[COL_SHOW.LAST_EDITOR - 1],
      lastUpdated: r[COL_SHOW.LAST_UPDATED - 1] ? String(r[COL_SHOW.LAST_UPDATED - 1]) : '', note: r[COL_SHOW.NOTE - 1]
    };
  });
}

function getShowMeta_(show) {
  var sheet = cueSs_().getSheetByName(SHEET_SHOWS);
  var rowNum = findRowByShow_(sheet, COL_SHOW.NAME, show);
  if (!rowNum) return null;
  var r = sheet.getRange(rowNum, 1, 1, COL_SHOW.NOTE).getValues()[0];
  return {
    name: r[COL_SHOW.NAME - 1], audioFile: r[COL_SHOW.AUDIO_FILE - 1], duration: r[COL_SHOW.DURATION - 1],
    bpm: r[COL_SHOW.BPM - 1], owner: r[COL_SHOW.OWNER - 1], lastEditor: r[COL_SHOW.LAST_EDITOR - 1],
    lastUpdated: r[COL_SHOW.LAST_UPDATED - 1] ? String(r[COL_SHOW.LAST_UPDATED - 1]) : '', note: r[COL_SHOW.NOTE - 1]
  };
}

function getCues_(show) {
  var sheet = cueSs_().getSheetByName(SHEET_CUES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, COL_CUE.MUST).getValues();
  return data.filter(function (r) { return String(r[COL_CUE.SHOW - 1]).trim() === show; })
    .map(function (r) {
      return {
        no: r[COL_CUE.NO - 1], time: Number(r[COL_CUE.TIME - 1]) || 0,
        timeEnd: r[COL_CUE.TIME_END - 1] === '' ? null : Number(r[COL_CUE.TIME_END - 1]),
        part: r[COL_CUE.PART - 1], fixtures: r[COL_CUE.FIXTURES - 1], color: r[COL_CUE.COLOR - 1],
        intensity: r[COL_CUE.INTENSITY - 1], memo: r[COL_CUE.MEMO - 1], measured: r[COL_CUE.MEASURED - 1] === true || String(r[COL_CUE.MEASURED - 1]).toUpperCase() === 'TRUE',
        must: r[COL_CUE.MUST - 1] === true || String(r[COL_CUE.MUST - 1]).toUpperCase() === 'TRUE'
      };
    })
    .sort(function (a, b) { return a.time - b.time; });
}

/** 演目を新規登録(演目一覧に1行追加)。既存なら何もしない */
function createShow_(show, audioFile, duration, bpm, owner, editor) {
  var sheet = cueSs_().getSheetByName(SHEET_SHOWS);
  var rowNum = findRowByShow_(sheet, COL_SHOW.NAME, show);
  if (rowNum) return { ok: false, error: 'show_exists' };
  var nextRow = sheet.getLastRow() + 1;
  var now = new Date();
  sheet.getRange(nextRow, 1, 1, COL_SHOW.NOTE).setValues([[show, audioFile || '', duration || 0, bpm || 0, owner || '', editor || '', now, '']]);
  return { ok: true, lastUpdated: String(now) };
}

/** 既存演目のメタ(音源ファイル/尺/BPM/備考)を更新。音源差替え時に使う。
 *  渡されなかった項目は現状維持する。 */
function updateShow_(show, audioFile, duration, bpm, note, editor) {
  var sheet = cueSs_().getSheetByName(SHEET_SHOWS);
  var rowNum = findRowByShow_(sheet, COL_SHOW.NAME, show);
  if (!rowNum) return { ok: false, error: 'show_not_found' };
  var before = sheet.getRange(rowNum, 1, 1, COL_SHOW.NOTE).getValues()[0];
  if (audioFile != null && audioFile !== '') sheet.getRange(rowNum, COL_SHOW.AUDIO_FILE).setValue(audioFile);
  if (duration != null && duration !== '') sheet.getRange(rowNum, COL_SHOW.DURATION).setValue(duration);
  if (bpm != null && bpm !== '') sheet.getRange(rowNum, COL_SHOW.BPM).setValue(bpm);
  if (note != null) sheet.getRange(rowNum, COL_SHOW.NOTE).setValue(note);
  var now = new Date();
  sheet.getRange(rowNum, COL_SHOW.LAST_EDITOR).setValue(editor || '');
  sheet.getRange(rowNum, COL_SHOW.LAST_UPDATED).setValue(now);
  return {
    ok: true, lastUpdated: String(now),
    before: { audioFile: before[COL_SHOW.AUDIO_FILE - 1], duration: before[COL_SHOW.DURATION - 1], bpm: before[COL_SHOW.BPM - 1] }
  };
}

/** 指定演目のキューを全置換。baseVersionが演目一覧の最終更新日時と一致しない場合は競合として拒否 */
function saveCues_(show, cues, editor, baseVersion) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var showSheet = cueSs_().getSheetByName(SHEET_SHOWS);
    var showRow = findRowByShow_(showSheet, COL_SHOW.NAME, show);
    if (!showRow) return { ok: false, error: 'show_not_found' };

    var currentVersion = String(showSheet.getRange(showRow, COL_SHOW.LAST_UPDATED).getValue());
    if (baseVersion && currentVersion && baseVersion !== currentVersion) {
      return { ok: false, error: 'conflict', currentVersion: currentVersion, currentEditor: showSheet.getRange(showRow, COL_SHOW.LAST_EDITOR).getValue() };
    }

    var cueSheet = cueSs_().getSheetByName(SHEET_CUES);
    // MUST列のヘッダーが未設定なら入れる(人がシートを直接見た時のため)
    if (!String(cueSheet.getRange(1, COL_CUE.MUST).getValue()).trim()) {
      cueSheet.getRange(1, COL_CUE.MUST).setValue('外せない');
    }
    var lastRow = cueSheet.getLastRow();
    // 既存の該当演目行を全削除(下から上へ)
    if (lastRow >= 2) {
      var data = cueSheet.getRange(2, COL_CUE.SHOW, lastRow - 1, 1).getValues();
      for (var i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]).trim() === show) cueSheet.deleteRow(i + 2);
      }
    }
    var now = new Date();
    var startRow = cueSheet.getLastRow() + 1;
    var rows = cues.map(function (c, idx) {
      return [show, idx + 1, c.time, c.timeEnd == null ? '' : c.timeEnd, c.part || '', c.fixtures || '', c.color || '', c.intensity || '', c.memo || '', !!c.measured, editor || '', now, !!c.must];
    });
    if (rows.length) cueSheet.getRange(startRow, 1, rows.length, COL_CUE.MUST).setValues(rows);

    showSheet.getRange(showRow, COL_SHOW.LAST_EDITOR).setValue(editor || '');
    showSheet.getRange(showRow, COL_SHOW.LAST_UPDATED).setValue(now);
    return { ok: true, lastUpdated: String(now), count: rows.length };
  } finally {
    lock.releaseLock();
  }
}

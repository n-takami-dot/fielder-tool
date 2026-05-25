const SS = SpreadsheetApp.getActiveSpreadsheet();
const GOALS = ['平台獲得','売場改善','多面展開','取り扱いアイテム増'];
const INFO_HEADERS = [
  '店舗名','弊社担当者','ゴール','現在のフェーズ',
  '担当者名','役職','権限レベル','関係性','キーマン',
  'お困りごと','打ち手','打ち手詳細',
  '本部制約','POP設置','プロモーション活動','最終訪問日'
];
const REC_HEADERS = [
  'ID','フィールダー','日付','店舗名',
  'ゴール','フェーズ','結果',
  'メモ','次回アクション','次回訪問日','優先度'
];

function doGet(e) {
  const action = e.parameter.action || 'getStores';
  if (action === 'getStores')    return getStores();
  if (action === 'getRecords')   return getRecords(e.parameter.fielder);
  if (action === 'getStoreInfo') return getStoreInfo();
  if (action === 'getMembers')        return getMembers();
  if (action === 'getAdminPassword') return getAdminPassword();
  return json({ error: 'unknown action' });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === 'saveRecord')    return saveRecord(data);
  if (data.action === 'saveStoreInfo') return saveStoreInfo(data);
  if (data.action === 'syncInfoSheet') return syncInfoSheet();
  if (data.action === 'deleteRecord')  return deleteRecord(data);
  return json({ error: 'unknown action' });
}

function getStores() {
  const sheet = SS.getSheetByName('店舗マスタ');
  const rows = sheet.getDataRange().getValues();
  const stores = rows.slice(1).filter(r => r[1]).map(r => ({
    chain:      String(r[0] || ''),
    name:       String(r[1] || ''),
    rank:       String(r[2] || ''),
    pic:        String(r[3] || ''),
    prevResult: String(r[4] || ''),
    m12: String(r[5] || ''),
    m1:  String(r[6] || ''),
    m2:  String(r[7] || ''),
    m3:  String(r[8] || ''),
    m4:  String(r[9] || ''),
    m5:  String(r[10] || ''),
    m6:  String(r[11] || ''),
    m7:  String(r[12] || ''),
    m8:  String(r[13] || ''),
    m9:  String(r[14] || ''),
    m10: String(r[15] || ''),
    m11: String(r[16] || ''),
    note: String(r[17] || '')
  }));
  return json({ stores });
}

function getRecords(fielder) {
  const sheet = SS.getSheetByName('訪問記録');
  const rows = sheet.getDataRange().getValues();
  const records = rows.slice(1)
    .filter(r => !fielder || r[1] === fielder)
    .map(r => ({
      id: r[0], fielder: r[1], date: r[2],
      storeName: r[3], goal: r[4], phase: r[5],
      result: r[6], memo: r[7],
      next: r[8], nextDate: r[9], priority: r[10]
    }));
  return json({ records });
}

function getStoreInfo() {
  const sheet = SS.getSheetByName('店舗情報');
  const rows = sheet.getDataRange().getValues();
  const info = rows.slice(1).filter(r => r[0]).map(r => ({
    storeName:    String(r[0]  || ''),
    fielder:      String(r[1]  || ''),
    goal:         r[2] !== '' ? r[2] : null,
    phase:        Number(r[3]  || 0),
    contactName:  String(r[4]  || ''),
    contactRole:  String(r[5]  || ''),
    contactAuth:  String(r[6]  || ''),
    contactRel:   String(r[7]  || ''),
    contactKM:    String(r[8]  || ''),
    issue1:       String(r[9]  || ''),
    uchiteSel:    String(r[10] || ''),
    uchiteDetail: String(r[11] || ''),
    hqConstraint: String(r[12] || ''),
    popOk:        String(r[13] || ''),
    promoOk:      String(r[14] || ''),
    lastVisit:    String(r[15] || '')
  }));
  return json({ info });
}

function getMembers() {
  const sheet = SS.getSheetByName('店舗マスタ');
  const rows = sheet.getDataRange().getValues();
  const members = rows.slice(1)
    .map(r => String(r[19] || ''))
    .filter(v => v !== '');
  return json({ members });
}

function getAdminPassword() {
  const sheet = SS.getSheetByName('店舗マスタ');
  // V1セル（行1・列22）にパスワードを格納
  const pw = String(sheet.getRange(1, 22).getValue() || '');
  return json({ pw });
}

function saveRecord(data) {
  const sheet = SS.getSheetByName('訪問記録');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(REC_HEADERS);
  }
  sheet.appendRow([
    data.id, data.fielder, data.date, data.storeName,
    data.goal, data.phase, data.result,
    data.memo, data.next, data.nextDate, data.priority
  ]);
  return json({ status: 'ok' });
}

function saveStoreInfo(data) {
  const sheet = SS.getSheetByName('店舗情報');
  ensureInfoHeaders(sheet);

  let goalText = '';
  if (data.goal !== null && data.goal !== undefined) {
    goalText = typeof data.goal === 'number'
      ? (GOALS[data.goal] || '')
      : String(data.goal);
  }

  const c = (data.contacts && data.contacts[0]) ? data.contacts[0] : {};
  const issues = data.issues || [];
  const issue1 = issues.length >= 1 ? issues[0].text : '';

  // 既存行のフィールダー名を取得
  const rows = sheet.getDataRange().getValues();
  const idx = rows.findIndex(r => r[0] === data.storeName);
  const existingFielder = idx > 0 ? String(rows[idx][1] || '') : '';

  // 送信値が空・「管理者」の場合は既存値を優先
  const fielder = (data.fielder && data.fielder !== '管理者')
    ? data.fielder
    : existingFielder;

  const row = [
    data.storeName,
    fielder,
    goalText,
    data.phase        || 0,
    c.name            || '',
    c.role            || '',
    c.auth            || '',
    c.rel             || '',
    c.keyman          || '',
    issue1,
    data.uchiteSel    || '',
    data.uchiteDetail || '',
    data.hqConstraint || '',
    data.popOk        || '',
    data.promoOk      || '',
    data.lastVisit    || ''
  ];

  if (idx > 0) {
    sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return json({ status: 'ok' });
}

function deleteRecord(data) {
  const sheet = SS.getSheetByName('訪問記録');
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      return json({ status: 'ok' });
    }
  }
  return json({ status: 'notfound' });
}

function initStoreInfo() {
  const masterSheet = SS.getSheetByName('店舗マスタ');
  const infoSheet   = SS.getSheetByName('店舗情報');
  ensureInfoHeaders(infoSheet);

  const masterRows = masterSheet.getDataRange().getValues();
  const storeNames = masterRows.slice(1).filter(r => r[1]).map(r => String(r[1]));

  const infoRows = infoSheet.getDataRange().getValues();
  const existingNames = infoRows.slice(1).map(r => String(r[0]));

  storeNames.forEach(name => {
    if (!existingNames.includes(name)) {
      const masterRow = masterRows.slice(1).find(r => String(r[1]) === name);
      const fielder = masterRow ? String(masterRow[3] || '') : '';
      infoSheet.appendRow([name, fielder, '', 0, '', '', '', '', '', '', '', '', '', '', '', '']);
    } else {
      const rowIdx = infoRows.findIndex(r => String(r[0]) === name);
      if (rowIdx > 0 && !infoRows[rowIdx][1]) {
        const masterRow = masterRows.slice(1).find(r => String(r[1]) === name);
        const fielder = masterRow ? String(masterRow[3] || '') : '';
        infoSheet.getRange(rowIdx + 1, 2).setValue(fielder);
      }
    }
  });
}

function syncInfoSheet() {
  initStoreInfo();
  return json({ status: 'ok' });
}

function ensureInfoHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(INFO_HEADERS);
  }
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

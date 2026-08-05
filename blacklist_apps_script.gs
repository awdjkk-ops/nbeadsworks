/**
 * 블랙리스트 관리용 Google Apps Script
 * - 이 코드를 구글 시트의 확장 프로그램 > Apps Script 에 붙여넣고
 *   "웹 앱"으로 배포(액세스 권한: 모두)하면 됩니다.
 * - 시트의 첫 번째 탭을 그대로 사용하며, 첫 실행 시 헤더 행을 자동으로 만듭니다.
 */

/**
 * 블랙리스트 관리용 Google Apps Script (JSONP 방식 — CORS 문제 없음)
 * - 이 코드를 구글 시트의 확장 프로그램 > Apps Script 에 붙여넣고
 *   "웹 앱"으로 배포(액세스 권한: 모두)하면 됩니다.
 * - 코드를 수정했다면 "배포 관리 > 연필 아이콘 > 새 버전"으로 다시 배포해야 반영됩니다.
 */

const SHEET_NAME = null; // null이면 첫 번째 시트를 사용

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', '이름', '주소', '휴대폰', '등록일시']);
  }
  return sheet;
}

function listItems_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  return values.slice(1).filter(r => r[0] !== '').map(r => ({
    id: String(r[0]),
    name: r[1],
    address: r[2],
    phone: r[3],
    createdAt: r[4]
  }));
}

function addItem_(name, address, phone) {
  const sheet = getSheet_();
  const id = String(Date.now()) + '-' + Math.floor(Math.random() * 1000);
  sheet.appendRow([id, name || '', address || '', phone || '', new Date()]);
  return { ok: true, id: id };
}

function deleteItem_(id) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'not found' };
}

function doGet(e) {
  const action = (e.parameter.action || 'list');
  const callback = e.parameter.callback;

  let result;
  try {
    if (action === 'list') {
      result = { ok: true, items: listItems_() };
    } else if (action === 'add') {
      result = addItem_(e.parameter.name, e.parameter.address, e.parameter.phone);
    } else if (action === 'delete') {
      result = deleteItem_(e.parameter.id);
    } else {
      result = { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

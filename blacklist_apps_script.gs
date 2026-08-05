/**
 * 블랙리스트 관리용 Google Apps Script
 * - 이 코드를 구글 시트의 확장 프로그램 > Apps Script 에 붙여넣고
 *   "웹 앱"으로 배포(액세스 권한: 모두)하면 됩니다.
 * - 시트의 첫 번째 탭을 그대로 사용하며, 첫 실행 시 헤더 행을 자동으로 만듭니다.
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

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(r => r[0] !== '').map(r => ({
    id: String(r[0]),
    name: r[1],
    address: r[2],
    phone: r[3],
    createdAt: r[4]
  }));
  return jsonOut_({ ok: true, items: rows });
}

function doPost(e) {
  const sheet = getSheet_();
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid body' });
  }

  const action = body.action;

  if (action === 'add') {
    const id = String(Date.now()) + '-' + Math.floor(Math.random() * 1000);
    const now = new Date();
    sheet.appendRow([id, body.name || '', body.address || '', body.phone || '', now]);
    return jsonOut_({ ok: true, id: id });
  }

  if (action === 'delete') {
    const id = String(body.id);
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === id) {
        sheet.deleteRow(i + 1);
        return jsonOut_({ ok: true });
      }
    }
    return jsonOut_({ ok: false, error: 'not found' });
  }

  return jsonOut_({ ok: false, error: 'unknown action' });
}

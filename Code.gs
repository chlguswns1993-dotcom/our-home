/**
 * ==============================================================================
 * [최현준 ♥ 조아영] 우리 집 공동생활 포털 & 가계부 Google Apps Script (Web App API)
 * ==============================================================================
 * 
 * 구글 드라이브(5TB) 내 스프레드시트를 클라우드 DB로 활용하여
 * 프론트엔드와 실시간 양방향 데이터 통신(GET / POST)을 수행합니다.
 * 
 * [시트 구성]
 * 1. Transactions: id, month, date, who, category, memo, amount, createdAt
 * 2. Tolls       : id, month, date, desc, amount, settled
 * 3. Checklist   : key, label, checked, updatedAt
 */

// 시트 이름 상수
const SHEET_TRANSACTIONS = 'Transactions';
const SHEET_TOLLS = 'Tolls';
const SHEET_CHECKLIST = 'Checklist';

/**
 * [1] 최초 설정 함수 (setupSheets)
 * 구글 스프레드시트 상단 메뉴 [실행] > [setupSheets] 선택하여 1회 실행합니다.
 * 3개 시트 생성, 헤더 서식 지정, 기본 가사 체크리스트 항목을 자동 세팅합니다.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Transactions 시트 설정
  let txSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
  if (!txSheet) {
    txSheet = ss.insertSheet(SHEET_TRANSACTIONS);
  }
  const txHeaders = ['id', 'month', 'date', 'who', 'category', 'memo', 'amount', 'createdAt'];
  txSheet.getRange(1, 1, 1, txHeaders.length).setValues([txHeaders]);
  formatHeaderRow(txSheet, txHeaders.length, '#4F46E5'); // 인디고 색상
  txSheet.setFrozenRows(1);

  // 2. Tolls 시트 설정
  let tollSheet = ss.getSheetByName(SHEET_TOLLS);
  if (!tollSheet) {
    tollSheet = ss.insertSheet(SHEET_TOLLS);
  }
  const tollHeaders = ['id', 'month', 'date', 'desc', 'amount', 'settled'];
  tollSheet.getRange(1, 1, 1, tollHeaders.length).setValues([tollHeaders]);
  formatHeaderRow(tollSheet, tollHeaders.length, '#E11D48'); // 로즈 색상
  tollSheet.setFrozenRows(1);

  // 3. Checklist 시트 설정
  let chkSheet = ss.getSheetByName(SHEET_CHECKLIST);
  if (!chkSheet) {
    chkSheet = ss.insertSheet(SHEET_CHECKLIST);
  }
  const chkHeaders = ['key', 'label', 'checked', 'updatedAt'];
  chkSheet.getRange(1, 1, 1, chkHeaders.length).setValues([chkHeaders]);
  formatHeaderRow(chkSheet, chkHeaders.length, '#0D9488'); // 틸 색상
  chkSheet.setFrozenRows(1);

  // 기본 가사 체크리스트 초기 데이터 주입 (비어있을 경우)
  if (chkSheet.getLastRow() <= 1) {
    const defaultChecklists = [
      ['chk_trash', '일요일 일반쓰레기 및 분리수거 배출 완료', false, new Date().toISOString()],
      ['chk_laundry', '일요일 기본 빨래(속옷·양말·면티·수건) & 면류 건조기 가동 + 건조 후 빨래 함께 개기', false, new Date().toISOString()],
      ['chk_vacuum', '주말 일반 진공청소기 1회 가동 (격주 교대 / 부재 시 2주 연속 후 교대)', false, new Date().toISOString()],
      ['chk_bath_master', '[2주 1회] 안방 화장실 청소 (최현준 전담)', false, new Date().toISOString()],
      ['chk_bath_living', '[2주 1회] 바깥 거실 화장실 청소 (조아영 전담)', false, new Date().toISOString()],
      ['chk_robot_deep', '[월 1회] 로봇청소기 센서·브러시·먼지통 종합 세척', false, new Date().toISOString()]
    ];
    chkSheet.getRange(2, 1, defaultChecklists.length, 4).setValues(defaultChecklists);
  }

  Logger.log('✅ 모든 시트와 헤더 설정이 성공적으로 완료되었습니다!');
}

/**
 * 헤더 행 스타일 서식 지정 헬퍼
 */
function formatHeaderRow(sheet, numCols, bgColor) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground(bgColor);
  range.setFontColor('#FFFFFF');
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, numCols);
}

/**
 * [2] GET 요청 처리 (doGet)
 * 웹앱 프론트엔드에서 데이터를 불러올 때 호출됩니다.
 * 파라미터:
 *  - action: 'getAll' (기본값) | 'checkHealth'
 *  - month: '2026-11' (특정 월 필터링 시 사용, 생략 시 전체 반환)
 */
function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = params.action || 'getAll';

    if (action === 'checkHealth') {
      return createJsonResponse({ status: 'ok', timestamp: new Date().toISOString() });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetMonth = params.month || '';

    // Transactions 조회
    const txSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
    const transactions = [];
    if (txSheet && txSheet.getLastRow() > 1) {
      const data = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, 8).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const rowMonth = String(row[1]);
        if (!targetMonth || rowMonth === targetMonth) {
          transactions.push({
            id: String(row[0]),
            month: rowMonth,
            date: formatDateString(row[2]),
            who: String(row[3]),
            category: String(row[4]),
            memo: String(row[5]),
            amount: Number(row[6]) || 0,
            createdAt: String(row[7])
          });
        }
      }
    }

    // Tolls 조회
    const tollSheet = ss.getSheetByName(SHEET_TOLLS);
    const tolls = [];
    if (tollSheet && tollSheet.getLastRow() > 1) {
      const data = tollSheet.getRange(2, 1, tollSheet.getLastRow() - 1, 6).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const rowMonth = String(row[1]);
        if (!targetMonth || rowMonth === targetMonth) {
          tolls.push({
            id: String(row[0]),
            month: rowMonth,
            date: formatDateString(row[2]),
            desc: String(row[3]),
            amount: Number(row[4]) || 0,
            settled: Boolean(row[5])
          });
        }
      }
    }

    // Checklist 조회
    const chkSheet = ss.getSheetByName(SHEET_CHECKLIST);
    const checklist = [];
    if (chkSheet && chkSheet.getLastRow() > 1) {
      const data = chkSheet.getRange(2, 1, chkSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        checklist.push({
          key: String(row[0]),
          label: String(row[1]),
          checked: Boolean(row[2]),
          updatedAt: String(row[3])
        });
      }
    }

    return createJsonResponse({
      success: true,
      month: targetMonth || 'all',
      transactions: transactions,
      tolls: tolls,
      checklist: checklist
    });

  } catch (error) {
    return createJsonResponse({
      success: false,
      error: error.toString()
    });
  }
}

/**
 * [3] POST 요청 처리 (doPost)
 * 웹앱 프론트엔드에서 데이터 추가/삭제/체크리스트 토글 시 호출됩니다.
 * CORS 프리플라이트 회피를 위해 브라우저에서는 text/plain 형식으로 JSON을 전송할 수 있습니다.
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    switch (action) {
      // 1. 거래 내역 추가
      case 'addTransaction': {
        const txSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
        if (!txSheet) throw new Error('Transactions 시트가 존재하지 않습니다.');
        
        const item = payload.data;
        const id = item.id || 'tx_' + new Date().getTime();
        const month = item.month || (item.date ? item.date.substring(0, 7) : '');
        const date = item.date || '';
        const who = item.who || '공동';
        const category = item.category || '기타';
        const memo = item.memo || '';
        const amount = Number(item.amount) || 0;
        const createdAt = item.createdAt || new Date().toISOString();

        txSheet.appendRow([id, month, date, who, category, memo, amount, createdAt]);
        return createJsonResponse({ success: true, action: action, id: id });
      }

      // 2. 거래 내역 삭제
      case 'deleteTransaction': {
        const txSheet = ss.getSheetByName(SHEET_TRANSACTIONS);
        const id = payload.id;
        const deleted = deleteRowById(txSheet, id, 1);
        return createJsonResponse({ success: deleted, action: action, id: id });
      }

      // 3. 톨비 내역 추가
      case 'addToll': {
        const tollSheet = ss.getSheetByName(SHEET_TOLLS);
        if (!tollSheet) throw new Error('Tolls 시트가 존재하지 않습니다.');

        const item = payload.data;
        const id = item.id || 'toll_' + new Date().getTime();
        const month = item.month || (item.date ? item.date.substring(0, 7) : '');
        const date = item.date || '';
        const desc = item.desc || '';
        const amount = Number(item.amount) || 0;
        const settled = Boolean(item.settled);

        tollSheet.appendRow([id, month, date, desc, amount, settled]);
        return createJsonResponse({ success: true, action: action, id: id });
      }

      // 4. 톨비 내역 삭제
      case 'deleteToll': {
        const tollSheet = ss.getSheetByName(SHEET_TOLLS);
        const id = payload.id;
        const deleted = deleteRowById(tollSheet, id, 1);
        return createJsonResponse({ success: deleted, action: action, id: id });
      }

      // 5. 톨비 정산 여부 토글 (settled)
      case 'toggleTollSettled': {
        const tollSheet = ss.getSheetByName(SHEET_TOLLS);
        const id = payload.id;
        const settled = Boolean(payload.settled);
        const updated = updateCellById(tollSheet, id, 1, 6, settled);
        return createJsonResponse({ success: updated, action: action, id: id, settled: settled });
      }

      // 6. 체크리스트 단일 항목 토글
      case 'toggleChecklist': {
        const chkSheet = ss.getSheetByName(SHEET_CHECKLIST);
        const key = payload.key;
        const checked = Boolean(payload.checked);
        const now = new Date().toISOString();

        let updated = false;
        if (chkSheet && chkSheet.getLastRow() > 1) {
          const values = chkSheet.getRange(2, 1, chkSheet.getLastRow() - 1, 1).getValues();
          for (let i = 0; i < values.length; i++) {
            if (String(values[i][0]) === String(key)) {
              chkSheet.getRange(i + 2, 3).setValue(checked);
              chkSheet.getRange(i + 2, 4).setValue(now);
              updated = true;
              break;
            }
          }
        }
        return createJsonResponse({ success: updated, action: action, key: key, checked: checked });
      }

      // 7. 체크리스트 일괄 초기화 (새로운 주 시작)
      case 'resetChecklist': {
        const chkSheet = ss.getSheetByName(SHEET_CHECKLIST);
        if (chkSheet && chkSheet.getLastRow() > 1) {
          const numRows = chkSheet.getLastRow() - 1;
          const now = new Date().toISOString();
          for (let i = 2; i <= chkSheet.getLastRow(); i++) {
            chkSheet.getRange(i, 3).setValue(false);
            chkSheet.getRange(i, 4).setValue(now);
          }
        }
        return createJsonResponse({ success: true, action: action });
      }

      // 8. 체크리스트 항목 추가
      case 'addChecklistItem': {
        const chkSheet = ss.getSheetByName(SHEET_CHECKLIST);
        const key = payload.key || 'chk_' + new Date().getTime();
        const label = payload.label || '새로운 가사 체크리스트';
        const checked = false;
        const now = new Date().toISOString();
        chkSheet.appendRow([key, label, checked, now]);
        return createJsonResponse({ success: true, action: action, key: key });
      }

      default:
        return createJsonResponse({ success: false, error: '알 수 없는 action: ' + action });
    }

  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * ID 기준으로 시트 행 삭제 유틸 함수
 */
function deleteRowById(sheet, id, idColIndex) {
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const values = sheet.getRange(2, idColIndex, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

/**
 * ID 기준으로 특정 열 셀 값 변경 유틸 함수
 */
function updateCellById(sheet, id, idColIndex, targetColIndex, value) {
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const values = sheet.getRange(2, idColIndex, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.getRange(i + 2, targetColIndex).setValue(value);
      return true;
    }
  }
  return false;
}

/**
 * 날짜 객체를 YYYY-MM-DD 문자열로 변환하는 유틸 함수
 */
function formatDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val);
}

/**
 * JSON 응답 생성 헬퍼 (CORS 헤더 및 MimeType 설정)
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

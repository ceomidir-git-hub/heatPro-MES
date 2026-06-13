/**
 * ════════════════════════════════════════════════════════════
 * HeatPro MES - 작업일보 Google Sheets 연동 Apps Script
 * ════════════════════════════════════════════════════════════
 * 사용법:
 * 1. "Github_Raw_Data___Worksheet_Db" 스프레드시트를 엽니다.
 * 2. 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 3. 기존 코드를 모두 지우고 이 파일의 내용을 전체 붙여넣기
 * 4. 저장 (Ctrl+S)
 * 5. 우측 상단 [배포] > [새 배포]
 *    - 유형 선택: 웹 앱
 *    - 실행 권한: 나 (Me)
 *    - 액세스 권한: 모든 사용자 (Anyone)
 * 6. [배포] 클릭 → 웹 앱 URL 복사
 * 7. GitHub_Worksheet.html 파일에서 API_URL 값을
 *    복사한 URL로 교체합니다.
 *
 * ※ 코드를 수정한 뒤에는 [배포] > [배포 관리] > 연필 아이콘 >
 *    "새 버전"으로 다시 배포해야 변경사항이 반영됩니다.
 * ════════════════════════════════════════════════════════════
 */

const DB_SHEET_NAME = 'DB DATA 2026';

// ── GET: RAW DATA 전체 불러오기 ────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const result = {
      employees: readColumn(ss, '직원코드', '이름'),
      customers: readColumn(ss, '고객사', '고객사'),
      materials: readColumn(ss, '재질', '재질'),
      equipment: readEquipment(ss),
      products:  readProducts(ss),
      dbCount:   getDbDataCount(ss)
    };

    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// ── POST: 작업일보 1건 저장 (DB DATA 2026 시트에 누적/수정) ──
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'save';

    if (action === 'save') {
      return jsonOut(saveRecord(ss, body.record));
    } else if (action === 'query') {
      return jsonOut(queryByDate(ss, body.date));
    }
    return jsonOut({ result: 'error', message: 'unknown action' });
  } catch (err) {
    return jsonOut({ result: 'error', message: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 헤더 행을 찾아서 특정 컬럼의 값 목록을 반환 ──────────────
function readColumn(ss, sheetName, headerName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i].map(c => String(c).trim());
    const col = row.indexOf(headerName);
    if (col >= 0) {
      const out = [];
      for (let r = i + 1; r < data.length; r++) {
        const v = String(data[r][col] || '').trim();
        if (v) out.push(v);
      }
      return out;
    }
  }
  return [];
}

// ── 제조설비번호: 설비번호 + 제조설비명 ─────────────────────
function readEquipment(ss) {
  const sheet = ss.getSheetByName('제조설비번호');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i].map(c => String(c).trim());
    const noCol = row.indexOf('설비번호');
    const nameCol = row.indexOf('제조설비명');
    if (noCol >= 0 && nameCol >= 0) {
      const out = [];
      for (let r = i + 1; r < data.length; r++) {
        const no = String(data[r][noCol] || '').trim();
        const name = String(data[r][nameCol] || '').trim();
        if (no && name) out.push({ no: no, name: name });
      }
      return out;
    }
  }
  return [];
}

// ── 품명품번: 고객사 > 품명 > [품번...] (forward-fill 병합셀 처리) ──
function readProducts(ss) {
  const sheet = ss.getSheetByName('품명품번');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i].map(c => String(c).trim());
    const custCol = row.indexOf('고객사');
    const nameCol = row.indexOf('품명');
    const noCol   = row.indexOf('품번');
    if (custCol >= 0 && nameCol >= 0 && noCol >= 0) {
      const result = {};
      let lastCust = '', lastName = '';
      for (let r = i + 1; r < data.length; r++) {
        const cust = String(data[r][custCol] || '').trim() || lastCust;
        const name = String(data[r][nameCol] || '').trim() || lastName;
        const no   = String(data[r][noCol]   || '').trim();
        if (!cust || !name || !no) continue;
        lastCust = cust; lastName = name;
        if (!result[cust]) result[cust] = {};
        if (!result[cust][name]) result[cust][name] = [];
        result[cust][name].push(no);
      }
      return result;
    }
  }
  return {};
}

function getDbDataCount(ss) {
  const sheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!sheet) return 0;
  const last = sheet.getLastRow();
  return last > 1 ? last - 1 : 0;
}

// ── 작업일보 저장: 5개 키(작업일/입고일/고객사/품명/품번) 일치 시 수정, 아니면 신규 행 추가 ──
function saveRecord(ss, record) {
  const sheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!sheet) return { result: 'error', message: 'DB DATA 2026 시트를 찾을 수 없습니다.' };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  // 날짜 비교용: 시트의 Date 객체 또는 문자열을 yyyy-MM-dd 문자열로 변환
  function cellToDateStr(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    const m = String(v).match(/(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[0] : String(v).trim();
  }

  const dateCol = headers.indexOf('작업일');
  const inspCol = headers.indexOf('입고일');
  const custCol = headers.indexOf('고객사');
  const nameCol = headers.indexOf('품명');
  const noCol   = headers.indexOf('품번');

  let targetRow = -1; // 1-indexed sheet row
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (let i = 0; i < range.length; i++) {
      const r = range[i];
      if (cellToDateStr(r[dateCol]) === (record['작업일'] || '') &&
          cellToDateStr(r[inspCol]) === (record['입고일'] || '') &&
          String(r[custCol] || '').trim() === (record['고객사'] || '') &&
          String(r[nameCol] || '').trim() === (record['품명'] || '') &&
          String(r[noCol]   || '').trim() === (record['품번'] || '')) {
        targetRow = i + 2; // sheet row number
        break;
      }
    }
  }

  const DATE_FIELDS = ['작업일', '입고일', '출고일'];

  // 헤더 순서에 맞춰 값 배열 구성
  const rowValues = headers.map((h, idx) => {
    if (h === 'No') {
      return targetRow > 0 ? sheet.getRange(targetRow, idx + 1).getValue() : (lastRow); // No는 신규 시 자동 계산
    }
    const v = record.hasOwnProperty(h) ? record[h] : '';
    if (v === '' || v === null || v === undefined) return '';
    if (DATE_FIELDS.indexOf(h) >= 0) {
      const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
      return v;
    }
    // 숫자형 필드 자동 변환
    const num = Number(String(v).replace(/,/g, ''));
    if (v !== '' && !isNaN(num) && /^-?\d+(\.\d+)?$/.test(String(v).trim())) return num;
    return v;
  });

  if (targetRow > 0) {
    // 수정 모드
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowValues]);
    return { result: 'updated', row: targetRow, no: rowValues[headers.indexOf('No')] };
  } else {
    // 신규 행 추가
    const newRowNo = lastRow > 1 ? lastRow : 1; // No 컬럼 자동 번호 (기존 데이터 행 수 기준)
    const noIdx = headers.indexOf('No');
    if (noIdx >= 0) rowValues[noIdx] = newRowNo;
    sheet.appendRow(rowValues);
    return { result: 'inserted', row: sheet.getLastRow(), no: newRowNo };
  }
}

// ── 작업일자로 DB 조회 ───────────────────────────────────────
function queryByDate(ss, dateStr) {
  const sheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!sheet) return { result: 'error', message: 'DB DATA 2026 시트를 찾을 수 없습니다.' };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { result: 'ok', rows: [] };

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const dateCol = headers.indexOf('작업일');
  const inspCol = headers.indexOf('입고일');
  const tagCol  = headers.indexOf('출고일');
  const DATE_COLS = [dateCol, inspCol, tagCol];

  function cellToDateStr(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    const m = String(v).match(/(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[0] : String(v).trim();
  }

  const rows = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (cellToDateStr(r[dateCol]) !== dateStr) continue;
    const obj = {};
    headers.forEach((h, c) => {
      let v = r[c];
      if (DATE_COLS.indexOf(c) >= 0) v = cellToDateStr(v);
      obj[h] = (v === null || v === undefined) ? '' : v;
    });
    rows.push(obj);
  }

  return { result: 'ok', rows: rows };
}

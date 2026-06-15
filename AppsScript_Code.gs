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

// 스프레드시트의 시간대를 안전하게 가져옴
// (getSpreadsheetTimeZone()이 실패하거나 비정상 값을 반환할 경우를 대비해
//  Session.getScriptTimeZone() → 'Asia/Seoul' 순으로 대체)
function getTz(ss) {
  try {
    var tz = ss.getSpreadsheetTimeZone();
    if (tz && typeof tz === 'string') return tz;
  } catch (e) {}
  try {
    var tz2 = Session.getScriptTimeZone();
    if (tz2 && typeof tz2 === 'string') return tz2;
  } catch (e) {}
  return 'Asia/Seoul';
}

// ════════════════════════════════════════════════════════════
// [임시 디버깅용] 날짜 값 진단 함수
// Apps Script 편집기에서 이 함수를 선택 → ▶실행 → 보기>실행 기록(로그)에서 결과 확인
// 확인 후에는 이 함수를 지워도 됩니다 (doGet/doPost와 무관, 안전함)
// ════════════════════════════════════════════════════════════
function debugDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = getTz(ss);
  Logger.log('getTz() 결과 (실제 코드가 쓰는 시간대): ' + tz);
  Logger.log('Spreadsheet TimeZone (raw, 빈 값일 수 있음): "' + ss.getSpreadsheetTimeZone() + '"');
  Logger.log('Script TimeZone: ' + Session.getScriptTimeZone());
  Logger.log('Spreadsheet Locale: ' + ss.getSpreadsheetLocale());

  const sheet = ss.getSheetByName(DB_SHEET_NAME);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  const dateCol = headers.indexOf('작업일');
  const lastRow = sheet.getLastRow();

  for (let r = lastRow - 4; r <= lastRow; r++) {
    const cell = sheet.getRange(r, dateCol + 1);
    const v = cell.getValue();
    const isDate = (Object.prototype.toString.call(v) === '[object Date]');
    Logger.log(
      'Row ' + r +
      ' | raw=' + v +
      ' | isDate=' + isDate +
      ' | ISO=' + (isDate ? v.toISOString() : 'N/A') +
      ' | fmt(getTz)=' + (isDate ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : 'N/A') +
      ' | cellNumberFormat=' + cell.getNumberFormat() +
      ' | displayValue=' + cell.getDisplayValue()
    );
  }
}

// ── GET: RAW DATA 전체 불러오기 + 대시보드 집계 ────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // index.html의 날짜 이동(◀▶) 버튼에서 전달하는 ?date=yyyy-MM-dd 파라미터
    // 값이 있으면 그 날짜를 "오늘"로 취급해서 대시보드를 계산함
    const selectedDate = (e && e.parameter && e.parameter.date) ? e.parameter.date : null;

    const result = {
      employees: readColumn(ss, '직원코드', '이름'),
      customers: readColumn(ss, '고객사', '고객사'),
      materials: readColumn(ss, '재질', '재질'),
      equipment: readEquipment(ss),
      products:  readProducts(ss),
      dbCount:   getDbDataCount(ss),
      dashboard: buildDashboard(ss, selectedDate)
    };

    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// ── 대시보드 집계: 선택한 날짜(또는 오늘) 기준 KPI + 월별/누적 차트 데이터 ─────────
// selectedDate: 'yyyy-MM-dd' 형식 문자열. 없거나 형식이 잘못되면 시스템 오늘 날짜 사용
function buildDashboard(ss, selectedDate) {
  var sheet = ss.getSheetByName(DB_SHEET_NAME);
  var empty = {
    todayJobCount: 0, todayProduction: 0, todayShipment: 0, todayEquipment: [],
    monthlyProduction: [], monthlyRate: [],
    cumulativeProduction: [], cumulativeRate: [], cumulativeDefectRate: []
  };
  if (!sheet) return empty;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return empty;

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // ⚠️ 스프레드시트의 시간대(셀에 표시되는 날짜 기준)를 사용해야
  //    "오늘작업건수" 등의 날짜 비교가 시트에 보이는 날짜와 정확히 일치합니다.
  //    (Session.getScriptTimeZone()은 Apps Script 프로젝트 설정 시간대로,
  //     스프레드시트 시간대와 다를 경우 날짜가 하루 어긋날 수 있습니다)
  var tz = getTz(ss);

  // 'date' 파라미터(yyyy-MM-dd)가 유효하면 그 날짜를 "오늘"로 사용, 그렇지 않으면 시스템 오늘 날짜 사용
  var todayStr;
  if (selectedDate && /^\d{4}-\d{2}-\d{2}$/.test(selectedDate)) {
    todayStr = selectedDate;
  } else {
    todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  }
  var thisYear = todayStr.substring(0, 4); // 월별/누적 차트도 선택한 날짜의 연도를 기준으로 집계

  function colIdx(name) { return headers.indexOf(name); }
  var dateCol   = colIdx('작업일');
  var shipCol   = colIdx('출고일');
  var inputCol  = colIdx('투입수량');
  var badCol    = colIdx('불량수량');
  var opCol     = colIdx('가동시간');
  var nonOpCol  = colIdx('비가동시간');
  var eqNameCol = colIdx('설비명');
  var eqNoCol   = colIdx('설비번호');

  function toDateStr(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    }
    var m = String(v).match(/(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[0] : String(v).trim();
  }

  function toNum(v) {
    var n = Number(String(v || '').replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }

  var todayJobCount = 0, todayProduction = 0, todayShipment = 0;
  var todayEquipMap = {};
  var monthlyMap = {};

  data.forEach(function(row) {
    var workDate = toDateStr(row[dateCol]);
    var shipDate = toDateStr(row[shipCol]);
    var inputQty = toNum(row[inputCol]);
    var badQty   = toNum(row[badCol]);
    var opTime   = toNum(row[opCol]);
    var nonOp    = toNum(row[nonOpCol]);

    if (workDate === todayStr) {
      todayJobCount++;
      todayProduction += inputQty;
      var eqName = String(row[eqNameCol] || '').trim();
      var eqNo   = String(row[eqNoCol]   || '').trim();
      if (eqName) todayEquipMap[eqName] = { name: eqName, no: eqNo };
    }
    if (shipDate === todayStr) todayShipment++;

    if (workDate && workDate.substring(0, 4) === thisYear) {
      var month = workDate.substring(0, 7);
      if (!monthlyMap[month]) monthlyMap[month] = { input: 0, bad: 0, opTime: 0, nonOp: 0 };
      monthlyMap[month].input  += inputQty;
      monthlyMap[month].bad    += badQty;
      monthlyMap[month].opTime += opTime;
      monthlyMap[month].nonOp  += nonOp;
    }
  });

  var months = Object.keys(monthlyMap).sort();

  var monthlyProduction = months.map(function(m) {
    return { month: m, value: monthlyMap[m].input };
  });
  var monthlyRate = months.map(function(m) {
    var d = monthlyMap[m];
    var total = d.opTime + d.nonOp;
    var rate  = total > 0 ? (d.opTime / total) * 100 : 0;
    return { month: m, value: Math.round(rate * 10) / 10 };
  });

  var cumInput = 0, cumBad = 0, rateSum = 0;
  var cumulativeProduction = [], cumulativeRate = [], cumulativeDefectRate = [];

  months.forEach(function(m, i) {
    var d = monthlyMap[m];
    cumInput += d.input;
    cumBad   += d.bad;
    rateSum  += monthlyRate[i].value;

    cumulativeProduction.push({ month: m, value: cumInput });
    cumulativeRate.push({ month: m, value: Math.round((rateSum / (i + 1)) * 10) / 10 });
    var defRate = cumInput > 0 ? (cumBad / cumInput) * 100 : 0;
    cumulativeDefectRate.push({ month: m, value: Math.round(defRate * 100) / 100 });
  });

  var todayEquipArr = Object.keys(todayEquipMap).map(function(k) { return todayEquipMap[k]; });

  return {
    todayJobCount:        todayJobCount,
    todayProduction:      todayProduction,
    todayShipment:        todayShipment,
    todayEquipment:       todayEquipArr,
    monthlyProduction:    monthlyProduction,
    monthlyRate:          monthlyRate,
    cumulativeProduction: cumulativeProduction,
    cumulativeRate:       cumulativeRate,
    cumulativeDefectRate: cumulativeDefectRate
  };
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
  // (스프레드시트 시간대 기준 — 셀에 표시되는 날짜와 일치시킴)
  const tz = getTz(ss);
  function cellToDateStr(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
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

  // DATE_FIELDS에 해당하는 컬럼 인덱스(0-based) — 저장 후 표시 형식 통일에 사용
  const dateColIdxs = DATE_FIELDS.map(h => headers.indexOf(h)).filter(i => i >= 0);

  function applyDateFormat(rowNum) {
    dateColIdxs.forEach(idx => {
      sheet.getRange(rowNum, idx + 1).setNumberFormat('yyyy-mm-dd');
    });
  }

  if (targetRow > 0) {
    // 수정 모드
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowValues]);
    applyDateFormat(targetRow);
    return { result: 'updated', row: targetRow, no: rowValues[headers.indexOf('No')] };
  } else {
    // 신규 행 추가
    const newRowNo = lastRow > 1 ? lastRow : 1; // No 컬럼 자동 번호 (기존 데이터 행 수 기준)
    const noIdx = headers.indexOf('No');
    if (noIdx >= 0) rowValues[noIdx] = newRowNo;
    sheet.appendRow(rowValues);
    applyDateFormat(sheet.getLastRow());
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

  // 스프레드시트 시간대 기준 — 셀에 표시되는 날짜와 일치시킴
  const tz = getTz(ss);
  function cellToDateStr(v) {
    if (!v) return '';
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
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

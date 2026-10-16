/**
 * 6학년 통합 관리 시스템 - 개요 페이지 전용 백엔드 (최적화 버전)
 */

// 상수 관리
const CONFIG = {
  IDS: {
    OVERVIEW: '1ZCGoSHlyBUs0IPxPr7zIMlVbUvKWc8cJlV3XxMXzDEE',
    BUDGET: '1E0b5Gw5zHEjAlewpNYDs0Fg0uEUr5fJJ4bskTrVL95E'
  },
  SHEET_NAMES: {
    OVERVIEW: '6학년 개요',
    BUDGET: '예산(학년)',
    EVENTS: '학교행사 및 정보' // [추가] 방학식 데이터를 가져올 시트 이름
  }
};

function doGet() {
  // 템플릿 생성
  var template = HtmlService.createTemplateFromFile('index');
  
  // [속도 개선 핵심] 
  // 페이지를 로드할 때 데이터를 미리 가져와서 템플릿에 심어줍니다.
  // 이렇게 하면 클라이언트에서 다시 요청할 필요가 없어 로딩이 매우 빨라집니다.
  try {
    template.initialJson = getInitialData();
  } catch (e) {
    template.initialJson = JSON.stringify({ error: e.toString() });
  }

  return template.evaluate()
    .setTitle('6학년 개요')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * [추가된 함수] 외부 시트(학교행사 및 정보)에서 방학식 날짜 추출
 */
function getVacationDates(ss) {
  try {
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EVENTS);
    if (!sheet) return { summer: '', winter: '' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { summer: '', winter: '' };

    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    let summerDate = '';
    let winterDate = '';

    for (let i = 0; i < data.length; i++) {
      let dateVal = data[i][0];
      let eventVal = String(data[i][2]);

      if (eventVal.indexOf("여름방학식") !== -1 && !summerDate) {
        summerDate = (dateVal instanceof Date) ? dateVal.toISOString() : dateVal;
      }
      if (eventVal.indexOf("겨울방학식") !== -1 && !winterDate) {
        winterDate = (dateVal instanceof Date) ? dateVal.toISOString() : dateVal;
      }
      if (summerDate && winterDate) break;
    }
    return { summer: summerDate, winter: winterDate };
  } catch (e) {
    Logger.log("방학식 로드 에러: " + e.message);
    return { summer: '', winter: '' };
  }
}

/**
 * 초기 데이터 로드 (최적화)
 * - 한 번의 호출로 필요한 모든 데이터를 병렬적(개념적)으로 수집하여 반환
 */
function getInitialData() {
  try {
    // 1. 개요 시트 데이터 가져오기
    const ssOverview = SpreadsheetApp.openById(CONFIG.IDS.OVERVIEW);
    const sheetOverview = ssOverview.getSheetByName(CONFIG.SHEET_NAMES.OVERVIEW);
    
    // 데이터 범위를 한 번에 가져와서 메모리에서 처리
    const lastRow = sheetOverview.getLastRow();
    const overviewAllData = sheetOverview.getRange(1, 1, lastRow, 15).getValues();

    // 2. 예산 시트 데이터 가져오기
    const ssBudget = SpreadsheetApp.openById(CONFIG.IDS.BUDGET);
    const sheetBudget = ssBudget.getSheetByName(CONFIG.SHEET_NAMES.BUDGET);
    const budgetRawData = sheetBudget.getRange('B1:K5').getValues();

    // 3. [추가] 방학식 데이터 가져오기
    const vacationDates = getVacationDates(ssOverview);

    return JSON.stringify({
      overview: processOverviewData(overviewAllData, vacationDates), // 방학식 데이터 같이 넘김
      budget: processBudgetData(budgetRawData)
    });

  } catch (e) {
    Logger.log(e);
    return JSON.stringify({ error: "Initial Load Failed: " + e.toString() });
  }
}

/**
 * 데이터 가공 - 개요 데이터
 */
function processOverviewData(dataValues, vacationDates) {
  let infoRows = [];
  
  for (let i = 1; i <= 5; i++) {
    let row = dataValues[i];
    if (!row) continue;

    let label = row[0];
    if (i === 5) label = '담임'; 

    infoRows.push({
      id: `r${i+1}`,
      label: label,
      c1: row[1], c2: row[2], c3: row[3], c4: row[4], c5: row[5],
      c6: row[6], c7: row[7], c8: row[8], c9: row[9], c10: row[10],
      total: row[11],
      isEditable: (i >= 4)
    });
  }

  // [수정됨] 기존에 overview 시트에서 직접 가져오던 코드를 삭제하고, 새로 찾은 데이터를 사용하도록 연결
  const dates = { 
    summer: vacationDates && vacationDates.summer ? vacationDates.summer : '', 
    winter: vacationDates && vacationDates.winter ? vacationDates.winter : '' 
  };

  let schedule = [];
  if (dataValues.length > 10) {
    for (let i = 10; i < dataValues.length; i++) {
      const r = dataValues[i];
      if (!r[0]) continue;

      schedule.push({
        rowIndex: i,
        date: r[0],
        event: r[2],  
        grade: r[11], 
        manual: r[14] 
      });
    }
  }

  return { info: infoRows, dates: dates, schedule: schedule };
}

/**
 * 데이터 가공 - 예산 데이터
 */
function processBudgetData(rawData) {
  const headers = rawData[0];
  const assigned = rawData[2];
  const used = rawData[3];
  const balance = rawData[4];

  let cards = [];
  
  for (let i = 0; i < headers.length; i++) {
    if (!headers[i]) continue;

    cards.push({
      name: headers[i],
      assigned: assigned[i],
      used: used[i],
      balance: balance[i]
    });
  }
  
  return cards;
}

// 데이터 새로고침
function refreshData() {
  return getInitialData();
}

/**
 * 데이터 업데이트 함수
 */
function updateSheetData(params) {
  const ss = SpreadsheetApp.openById(CONFIG.IDS.OVERVIEW);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.OVERVIEW);

  try {
    if (params.type === 'overview_edit') {
      const row = parseInt(params.rowId.replace('r','')); 
      const col = parseInt(params.colIndex) + 2; 
      sheet.getRange(row, col).setValue(params.value);
    }

    if (params.type === 'overview_schedule_edit') {
      const row = parseInt(params.rowIndex) + 1;
      const col = 15; 
      sheet.getRange(row, col).setValue(params.value);
    }

    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ error: e.toString() });
  }
}



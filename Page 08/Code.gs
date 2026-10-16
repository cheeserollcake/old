const SHEET_ID = '1Ey1Y89SldUwnoQtmdh8OAlV4_Z0sZuPM_jojv4_msuQ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('반별 일정 조회')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function formatDate(dateVal, timeZone) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, timeZone, "yyyy-MM-dd");
  }
  const str = String(dateVal).trim();
  if (str.length >= 10) return str.substring(0, 10);
  return str;
}

// 1. 데이터 조회 함수 (범위 수정 반영)
function getAllData(isRefresh) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "schedule_data_v7"; // 범위 변경으로 인한 캐시 키 업데이트
  
  if (!isRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      return JSON.parse(cachedData);
    }
  }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const baseSheet = ss.getSheetByName('기초자료');
    const addSheet = ss.getSheetByName('추가된일정');
    const timeZone = ss.getSpreadsheetTimeZone();
    
    // 빈 행 읽기 방지를 위한 LastRow 계산
    const baseLastRow = Math.max(baseSheet.getLastRow(), 2);
    const addLastRow = Math.max(addSheet.getLastRow(), 2);

    // [범위 수정 1] 휴일: 기초자료 B2:C
    const holidayData = baseSheet.getRange('B2:C' + baseLastRow).getValues();
    
    // [범위 수정 2] 강당: 기초자료 G2:J (날짜, 학급, 교시, 강좌명)
    const audData = baseSheet.getRange('G2:J' + baseLastRow).getValues();
    
    // [범위 수정 3] 전담: 기초자료 T2:W (날짜, 학급, 교시, 강좌명)
    const specData = baseSheet.getRange('T2:W' + baseLastRow).getValues();
    
    // [범위 수정 4] 추가된 일정: 추가된일정 B2:F (F열 포함)
    const addData = addSheet.getRange('B2:F' + addLastRow).getValues();

    // --- 데이터 매핑 ---
    
    // 휴일
    const holidays = holidayData
      .filter(r => r[0])
      .map(r => ({ date: formatDate(r[0], timeZone), name: String(r[1]) }));

    // 강당 (인덱스: 0=날짜, 1=학급, 2=교시, 3=강좌명)
    const auditorium = audData
      .filter(r => r[0])
      .map(r => ({
        date: formatDate(r[0], timeZone), 
        targetClass: String(r[1]),
        period: String(r[2]),
        course: String(r[3])
      }));

    // 전담 (인덱스: 0=날짜, 1=학급, 2=교시, 3=강좌명)
    const special = specData
      .filter(r => r[0])
      .map(r => ({
        date: formatDate(r[0], timeZone), 
        targetClass: String(r[1]),
        period: String(r[2]),
        course: String(r[3])
      }));

    const rawEvents = [];
    
    // 추가된 일정 (인덱스: 0=날짜, 1=학급, 2=교시, 3=강좌명. 4=F열은 무시)
    for (let i = 0; i < addData.length; i++) {
      if (!addData[i][0]) continue;
      rawEvents.push({
        date: formatDate(addData[i][0], timeZone), 
        classes: String(addData[i][1]),
        periods: String(addData[i][2]), 
        course: String(addData[i][3]).trim()
      });
    }

    // 강당/전담 일정을 rawEvents에 병합
    for (let i = 0; i < auditorium.length; i++) {
      const a = auditorium[i];
      if(a.targetClass) rawEvents.push({ date: a.date, classes: a.targetClass, periods: a.period, course: a.course });
    }
    for (let i = 0; i < special.length; i++) {
      const s = special[i];
      if(s.targetClass) rawEvents.push({ date: s.date, classes: s.targetClass, periods: s.period, course: s.course });
    }

    const result = { success: true, holidays: holidays, rawEvents: rawEvents };
    
    // 데이터 캐싱 (6시간)
    try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch (e) {}

    return result;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 2. 데이터 수정 (범위 수정 반영)
function modifyScheduleData(action, oldData, newData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const addSheet = ss.getSheetByName('추가된일정');
    const addLastRow = Math.max(addSheet.getLastRow(), 1);
    const timeZone = ss.getSpreadsheetTimeZone();
    
    // 수정 시 캐시 삭제
    CacheService.getScriptCache().remove("schedule_data_v7");

    if (action === 'add') {
      // 추가 시에는 B열부터 E열까지만 데이터 입력 (F열은 건드리지 않음)
      const bValues = addSheet.getRange('B1:B' + addLastRow).getValues();
      let lastRow = 1;
      for (let i = bValues.length - 1; i >= 0; i--) {
        if (bValues[i][0] !== '') { lastRow = i + 1; break; }
      }
      addSheet.getRange(lastRow + 1, 2, 1, 4).setValues([[newData.date, newData.classes, newData.periods, newData.course]]);
      return { success: true };
    }

    // 검색 시에는 B~F 열까지 포함하여 행을 찾음 (요청하신 대로 F열 포함)
    const values = addSheet.getRange('B1:F' + addLastRow).getValues();
    let targetRowIndex = -1;

    for (let i = values.length - 1; i >= 1; i--) { 
      if (!values[i][0]) continue;
      const rowDate = formatDate(values[i][0], timeZone);
      const rowClass = String(values[i][1]);
      const rowPeriod = String(values[i][2]);
      const rowCourse = String(values[i][3]).trim();

      if (rowDate === oldData.date && rowClass === oldData.classes && 
          rowPeriod === oldData.periods && rowCourse === oldData.course) {
        targetRowIndex = i + 1; 
        break;
      }
    }

    if (targetRowIndex === -1) return { success: false, error: "해당 일정을 원본 시트에서 찾을 수 없습니다." };

    if (action === 'delete') addSheet.deleteRow(targetRowIndex);
    else if (action === 'edit') addSheet.getRange(targetRowIndex, 2, 1, 4).setValues([[newData.date, newData.classes, newData.periods, newData.course]]);

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

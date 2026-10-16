/**
 * 2026학년도 6학년 전담 시간표 전용 백엔드 (Optimized for Speed)
 */

const FILE_ID = '1Ey1Y89SldUwnoQtmdh8OAlV4_Z0sZuPM_jojv4_msuQ';
// 캐시 키를 V2로 변경하여 기존 4개 탭 시절의 캐시를 무시하고 새로 시트를 읽게 만듭니다.
const CACHE_KEY = 'TIMETABLE_FULL_DATA_V2';
const CACHE_TIME = 21600; // 6시간 (초 단위)

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  
  // [속도 개선] 페이지 로드 시 데이터를 미리 가져와 HTML에 심어줍니다.
  // 클라이언트에서 별도로 로딩 요청을 보낼 필요가 없어집니다.
  let initialData = null;
  try {
    const cached = CacheService.getScriptCache().get(CACHE_KEY);
    if (cached) {
      initialData = cached; // 이미 JSON 문자열임
    } else {
      const dataObj = fetchFromSheet_();
      const responseObj = { success: true, data: dataObj };
      initialData = JSON.stringify(responseObj);
      // 데이터가 너무 크면 캐시에 안 담길 수 있으므로 try-catch 처리
      try {
        CacheService.getScriptCache().put(CACHE_KEY, initialData, CACHE_TIME);
      } catch (e) {
        console.log("Cache save failed (too large):", e);
      }
    }
  } catch (e) {
    initialData = JSON.stringify({ success: false, error: e.toString() });
  }

  template.initialData = initialData;

  return template.evaluate()
    .setTitle('전담 시간표')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * [내부 함수] 실제 시트에서 데이터 읽어오기
 */
function fetchFromSheet_() {
  const ss = SpreadsheetApp.openById(FILE_ID);
  const tabs = ['전담(과학1)', '전담(과학2)', '전담(영어1)', '전담(영어2)', '전담(영어3)'];
  const allData = {};

  tabs.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      const rangeData = sheet.getRange('A1:AF70').getValues();
      allData[sheetName] = processSheetDataInMemory_(rangeData);
    }
  });
  return allData;
}

/**
 * 프론트엔드에서 강제 리로드가 필요할 때 호출
 */
function getAllTimetableData() {
  try {
    // 1. 캐시 확인
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_KEY);
    
    if (cached) {
      return cached;
    }

    // 2. 캐시 없으면 시트 조회
    const data = fetchFromSheet_();
    const result = JSON.stringify({
      success: true,
      data: data
    });

    // 3. 캐시 저장
    try {
      cache.put(CACHE_KEY, result, CACHE_TIME);
    } catch(e) {}

    return result;
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 메모리 상의 2D 배열에서 필요한 데이터 추출
 */
function processSheetDataInMemory_(values) {
  const sem1Top = values.slice(2, 4).map(row => row.slice(1, 12));
  const sem1Bottom = values.slice(12, 33).map(row => row.slice(0, 32));
  const sem1Options = values[2].slice(2, 12);
  const sem1Title = values[9][0];

  const sem2Top = values.slice(38, 40).map(row => row.slice(1, 12));
  const sem2Bottom = values.slice(48, 69).map(row => row.slice(0, 32));
  const sem2Options = values[38].slice(2, 12);
  const sem2Title = values[45][0];

  return {
    sem1: { top: sem1Top, bottom: sem1Bottom, options: sem1Options, title: sem1Title },
    sem2: { top: sem2Top, bottom: sem2Bottom, options: sem2Options, title: sem2Title }
  };
}

/**
 * 데이터 일괄 수정
 */
function updateBatchSpecialTimetable(paramsList) {
  try {
    const ss = SpreadsheetApp.openById(FILE_ID);
    const updatesBySheet = {};
    
    paramsList.forEach(param => {
      if (!updatesBySheet[param.sheetName]) {
        updatesBySheet[param.sheetName] = [];
      }
      updatesBySheet[param.sheetName].push(param);
    });

    Object.keys(updatesBySheet).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const updates = updatesBySheet[sheetName];
      updates.forEach(u => {
        let row, col;
        if (u.semester === 1) {
          row = 13 + u.r;
          col = 1 + u.c;
        } else {
          row = 49 + u.r;
          col = 1 + u.c;
        }
        sheet.getRange(row, col).setValue(u.value);
      });
    });

    // [중요] 데이터가 수정되었으므로 캐시를 날려야 다음 조회 때 갱신된 데이터를 가져옴
    CacheService.getScriptCache().remove(CACHE_KEY);

    return JSON.stringify({ success: true, count: paramsList.length });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 단건 수정 (하위 호환성)
 */
function updateSpecialTimetable(params) {
  try {
    const ss = SpreadsheetApp.openById(FILE_ID);
    const sheet = ss.getSheetByName(params.sheetName);
    let row, col;
    if (params.semester === 1) {
      row = 13 + params.r;
      col = 1 + params.c;
    } else {
      row = 49 + params.r;
      col = 1 + params.c;
    }
    sheet.getRange(row, col).setValue(params.value);
    
    // 캐시 삭제
    CacheService.getScriptCache().remove(CACHE_KEY);
    
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

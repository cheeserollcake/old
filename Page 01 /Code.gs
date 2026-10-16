/**
 * 2026학년도 6학년 통합 관리 시스템 백엔드 (초정밀 예외 처리 & 대용량 캐시 방어 버전)
 */

const FILE_IDS = {
  OVERVIEW: '1ZCGoSHlyBUs0IPxPr7zIMlVbUvKWc8cJlV3XxMXzDEE'
};

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  
  let initialData = null;
  try {
    initialData = getInitialData();
  } catch (e) {
    initialData = JSON.stringify({ error: e.message });
  }
  
  if (!initialData) {
    initialData = JSON.stringify({});
  }

  template.initialData = initialData;

  return template.evaluate()
    .setTitle('2026학년도 6학년 통합 관리 시스템')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 초기 데이터 로드 (캐시 초과 방지 필터링 적용)
 */
function getInitialData() {
  const cache = CacheService.getScriptCache();
  let cachedData = null;
  
  try {
    cachedData = cache.get("mainOverviewData");
  } catch (cacheError) {
    console.warn("캐시 읽기 실패: " + cacheError.message);
  }

  if (cachedData != null) {
    return cachedData;
  }

  // 예외 처리 구조로 안전하게 데이터 취득
  let data;
  try {
    data = {
      overview: getOverviewData(),
      error: null
    };
  } catch (err) {
    data = {
      overview: { dates: { summer: null, winter: null }, schedule: [] },
      error: err.message
    };
  }
  
  const payload = JSON.stringify(data);
  
  // [핵심 패치] 캐시 용량(100KB) 제한 우회 보호막
  try {
    // UTF-8 바이트 단위 계산 대용으로 안전하게 90KB(약 90,000자) 이하일 때만 캐싱 적용
    if (payload.length < 90000) {
      cache.put("mainOverviewData", payload, 600); 
    } else {
      console.warn("데이터가 너무 커서 캐시를 생략하고 실시간 로드합니다. (크기: " + payload.length + " 자)");
    }
  } catch (cacheWriteError) {
    console.warn("캐시 쓰기 실패: " + cacheWriteError.message);
  }
  
  return payload;
}

/**
 * 수동 새로고침용 함수
 */
function refreshData() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove("mainOverviewData");
  } catch (e) {
    console.warn("캐시 삭제 실패: " + e.message);
  }
  return getInitialData();
}

/**
 * 스프레드시트 데이터 안전 추출
 */
function getOverviewData() {
  try {
    const ss = SpreadsheetApp.openById(FILE_IDS.OVERVIEW);
    if (!ss) {
      throw new Error("스프레드시트를 읽을 권한이 없거나 파일 ID가 유효하지 않습니다.");
    }
    
    // 1. 학교행사 및 정보 시트에서 방학식 날짜 찾기
    const eventSheet = ss.getSheetByName('학교행사 및 정보');
    let summerDate = null;
    let winterDate = null;
    
    if (eventSheet) {
      const lastEventRow = eventSheet.getLastRow();
      if (lastEventRow >= 2) {
        const eventData = eventSheet.getRange(2, 1, lastEventRow - 1, 3).getValues();
        for (let i = 0; i < eventData.length; i++) {
          const date = eventData[i][0];
          const eventName = String(eventData[i][2] || ""); 
          
          if (eventName.includes("여름방학식")) summerDate = date;
          if (eventName.includes("겨울방학식")) winterDate = date;
          
          if (summerDate && winterDate) break;
        }
      }
    }

    const dates = {
      summer: summerDate,
      winter: winterDate
    };

    // 2. 6학년 개요 시트에서 주요 일정 가져오기
    const sheet = ss.getSheetByName('6학년 개요');
    let schedule = [];

    if (!sheet) {
      throw new Error("'6학년 개요' 시트 이름을 찾을 수 없습니다.");
    }

    const lastRow = sheet.getLastRow();
    if (lastRow >= 11) {
      const rawSch = sheet.getRange(11, 1, lastRow - 10, 15).getValues();
      
      schedule = rawSch
        .map((r, idx) => ({
          rowIndex: idx,
          date: r[0],      
          event: r[2],     
          grade: r[11],    
          manual: r[14]    
        }))
        .filter(r => r.date); 
    }

    return {
      dates: dates,
      schedule: schedule
    };
  } catch (e) {
    throw new Error("스프레드시트 처리 중 장애 발생: " + e.toString());
  }
}

/**
 * 2026학년도 6학년 월별 시간표 전용 백엔드 (필터링 최적화)
 */

const DB_ID = '1Ey1Y89SldUwnoQtmdh8OAlV4_Z0sZuPM_jojv4_msuQ'; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('월별 시간표')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function changeMonthAndGetData(month) {
  try {
    // DB에서 데이터 직접 로드
    const data = fetchRawDataFromDB();
    
    // JSON 문자열로 변환하여 반환 (성공/실패 구조)
    return JSON.stringify({
      success: true,
      data: data,
      targetMonth: month
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 기초자료 시트에서 데이터를 가져옵니다.
 * [최적화 및 에러방지 핵심] getRange 사용 시 행/열 삭제로 인한 범위 오류를 원천 차단하기 위해
 * getDataRange()로 실제 존재하는 영역만 배열로 가져와 메모리에서 안전하게 처리합니다.
 */
function fetchRawDataFromDB() {
  const ss = SpreadsheetApp.openById(DB_ID);
  const sheetBase = ss.getSheetByName('기초자료');
  const sheetAdd = ss.getSheetByName('추가된일정');

  const holidays = [];
  const auditorium = [];
  const subject = [];
  let added = [];

  // 1. 기초자료 시트 데이터 (행이나 열이 삭제되어도 절대 다운되지 않는 방식)
  if (sheetBase) {
    const baseValues = sheetBase.getDataRange().getValues();
    
    // i=1 (2행)부터 시작하여 헤더 스킵
    for (let i = 1; i < baseValues.length; i++) {
      const row = baseValues[i];
      
      // B열(인덱스 1), C열(인덱스 2) - 휴일
      if (row.length > 1 && row[1]) {
        holidays.push([row[1], row[2]]);
      }
      
      // G열(6), H열(7), I열(8), J열(9) - 강당
      if (row.length > 9 && row[6] && row[9]) {
        auditorium.push([row[6], row[7], row[8], row[9]]);
      }
      
      // T열(19), U열(20), V열(21), W열(22) - 전담
      if (row.length > 22 && row[19] && row[22]) {
        subject.push([row[19], row[20], row[21], row[22]]);
      }
    }
  }

  // 2. 추가된일정 시트 데이터
  if (sheetAdd) {
    const addValues = sheetAdd.getDataRange().getValues();
    
    for (let i = 1; i < addValues.length; i++) {
      const row = addValues[i];
      
      // B열(1), C열(2), D열(3), E열(4) - 추가일정
      if (row.length > 1 && row[1] !== "" && row[1] != null) {
        added.push([
          row[1], 
          row.length > 2 ? row[2] : "", 
          row.length > 3 ? row[3] : "", 
          row.length > 4 ? row[4] : ""
        ]);
      }
    }
  }

  return {
    holidays: holidays,
    auditorium: auditorium,
    subject: subject,
    added: added
  };
}

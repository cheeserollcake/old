/**
 * 2026학년도 6학년 학급 기초 시간표 전용 백엔드 (Optimized Version v2)
 */

const FILE_IDS = {
  TIMETABLE: '10ecshZV4QwGeb9KJxqCM69wQP3p9euOKZa5N5etIFJg'
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('학급 기초 시간표')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  try {
    return JSON.stringify({ 
      success: true,
      type: 'basic_timetable', 
      data: getBasicTimetableData(1) 
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function refreshData(request) {
  try {
    if (request.classNum === 'all') {
       return JSON.stringify({ 
         success: true,
         type: 'basic_timetable_all', 
         data: getAllBasicTimetables() 
       });
    }
    const classNum = parseInt(request.classNum) || 1;
    return JSON.stringify({ 
      success: true,
      type: 'basic_timetable', 
      data: getBasicTimetableData(classNum) 
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

// --- Helper Functions (Optimized) ---

// 핵심 최적화: 전체 시트(getDataRange) 대신 필요한 데이터 영역(A1:AA50)만 가져옴
// 빈 행이 수천 개 있을 경우 getDataRange는 매우 느려짐
function getRawSheetData() {
  const ss = SpreadsheetApp.openById(FILE_IDS.TIMETABLE);
  const sheet = ss.getSheetByName('학급기초시간표');
  // 1반~10반 데이터가 대략 40행 내외에 있으므로 50행까지만 가져오면 충분함
  return sheet.getRange("A1:AA50").getValues();
}

function extractCommonData(allValues) {
  // 배열 인덱스는 0부터 시작하므로 엑셀 행번호 - 1
  // 과목: Q19:AA19 -> 행 18, 열 16~26
  // 시수: Q20:AA20 -> 행 19, 열 16~26
  // P12 -> 행 11, 열 15
  // P13 -> 행 12, 열 15
  
  // 데이터 범위 체크 후 추출
  const subjects = allValues[18] ? allValues[18].slice(16, 27) : [];
  const goals = allValues[19] ? allValues[19].slice(16, 27) : [];
  const p12 = allValues[11] ? allValues[11][15] : '';
  const p13 = allValues[12] ? allValues[12][15] : '';

  return { subjects, goals, p12, p13 };
}

function getScheduleFromRaw(allValues, classNum) {
  // 1,2반: 1행~ / 3,4반: 9행~ (8행 간격)
  // 홀수반: 1열(A, 인덱스0) / 짝수반: 8열(H, 인덱스7)
  const startRowIndex = (1 + (Math.ceil(classNum / 2) - 1) * 8) - 1;
  const startColIndex = classNum % 2 !== 0 ? 0 : 7;

  const scheduleData = [];
  for (let i = 0; i < 7; i++) {
    const rowData = [];
    if (allValues[startRowIndex + i]) {
      for (let j = 0; j < 6; j++) {
        rowData.push(allValues[startRowIndex + i][startColIndex + j]);
      }
    } else {
      for (let j = 0; j < 6; j++) rowData.push("");
    }
    scheduleData.push(rowData);
  }
  return scheduleData;
}

function getBasicTimetableData(classNum) {
  const allValues = getRawSheetData();
  const common = extractCommonData(allValues);
  const schedule = getScheduleFromRaw(allValues, classNum);

  return {
    subjects: common.subjects,
    goals: common.goals,
    p12: common.p12,
    p13: common.p13,
    schedule: schedule
  };
}

function getAllBasicTimetables() {
  const allValues = getRawSheetData();
  const common = extractCommonData(allValues);
  
  const allSchedules = [];
  for (let i = 1; i <= 10; i++) {
    allSchedules.push(getScheduleFromRaw(allValues, i));
  }

  return {
    subjects: common.subjects,
    goals: common.goals,
    p12: common.p12,
    p13: common.p13,
    allSchedules: allSchedules
  };
}

// --- Action Functions ---

function updateSheetData(params) {
  const ssTimetable = SpreadsheetApp.openById(FILE_IDS.TIMETABLE);

  try {
    if (params.type === 'basic_timetable_edit') {
      const sheet = ssTimetable.getSheetByName('학급기초시간표');
      const classNum = params.classNum;
      const startRow = 1 + (Math.ceil(classNum / 2) - 1) * 8;
      const startCol = classNum % 2 !== 0 ? 1 : 8;
      
      const targetRow = startRow + params.r;
      const targetCol = startCol + params.c;
      sheet.getRange(targetRow, targetCol).setValue(params.value);
    }
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

function deleteBasicTimetable(params) {
  const ss = SpreadsheetApp.openById(FILE_IDS.TIMETABLE);
  const sheet = ss.getSheetByName('학급기초시간표');
  const classNum = params.classNum;
  
  const startRow = 1 + (Math.ceil(classNum / 2) - 1) * 8;
  const startCol = classNum % 2 !== 0 ? 1 : 8;
  
  sheet.getRange(startRow + 1, startCol + 1, 6, 5).clearContent();
  
  return JSON.stringify({ success: true });
}

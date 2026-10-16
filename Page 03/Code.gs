/**
 * [통합 명렬표 시스템]
 * 학년 전체 명렬표 + 반별 명렬표 + 비밀번호/관리자 기능 통합
 */

// ---------------------------------------------------------------
// 1. 공통 설정 (스프레드시트 ID 및 비밀번호)
// ---------------------------------------------------------------

// 두 앱 모두 같은 시트를 사용하므로 ID 통합
const SPREADSHEET_ID = '1ZCGoSHlyBUs0IPxPr7zIMlVbUvKWc8cJlV3XxMXzDEE';

const DEFAULT_ACCESS_PW = '1111';
const DEFAULT_ADMIN_PW = '00admin';

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('통합 명렬표 시스템')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------
// 2. 인증 및 관리자 기능
// ---------------------------------------------------------------

function verifyPassword(password) {
  const props = PropertiesService.getScriptProperties();
  const savedPw = props.getProperty('ACCESS_PASSWORD') || DEFAULT_ACCESS_PW;
  
  if (password === savedPw) {
    return { success: true };
  } else {
    return { success: false, message: '비밀번호가 일치하지 않습니다.' };
  }
}

function verifyAdminPassword(password) {
  const props = PropertiesService.getScriptProperties();
  const savedAdminPw = props.getProperty('ADMIN_PASSWORD') || DEFAULT_ADMIN_PW;
  
  if (password === savedAdminPw) {
    return { success: true };
  } else {
    return { success: false, message: '관리자 비밀번호가 일치하지 않습니다.' };
  }
}

function changePasswords(adminPw, newAccessPw, newAdminPw) {
  const verify = verifyAdminPassword(adminPw);
  if (!verify.success) return verify;

  const props = PropertiesService.getScriptProperties();
  if (newAccessPw) props.setProperty('ACCESS_PASSWORD', newAccessPw);
  if (newAdminPw) props.setProperty('ADMIN_PASSWORD', newAdminPw);

  return { success: true, message: '비밀번호가 성공적으로 변경되었습니다.' };
}

// ---------------------------------------------------------------
// 3. 학년 전체 명렬표 로직
// ---------------------------------------------------------------

function getGradeData() {
  const data = {
    horizontal: getHorizontalNamelistData()
  };
  return JSON.stringify(data);
}

function refreshGradeData(request) {
  try {
    if (request.sub === '가로 명렬표') {
      return JSON.stringify({ type: 'horizontal_namelist', data: getHorizontalNamelistData() });
    }
    return getGradeData();
  } catch (e) {
    return JSON.stringify({ error: e.toString() });
  }
}

function getHorizontalNamelistData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetHoriz = ss.getSheetByName('가로명렬표');
  return {
    group1: sheetHoriz.getRange('C5:V34').getValues(),
    group2: sheetHoriz.getRange('C38:V67').getValues()
  };
}

// ---------------------------------------------------------------
// 4. 반별 명렬표 로직
// ---------------------------------------------------------------

/**
 * 특정 반의 명렬표 데이터 가져오기
 */
function getClassData(classNum) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(`${classNum}반`);
    if (!sheet) return JSON.stringify({ success: false, error: "시트를 찾을 수 없습니다." });
    
    // B1:F31 (헤더 포함 전체 범위)
    const data = sheet.getRange('B1:F31').getValues();
    return JSON.stringify({ success: true, data: data });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 학생 정보 수정
 */
function updateStudent(params) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(`${params.classNum}반`);
    
    // 데이터는 1행(헤더) 다음인 2행부터 시작하지만, 
    // 프론트엔드 인덱스(rowIndex)는 헤더(0)를 포함한 배열의 인덱스이므로
    // 시트의 실제 행 번호 = params.rowIndex + 1 (B1 시작 기준)
    const row = params.rowIndex + 1; 
    let col = 0;
    
    // B(2), C(3), D(4), E(5), F(6)
    if (params.field === 'name') col = 3;   
    if (params.field === 'gender') col = 4; 
    if (params.field === 'note') col = 5;   
    if (params.field === 'birth') col = 6;  
    
    if (col > 0) {
      sheet.getRange(row, col).setValue(params.value);
    }
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

/**
 * 학생 정보 삭제 (행 내용 지우기)
 */
function deleteStudent(params) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(`${params.classNum}반`);
    const row = params.rowIndex + 1;
    // C, D, E, F 열 삭제 (3,4,5,6) - 번호(B)는 유지
    sheet.getRange(row, 3, 1, 4).clearContent();
    return JSON.stringify({ success: true });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

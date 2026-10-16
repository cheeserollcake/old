// 권한 설정: 전체 시트 접근 허용
// appsscript.json에서 oauthScopes에 https://www.googleapis.com/auth/spreadsheets 가 포함되어야 함

const SPREADSHEET_ID = '1E0b5Gw5zHEjAlewpNYDs0Fg0uEUr5fJJ4bskTrVL95E'; 
const SHEET_NAME = '학교ID';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('쇼핑몰 학교 ID')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 설정(비밀번호)을 구글 서버 속성에 안전하게 저장하고 불러오는 함수 추가
 */
function getAppConfig() {
  const saved = PropertiesService.getScriptProperties().getProperty('schoolIdAppConfig');
  if (saved) {
    return JSON.parse(saved);
  }
  // 저장된 설정이 없으면 기본 비밀번호 반환
  return { userPw: "1111", adminPw: "0000" }; 
}

function saveAppConfig(config) {
  PropertiesService.getScriptProperties().setProperty('schoolIdAppConfig', JSON.stringify(config));
  return true;
}

/**
 * 데이터 가져오기
 * 수정/삭제를 위해 '행 번호(rowIndex)'를 함께 반환하도록 변경
 */
function getData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const range = sheet.getRange('A6:E50');
  const values = range.getValues();
  
  // 데이터와 함께 실제 행 번호(rowIndex)를 객체로 매핑
  // A6가 시작이므로 인덱스 i + 6이 실제 행 번호
  const mappedData = values.map((row, i) => ({
    rowIndex: i + 6,
    mall: row[0],
    id: row[1],
    pw: row[2],
    url: row[3],
    note: row[4]
  }));

  // 빈 행(쇼핑몰 이름이 없는 행) 필터링
  return mappedData.filter(item => item.mall !== "" || item.id !== "");
}

/**
 * 데이터 추가하기
 */
function addData(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  // 빈 행 찾기 로직
  const range = sheet.getRange('A6:E50');
  const values = range.getValues();
  let emptyRowIndex = -1;

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "") {
      emptyRowIndex = i;
      break;
    }
  }

  if (emptyRowIndex === -1) {
    return { success: false, message: "데이터 입력 공간(50행)이 가득 찼습니다." };
  }

  const targetRow = emptyRowIndex + 6;
  return saveRowData(sheet, targetRow, formObject);
}

/**
 * 데이터 수정하기 (New)
 */
function updateData(rowIndex, formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  // rowIndex 유효성 검사 (A6:E50 범위 내인지)
  if (rowIndex < 6 || rowIndex > 50) {
    return { success: false, message: "유효하지 않은 행 번호입니다." };
  }
  
  return saveRowData(sheet, rowIndex, formObject);
}

/**
 * 데이터 삭제하기 (New)
 */
function deleteData(rowIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (rowIndex < 6 || rowIndex > 50) {
    return { success: false, message: "유효하지 않은 행 번호입니다." };
  }

  // 행 삭제 대신 내용을 지우고 정렬하거나, 그냥 행을 삭제할 수 있음.
  // 여기서는 행 자체를 삭제하면 아래 데이터가 올라오므로 deleteRow 사용
  // 주의: deleteRow를 하면 50행 범위가 하나 줄어들 수 있으므로, 
  // 50행 고정 포맷이라면 clearContent()가 나을 수 있으나 
  // 일반적으로는 deleteRow가 깔끔함. 여기서는 clearContent 후 정렬 로직이 복잡하므로
  // 단순하게 해당 행의 내용을 지우는(Clear) 방식을 택함 (빈 행으로 만듦)
  // 이렇게 해야 A6:E50 구조가 깨지지 않음.
  
  sheet.getRange(rowIndex, 1, 1, 5).clearContent();
  
  return { success: true, message: "삭제되었습니다." };
}

/**
 * 공통 저장 로직 (추가/수정 시 사용)
 */
function saveRowData(sheet, targetRow, formObject) {
  let siteUrl = formObject.siteUrl;
  if (siteUrl && !siteUrl.match(/^https?:\/\//i)) {
    siteUrl = 'http://' + siteUrl;
  }

  // 텍스트 데이터 저장
  sheet.getRange(targetRow, 1).setValue(formObject.mallName);
  sheet.getRange(targetRow, 2).setValue(formObject.mallId);
  sheet.getRange(targetRow, 3).setValue(formObject.mallPw);
  sheet.getRange(targetRow, 5).setValue(formObject.note);

  // URL RichText 저장
  if (siteUrl) {
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(siteUrl)
      .setLinkUrl(siteUrl)
      .build();
    sheet.getRange(targetRow, 4).setRichTextValue(richText);
  } else {
    sheet.getRange(targetRow, 4).setValue("");
  }

  return { success: true, message: "저장되었습니다." };
}

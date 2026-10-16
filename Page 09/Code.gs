/**
 * 스프레드시트 ID
 * [수정] 사용자가 제공한 올바른 ID로 변경됨
 */
const SPREADSHEET_ID = '1Ey1Y89SldUwnoQtmdh8OAlV4_Z0sZuPM_jojv4_msuQ';
const SHEET_NAME = '추가된일정';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('강좌별 시간표 조회')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 안전한 데이터 로딩 함수
 */
function getAllData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`'${SHEET_NAME}' 시트를 찾을 수 없습니다.`);
    }

    const lastRow = sheet.getLastRow();
    const maxRows = sheet.getMaxRows();
    
    // --- 1. 드롭다운 데이터 가져오기 (G2:G100) ---
    let specialKeywords = { all: '', partial1: '', partial2: '', partial3: '', partial4: '', partial5: '', partial6: '', partial7: '' };
    const dropdowns = []; // 추가됨: dropdowns 배열 초기화

    const endRow = Math.min(100, maxRows);

    if (endRow >= 2) {
      const rangeA1 = `G2:G${endRow}`;
      const dropdownRange = sheet.getRange(rangeA1);
      const dropdownValues = dropdownRange.getDisplayValues(); 

      specialKeywords.all = dropdownValues[0] ? dropdownValues[0][0] : '';      
      specialKeywords.partial1 = dropdownValues[1] ? dropdownValues[1][0] : ''; 
      specialKeywords.partial2 = dropdownValues[2] ? dropdownValues[2][0] : ''; 
      specialKeywords.partial3 = dropdownValues[3] ? dropdownValues[3][0] : ''; 
      specialKeywords.partial4 = dropdownValues[4] ? dropdownValues[4][0] : ''; 
      specialKeywords.partial5 = dropdownValues[5] ? dropdownValues[5][0] : ''; 
      specialKeywords.partial6 = dropdownValues[6] ? dropdownValues[6][0] : ''; 
      specialKeywords.partial7 = dropdownValues[7] ? dropdownValues[7][0] : ''; 

      dropdownValues.forEach(row => {
        const val = row[0];
        if (val && String(val).trim() !== '') {
          dropdowns.push(val);
        }
      });
    }

    // --- 2. 일정 데이터 가져오기 (B2:E) ---
    const events = [];
    if (lastRow >= 2) {
      const eventRange = sheet.getRange(2, 2, lastRow - 1, 4);
      const eventValues = eventRange.getValues(); 

      for (let i = 0; i < eventValues.length; i++) {
        const row = eventValues[i];
        
        if (!row[0]) continue;

        let dateStr = "";
        if (row[0] instanceof Date) {
          dateStr = Utilities.formatDate(row[0], 'Asia/Seoul', 'yyyy-MM-dd');
        } else {
          try {
            const parsedDate = new Date(row[0]);
            if (!isNaN(parsedDate.getTime())) {
              dateStr = Utilities.formatDate(parsedDate, 'Asia/Seoul', 'yyyy-MM-dd');
            }
          } catch (e) { continue; }
        }

        if (dateStr) {
          events.push({
            date: dateStr,
            classNm: String(row[1] || ''), 
            period: String(row[2] || ''),  
            course: String(row[3] || '')   
          });
        }
      }
    }

    // --- 3. 휴일 데이터 가져오기 ('기초자료' 시트 B2:C) ---
    const basicSheet = ss.getSheetByName('기초자료');
    let holidays = {};
    if (basicSheet) {
      const basicLastRow = basicSheet.getLastRow();
      if (basicLastRow >= 2) {
        const holidayRange = basicSheet.getRange(2, 2, basicLastRow - 1, 2);
        const holidayValues = holidayRange.getValues();
        
        for (let i = 0; i < holidayValues.length; i++) {
          const row = holidayValues[i];
          if (!row[0]) continue;
          
          let dateStr = "";
          if (row[0] instanceof Date) {
            dateStr = Utilities.formatDate(row[0], 'Asia/Seoul', 'yyyy-MM-dd');
          } else {
            try {
              const parsedDate = new Date(row[0]);
              if (!isNaN(parsedDate.getTime())) {
                dateStr = Utilities.formatDate(parsedDate, 'Asia/Seoul', 'yyyy-MM-dd');
              }
            } catch (e) { continue; }
          }
          
          if (dateStr && row[1]) {
            holidays[dateStr] = String(row[1]);
          }
        }
      }
    }

    return {
      success: true,
      dropdowns: dropdowns,
      specials: specialKeywords,
      events: events,
      holidays: holidays
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

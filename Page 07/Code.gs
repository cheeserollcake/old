const SPREADSHEET_ID = '1Ey1Y89SldUwnoQtmdh8OAlV4_Z0sZuPM_jojv4_msuQ';
const DRIVE_FOLDER_ID = '1CcQG3KcA9lXtkGOVt8oAWATsUTpONRPm';
const DATA_SHEET_NAME = '기초자료';
const KEY_SETTINGS = 'APP_SETTINGS';
const KEY_PASSWORDS = 'ROOM_PASSWORDS'; // 특별실 비밀번호 저장 키

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('특별실 시간표')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET_NAME);
    const lastRow = Math.max(sheet.getLastRow(), 2);

    const holidays = [];
    const schedule = [];
    let holidayValues = [];
    let sportsValues = [];
    let comValues = [];

    if (lastRow >= 2) {
      holidayValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // 휴일 B열
      sportsValues = sheet.getRange(2, 7, lastRow - 1, 4).getValues();  // 체육공간 G:J열
      comValues = sheet.getRange(2, 27, lastRow - 1, 4).getValues();    // 컴실등 AA:AD열
    }

    holidayValues.forEach(row => {
      if (row[0]) holidays.push(formatToAppDate(row[0]));
    });

    const processSchedule = (values) => {
      values.forEach(row => {
        if (row[0] && row[1]) { 
          schedule.push({
            date: formatToAppDate(row[0]),
            className: row[1],
            period: String(row[2]).trim(),
            space: String(row[3]).trim()
          });
        }
      });
    };
    processSchedule(sportsValues);
    processSchedule(comValues);

    const scriptProps = PropertiesService.getScriptProperties();
    const settingsJson = scriptProps.getProperty(KEY_SETTINGS);
    const passwordsJson = scriptProps.getProperty(KEY_PASSWORDS);

    let settings = settingsJson ? JSON.parse(settingsJson) : { 
      adminPwd: '0000', 
      specialPwd: '1111', 
      sem1: {}, 
      sem2: {}, 
      exceptions: [],
      closeExceptions: [] // 예외 불가 설정 추가
    };

    let roomPasswords = passwordsJson ? JSON.parse(passwordsJson) : [];

    return {
      success: true,
      holidays: [...new Set(holidays)],
      schedule: schedule,
      settings: settings,
      roomPasswords: roomPasswords
    };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function saveScheduleData(updates) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false, error: "서버가 혼잡합니다. 잠시 후 시도해주세요." };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(DATA_SHEET_NAME);
    const lastRow = Math.max(sheet.getLastRow(), 2);

    // VR 열 추가 및 다목 분리 반영
    const sportsSpaces = ['강당', '다목(2층)', '다목(3층)', '운동장', '나래', 'VR'];
    const comSpaces = ['컴1', '컴2', '음악'];

    const sportsUpdates = updates.filter(u => sportsSpaces.includes(u.space));
    const comUpdates = updates.filter(u => comSpaces.includes(u.space));

    updateSheetRange(sheet, sportsUpdates, 7, lastRow); // G열(7)
    updateSheetRange(sheet, comUpdates, 27, lastRow);   // AA열(27)

    return { success: true };

  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateSheetRange(sheet, updates, startCol, lastRow) {
  if (updates.length === 0) return false;
  let values = [];
  if (lastRow >= 2) {
    values = sheet.getRange(2, startCol, lastRow - 1, 4).getValues();
  }

  let isChanged = false;
  updates.forEach(update => {
    const appDateStr = update.date;
    const targetPeriodStr = String(update.period).trim();
    const targetSpace = String(update.space).trim();
    const newValue = update.className;
    let foundIndex = -1;

    for (let i = 0; i < values.length; i++) {
      const sheetDateInAppFormat = formatToAppDate(values[i][0]);
      const sheetPeriodStr = String(values[i][2]).trim();
      const sheetSpaceStr = String(values[i][3]).trim();
      if (sheetDateInAppFormat === appDateStr && 
          sheetPeriodStr === targetPeriodStr && 
          sheetSpaceStr === targetSpace) {
        foundIndex = i;
        break;
      }
    }

    // "불가"가 입력되면 지워지지 않고 일반 데이터처럼 입력되도록 조건문 수정
    if (newValue === "") {
      if (foundIndex !== -1) {
        values[foundIndex] = ["", "", "", ""];
        isChanged = true;
      }
    } else {
      if (foundIndex !== -1) {
        if (values[foundIndex][1] !== newValue) {
          values[foundIndex][1] = newValue;
          values[foundIndex][2] = targetPeriodStr;
          isChanged = true;
        }
      } else {
        let emptyIndex = -1;
        for (let i = 0; i < values.length; i++) {
          if (!values[i][0] && !values[i][1]) {
            emptyIndex = i;
            break;
          }
        }
        const sheetDateStr = formatToSheetDate(appDateStr);
        const newRow = [sheetDateStr, newValue, targetPeriodStr, targetSpace];
        if (emptyIndex !== -1) values[emptyIndex] = newRow;
        else values.push(newRow);
        isChanged = true;
      }
    }
  });

  if (isChanged && values.length > 0) {
    let endRow = values.length;
    while (endRow > 0 && !values[endRow - 1][0]) { endRow--; }
    const finalValues = values.slice(0, Math.max(endRow, 1));
    sheet.getRange(2, startCol, finalValues.length, 4).setValues(finalValues);
    if (lastRow - 1 > finalValues.length) {
      sheet.getRange(2 + finalValues.length, startCol, (lastRow - 1) - finalValues.length, 4).clearContent();
    }
  } else if (values.length === 0 && lastRow >= 2) {
    sheet.getRange(2, startCol, lastRow - 1, 4).clearContent();
  }
  return isChanged;
}

function saveSettings(settings) {
  try {
    PropertiesService.getScriptProperties().setProperty(KEY_SETTINGS, JSON.stringify(settings));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function saveRoomPasswords(passwords) {
  try {
    PropertiesService.getScriptProperties().setProperty(KEY_PASSWORDS, JSON.stringify(passwords));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getImages() {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const files = folder.getFiles();
    const result = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType().includes('image') && !file.isTrashed()) {
        result.push({
          id: file.getId(),
          url: "https://lh3.googleusercontent.com/d/" + file.getId(),
          name: file.getName()
        });
      }
    }
    return result;
  } catch (e) { return []; }
}

function uploadImage(data, name, mimeType) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, name);
    folder.createFile(blob);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteImage(id) {
  try {
    DriveApp.getFileById(id).setTrashed(true);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function formatToAppDate(val) {
  if (!val) return "";
  let date;
  if (typeof val === 'string') {
    if (val.includes('.')) {
      const parts = val.split('.').map(s => s.trim());
      if (parts.length >= 3) date = new Date(parts[0], parts[1] - 1, parts[2]);
      else return val;
    } else return val;
  } else date = val;
  if (isNaN(date.getTime())) return "";
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  return `${date.getFullYear()}년 ${String(date.getMonth() + 1).padStart(2, '0')}월 ${String(date.getDate()).padStart(2, '0')}일 (${days[date.getDay()]})`;
}

function formatToSheetDate(appDateStr) {
  if (!appDateStr) return "";
  const match = appDateStr.match(/(\d{4})년\s(\d{2})월\s(\d{2})일/);
  if (match) {
    return `${match[1]}. ${parseInt(match[2], 10)}. ${parseInt(match[3], 10)}`;
  }
  return appDateStr;
}

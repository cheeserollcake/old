/**
 * 웹앱으로 접속했을 때 화면(index.html)을 렌더링합니다.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('내선번호')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=5, user-scalable=yes');
}

/**
 * [중요] 이 함수를 구글 앱스 스크립트 에디터에서 선택하고 [실행] 버튼을 누르면
 * 구글 드라이브(DriveApp) 권한 승인창이 확실하게 강제로 표시됩니다.
 */
function forceAuthorize() {
  const root = DriveApp.getRootFolder();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('AUTH_TEST', 'SUCCESS');
  Logger.log("구글 드라이브 및 속성 스토리지 권한 승인이 완료되었습니다: " + root.getName());
  return "권한 승인 완료!";
}

/**
 * 현재 저장된 파일의 ID, 파일명, MimeType 정보를 안전하게 분석하여 반환합니다.
 */
function getFileMetadata() {
  try {
    const props = PropertiesService.getScriptProperties();
    const fileId = props.getProperty('FILE_ID') || '1B73wD3pzJYCOxFcsZg-z2GFq0GJT2yF9';
    const file = DriveApp.getFileById(fileId);
    return {
      id: fileId,
      mimeType: file.getMimeType(),
      name: file.getName()
    };
  } catch (e) {
    // 권한 오류나 미로그인 세션 우회 시 기본 제공 이미지 속성 적용
    return {
      id: '1B73wD3pzJYCOxFcsZg-z2GFq0GJT2yF9',
      mimeType: 'image/png',
      name: 'sample.png'
    };
  }
}

/**
 * 사용자가 입력한 비밀번호가 맞는지 확인합니다.
 * @param {string} type - 'user' 또는 'admin'
 * @param {string} pwd - 입력한 비밀번호
 */
function verifyPassword(type, pwd) {
  const props = PropertiesService.getScriptProperties();
  
  if (type === 'user') {
    const userPwd = props.getProperty('USER_PWD') || 'qlcrkdhs01*';
    return pwd === userPwd;
  } else if (type === 'admin') {
    const adminPwd = props.getProperty('ADMIN_PWD') || '00admin';
    return pwd === adminPwd;
  }
  return false;
}

/**
 * 관리자 권한으로 열람 비밀번호를 변경합니다.
 * @param {string} adminPwd - 관리자 비밀번호 확인용
 * @param {string} newPwd - 새롭게 설정할 열람 비밀번호
 */
function changeUserPassword(adminPwd, newPwd) {
  const props = PropertiesService.getScriptProperties();
  const currentAdminPwd = props.getProperty('ADMIN_PWD') || '00admin';
  
  if (adminPwd === currentAdminPwd) {
    props.setProperty('USER_PWD', newPwd);
    return { success: true, message: '비밀번호가 성공적으로 변경되었습니다.' };
  } else {
    return { success: false, message: '관리자 권한이 유효하지 않습니다.' };
  }
}

/**
 * 관리자 권한으로 이미지, PDF 또는 엑셀 파일을 업로드하고 기존 등록 파일을 드라이브에서 삭제합니다.
 * @param {string} adminPwd - 관리자 비밀번호
 * @param {string} base64Data - 파일 Base64 데이터
 * @param {string} fileName - 원본 파일명
 * @param {string} mimeType - 파일 MIME 타입
 */
function uploadFile(adminPwd, base64Data, fileName, mimeType) {
  const props = PropertiesService.getScriptProperties();
  const currentAdminPwd = props.getProperty('ADMIN_PWD') || '00admin';
  
  if (adminPwd !== currentAdminPwd) {
    return { success: false, message: '관리자 권한이 유효하지 않습니다.' };
  }
  
  try {
    // 1. 기존에 등록되어 사용 중이던 구글 드라이브 파일 삭제 (휴지통 이동)
    const oldFileId = props.getProperty('FILE_ID');
    if (oldFileId) {
      try {
        DriveApp.getFileById(oldFileId).setTrashed(true);
      } catch (e) {
        console.warn('기존 파일 삭제 우회 (이미 삭제되었거나 찾을 수 없음):', e.toString());
      }
    }
    
    // 2. Base64 형식 데이터를 바이너리 디코딩하여 새 파일 생성
    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, fileName);
    const newFile = DriveApp.createFile(blob);
    
    // 3. 누구나 링크가 있다면 열람할 수 있도록 보기 권한 개방
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 4. 새로운 파일의 구글 드라이브 고유 ID를 스크립트 속성에 영구 업데이트
    const newFileId = newFile.getId();
    props.setProperty('FILE_ID', newFileId);
    
    return { 
      success: true, 
      message: '새 파일이 정상적으로 교체 업로드되었습니다.', 
      fileId: newFileId 
    };
  } catch (error) {
    return { success: false, message: '파일 업로드 실패: ' + error.toString() };
  }
}

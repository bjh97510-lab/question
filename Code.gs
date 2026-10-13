/**
 * ============================================================
 *  수학 질문 톡! (Math Q&A Board) - 백엔드 (Google Apps Script)
 * ============================================================
 *  - 데이터베이스 : Google Sheets ('질문목록' 시트)
 *  - 이미지 저장소 : Google Drive (지정 폴더)
 *  - 프론트엔드   : Index.html (HtmlService 로 서빙)
 *
 *  시트 컬럼 구조 (A~G)
 *  [고유ID, 등록시간, 질문내용, 이미지URL, 선생님답변, 답변상태, 수업반영여부]
 *
 *  ※ 학년/반/번호, 이름, 단원명 필드는 존재하지 않습니다.
 *
 *  통신 방식 (두 가지 모두 지원)
 *  1) Index.html 을 이 스크립트가 직접 서빙하는 경우 → google.script.run (CORS 불필요)
 *  2) Index.html 을 외부(GitHub Pages 등)에 올리는 경우 → doGet/doPost JSON API
 * ============================================================
 */

// ------------------------------------------------------------
// 설정
// ------------------------------------------------------------
const CONFIG = {
  // 비워두면 이 스크립트가 연결된(컨테이너 바인딩) 스프레드시트를 사용합니다.
  // 독립형(standalone) 스크립트라면 시트 URL 의 /d/ 와 /edit 사이 ID 를 넣어주세요.
  SPREADSHEET_ID: '',

  SHEET_NAME: '질문목록',

  // 이미지가 저장될 드라이브 폴더 이름 (없으면 자동 생성)
  DRIVE_FOLDER_NAME: '수학질문톡_이미지',

  // 선생님 관리자 모드 비밀번호
  // ★ 코드에 쓰지 말고 [프로젝트 설정 > 스크립트 속성] 에 TEACHER_PASSWORD 로 넣어 주세요.
  //   (이 저장소는 공개되어 있어 여기에 적으면 누구나 볼 수 있습니다)
  //   비워 두면 속성 값만 사용하며, 속성도 없으면 관리자 모드가 잠깁니다.
  TEACHER_PASSWORD: '',

  MAX_IMAGE_BYTES: 8 * 1024 * 1024, // 8MB (프론트에서 이미 압축해서 보냄)
  MAX_QUESTION_LENGTH: 1000,

  APP_TITLE: '수학 질문 톡!'
};

// H열 '삭제키' 는 학생이 자기 질문을 삭제할 때 쓰는 비공개 토큰 (목록 API 로는 노출되지 않음)
const HEADERS = ['고유ID', '등록시간', '질문내용', '이미지URL', '선생님답변', '답변상태', '수업반영여부', '삭제키'];

// 컬럼 번호 (1부터 시작)
const COL = { ID: 1, TIME: 2, QUESTION: 3, IMAGE: 4, ANSWER: 5, STATUS: 6, CLASS: 7, KEY: 8 };

const STATUS = { PENDING: '대기중', ANSWERED: '답변완료' };
const CLASS_FLAG = { NO: '미반영', YES: '반영' };

// ------------------------------------------------------------
// 최초 1회 실행: 시트/폴더 준비 (에디터에서 setupSheet 선택 후 ▶ 실행)
// ------------------------------------------------------------
function setupSheet() {
  const sheet = getSheet_();
  const folder = getImageFolder_();

  // 헤더 스타일
  const header = sheet.getRange(1, 1, 1, HEADERS.length);
  header.setValues([HEADERS]);
  header.setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(COL.ID, 280);
  sheet.setColumnWidth(COL.TIME, 150);
  sheet.setColumnWidth(COL.QUESTION, 380);
  sheet.setColumnWidth(COL.IMAGE, 280);
  sheet.setColumnWidth(COL.ANSWER, 380);
  sheet.setColumnWidth(COL.STATUS, 100);
  sheet.setColumnWidth(COL.CLASS, 110);
  sheet.setColumnWidth(COL.KEY, 120);
  try { sheet.hideColumns(COL.KEY); } catch (e) { /* 이미 숨김 */ }

  console.log('✅ 설정 완료');
  console.log('시트: ' + sheet.getParent().getUrl());
  console.log('이미지 폴더: ' + folder.getUrl());
  console.log('선생님 비밀번호: ' + (getTeacherPassword_() ? '설정됨 (스크립트 속성)' : '⚠️ 미설정 - 스크립트 속성에 TEACHER_PASSWORD 를 추가하세요'));
}

// ------------------------------------------------------------
// 웹 앱 엔트리 포인트
// ------------------------------------------------------------

/**
 * GET
 *  - ?action=list            → JSON (질문 목록)
 *  - 파라미터 없음            → Index.html 서빙
 */
function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.action) {
    return jsonOutput_(routeAction_(params.action, params));
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * POST (JSON API)
 *  body: { action: 'submit'|'answer'|'class'|'verify'|'list', ...payload }
 *  - 외부 호스팅 프론트에서 fetch 로 호출할 때 사용
 *  - Content-Type 은 'text/plain;charset=utf-8' 로 보내면 CORS preflight 가 발생하지 않음
 */
function doPost(e) {
  let body = {};
  try {
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      body = e.parameter; // form-urlencoded 전송 대응
    }
  } catch (err) {
    return jsonOutput_({ ok: false, error: '요청 본문(JSON)을 해석할 수 없습니다.' });
  }

  const action = body.action || (e && e.parameter && e.parameter.action);
  return jsonOutput_(routeAction_(action, body));
}

/** action 문자열 → 실제 함수 라우팅 */
function routeAction_(action, payload) {
  payload = payload || {};
  switch (String(action || '')) {
    case 'list':   return apiGetQuestions();
    case 'submit': return apiSubmitQuestion(payload);
    case 'verify': return apiVerifyTeacher(payload.password);
    case 'answer': return apiSaveAnswer(payload);
    case 'class':  return apiSetClassFlag(payload);
    case 'delete': return apiDeleteQuestion(payload);
    default:       return { ok: false, error: '알 수 없는 action 입니다: ' + action };
  }
}

/** JSON 응답 (ContentService 는 웹앱에서 자동으로 교차 출처 GET/POST 허용) */
function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// API 함수 (google.script.run 에서 직접 호출 가능)
//  모든 함수는 { ok: true, data } 또는 { ok: false, error } 를 반환
// ------------------------------------------------------------

/** 질문 전체 목록 (최신순) */
function apiGetQuestions() {
  return safe_(function () {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const tz = Session.getScriptTimeZone();
    const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

    return values
      .filter(function (r) { return r[COL.ID - 1] !== ''; })
      .map(function (r) {
        const t = r[COL.TIME - 1];
        return {
          id: String(r[COL.ID - 1]),
          time: formatTime_(t, tz),
          epoch: (t instanceof Date) ? t.getTime() : 0,
          question: String(r[COL.QUESTION - 1] || ''),
          imageUrl: String(r[COL.IMAGE - 1] || ''),
          answer: String(r[COL.ANSWER - 1] || ''),
          status: String(r[COL.STATUS - 1] || STATUS.PENDING),
          inClass: String(r[COL.CLASS - 1] || CLASS_FLAG.NO) === CLASS_FLAG.YES
        };
      })
      .sort(function (a, b) { return b.epoch - a.epoch; });
  });
}

/**
 * 질문 등록
 * payload: { question: string, imageBase64?: string, mimeType?: string, ownerKey?: string }
 *  - imageBase64 는 'data:image/jpeg;base64,....' 전체 또는 순수 base64 모두 허용
 *  - ownerKey 는 브라우저가 만든 무작위 토큰. 같은 브라우저에서 자기 질문을 삭제할 때 사용
 */
function apiSubmitQuestion(payload) {
  return safe_(function () {
    payload = payload || {};
    const question = String(payload.question || '').trim();
    const imageBase64 = payload.imageBase64 ? String(payload.imageBase64) : '';
    const ownerKey = String(payload.ownerKey || '').replace(/[^A-Za-z0-9_-]/g, '').substring(0, 64);

    if (!question && !imageBase64) {
      throw new Error('질문 내용이나 사진 중 하나는 꼭 필요해요.');
    }
    if (question.length > CONFIG.MAX_QUESTION_LENGTH) {
      throw new Error('질문은 ' + CONFIG.MAX_QUESTION_LENGTH + '자 이내로 작성해 주세요.');
    }

    // 1) 이미지가 있으면 드라이브에 저장
    let imageUrl = '';
    if (imageBase64) {
      imageUrl = saveImageToDrive_(imageBase64, payload.mimeType).url;
    }

    // 2) 시트에 행 추가 (동시 등록 충돌 방지를 위해 Lock 사용)
    const id = Utilities.getUuid();
    const now = new Date();
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      getSheet_().appendRow([id, now, question, imageUrl, '', STATUS.PENDING, CLASS_FLAG.NO, ownerKey]);
    } finally {
      lock.releaseLock();
    }

    return { id: id, time: formatTime_(now), imageUrl: imageUrl };
  });
}

/** 선생님 비밀번호 확인 */
function apiVerifyTeacher(password) {
  return safe_(function () {
    checkTeacher_(password);
    return true;
  });
}

/**
 * 답변 저장 → 답변이 비어있지 않으면 상태를 '답변완료' 로 자동 변경
 * payload: { id, answer, password }
 */
function apiSaveAnswer(payload) {
  return safe_(function () {
    payload = payload || {};
    checkTeacher_(payload.password);

    const sheet = getSheet_();
    const row = findRowById_(sheet, payload.id);
    const answer = String(payload.answer || '').trim();
    const status = answer ? STATUS.ANSWERED : STATUS.PENDING;

    sheet.getRange(row, COL.ANSWER, 1, 2).setValues([[answer, status]]);
    return { id: payload.id, answer: answer, status: status };
  });
}

/**
 * 수업 반영 여부 토글
 * payload: { id, inClass: boolean, password }
 */
function apiSetClassFlag(payload) {
  return safe_(function () {
    payload = payload || {};
    checkTeacher_(payload.password);

    const sheet = getSheet_();
    const row = findRowById_(sheet, payload.id);
    const inClass = (payload.inClass === true || payload.inClass === 'true');

    sheet.getRange(row, COL.CLASS).setValue(inClass ? CLASS_FLAG.YES : CLASS_FLAG.NO);
    return { id: payload.id, inClass: inClass };
  });
}

/**
 * 질문 삭제
 * payload: { id, password?: string, ownerKey?: string }
 *  - 선생님: password 가 맞으면 어떤 질문이든 삭제
 *  - 학생  : 등록할 때 보낸 ownerKey 가 시트의 삭제키와 일치할 때만 삭제 (같은 브라우저)
 *  - 첨부 사진은 드라이브 휴지통으로 이동
 */
function apiDeleteQuestion(payload) {
  return safe_(function () {
    payload = payload || {};
    const id = String(payload.id || '').trim();
    const password = String(payload.password || '');
    const ownerKey = String(payload.ownerKey || '').trim();

    const sheet = getSheet_();
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const row = findRowById_(sheet, id);
      const values = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
      const storedKey = String(values[COL.KEY - 1] || '');
      const imageUrl = String(values[COL.IMAGE - 1] || '');

      const isTeacher = isTeacherPassword_(password);
      const isOwner = ownerKey && storedKey && ownerKey === storedKey;
      if (!isTeacher && !isOwner) {
        throw new Error(password
          ? '선생님 비밀번호가 올바르지 않습니다.'
          : '이 질문은 등록한 기기에서만 삭제할 수 있어요. 선생님께 요청해 주세요.');
      }

      sheet.deleteRow(row);

      // 첨부 사진: 공유를 끊고(링크 접근 즉시 차단) 휴지통으로 (실패해도 질문 삭제는 유지)
      const m = imageUrl.match(/[-\w]{25,}/);
      if (m) {
        try {
          const file = DriveApp.getFileById(m[0]);
          try { file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW); } catch (e) { /* ignore */ }
          try { file.moveTo(DriveApp.getRootFolder()); } catch (e) { /* 폴더 권한 상속 차단용, 실패해도 진행 */ }
          file.setTrashed(true);
        } catch (err) { console.warn('이미지 휴지통 이동 실패: ' + err.message); }
      }
      return { id: id, deletedBy: isTeacher ? 'teacher' : 'owner' };
    } finally {
      lock.releaseLock();
    }
  });
}

// ------------------------------------------------------------
// 드라이브 이미지 저장
// ------------------------------------------------------------

/**
 * Base64 이미지를 드라이브 폴더에 파일로 저장하고 공개 URL 반환
 * @return {{id: string, url: string}}
 */
function saveImageToDrive_(base64, mimeType) {
  let data = String(base64).trim();
  let mime = String(mimeType || '').trim();

  // data URI 형태라면 MIME 과 데이터 분리
  if (data.indexOf('data:') === 0) {
    const commaIdx = data.indexOf(',');
    const meta = data.substring(5, commaIdx);          // image/jpeg;base64
    data = data.substring(commaIdx + 1);
    const detected = meta.split(';')[0];
    if (!mime && detected) mime = detected;
  }
  if (!mime || mime.indexOf('image/') !== 0) mime = 'image/jpeg';

  let bytes;
  try {
    bytes = Utilities.base64Decode(data);
  } catch (err) {
    throw new Error('이미지 데이터(Base64)를 해석할 수 없습니다.');
  }
  if (bytes.length === 0) throw new Error('이미지 데이터가 비어 있습니다.');
  if (bytes.length > CONFIG.MAX_IMAGE_BYTES) {
    throw new Error('이미지 용량이 너무 큽니다. (최대 ' + Math.round(CONFIG.MAX_IMAGE_BYTES / 1024 / 1024) + 'MB)');
  }

  const ext = ({ 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif', 'image/heic': 'heic' })[mime] || 'jpg';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const rand = Math.random().toString(36).substring(2, 7);
  const fileName = 'Q_' + stamp + '_' + rand + '.' + ext;

  const blob = Utilities.newBlob(bytes, mime, fileName);
  const file = getImageFolder_().createFile(blob);

  // 링크가 있는 모든 사용자가 볼 수 있게 설정 (학교 Workspace 정책으로 막혀 있으면 도메인 공유로 대체)
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err1) {
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      console.warn('ANYONE_WITH_LINK 공유 불가 → DOMAIN_WITH_LINK 로 설정: ' + err1.message);
    } catch (err2) {
      console.warn('파일 공유 설정 실패 (관리자 정책 확인 필요): ' + err2.message);
    }
  }

  const id = file.getId();
  return { id: id, url: 'https://drive.google.com/file/d/' + id + '/view' };
}

/** 이미지 폴더 가져오기 (스크립트 속성에 ID 캐시, 없으면 이름으로 찾고, 그래도 없으면 생성) */
function getImageFolder_() {
  const props = PropertiesService.getScriptProperties();
  const cachedId = props.getProperty('IMAGE_FOLDER_ID');

  if (cachedId) {
    try {
      const f = DriveApp.getFolderById(cachedId);
      if (!f.isTrashed()) return ensureFolderPrivate_(f, props);
    } catch (err) { /* 삭제된 경우 아래에서 재생성 */ }
  }

  let folder;
  const it = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (it.hasNext()) {
    folder = it.next();
  } else {
    folder = DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
  }

  props.setProperty('IMAGE_FOLDER_ID', folder.getId());
  return ensureFolderPrivate_(folder, props);
}

/**
 * 폴더는 공유하지 않음 (파일마다 개별 공유). 폴더가 공유되어 있으면 삭제한 사진이
 * 폴더 권한을 상속받아 계속 열리므로, 한 번만 비공개로 되돌린다.
 */
function ensureFolderPrivate_(folder, props) {
  if (props.getProperty('FOLDER_PRIVATE_DONE') === folder.getId()) return folder;
  try {
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
  } catch (err) {
    console.warn('폴더 비공개 설정 실패: ' + err.message);
  }
  props.setProperty('FOLDER_PRIVATE_DONE', folder.getId());
  return folder;
}

// ------------------------------------------------------------
// 시트 유틸
// ------------------------------------------------------------
function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('스프레드시트를 찾을 수 없습니다. CONFIG.SPREADSHEET_ID 를 설정해 주세요.');
  }
  return ss;
}

/** '질문목록' 시트 가져오기 (없으면 헤더와 함께 생성) */
function getSheet_() {
  const ss = getSpreadsheet_();
  syncSpreadsheetTimeZone_(ss);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < HEADERS.length) {
    // 예전 7컬럼 시트 → 누락된 헤더(삭제키 등) 자동 추가
    const from = sheet.getLastColumn() + 1;
    sheet.getRange(1, from, 1, HEADERS.length - from + 1).setValues([HEADERS.slice(from - 1)]);
    try { sheet.hideColumns(COL.KEY); } catch (e) { /* ignore */ }
  }
  return sheet;
}

/**
 * 스프레드시트 시간대를 스크립트 시간대(appsscript.json 의 timeZone)와 동일하게 맞춤
 *  - 둘이 다르면 Date 를 시트에 쓰고 읽을 때 등록시간이 몇 시간씩 어긋남
 */
function syncSpreadsheetTimeZone_(ss) {
  try {
    const scriptTz = Session.getScriptTimeZone();
    if (ss.getSpreadsheetTimeZone() !== scriptTz) {
      ss.setSpreadsheetTimeZone(scriptTz);
      console.log('스프레드시트 시간대를 ' + scriptTz + ' 로 변경했습니다.');
    }
  } catch (err) {
    console.warn('시간대 동기화 실패: ' + err.message);
  }
}

/** 고유ID 로 행 번호 찾기 (없으면 예외) */
function findRowById_(sheet, id) {
  id = String(id || '').trim();
  if (!id) throw new Error('질문 ID 가 없습니다.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('등록된 질문이 없습니다.');

  const ids = sheet.getRange(2, COL.ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 2;
  }
  throw new Error('해당 질문을 찾을 수 없습니다. (삭제되었을 수 있어요)');
}

// ------------------------------------------------------------
// 공통 유틸
// ------------------------------------------------------------
function getTeacherPassword_() {
  const fromProps = PropertiesService.getScriptProperties().getProperty('TEACHER_PASSWORD');
  return String(fromProps || CONFIG.TEACHER_PASSWORD || '');
}

/** 입력한 비밀번호가 맞으면 true (비밀번호 미설정이면 항상 false) */
function isTeacherPassword_(password) {
  const expected = getTeacherPassword_();
  return !!expected && String(password || '') === expected;
}

function checkTeacher_(password) {
  if (!getTeacherPassword_()) {
    throw new Error('선생님 비밀번호가 아직 설정되지 않았습니다. Apps Script [프로젝트 설정 > 스크립트 속성] 에 TEACHER_PASSWORD 를 추가해 주세요.');
  }
  if (!isTeacherPassword_(password)) {
    throw new Error('선생님 비밀번호가 올바르지 않습니다.');
  }
}

function formatTime_(value, tz) {
  tz = tz || Session.getScriptTimeZone();
  if (value instanceof Date) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd HH:mm');
  }
  return String(value || '');
}

/** 예외를 { ok:false, error } 로 변환해 프론트에서 일관되게 처리 */
function safe_(fn) {
  try {
    return { ok: true, data: fn() };
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
}

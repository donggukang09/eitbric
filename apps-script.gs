/**
 * 이이티브릭 서비스 시스템 - Apps Script
 *
 * 이 코드는 구글시트의 데이터를 읽어오기만 합니다.
 * 시트의 데이터는 사장님이 직접 자유롭게 수정/추가/삭제 가능합니다.
 *
 * 시트 탭 구성:
 *  - 거래처 : 고객번호, 상호, 주소, 전화1, 전화2, 담당자성명, 담당자휴대폰, 사업자번호, 정기점검일, 점검주기, AS담당직원, 비고
 *  - 기기   : 고객번호, 기종, 시리얼번호, 기기종류, 임대시작일, 임대종료일, 임대료, 중요사항
 *  - 점검이력 : (다음 단계에서 정의)
 *  - 담당자  : 이름, PIN
 *
 * ⚠ 중요: 두 탭을 연결하는 키는 "고객번호"입니다.
 *   거래처 탭과 기기 탭의 고객번호가 정확히 일치해야 합니다.
 */

const SHEET_ID = '11PS_LW35Xd7RtnINHTh6lrB6SIL9nLKAjQUbyaWMnQs';

const TAB_CUSTOMER = '거래처';
const TAB_DEVICE   = '기기';
const TAB_HISTORY  = '점검이력';
const TAB_STAFF    = '담당자';
const TAB_INSPECT_ITEMS = '점검항목';
const TAB_PARTS    = '사용부품';
const TAB_AS       = 'AS접수';

const HISTORY_HEADERS = ['점검ID','점검일','점검자','점검유형','고객번호','상호','시리얼번호','기종','점검항목','사용부품','기타사항','등록시각'];
const AS_HEADERS = ['접수ID','접수일시','접수자','분류','고객번호','상호','시리얼번호목록','기종목록','요청사항','방문요청일시','담당기사','상태','처리완료일시','비고'];


/* ============================================
 * 웹앱 진입점
 * ============================================ */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let result;

  try {
    if (action === 'staff') {
      result = getStaffList();
    } else if (action === 'login') {
      result = login(e.parameter.name, e.parameter.pin);
    } else if (action === 'search') {
      result = searchCustomers(e.parameter.q || '');
    } else if (action === 'customer') {
      result = getCustomerDetail(e.parameter.id);
    } else if (action === 'device') {
      result = getDeviceDetail(e.parameter.sn);
    } else if (action === 'dashboard') {
      result = getDashboard();
    } else if (action === 'inspectOptions') {
      result = getInspectOptions();
    } else if (action === 'saveInspect') {
      // POST 대신 GET 파라미터로 받기 (Apps Script CORS 회피)
      result = saveInspect({
        점검일: e.parameter.date || '',
        점검자: e.parameter.worker || '',
        점검유형: e.parameter.type || '',
        고객번호: e.parameter.customerId || '',
        상호: e.parameter.customerName || '',
        시리얼번호: e.parameter.sn || '',
        기종: e.parameter.model || '',
        점검항목: e.parameter.items || '',
        사용부품: e.parameter.parts || '',
        기타사항: e.parameter.memo || ''
      });
    } else if (action === 'asList') {
      result = getAsList(e.parameter.status || '');
    } else if (action === 'saveAs') {
      result = saveAs({
        접수자: e.parameter.receiver || '',
        분류: e.parameter.category || '',
        고객번호: e.parameter.customerId || '',
        상호: e.parameter.customerName || '',
        시리얼번호목록: e.parameter.snList || '',
        기종목록: e.parameter.modelList || '',
        요청사항: e.parameter.request || '',
        방문요청일시: e.parameter.visitWhen || '',
        담당기사: e.parameter.assignee || ''
      });
    } else if (action === 'updateAsStatus') {
      result = updateAsStatus(e.parameter.asId || '', e.parameter.status || '');
    } else if (action === 'asDetail') {
      result = getAsDetail(e.parameter.asId || '');
    } else {
      result = { error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // POST도 동일하게 처리 (preflight 회피용)
  return doGet(e);
}


/* ============================================
 * 시트 읽기 헬퍼
 * ============================================ */

function readSheet(tabName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error('탭을 찾을 수 없습니다: ' + tabName);

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return { headers: [], rows: [] };

  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    let hasValue = false;
    for (let j = 0; j < headers.length; j++) {
      const v = data[i][j];
      let s;
      if (v === null || v === undefined) {
        s = '';
      } else if (v instanceof Date) {
        // Date 객체는 yyyy-MM-dd 형식으로 변환
        s = Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
      } else {
        s = String(v);
      }
      obj[headers[j]] = s;
      if (s.trim() !== '') hasValue = true;
    }
    if (hasValue) rows.push(obj);
  }
  return { headers: headers, rows: rows };
}


/* ============================================
 * 담당자 / 로그인
 * ============================================ */

function getStaffList() {
  const data = readSheet(TAB_STAFF);
  const staff = data.rows.map(r => ({ name: r['이름'] }));
  return { staff: staff };
}

function login(name, pin) {
  if (!name || !pin) return { success: false };
  const data = readSheet(TAB_STAFF);
  for (let i = 0; i < data.rows.length; i++) {
    const r = data.rows[i];
    if (r['이름'] === name && String(r['PIN']) === String(pin)) {
      return { success: true, name: name };
    }
  }
  return { success: false };
}


/* ============================================
 * 거래처 검색 / 조회
 * ============================================ */

function searchCustomers(query) {
  const q = String(query || '').trim().toLowerCase();
  if (!q) return { customers: [] };

  const data = readSheet(TAB_CUSTOMER);
  const results = [];
  for (let i = 0; i < data.rows.length; i++) {
    const r = data.rows[i];
    const name = String(r['상호'] || '');
    if (name.toLowerCase().indexOf(q) >= 0) {
      results.push({
        고객번호: r['고객번호'],
        상호: name,
        주소: r['주소'] || ''
      });
    }
  }
  return { customers: results };
}

function getCustomerDetail(customerId) {
  if (!customerId) return { error: '고객번호 없음' };
  const cid = String(customerId).trim();

  // 거래처 정보
  const custData = readSheet(TAB_CUSTOMER);
  let customer = null;
  for (let i = 0; i < custData.rows.length; i++) {
    if (String(custData.rows[i]['고객번호']).trim() === cid) {
      customer = custData.rows[i];
      break;
    }
  }
  if (!customer) return { error: 'not found' };

  // 해당 거래처의 기기 리스트
  const devData = readSheet(TAB_DEVICE);
  const devices = [];
  for (let i = 0; i < devData.rows.length; i++) {
    if (String(devData.rows[i]['고객번호']).trim() === cid) {
      devices.push(devData.rows[i]);
    }
  }

  return { customer: customer, devices: devices };
}


/* ============================================
 * 기기 상세 / 점검이력
 * ============================================ */

function getDeviceDetail(serialNumber) {
  if (!serialNumber) return { error: '시리얼번호 없음' };
  const sn = String(serialNumber).trim();

  const devData = readSheet(TAB_DEVICE);
  let device = null;
  for (let i = 0; i < devData.rows.length; i++) {
    if (String(devData.rows[i]['시리얼번호']).trim() === sn) {
      device = devData.rows[i];
      break;
    }
  }
  if (!device) return { error: 'not found' };

  // 점검이력 (탭이 비어있거나 컬럼이 미정인 경우 빈 배열 반환)
  let history = [];
  try {
    const histData = readSheet(TAB_HISTORY);
    for (let i = 0; i < histData.rows.length; i++) {
      if (String(histData.rows[i]['시리얼번호'] || '').trim() === sn) {
        history.push(histData.rows[i]);
      }
    }
  } catch (err) {
    history = [];
  }

  return { device: device, history: history };
}


/* ============================================
 * 거래처 대시보드
 * ============================================ */

function getDashboard() {
  const custData = readSheet(TAB_CUSTOMER);
  const devData = readSheet(TAB_DEVICE);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let activeCount = 0;
  let expiredCount = 0;
  let soonCount = 0;
  let totalFee = 0;

  const expiredList = [];
  const soonList = [];
  const activeList = [];

  // 고객번호 → 상호 매핑
  const custMap = {};
  for (let i = 0; i < custData.rows.length; i++) {
    const r = custData.rows[i];
    const cid = String(r['고객번호'] || '').trim();
    if (cid) custMap[cid] = r['상호'] || '';
  }

  for (let i = 0; i < devData.rows.length; i++) {
    const d = devData.rows[i];
    const endStr = String(d['임대종료일'] || '').trim();
    const cid = String(d['고객번호'] || '').trim();
    const company = custMap[cid] || '';

    let status = 'unknown';
    let daysLeft = null;

    if (endStr) {
      const cleaned = endStr.replace(/\./g, '-').replace(/\//g, '-').replace(/\s/g, '');
      const m = cleaned.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) {
        const end = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
        end.setHours(0, 0, 0, 0);
        if (!isNaN(end.getTime())) {
          const diffDays = Math.round((end - today) / (1000 * 60 * 60 * 24));
          daysLeft = diffDays;
          if (diffDays < 0) status = 'expired';
          else if (diffDays <= 30) status = 'soon';
          else status = 'active';
        }
      }
    }

    // 임대료
    const feeRaw = String(d['임대료'] || '').replace(/[^0-9.-]/g, '');
    const fee = feeRaw ? Number(feeRaw) : 0;

    if (status === 'active') {
      activeCount++;
      if (!isNaN(fee)) totalFee += fee;
      activeList.push({
        고객번호: cid,
        상호: company,
        기종: d['기종'] || '',
        시리얼번호: d['시리얼번호'] || '',
        임대종료일: endStr,
        남은일수: daysLeft,
        임대료: fee
      });
    } else if (status === 'soon') {
      activeCount++;
      soonCount++;
      if (!isNaN(fee)) totalFee += fee;
      soonList.push({
        고객번호: cid,
        상호: company,
        기종: d['기종'] || '',
        시리얼번호: d['시리얼번호'] || '',
        임대종료일: endStr,
        남은일수: daysLeft,
        임대료: fee
      });
    } else if (status === 'expired') {
      expiredCount++;
      expiredList.push({
        고객번호: cid,
        상호: company,
        기종: d['기종'] || '',
        시리얼번호: d['시리얼번호'] || '',
        임대종료일: endStr,
        지난일수: -daysLeft,
        임대료: fee
      });
    }
  }

  // 정렬: 만료임박은 가까운 순, 만료된 것은 최근 만료된 순, 임대중은 종료일 가까운 순
  soonList.sort((a, b) => a.남은일수 - b.남은일수);
  expiredList.sort((a, b) => a.지난일수 - b.지난일수);
  activeList.sort((a, b) => {
    if (a.남은일수 === null) return 1;
    if (b.남은일수 === null) return -1;
    return a.남은일수 - b.남은일수;
  });

  return {
    총거래처수: custData.rows.length,
    임대중기기수: activeCount,
    월임대료총합: totalFee,
    만료임박수: soonCount,
    만료된기기수: expiredCount,
    임대중리스트: activeList,
    만료임박리스트: soonList,
    만료된리스트: expiredList
  };
}


/* ============================================
 * 점검 옵션 (점검항목 / 사용부품)
 * ============================================ */

function getInspectOptions() {
  const items = [];
  const parts = [];

  try {
    const itemData = readSheet(TAB_INSPECT_ITEMS);
    for (let i = 0; i < itemData.rows.length; i++) {
      const r = itemData.rows[i];
      const name = String(r['항목명'] || '').trim();
      const active = String(r['활성여부'] || '').trim().toUpperCase();
      if (name && (active === 'Y' || active === '')) {
        items.push(name);
      }
    }
  } catch (err) {}

  try {
    const partData = readSheet(TAB_PARTS);
    for (let i = 0; i < partData.rows.length; i++) {
      const r = partData.rows[i];
      const name = String(r['부품명'] || '').trim();
      const active = String(r['활성여부'] || '').trim().toUpperCase();
      if (name && (active === 'Y' || active === '')) {
        parts.push(name);
      }
    }
  } catch (err) {}

  return { items: items, parts: parts };
}


/* ============================================
 * 점검 저장
 * ============================================ */

function saveInspect(data) {
  if (!data.점검일) return { success: false, error: '점검일이 없습니다' };
  if (!data.점검자) return { success: false, error: '점검자가 없습니다' };
  if (!data.시리얼번호) return { success: false, error: '시리얼번호가 없습니다' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(TAB_HISTORY);

  // 점검이력 탭이 없거나 비어있으면 생성/헤더 추가
  if (!sheet) {
    sheet = ss.insertSheet(TAB_HISTORY);
  }

  // 헤더 확인 - 비어있으면 헤더 생성
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#f0f0f0');
    sheet.setFrozenRows(1);
  }

  // 점검ID 생성: INSP-YYYYMMDD-NNN
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = ('0' + (today.getMonth() + 1)).slice(-2);
  const dd = ('0' + today.getDate()).slice(-2);
  const datePrefix = 'INSP-' + yyyy + mm + dd;

  // 오늘 날짜 등록 건수 카운트
  const lastRow = sheet.getLastRow();
  let count = 0;
  if (lastRow > 1) {
    const idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < idCol.length; i++) {
      if (String(idCol[i][0]).indexOf(datePrefix) === 0) count++;
    }
  }
  const seq = ('00' + (count + 1)).slice(-3);
  const inspectId = datePrefix + '-' + seq;

  // 등록시각
  const now = new Date();
  const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  // 행 추가
  const row = [
    inspectId,
    data.점검일,
    data.점검자,
    data.점검유형,
    data.고객번호,
    data.상호,
    data.시리얼번호,
    data.기종,
    data.점검항목,
    data.사용부품,
    data.기타사항,
    nowStr
  ];
  sheet.appendRow(row);

  return { success: true, 점검ID: inspectId };
}


/* ============================================
 * AS 접수 - 저장 / 목록 / 상태 변경
 * ============================================ */

function ensureAsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(TAB_AS);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_AS);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, AS_HEADERS.length).setValues([AS_HEADERS]);
    sheet.getRange(1, 1, 1, AS_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#f0f0f0');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveAs(data) {
  if (!data.접수자) return { success: false, error: '접수자가 없습니다' };
  if (!data.고객번호) return { success: false, error: '거래처가 없습니다' };
  if (!data.시리얼번호목록) return { success: false, error: '기기가 선택되지 않았습니다' };

  const sheet = ensureAsSheet();

  // 접수ID 생성: AS-YYYYMMDD-NNN
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = ('0' + (today.getMonth() + 1)).slice(-2);
  const dd = ('0' + today.getDate()).slice(-2);
  const datePrefix = 'AS-' + yyyy + mm + dd;

  const lastRow = sheet.getLastRow();
  let count = 0;
  if (lastRow > 1) {
    const idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < idCol.length; i++) {
      if (String(idCol[i][0]).indexOf(datePrefix) === 0) count++;
    }
  }
  const seq = ('00' + (count + 1)).slice(-3);
  const asId = datePrefix + '-' + seq;

  const now = new Date();
  const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  const row = [
    asId,
    nowStr,
    data.접수자,
    data.분류,
    data.고객번호,
    data.상호,
    data.시리얼번호목록,
    data.기종목록,
    data.요청사항,
    data.방문요청일시,
    data.담당기사,
    '대기중',
    '',
    ''
  ];
  sheet.appendRow(row);

  return { success: true, 접수ID: asId };
}

function getAsList(statusFilter) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_AS);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { items: [] };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const items = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    let hasValue = false;
    for (let j = 0; j < headers.length; j++) {
      const v = data[i][j];
      let s;
      if (v === null || v === undefined) {
        s = '';
      } else if (v instanceof Date) {
        s = Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      } else {
        s = String(v);
      }
      obj[headers[j]] = s;
      if (s.trim() !== '') hasValue = true;
    }
    if (!hasValue) continue;
    if (statusFilter && obj['상태'] !== statusFilter) continue;
    items.push(obj);
  }

  // 최신순 정렬 (접수일시 desc)
  items.sort((a, b) => (b['접수일시'] || '').localeCompare(a['접수일시'] || ''));

  return { items: items };
}

function getAsDetail(asId) {
  if (!asId) return { error: '접수ID 없음' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_AS);
  if (!sheet) return { error: 'AS접수 탭이 없습니다' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === asId) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        const v = data[i][j];
        let s;
        if (v === null || v === undefined) s = '';
        else if (v instanceof Date) s = Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        else s = String(v);
        obj[headers[j]] = s;
      }
      return { item: obj };
    }
  }
  return { error: 'not found' };
}

function updateAsStatus(asId, newStatus) {
  if (!asId) return { success: false, error: '접수ID 없음' };
  if (!newStatus) return { success: false, error: '상태 없음' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(TAB_AS);
  if (!sheet) return { success: false, error: 'AS접수 탭이 없습니다' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const statusCol = headers.indexOf('상태') + 1;
  const completeCol = headers.indexOf('처리완료일시') + 1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === asId) {
      sheet.getRange(i + 1, statusCol).setValue(newStatus);
      if (newStatus === '완료') {
        const now = new Date();
        const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        sheet.getRange(i + 1, completeCol).setValue(nowStr);
      } else {
        // 다시 대기/진행으로 변경시 완료시각 비움
        sheet.getRange(i + 1, completeCol).setValue('');
      }
      return { success: true };
    }
  }
  return { success: false, error: 'not found' };
}


/* ============================================
 * 테스트 함수 (Apps Script 편집기에서 직접 실행 가능)
 * ============================================ */

function testStaff() {
  Logger.log(JSON.stringify(getStaffList()));
}

function testSearch() {
  Logger.log(JSON.stringify(searchCustomers('비키')));
}

function testCustomer() {
  // 거래처 탭의 첫 고객번호로 테스트
  const data = readSheet(TAB_CUSTOMER);
  if (data.rows.length > 0) {
    const cid = data.rows[0]['고객번호'];
    Logger.log('고객번호: ' + cid);
    Logger.log(JSON.stringify(getCustomerDetail(cid)));
  }
}

function testDashboard() {
  Logger.log(JSON.stringify(getDashboard()));
}

function testInspectOptions() {
  Logger.log(JSON.stringify(getInspectOptions()));
}

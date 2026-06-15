/** 2026 IOIA 심화과정 신청·관리 백엔드 — 구글시트 저장 + 신청자 이메일 발송 + 관리자 API */

// ===== 설정 =====
// 관리자 비밀번호는 코드에 적지 않고 Apps Script의 Script Properties(비공개 설정값)에 저장한다.
// 키: ADMIN_PASSWORD, 값: isidor-ioia-2026  → 프로젝트 설정 > 스크립트 속성에서 입력 (SETUP-backend.md 참조)
const CAPACITY = 20;                              // 회차당 정원
const FROM_NAME = '이시도르 지속가능연구소';       // 발신자 표시 이름 (실제 발신 주소는 스크립트 소유 계정 = isidor.yu@gmail.com)
const CONTACT = 'isidor.yu@gmail.com';
const SESSIONS = [
  '매스밸런스 — 9월 18일(금)',
  'ISO 17065 — 10월 14일(수)',
  'ISO 17065 — 10월 15일(목)',
  'ISO 17065 — 10월 16일(금)',
];
// ================================

function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

function authOk_(pw) {
  const real = getAdminPassword_();
  return real && pw === real; // 설정값이 비어 있으면 항상 거부
}

const HEADERS = ['접수ID', '신청일시', '성명', '소속', '이메일', '연락처', '신청회차', '문의', '상태', '입금확인일시'];

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { id = null; } }
  if (!id) {
    ss = SpreadsheetApp.create('2026 IOIA 심화과정 신청자');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName('신청자');
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName('신청자');
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  // 비공개 관리자 페이지: 공개 사이트가 아닌 이 배포 URL(?page=admin)로만 제공된다.
  if (params.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('IOIA 심화과정 관리자')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const action = params.action || 'status';
  if (action === 'status') return json_(publicStatus_());
  return json_({ ok: false, error: 'unknown' });
}

// 관리자 페이지(HtmlService)에서 google.script.run으로 호출하는 API
function adminList(pw) { return handleList_({ pw: pw }); }
function adminConfirm(pw, row) { return handleConfirm_({ pw: pw, row: row }); }
function adminUnconfirm(pw, row) { return handleUnconfirm_({ pw: pw, row: row }); }
function adminDelete(pw, id) { return handleDelete_({ pw: pw, id: id }); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  if (p.action === 'apply') return json_(handleApply_(p));
  if (p.action === 'list') return json_(handleList_(p));
  if (p.action === 'confirm') return json_(handleConfirm_(p));
  if (p.action === 'unconfirm') return json_(handleUnconfirm_(p));
  if (p.action === 'delete') return json_(handleDelete_(p));
  return json_({ ok: false, error: 'unknown action' });
}

function countPaidBySession_() {
  const rows = getSheet_().getDataRange().getValues();
  const paid = {};
  SESSIONS.forEach(function (s) { paid[s] = 0; });
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][8] === '입금확인' && paid.hasOwnProperty(rows[i][6])) paid[rows[i][6]]++;
  }
  return paid;
}

function publicStatus_() {
  const paid = countPaidBySession_();
  const out = {};
  SESSIONS.forEach(function (s) { out[s] = { paid: paid[s], full: paid[s] >= CAPACITY, capacity: CAPACITY }; });
  return { ok: true, sessions: out };
}

function handleApply_(p) {
  const name = (p['이름'] || '').toString().trim();
  const org = (p['소속'] || '').toString().trim();
  const email = (p['이메일'] || '').toString().trim();
  const phone = (p['연락처'] || '').toString().trim();
  const sess = (p['신청회차'] || '').toString().trim();
  const note = (p['문의'] || '').toString().trim();
  if (!name || !email || !sess) return { ok: false, error: '필수 항목이 누락되었습니다.' };
  if (SESSIONS.indexOf(sess) === -1) return { ok: false, error: '유효하지 않은 회차입니다.' };

  const paid = countPaidBySession_();
  const remaining = Math.max(0, CAPACITY - (paid[sess] || 0));
  const isWaitlist = remaining <= 0;
  const status = isWaitlist ? '대기' : '신청';
  const id = 'A' + new Date().getTime();
  getSheet_().appendRow([id, new Date(), name, org, email, phone, sess, note, status, '']);
  try { sendApplyEmail_(email, name, sess, isWaitlist, remaining); } catch (e) { /* 메일 실패해도 접수는 유지 */ }
  return { ok: true, waitlist: isWaitlist, remaining: remaining };
}

function handleList_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const rows = getSheet_().getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    list.push({
      row: i + 1, id: r[0], 신청일시: fmt_(r[1]), 성명: r[2], 소속: r[3], 이메일: r[4],
      연락처: r[5], 신청회차: r[6], 문의: r[7], 상태: r[8], 입금확인일시: fmt_(r[9]),
    });
  }
  const paid = countPaidBySession_();
  const summary = SESSIONS.map(function (s) {
    let applied = 0, wait = 0;
    list.forEach(function (x) {
      if (x.신청회차 !== s) return;
      if (x.상태 === '입금확인') return;
      if (x.상태 === '대기') wait++; else applied++;
    });
    return { 회차: s, 정원: CAPACITY, 입금확인: paid[s] || 0, 미입금: applied, 대기: wait, 잔여: Math.max(0, CAPACITY - (paid[s] || 0)) };
  });
  return { ok: true, summary: summary, list: list };
}

function handleConfirm_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const sh = getSheet_();
  const row = parseInt(p.row, 10);
  if (!row || row < 2) return { ok: false, error: 'bad row' };
  const data = sh.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  sh.getRange(row, 9).setValue('입금확인');
  sh.getRange(row, 10).setValue(new Date());
  try { sendConfirmEmail_(data[4], data[2], data[6]); } catch (e) {}
  return { ok: true };
}

function handleUnconfirm_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const sh = getSheet_();
  const row = parseInt(p.row, 10);
  if (!row || row < 2) return { ok: false, error: 'bad row' };
  sh.getRange(row, 9).setValue('신청');
  sh.getRange(row, 10).setValue('');
  return { ok: true };
}

function handleDelete_(p) {
  if (!authOk_(p.pw)) return { ok: false, error: 'auth' };
  const id = (p.id || '').toString();
  if (!id) return { ok: false, error: 'bad id' };
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {  // 접수ID(A열)로 매칭 → 행 번호 이동에 안전
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'not found' };
}

function fmt_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) !== '[object Date]') return d.toString();
  return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

// ===== 이메일 (발신 주소 = 스크립트 소유 계정) =====
function sendApplyEmail_(email, name, sess, isWaitlist, remaining) {
  const subject = isWaitlist
    ? '[2026 IOIA 심화과정] 대기자로 등록되었습니다'
    : '[2026 IOIA 심화과정] 신청이 접수되었습니다';
  const body = isWaitlist ? waitlistBody_(name, sess) : applyBody_(name, sess, remaining);
  GmailApp.sendEmail(email, subject, body, { name: FROM_NAME });
}

function sendConfirmEmail_(email, name, sess) {
  GmailApp.sendEmail(email, '[2026 IOIA 심화과정] 접수가 완료되었습니다', confirmBody_(name, sess), { name: FROM_NAME });
}

function applyBody_(name, sess, remaining) {
  return name + ' 님, 안녕하세요.\n\n'
    + '2026 IOIA 심화과정 "' + sess + '" 신청이 정상적으로 제출되었습니다.\n'
    + '수강료 입금이 확인되면 접수가 최종 완료됩니다.\n\n'
    + '· 신청 시점 잔여석: ' + remaining + '석 / 정원 ' + CAPACITY + '명\n\n'
    + '아래 계좌로 입금해 주세요.\n'
    + '· 수강료: 500,000원 (1인 · 1과정)\n'
    + '· 입금 은행: 기업은행(IBK)\n'
    + '· 계좌번호: 696-010037-04-016\n'
    + '· 예금주: 이시도르지속가능연구소(주)\n'
    + '· 입금자명: 신청자 본인 성함\n\n'
    + '※ 접수는 입금이 확인된 순서(입금자 순)로 처리됩니다.\n'
    + '※ 신청서를 제출하셨더라도 입금이 지연되어 그 사이 정원(' + CAPACITY + '명)이 마감되면,\n'
    + '   해당 회차는 대기자로 순연 처리될 수 있습니다. 빠른 입금을 권해드립니다.\n\n'
    + '입금이 확인되면 "접수 완료" 안내 메일을 다시 보내드리겠습니다.\n\n'
    + '문의: ' + CONTACT + '\n이시도르 지속가능연구소';
}

function waitlistBody_(name, sess) {
  return name + ' 님, 안녕하세요.\n\n'
    + '2026 IOIA 심화과정 "' + sess + '" 회차는 정원(' + CAPACITY + '명)이 마감되어\n'
    + '현재 잔여석이 없어, 대기자 명단에 등록되었습니다. (잔여석: 0석)\n\n'
    + '접수는 입금이 확인된 순서(입금자 순)로 처리되며,\n'
    + '결원이 발생하면 대기 신청 순서대로 개별 안내드리겠습니다.\n'
    + '다른 회차(특히 ISO 17065 수·목·금 회차)에 여석이 있을 수 있으니,\n'
    + '다른 요일 수강을 원하시면 본 메일에 회신해 주세요.\n\n'
    + '문의: ' + CONTACT + '\n이시도르 지속가능연구소';
}

function confirmBody_(name, sess) {
  return name + ' 님, 안녕하세요.\n\n'
    + '2026 IOIA 심화과정 "' + sess + '" 수강료 입금이 확인되어\n'
    + '접수가 최종 완료되었습니다.\n\n'
    + '교육 일정과 접속 안내(Zoom 링크 등)는 교육일 전 별도로 보내드리겠습니다.\n'
    + '감사합니다.\n\n'
    + '문의: ' + CONTACT + '\n이시도르 지속가능연구소';
}

/**
 * 생산·품질 대시보드 — 설치_한번에 (편집기에 이 파일 하나만 붙여넣기)
 * ------------------------------------------------------------
 * Code.js + 대시보드.js + 설치.js 를 합친 파일입니다.
 * 붙여넣은 뒤 실행 순서:  설치_전체  →  설치_샘플데이터  →  설치_확인
 * (clasp 로 올릴 때는 이 파일을 올리지 않습니다 — 함수 이름이 겹칩니다.
 *  .claspignore 에서 제외해 두었습니다.)
 * ------------------------------------------------------------
 */

/* ===================== Code.js ===================== */

/**
 * 생산·품질 대시보드 — 본체 (공통 도구 · 생산일보 수식/서식 · 집계 수식 · 내보내기 · 측정로그)
 * ------------------------------------------------------------
 * 실행 환경: Google Sheets + Apps Script (무료 계정, 부가기능 없음)
 * 데이터: 전부 합성(가상) 데이터. 실제 회사 자료 사용 금지.
 *
 * 만드는 것
 *   생산일보(12개월 × 5라인 × 3교대 = 5,475행)
 *     → 집계 5탭 (SUMIFS·COUNTIFS 수식)
 *     → 대시보드 (KPI 카드 6 + 차트 6 + 자동 코멘트)
 *     → 리포트(경영진 보고용 1장) PDF · 엑셀(.xlsx) 내보내기 · 월간 트리거
 *
 * 시트 탭: 안내 / 라인마스터 / 생산일보 / 목표 / 설정 /
 *          집계_월별 / 집계_라인별 / 집계_불량유형별 / 집계_교대조별 / 집계_정지사유별 /
 *          대시보드 / 리포트 / 측정로그
 *
 * 수식은 엑셀에서도 도는 것만 쓴다(SUMIFS·COUNTIFS·AVERAGEIFS·IFERROR·IF·TEXT·MAXIFS·XLOOKUP·INDEX/MATCH).
 * QUERY·FILTER·ARRAYFORMULA 는 엑셀에 없으므로 금지 — 행마다 수식을 setFormulas 로 쓴다.
 * 수식 범위는 전체 열($A:$A) 대신 실제 데이터 범위($A$2:$A$5476)로 잡는다.
 *
 * 파일 3개(Code.js · 대시보드.js · 설치.js)는 로드 순서가 보장되지 않는다 —
 * 다른 파일의 최상위 값(SH 등)은 반드시 함수 안에서만 참조한다.
 */

const TZ = 'Asia/Seoul';

const SH = {
  안내: '안내',
  라인: '라인마스터',
  생산: '생산일보',
  목표: '목표',
  설정: '설정',
  월별: '집계_월별',
  라인별: '집계_라인별',
  유형별: '집계_불량유형별',
  교대별: '집계_교대조별',
  사유별: '집계_정지사유별',
  대시보드: '대시보드',
  리포트: '리포트',
  측정: '측정로그',
};

const 교대조목록 = ['A', 'B', 'C'];
const 정지사유목록 = ['설비고장', '자재대기', '금형교체', '품질이상', '정전', '계획정비', '인원부족'];
const 불량유형목록 = ['치수불량', '외관불량', '이물혼입', '포장불량', '수분초과', '색상불량'];

/* ============================ 공통 도구 ============================ */

const 로그모음 = [];

/** Logger.log + 배열 누적(설치 결과를 한 번에 돌려주기 위함) */
function _로그(...args) {
  const s = args.map(String).join(' ');
  Logger.log(s);
  로그모음.push(s);
  return s;
}

function 로그비우기() {
  로그모음.length = 0;
  _캐시비우기();
}

function 로그전체() {
  return 로그모음.join('\n');
}

/** 실패해도 전체를 멈추지 않는 실행. 건너뛴 이유는 반드시 로그에 남긴다 */
function _안전(이름, fn) {
  try {
    return fn();
  } catch (e) {
    _로그('[건너뜀] ' + 이름 + ': ' + e.message);
    return null;
  }
}

/**
 * 구글 시트가 잠깐 멈추는 오류(`Service Spreadsheets timed out`)는 대개 일시적이라
 * 잠깐 쉬었다 다시 부르면 지나간다. 수식을 수천 개 쓰거나 행을 지운 직후에 잘 나므로
 * 그런 무거운 호출만 감싼다. 다시 시도해도 안 되면 그대로 던진다(원인을 숨기지 않는다).
 */
const 재시도문구 = /timed out|timeout|Service Spreadsheets|try again|backend error|rate limit|too many|internal error/i;

function _재시도(이름, fn, 횟수) {
  const 최대 = Math.max(Number(횟수) || 3, 1);
  for (let i = 1; ; i++) {
    try {
      return fn();
    } catch (e) {
      const 메시지 = String((e && e.message) || e);
      if (i >= 최대 || !재시도문구.test(메시지)) throw e;
      _로그('[재시도 ' + i + '/' + (최대 - 1) + '] ' + 이름 + ': ' + 메시지);
      Utilities.sleep(i * 1500);
    }
  }
}

function _시트(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('탭이 없습니다: ' + name + ' — 편집기에서 설치_전체를 먼저 실행하세요.');
  return sh;
}

function _헤더행(sh, 행) {
  const 열수 = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(행, 1, 1, 열수).getValues()[0].map(String);
}

function _헤더(sh) {
  return _헤더행(sh, 1);
}

function _열(sh, name) {
  const i = _헤더(sh).indexOf(name);
  if (i < 0) throw new Error('헤더 없음: "' + name + '" in ' + sh.getName());
  return i + 1;
}

/** 1 → A, 27 → AA */
function _열문자(n) {
  let s = '';
  let x = Number(n);
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/* ---------- 캐시 ---------- */

/**
 * 설정·생산일보·라인마스터·목표는 한 번 실행 안에서 여러 번 읽는다.
 * 읽을 때마다 시트를 다시 부르면 호출이 수천 번이 되어 6분 제한에 걸린다(②에서 실제로 그랬다).
 * 한 실행 안에서만 쓰는 메모리 캐시를 두고, 값을 쓰면 반드시 _데이터캐시비우기() 로 무효화한다.
 */
const _설정캐시 = {};
let _생산값캐시 = null;
let _라인값캐시 = null;
let _목표값캐시 = null;
let _집계값캐시 = null;

function _캐시비우기() {
  Object.keys(_설정캐시).forEach((k) => delete _설정캐시[k]);
  _데이터캐시비우기();
}

/** 데이터 탭(생산일보·라인마스터·목표)을 다시 읽게 한다. 데이터를 새로 쓴 뒤 반드시 부른다 */
function _데이터캐시비우기() {
  _생산값캐시 = null;
  _라인값캐시 = null;
  _목표값캐시 = null;
  _집계값캐시 = null;
}

/** 생산일보 전체 값(헤더 포함). 부르는 쪽에서 slice() 해서 쓴다 */
function _생산값() {
  if (!_생산값캐시) _생산값캐시 = _재시도('생산일보 읽기', () => _시트(SH.생산).getDataRange().getValues());
  return _생산값캐시;
}

function _라인값() {
  if (!_라인값캐시) _라인값캐시 = _재시도('라인마스터 읽기', () => _시트(SH.라인).getDataRange().getValues());
  return _라인값캐시;
}

function _목표값() {
  if (!_목표값캐시) _목표값캐시 = _재시도('목표 읽기', () => _시트(SH.목표).getDataRange().getValues());
  return _목표값캐시;
}

function 설정(key) {
  if (key in _설정캐시) return _설정캐시[key];
  const v = _시트(SH.설정).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) {
      _설정캐시[key] = v[i][1];
      return v[i][1];
    }
  }
  throw new Error('설정 탭에 없는 키: ' + key);
}

function _설정안전(key, 기본) {
  try {
    const v = 설정(key);
    return v === '' || v === null || v === undefined ? 기본 : v;
  } catch (e) {
    return 기본;
  }
}

/**
 * 설정 값 쓰기. 값 칸은 미리 '일반 텍스트(@)' 서식을 준다 —
 * '2025-10' 같은 값을 그냥 쓰면 시트가 날짜로 바꿔 버려 월 비교가 어긋난다.
 */
function _설정쓰기(key, 값) {
  const sh = _시트(SH.설정);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setNumberFormat('@').setValue(값);
      _설정캐시[key] = 값;
      return;
    }
  }
  const 행 = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(행, 1, 1, 3).setNumberFormat('@').setValues([[key, 값, '']]);
  _설정캐시[key] = 값;
}

/* ---------- 날짜·월 ---------- */

function _날짜(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function _지금() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

/**
 * 셀 값을 'yyyy-MM-dd' 문자열로 맞춘다.
 * 시트가 날짜로 바꿔 저장하면 Date 객체가 되고, 문자열이면 시각이 붙을 수 있어
 * 그대로 두면 날짜 비교가 어긋난다.
 */
function _날짜문자열(raw) {
  if (raw === null || raw === undefined) return '';
  if (Object.prototype.toString.call(raw) === '[object Date]') return _날짜(raw);
  const s = String(raw).trim();
  return s.length > 10 ? s.slice(0, 10) : s;
}

/** 셀 값을 'yyyy-MM' 문자열로 맞춘다(Date 로 바뀐 값도 받아 준다) */
function _월문자열(raw) {
  if (raw === null || raw === undefined) return '';
  if (Object.prototype.toString.call(raw) === '[object Date]') return Utilities.formatDate(raw, TZ, 'yyyy-MM');
  const s = String(raw).trim();
  const m = s.match(/^(\d{4})[-./년\s]*(\d{1,2})/);
  if (!m) return s.slice(0, 7);
  return m[1] + '-' + ('0' + m[2]).slice(-2);
}

/** 설정의 기간(시작월 + 개월 수)으로 'yyyy-MM' 목록을 만든다 */
function _월목록() {
  const 시작 = _월문자열(_설정안전('기간_시작월', '2025-10')) || '2025-10';
  const 개월 = Math.max(Number(_설정안전('기간_개월', 12)) || 12, 1);
  const y = Number(시작.slice(0, 4));
  const m = Number(시작.slice(5, 7));
  const 목록 = [];
  for (let i = 0; i < 개월; i++) {
    const t = m - 1 + i;
    목록.push((y + Math.floor(t / 12)) + '-' + ('0' + ((t % 12) + 1)).slice(-2));
  }
  return 목록;
}

/** 'yyyy-MM' 의 일수 */
function _월일수(월) {
  const y = Number(월.slice(0, 4));
  const m = Number(월.slice(5, 7));
  return new Date(y, m, 0).getDate();
}

/** 폴더를 이름으로 찾고, 없으면 만든다(멱등) */
function _폴더보장(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 설정의 출력 폴더. 비어 있으면 원시 드라이브 오류 대신 할 일을 알려 준다 */
function _출력폴더() {
  const id = String(_설정안전('출력폴더ID', '')).trim();
  if (!id) throw new Error('설정 탭의 출력폴더ID 가 비어 있습니다. 설치_전체를 먼저 실행하세요.');
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    throw new Error('출력 폴더(' + id + ')에 접근할 수 없습니다. 설정 탭의 출력폴더ID 를 지우고 설치_전체를 다시 실행하세요. (' + e.message + ')');
  }
}

/* ---------- 숫자 표기 ---------- */

function _율(a, b) {
  return Number(b) ? (Number(a) / Number(b)) * 100 : 0;
}

/** 12345.678 → '12,346' (자릿수 지정 가능) */
function _천단위(n, 소수) {
  const d = Number(소수) || 0;
  const x = Number(n) || 0;
  const s = x.toFixed(d);
  const 부분 = s.split('.');
  부분[0] = 부분[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return 부분.join('.');
}

/** +3.2 / -1.5 처럼 부호를 붙인다 */
function _부호(n, 소수) {
  const x = Number(n) || 0;
  return (x > 0 ? '+' : '') + x.toFixed(Number(소수) || 0);
}

/* ============================ 생산일보 수식·서식 ============================ */

/** 라인마스터의 실제 데이터 범위($A$2:$A$6 형태) */
function _라인범위() {
  const 마지막 = Math.max(_시트(SH.라인).getLastRow(), 2);
  return { 마지막: 마지막, 열: (c) => SH.라인 + '!$' + c + '$2:$' + c + '$' + 마지막 };
}

/** 생산일보의 실제 데이터 범위(생산일보!$C$2:$C$5476 형태) */
function _생산범위() {
  const 마지막 = Math.max(_시트(SH.생산).getLastRow(), 2);
  return { 마지막: 마지막, 열: (c) => SH.생산 + '!$' + c + '$2:$' + c + '$' + 마지막 };
}

/**
 * 설정 조회함수 값에 따라 XLOOKUP 또는 INDEX/MATCH 수식 문자열.
 * 찾는범위·반환범위는 '시트!$A$2:$A$6' 형태의 문자열.
 */
function _조회수식(찾을셀, 찾는범위, 반환범위, 없을때) {
  const 방식 = String(_설정안전('조회함수', 'XLOOKUP')).toUpperCase();
  const 기본 = 없을때 === undefined ? '"미등록"' : 없을때;
  if (방식 === 'INDEXMATCH') {
    return '=IFERROR(INDEX(' + 반환범위 + ',MATCH(' + 찾을셀 + ',' + 찾는범위 + ',0)),' + 기본 + ')';
  }
  return '=IFERROR(XLOOKUP(' + 찾을셀 + ',' + 찾는범위 + ',' + 반환범위 + '),' + 기본 + ')';
}

/**
 * 데이터 검증(드롭다운) 한 열.
 * requireValueInRange 는 같은 스프레드시트 안에서만 허용된다 —
 * 내보낸 xlsx 는 임시 사본(다른 파일)이라 범위 참조를 쓰면 거기서 죽는다.
 * 값 목록(문자열 배열)으로 넣으면 어느 시트에서나 동작한다.
 */
function _검증목록(sh, 머리행, 열이름, 값들) {
  const h = _헤더행(sh, 머리행);
  const i = h.indexOf(열이름);
  if (i < 0) return false;
  const 목록 = 값들.map(String).filter((v) => v.trim() !== '');
  if (!목록.length) return false;
  const 검증 = SpreadsheetApp.newDataValidation().requireValueInList(목록, true)
    .setAllowInvalid(true).setHelpText(열이름 + ' 목록에서 고르세요').build();
  const 데이터수 = sh.getLastRow() - 머리행;
  if (데이터수 > 0) sh.getRange(머리행 + 1, i + 1, 데이터수, 1).setDataValidation(검증);
  return true;
}

/**
 * 큰 범위에 수식을 쓸 때는 '한번에처리행' 단위로 나눠 쓴다.
 * 5,475행 × 3열을 한 번에 쓰면 일시적 오류가 잦고, 실패하면 처음부터 다시 해야 한다.
 */
function _수식나눠쓰기(sh, 시작행, 열, 수식들, 이름) {
  const 묶음 = Math.max(Number(_설정안전('한번에처리행', 2000)) || 2000, 100);
  for (let i = 0; i < 수식들.length; i += 묶음) {
    const 조각 = 수식들.slice(i, i + 묶음);
    _재시도(이름 + ' 수식 쓰기 ' + (i + 1) + '~' + (i + 조각.length), () =>
      sh.getRange(시작행 + i, 열, 조각.length, 1).setFormulas(조각));
  }
  return 수식들.length;
}

/**
 * 생산일보 시트에 수식 3열(월·라인명·불량률)·드롭다운 4열·조건부 서식·숫자 서식·틀 고정·필터를 넣는다.
 * 행을 추가한 뒤 다시 실행해도 같은 상태가 되도록 만든다(멱등).
 */
function 생산일보서식적용(sh) {
  const 대상 = sh || _시트(SH.생산);
  const h = _헤더(대상);
  const 마지막 = 대상.getLastRow();
  const 데이터수 = 마지막 - 1;
  if (데이터수 <= 0) {
    _로그('생산일보 서식: 데이터가 없어 건너뜁니다 (설치_샘플데이터 를 먼저 실행하세요)');
    return;
  }
  const 열 = (이름) => h.indexOf(이름) + 1;
  const 글자 = (이름) => _열문자(열(이름));
  const 라인 = _라인범위();

  // 1) 수식 3열 — 엑셀에 ARRAYFORMULA 가 없으므로 행마다 기록한다
  let 수식수 = 0;
  const 월수식 = [];
  const 라인명수식 = [];
  const 불량률수식 = [];
  for (let r = 2; r <= 마지막; r++) {
    월수식.push(['=TEXT($' + 글자('일자') + r + ',"yyyy-mm")']);
    라인명수식.push([_조회수식('$' + 글자('라인ID') + r, 라인.열('A'), 라인.열('B'))]);
    불량률수식.push(['=IFERROR($' + 글자('불량수') + r + '/$' + 글자('생산수량') + r + '*100,0)']);
  }
  수식수 += _수식나눠쓰기(대상, 2, 열('월'), 월수식, '월');
  수식수 += _수식나눠쓰기(대상, 2, 열('라인명'), 라인명수식, '라인명');
  수식수 += _수식나눠쓰기(대상, 2, 열('불량률(%)'), 불량률수식, '불량률');

  // 2) 드롭다운 4열 (값 목록 — 다른 파일로 내보내도 살아남는다)
  const 라인ID들 = _라인값().slice(1).map((r) => String(r[0]).trim()).filter((v) => v);
  let 검증수 = 0;
  if (_검증목록(대상, 1, '라인ID', 라인ID들)) 검증수++;
  if (_검증목록(대상, 1, '교대조', 교대조목록)) 검증수++;
  if (_검증목록(대상, 1, '정지사유', 정지사유목록)) 검증수++;
  if (_검증목록(대상, 1, '불량유형', 불량유형목록)) 검증수++;

  // 3) 숫자 서식
  _안전('숫자 서식', () => {
    대상.getRange(2, 열('일자'), 데이터수, 1).setNumberFormat('yyyy-mm-dd');
    ['계획수량', '생산수량', '불량수'].forEach((n) => 대상.getRange(2, 열(n), 데이터수, 1).setNumberFormat('#,##0'));
    ['가동시간(h)', '정지시간(h)'].forEach((n) => 대상.getRange(2, 열(n), 데이터수, 1).setNumberFormat('0.0'));
    대상.getRange(2, 열('불량률(%)'), 데이터수, 1).setNumberFormat('0.00');
  });

  // 4) 조건부 서식 — 같은 시트만 참조해야 엑셀에서도 산다(다른 시트 참조는 구글시트도 거부)
  const 목표 = Number(_설정안전('목표불량률기본(%)', 1.2)) || 1.2;
  const 규칙 = [
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(목표)
      .setBackground('#f4cccc').setFontColor('#990000')
      .setRanges([대상.getRange(2, 열('불량률(%)'), 데이터수, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(2)
      .setBackground('#fce5cd')
      .setRanges([대상.getRange(2, 열('정지시간(h)'), 데이터수, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('품질이상')
      .setBold(true).setFontColor('#c00000')
      .setRanges([대상.getRange(2, 열('정지사유'), 데이터수, 1)]).build(),
  ];
  _재시도('생산일보 조건부 서식', () => 대상.setConditionalFormatRules(규칙));

  // 5) 틀 고정·열 너비·필터
  대상.setFrozenRows(1);
  대상.setFrozenColumns(3);
  const 너비 = {
    일자: 92, 월: 68, 라인ID: 60, 라인명: 110, 교대조: 55, 계획수량: 80, 생산수량: 80,
    '가동시간(h)': 80, '정지시간(h)': 80, 정지사유: 80, 불량수: 65, 불량유형: 80, '불량률(%)': 75, 작업자: 65, 비고: 150,
  };
  Object.keys(너비).forEach((이름) => {
    const i = h.indexOf(이름);
    if (i >= 0) 대상.setColumnWidth(i + 1, 너비[이름]);
  });
  // 이미 필터가 있으면 createFilter 가 거부된다 → 지우고 다시 만든다(두 번 돌려도 같은 상태)
  _안전('필터', () => {
    const 기존 = 대상.getFilter();
    if (기존) 기존.remove();
    _재시도('필터', () => 대상.getRange(1, 1, 데이터수 + 1, h.length).createFilter());
  });

  _로그('생산일보 서식: 수식 ' + _천단위(수식수) + '개 / 드롭다운 ' + 검증수 + '열 / 조건부 서식 ' + 규칙.length + '규칙 / 범위 $2:$' + 마지막);
}

/* ============================ 집계 (SUMIFS·COUNTIFS) ============================ */

/**
 * 생산일보 값을 한 번 읽어 월별·라인별·교대조별·불량유형별·정지사유별로 스크립트 안에서 합산한다.
 * 시트 수식(SUMIFS)과 같은 정의를 쓴다:
 *   달성률 = 생산/계획×100, 불량률 = 불량/생산×100, 가동률 = 가동/(가동+정지)×100
 * 용도: 불량유형 정렬(파레토), 자동 코멘트, 리포트 값 — 수식이 아직 계산되지 않아도 숫자를 얻기 위해서다.
 */
function _집계값() {
  if (_집계값캐시) return _집계값캐시;
  const v = _생산값();
  const h = v[0].map(String);
  const i = {};
  ['일자', '라인ID', '교대조', '계획수량', '생산수량', '가동시간(h)', '정지시간(h)', '정지사유', '불량수', '불량유형'].forEach((k) => {
    i[k] = h.indexOf(k);
    if (i[k] < 0) throw new Error('생산일보 헤더 없음: ' + k);
  });
  const 빈 = () => ({ 계획: 0, 생산: 0, 가동: 0, 정지: 0, 불량: 0, 행: 0 });
  const 더하기 = (o, r) => {
    o.계획 += Number(r[i.계획수량]) || 0;
    o.생산 += Number(r[i.생산수량]) || 0;
    o.가동 += Number(r[i['가동시간(h)']]) || 0;
    o.정지 += Number(r[i['정지시간(h)']]) || 0;
    o.불량 += Number(r[i.불량수]) || 0;
    o.행++;
  };
  const 월별 = {};
  const 라인별 = {};
  const 교대별 = {};
  const 월라인 = {};
  const 유형별 = {};
  const 월유형 = {};
  const 사유별 = {};
  const 전체 = 빈();
  const 라인목록 = [];

  v.slice(1).forEach((r) => {
    const 월 = _월문자열(r[i.일자]);
    const 라인 = String(r[i.라인ID]).trim();
    if (!월 || !라인) return;
    const 교대 = String(r[i.교대조]).trim();
    if (!월별[월]) 월별[월] = 빈();
    if (!라인별[라인]) { 라인별[라인] = 빈(); 라인목록.push(라인); }
    if (!교대별[교대]) 교대별[교대] = 빈();
    const mk = 월 + '|' + 라인;
    if (!월라인[mk]) 월라인[mk] = 빈();
    [월별[월], 라인별[라인], 교대별[교대], 월라인[mk], 전체].forEach((o) => 더하기(o, r));

    const 불량 = Number(r[i.불량수]) || 0;
    const 유형 = String(r[i.불량유형]).trim();
    if (불량 > 0 && 유형) {
      if (!유형별[유형]) 유형별[유형] = { 불량: 0, 건수: 0 };
      유형별[유형].불량 += 불량;
      유형별[유형].건수++;
      if (!월유형[월]) 월유형[월] = {};
      월유형[월][유형] = (월유형[월][유형] || 0) + 불량;
    }
    const 정지 = Number(r[i['정지시간(h)']]) || 0;
    const 사유 = String(r[i.정지사유]).trim();
    if (정지 > 0 && 사유) {
      if (!사유별[사유]) 사유별[사유] = { 정지: 0, 건수: 0, 라인: {} };
      사유별[사유].정지 += 정지;
      사유별[사유].건수++;
      사유별[사유].라인[라인] = (사유별[사유].라인[라인] || 0) + 정지;
    }
  });

  const 유형순서 = Object.keys(유형별).sort((a, b) => 유형별[b].불량 - 유형별[a].불량);
  const 사유순서 = Object.keys(사유별).sort((a, b) => 사유별[b].정지 - 사유별[a].정지);
  _집계값캐시 = {
    월목록: Object.keys(월별).sort(),
    라인목록: 라인목록.sort(),
    월별, 라인별, 교대별, 월라인, 유형별, 월유형, 사유별, 전체, 유형순서, 사유순서,
  };
  return _집계값캐시;
}

/** 목표 탭에서 그 달의 목표(생산량·불량률)를 찾는다. 없으면 설정 기본값 */
function _목표(월) {
  const 기본률 = Number(_설정안전('목표불량률기본(%)', 1.2)) || 1.2;
  const v = _목표값();
  for (let i = 1; i < v.length; i++) {
    if (_월문자열(v[i][0]) === 월) return { 생산: Number(v[i][1]) || 0, 불량률: Number(v[i][2]) || 기본률 };
  }
  return { 생산: 0, 불량률: 기본률 };
}

/** SUMIFS 한 줄: 합계열 + 조건들 [[열문자, 조건], ...] — 생산일보의 실제 데이터 범위만 본다 */
function _합계(R, 합계열, 조건들) {
  return '=SUMIFS(' + R.열(합계열) + ',' + 조건들.map(([c, v]) => R.열(c) + ',' + v).join(',') + ')';
}

function _건수(R, 조건들) {
  return '=COUNTIFS(' + 조건들.map(([c, v]) => R.열(c) + ',' + v).join(',') + ')';
}

/** 집계 탭 한 장을 값(수식 문자열 포함)으로 다시 쓴다. 월 열은 텍스트 서식(@)을 먼저 준다 */
function _집계탭쓰기(이름, 행들, 서식) {
  const sh = _시트(이름);
  const 기존 = sh.getLastRow() - 1;
  if (기존 > 0) _재시도(이름 + ' 비우기', () => sh.getRange(2, 1, 기존, Math.max(sh.getLastColumn(), 1)).clearContent());
  if (!행들.length) return sh;
  sh.getRange(2, 1, 행들.length, 1).setNumberFormat('@');
  _재시도(이름 + ' 쓰기', () => sh.getRange(2, 1, 행들.length, 행들[0].length).setValues(행들));
  _안전(이름 + ' 숫자 서식', () => {
    Object.keys(서식 || {}).forEach((열문자) => {
      const c = 열문자.charCodeAt(0) - 64;
      sh.getRange(2, c, 행들.length, 1).setNumberFormat(서식[열문자]);
    });
  });
  sh.setFrozenRows(1);
  return sh;
}

/**
 * 집계 5탭을 현재 생산일보 범위에 맞춰 다시 쓴다.
 * 행을 추가한 뒤에도 이 함수(또는 대시보드갱신)를 돌리면 범위($2:$N)가 다시 잡힌다.
 */
function 집계갱신() {
  const 생산 = _시트(SH.생산);
  if (생산.getLastRow() < 2) {
    _로그('집계: 생산일보가 비어 있어 건너뜁니다');
    return null;
  }
  const h = _헤더(생산);
  const 글자 = (이름) => {
    const i = h.indexOf(이름);
    if (i < 0) throw new Error('생산일보 헤더 없음: ' + 이름);
    return _열문자(i + 1);
  };
  const R = _생산범위();
  const c = {
    월: 글자('월'), 라인: 글자('라인ID'), 교대: 글자('교대조'), 계획: 글자('계획수량'), 생산: 글자('생산수량'),
    가동: 글자('가동시간(h)'), 정지: 글자('정지시간(h)'), 사유: 글자('정지사유'), 불량: 글자('불량수'), 유형: 글자('불량유형'),
  };
  const 집 = _집계값();
  const 월들 = _월목록();
  const 목표마지막 = Math.max(_시트(SH.목표).getLastRow(), 2);
  const 목표범위 = (열) => SH.목표 + '!$' + 열 + '$2:$' + 열 + '$' + 목표마지막;
  const 라인 = _라인범위();

  // 1) 집계_월별
  const 월행들 = 월들.map((월, k) => {
    const r = k + 2;
    return [
      월,
      _합계(R, c.계획, [[c.월, '$A' + r]]),
      _합계(R, c.생산, [[c.월, '$A' + r]]),
      '=IFERROR(C' + r + '/B' + r + '*100,0)',
      _조회수식('$A' + r, 목표범위('A'), 목표범위('B'), '0'),
      '=IFERROR(C' + r + '/E' + r + '*100,0)',
      _합계(R, c.불량, [[c.월, '$A' + r]]),
      '=IFERROR(G' + r + '/C' + r + '*100,0)',
      _조회수식('$A' + r, 목표범위('A'), 목표범위('C'), String(_설정안전('목표불량률기본(%)', 1.2))),
      '=H' + r + '-I' + r,
      _합계(R, c.가동, [[c.월, '$A' + r]]),
      _합계(R, c.정지, [[c.월, '$A' + r]]),
      '=IFERROR(K' + r + '/(K' + r + '+L' + r + ')*100,0)',
      '=IF(H' + r + '<=I' + r + ',"달성","초과")',
    ];
  });
  const 월별탭 = _집계탭쓰기(SH.월별, 월행들, { B: '#,##0', C: '#,##0', D: '0.0', E: '#,##0', F: '0.0', G: '#,##0', H: '0.00', I: '0.00', J: '+0.00;-0.00', K: '#,##0.0', L: '#,##0.0', M: '0.0' });
  _안전('집계_월별 조건부 서식', () => 월별탭.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('초과').setBackground('#f4cccc').setFontColor('#990000')
      .setRanges([월별탭.getRange(2, 14, 월행들.length, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$H2>$I2').setFontColor('#c00000').setBold(true)
      .setRanges([월별탭.getRange(2, 8, 월행들.length, 1)]).build(),
  ]));

  // 2) 집계_라인별
  const 라인들 = _라인값().slice(1).map((r) => String(r[0]).trim()).filter((v) => v);
  const 라인행들 = 라인들.map((id, k) => {
    const r = k + 2;
    return [
      id,
      _조회수식('$A' + r, 라인.열('A'), 라인.열('B')),
      _합계(R, c.계획, [[c.라인, '$A' + r]]),
      _합계(R, c.생산, [[c.라인, '$A' + r]]),
      '=IFERROR(D' + r + '/C' + r + '*100,0)',
      _합계(R, c.불량, [[c.라인, '$A' + r]]),
      '=IFERROR(F' + r + '/D' + r + '*100,0)',
      '=IFERROR(AVERAGE(' + 목표범위('C') + '),' + String(_설정안전('목표불량률기본(%)', 1.2)) + ')',
      '=IF(G' + r + '<=H' + r + ',"달성","초과")',
      _합계(R, c.가동, [[c.라인, '$A' + r]]),
      _합계(R, c.정지, [[c.라인, '$A' + r]]),
      '=IFERROR(J' + r + '/(J' + r + '+K' + r + ')*100,0)',
    ];
  });
  const 라인별탭 = _집계탭쓰기(SH.라인별, 라인행들, { C: '#,##0', D: '#,##0', E: '0.0', F: '#,##0', G: '0.00', H: '0.00', J: '#,##0.0', K: '#,##0.0', L: '0.0' });
  // 목표 초과 라인은 빨강 — 같은 시트의 목표 열($H)만 참조한다
  _안전('집계_라인별 조건부 서식', () => 라인별탭.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$G2>$H2').setBackground('#f4cccc').setFontColor('#990000').setBold(true)
      .setRanges([라인별탭.getRange(2, 7, 라인행들.length, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('초과').setBackground('#f4cccc').setFontColor('#990000')
      .setRanges([라인별탭.getRange(2, 9, 라인행들.length, 1)]).build(),
  ]));

  // 3) 집계_불량유형별 — 파레토: 불량수 많은 순으로 정렬해 쓰고, 누적 비율을 붙인다
  const 유형순서 = 집.유형순서.slice();
  불량유형목록.forEach((t) => { if (유형순서.indexOf(t) < 0) 유형순서.push(t); });
  const 유형마지막 = 유형순서.length + 1;
  const 유형행들 = 유형순서.map((유형, k) => {
    const r = k + 2;
    return [
      유형,
      _합계(R, c.불량, [[c.유형, '$A' + r]]),
      '=IFERROR(B' + r + '/SUM($B$2:$B$' + 유형마지막 + ')*100,0)',
      '=IFERROR(SUM($B$2:B' + r + ')/SUM($B$2:$B$' + 유형마지막 + ')*100,0)',
      _건수(R, [[c.유형, '$A' + r], [c.불량, '">0"']]),
    ];
  });
  _집계탭쓰기(SH.유형별, 유형행들, { B: '#,##0', C: '0.0', D: '0.0', E: '#,##0' });

  // 4) 집계_교대조별
  const 교대행들 = 교대조목록.map((교대, k) => {
    const r = k + 2;
    return [
      교대,
      _합계(R, c.계획, [[c.교대, '$A' + r]]),
      _합계(R, c.생산, [[c.교대, '$A' + r]]),
      '=IFERROR(C' + r + '/B' + r + '*100,0)',
      _합계(R, c.불량, [[c.교대, '$A' + r]]),
      '=IFERROR(E' + r + '/C' + r + '*100,0)',
      _합계(R, c.가동, [[c.교대, '$A' + r]]),
      _합계(R, c.정지, [[c.교대, '$A' + r]]),
      '=IFERROR(G' + r + '/(G' + r + '+H' + r + ')*100,0)',
    ];
  });
  _집계탭쓰기(SH.교대별, 교대행들, { B: '#,##0', C: '#,##0', D: '0.0', E: '#,##0', F: '0.00', G: '#,##0.0', H: '#,##0.0', I: '0.0' });

  // 5) 집계_정지사유별 — 정지시간 많은 순, 라인별 열(누적 막대용)
  const 사유순서 = 집.사유순서.slice();
  정지사유목록.forEach((t) => { if (사유순서.indexOf(t) < 0) 사유순서.push(t); });
  const 사유마지막 = 사유순서.length + 1;
  const 사유행들 = 사유순서.map((사유, k) => {
    const r = k + 2;
    const 행 = [
      사유,
      _건수(R, [[c.사유, '$A' + r], [c.정지, '">0"']]),
      _합계(R, c.정지, [[c.사유, '$A' + r]]),
      '=IFERROR(C' + r + '/SUM($C$2:$C$' + 사유마지막 + ')*100,0)',
    ];
    라인들.forEach((id, j) => {
      행.push(_합계(R, c.정지, [[c.사유, '$A' + r], [c.라인, _열문자(5 + j) + '$1']]));
    });
    return 행;
  });
  const 사유서식 = { B: '#,##0', C: '#,##0.0', D: '0.0' };
  라인들.forEach((id, j) => (사유서식[_열문자(5 + j)] = '#,##0.0'));
  _집계탭쓰기(SH.사유별, 사유행들, 사유서식);

  _로그('집계: 월별 ' + 월행들.length + '행 / 라인별 ' + 라인행들.length + '행 / 불량유형별 ' + 유형행들.length +
    '행 / 교대조별 ' + 교대행들.length + '행 / 정지사유별 ' + 사유행들.length + '행 (범위 $2:$' + R.마지막 + ')');
  return { 월: 월행들.length, 라인: 라인행들.length, 유형: 유형행들.length, 교대: 교대행들.length, 사유: 사유행들.length };
}

/* ============================ 내보내기 (3단계 폴백) ============================ */

/**
 * 구글 파일(시트)을 다른 형식 블롭으로 내보낸다.
 * getAs 만으로는 계정·파일 상태에 따라
 * "Converting from application/vnd.google-apps.spreadsheet to ... is not supported" 로 막힌다(②에서 실제로 났다).
 * 그래서 다운로드 URL(파일 > 다운로드 와 같은 경로)까지 시도하고,
 * 어느 경로가 쓰였는지 로그에 남긴다(실패를 조용히 삼키지 않는다). 둘 다 안 되면 null.
 */
function _내보내기URL(파일, 형식, 추가) {
  const m = String(파일.getMimeType());
  const 종류 = m.indexOf('spreadsheet') >= 0 ? 'spreadsheets' : 'document';
  return 'https://docs.google.com/' + 종류 + '/d/' + 파일.getId() + '/export?format=' + 형식 + (추가 ? '&' + 추가 : '');
}

function _내보내기(파일, mime, 형식, 추가) {
  if (!파일 || !mime || !형식) throw new Error('내부 함수입니다. 메뉴 [대시보드] → 리포트 PDF 내보내기 / 엑셀 내보내기 로 실행하세요.');
  try {
    const b = 파일.getAs(mime);
    _로그(형식 + ' 변환: getAs 경로 사용');
    return b;
  } catch (e) {
    _로그('[건너뜀] getAs(' + 형식 + '): ' + e.message);
  }
  try {
    const 응답 = UrlFetchApp.fetch(_내보내기URL(파일, 형식, 추가), {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (응답.getResponseCode() === 200) {
      _로그(형식 + ' 변환: 다운로드 URL 경로 사용');
      return 응답.getBlob();
    }
    _로그('[건너뜀] ' + 형식 + ' 다운로드 URL: HTTP ' + 응답.getResponseCode());
  } catch (e2) {
    _로그('[건너뜀] ' + 형식 + ' 다운로드 URL: ' + e2.message);
  }
  return null;
}

/* ============================ 측정로그 ============================ */

function _측정로그(구분, 작업, 건수, 소요초, 방식, 비고) {
  const sh = _시트(SH.측정);
  const 초 = Number(소요초) || 0;
  const n = Number(건수) || 0;
  sh.appendRow([_지금(), 구분, 작업, n, 초.toFixed(1), n ? (초 / n).toFixed(2) : '', 방식 || '자동', 비고 || '']);
  return 초;
}

/**
 * 측정로그의 빈 줄·중복 줄을 정리한다(②의 발행이력 정리와 같은 역할).
 * 측정일시나 작업이 빈 줄, 그리고 (측정일시·구분·작업·건수·소요초)가 완전히 같은 줄은 마지막 줄만 남긴다.
 * 인수 없이 실행된 흔적이나 두 번 붙은 기록이 섞이면 진단 건수가 부풀려지기 때문이다.
 */
function 정리_측정로그중복() {
  const sh = _시트(SH.측정);
  if (sh.getLastRow() < 2) {
    _로그('측정로그: 정리할 기록이 없습니다');
    return 0;
  }
  const v = sh.getDataRange().getValues();
  const h = v.shift().map(String);
  const i시 = h.indexOf('측정일시');
  const i구 = h.indexOf('구분');
  const i작 = h.indexOf('작업');
  const i건 = h.indexOf('건수');
  const i초 = h.indexOf('소요초');
  const 본것 = {};
  const 지울행 = [];
  let 빈행 = 0;

  v.forEach((r, i) => {
    const 행번호 = i + 2;
    const 시각 = _날짜문자열(r[i시]) + ' ' + String(r[i시]).slice(-8);
    const 작업 = String(r[i작]).trim();
    if (!String(r[i시]).trim() || !작업) {
      빈행++;
      지울행.push(행번호);
      return;
    }
    const k = [시각, String(r[i구]).trim(), 작업, String(r[i건]).trim(), String(r[i초]).trim()].join('|');
    if (k in 본것) 지울행.push(본것[k]); // 앞선 줄을 지우고 최신 줄을 남긴다
    본것[k] = 행번호;
  });

  지울행.sort((a, b) => b - a).forEach((행번호) => _재시도('측정로그 줄 삭제', () => sh.deleteRow(행번호)));
  _안전('측정로그 정리 후 반영', () => SpreadsheetApp.flush());
  _로그('측정로그 정리: 중복 ' + (지울행.length - 빈행) + '행 / 빈 기록 ' + 빈행 + '행 삭제 → 남은 ' + Math.max(sh.getLastRow() - 1, 0) + '건');
  return 지울행.length;
}

/* ===================== 대시보드.js ===================== */

/**
 * 생산·품질 대시보드 — 대시보드 · 리포트 · 내보내기 · 트리거
 * ------------------------------------------------------------
 * 대시보드갱신: 생산일보 수식 범위를 다시 잡고 → 집계 5탭 갱신 → KPI 카드 6 · 코멘트 · 차트 6 을 다시 그린다.
 * 리포트_PDF:   경영진 보고용 1장(값으로 굳힘)을 만들어 출력 폴더에 PDF 로 저장한다.
 * 엑셀내보내기: 집계 탭을 값으로 굳힌 사본을 .xlsx 로 저장한다(엑셀에서 열림 — 차트는 빠진다).
 * 월간트리거설치: 매월 1일 아침에 위 둘을 자동 실행한다.
 *
 * 차트는 만들기 전에 기존 차트를 모두 지운다 — 다시 돌릴 때 겹겹이 쌓이지 않도록.
 * 이 파일의 최상위에는 다른 파일 값을 쓰지 않는다(로드 순서 문제).
 */

/* ============================ 대시보드 ============================ */

/** A4 세로 1장에 들어가는 대략의 높이(px). 96dpi · 여백 0.75인치 기준 — 이 파일 안에서만 쓴다 */
const A4한장 = 978;

/** 대시보드 카드·차트 배치 상수 (이 파일 안에서만 쓴다) */
function _대시보드배치() {
  return {
    열수: 12,
    카드행: 5,        // 5행 제목, 6행 값, 7행 전월 대비
    기준행: 8,        // 조건부 서식용 기준값(목표 불량률·달성 기준)
    코멘트행: 10,
    차트시작행: 13,
    차트행간격: 18,
    차트열: [1, 7],
    차트폭: 560,
    차트높이: 340,
  };
}

/**
 * 대시보드 탭을 처음부터 다시 그린다. 두 번 돌려도 같은 결과(차트는 지우고 다시 만든다).
 * 값은 집계 탭을 가리키는 수식이라 시트를 열 때마다 최신 숫자가 보이고,
 * 코멘트·갱신시각은 이 함수가 실행된 시점의 값이다.
 */
function 대시보드작성(ss) {
  const 시트 = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sh = 시트.getSheetByName(SH.대시보드) || 시트.insertSheet(SH.대시보드);
  const L = _대시보드배치();
  const 지운수 = _차트전부삭제(sh);
  sh.clear();
  _안전('대시보드 조건부 서식 초기화', () => sh.clearConditionalFormatRules());

  const 회사 = String(_설정안전('회사명', '(주)가상케미칼'));
  const 부서 = String(_설정안전('부서명', '생산1팀'));
  const 월들 = _월목록();
  const 월별탭 = _시트(SH.월별);
  const N = 월별탭.getLastRow();          // 집계_월별 마지막 행(=1+개월 수)
  const 데이터있음 = N >= 2 && _시트(SH.생산).getLastRow() >= 2;

  // 제목 블록
  sh.getRange(1, 1).setValue(회사 + ' ' + 부서 + ' 생산·품질 대시보드').setFontSize(18).setFontWeight('bold').setFontColor('#1F3864');
  sh.getRange(2, 1).setValue('기간 ' + 월들[0] + ' ~ ' + 월들[월들.length - 1] + ' · 갱신 ' + _지금() + ' · 단위: 포(25kg) / % / 시간(h)').setFontColor('#666666');
  sh.getRange(3, 1).setValue('※ ' + String(_설정안전('합성데이터고지', '이 문서의 모든 데이터는 포트폴리오 시연용 합성(가상) 데이터입니다.'))).setFontColor('#999999').setFontSize(9);
  [1, 2, 3].forEach((r) => _안전('제목 병합', () => sh.getRange(r, 1, 1, L.열수).merge()));

  // KPI 카드 6장 — 각 2열. 값은 집계_월별 범위($2:$N)를 가리키는 수식
  const 월 = (열, 행) => SH.월별 + '!$' + 열 + '$' + 행;
  const 월범위 = (열) => SH.월별 + '!$' + 열 + '$2:$' + 열 + '$' + N;
  const 전월있음 = N >= 3;
  const 증감 = (열, 서식, 단위) => (전월있음
    ? '=TEXT(' + 월(열, N) + '-' + 월(열, N - 1) + ',"' + 서식 + '")&"' + 단위 + ' (전월 대비)"'
    : '전월 데이터 없음');
  const 카드 = [
    { 이름: '총생산량 (포)', 값: '=SUM(' + 월범위('C') + ')', 서식: '#,##0', 보조: 증감('C', '+#,##0;-#,##0', ' 포') },
    { 이름: '평균 불량률 (%)', 값: '=IFERROR(SUM(' + 월범위('G') + ')/SUM(' + 월범위('C') + ')*100,0)', 서식: '0.00', 보조: 증감('H', '+0.00;-0.00', '%p') },
    { 이름: '목표 달성률 (%)', 값: '=IFERROR(SUM(' + 월범위('C') + ')/SUM(' + 월범위('E') + ')*100,0)', 서식: '0.0', 보조: 증감('F', '+0.0;-0.0', '%p') },
    { 이름: '평균 가동률 (%)', 값: '=IFERROR(SUM(' + 월범위('K') + ')/(SUM(' + 월범위('K') + ')+SUM(' + 월범위('L') + '))*100,0)', 서식: '0.0', 보조: 증감('M', '+0.0;-0.0', '%p') },
    { 이름: '총 정지시간 (h)', 값: '=SUM(' + 월범위('L') + ')', 서식: '#,##0.0', 보조: 증감('L', '+#,##0.0;-#,##0.0', ' h') },
    { 이름: '최다 불량유형', 값: '=' + SH.유형별 + '!$A$2', 서식: '@',
      보조: '=TEXT(' + SH.유형별 + '!$B$2,"#,##0")&" 포 · 전체의 "&TEXT(' + SH.유형별 + '!$C$2,"0.0")&"%"' },
  ];
  카드.forEach((k, i) => {
    const c = 1 + i * 2;
    const 제목 = sh.getRange(L.카드행, c, 1, 2);
    const 값 = sh.getRange(L.카드행 + 1, c, 1, 2);
    const 보조 = sh.getRange(L.카드행 + 2, c, 1, 2);
    _안전('카드 병합', () => { 제목.merge(); 값.merge(); 보조.merge(); });
    제목.setValue(k.이름).setBackground('#1F3864').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    if (데이터있음) {
      값.setFormula(k.값).setNumberFormat(k.서식);
      if (String(k.보조).charAt(0) === '=') 보조.setFormula(k.보조); else 보조.setValue(k.보조);
    } else {
      값.setValue('—');
      보조.setValue('설치_샘플데이터 → 대시보드갱신');
    }
    값.setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#F3F6FA');
    보조.setFontSize(9).setFontColor('#666666').setHorizontalAlignment('center').setBackground('#F3F6FA');
  });
  sh.setRowHeight(L.카드행 + 1, 40);

  // 기준값(조건부 서식은 다른 시트를 참조할 수 없어 같은 시트에 둔다)
  const 목표범위 = SH.목표 + '!$C$2:$C$' + Math.max(_시트(SH.목표).getLastRow(), 2);
  const 달성기준 = Number(_설정안전('계획달성기준(%)', 95)) || 95;
  sh.getRange(L.기준행, 1).setValue('기준값(참고)').setFontColor('#999999').setFontSize(9);
  sh.getRange(L.기준행, 2).setValue('목표 불량률(%)').setFontColor('#999999').setFontSize(9);
  sh.getRange(L.기준행, 3).setFormula('=IFERROR(AVERAGE(' + 목표범위 + '),' + String(_설정안전('목표불량률기본(%)', 1.2)) + ')').setNumberFormat('0.00').setFontColor('#999999').setFontSize(9);
  sh.getRange(L.기준행, 4).setValue('목표 달성 기준(%)').setFontColor('#999999').setFontSize(9);
  sh.getRange(L.기준행, 5).setValue(달성기준).setNumberFormat('0.0').setFontColor('#999999').setFontSize(9);
  if (데이터있음) {
    _안전('대시보드 조건부 서식', () => sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C$' + (L.카드행 + 1) + '>$C$' + L.기준행)
        .setFontColor('#c00000').setRanges([sh.getRange(L.카드행 + 1, 3, 1, 2)]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E$' + (L.카드행 + 1) + '<$E$' + L.기준행)
        .setFontColor('#c00000').setRanges([sh.getRange(L.카드행 + 1, 5, 1, 2)]).build(),
    ]));
  }

  // 종합 코멘트(실행 시점 값)
  sh.getRange(L.코멘트행 - 1, 1).setValue('■ 종합 코멘트 (자동 생성 · ' + _지금().slice(0, 16) + ')').setFontWeight('bold').setFontColor('#1F3864');
  const 코멘트 = 데이터있음 ? _종합코멘트(_집계값()) : '생산일보에 데이터가 없습니다. 설치_샘플데이터 를 실행한 뒤 대시보드갱신 을 실행하세요.';
  const 코멘트칸 = sh.getRange(L.코멘트행, 1, 1, L.열수);
  _안전('코멘트 병합', () => 코멘트칸.merge());
  코멘트칸.setValue(코멘트).setWrap(true).setVerticalAlignment('top').setBackground('#FFF8E1');
  sh.setRowHeight(L.코멘트행, 96);

  // 열 너비·틀
  for (let c = 1; c <= L.열수; c++) sh.setColumnWidth(c, 96);
  sh.setFrozenRows(0);
  _안전('격자 숨김', () => sh.setHiddenGridlines(true));

  // 차트 6개
  let 차트수 = 0;
  if (데이터있음) 차트수 = _차트전부(sh);
  else _로그('차트: 데이터가 없어 건너뜁니다 (설치_샘플데이터 → 대시보드갱신)');

  _로그('대시보드: KPI 카드 ' + 카드.length + ' / 차트 ' + 차트수 + '개 (기존 ' + 지운수 + '개 삭제 후 재생성) / 코멘트 ' + 코멘트.length + '자');
  return sh;
}

/** 시트의 차트를 모두 지운다(재실행 때 겹치지 않도록). 지운 수를 돌려준다 */
function _차트전부삭제(sh) {
  if (!sh) throw new Error('내부 함수입니다. 메뉴 [대시보드] → 대시보드 갱신 으로 실행하세요.');
  const 차트들 = sh.getCharts();
  차트들.forEach((c) => _재시도('차트 삭제', () => sh.removeChart(c)));
  return 차트들.length;
}

/** 차트 6개 정의 → 생성. 하나가 실패해도 나머지는 만든다([건너뜀] 로그) */
function _차트전부(sh) {
  if (!sh) throw new Error('내부 함수입니다. 메뉴 [대시보드] → 대시보드 갱신 으로 실행하세요.');
  const L = _대시보드배치();
  const 월별 = _시트(SH.월별);
  const 라인별 = _시트(SH.라인별);
  const 유형별 = _시트(SH.유형별);
  const 교대별 = _시트(SH.교대별);
  const 사유별 = _시트(SH.사유별);
  const mN = Math.max(월별.getLastRow(), 2);
  const lN = Math.max(라인별.getLastRow(), 2);
  const tN = Math.max(유형별.getLastRow(), 2);
  const sN = Math.max(교대별.getLastRow(), 2);
  const rN = Math.max(사유별.getLastRow(), 2);
  const 라인수 = Math.max(lN - 1, 1);

  const 정의 = [
    { 제목: '① 월별 생산량 vs 목표 (포)', 종류: 'COLUMN',
      범위: [월별.getRange(1, 1, mN, 1), 월별.getRange(1, 2, mN, 1), 월별.getRange(1, 3, mN, 1), 월별.getRange(1, 5, mN, 1)] },
    { 제목: '② 월별 불량률(%) 추이 · 목표선', 종류: 'LINE',
      범위: [월별.getRange(1, 1, mN, 1), 월별.getRange(1, 8, mN, 1), 월별.getRange(1, 9, mN, 1)] },
    { 제목: '③ 라인별 불량률(%) vs 목표', 종류: 'COLUMN',
      범위: [라인별.getRange(1, 2, lN, 1), 라인별.getRange(1, 7, lN, 1), 라인별.getRange(1, 8, lN, 1)] },
    { 제목: '④ 불량유형 파레토 (불량수 · 누적 %)', 종류: 'COMBO',
      범위: [유형별.getRange(1, 1, tN, 1), 유형별.getRange(1, 2, tN, 1), 유형별.getRange(1, 4, tN, 1)],
      옵션: { series: { 0: { type: 'bars' }, 1: { type: 'line', targetAxisIndex: 1 } } } },
    { 제목: '⑤ 교대조별 불량률(%)', 종류: 'COLUMN',
      범위: [교대별.getRange(1, 1, sN, 1), 교대별.getRange(1, 6, sN, 1)] },
    { 제목: '⑥ 정지사유별 정지시간(h) — 라인별 누적', 종류: 'COLUMN',
      범위: [사유별.getRange(1, 1, rN, 1), 사유별.getRange(1, 5, rN, 라인수)],
      옵션: { isStacked: true } },
  ];
  let 만든수 = 0;
  정의.forEach((d, i) => {
    d.행 = L.차트시작행 + Math.floor(i / 2) * L.차트행간격;
    d.열 = L.차트열[i % 2];
    if (_안전('차트 ' + d.제목, () => _차트만들기(sh, d))) 만든수++;
  });
  return 만든수;
}

/** 차트 1개. 정의 = {제목, 종류, 범위[], 행, 열, 옵션} */
function _차트만들기(sh, d) {
  if (!sh || !d || !d.범위) throw new Error('내부 함수입니다. 메뉴 [대시보드] → 대시보드 갱신 으로 실행하세요.');
  const L = _대시보드배치();
  let b = sh.newChart().setChartType(Charts.ChartType[d.종류]);
  d.범위.forEach((r) => (b = b.addRange(r)));
  b = b.setNumHeaders(1)
    .setPosition(d.행, d.열, 0, 0)
    .setOption('title', d.제목)
    .setOption('width', L.차트폭)
    .setOption('height', L.차트높이)
    .setOption('legend', { position: 'bottom' })
    .setOption('titleTextStyle', { fontSize: 13, bold: true });
  Object.keys(d.옵션 || {}).forEach((k) => (b = b.setOption(k, d.옵션[k])));
  _재시도('차트 생성 ' + d.제목, () => sh.insertChart(b.build()));
  return true;
}

/* ============================ 자동 코멘트 ============================ */

/**
 * 집계값(스크립트 합산)으로 경영진용 요약 문장을 만든다.
 * 이달(마지막 달) 실적 → 기간 누적 → 목표 초과 라인 → 최다 불량유형·정지사유 → 특이 급등 순.
 */
function _종합코멘트(집) {
  if (!집 || !집.월목록) throw new Error('내부 함수입니다. 메뉴 [대시보드] → 대시보드 갱신 으로 실행하세요.');
  const 월들 = 집.월목록;
  if (!월들.length) return '생산일보에 집계할 데이터가 없습니다.';
  const 마지막 = 월들[월들.length - 1];
  const 전 = 월들.length > 1 ? 월들[월들.length - 2] : null;
  const m = 집.월별[마지막];
  const p = 전 ? 집.월별[전] : null;
  const 목표 = _목표(마지막);
  const 이달불량률 = _율(m.불량, m.생산);
  const 문장 = [];

  // 1) 이달
  let s = 마지막 + ' 생산 ' + _천단위(m.생산) + '포';
  if (p && p.생산) s += '(전월 대비 ' + _부호(_율(m.생산 - p.생산, p.생산), 1) + '%)';
  s += ', 불량률 ' + 이달불량률.toFixed(2) + '%(목표 ' + 목표.불량률.toFixed(2) + '% ' + (이달불량률 <= 목표.불량률 ? '달성' : '초과') + ')';
  if (p) s += ', 전월 대비 ' + _부호(이달불량률 - _율(p.불량, p.생산), 2) + '%p';
  s += ', 가동률 ' + _율(m.가동, m.가동 + m.정지).toFixed(1) + '%.';
  문장.push(s);

  // 2) 기간 누적
  const T = 집.전체;
  let 목표합 = 0;
  월들.forEach((월) => (목표합 += _목표(월).생산));
  문장.push('기간 누적(' + 월들[0] + '~' + 마지막 + ') 생산 ' + _천단위(T.생산) + '포 · 목표 달성률 ' + _율(T.생산, 목표합).toFixed(1) +
    '% · 평균 불량률 ' + _율(T.불량, T.생산).toFixed(2) + '% · 평균 가동률 ' + _율(T.가동, T.가동 + T.정지).toFixed(1) +
    '% · 총 정지시간 ' + _천단위(T.정지, 1) + 'h.');

  // 3) 목표 초과 라인
  const 기준률 = Number(_설정안전('목표불량률기본(%)', 1.2)) || 1.2;
  const 초과 = 집.라인목록
    .map((id) => ({ id, 률: _율(집.라인별[id].불량, 집.라인별[id].생산) }))
    .filter((x) => x.률 > 기준률)
    .sort((a, b) => b.률 - a.률);
  문장.push(초과.length
    ? '목표 불량률(' + 기준률.toFixed(2) + '%)을 넘긴 라인: ' + 초과.map((x) => x.id + ' ' + x.률.toFixed(2) + '%').join(', ') + '.'
    : '모든 라인이 기간 평균 기준 목표 불량률(' + 기준률.toFixed(2) + '%) 이내입니다.');

  // 4) 최다 불량유형·정지사유
  const 유형 = 집.유형순서[0];
  const 사유 = 집.사유순서[0];
  if (유형 || 사유) {
    const 부분 = [];
    if (유형) 부분.push('최다 불량유형은 ' + 유형 + '(' + _천단위(집.유형별[유형].불량) + '포, 전체의 ' + _율(집.유형별[유형].불량, T.불량).toFixed(1) + '%)');
    if (사유) 부분.push('최다 정지사유는 ' + 사유 + '(' + _천단위(집.사유별[사유].정지, 1) + 'h, ' + 집.사유별[사유].건수 + '건)');
    문장.push(부분.join(', ') + '.');
  }

  // 5) 특이 급등: 어떤 달·라인의 불량률이 그 라인 연평균의 2배를 넘고 목표도 넘으면 짚어 준다
  const 급등 = [];
  집.라인목록.forEach((id) => {
    const 연평균 = _율(집.라인별[id].불량, 집.라인별[id].생산);
    월들.forEach((월) => {
      const x = 집.월라인[월 + '|' + id];
      if (!x) return;
      const 률 = _율(x.불량, x.생산);
      if (률 > 연평균 * 2 && 률 > 기준률) 급등.push({ 월, id, 률, 배: 연평균 ? 률 / 연평균 : 0 });
    });
  });
  if (급등.length) {
    급등.sort((a, b) => b.률 - a.률);
    const g = 급등[0];
    const 그달유형 = Object.keys(집.월유형[g.월] || {}).sort((a, b) => 집.월유형[g.월][b] - 집.월유형[g.월][a])[0];
    문장.push('특이사항: ' + g.월 + ' ' + g.id + ' 불량률 ' + g.률.toFixed(2) + '%(라인 연평균의 ' + g.배.toFixed(1) + '배)' +
      (그달유형 ? ' — 그 달 최다 불량유형은 ' + 그달유형 + '.' : '.') + ' 원인 조사와 재발 방지 조치 확인이 필요합니다.');
  }
  return 문장.join(' ');
}

/* ============================ 대시보드갱신 (공개) ============================ */

/**
 * 생산일보에 행을 추가했거나 목표를 바꿨을 때 실행한다.
 * ① 생산일보 수식 범위를 다시 잡고 ② 집계 5탭을 다시 쓰고 ③ 대시보드(카드·코멘트·차트)를 다시 그린다.
 */
function 대시보드갱신() {
  로그비우기();
  const 시작 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SH.생산) || !ss.getSheetByName(SH.월별)) {
    throw new Error('탭이 없습니다. 편집기에서 설치_전체 → 설치_샘플데이터 를 먼저 실행하세요.');
  }
  _로그('=== 대시보드 갱신 시작 ===');
  const 생산 = _시트(SH.생산);
  const 행수 = Math.max(생산.getLastRow() - 1, 0);
  생산일보서식적용(생산);
  집계갱신();
  대시보드작성(ss);
  _설정쓰기('마지막갱신', _지금());
  const 초 = (Date.now() - 시작) / 1000;
  _측정로그('자동', '대시보드 갱신', 행수, 초, '자동', '집계 5탭 + KPI 6 + 차트 6');
  _로그('대시보드 갱신 완료: ' + 행수 + '행 / ' + 초.toFixed(1) + '초');
  return 로그전체();
}

/* ============================ 리포트 (경영진 1장) ============================ */

/**
 * 리포트 탭을 값으로 다시 쓴다(수식 없음 → PDF·엑셀 어디서나 같은 숫자).
 * 구성: 제목·기간 → 핵심 지표 6줄 → 월별 실적 표 → 라인별 실적 표 → 종합 코멘트 → 조치 제안 → 결재란
 */
function _리포트작성(ss) {
  if (!ss) throw new Error('내부 함수입니다. 메뉴 [대시보드] → 리포트 PDF 내보내기 로 실행하세요.');
  const sh = ss.getSheetByName(SH.리포트) || ss.insertSheet(SH.리포트);
  sh.clear();
  const 열수 = 8;
  const 회사 = String(_설정안전('회사명', '(주)가상케미칼'));
  const 부서 = String(_설정안전('부서명', '생산1팀'));
  const 제목 = String(_설정안전('리포트_제목', '월간 생산·품질 리포트'));
  const 줄 = [];   // [값 배열, 스타일]
  const 넣기 = (값들, 스타일) => {
    const r = 값들.slice();
    while (r.length < 열수) r.push('');
    줄.push([r, 스타일 || '본문']);
  };

  const 집 = _집계값();
  const 월들 = 집.월목록;
  if (!월들.length) {
    넣기([제목], '제목');
    넣기(['생산일보에 데이터가 없습니다. 설치_샘플데이터 를 먼저 실행하세요.']);
    sh.getRange(1, 1, 줄.length, 열수).setValues(줄.map((x) => x[0]));
    _로그('리포트: 데이터가 없어 안내만 썼습니다');
    return sh;
  }
  const 마지막 = 월들[월들.length - 1];
  const 전 = 월들.length > 1 ? 월들[월들.length - 2] : null;
  const m = 집.월별[마지막];
  const p = 전 ? 집.월별[전] : null;
  const T = 집.전체;
  const 목표 = _목표(마지막);
  const 달성기준 = Number(_설정안전('계획달성기준(%)', 95)) || 95;
  let 목표합 = 0;
  월들.forEach((월) => (목표합 += _목표(월).생산));
  const 률 = (o) => (o ? _율(o.불량, o.생산) : 0);
  const 가동 = (o) => (o ? _율(o.가동, o.가동 + o.정지) : 0);
  const 최다유형 = (월) => {
    const t = 집.월유형[월] || {};
    const k = Object.keys(t).sort((a, b) => t[b] - t[a])[0];
    return k ? k + ' ' + _천단위(t[k]) + '포' : '-';
  };

  넣기([회사 + ' ' + 부서 + ' ' + 제목 + ' — ' + 마지막], '제목');
  넣기(['기간 ' + 월들[0] + ' ~ ' + 마지막 + ' · 작성 ' + _날짜(new Date()) + ' · 단위 포(25kg) / % / h'], '부제');
  넣기(['※ ' + String(_설정안전('합성데이터고지', '포트폴리오 시연용 합성(가상) 데이터입니다.'))], '고지');
  넣기(['']);
  넣기(['1. 핵심 지표'], '소제목');
  넣기(['지표', '기간 누적', '이달(' + 마지막 + ')', '전월', '전월 대비', '목표/기준', '판정'], '머리');
  const 이달달성 = _율(m.생산, 목표.생산);
  넣기(['생산량(포)', _천단위(T.생산), _천단위(m.생산), p ? _천단위(p.생산) : '-', p ? _부호(_율(m.생산 - p.생산, p.생산), 1) + '%' : '-',
    _천단위(목표.생산), 이달달성 >= 달성기준 ? '달성' : '미달']);
  넣기(['불량률(%)', 률(T).toFixed(2), 률(m).toFixed(2), p ? 률(p).toFixed(2) : '-', p ? _부호(률(m) - 률(p), 2) + '%p' : '-',
    목표.불량률.toFixed(2), 률(m) <= 목표.불량률 ? '달성' : '초과']);
  넣기(['목표 달성률(%)', _율(T.생산, 목표합).toFixed(1), 이달달성.toFixed(1), p ? _율(p.생산, _목표(전).생산).toFixed(1) : '-',
    p ? _부호(이달달성 - _율(p.생산, _목표(전).생산), 1) + '%p' : '-', 달성기준.toFixed(1), 이달달성 >= 달성기준 ? '달성' : '미달']);
  넣기(['가동률(%)', 가동(T).toFixed(1), 가동(m).toFixed(1), p ? 가동(p).toFixed(1) : '-', p ? _부호(가동(m) - 가동(p), 1) + '%p' : '-', '-', '-']);
  넣기(['정지시간(h)', _천단위(T.정지, 1), _천단위(m.정지, 1), p ? _천단위(p.정지, 1) : '-', p ? _부호(m.정지 - p.정지, 1) + 'h' : '-', '-', '-']);
  const 연간유형 = 집.유형순서[0];
  넣기(['최다 불량유형', 연간유형 ? 연간유형 + ' ' + _율(집.유형별[연간유형].불량, T.불량).toFixed(1) + '%' : '-', 최다유형(마지막), 전 ? 최다유형(전) : '-', '-', '-', '-']);
  넣기(['']);

  넣기(['2. 월별 실적'], '소제목');
  넣기(['월', '계획(포)', '생산(포)', '달성률(%)', '불량률(%)', '목표 불량률(%)', '가동률(%)', '정지시간(h)'], '머리');
  월들.forEach((월) => {
    const x = 집.월별[월];
    넣기([월, _천단위(x.계획), _천단위(x.생산), _율(x.생산, x.계획).toFixed(1), 률(x).toFixed(2), _목표(월).불량률.toFixed(2), 가동(x).toFixed(1), _천단위(x.정지, 1)], '표');
  });
  넣기(['']);

  넣기(['3. 라인별 실적 (기간 누적)'], '소제목');
  넣기(['라인', '라인명', '생산(포)', '달성률(%)', '불량률(%)', '판정', '가동률(%)', '정지시간(h)'], '머리');
  const 라인명 = {};
  _라인값().slice(1).forEach((r) => (라인명[String(r[0]).trim()] = r[1]));
  const 기준률 = Number(_설정안전('목표불량률기본(%)', 1.2)) || 1.2;
  집.라인목록.forEach((id) => {
    const x = 집.라인별[id];
    넣기([id, 라인명[id] || '', _천단위(x.생산), _율(x.생산, x.계획).toFixed(1), 률(x).toFixed(2), 률(x) <= 기준률 ? '달성' : '초과', 가동(x).toFixed(1), _천단위(x.정지, 1)], '표');
  });
  넣기(['']);

  넣기(['4. 종합 코멘트'], '소제목');
  넣기([_종합코멘트(집)], '긴글');
  넣기(['']);

  넣기(['5. 조치 제안'], '소제목');
  _조치제안(집).forEach((s) => 넣기(['· ' + s], '조치'));
  넣기(['']);
  넣기(['', '', '', '', '', '담당', '검토', '승인'], '결재');
  넣기(['', '', '', '', '', '', '', ''], '서명');

  // 쓰기 + 서식
  _재시도('리포트 쓰기', () => sh.getRange(1, 1, 줄.length, 열수).setValues(줄.map((x) => x[0])));
  const 긴글행 = [];
  const 조치행 = [];
  줄.forEach((x, i) => {
    const r = i + 1;
    const 스타일 = x[1];
    const 전체 = sh.getRange(r, 1, 1, 열수);
    if (스타일 === '제목') { _안전('병합', () => 전체.merge()); 전체.setFontSize(15).setFontWeight('bold').setFontColor('#1F3864'); }
    else if (스타일 === '부제') { _안전('병합', () => 전체.merge()); 전체.setFontSize(10).setFontColor('#666666'); }
    else if (스타일 === '고지') { _안전('병합', () => 전체.merge()); 전체.setFontSize(8).setFontColor('#999999'); }
    else if (스타일 === '소제목') { 전체.setFontSize(11).setFontWeight('bold').setFontColor('#1F3864').setBorder(false, false, true, false, false, false, '#1F3864', SpreadsheetApp.BorderStyle.SOLID); }
    else if (스타일 === '머리') { 전체.setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF').setHorizontalAlignment('center').setFontSize(9); }
    else if (스타일 === '표') { 전체.setFontSize(9).setHorizontalAlignment('right'); sh.getRange(r, 1, 1, 2).setHorizontalAlignment('left'); }
    else if (스타일 === '긴글') { _안전('병합', () => 전체.merge()); 전체.setWrap(true).setVerticalAlignment('top').setFontSize(9); 긴글행.push(r); }
    else if (스타일 === '조치') { _안전('병합', () => 전체.merge()); 전체.setWrap(true).setVerticalAlignment('top').setFontSize(9); 조치행.push(r); }
    else if (스타일 === '결재') { sh.getRange(r, 6, 1, 3).setBackground('#F1F3F4').setFontColor('#666666').setHorizontalAlignment('center').setFontSize(9); }
    else if (스타일 === '서명') { sh.getRange(r, 6, 1, 3).setBorder(true, true, true, true, true, false, '#999999', SpreadsheetApp.BorderStyle.SOLID); }
    else if (스타일 === '본문') { 전체.setFontSize(9); }
  });
  // A4 세로 1장에 담기게 행 높이를 조인다 — 기본 21px 로 두면 47행이 1장을 넘겨
// '5. 조치 제안' 제목만 앞장 끝에 남고 내용이 다음 장으로 밀린다(실제로 그랬다).
  // 표·본문은 17px, 종합 코멘트(4줄) 60px, 조치 제안(1~2줄) 30px, 서명란 32px.
  const 행높이 = { 본문: 17, 긴글: 60, 조치: 30, 서명: 32 };
  _안전('행 높이', () => {
    sh.setRowHeights(1, 줄.length, 행높이.본문);
    긴글행.forEach((r) => sh.setRowHeight(r, 행높이.긴글));
    조치행.forEach((r) => sh.setRowHeight(r, 행높이.조치));
    sh.setRowHeight(줄.length, 행높이.서명);
  });
  const 예상높이 = (줄.length - 긴글행.length - 조치행.length - 1) * 행높이.본문 +
    긴글행.length * 행높이.긴글 + 조치행.length * 행높이.조치 + 행높이.서명;
  // 핵심 지표 표: 판정 '초과'·'미달' 빨강 (같은 시트만 참조)
  _안전('리포트 조건부 서식', () => sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('초과').setFontColor('#c00000').setBold(true).setRanges([sh.getRange(1, 1, 줄.length, 열수)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('미달').setFontColor('#c00000').setBold(true).setRanges([sh.getRange(1, 1, 줄.length, 열수)]).build(),
  ]));
  const 너비 = [110, 90, 90, 80, 80, 90, 80, 80];
  너비.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  _안전('격자 숨김', () => sh.setHiddenGridlines(true));
  _로그('리포트: ' + 줄.length + '행 (핵심 지표 6 · 월별 ' + 월들.length + ' · 라인별 ' + 집.라인목록.length + ' · 조치 제안 ' + _조치제안(집).length + ')');
  _로그('리포트: 예상 높이 약 ' + 예상높이 + 'px / A4 세로 1장 ' + A4한장 + 'px — 넘으면 PDF 가 2장으로 갈라집니다');
  return sh;
}

/** 집계값으로 조치 제안 문장 2~4개 */
function _조치제안(집) {
  if (!집 || !집.라인목록) throw new Error('내부 함수입니다. 메뉴 [대시보드] → 리포트 PDF 내보내기 로 실행하세요.');
  const 제안 = [];
  const 기준률 = Number(_설정안전('목표불량률기본(%)', 1.2)) || 1.2;
  const 초과 = 집.라인목록.filter((id) => _율(집.라인별[id].불량, 집.라인별[id].생산) > 기준률);
  const 유형 = 집.유형순서[0];
  if (초과.length) 제안.push('목표 불량률 초과 라인(' + 초과.join(', ') + '): ' + (유형 ? '최다 불량유형 "' + 유형 + '" 의 ' : '') + '발생 공정 조건(온도·압력·원료 로트)을 재점검하고 주간 품질 회의에서 추적한다.');
  const 사유 = 집.사유순서[0];
  if (사유) {
    const 문구 = { 설비고장: '예방정비 주기 단축과 예비 부품 확보를 검토한다', 자재대기: '자재 입고 일정과 안전 재고 기준을 조정한다', 금형교체: '교체 표준 작업(SMED)으로 교체 시간을 줄인다',
      품질이상: '품질 이상 발생 즉시 라인 정지 기준과 원인 분석 절차를 점검한다', 정전: '비상 전원·재기동 절차를 점검한다', 계획정비: '정비 계획을 생산 계획에 미리 반영한다', 인원부족: '교대 인원 배치와 대체 인력 계획을 세운다' };
    제안.push('정지사유 1위 "' + 사유 + '"(' + _천단위(집.사유별[사유].정지, 1) + 'h): ' + (문구[사유] || '원인별 대책을 세운다') + '.');
  }
  const 교대률 = 교대조목록.map((c) => ({ c, 률: 집.교대별[c] ? _율(집.교대별[c].불량, 집.교대별[c].생산) : 0 }));
  const 최고 = 교대률.slice().sort((a, b) => b.률 - a.률)[0];
  const 평균 = 교대률.reduce((a, x) => a + x.률, 0) / Math.max(교대률.length, 1);
  if (최고 && 평균 && 최고.률 > 평균 * 1.15) 제안.push(최고.c + '조 불량률(' + 최고.률.toFixed(2) + '%)이 평균(' + 평균.toFixed(2) + '%)보다 높다: 교대 인수인계 점검 항목과 야간 조명·순회 점검을 강화한다.');
  if (!제안.length) 제안.push('모든 지표가 목표 이내이다. 현재 관리 수준을 유지하며 다음 달 목표를 단계적으로 상향 검토한다.');
  return 제안;
}

/** 리포트 탭만 다시 만든다(PDF 없이) */
function 리포트작성() {
  로그비우기();
  _리포트작성(SpreadsheetApp.getActiveSpreadsheet());
  return 로그전체();
}

/**
 * 리포트 탭 → PDF. 임시 사본을 만들어 리포트 탭만 남기고 변환한다.
 * 변환 경로: ① getAs → ② 다운로드 URL(/export?format=pdf) → ③ 둘 다 막히면 구글시트 사본을 출력 폴더에 보존.
 */
function 리포트_PDF() {
  로그비우기();
  const 시작 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _리포트작성(ss);
  const 집 = _집계값();
  const 마지막월 = 집.월목록.length ? 집.월목록[집.월목록.length - 1] : _날짜(new Date()).slice(0, 7);
  const 이름 = '생산품질_리포트_' + 마지막월;
  const 출력 = _출력폴더();
  SpreadsheetApp.flush();

  let 임시ID = null;
  let 결과 = null;
  let 보존 = false;
  try {
    임시ID = DriveApp.getFileById(ss.getId()).makeCopy('임시_' + 이름).getId();
    const 임시 = SpreadsheetApp.openById(임시ID);
    const 리포트 = 임시.getSheetByName(SH.리포트);
    if (!리포트) throw new Error('임시 사본에 리포트 탭이 없습니다');
    // 리포트는 값만 있으므로 다른 탭을 지워도 #REF! 가 나지 않는다 → PDF 에 리포트 한 장만 담긴다
    임시.getSheets().forEach((s) => { if (s.getName() !== SH.리포트) 임시.deleteSheet(s); });
    SpreadsheetApp.flush();
    Utilities.sleep(1000);

    const 파일 = DriveApp.getFileById(임시ID);
    const pdf = _내보내기(파일, 'application/pdf', 'pdf',
      // fitw(가로 맞춤)만으로는 세로가 넘쳐 2장이 됐다 → fith(세로 맞춤)까지 걸고 여백을 줄인다
      'gid=' + 리포트.getSheetId() + '&portrait=true&fitw=true&fith=true&size=A4&gridlines=false&printtitle=false&sheetnames=false' +
      '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4');
    if (pdf) {
      결과 = 출력.createFile(pdf.setName(이름 + '.pdf')).getUrl();
      _로그('PDF 저장: ' + 결과);
    } else {
      출력.addFile(파일);
      보존 = true;
      결과 = 파일.getUrl();
      _로그('[대체] PDF 변환이 막혀 구글시트 사본(리포트 탭)으로 저장했습니다: ' + 결과);
    }
  } finally {
    if (임시ID && !보존) _안전('임시본 정리', () => DriveApp.getFileById(임시ID).setTrashed(true));
  }

  _메일알림(이름, '월간 생산·품질 리포트가 만들어졌습니다.\n' + 결과 + '\n\n' + _종합코멘트(집));
  _측정로그('자동', '리포트 PDF', 1, (Date.now() - 시작) / 1000, '자동', 이름 + '.pdf');
  _로그('리포트 PDF 완료: ' + ((Date.now() - 시작) / 1000).toFixed(1) + '초');
  return 결과;
}

/** 설정 '메일알림' 에 주소가 있을 때만 보낸다. 없거나 실패하면 [건너뜀] 로그 */
function _메일알림(제목, 본문) {
  const 주소 = String(_설정안전('메일알림', '')).trim();
  if (!주소) {
    _로그('[건너뜀] 메일 알림: 설정 탭 메일알림 이 비어 있습니다(원하면 받을 주소를 적으세요)');
    return false;
  }
  return _안전('메일 알림', () => {
    MailApp.sendEmail(주소, '[대시보드] ' + 제목, 본문);
    _로그('메일 알림: ' + 주소);
    return true;
  }) || false;
}

/* ============================ 엑셀(.xlsx) 내보내기 ============================ */

/**
 * 집계 5탭·대시보드 카드를 값으로 굳힌 사본을 .xlsx 로 저장한다.
 * 차트는 xlsx 변환에서 빠진다 — 그래서 집계 탭이 통합문서의 본체다(②의 요약 탭과 같은 이유).
 * 생산일보의 수식(TEXT·XLOOKUP·IFERROR)은 엑셀에서도 계산된다. 구버전 엑셀이면 설정 조회함수=INDEXMATCH.
 */
function 엑셀내보내기() {
  로그비우기();
  const 시작 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 집 = _집계값();
  const 마지막월 = 집.월목록.length ? 집.월목록[집.월목록.length - 1] : _날짜(new Date()).slice(0, 7);
  const 이름 = '생산품질_대시보드_' + 마지막월;
  const 출력 = _출력폴더();
  const 행수 = Math.max(_시트(SH.생산).getLastRow() - 1, 0);

  let 임시ID = null;
  let 결과 = null;
  let 보존 = false;
  try {
    임시ID = DriveApp.getFileById(ss.getId()).makeCopy('임시_' + 이름).getId();
    const 임시 = SpreadsheetApp.openById(임시ID);

    // 집계·대시보드는 값으로 굳힌다 — 엑셀에서 수식 재계산 없이 같은 숫자가 보이도록
    let 굳힌탭 = 0;
    [SH.월별, SH.라인별, SH.유형별, SH.교대별, SH.사유별, SH.대시보드].forEach((n) => {
      const s = 임시.getSheetByName(n);
      if (!s || s.getLastRow() < 1) return;
      const 범위 = s.getDataRange();
      _재시도(n + ' 값 굳히기', () => 범위.setValues(범위.getValues()));
      굳힌탭++;
    });
    // 내보내기에 필요 없는 탭 제거
    [SH.안내, SH.설정, SH.측정].forEach((n) => {
      const s = 임시.getSheetByName(n);
      if (s) 임시.deleteSheet(s);
    });
    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    const 파일 = DriveApp.getFileById(임시ID);
    const xlsx = _내보내기(파일, MimeType.MICROSOFT_EXCEL, 'xlsx');
    if (xlsx) {
      결과 = 출력.createFile(xlsx.setName(이름 + '.xlsx')).getUrl();
      _로그('xlsx 저장: ' + 결과 + ' (값으로 굳힌 탭 ' + 굳힌탭 + ')');
    } else {
      출력.addFile(파일);
      보존 = true;
      결과 = 파일.getUrl();
      _로그('[대체] 엑셀 변환이 막혀 구글시트 사본으로 저장했습니다: ' + 결과);
    }
  } finally {
    if (임시ID && !보존) _안전('임시본 정리', () => DriveApp.getFileById(임시ID).setTrashed(true));
  }

  _측정로그('자동', '엑셀 내보내기', 행수, (Date.now() - 시작) / 1000, '자동', 이름 + '.xlsx');
  _로그('엑셀 내보내기 완료: ' + ((Date.now() - 시작) / 1000).toFixed(1) + '초');
  return 결과;
}

/* ============================ 월간 트리거 ============================ */

/** 트리거가 부르는 함수: 대시보드 갱신 + 리포트 PDF */
function 월간자동실행() {
  대시보드갱신();
  const url = 리포트_PDF();
  _로그('월간 자동 실행 완료: ' + url);
  return 로그전체();
}

function 월간트리거설치() {
  로그비우기();
  const 시각 = Math.min(Math.max(Number(_설정안전('트리거시각', 7)) || 7, 0), 23);
  let 지운수 = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === '월간자동실행') { ScriptApp.deleteTrigger(t); 지운수++; }
  });
  ScriptApp.newTrigger('월간자동실행').timeBased().onMonthDay(1).atHour(시각).create();
  _로그('월간 트리거 설치: 매월 1일 ' + 시각 + '시 대시보드갱신 + 리포트_PDF (기존 ' + 지운수 + '개 교체)');
  return 로그전체();
}

function 월간트리거해제() {
  로그비우기();
  let 지운수 = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === '월간자동실행') { ScriptApp.deleteTrigger(t); 지운수++; }
  });
  _로그('월간 트리거 해제: ' + 지운수 + '개 삭제');
  return 로그전체();
}

/* ===================== 설치.js ===================== */

/**
 * 생산·품질 대시보드 — 설치 · 샘플 데이터 · 자가 진단 · 메뉴
 * ------------------------------------------------------------
 * 설치_전체:       탭·설정·폴더를 만들고 수식·서식·대시보드 틀을 적용한다(두 번 돌려도 데이터가 지워지지 않는다).
 * 설치_샘플데이터: 합성 생산일보 5,475행(12개월 × 5라인 × 3교대) + 목표 12행 + 라인마스터 5행을 만들고 집계·대시보드까지 갱신한다.
 * 설치_확인:       12항목 ○/× 자가 진단.
 *
 * 처음 한 번:  설치_전체  →  설치_샘플데이터  →  설치_확인
 * 그 다음:     메뉴 [대시보드] 에서 갱신 · 리포트 PDF · 엑셀 내보내기 · 월간 트리거
 *
 * 이 파일의 최상위에서는 다른 파일(Code.js)의 SH 등을 참조하지 않는다 —
 * Apps Script 는 파일 실행 순서를 보장하지 않아 "SH is not defined" 로 터진다(②에서 실제로 났다).
 * 탭정의·설정기본값·라인정의는 함수로 감싸 호출 시점에 만든다.
 */

/** 작업자 15명 (라인 5 × 교대 3, 전부 가상 인명) */
const 작업자목록 = ['강도윤', '고은채', '남재현', '문서준', '배지민', '서하늘', '신예은', '오태양', '유지호', '윤소율', '임가온', '장우진', '조하린', '하준서', '한수아'];

/** 라인 5개 (전부 가상). 설계능력은 톤/일, 교대당 능력은 포(25kg) 단위로 환산한다 */
function 라인정의() {
  return [
    { 라인ID: 'L1', 라인명: '1호 압출라인', 제품: 'PE 펠릿', 설비수: 6, '설계능력(톤/일)': 120, 담당자: '김민수', 불량기본: 0.7, 유형가중: { 치수불량: 4, 외관불량: 3, 이물혼입: 1, 포장불량: 1, 수분초과: 1, 색상불량: 2 } },
    { 라인ID: 'L2', 라인명: '2호 압출라인', 제품: 'PP 펠릿', 설비수: 5, '설계능력(톤/일)': 100, 담당자: '이서연', 불량기본: 1.0, 유형가중: { 치수불량: 3, 외관불량: 4, 이물혼입: 1, 포장불량: 1, 수분초과: 1, 색상불량: 2 } },
    { 라인ID: 'L3', 라인명: '3호 배합라인', 제품: '접착제 A', 설비수: 4, '설계능력(톤/일)': 90, 담당자: '박지훈', 불량기본: 0.9, 유형가중: { 치수불량: 1, 외관불량: 2, 이물혼입: 3, 포장불량: 2, 수분초과: 3, 색상불량: 1 } },
    { 라인ID: 'L4', 라인명: '4호 배합라인', 제품: '수지 B', 설비수: 7, '설계능력(톤/일)': 150, 담당자: '최유진', 불량기본: 1.5, 유형가중: { 치수불량: 5, 외관불량: 3, 이물혼입: 2, 포장불량: 1, 수분초과: 2, 색상불량: 1 } },
    { 라인ID: 'L5', 라인명: '5호 포장라인', 제품: '첨가제 C', 설비수: 3, '설계능력(톤/일)': 60, 담당자: '정도현', 불량기본: 0.6, 유형가중: { 치수불량: 1, 외관불량: 2, 이물혼입: 1, 포장불량: 5, 수분초과: 1, 색상불량: 1 } },
  ];
}

/* ============================ 탭 정의 ============================ */

/**
 * 탭 정의. 함수로 감싸 호출 시점에 만든다 —
 * 최상위 배열이 다른 파일(Code.js)의 SH 를 로드 시점에 참조하면 "SH is not defined" 로 터진다.
 */
function 탭정의() {
  const 라인ID들 = 라인정의().map((d) => d.라인ID);
  return [
    { 이름: SH.라인, 헤더: ['라인ID', '라인명', '제품', '설비수', '설계능력(톤/일)', '담당자'] },
    { 이름: SH.생산, 헤더: ['일자', '월', '라인ID', '라인명', '교대조', '계획수량', '생산수량', '가동시간(h)', '정지시간(h)', '정지사유', '불량수', '불량유형', '불량률(%)', '작업자', '비고'] },
    { 이름: SH.목표, 헤더: ['월', '목표생산량', '목표불량률(%)'] },
    { 이름: SH.설정, 헤더: ['키', '값', '메모'] },
    { 이름: SH.월별, 헤더: ['월', '계획수량', '생산수량', '달성률(%)', '목표생산량', '목표달성률(%)', '불량수', '불량률(%)', '목표불량률(%)', '목표대비(%p)', '가동시간(h)', '정지시간(h)', '가동률(%)', '판정'] },
    { 이름: SH.라인별, 헤더: ['라인ID', '라인명', '계획수량', '생산수량', '달성률(%)', '불량수', '불량률(%)', '목표불량률(%)', '판정', '가동시간(h)', '정지시간(h)', '가동률(%)'] },
    { 이름: SH.유형별, 헤더: ['불량유형', '불량수', '비율(%)', '누적(%)', '발생건수'] },
    { 이름: SH.교대별, 헤더: ['교대조', '계획수량', '생산수량', '달성률(%)', '불량수', '불량률(%)', '가동시간(h)', '정지시간(h)', '가동률(%)'] },
    { 이름: SH.사유별, 헤더: ['정지사유', '발생건수', '정지시간(h)', '비율(%)'].concat(라인ID들) },
    { 이름: SH.측정, 헤더: ['측정일시', '구분', '작업', '건수', '소요초', '건당초', '방식', '비고'] },
    { 이름: SH.대시보드, 헤더: ['대시보드'] },
    { 이름: SH.리포트, 헤더: ['리포트'] },
    { 이름: SH.안내, 헤더: ['안내'] },
  ];
}

function 설정기본값() {
  return [
    ['회사명', '(주)가상케미칼', '실제 회사명이 아닙니다(포트폴리오용)'],
    ['부서명', '생산1팀', ''],
    ['기간_시작월', '2025-10', '생산일보 첫 달(yyyy-MM). 바꾸면 설치_샘플데이터를 다시 실행'],
    ['기간_개월', 12, '집계 개월 수'],
    ['조회함수', 'XLOOKUP', 'XLOOKUP 또는 INDEXMATCH (구버전 엑셀이면 INDEXMATCH)'],
    ['한번에처리행', 2000, '큰 범위를 나눠 쓰는 단위(6분 제한 대비)'],
    ['시간제한초', 270, '이 시간을 넘으면 다음 실행으로 넘김(샘플 데이터 이어쓰기)'],
    ['목표불량률기본(%)', 1.2, '목표 탭에 값이 없을 때 쓰는 기본 목표'],
    ['계획달성기준(%)', 95, '이 값보다 낮으면 리포트 판정이 미달'],
    ['트리거시각', 7, '월간 트리거 실행 시각(0~23시)'],
    ['리포트_제목', '월간 생산·품질 리포트', ''],
    ['메일알림', '', '리포트 PDF 를 받을 주소(비우면 보내지 않음)'],
    ['합성데이터고지', '이 문서의 모든 데이터는 포트폴리오 시연용 합성(가상) 데이터입니다.', ''],
    ['출력폴더ID', '', '설치_전체가 자동 입력'],
    ['마지막갱신', '', '대시보드갱신이 자동 입력'],
    ['시트버전', '1.0', ''],
  ];
}

/* ============================ 설치 ============================ */

function _설정읽기(ss) {
  const sh = ss.getSheetByName(SH.설정);
  const 값 = {};
  if (!sh || sh.getLastRow() < 2) return 값;
  sh.getDataRange().getValues().slice(1).forEach((r) => {
    const k = String(r[0]).trim();
    if (k) 값[k] = r[1];
  });
  return 값;
}

function _서식헤더(sh, 열수) {
  _안전('헤더 서식', () => {
    sh.getRange(1, 1, 1, 열수).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });
}

/** 탭이 없으면 만들고, 헤더가 다를 때만 헤더를 다시 쓴다(데이터는 건드리지 않는다) */
function _탭보장(ss, 이름, 헤더) {
  let sh = ss.getSheetByName(이름);
  if (!sh) {
    sh = ss.insertSheet(이름);
    sh.getRange(1, 1, 1, 헤더.length).setValues([헤더]);
    _서식헤더(sh, 헤더.length);
    return true;
  }
  if (이름 === SH.대시보드 || 이름 === SH.리포트 || 이름 === SH.안내) return false; // 코드가 통째로 그리는 탭
  const 현재 = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 헤더.length)).getValues()[0].map(String);
  const 같음 = 헤더.every((h, i) => String(현재[i] || '').trim() === h);
  if (!같음) {
    sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 헤더.length)).clearContent();
    sh.getRange(1, 1, 1, 헤더.length).setValues([헤더]);
    _서식헤더(sh, 헤더.length);
    _로그('[주의] ' + 이름 + ' 탭 헤더를 다시 썼습니다(데이터 행은 그대로 두었습니다)');
  }
  return false;
}

function 설치_전체() {
  로그비우기();
  const 시작 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _로그('=== 설치 시작: ' + ss.getName() + ' ===');

  _안전('시간대·로케일', () => {
    ss.setSpreadsheetTimeZone(TZ);
    ss.setSpreadsheetLocale('ko_KR');
  });

  const 만든탭 = [];
  탭정의().forEach((t) => {
    if (_탭보장(ss, t.이름, t.헤더)) 만든탭.push(t.이름);
  });
  _안전('빈 기본 시트 정리', () => {
    ss.getSheets()
      .filter((s) => /^(Sheet|시트)\d*$/.test(s.getName()) && s.getLastRow() === 0)
      .forEach((s) => ss.deleteSheet(s));
  });
  _로그('탭 ' + ss.getSheets().length + '개 (새로 만든 탭: ' + (만든탭.join(', ') || '없음') + ')');

  // 설정: 기존 값은 절대 덮지 않고, 없는 키만 채운다. 값 열은 텍스트 서식(@) — '2025-10' 이 날짜로 바뀌지 않게
  _안전('설정 값 열 텍스트 서식', () => _시트(SH.설정).getRange(2, 2, 60, 1).setNumberFormat('@'));
  const 기존 = _설정읽기(ss);
  let 채운수 = 0;
  설정기본값().forEach(([k, v]) => {
    if (!(k in 기존)) {
      _설정쓰기(k, v);
      채운수++;
    }
  });
  _로그('설정 ' + Object.keys(기존).length + '개 중 ' + 채운수 + '개 채움(기존 값은 유지)');

  // 드라이브 폴더: 루트/출력
  const 루트 = _폴더보장(DriveApp.getRootFolder(), '생산품질대시보드_포트폴리오');
  const 출력 = _폴더보장(루트, '출력');
  if (!String(_설정안전('출력폴더ID', '')).trim()) _설정쓰기('출력폴더ID', 출력.getId());
  _로그('드라이브 폴더: ' + 루트.getName() + ' / 출력');

  _안내작성(ss);
  _안전('측정로그 정리', 정리_측정로그중복);

  // 데이터가 있으면 수식·집계까지, 없으면 대시보드 틀만
  const 생산 = _시트(SH.생산);
  if (생산.getLastRow() >= 2) {
    생산일보서식적용(생산);
    집계갱신();
  } else {
    _로그('생산일보가 비어 있어 수식·집계는 건너뜁니다 (다음: 설치_샘플데이터)');
  }
  대시보드작성(ss);
  _안전('탭 순서', () => _탭순서정리(ss));

  _로그('설치 완료: ' + ((Date.now() - 시작) / 1000).toFixed(1) + '초');
  _로그('※ 다음: ' + (생산.getLastRow() >= 2 ? '메뉴 [대시보드] → 대시보드 갱신' : '설치_샘플데이터 실행'));
  return 로그전체();
}

/** 안내 → 대시보드 → 리포트 → 집계 → 데이터 → 설정 → 측정로그 순으로 정렬 */
function _탭순서정리(ss) {
  const 순서 = [SH.안내, SH.대시보드, SH.리포트, SH.월별, SH.라인별, SH.유형별, SH.교대별, SH.사유별, SH.생산, SH.목표, SH.라인, SH.설정, SH.측정];
  순서.forEach((이름, i) => {
    const sh = ss.getSheetByName(이름);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  const 안내 = ss.getSheetByName(SH.안내);
  if (안내) ss.setActiveSheet(안내);
}

/* ============================ 안내 탭 ============================ */

function _안내작성(ss) {
  const sh = ss.getSheetByName(SH.안내) || ss.insertSheet(SH.안내, 0);
  sh.clear();
  const 줄 = [
    ['생산·품질 대시보드', '제목'],
    ['구글시트 + Apps Script (무료 계정, 부가기능 없음)', '부제'],
    ['', ''],
    ['■ 무엇을 하는 예제인가', '소제목'],
    ['생산일보(일자 × 라인 × 교대조, 5,475행)를 쌓으면 집계 5탭 → KPI 카드·차트 6개 대시보드 → 경영진 리포트 PDF 가 자동으로 만들어집니다.', '본문'],
    ['회사 업무에 실제로 쓰는 자료가 아니라, 자동화 실력을 보여주기 위한 포트폴리오입니다.', '본문'],
    ['데이터는 전부 합성(가상)입니다 — 실제 회사 자료, 라인명, 인명을 쓰지 않았습니다. 예제 회사는 (주)가상케미칼 생산1팀입니다.', '본문'],
    ['', ''],
    ['■ 사용 순서', '소제목'],
    ['1. 설치_전체 — 탭·설정·폴더를 만들고 수식·서식·대시보드 틀을 적용합니다. 두 번 돌려도 데이터가 지워지지 않습니다.', '본문'],
    ['2. 설치_샘플데이터 — 생산일보 5,475행(12개월 × 5라인 × 3교대)과 목표 12행을 만들고 집계·대시보드까지 갱신합니다.', '본문'],
    ['3. 설치_확인 — 12항목을 스스로 검사해 ○/× 로 보여줍니다. 마지막 줄이 "정상 12 / 문제 0" 이면 정상입니다.', '본문'],
    ['4. 메뉴 [대시보드] → 대시보드 갱신 — 행을 추가했거나 목표를 바꿨을 때. 수식 범위를 다시 잡고 차트를 다시 그립니다.', '본문'],
    ['5. 메뉴 [대시보드] → 리포트 PDF 내보내기 — 경영진 보고용 1장을 드라이브 [출력] 폴더에 PDF 로 저장합니다.', '본문'],
    ['6. 메뉴 [대시보드] → 엑셀(.xlsx) 내보내기 — 엑셀에서 열리는 사본(집계 탭은 값으로 굳힘). 차트는 빠집니다.', '본문'],
    ['7. 메뉴 [대시보드] → 월간 트리거 설치 — 매월 1일 아침에 4·5번을 자동 실행합니다.', '본문'],
    ['', ''],
    ['■ 시트 구성', '소제목'],
    ['대시보드 — KPI 카드 6(총생산량·평균 불량률·목표 달성률·평균 가동률·총 정지시간·최다 불량유형, 전월 대비 포함) + 자동 코멘트 + 차트 6', '본문'],
    ['리포트 — 경영진 보고용 1장(핵심 지표·월별·라인별·코멘트·조치 제안·결재란). 값으로 굳혀 PDF 로 내보냅니다.', '본문'],
    ['집계_월별 · 집계_라인별 · 집계_불량유형별(파레토) · 집계_교대조별 · 집계_정지사유별 — SUMIFS/COUNTIFS 수식(엑셀에서도 계산됩니다)', '본문'],
    ['생산일보 — 1행 = 일자 × 라인 × 교대조. 월·라인명·불량률(%)은 수식이 채웁니다. 단위: 포(25kg) / 시간(h)', '본문'],
    ['목표 — 월별 목표 생산량·목표 불량률(%). 바꾸면 대시보드 갱신으로 반영됩니다.', '본문'],
    ['라인마스터 — 라인 5개 기준 정보(라인ID·라인명·제품·설비수·설계능력·담당자)', '본문'],
    ['설정 — 회사명·기간·조회함수·한 번에 처리할 행 수·출력 폴더 등 16개', '본문'],
    ['측정로그 — 자동/수작업 소요 시간 비교 기록', '본문'],
    ['', ''],
    ['■ 참고', '소제목'],
    ['수식은 엑셀 호환이 되는 것만 씁니다(SUMIFS·COUNTIFS·IFERROR·IF·TEXT·XLOOKUP). QUERY·FILTER·ARRAYFORMULA 는 쓰지 않았습니다.', '본문'],
    ['수식 범위는 전체 열(100만 행) 대신 실제 데이터 범위(예: $2:$5476)로 잡아 계산을 가볍게 했습니다. 행을 추가했다면 [대시보드 갱신] 을 한 번 실행하세요.', '본문'],
    ['비율은 전부 % 숫자입니다(1.23 = 1.23%). 불량률 = 불량수 ÷ 생산수량 × 100, 가동률 = 가동시간 ÷ (가동 + 정지) × 100.', '본문'],
    ['차트는 .xlsx 로 내보내면 빠집니다(구글 차트는 엑셀로 변환되지 않음). 그래서 집계 5탭이 통합문서의 본체입니다.', '본문'],
    ['무료 계정 할당량: 스크립트 실행 6분/회. 샘플 데이터는 2,000행씩 나눠 쓰고, 시간이 넘으면 다음 실행에서 이어서 씁니다.', '본문'],
  ];
  const 값 = 줄.map((r) => [r[0]]);
  sh.getRange(1, 1, 값.length, 1).setValues(값);
  줄.forEach((r, i) => {
    const p = sh.getRange(i + 1, 1);
    if (r[1] === '제목') p.setFontSize(16).setFontWeight('bold').setFontColor('#1F3864');
    else if (r[1] === '부제') p.setFontSize(11).setFontColor('#666666');
    else if (r[1] === '소제목') p.setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
    else p.setFontSize(10);
    sh.setRowHeight(i + 1, r[1] === '제목' ? 26 : 18);
  });
  sh.setColumnWidth(1, 720);
  return sh;
}

/* ============================ 샘플 데이터 ============================ */

/** 같은 결과가 나오도록 고정된 난수(재현 가능) */
function _난수기계(seed) {
  let s = seed >>> 0;
  return function () {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 4294967296;
  };
}

/** 가중치 객체에서 하나 고르기 */
function _가중선택(r, 가중) {
  const 키 = Object.keys(가중);
  const 합 = 키.reduce((a, k) => a + 가중[k], 0);
  let x = r() * 합;
  for (let i = 0; i < 키.length; i++) {
    x -= 가중[키[i]];
    if (x < 0) return 키[i];
  }
  return 키[키.length - 1];
}

/**
 * 합성 생산일보를 메모리에서 만든다(시트에 쓰지 않는다).
 * 계절성(겨울 낮음·5~6월 높음) · 주말(토 80%·일 55%) · 야간(C조) 불량률 1.3배 ·
 * 8번째 달의 L3 라인에 이물혼입 급증(불량률 약 4배) — 대시보드가 짚어낼 '이야기'를 심어 둔다.
 */
function _샘플생성(월들) {
  if (!월들 || !월들.length) throw new Error('내부 함수입니다. 설치_샘플데이터 를 실행하세요.');
  const r = _난수기계(20260925);
  const 라인들 = 라인정의();
  const 월계수 = { '01': 0.82, '02': 0.80, '03': 0.90, '04': 0.93, '05': 0.95, '06': 0.95, '07': 0.90, '08': 0.88, '09': 0.94, '10': 0.92, '11': 0.90, '12': 0.85 };
  const 사고 = { 월: 월들[Math.min(7, 월들.length - 1)], 라인: 'L3', 배수: 3.8, 유형: '이물혼입' };
  const 사유가중기본 = { 설비고장: 5, 자재대기: 3, 금형교체: 3, 품질이상: 1, 정전: 0.5, 계획정비: 2, 인원부족: 1.5 };
  const 목표률 = 1.2;
  const 행들 = [];
  const 목표행들 = [];
  const 통계 = { 불량: 0, 생산: 0, 정지건: 0, 사고월: 사고.월, 사고라인: 사고.라인 };

  월들.forEach((월) => {
    const y = Number(월.slice(0, 4));
    const mo = Number(월.slice(5, 7));
    const 일수 = _월일수(월);
    const 계수 = 월계수[월.slice(5, 7)] || 0.9;
    let 목표생산 = 0;
    // 주말 감산(토 80%·일 55%)과 평균 가동률(약 95%)을 감안한 현실적인 목표 — 달성률이 95% 언저리에서 오르내린다
    라인들.forEach((L) => (목표생산 += (L['설계능력(톤/일)'] * 1000 / 25) * 일수 * 계수 * 0.86));
    목표행들.push([월, Math.round(목표생산 / 100) * 100, 목표률]);

    for (let d = 1; d <= 일수; d++) {
      const 날 = new Date(y, mo - 1, d);
      const 요일 = 날.getDay();
      const 요일계수 = 요일 === 0 ? 0.55 : 요일 === 6 ? 0.8 : 1;
      라인들.forEach((L, li) => {
        const 교대능력 = (L['설계능력(톤/일)'] * 1000 / 25) / 3;
        const 사고중 = 월 === 사고.월 && L.라인ID === 사고.라인;
        교대조목록.forEach((교대, si) => {
          const 계획 = Math.round((교대능력 * 계수 * 요일계수 * (0.97 + r() * 0.06)) / 10) * 10;
          const 정지확률 = 사고중 ? 0.5 : 0.28;
          let 정지 = r() < 정지확률 ? Math.round((0.3 + r() * 2.2) * 10) / 10 : 0;
          let 사유 = '';
          if (정지 > 0) {
            const 가중 = Object.assign({}, 사유가중기본);
            if (사고중) 가중.품질이상 = 6;
            사유 = _가중선택(r, 가중);
            통계.정지건++;
          }
          const 가동 = Math.round((8 - 정지) * 10) / 10;
          const 효율 = 0.93 + r() * 0.10;
          const 생산 = Math.max(Math.round(계획 * (가동 / 8) * 효율), 0);
          const 교대계수 = 교대 === 'C' ? 1.3 : 교대 === 'B' ? 1.05 : 1;
          const 률 = (L.불량기본 * 교대계수 * (0.75 + r() * 0.5) * (사고중 ? 사고.배수 : 1)) / 100;
          const 불량 = Math.round(생산 * 률);
          let 유형 = '';
          if (불량 > 0) 유형 = 사고중 && r() < 0.7 ? 사고.유형 : _가중선택(r, L.유형가중);
          let 비고 = '';
          if (사고중 && 불량 > 생산 * 0.03 && r() < 0.35) 비고 = '이물 혼입 집중 점검';
          else if (사유 === '정전') 비고 = '전력 복구 후 재가동';
          else if (사유 === '계획정비') 비고 = '정비 계획 반영';
          통계.불량 += 불량;
          통계.생산 += 생산;
          // 월·라인명·불량률(%) 은 수식 열이므로 비워 둔다(생산일보서식적용이 채운다)
          행들.push([날, '', L.라인ID, '', 교대, 계획, 생산, 가동, 정지, 사유, 불량, 유형, '', 작업자목록[(li * 3 + si) % 작업자목록.length], 비고]);
        });
      });
    }
  });
  return { 행들, 목표행들, 통계 };
}

function 설치_샘플데이터() {
  로그비우기();
  const 시작 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 설치를 건너뛰고 실행해도 "탭이 없습니다" 로 죽지 않게 먼저 확인한다
  if (!ss.getSheetByName(SH.라인) || !ss.getSheetByName(SH.생산) || !ss.getSheetByName(SH.설정) || !ss.getSheetByName(SH.목표)) {
    _로그('[알림] 탭이 아직 없어 설치_전체를 먼저 실행합니다');
    설치_전체();
    로그모음.push('');
  }

  const 월들 = _월목록();
  const 라인들 = 라인정의();
  const 라인헤더 = 탭정의()[0].헤더;
  const { 행들, 목표행들, 통계 } = _샘플생성(월들);
  const 생산탭 = _시트(SH.생산);
  const 시간제한 = Number(_설정안전('시간제한초', 270)) || 270;
  const 묶음 = Math.max(Number(_설정안전('한번에처리행', 2000)) || 2000, 100);
  const p = PropertiesService.getDocumentProperties();

  // 이전 실행이 시간 제한으로 멈췄으면 그 자리부터 이어서 쓴다(생성 결과는 시드가 고정이라 항상 같다)
  let 시작행 = 0;
  const 진행 = Number(p.getProperty('샘플데이터_진행') || 0);
  if (진행 > 0 && 진행 < 행들.length && 생산탭.getLastRow() - 1 >= 진행) {
    시작행 = 진행;
    _로그('[이어서] 이전 실행이 ' + _천단위(진행) + '행에서 멈춰 그 다음부터 씁니다');
  } else {
    // 1) 라인마스터·목표 (작으므로 한 번에)
    const 라인탭 = _시트(SH.라인);
    라인탭.getRange(2, 1, Math.max(라인탭.getLastRow() - 1, 1), 라인탭.getLastColumn()).clearContent();
    라인탭.getRange(2, 1, 라인들.length, 라인헤더.length).setValues(라인들.map((d) => 라인헤더.map((h) => d[h])));
    _로그('가상 라인 ' + 라인들.length + '개 등록 (' + 라인들.map((d) => d.라인ID).join('/') + ')');

    const 목표탭 = _시트(SH.목표);
    목표탭.getRange(2, 1, Math.max(목표탭.getLastRow() - 1, 1), 목표탭.getLastColumn()).clearContent();
    목표탭.getRange(2, 1, 목표행들.length, 1).setNumberFormat('@');
    목표탭.getRange(2, 1, 목표행들.length, 3).setValues(목표행들);
    _안전('목표 숫자 서식', () => {
      목표탭.getRange(2, 2, 목표행들.length, 1).setNumberFormat('#,##0');
      목표탭.getRange(2, 3, 목표행들.length, 1).setNumberFormat('0.00');
    });
    _로그('목표 ' + 목표행들.length + '개월 등록 (목표 불량률 ' + 목표행들[0][2] + '%)');

    // 2) 생산일보 비우기
    const 기존행 = 생산탭.getLastRow() - 1;
    if (기존행 > 0) _재시도('생산일보 비우기', () => 생산탭.getRange(2, 1, 기존행, 생산탭.getLastColumn()).clearContent());
  }

  // 3) 생산일보 쓰기 — 묶음 단위, 시간 제한이 오면 진행 위치를 남기고 멈춘다
  for (let i = 시작행; i < 행들.length; i += 묶음) {
    const 조각 = 행들.slice(i, i + 묶음);
    _재시도('생산일보 쓰기 ' + (i + 1) + '~' + (i + 조각.length), () =>
      생산탭.getRange(2 + i, 1, 조각.length, 조각[0].length).setValues(조각));
    p.setProperty('샘플데이터_진행', String(i + 조각.length));
    const 남음 = 행들.length - (i + 조각.length);
    if (남음 > 0 && (Date.now() - 시작) / 1000 > 시간제한) {
      _로그('[중단] 시간 제한(' + 시간제한 + '초) — ' + _천단위(i + 조각.length) + '행까지 썼습니다. 설치_샘플데이터 를 다시 실행하면 이어서 씁니다.');
      return 로그전체();
    }
  }
  p.deleteProperty('샘플데이터_진행');
  _데이터캐시비우기(); // 방금 쓴 데이터를 다시 읽도록
  _로그('가상 생산일보 ' + _천단위(행들.length) + '행 생성 (' + 월들[0] + ' ~ ' + 월들[월들.length - 1] + ', 라인 ' + 라인들.length + ' × 교대 ' + 교대조목록.length + ')');
  _로그('생산 ' + _천단위(통계.생산) + '포 / 불량 ' + _천단위(통계.불량) + '포 (' + _율(통계.불량, 통계.생산).toFixed(2) + '%) / 정지 ' + _천단위(통계.정지건) + '건 · 급증 시나리오: ' + 통계.사고월 + ' ' + 통계.사고라인);

  // 4) 수식·집계·대시보드
  생산일보서식적용(생산탭);
  집계갱신();
  대시보드작성(ss);
  _설정쓰기('마지막갱신', _지금());
  _측정로그('자동', '샘플 데이터 + 집계 + 대시보드', 행들.length, (Date.now() - 시작) / 1000, '자동', 월들[0] + '~' + 월들[월들.length - 1]);

  _로그('샘플 데이터 완료: ' + ((Date.now() - 시작) / 1000).toFixed(1) + '초');
  _로그('※ 다음: 설치_확인 → 메뉴 [대시보드] → 리포트 PDF 내보내기');
  return 로그전체();
}

/* ============================ 설치 확인 ============================ */

/**
 * 항목별 자체 진단. 무엇이 잘못됐는지 한 줄씩 ○/× 로 보여준다 —
 * 사용자가 로그를 그대로 붙여주면 원인을 바로 찾을 수 있다.
 */
function 설치_확인() {
  로그비우기();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let 통과 = 0;
  let 문제 = 0;

  const 검사 = (이름, fn) => {
    let 결과;
    try {
      결과 = fn();
    } catch (e) {
      결과 = e.message;
    }
    if (결과 === true || 결과 === undefined || 결과 === null) {
      통과++;
      _로그('  ○ ' + 이름);
    } else {
      문제++;
      _로그('  × ' + 이름 + ' — ' + 결과);
    }
  };

  _로그('=== 진단: ' + ss.getName() + ' ===');
  const 탭수 = 탭정의().length;

  검사('탭 ' + 탭수 + '개', () => {
    const 빠진것 = 탭정의().map((t) => t.이름).filter((n) => !ss.getSheetByName(n));
    return 빠진것.length ? '없음: ' + 빠진것.join(', ') + ' → 설치_전체를 실행하세요' : true;
  });

  검사('탭 헤더', () => {
    const 틀린것 = [];
    탭정의().forEach((t) => {
      if (t.이름 === SH.안내 || t.이름 === SH.대시보드 || t.이름 === SH.리포트) return;
      const sh = ss.getSheetByName(t.이름);
      if (!sh) return;
      const 현재 = _헤더행(sh, 1);
      if (t.헤더.some((hd, i) => String(현재[i] || '').trim() !== hd)) 틀린것.push(t.이름);
    });
    return 틀린것.length ? '헤더 다름: ' + 틀린것.join(', ') : true;
  });

  const 라인수 = 라인정의().length;
  검사('라인마스터 ' + 라인수 + '행', () => {
    const sh = ss.getSheetByName(SH.라인);
    const 행 = sh ? sh.getLastRow() - 1 : 0;
    return 행 === 라인수 ? true : 행 + '행 (' + 라인수 + '행이어야 함 → 설치_샘플데이터)';
  });

  const 월들 = _월목록();
  let 기대행 = 0;
  월들.forEach((월) => (기대행 += _월일수(월)));
  기대행 *= 라인수 * 교대조목록.length;
  검사('생산일보 ' + _천단위(기대행) + '행', () => {
    const sh = ss.getSheetByName(SH.생산);
    const 행 = sh ? sh.getLastRow() - 1 : 0;
    if (행 === 기대행) return true;
    const 진행 = PropertiesService.getDocumentProperties().getProperty('샘플데이터_진행');
    return 행 + '행 (' + _천단위(기대행) + '행이어야 함' + (진행 ? ' — 이어쓰기 대기 중, 설치_샘플데이터를 다시 실행' : ' → 설치_샘플데이터') + ')';
  });

  검사('생산일보 수식 3열(월·라인명·불량률)', () => {
    const sh = ss.getSheetByName(SH.생산);
    if (!sh || sh.getLastRow() < 2) return '생산일보가 비어 있음';
    const h = _헤더(sh);
    const 마지막 = sh.getLastRow();
    const 없는것 = ['월', '라인명', '불량률(%)'].filter((n) => {
      const c = h.indexOf(n) + 1;
      if (!c) return true;
      const f = sh.getRange(마지막, c).getFormulas()[0][0];
      return !f;
    });
    return 없는것.length ? '마지막 행에 수식 없음: ' + 없는것.join(', ') + ' (대시보드 갱신을 실행하세요)' : true;
  });

  검사('생산일보 드롭다운 4열', () => {
    const sh = ss.getSheetByName(SH.생산);
    if (!sh || sh.getLastRow() < 2) return '생산일보가 비어 있음';
    const h = _헤더(sh);
    const 있음 = ['라인ID', '교대조', '정지사유', '불량유형'].filter((n) => {
      const c = h.indexOf(n);
      return c >= 0 && sh.getRange(2, c + 1).getDataValidation() !== null;
    });
    return 있음.length === 4 ? true : 있음.length + '/4열만 있음';
  });

  검사('목표 ' + 월들.length + '행 · 월 값이 텍스트', () => {
    const sh = ss.getSheetByName(SH.목표);
    const 행 = sh ? sh.getLastRow() - 1 : 0;
    if (행 !== 월들.length) return 행 + '행 (' + 월들.length + '행이어야 함)';
    const 첫 = sh.getRange(2, 1).getValues()[0][0];
    return typeof 첫 === 'string' ? true : '월 칸이 날짜로 바뀜(' + 첫 + ') → 설치_샘플데이터를 다시 실행';
  });

  검사('집계 수식 결과', () => {
    const sh = ss.getSheetByName(SH.월별);
    if (!sh || sh.getLastRow() < 2) return '집계_월별이 비어 있음 (대시보드 갱신을 실행하세요)';
    const 첫 = sh.getRange(2, 1).getValues()[0][0];
    if (typeof 첫 !== 'string') return '월 칸이 날짜로 바뀜(' + 첫 + ') → 대시보드 갱신';
    const 값 = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
    const 오류 = 값.filter((r) => r.some((x) => String(x).charAt(0) === '#')).length;
    if (오류) return 오류 + '행에 수식 오류(#REF! 등)';
    const 빈탭 = [SH.라인별, SH.유형별, SH.교대별, SH.사유별].filter((n) => (ss.getSheetByName(n) || { getLastRow: () => 0 }).getLastRow() < 2);
    return 빈탭.length ? '비어 있음: ' + 빈탭.join(', ') : true;
  });

  검사('설정 키 ' + 설정기본값().length + '개', () => {
    const 값 = _설정읽기(ss);
    const 빠진것 = 설정기본값().map(([k]) => k).filter((k) => !(k in 값));
    return 빠진것.length ? '없는 키: ' + 빠진것.join(', ') + ' → 설치_전체' : true;
  });

  검사('드라이브 출력 폴더', () => {
    const id = String(_설정안전('출력폴더ID', ''));
    if (!id) return '설정에 폴더 ID가 없음 → 설치_전체를 실행하세요';
    const 이름 = _안전('폴더 확인', () => DriveApp.getFolderById(id).getName());
    return 이름 ? true : '접근 실패: ' + id;
  });

  검사('대시보드 차트 6개 · 코멘트', () => {
    const sh = ss.getSheetByName(SH.대시보드);
    if (!sh) return '대시보드 탭 없음';
    const n = sh.getCharts().length;
    const 코멘트 = String(sh.getRange(_대시보드배치().코멘트행, 1).getValues()[0][0]);
    if (n !== 6) return '차트 ' + n + '개 (6개여야 함 → 대시보드 갱신)';
    if (코멘트.indexOf('데이터가 없습니다') >= 0) return '코멘트가 아직 없음 → 설치_샘플데이터';
    _로그('      · ' + 코멘트.slice(0, 80) + (코멘트.length > 80 ? '…' : ''));
    return true;
  });

  검사('측정로그 (빈 줄·중복 없음)', () => {
    const sh = ss.getSheetByName(SH.측정);
    if (!sh || sh.getLastRow() < 2) return true;
    const v = sh.getDataRange().getValues().slice(1);
    const 빈 = v.filter((r) => !String(r[0]).trim() || !String(r[2]).trim()).length;
    const 본것 = {};
    let 중복 = 0;
    v.forEach((r) => {
      const k = [String(r[0]), r[1], r[2], r[3], r[4]].join('|');
      if (k in 본것) 중복++;
      본것[k] = true;
    });
    v.slice(-3).forEach((r) => _로그('      · ' + r[1] + ' / ' + r[2] + ' / ' + r[3] + '건 / ' + r[4] + '초'));
    return 빈 || 중복 ? '빈 줄 ' + 빈 + ' · 중복 ' + 중복 + ' → 설치_전체를 실행하면 정리됩니다' : true;
  });

  _로그('=== 결과: 정상 ' + 통과 + ' / 문제 ' + 문제 + ' ===');
  if (문제) _로그('※ × 항목을 그대로 복사해서 보내주시면 원인을 잡습니다.');
  return 로그전체();
}

/* ============================ 수작업 측정 ============================ */

function 측정_수작업시작() {
  로그비우기();
  PropertiesService.getDocumentProperties().setProperty('측정_시작', String(Date.now()));
  _로그('수작업 측정 시작. 엑셀로 같은 월별 집계·차트·보고서를 직접 만들어 보세요.');
  _로그('끝나면 측정_수작업종료 를 실행하고 만든 건수(예: 보고서 1건)를 입력합니다.');
  return 로그전체();
}

function 측정_수작업종료(건수) {
  로그비우기();
  const p = PropertiesService.getDocumentProperties();
  const 시작 = Number(p.getProperty('측정_시작') || 0);
  if (!시작) throw new Error('측정_수작업시작을 먼저 실행하세요.');
  const 초 = (Date.now() - 시작) / 1000;
  let n = Number(건수);
  if (!n) {
    // 편집기 실행은 인수를 못 넣으므로 물어본다(대화상자가 안 되면 1건으로 기록)
    n = 1;
    try {
      const ui = SpreadsheetApp.getUi();
      const 답 = ui.prompt('수작업 측정', '직접 만든 집계·보고서가 몇 건인가요?', ui.ButtonSet.OK_CANCEL);
      if (답.getSelectedButton() === ui.Button.OK) n = Number(답.getResponseText()) || 1;
    } catch (e) {}
  }
  _측정로그('수작업', '집계·보고서 작성', n, 초, '수작업', '엑셀 수작업 기준');
  p.deleteProperty('측정_시작');
  _로그('수작업 ' + n + '건 / ' + 초.toFixed(1) + '초 / 건당 ' + (초 / n).toFixed(1) + '초');
  return 로그전체();
}

/* ============================ 초기화 ============================ */

function 초기화_출력폴더() {
  로그비우기();
  const 폴더 = _출력폴더();
  let n = 0;
  const 파일들 = 폴더.getFiles();
  while (파일들.hasNext()) {
    파일들.next().setTrashed(true);
    n++;
  }
  _로그('출력 폴더 파일 ' + n + '개를 휴지통으로 보냈습니다(측정로그는 그대로 둡니다)');
  return 로그전체();
}

/* ============================ 메뉴 ============================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('대시보드')
    .addItem('① 전체 설치', '메뉴_설치전체')
    .addItem('② 샘플 데이터 생성', '메뉴_샘플데이터')
    .addSeparator()
    .addItem('대시보드 갱신', '메뉴_대시보드갱신')
    .addItem('리포트 PDF 내보내기', '메뉴_리포트PDF')
    .addItem('엑셀(.xlsx) 내보내기', '메뉴_엑셀내보내기')
    .addSeparator()
    .addItem('월간 트리거 설치', '메뉴_트리거설치')
    .addItem('월간 트리거 해제', '메뉴_트리거해제')
    .addSeparator()
    .addItem('수작업 측정 시작', '메뉴_수작업시작')
    .addItem('수작업 측정 종료', '메뉴_수작업종료')
    .addItem('설치 상태 확인', '메뉴_설치확인')
    .addToUi();
}

/** 메뉴 실행 공통. 오류가 나면 이유를 대화상자로 보여준다(원시 예외 문구만 던지지 않는다) */
function _메뉴실행(제목, 일) {
  로그비우기();
  try {
    const 본문 = 일();
    _알림(제목 + ' 완료', 본문 || _마지막줄(8));
  } catch (e) {
    _로그('[오류] ' + e.message);
    _알림(제목 + ' 오류', e.message + '\n\n' + _마지막줄(6));
  }
}

function _알림(제목, 본문) {
  const 글 = String(본문 === null || 본문 === undefined ? '' : 본문);
  try {
    SpreadsheetApp.getUi().alert(제목, 글, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(제목 + ' / ' + 글);
  }
  return 글;
}

/** 설치·생성처럼 시간이 걸리는 일은 실행 로그 대신 마지막 몇 줄만 보여준다 */
function _마지막줄(n) {
  return 로그전체().split('\n').slice(-Number(n || 8)).join('\n');
}

function 메뉴_설치전체() {
  _메뉴실행('설치', () => {
    설치_전체();
    return _마지막줄(8);
  });
}

function 메뉴_샘플데이터() {
  let 답 = 'YES';
  try {
    답 = SpreadsheetApp.getUi().alert('샘플 데이터', '기존 생산일보·목표·라인마스터를 지우고 합성 데이터 5,475행을 다시 만듭니다. 계속할까요?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO);
  } catch (e) {}
  // V8 에서 alert(title, message, buttonSet) 은 'YES'/'NO' 문자열을 돌려준다
  const 답글 = String(답).toUpperCase();
  if (답글.indexOf('YES') < 0 && 답글.indexOf('예') < 0) return;
  _메뉴실행('샘플 데이터', () => {
    설치_샘플데이터();
    return _마지막줄(8);
  });
}

function 메뉴_대시보드갱신() {
  _메뉴실행('대시보드 갱신', () => {
    대시보드갱신();
    return _마지막줄(5);
  });
}

function 메뉴_리포트PDF() {
  _메뉴실행('리포트 PDF', () => {
    const url = 리포트_PDF();
    return _마지막줄(4) + '\n\n' + url;
  });
}

function 메뉴_엑셀내보내기() {
  _메뉴실행('엑셀 내보내기', () => {
    const url = 엑셀내보내기();
    return _마지막줄(3) + '\n\n' + url;
  });
}

function 메뉴_트리거설치() {
  _메뉴실행('월간 트리거', () => 월간트리거설치());
}

function 메뉴_트리거해제() {
  _메뉴실행('월간 트리거 해제', () => 월간트리거해제());
}

function 메뉴_수작업시작() {
  _메뉴실행('수작업 측정', () => 측정_수작업시작());
}

function 메뉴_수작업종료() {
  _메뉴실행('수작업 측정', () => 측정_수작업종료());
}

function 메뉴_설치확인() {
  _메뉴실행('진단', () => 설치_확인());
}


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

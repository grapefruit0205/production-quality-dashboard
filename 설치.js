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

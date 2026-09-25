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

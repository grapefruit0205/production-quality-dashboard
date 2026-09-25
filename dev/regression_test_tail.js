/**
 * 회귀 테스트(주입 실패) — stub_test.js 가 본 흐름을 다 돌린 뒤 부른다.
 *   1. 일시적 시트 오류(Service Spreadsheets timed out) 1회 → 재시도로 넘어가야 한다
 *   2. 계속 실패 → 오류를 삼키지 않고 그대로 올려야 한다
 *   3. 측정로그에 빈 줄·중복 줄 주입 → 설치_전체 가 자동 정리해야 한다
 *   4. 필터가 이미 있는 시트에 createFilter → 서식 적용이 [건너뜀] 없이 지나가야 한다(제거 후 재생성)
 *   5. 내부 함수를 인수 없이 실행 → 친절한 한글 안내 예외
 *   6. 6분 제한 시뮬레이션 → [중단] 후 재실행하면 [이어서] 로 5,475행 완성
 *   7. 목표 변경 → 대시보드갱신 후 코멘트 반영
 */
module.exports = function ({ sandbox, ss, 탭, 실행, 단언, 로그포함, Sheet, properties, 마지막로그 }) {
  const 생산 = 탭('생산일보');

  /* ===== 1. 일시적 오류 1회 → 재시도 ===== */
  const 원본_setValues = sandbox.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('집계_월별').getRange(1, 1).constructor.prototype.setValues;
  const RangeProto = Object.getPrototypeOf(탭('집계_월별').getRange(1, 1));
  let 남은실패 = 1;
  RangeProto.setValues = function (v) {
    if (this.sheet.name === '집계_월별' && this.nr > 1 && 남은실패 > 0) {
      남은실패--;
      throw new Error('Service Spreadsheets timed out while accessing document with id TEST.');
    }
    return 원본_setValues.call(this, v);
  };
  sandbox._데이터캐시비우기();
  실행('회귀1: 집계_월별 쓰기 1회 타임아웃 주입 → 집계갱신', () => sandbox.집계갱신());
  단언(남은실패 === 0 && 로그포함('[재시도 1/2] 집계_월별 쓰기: Service Spreadsheets timed out'), '타임아웃 1회를 재시도로 넘김');
  단언(로그포함('집계: 월별 12행'), '재시도 뒤 집계 완료');
  RangeProto.setValues = 원본_setValues;

  /* ===== 2. 계속 실패 → 오류 그대로 ===== */
  let 계속실패 = 99;
  RangeProto.setValues = function (v) {
    if (this.sheet.name === '집계_월별' && this.nr > 1 && 계속실패 > 0) { 계속실패--; throw new Error('Service Spreadsheets timed out.'); }
    return 원본_setValues.call(this, v);
  };
  let 올라온오류 = '(없음 — 오류를 삼켰다)';
  try { sandbox.집계갱신(); } catch (e) { 올라온오류 = e.message; }
  단언(/timed out/.test(올라온오류) && 계속실패 === 96, '3회 다 실패하면 오류를 그대로 올림 (시도 3회): ' + 올라온오류);
  RangeProto.setValues = 원본_setValues;
  sandbox._데이터캐시비우기();

  /* ===== 3. 측정로그 빈 줄·중복 → 설치_전체 가 정리 ===== */
  const 측정 = 탭('측정로그');
  const 전건수 = 측정.getLastRow() - 1;
  const 첫행 = 측정.getRange(2, 1, 1, 8).getValues()[0];
  측정.appendRow(['', '자동', '', 3, '1.0', '', '자동', '']);   // 측정일시가 빈 줄(인수 없이 실행된 흔적)
  측정.appendRow([new Date(), '', '', 0, '0.0', '', '', '']);   // 작업이 빈 줄
  측정.appendRow(첫행);                                        // 완전 중복
  console.log('\n[테스트] 주입 후 측정로그 = ' + (측정.getLastRow() - 1) + '행 (기대 ' + (전건수 + 3) + ')');
  실행('회귀3: 설치_전체 (측정로그 자동 정리)', () => sandbox.설치_전체());
  단언(로그포함('측정로그 정리: 중복 1행 / 빈 기록 2행 삭제'), '중복 1 · 빈 줄 2 정리 로그');
  단언(측정.getLastRow() - 1 === 전건수, '정리 후 측정로그 ' + (측정.getLastRow() - 1) + '행 (기대 ' + 전건수 + ')');

  /* ===== 4. 필터 중복 거부 ===== */
  let 필터오류 = '';
  try { 생산.getRange(1, 1, 10, 15).createFilter(); } catch (e) { 필터오류 = e.message; }
  단언(/already has a filter/.test(필터오류), '하네스가 필터 중복 생성을 거부함(실제 API 재현)');
  실행('회귀4: 필터 있는 상태에서 생산일보서식적용 재실행', () => sandbox.생산일보서식적용(생산));
  단언(!로그포함('[건너뜀] 필터') && 로그포함('생산일보 서식: 수식 16,425개'), '기존 필터 제거 후 재생성 — [건너뜀] 없음');

  /* ===== 5. 내부 함수 인수 없이 실행 → 친절한 안내 ===== */
  const 안내검사 = (이름, fn) => {
    let m = '(예외 없음)';
    try { fn(); } catch (e) { m = e.message; }
    단언(/메뉴|실행하세요/.test(m) && !/undefined|null/.test(m), 이름 + '() → 안내 예외: ' + m.slice(0, 60));
  };
  안내검사('_차트만들기', () => sandbox._차트만들기());
  안내검사('_차트전부', () => sandbox._차트전부());
  안내검사('_내보내기', () => sandbox._내보내기());
  안내검사('_리포트작성', () => sandbox._리포트작성());
  안내검사('_종합코멘트', () => sandbox._종합코멘트());
  안내검사('_조치제안', () => sandbox._조치제안());
  안내검사('_샘플생성', () => sandbox._샘플생성());
  안내검사('_차트전부삭제', () => sandbox._차트전부삭제());

  /* ===== 6. 6분 제한 시뮬레이션 → 이어쓰기 ===== */
  sandbox._설정쓰기('시간제한초', 0.001);   // 첫 묶음 뒤 바로 멈춘다(0 은 '비어 있음' 으로 보고 기본값 270 을 쓴다)
  sandbox._설정쓰기('한번에처리행', 2000);
  실행('회귀6a: 시간제한 0.001초 → 샘플데이터 중단', () => sandbox.설치_샘플데이터());
  단언(로그포함('[중단] 시간 제한(0.001초) — 2,000행까지 썼습니다'), '2,000행 뒤 중단 로그');
  단언(properties['샘플데이터_진행'] === '2000', '진행 속성 2000 저장');
  단언(생산.getLastRow() - 1 === 2000, '생산일보 2,000행만 있음');
  sandbox._설정쓰기('시간제한초', 270);
  실행('회귀6b: 재실행 → 이어쓰기', () => sandbox.설치_샘플데이터());
  단언(로그포함('[이어서] 이전 실행이 2,000행에서 멈춰'), '[이어서] 로그');
  단언(생산.getLastRow() - 1 === 5475 && !properties['샘플데이터_진행'], '5,475행 완성 · 진행 속성 삭제');
  단언(탭('대시보드').getCharts().length === 6, '이어쓰기 후 차트 6개');

  /* ===== 7. 목표 변경 반영 ===== */
  const 목표 = 탭('목표');
  목표.getRange(2, 3, 12, 1).setValues(Array.from({ length: 12 }, () => [0.5]));   // 목표를 0.5% 로 낮춤
  sandbox._설정쓰기('목표불량률기본(%)', 0.5);
  실행('회귀7: 목표 0.5% 로 낮춘 뒤 대시보드갱신', () => sandbox.대시보드갱신());
  const 코멘트 = String(탭('대시보드').getRange(10, 1).getValues()[0][0]);
  단언(코멘트.indexOf('목표 0.50% 초과') >= 0, '코멘트가 새 목표(0.50%) 기준으로 초과 판정');
  단언(코멘트.indexOf('목표 불량률(0.50%)을 넘긴 라인: L4') >= 0, '초과 라인 목록 갱신');
  // 원상 복구
  목표.getRange(2, 3, 12, 1).setValues(Array.from({ length: 12 }, () => [1.2]));
  sandbox._설정쓰기('목표불량률기본(%)', 1.2);
  실행('회귀7 복구: 대시보드갱신', () => sandbox.대시보드갱신());

  /* ===== 최종 진단 ===== */
  실행('설치_확인 (회귀 테스트 후)', () => sandbox.설치_확인());
  단언(마지막로그().trim().split('\n').pop() === '=== 결과: 정상 12 / 문제 0 ===', '회귀 테스트 후에도 설치_확인 12/0');
};

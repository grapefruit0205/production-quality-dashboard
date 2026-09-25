/**
 * 생산·품질 대시보드 로컬 검증 하네스 (개발용 — Apps Script 에 올리지 않는다)
 * Apps Script 런타임을 최소한으로 흉내 내서
 *   설치_전체 → 설치_샘플데이터 → 설치_전체(재실행) → 대시보드갱신 → 리포트_PDF → 엑셀내보내기 → 트리거 → 설치_확인
 * 까지 돌리고, 로그 줄과 결과 숫자를 단언(assert)한다.
 *
 * 목적: 구글에 올리기 전에 오타·미정의 변수·인덱스 오류·로드 순서 문제를 잡는 것.
 * 실제 API 의 제약(필터 중복 거부·requireValueInRange 거부·xlsx getAs 거부)을 일부러 재현해 회귀를 막는다.
 *
 * 실행:  node dev/stub_test.js      (종료 코드 0 = 전부 통과)
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');

const ROOT = path.join(__dirname, '..');

/* ---------- 단언 ---------- */
let 단언통과 = 0;
let 단언실패 = 0;
function 단언(조건, 메시지) {
  if (조건) { 단언통과++; console.log('  ✔ ' + 메시지); }
  else { 단언실패++; console.log('  ✘ ' + 메시지); process.exitCode = 1; }
}

/* ---------- 시트 모형 ---------- */
class Range {
  constructor(sheet, row, col, nr, nc) {
    this.sheet = sheet; this.row = row; this.col = col; this.nr = nr; this.nc = nc;
    sheet.touch(row + nr - 1, col + nc - 1);
  }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getA1Notation() { return 'R' + this.row + 'C' + this.col + ':R' + (this.row + this.nr - 1) + 'C' + (this.col + this.nc - 1); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nr; r++) {
      const line = [];
      for (let c = 0; c < this.nc; c++) {
        const v = this.sheet.get(this.row + r, this.col + c);
        line.push(v === undefined ? '' : v);
      }
      out.push(line);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(v) {
    if (v.length !== this.nr || (v[0] || []).length !== this.nc) {
      throw new Error('The number of rows or columns in the data does not match (range ' + this.nr + 'x' + this.nc + ', data ' + v.length + 'x' + (v[0] || []).length + ')');
    }
    v.forEach((line, r) => line.forEach((x, c) => this.sheet.set(this.row + r, this.col + c, x)));
    return this;
  }
  setValue(x) { this.sheet.set(this.row, this.col, x); return this; }
  setFormula(f) { return this.setValue(f); }
  setFormulas(v) { return this.setValues(v); }
  getFormulas() { return this.getValues().map((r) => r.map((x) => (typeof x === 'string' && x.charAt(0) === '=' ? x : ''))); }
  getFormula() { return this.getFormulas()[0][0]; }
  clearContent() {
    for (let r = 0; r < this.nr; r++) for (let c = 0; c < this.nc; c++) this.sheet.set(this.row + r, this.col + c, '');
    return this;
  }
  getNumRows() { return this.nr; }
  getNumColumns() { return this.nc; }
  setDataValidation(v) {
    for (let r = 0; r < this.nr; r++) for (let c = 0; c < this.nc; c++) this.sheet.validations[(this.row + r) + ':' + (this.col + c)] = v;
    return this;
  }
  getDataValidation() { return this.sheet.validations[this.row + ':' + this.col] || null; }
  getDisplayValues() { return this.getValues().map((r) => r.map((x) => String(x))); }
  merge() { return this; }
}
class Sheet {
  constructor(name, ss) {
    this.name = name; this.ss = ss; this.grid = {}; this.maxRow = 0; this.maxCol = 0; this.frozenRows = 0; this.validations = {}; this.rowHeights = {};
    this.charts = []; this.id = Math.floor(Math.random() * 1e9); this.cfRules = [];
  }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getParent() { return this.ss; }
  get(r, c) { const k = r + ':' + c; return this.grid[k] === undefined ? '' : this.grid[k]; }
  set(r, c, v) { this.grid[r + ':' + c] = (v === null || v === undefined) ? '' : v; this.touch(r, c); }
  touch(r, c) { this.maxRow = Math.max(this.maxRow, r); this.maxCol = Math.max(this.maxCol, c); }
  getLastRow() {
    for (let r = this.maxRow; r >= 1; r--) {
      for (let c = 1; c <= this.maxCol; c++) if (String(this.get(r, c)).trim() !== '') return r;
    }
    return 0;
  }
  getLastColumn() {
    for (let c = this.maxCol; c >= 1; c--) {
      for (let r = 1; r <= Math.min(this.maxRow, 3); r++) if (String(this.get(r, c)).trim() !== '') return c;
    }
    return 0;
  }
  getRange(r, c, nr, nc) {
    if (typeof r === 'string') throw new Error('하네스는 A1 표기를 지원하지 않습니다: ' + r);
    return new Range(this, r, c, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc);
  }
  getDataRange() { return new Range(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
  appendRow(v) { const r = this.getLastRow() + 1; v.forEach((x, i) => this.set(r, i + 1, x)); }
  deleteRow(r) { this.deleteRows(r, 1); }
  deleteRows(r, n) {
    const 마지막 = this.maxRow;
    for (let i = r; i <= 마지막; i++) {
      for (let c = 1; c <= this.maxCol; c++) this.set(i, c, (i + n <= 마지막) ? this.get(i + n, c) : '');
    }
  }
  getMaxRows() { return Math.max(this.maxRow, 1000); }
  clear() { this.grid = {}; this.maxRow = 0; this.maxCol = 0; this.validations = {}; return this; }
  setName(n) { this.name = n; }
  setFrozenRows(n) { this.frozenRows = n; }
  setFrozenColumns() {}
  setColumnWidth() {}
  setRowHeight(r, h) { this.rowHeights[r] = h; }
  setRowHeights(r, n, h) { for (let i = r; i < r + n; i++) this.rowHeights[i] = h; }
  setHiddenGridlines() {}
  setTabColor() {}
  setConditionalFormatRules(r) { this.cfRules = r; return this; }
  getConditionalFormatRules() { return this.cfRules; }
  clearConditionalFormatRules() { this.cfRules = []; return this; }
  // 차트: 실제 API 처럼 시트에 붙고, 지우기 전에는 계속 쌓인다
  newChart() { return new ChartBuilder(this); }
  insertChart(c) { this.charts.push(c); }
  removeChart(c) { this.charts = this.charts.filter((x) => x !== c); }
  getCharts() { return this.charts.slice(); }
}
class ChartBuilder {
  constructor(sheet) { this.sheet = sheet; this.ranges = []; this.options = {}; this.type = null; }
  setChartType(t) { if (!t) throw new Error('알 수 없는 차트 종류'); this.type = t; return this; }
  addRange(r) { if (!(r instanceof Range)) throw new Error('addRange 에는 Range 를 넘겨야 합니다'); this.ranges.push(r); return this; }
  setNumHeaders() { return this; }
  setPosition(r, c) { this.pos = [r, c]; return this; }
  setOption(k, v) { this.options[k] = v; return this; }
  build() { if (!this.ranges.length) throw new Error('범위 없는 차트'); return { type: this.type, ranges: this.ranges, options: this.options, pos: this.pos }; }
}
class Spreadsheet {
  constructor(name) { this.name = name; this.sheets = [new Sheet('Sheet1', this)]; this.active = null; }
  getId() { return 'SS-' + this.name; }
  getName() { return this.name; }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(n) { return this.sheets.filter((s) => s.getName() === n)[0] || null; }
  insertSheet(n, i) { const s = new Sheet(n, this); if (i === 0) this.sheets.unshift(s); else this.sheets.push(s); return s; }
  deleteSheet(s) { if (this.sheets.length <= 1) throw new Error('마지막 시트는 지울 수 없습니다'); this.sheets = this.sheets.filter((x) => x !== s); }
  setSpreadsheetTimeZone() {}
  setSpreadsheetLocale() {}
  setActiveSheet(s) { this.active = s; }
  moveActiveSheet(pos) {
    const s = this.active; if (!s) return;
    this.sheets = this.sheets.filter((x) => x !== s);
    this.sheets.splice(Math.min(pos - 1, this.sheets.length), 0, s);
  }
}
// 서식 메서드는 전부 무시하고 체이닝만 되면 된다
const 무시 = ['setFontWeight', 'setBackground', 'setFontColor', 'setFontSize', 'setHorizontalAlignment',
  'setBorderColor', 'setBorderWidth', 'setBorder', 'setPaddingTop', 'setPaddingBottom', 'setPaddingLeft', 'setPaddingRight',
  'setVerticalAlignment', 'setNumberFormat', 'setWrap', 'setNote', 'setBold', 'setItalic', 'setFontFamily', 'setFontLine', 'setFontStyle'];
무시.forEach((m) => {
  Range.prototype[m] = function () { return this; };
});

// 실제 API 는 '필터가 이미 있는 시트' 에 createFilter 를 또 부르면 오류를 낸다 — 하네스도 같은 제약을 흉내 낸다
Range.prototype.createFilter = function () {
  if (this.sheet.filter) throw new Error("You can't create a filter in a sheet that already has a filter.");
  this.sheet.filter = { remove: () => { this.sheet.filter = null; } };
  return this;
};
Sheet.prototype.getFilter = function () { return this.filter || null; };

/* ---------- 드라이브 모형 ---------- */
const 파일저장소 = {};
class Folder {
  constructor(id, name) { this.id = id; this.name = name; this.files = []; this.folders = []; }
  getId() { return this.id; }
  getName() { return this.name; }
  getFoldersByName(n) { const f = this.folders.filter((x) => x.name === n); let i = 0; return { hasNext: () => i < f.length, next: () => f[i++] }; }
  createFolder(n) { const f = new Folder('F' + Math.random().toString(36).slice(2, 8), n); this.folders.push(f); return f; }
  createFile(blob) { const f = new FakeFile('FILE-' + Math.random().toString(36).slice(2, 8), blob.getName ? blob.getName() : 'x'); this.files.push(f); return f; }
  getFiles() { let i = 0; const self = this; return { hasNext: () => i < self.files.length, next: () => self.files[i++] }; }
  addFile(f) { this.files.push(f); return this; }
  removeFile(f) { this.files = this.files.filter((x) => x !== f); return this; }
}
class FakeFile {
  constructor(id, name, 원본시트) { this.id = id; this.name = name; this.trashed = false; this.원본시트 = 원본시트; 파일저장소[id] = this; }
  makeCopy(name) { return new FakeFile(시트복사(this.원본시트, name).getId(), name); }
  getId() { return this.id; }
  getName() { return this.name; }
  getMimeType() { return 'application/vnd.google-apps.spreadsheet'; }
  getUrl() { return 'https://drive/' + this.id; }
  getAs(mime) {
    const m = String(mime);
    // ②에서 실제로 났던 제약: 시트 → xlsx 변환이 getAs 로는 막힌다(다운로드 URL 경로로 넘어가야 한다)
    if (m === 'excel' || m.indexOf('spreadsheetml') >= 0) {
      throw new Error('Converting from application/vnd.google-apps.spreadsheet to ' + m + ' is not supported.');
    }
    // PDF 도 계정에 따라 막힌다 — 회귀 테스트에서 스위치로 켠다(스텁 전역)
    if (mime === 'application/pdf' && globalThis.__pdfGetAs거부) {
      throw new Error('Converting from application/vnd.google-apps.spreadsheet to application/pdf is not supported.');
    }
    return { mime: m, name: this.name, setName(n) { this.name = n; return this; }, getName() { return this.name; } };
  }
  setName(n) { this.name = n; return this; }
  setTrashed() { this.trashed = true; return this; }
}
const myDrive = new Folder('ROOT', 'My Drive');
const 복사원본 = {};
const 시트복사 = (원본, 이름) => {
  const 새것 = new Spreadsheet(이름);
  새것.sheets = 원본.getSheets().map((s) => {
    const c = new Sheet(s.getName(), 새것);
    Object.keys(s.grid).forEach((k) => (c.grid[k] = s.grid[k]));
    c.maxRow = s.maxRow; c.maxCol = s.maxCol;
    c.charts = s.charts.slice();
    return c;
  });
  복사원본[새것.getId()] = 새것;
  파일저장소[새것.getId()] = new FakeFile(새것.getId(), 이름, 새것);
  return 새것;
};

/* ---------- 기타 서비스 ---------- */
const Utils = {
  formatDate: (d, tz, fmt) => {
    const p = (n) => ('0' + n).slice(-2);
    if (!d || typeof d.getFullYear !== 'function') return '';
    return fmt.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()))
      .replace('HH', p(d.getHours())).replace('mm', p(d.getMinutes())).replace('ss', p(d.getSeconds()));
  },
  sleep: () => {},
};
const Logger = { log: () => {} };
const properties = {};
const PropertiesService = {
  getDocumentProperties: () => ({
    setProperty: (k, v) => (properties[k] = v),
    getProperty: (k) => (k in properties ? properties[k] : null),
    deleteProperty: (k) => delete properties[k],
  }),
};
const 트리거들 = [];
const ScriptApp = {
  getOAuthToken: () => 'fake-token',
  getProjectTriggers: () => 트리거들.slice(),
  deleteTrigger: (t) => { const i = 트리거들.indexOf(t); if (i >= 0) 트리거들.splice(i, 1); },
  newTrigger: (fn) => ({
    timeBased: () => ({
      onMonthDay: (d) => ({ atHour: (h) => ({ create: () => { const t = { getHandlerFunction: () => fn, 일: d, 시: h }; 트리거들.push(t); return t; } }) }),
    }),
  }),
};
const 보낸메일 = [];
const MailApp = { sendEmail: (to, subject, body) => 보낸메일.push({ to, subject, body }) };

/* ---------- 가짜 시계 ----------
 * 실제 실행은 몇 초씩 걸리지만 하네스는 0.1초 안에 끝나 측정로그 시각이 전부 같아진다.
 * 부를 때마다 1초씩 앞으로 가는 Date 를 넣어 실제와 같은 간격을 흉내 낸다. */
let 시계오프셋 = 0;
class FakeDate extends Date {
  constructor(...a) {
    if (a.length) super(...a);
    else { 시계오프셋 += 1000; super(Date.now() + 시계오프셋); }
  }
  static now() { 시계오프셋 += 1000; return Date.now() + 시계오프셋; }
}

/* ---------- 컨텍스트 ---------- */
let 케이스 = null;
const sandbox = {
  console,
  Date: FakeDate,
  Logger,
  Utilities: Utils,
  PropertiesService,
  ScriptApp,
  MailApp,
  SpreadsheetApp: {
    getActiveSpreadsheet: () => 케이스.ss,
    openById: (id) => 복사원본[id],
    getUi: () => { throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); },
    newDataValidation: () => {
      const b = {};
      // 실제 런타임은 다른 스프레드시트의 범위를 참조하면 예외를 던진다(②의 내보낸 대장에서 터졌던 원인)
      b.requireValueInRange = () => { throw new Error('데이터 검증 범위는 같은 스프레드시트 안에서만 참조할 수 있습니다'); };
      ['requireValueInList', 'setAllowInvalid', 'setHelpText'].forEach((m) => (b[m] = () => b));
      b.build = () => ({ 종류: '검증' });
      return b;
    },
    newConditionalFormatRule: () => {
      const b = {};
      ['whenTextEqualTo', 'whenFormulaSatisfied', 'whenNumberGreaterThan', 'whenNumberLessThan', 'setBackground', 'setFontColor', 'setBold', 'setRanges',
        'setGradientMinpointWithValue', 'setGradientMaxpointWithValue'].forEach((m) => (b[m] = () => b));
      b.whenFormulaSatisfied = (f) => {
        // 실제 API: 조건부 서식 수식은 다른 시트를 참조할 수 없다
        if (/[가-힣A-Za-z_]+!/.test(String(f))) throw new Error('조건부 서식 수식은 다른 시트를 참조할 수 없습니다: ' + f);
        return b;
      };
      b.build = () => ({});
      return b;
    },
    InterpolationType: { NUMBER: 'NUMBER' },
    BorderStyle: { SOLID: 'SOLID' },
    flush: () => {},
  },
  Charts: { ChartType: { COLUMN: 'COLUMN', LINE: 'LINE', COMBO: 'COMBO', BAR: 'BAR', PIE: 'PIE' } },
  DriveApp: {
    getRootFolder: () => myDrive,
    getFolderById: (id) => {
      const 찾기 = (f) => {
        if (f.id === id) return f;
        for (const g of f.folders) { const r = 찾기(g); if (r) return r; }
        return null;
      };
      return 찾기(myDrive);
    },
    getFileById: (id) => 파일저장소[id],
  },
  UrlFetchApp: {
    fetch: (url) => {
      if (String(url).indexOf('/export?format=') < 0) throw new Error('예상 밖 URL: ' + url);
      내보내기URL들.push(String(url));
      const 형식 = String(url).split('format=')[1].split('&')[0];
      return {
        getResponseCode: () => 200,
        getBlob: () => ({ name: 'blob.' + 형식, setName(n) { this.name = n; return this; }, getName() { return this.name; } }),
      };
    },
  },
  MimeType: { MICROSOFT_EXCEL: 'excel', PDF: 'application/pdf' },
  Sheets: undefined,
};
sandbox.global = sandbox;
vm.createContext(sandbox);

// 최악의 로드 순서: 설치.js → 대시보드.js → Code.js (SH 를 정의하는 Code.js 가 마지막)
// ONEFILE=1 이면 붙여넣기용 합본(설치_한번에.js) 하나만 로드해 같은 검증을 돌린다
const 로드순서 = process.env.ONEFILE ? ['설치_한번에.js'] : ['설치.js', '대시보드.js', 'Code.js'];
로드순서.forEach((f) => {
  vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f });
});
console.log('로드 순서(최악): ' + 로드순서.join(' → ') + ' — 최상위 참조 오류 없음');
// 최상위 const 는 컨텍스트의 렉시컬 바인딩이라 sandbox.SH 로는 안 보인다 — 컨텍스트 안에서 확인한다
단언(vm.runInContext('typeof SH', sandbox) === 'object' && typeof sandbox.탭정의 === 'function', '파일 3개 로드 (최상위 교차 참조 없음)');

/* ---------- 실행 도구 ---------- */
const 내보내기URL들 = [];
let 마지막로그 = '';
function 실행(이름, fn) {
  process.stdout.write('\n========== ' + 이름 + ' ==========\n');
  let 반환;
  try {
    반환 = fn();
    마지막로그 = sandbox.로그전체();
    console.log(마지막로그.split('\n').slice(-30).join('\n'));
  } catch (e) {
    마지막로그 = sandbox.로그전체() + '\n!! 오류: ' + e.message;
    console.log('!! 오류: ' + e.message + '\n' + (e.stack || '').split('\n').slice(0, 6).join('\n'));
    단언(false, 이름 + ' 이 오류 없이 끝남');
  }
  return 반환;
}
const 로그포함 = (문구) => 마지막로그.indexOf(문구) >= 0;

케이스 = { ss: new Spreadsheet('생산품질_대시보드') };
파일저장소[케이스.ss.getId()] = new FakeFile(케이스.ss.getId(), '생산품질_대시보드', 케이스.ss);
const ss = 케이스.ss;
const 탭 = (n) => ss.getSheetByName(n);

/* ========== 1. 설치_전체 (빈 시트) ========== */
실행('설치_전체 (빈 시트)', () => sandbox.설치_전체());
단언(로그포함('설치 완료'), '설치 완료 로그');
단언(로그포함('생산일보가 비어 있어 수식·집계는 건너뜁니다'), '데이터 없을 때 집계 건너뜀 안내');
단언(ss.getSheets().length === 13, '탭 13개 (실제 ' + ss.getSheets().length + ')');
단언(탭('설정').getLastRow() - 1 === 16, '설정 키 16개');
단언(String(탭('설정').getDataRange().getValues().find((r) => r[0] === '출력폴더ID')[1]).length > 0, '출력폴더ID 자동 입력');
단언(ss.getSheets()[0].getName() === '안내' && ss.getSheets()[1].getName() === '대시보드', '탭 순서: 안내 → 대시보드');
단언(로그포함('차트: 데이터가 없어 건너뜁니다'), '데이터 없을 때 차트 건너뜀');

/* ========== 2. 설치_샘플데이터 ========== */
실행('설치_샘플데이터', () => sandbox.설치_샘플데이터());
const 생산 = 탭('생산일보');
단언(생산.getLastRow() - 1 === 5475, '생산일보 5,475행 (실제 ' + (생산.getLastRow() - 1) + ')');
단언(탭('라인마스터').getLastRow() - 1 === 5, '라인마스터 5행');
단언(탭('목표').getLastRow() - 1 === 12, '목표 12행');
단언(로그포함('가상 생산일보 5,475행 생성'), '생성 로그 5,475행');
단언(로그포함('생산일보 서식: 수식 16,425개'), '수식 16,425개 (5,475 × 3열)');
단언(로그포함('드롭다운 4열'), '드롭다운 4열');
단언(로그포함('집계: 월별 12행 / 라인별 5행 / 불량유형별 6행 / 교대조별 3행 / 정지사유별 7행'), '집계 5탭 행 수');
단언(로그포함('대시보드: KPI 카드 6 / 차트 6개'), 'KPI 6 · 차트 6');
단언(탭('대시보드').getCharts().length === 6, '대시보드 차트 실제 6개');
단언(!properties['샘플데이터_진행'], '진행 속성 정리됨');

// 수식 검사: 엑셀에 없는 함수 금지, 범위는 실제 데이터 범위로 한정
const 행2 = 생산.getRange(2, 1, 1, 15).getValues()[0];
console.log('[확인] 생산일보 2행 = ' + JSON.stringify(행2.map((x) => (x instanceof Date ? Utils.formatDate(x, '', 'yyyy-MM-dd') : x))));
단언(String(행2[1]).indexOf('=TEXT(') === 0, '월 열은 TEXT 수식');
단언(String(행2[3]).indexOf('XLOOKUP') > 0 && String(행2[3]).indexOf('라인마스터!$A$2:$A$6') > 0, '라인명은 XLOOKUP, 라인마스터 범위 $2:$6');
단언(String(행2[12]).indexOf('=IFERROR(') === 0, '불량률(%)은 IFERROR 수식');
const 월별 = 탭('집계_월별');
const 월별2행 = 월별.getRange(2, 1, 1, 14).getValues()[0];
console.log('[확인] 집계_월별 2행 = ' + JSON.stringify(월별2행).slice(0, 400));
단언(월별2행[0] === '2025-10', '집계_월별 A2 = "2025-10" (텍스트)');
단언(String(월별2행[1]).indexOf('생산일보!$B$2:$B$5476') > 0, 'SUMIFS 범위가 $2:$5476 으로 한정');
const 모든수식 = [];
['집계_월별', '집계_라인별', '집계_불량유형별', '집계_교대조별', '집계_정지사유별', '대시보드'].forEach((n) => {
  const s = 탭(n);
  s.getDataRange().getValues().forEach((r) => r.forEach((x) => { if (typeof x === 'string' && x.charAt(0) === '=') 모든수식.push(x); }));
});
모든수식.push(String(행2[1]), String(행2[3]), String(행2[12]));
const 금지 = 모든수식.filter((f) => /\b(QUERY|FILTER|ARRAYFORMULA|SORT|UNIQUE|SPARKLINE)\(/i.test(f));
단언(금지.length === 0, '엑셀 비호환 함수(QUERY/FILTER/ARRAYFORMULA) 0개 (검사한 수식 ' + 모든수식.length + '개)');
const 전체열 = 모든수식.filter((f) => /![$]?[A-Z]+:[$]?[A-Z]+/.test(f));
단언(전체열.length === 0, '전체 열($A:$A) 참조 0개');

// 데이터 이야기 검사 — 스크립트 집계값으로
const 집 = sandbox._집계값();
const 률 = (o) => (o.생산 ? (o.불량 / o.생산) * 100 : 0);
console.log('[확인] 전체: 생산 ' + 집.전체.생산 + ' / 불량률 ' + 률(집.전체).toFixed(2) + '% / 가동률 ' + ((집.전체.가동 / (집.전체.가동 + 집.전체.정지)) * 100).toFixed(1) + '%');
console.log('[확인] 라인별 불량률: ' + 집.라인목록.map((id) => id + ' ' + 률(집.라인별[id]).toFixed(2) + '%').join(' · '));
console.log('[확인] 교대별 불량률: ' + ['A', 'B', 'C'].map((c) => c + ' ' + 률(집.교대별[c]).toFixed(2) + '%').join(' · '));
console.log('[확인] 월별 불량률: ' + 집.월목록.map((m) => m.slice(2) + ' ' + 률(집.월별[m]).toFixed(2)).join(' · '));
console.log('[확인] 불량유형 순서: ' + 집.유형순서.map((t) => t + ' ' + 집.유형별[t].불량).join(' · '));
console.log('[확인] 정지사유 순서: ' + 집.사유순서.map((t) => t + ' ' + 집.사유별[t].정지.toFixed(1) + 'h').join(' · '));
단언(집.월목록.length === 12 && 집.월목록[0] === '2025-10' && 집.월목록[11] === '2026-09', '월 12개 2025-10 ~ 2026-09');
단언(률(집.전체) > 0.7 && 률(집.전체) < 1.8, '전체 불량률이 현실적 범위(0.7~1.8%): ' + 률(집.전체).toFixed(2));
단언(률(집.라인별.L4) > 1.2, 'L4 는 목표(1.2%) 초과: ' + 률(집.라인별.L4).toFixed(2));
단언(률(집.교대별.C) > 률(집.교대별.A), '야간 C조 불량률 > A조');
const L3사고 = 집.월라인['2026-05|L3'];
단언(률(L3사고) > 2.5 * 률(집.라인별.L3), '2026-05 L3 급증(연평균의 ' + (률(L3사고) / 률(집.라인별.L3)).toFixed(1) + '배)');
단언(집.유형순서.length === 6 && 집.사유순서.length === 7, '불량유형 6종 · 정지사유 7종 전부 등장');
const 코멘트 = String(탭('대시보드').getRange(10, 1).getValues()[0][0]);
console.log('\n[코멘트] ' + 코멘트);
단언(코멘트.indexOf('특이사항') >= 0 && 코멘트.indexOf('L3') >= 0 && 코멘트.indexOf('2026-05') >= 0, '코멘트가 2026-05 L3 급증을 짚음');
단언(코멘트.indexOf('2026-09 생산') === 0, '코멘트가 마지막 달로 시작');

/* ========== 3. 설치_전체 재실행 (멱등) ========== */
실행('설치_전체 두 번째 (멱등)', () => sandbox.설치_전체());
단언(생산.getLastRow() - 1 === 5475, '재설치 후 생산일보 5,475행 유지');
단언(탭('설정').getLastRow() - 1 === 16 && 로그포함('설정 16개 중 0개 채움'), '설정 값 유지(0개 채움)');
단언(탭('대시보드').getCharts().length === 6, '재설치 후 차트 6개(겹쳐 쌓이지 않음)');
단언(!로그포함('[건너뜀] 필터'), '재실행 때 필터 거부 없음(기존 필터 제거 후 재생성)');
단언(로그포함('측정로그 정리'), '측정로그 정리 실행');

/* ========== 4. 대시보드갱신 ========== */
실행('대시보드갱신', () => sandbox.대시보드갱신());
단언(로그포함('대시보드 갱신 완료: 5475행'), '갱신 완료 로그');
단언(탭('대시보드').getCharts().length === 6, '갱신 후 차트 6개');
단언(String(탭('설정').getDataRange().getValues().find((r) => r[0] === '마지막갱신')[1]).length >= 16, '마지막갱신 시각 기록');

/* ========== 5. 리포트_PDF ========== */
const pdf결과 = 실행('리포트_PDF', () => sandbox.리포트_PDF());
단언(typeof pdf결과 === 'string' && pdf결과.indexOf('https://drive/') === 0, 'PDF URL 반환: ' + pdf결과);
단언(로그포함('pdf 변환: getAs 경로 사용') && 로그포함('PDF 저장:'), 'PDF 1단계(getAs) 경로');
단언(로그포함('[건너뜀] 메일 알림'), '메일 알림 설정 없음 → [건너뜀]');
단언(Object.values(파일저장소).filter((f) => f.trashed).length >= 1, '임시 사본 휴지통 처리');
const 리포트 = 탭('리포트');
단언(리포트.getLastRow() >= 40, '리포트 ' + 리포트.getLastRow() + '행');
console.log('\n----- 리포트 (앞 14행) -----');
리포트.getRange(1, 1, 14, 8).getValues().forEach((r) => console.log('  ' + r.join(' | ')));
const 리포트값 = 리포트.getDataRange().getValues();
단언(리포트값.every((r) => r.every((x) => !(typeof x === 'string' && x.charAt(0) === '='))), '리포트는 수식 없이 값만');
단언(리포트값.some((r) => r[0] === '4. 종합 코멘트') && 리포트값.some((r) => r[0] === '5. 조치 제안'), '리포트에 코멘트·조치 제안 있음');
/* A4 세로 1장에 담기는지 — 2026-09-25 실제 환경에서 2장으로 갈라졌던 문제의 회귀 테스트 */
const 리포트행수 = 리포트.getLastRow();
const 높이지정 = Object.keys(리포트.rowHeights).map(Number).filter((r) => r >= 1 && r <= 리포트행수);
const 리포트높이 = 높이지정.reduce((a, r) => a + 리포트.rowHeights[r], 0);
단언(높이지정.length === 리포트행수, '리포트 ' + 리포트행수 + '행 전부 행 높이 지정(기본값 21px 로 남는 행 없음)');
단언(리포트높이 > 0 && 리포트높이 < 978, '리포트 총 높이 ' + 리포트높이 + 'px < A4 1장 978px');
단언(로그포함('A4 세로 1장'), '리포트 높이 로그 남김');


/* ========== 6. 엑셀내보내기 ========== */
const xlsx결과 = 실행('엑셀내보내기', () => sandbox.엑셀내보내기());
단언(typeof xlsx결과 === 'string' && xlsx결과.indexOf('https://drive/') === 0, 'xlsx URL 반환');
단언(로그포함('[건너뜀] getAs(xlsx)') && 로그포함('xlsx 변환: 다운로드 URL 경로 사용'), 'xlsx 1단계 거부 → 2단계(다운로드 URL) 자동 전환');
단언(로그포함('값으로 굳힌 탭 6'), '집계 5탭 + 대시보드 값 굳힘');

/* ========== 7. 트리거 ========== */
실행('월간트리거설치', () => sandbox.월간트리거설치());
단언(트리거들.length === 1 && 트리거들[0].getHandlerFunction() === '월간자동실행', '트리거 1개 (월간자동실행)');
실행('월간트리거설치 재실행', () => sandbox.월간트리거설치());
단언(트리거들.length === 1 && 로그포함('기존 1개 교체'), '재설치해도 트리거 1개(교체)');
실행('월간트리거해제', () => sandbox.월간트리거해제());
단언(트리거들.length === 0, '트리거 해제');

/* ========== 8. 메일 알림 (설정 있을 때) ========== */
sandbox._설정쓰기('메일알림', 'test@example.invalid');
실행('리포트_PDF (메일 알림 켬)', () => sandbox.리포트_PDF());
단언(보낸메일.length === 1 && 보낸메일[0].to === 'test@example.invalid', '메일 1통 발송');
sandbox._설정쓰기('메일알림', '');

/* ========== 8-1. getAs 가 막힌 계정 → 다운로드 URL 경로 (PDF 옵션 확인) ========== */
globalThis.__pdfGetAs거부 = true;
실행('리포트_PDF (getAs 막힘 → 다운로드 URL 경로)', () => sandbox.리포트_PDF());
globalThis.__pdfGetAs거부 = false;
단언(로그포함('[건너뜀] getAs(pdf)') && 로그포함('pdf 변환: 다운로드 URL 경로 사용'), 'pdf 1단계 거부 → 2단계 자동 전환');
const pdfURL = 내보내기URL들.filter((u) => u.indexOf('format=pdf') >= 0).pop() || '';
단언(pdfURL.indexOf('fitw=true') >= 0 && pdfURL.indexOf('fith=true') >= 0 && pdfURL.indexOf('portrait=true') >= 0,
  'PDF 내보내기 URL 에 가로·세로 맞춤 + A4 세로 옵션');
단언(pdfURL.indexOf('top_margin=0.4') >= 0 && pdfURL.indexOf('bottom_margin=0.4') >= 0, 'PDF 여백 옵션(0.4인치)');

/* ========== 9. 설치_확인 ========== */
실행('설치_확인', () => sandbox.설치_확인());
단언(마지막로그.trim().split('\n').pop() === '=== 결과: 정상 12 / 문제 0 ===', '설치_확인 12/0');

/* ========== 회귀 테스트 ========== */
require('./regression_test_tail.js')({ sandbox, ss, 탭, 실행, 단언, 로그포함, Sheet, properties, 파일저장소, 마지막로그: () => 마지막로그 });

const 측정 = 탭('측정로그');
console.log('\n[확인] 측정로그:\n' + 측정.getDataRange().getValues().slice(1).map((r) => '  ' + r.join(' | ')).join('\n'));
console.log('\n=== 단언 ' + (단언통과 + 단언실패) + '개 중 통과 ' + 단언통과 + ' / 실패 ' + 단언실패 + ' ===');

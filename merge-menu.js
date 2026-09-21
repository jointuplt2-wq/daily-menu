// 새 xlsx 식단표를 기존 menu.json에 병합하고, 선택적으로 날짜 범위를 제거한다.
//
// 사용법:
//   node merge-menu.js <xlsx경로> [--remove YYYY-MM-DD YYYY-MM-DD]
//
// 예:
//   node merge-menu.js "../260621-0630(케어).xlsx" --remove 2026-06-01 2026-06-10
//
// parse-menu.js 와 달리 menu.json 을 통째로 덮어쓰지 않고, 파싱된 날짜만
// 갱신/추가한 뒤(병합), --remove 로 지정한 날짜 구간을 삭제한다.

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function pad(n) { return String(n).padStart(2, '0'); }

function xlSerialToDate(n) {
  if (typeof n !== 'number' || n < 1) return null;
  const d = new Date(Math.round((n - 25569) * 86400 * 1000));
  if (isNaN(d.getTime())) return null;
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}

function findDate(s) {
  if (!s) return null;
  const yr = new Date().getFullYear();
  let m;
  m = s.match(/(\d{4})-(\d{2})-(\d{2})/); if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/); if (m) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
  m = s.match(/(\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일/); if (m) return '20' + m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
  m = s.match(/(\d{1,2})월\s*(\d{1,2})일/); if (m) return yr + '-' + pad(+m[1]) + '-' + pad(+m[2]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})$/); if (m && +m[1] <= 12 && +m[2] <= 31) return yr + '-' + pad(+m[1]) + '-' + pad(+m[2]);
  return null;
}

function findDateXl(cell) {
  if (cell instanceof Date) return cell.getFullYear() + '-' + pad(cell.getMonth() + 1) + '-' + pad(cell.getDate());
  const n = +cell;
  if (!isNaN(n) && n > 40000 && n < 60000) return xlSerialToDate(n);
  const s = String(cell || '').trim();
  const d = findDate(s); if (d) return d;
  const yr = new Date().getFullYear();
  const m = s.match(/(\d{1,2})[\/\.](\d{1,2})/);
  if (m && +m[1] >= 1 && +m[1] <= 12 && +m[2] >= 1 && +m[2] <= 31) return yr + '-' + pad(+m[1]) + '-' + pad(+m[2]);
  return null;
}

// 식사 구분 라벨 -> menu.json 키
function mealFromLabel(s) {
  if (/조식|아침/.test(s)) return '아침';
  if (/중식|점심/.test(s)) return '점심';
  if (/석식|저녁/.test(s)) return '저녁';
  return null;
}

// 표의 마지막 행을 찾는다.
// 구분 열(조식/중식/석식)은 각 구간만큼 세로 병합돼 있으므로 마지막 병합의 끝이 곧 표의 끝이다.
// 표 아래 비고행("21 | 중,석 순두부" 등)이 석식 메뉴로 섞여 들어가는 것을 막는다.
// 병합 정보가 없으면 -1 을 돌려주고, 호출부에서 빈 행 기준으로 대체 판단한다.
function findTableEndRow(ws, mealCol) {
  const merges = (ws && ws['!merges']) || [];
  let end = -1;
  for (const m of merges) {
    if (m.s.c !== mealCol) continue;
    const addr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
    const label = ws[addr] ? String(ws[addr].v).trim() : '';
    if (mealFromLabel(label)) end = Math.max(end, m.e.r);
  }
  return end;
}

function parseXlsxRows(rowsRaw, rowsFmt, res, ws) {
  let dateRowIdx = -1, dateCols = {}, mealCol = 0;
  for (let r = 0; r < rowsFmt.length; r++) {
    const row = rowsFmt[r]; let found = 0; const tmp = {};
    for (let c = 0; c < row.length; c++) {
      const d = findDateXl(rowsRaw[r] ? rowsRaw[r][c] : '') || findDateXl(row[c]);
      if (d) { tmp[c] = d; found++; }
    }
    if (found >= 2) {
      dateRowIdx = r; dateCols = tmp;
      const minCol = Math.min(...Object.keys(tmp).map(Number));
      mealCol = Math.max(0, minCol - 1);
      break;
    }
  }
  if (dateRowIdx < 0) return;

  // 표 끝 판정: 병합 정보 우선, 없으면 빈 행 2회 연속
  // (표 중간에도 빈 행이 1개 있으므로 1회로는 끊지 않는다)
  const tableEndRow = findTableEndRow(ws, mealCol);
  const lastRow = tableEndRow >= 0 ? Math.min(tableEndRow, rowsFmt.length - 1) : rowsFmt.length - 1;
  const BLANK_RUN_END = 2;

  let curMeal = null;
  let blankRun = 0;
  for (let r = dateRowIdx + 1; r <= lastRow; r++) {
    const row = rowsFmt[r];
    const mealCell = String(row[mealCol] || '').trim();

    if (tableEndRow < 0) {
      const isBlank = !mealCell && Object.keys(dateCols).every(c => !String(row[c] || '').trim());
      if (isBlank) {
        if (++blankRun >= BLANK_RUN_END) break;
        continue;
      }
      blankRun = 0;
    }

    curMeal = mealFromLabel(mealCell) || curMeal;
    if (!curMeal) continue;
    for (const c in dateCols) {
      const dt = dateCols[c];
      const cell = String(row[c] || '').trim(); if (!cell) continue;
      if (!res[dt]) res[dt] = {};
      if (!res[dt][curMeal]) res[dt][curMeal] = [];
      const its = cell.split(/[,\n]+/).map(x => x.trim()).filter(x => x.length >= 1);
      res[dt][curMeal].push(...its);
    }
  }
}

function parseXlsx(targetFile) {
  const wb = XLSX.readFile(targetFile, { cellDates: true, raw: false });
  const res = {};
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rowsRaw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, dateNF: 'yyyy-mm-dd' });
    const rowsFmt = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
    parseXlsxRows(rowsRaw, rowsFmt, res, ws);
  }
  return res;
}

function main() {
  const args = process.argv.slice(2);
  const xlsxPath = args.find(a => !a.startsWith('--'));
  if (!xlsxPath) {
    console.error('❌ xlsx 경로를 지정하세요. 예: node merge-menu.js "../260621-0630(케어).xlsx" --remove 2026-06-01 2026-06-10');
    process.exit(1);
  }
  if (!fs.existsSync(xlsxPath)) {
    console.error(`❌ 파일을 찾을 수 없습니다: ${xlsxPath}`);
    process.exit(1);
  }

  const removeIdx = args.indexOf('--remove');
  let removeStart = null, removeEnd = null;
  if (removeIdx >= 0) {
    removeStart = args[removeIdx + 1];
    removeEnd = args[removeIdx + 2];
    if (!removeStart || !removeEnd) {
      console.error('❌ --remove 는 시작/끝 날짜 2개가 필요합니다. 예: --remove 2026-06-01 2026-06-10');
      process.exit(1);
    }
  }

  console.log(`📊 파싱 중: ${xlsxPath}`);
  const parsed = parseXlsx(xlsxPath);
  const parsedKeys = Object.keys(parsed).sort();
  if (!parsedKeys.length) {
    console.error('❌ 파싱 실패: 날짜/식사 데이터를 찾을 수 없습니다.');
    process.exit(1);
  }
  console.log(`   파싱된 날짜: ${parsedKeys[0]} ~ ${parsedKeys[parsedKeys.length - 1]} (${parsedKeys.length}일)`);

  const menuPath = path.join(__dirname, 'menu.json');
  const menu = JSON.parse(fs.readFileSync(menuPath, 'utf-8'));

  // 병합 (같은 날짜는 새 데이터로 교체)
  for (const d of parsedKeys) menu[d] = parsed[d];

  // 삭제
  if (removeStart && removeEnd) {
    let removed = 0;
    for (const d of Object.keys(menu)) {
      if (d >= removeStart && d <= removeEnd) { delete menu[d]; removed++; }
    }
    console.log(`   제거: ${removeStart} ~ ${removeEnd} (${removed}일)`);
  }

  // 날짜순 정렬 후 저장
  const sorted = {};
  for (const k of Object.keys(menu).sort()) sorted[k] = menu[k];
  fs.writeFileSync(menuPath, JSON.stringify(sorted, null, 2), 'utf-8');

  const ks = Object.keys(sorted);
  console.log(`✅ menu.json 갱신 완료. 범위: ${ks[0]} ~ ${ks[ks.length - 1]} (${ks.length}일)`);
}

main();

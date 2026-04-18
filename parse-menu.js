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

function findMeal(s) {
  if (/^(아침|조식)/.test(s)) return '아침';
  if (/^(점심|중식)/.test(s)) return '점심';
  if (/^(저녁|석식)/.test(s)) return '저녁';
  return null;
}

function splitItems(s) {
  if (!s.trim()) return [];
  if (s.indexOf(',') >= 0 || s.indexOf('，') >= 0)
    return s.split(/[,，]+/).map(x => x.trim()).filter(x => x.length > 0);
  return s.split(/\s+/).map(x => x.trim()).filter(x => x.length >= 2);
}

function parseXlsxRows(rowsRaw, rowsFmt, res) {
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

  if (dateRowIdx >= 0) {
    let curMeal = null;
    for (let r = dateRowIdx + 1; r < rowsFmt.length; r++) {
      const row = rowsFmt[r];
      const mealCell = String(row[mealCol] || '').trim();
      if (/조식|아침/.test(mealCell)) curMeal = '아침';
      else if (/중식|점심/.test(mealCell)) curMeal = '점심';
      else if (/석식|저녁/.test(mealCell)) curMeal = '저녁';
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
    return;
  }

  // 폴백: 텍스트 방식
  let cur = null;
  for (let r = 0; r < rowsFmt.length; r++) {
    const line = rowsFmt[r].map(x => String(x || '').trim()).filter(x => x).join(' ');
    if (!line) continue;
    const d = findDateXl(line) || findDate(line);
    if (d) { cur = d; if (!res[cur]) res[cur] = {}; continue; }
    if (!cur) continue;
    const meal = findMeal(line);
    if (meal) {
      const rest = line.replace(/^(아침|조식|점심|중식|저녁|석식)/, '').replace(/^[\s:：]+/, '');
      if (!res[cur][meal]) res[cur][meal] = [];
      const its = splitItems(rest);
      res[cur][meal].push(...its.filter(x => x));
    }
  }
}

// 저장소 루트에서 가장 최근 xlsx 파일 찾기
const xlsxFiles = fs.readdirSync('.').filter(f => /\.(xlsx|xls|xlsm)$/i.test(f));
if (!xlsxFiles.length) {
  console.error('❌ xlsx 파일이 없습니다.');
  process.exit(1);
}

// 수정일 기준 최신 파일 선택
xlsxFiles.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
const targetFile = xlsxFiles[0];
console.log(`📊 파싱 중: ${targetFile}`);

const wb = XLSX.readFile(targetFile, { cellDates: true, raw: false });
const res = {};

for (const sheetName of wb.SheetNames) {
  const ws = wb.Sheets[sheetName];
  const rowsRaw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, dateNF: 'yyyy-mm-dd' });
  const rowsFmt = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  parseXlsxRows(rowsRaw, rowsFmt, res);
}

const cnt = Object.keys(res).length;
if (!cnt) {
  console.error('❌ 파싱 실패: 날짜/식사 데이터를 찾을 수 없습니다.');
  process.exit(1);
}

fs.writeFileSync('menu.json', JSON.stringify(res, null, 2), 'utf-8');
console.log(`✅ menu.json 생성 완료 (${cnt}일치)`);

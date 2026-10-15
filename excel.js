// Excel export/import. ExcelJS is passed in: window.ExcelJS in the browser, the npm package in tests.
import { LEVELS, columnNumbers, allParts, hasProduct, code, ymd } from './logic.js';

export const MARKER = 'INVENTORY-APP-V1';
const HEADER = ['Row', 'Box', 'Partition', 'Code', 'Stock'];
const W = 6; // Excel columns per stillage: divides evenly into 1, 2 or 3 side-by-side partitions
// [column offset, row offset, width, height] of each partition inside a stillage's 6×2 cells.
const BLOCKS = {
  1: [[0, 0, 6, 2]],
  2: [[0, 0, 3, 2], [3, 0, 3, 2]],
  3: [[0, 0, 2, 2], [2, 0, 2, 2], [4, 0, 2, 2]],
  4: [[0, 0, 3, 1], [3, 0, 3, 1], [0, 1, 3, 1], [3, 1, 3, 1]],
};
const CENTER = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };
const THIN = { style: 'thin' }, THICK = { style: 'medium' };
const GREY = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
const col = i => 2 + i * W; // map column i and table field i both start here, so the table lines up under the map

// Writes a value and style, then merges. Style goes on first: ExcelJS copies it to every cell of the merge.
// ponytail: ExcelJS checks each merge against all earlier ones (O(n²)), ~1s per 600 products on a PC; write the table unmerged if exports get slow.
function block(ws, r, c, w, h, value, style = {}) {
  const cell = ws.getCell(r, c);
  cell.value = value;
  Object.assign(cell, style);
  if (w > 1 || h > 1) ws.mergeCells(r, c, r + h - 1, c + w - 1);
}

export async function exportWorkbook(ExcelJS, state, { date = new Date(), chunk = 30000 } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Warehouse');
  const widest = Math.max(HEADER.length, ...state.rows.map(row => row.columns.length));
  ws.getColumn(1).width = 4;
  for (let c = 2; c < col(widest); c++) ws.getColumn(c).width = 4;

  let r = 1;
  for (const row of state.rows) {
    block(ws, r, 1, 1, 1, row.name, { font: { bold: true, size: 14 } });
    const nums = columnNumbers(row);
    nums.forEach((n, c) => n && block(ws, r + 1, col(c), W, 1, n, { alignment: CENTER, font: { bold: true } }));
    const top = r + 2;
    for (let i = 0; i < 6; i++) ws.getRow(top + i).height = 20;
    LEVELS.forEach((letter, l) => block(ws, top + l * 2, 1, 1, 2, letter, { alignment: CENTER, font: { bold: true } }));
    row.columns.forEach((column, c) => {
      if (column.type === 'space') return block(ws, top, col(c), W, 6, 'Walkway', { alignment: CENTER, fill: GREY });
      column.levels.forEach((lvl, l) => BLOCKS[lvl.parts.length].forEach(([dx, dy, w, h], p) => {
        const border = { top: dy === 0 ? THICK : THIN, bottom: dy + h === 2 ? THICK : THIN, left: dx === 0 ? THICK : THIN, right: dx + w === W ? THICK : THIN };
        block(ws, top + l * 2 + dy, col(c) + dx, w, h, code(lvl.parts[p]) || 'NA', { alignment: CENTER, border });
      }));
    });
    r = top + 7; // 6 map rows, then 1 blank
  }

  r += 1;
  block(ws, r, 1, 1, 1, `Inventory count — ${ymd(date)}`, { font: { bold: true, size: 14 } });
  r += 1;
  HEADER.forEach((h, i) => block(ws, r, col(i), W, 1, h, { font: { bold: true }, border: { bottom: THIN } }));
  for (const x of allParts(state)) {
    if (!hasProduct(x.part)) continue;
    r += 1;
    [x.row, x.box, x.partition, code(x.part), x.part.stock].forEach((v, i) => block(ws, r, col(i), W, 1, v));
  }

  const meta = wb.addWorksheet('_layout', { state: 'hidden' });
  meta.getCell(1, 1).value = MARKER;
  const json = JSON.stringify(state);
  // Excel caps a cell at 32,767 characters, so the layout runs down column A in chunks.
  for (let i = 0; i * chunk < json.length; i++) meta.getCell(2 + i, 1).value = json.slice(i * chunk, (i + 1) * chunk);
  return wb.xlsx.writeBuffer();
}

const key = (row, box, partition) => `${row.trim().toUpperCase()}|${box.trim().toLowerCase()}|${Number(partition)}`;

// Layout comes from the hidden sheet; Stock comes from the count table, so counts edited in Excel win.
export async function importWorkbook(ExcelJS, buffer) {
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(buffer); } catch { return { error: "Couldn't open that file. Is it an .xlsx spreadsheet?" }; }
  const meta = wb.getWorksheet('_layout'), ws = wb.getWorksheet('Warehouse');
  if (!meta || !ws || meta.getCell(1, 1).text !== MARKER) return { error: 'Not a file exported by this app.' };

  let json = '';
  for (let r = 2; meta.getCell(r, 1).value != null; r++) json += meta.getCell(r, 1).value;
  let state;
  try { state = JSON.parse(json); } catch { return { error: 'The layout data in this file is damaged.' }; }
  if (!Array.isArray(state?.rows) || !Array.isArray(state?.styles)) return { error: 'The layout data in this file is damaged.' };

  const index = new Map(allParts(state).map(x => [key(x.row, x.box, x.partition), x.part]));
  let header = 0;
  ws.eachRow((row, n) => { if (!header && HEADER.every((h, i) => row.getCell(col(i)).text === h)) header = n; });
  if (!header) return { error: "Couldn't find the count table (Row | Box | Partition | Code | Stock)." };

  const skipped = [];
  for (let n = header + 1; n <= ws.rowCount; n++) {
    const [row, box, partition, , stock] = HEADER.map((_, i) => ws.getCell(n, col(i)).text.trim());
    if (!row && !box) continue;
    const part = index.get(key(row, box, partition));
    if (!part) skipped.push(`Line ${n}: ${row} ${box} partition ${partition} doesn't exist`);
    else if (stock === '') part.stock = null;
    else if (/^\d+$/.test(stock)) part.stock = Number(stock);
    else skipped.push(`Line ${n}: stock "${stock}" isn't a whole number`);
  }
  return { state, skipped };
}

// Excel export/import. ExcelJS is passed in: window.ExcelJS in the browser, the npm package in tests.
import { LEVELS, LAYOUTS, columnNumbers, allParts, hasProduct, code, ymd, migrate } from './logic.js';

export const MARKER = 'INVENTORY-APP-V1';
const HEADER = ['Row', 'Box', 'Partition', 'Code', 'Stock'];
const G = 12; // each stillage level is a 12×12 block of Excel cells: one cell per unit of the LAYOUTS grid
const CENTER = { horizontal: 'center', vertical: 'middle', shrinkToFit: true };
const THIN = { style: 'thin' }, THICK = { style: 'medium' };
const GREY = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
const col = i => 2 + i * G; // map column i and table field i both start here, so the table lines up under the map

// Writes a value and style, then merges. Style goes on first: ExcelJS copies it to every cell of the merge.
// ponytail: ExcelJS checks each merge against all earlier ones, so time grows as merges²; keep merges to ~2 per product.
function block(ws, r, c, w, h, value, style = {}) {
  const cell = ws.getCell(r, c);
  cell.value = value;
  Object.assign(cell, style);
  if (w > 1 || h > 1) ws.mergeCells(r, c, r + h - 1, c + w - 1);
}

function drawMap(ws, state) {
  let r = 1;
  for (const row of state.rows) {
    block(ws, r, 1, 1, 1, row.name, { font: { bold: true, size: 14 } });
    const nums = columnNumbers(row);
    nums.forEach((n, c) => n && block(ws, r + 1, col(c), G, 1, n, { alignment: CENTER, font: { bold: true } }));
    const top = r + 2;
    for (let i = 0; i < LEVELS.length * G; i++) ws.getRow(top + i).height = 6;
    LEVELS.forEach((letter, l) => block(ws, top + l * G, 1, 1, G, letter, { alignment: CENTER, font: { bold: true } }));
    row.columns.forEach((column, c) => {
      if (column.type === 'space') return block(ws, top, col(c), G, LEVELS.length * G, column.name, { alignment: CENTER, fill: GREY });
      column.levels.forEach((lvl, l) => LAYOUTS.get(lvl.layout).forEach(([x, y, w, h], p) => {
        const border = { top: y === 0 ? THICK : THIN, bottom: y + h === G ? THICK : THIN, left: x === 0 ? THICK : THIN, right: x + w === G ? THICK : THIN };
        block(ws, top + l * G + y, col(c) + x, w, h, code(lvl.parts[p]) || 'NA', { alignment: CENTER, border });
      }));
    });
    r = top + LEVELS.length * G + 1; // the three levels, then 1 blank row
  }
  return r;
}

export async function exportWorkbook(ExcelJS, state, { date = new Date(), chunk = 30000, map = true, count = true } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Warehouse');
  const widest = Math.max(HEADER.length, ...state.rows.map(row => row.columns.length));
  ws.getColumn(1).width = 4;
  for (let c = 2; c < col(widest); c++) ws.getColumn(c).width = 2;

  let r = map ? drawMap(ws, state) + 1 : 1;
  if (count) {
    block(ws, r, 1, 1, 1, `Inventory count — ${ymd(date)}`, { font: { bold: true, size: 14 } });
    r += 1;
    HEADER.forEach((h, i) => block(ws, r, col(i), G, 1, h, { font: { bold: true }, border: { bottom: THIN } }));
    for (const x of allParts(state)) {
      if (!hasProduct(x.part)) continue;
      r += 1;
      // Text spills into the empty cells to its right, so only Stock (a number, which doesn't spill) is merged.
      [x.row, x.box, x.partition, code(x.part)].forEach((v, i) => block(ws, r, col(i), 1, 1, v));
      block(ws, r, col(4), G, 1, x.part.stock);
    }
  }

  // The full layout always travels in the hidden sheet, so a map-only or count-only file still imports.
  const meta = wb.addWorksheet('_layout', { state: 'hidden' });
  meta.getCell(1, 1).value = MARKER;
  const json = JSON.stringify(state);
  // Excel caps a cell at 32,767 characters, so the layout runs down column A in chunks.
  for (let i = 0; i * chunk < json.length; i++) meta.getCell(2 + i, 1).value = json.slice(i * chunk, (i + 1) * chunk);
  return wb.xlsx.writeBuffer();
}

const key = (row, box, partition) => `${row.trim().toUpperCase()}|${box.trim().toLowerCase()}|${Number(partition)}`;

// The hidden sheet is untrusted input: accept only the shape the app itself saves.
const isText = v => typeof v === 'string';
const isPart = p => isText(p?.style) && isText(p.number) && isText(p.color) && (p.stock === null || (Number.isInteger(p.stock) && p.stock >= 0));
const isLevel = lvl => LAYOUTS.has(lvl?.layout) && Array.isArray(lvl.parts) &&
  lvl.parts.length === LAYOUTS.get(lvl.layout).length && lvl.parts.every(isPart);
const isColumn = col => (col?.type === 'space' && isText(col.name)) ||
  (col?.type === 'stillage' && Array.isArray(col.levels) && col.levels.length === LEVELS.length && col.levels.every(isLevel));
const isState = s => Array.isArray(s?.styles) && s.styles.every(isText) &&
  Array.isArray(s.rows) && s.rows.every(row => isText(row?.name) && Array.isArray(row.columns) && row.columns.every(isColumn));

// The count table's header row and each field's column, found by text: older versions put the table in other columns.
function findTable(ws) {
  let found = null;
  ws.eachRow((row, n) => {
    if (found) return;
    const at = {};
    row.eachCell(cell => { if (HEADER.includes(cell.text) && !at[cell.text]) at[cell.text] = cell.col; });
    if (HEADER.every(h => at[h])) found = { header: n, cols: HEADER.map(h => at[h]) };
  });
  return found;
}

// Layout comes from the hidden sheet; Stock comes from the count table if there is one, so counts edited in Excel win.
export async function importWorkbook(ExcelJS, buffer) {
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.load(buffer); } catch { return { error: "Couldn't open that file. Is it an .xlsx spreadsheet?" }; }
  const meta = wb.getWorksheet('_layout'), ws = wb.getWorksheet('Warehouse');
  if (!meta || !ws || meta.getCell(1, 1).text !== MARKER) return { error: 'Not a file exported by this app.' };

  let json = '';
  for (let r = 2; meta.getCell(r, 1).value != null; r++) json += meta.getCell(r, 1).value;
  let state;
  try { state = migrate(JSON.parse(json)); } catch { return { error: 'The layout data in this file is damaged.' }; }
  if (!isState(state)) return { error: 'The layout data in this file is damaged.' };

  const index = new Map(allParts(state).map(x => [key(x.row, x.box, x.partition), x.part]));
  const table = findTable(ws);
  const skipped = [];
  if (table) for (let n = table.header + 1; n <= ws.rowCount; n++) {
    const [row, box, partition, , stock] = table.cols.map(c => ws.getCell(n, c).text.trim());
    if (!row && !box) continue;
    const part = index.get(key(row, box, partition));
    if (!part) skipped.push(`Line ${n}: ${row} ${box} partition ${partition} doesn't exist`);
    else if (stock === '') part.stock = null;
    else if (/^\d+$/.test(stock)) part.stock = Number(stock);
    else skipped.push(`Line ${n}: stock "${stock}" isn't a whole number`);
  }
  return { state, skipped };
}

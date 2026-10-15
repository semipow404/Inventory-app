import assert from 'node:assert/strict';
import * as L from './logic.js';
import ExcelJS from 'exceljs';
import { exportWorkbook, importWorkbook } from './excel.js';

// --- Box labels skip walkways and renumber when a column is added at the left end
const s = L.newState();
s.rows.push(L.newRow('Mid Row'));
const row = s.rows[0];
L.addColumn(row, 'space', 'left');
L.addColumn(row, 'stillage', 'right');
assert.deepEqual(L.columnNumbers(row), [null, 1, 2]);
row.columns[1].levels[0].parts[0] = { style: 'LSW', number: '022', color: 'BR', stock: 10 };
L.addColumn(row, 'stillage', 'left');
assert.deepEqual(L.columnNumbers(row), [1, null, 2, 3]);
const found = L.allParts(s).find(x => L.code(x.part) === 'LSW022BR');
assert.equal(found.box, 'a2');
assert.equal(found.partition, 1);
assert.equal(L.allParts(s).length, 9); // 3 stillage columns × 3 levels × 1 partition

// --- Search: filled boxes must match that part of the code, ignoring case
const p = { style: 'LSW', number: '012', color: 'KK', stock: 20 };
assert.ok(L.matches(p, { style: 'LSW', number: '', color: '' }));
assert.ok(L.matches(p, { style: 'LSW', number: '012', color: 'kk' }));
assert.ok(L.matches(p, { style: '', number: '', color: 'KK' }));
assert.ok(!L.matches(p, { style: 'LSW', number: '12', color: '' }));
assert.ok(!L.matches(p, { style: '', number: '', color: '' }));
assert.ok(!L.matches(L.emptyPart(), { style: '', number: '', color: '' }));
assert.ok(!L.isSearching({ style: '', number: '', color: '' }));

// --- Partitions
const lvl = { parts: [L.emptyPart()] };
L.setPartitions(lvl, 4);
assert.equal(lvl.parts.length, 4);
lvl.parts[3] = { ...p };
assert.ok(L.dropsProducts(lvl, 3));
assert.ok(!L.dropsProducts(lvl, 4));
L.setPartitions(lvl, 2);
assert.deepEqual(lvl.parts, [L.emptyPart(), L.emptyPart()]);

// --- Product pop-up values
assert.deepEqual(L.makePart({ style: 'LSW', number: ' 022 ', color: 'wrbf', stock: '12' }).part,
  { style: 'LSW', number: '022', color: 'WRBF', stock: 12 });
assert.equal(L.makePart({ style: 'LSW', number: '022', color: '', stock: '' }).part.stock, null);
assert.equal(L.makePart({ style: 'LSW', number: '022', color: '', stock: '0' }).part.stock, 0);
assert.ok(L.makePart({ style: '', number: '022', color: '', stock: '' }).error);
assert.ok(L.makePart({ style: 'LSW', number: '022', color: '', stock: '-1' }).error);
assert.ok(L.makePart({ style: 'LSW', number: '022', color: '', stock: '2.5' }).error);
assert.deepEqual(L.makePart({ style: '', number: ' ', color: '', stock: '5' }).part, L.emptyPart());

// --- Row names
assert.ok(L.rowNameError(s, '  '));
assert.ok(L.rowNameError(s, ' mid row '));
assert.equal(L.rowNameError(s, 'Mid Row', s.rows[0]), null); // renaming a row to its own name
assert.equal(L.rowNameError(s, 'Back Row'), null);
assert.equal(L.newRow('  Back Row ').name, 'Back Row');

// --- Styles
assert.equal(L.addStyle(s, ' lsw '), null);
assert.deepEqual(s.styles, ['LSW']);
assert.ok(L.addStyle(s, 'LSW'));
assert.ok(L.addStyle(s, '  '));
// A style removed from Settings still shows for a product that uses it (Review Focus 4)
assert.deepEqual(L.styleChoices(['LSW', 'LAN'], 'LCW'), ['', 'LSW', 'LAN', 'LCW']);
assert.deepEqual(L.styleChoices(['LSW', 'LAN'], 'LSW'), ['', 'LSW', 'LAN']);
assert.deepEqual(L.styleChoices(['LSW'], ''), ['', 'LSW']);

// --- Summary and dates
assert.deepEqual(L.summary(s), { rows: 1, stillages: 9, products: 1 });
assert.equal(L.ymd(new Date(2026, 9, 5)), '2026-10-05');

// ===== Excel =====
// Sample: walkway, 1/2/3/4 partitions, counted, uncounted and zero stock, two rows.
function sample() {
  const s = L.newState();
  s.styles = ['LSW', 'LAN'];
  s.rows.push(L.newRow('Mid Row'), L.newRow('Back Row'));
  const [mid, back] = s.rows;
  L.addColumn(mid, 'space', 'left');        // columns: [walkway, stillage 1]
  L.addColumn(mid, 'stillage', 'right');    // columns: [walkway, stillage 1, stillage 2]
  const one = mid.columns[1], two = mid.columns[2];
  L.setPartitions(one.levels[0], 4);
  one.levels[0].parts[0] = { style: 'LSW', number: '022', color: 'BR', stock: 10 };
  one.levels[0].parts[3] = { style: 'LSW', number: '021', color: 'WKOF', stock: null };
  L.setPartitions(two.levels[1], 2);
  L.setPartitions(two.levels[2], 3);
  two.levels[2].parts[2] = { style: 'LAN', number: '911', color: 'FG', stock: 0 };
  back.columns[0].levels[0].parts[0] = { style: 'LAN', number: '042', color: 'W', stock: 7 };
  return s;
}

const W = 6, col = i => 2 + i * W; // each stillage / table field spans 6 Excel columns starting at column B
const opened = async buf => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf); return wb; };
const headerRow = ws => { let h = 0; ws.eachRow((r, n) => { if (!h && r.getCell(col(0)).text === 'Row' && r.getCell(col(4)).text === 'Stock') h = n; }); return h; };

// --- Map layout: Mid Row name in row 1, column numbers in row 2, levels from row 3 (2 Excel rows per level)
{
  const wb = await opened(await exportWorkbook(ExcelJS, sample(), { date: new Date(2026, 9, 15) }));
  const ws = wb.getWorksheet('Warehouse');
  assert.equal(wb.getWorksheet('_layout').state, 'hidden');
  assert.equal(ws.getCell(1, 1).text, 'Mid Row');
  assert.equal(ws.getCell(2, col(1)).text, '1');
  assert.equal(ws.getCell(2, col(2)).text, '2');
  assert.equal(ws.getCell(3, 1).text, 'a');
  assert.equal(ws.getCell(3, col(0)).text, 'Walkway');
  assert.equal(ws.getCell(8, col(0) + 5).text, 'Walkway');            // walkway covers 6 columns × 6 rows
  assert.equal(ws.getCell(3, col(1)).text, 'LSW022BR');               // 4-way: top-left is 3 wide × 1 high
  assert.equal(ws.getCell(3, col(1) + 2).text, 'LSW022BR');
  assert.equal(ws.getCell(3, col(1) + 3).text, 'NA');                 // top-right
  assert.equal(ws.getCell(4, col(1) + 3).text, 'LSW021WKOF');         // bottom-right
  assert.equal(ws.getCell(7, col(2) + 4).text, 'LAN911FG');           // 3-way, level c, third block (2 wide × 2 high)
  assert.equal(ws.getCell(8, col(2) + 5).text, 'LAN911FG');
  assert.equal(ws.getCell(5, col(2)).text, 'NA');                     // 2-way, level b, left half
  // --- Count table: walking order, blank stock for uncounted
  const h = headerRow(ws);
  assert.ok(h > 0, 'count table header found');
  assert.ok(ws.getCell(h - 1, 1).text.includes('2026-10-15'));
  const line = n => [0, 1, 2, 3, 4].map(i => ws.getCell(n, col(i)).text);
  assert.deepEqual(line(h + 1), ['Mid Row', 'a1', '1', 'LSW022BR', '10']);
  assert.deepEqual(line(h + 2), ['Mid Row', 'a1', '4', 'LSW021WKOF', '']);
  assert.deepEqual(line(h + 3), ['Mid Row', 'c2', '3', 'LAN911FG', '0']);
  assert.deepEqual(line(h + 4), ['Back Row', 'a1', '1', 'LAN042W', '7']);
  assert.equal(ws.getCell(h + 5, col(0)).text, '');
}

// --- Round trip with edits a user might make in Excel (Review Focus 2 and 3)
{
  const wb = await opened(await exportWorkbook(ExcelJS, sample(), { chunk: 7 })); // tiny chunks split the JSON across many cells
  const ws = wb.getWorksheet('Warehouse');
  const h = headerRow(ws);
  ws.getCell(h + 1, col(4)).value = 99;                 // LSW022BR: new count
  ws.getCell(h + 3, col(4)).value = 'abc';              // LAN911FG: not a number → skipped, keeps 0
  ws.getCell(h + 4, col(0)).value = '  back row ';      // LAN042W: different case/spaces still match
  ws.getCell(h + 4, col(1)).value = 'A1';
  ws.getCell(h + 4, col(4)).value = null;               // cleared → not counted
  ws.getCell(h + 5, col(0)).value = 'Mid Row';          // a line for a box that doesn't exist
  ws.getCell(h + 5, col(1)).value = 'b9';
  ws.getCell(h + 5, col(2)).value = 1;
  ws.getCell(h + 5, col(4)).value = 5;
  const out = await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer());
  const expected = sample();
  expected.rows[0].columns[1].levels[0].parts[0].stock = 99;
  expected.rows[1].columns[0].levels[0].parts[0].stock = null;
  assert.deepEqual(out.state, expected);
  assert.equal(out.skipped.length, 2);
  assert.match(out.skipped.join('\n'), /abc/);
  assert.match(out.skipped.join('\n'), /b9/);
}

// --- A layout too big for one Excel cell survives (Review Focus 1)
{
  const big = L.newState();
  big.rows.push(L.newRow('Long Row'));
  for (let i = 0; i < 299; i++) L.addColumn(big.rows[0], 'stillage', 'right');
  big.rows[0].columns[0].levels[0].parts[0] = { style: 'LSW', number: '022', color: 'BR', stock: 1 };
  assert.ok(JSON.stringify(big).length > 32767, 'layout must not fit in one Excel cell');
  assert.deepEqual(await importWorkbook(ExcelJS, await exportWorkbook(ExcelJS, big)), { state: big, skipped: [] });
}

// --- Files this app didn't make are rejected
{
  const plain = new ExcelJS.Workbook();
  plain.addWorksheet('Sheet1');
  assert.match((await importWorkbook(ExcelJS, await plain.xlsx.writeBuffer())).error, /Not a file exported by this app/);
  assert.ok((await importWorkbook(ExcelJS, Buffer.from('not a spreadsheet'))).error);
}

// --- Layout data the app couldn't have saved itself is rejected (a crafted file could inject HTML via stock)
{
  const tamper = async mutate => {
    const wb = await opened(await exportWorkbook(ExcelJS, sample()));
    const meta = wb.getWorksheet('_layout');
    const s = JSON.parse(meta.getCell(2, 1).value);
    mutate(s);
    meta.getCell(2, 1).value = JSON.stringify(s);
    return (await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer())).error ?? 'accepted';
  };
  const lsw = s => s.rows[0].columns[1].levels[0];
  assert.match(await tamper(s => { lsw(s).parts[0].stock = '<img src=x onerror=alert(1)>'; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].stock = -3; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].color = { x: 1 }; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts = []; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts.push(L.emptyPart()); }), /damaged/); // 5 partitions
  assert.match(await tamper(s => { s.rows[0].columns[1].levels.pop(); }), /damaged/);
  assert.match(await tamper(s => { s.rows[0].columns[0] = { type: 'shelf' }; }), /damaged/);
  assert.match(await tamper(s => { s.rows[0].name = 42; }), /damaged/);
  assert.match(await tamper(s => { s.styles = [null]; }), /damaged/);
  assert.equal(await tamper(() => {}), 'accepted');
}

// --- A big warehouse exports quickly enough to keep the phone's share sheet (it allows ~5 s after the tap).
// ExcelJS checks each merge against all earlier ones, so merging every count-table cell grew as n².
{
  const big = L.newState();
  big.rows.push(L.newRow('Big Row'));
  for (let i = 0; i < 200; i++) L.addColumn(big.rows[0], 'stillage', 'right');
  big.rows[0].columns.forEach((col, c) => col.levels.forEach(lvl => {
    L.setPartitions(lvl, 4);
    lvl.parts = lvl.parts.map((_, p) => ({ style: 'LSW', number: String(c).padStart(3, '0'), color: 'C' + p, stock: c * 100 + p }));
  }));
  const t = Date.now();
  const buf = await exportWorkbook(ExcelJS, big);
  const ms = Date.now() - t;
  assert.ok(ms < 3000, `exporting ${L.summary(big).products} products took ${ms} ms`);
  assert.deepEqual(await importWorkbook(ExcelJS, buf), { state: big, skipped: [] });
}

console.log('all tests passed');

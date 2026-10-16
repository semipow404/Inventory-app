import assert from 'node:assert/strict';
import * as L from './logic.js';
import ExcelJS from 'exceljs';
import { exportWorkbook, importWorkbook } from './excel.js';

// A partition as the app saves it; full number and previous counts default to none.
const P = (style, number, color, stock, more = {}) => ({ ...L.emptyPart(), style, number, color, stock, ...more });
const T0 = new Date('2026-10-16T09:00:00Z');

// --- Box labels skip walkways and renumber when a column is added at the left end
const s = L.newState();
s.rows.push(L.newRow('Mid Row'));
const row = s.rows[0];
L.addColumn(row, 'space', 'left');
L.addColumn(row, 'stillage', 'right');
assert.deepEqual(L.columnNumbers(row), [null, 1, 2]);
row.columns[1].levels[0].parts[0] = P('LSW', '022', 'BR', 10);
L.addColumn(row, 'stillage', 'left');
assert.deepEqual(L.columnNumbers(row), [1, null, 2, 3]);
const found = L.allParts(s).find(x => L.code(x.part) === 'LSW022BR');
assert.equal(found.box, 'a2');
assert.equal(found.partition, 1);
assert.equal(L.allParts(s).length, 9); // 3 stillage columns × 3 levels × 1 partition

// --- Search: filled boxes must match that part of the code, ignoring case
const p = P('LSW', '012', 'KK', 20);
assert.ok(L.matches(p, { style: 'LSW', number: '', color: '' }));
assert.ok(L.matches(p, { style: 'LSW', number: '012', color: 'kk' }));
assert.ok(L.matches(p, { style: '', number: '', color: 'KK' }));
assert.ok(!L.matches(p, { style: 'LSW', number: '12', color: '' }));
assert.ok(!L.matches(p, { style: '', number: '', color: '' }));
assert.ok(!L.matches(L.emptyPart(), { style: '', number: '', color: '' }));
assert.ok(!L.isSearching({ style: '', number: '', color: '' }));

// --- Partition layouts: each tiles the stillage's 12×12 grid exactly once, numbered left→right then top→bottom
assert.deepEqual([...L.LAYOUTS.keys()], ['1', '2II', '2=', '3II', '3T', '4', '4T', '6', '9', '12', '16']);
assert.deepEqual([...L.LAYOUTS.values()].map(b => b.length), [1, 2, 2, 3, 3, 4, 4, 6, 9, 12, 16]);
for (const [id, blocks] of L.LAYOUTS) {
  const cover = Array(144).fill(0);
  blocks.forEach(([x, y, w, h]) => { for (let i = x; i < x + w; i++) for (let j = y; j < y + h; j++) cover[j * 12 + i]++; });
  assert.ok(cover.every(n => n === 1), `${id} covers every cell exactly once`);
  const order = blocks.map(([x, y]) => y * 12 + x);
  assert.deepEqual(order, [...order].sort((a, b) => a - b), `${id} is numbered left→right, top→bottom`);
}
assert.deepEqual(L.LAYOUTS.get('2='), [[0, 0, 12, 6], [0, 6, 12, 6]]);                 // two stacked
assert.deepEqual(L.LAYOUTS.get('3T'), [[0, 0, 12, 6], [0, 6, 6, 6], [6, 6, 6, 6]]);    // one on top, two below
assert.deepEqual(L.LAYOUTS.get('4T'), [[0, 0, 12, 6], [0, 6, 4, 6], [4, 6, 4, 6], [8, 6, 4, 6]]);
assert.deepEqual(L.LAYOUTS.get('12')[4], [0, 4, 3, 4]);                                // 4 across × 3 down
const lvl = L.newStillage().levels[0];
assert.equal(lvl.layout, '1');
L.setLayout(lvl, '2II');
lvl.parts[1] = { ...p };
L.setLayout(lvl, '2=');                                  // same count: products stay put
assert.equal(lvl.layout, '2=');
assert.deepEqual(lvl.parts, [L.emptyPart(), p]);
L.setLayout(lvl, '16');
assert.equal(lvl.parts.length, 16);
lvl.parts[15] = { ...p };
assert.ok(L.dropsProducts(lvl, 12));
assert.ok(!L.dropsProducts(lvl, 16));
L.setLayout(lvl, '4T');
assert.equal(lvl.parts.length, 4);
assert.deepEqual(lvl.parts[1], p);

// --- Spaces are named, "Walkway" by default
const spaced = L.newRow('Spaced');
L.addColumn(spaced, 'space', 'left');
assert.deepEqual(spaced.columns[0], { type: 'space', name: 'Walkway' });

// --- Data saved before named layouts and space names still loads
const parts = n => Array.from({ length: n }, () => ({ style: '', number: '', color: '', stock: null })); // first-release partition shape
const old = { styles: [], rows: [{ name: 'R', columns: [
  { type: 'space' },
  { type: 'stillage', levels: [{ parts: parts(1) }, { parts: parts(2) }, { parts: parts(3) }] },
  { type: 'stillage', levels: [{ parts: parts(4) }, { parts: parts(1) }, { parts: parts(1) }] },
] }] };
const migrated = L.migrate(structuredClone(old));
assert.equal(migrated.rows[0].columns[0].name, 'Walkway');
assert.deepEqual(migrated.rows[0].columns[1].levels.map(x => x.layout), ['1', '2II', '3II']);
assert.equal(migrated.rows[0].columns[2].levels[0].layout, '4');
assert.deepEqual(migrated.rows[0].columns[1].levels[0].parts[0], L.emptyPart()); // gains full: null, counts: []
assert.deepEqual(L.migrate(structuredClone(migrated)), migrated); // already-current data is left alone
assert.doesNotThrow(() => L.migrate({ rows: [null, { columns: 'x' }] })); // garbage is left for validation to reject

// --- Product pop-up values
assert.deepEqual(L.makePart({ style: 'LSW', number: ' 022 ', color: 'wrbf', stock: '12', full: '' }, L.emptyPart(), T0).part,
  P('LSW', '022', 'WRBF', 12, { counts: [{ at: T0.toISOString(), stock: 12 }] }));
assert.equal(L.makePart({ style: 'LSW', number: '022', color: '', stock: '', full: '' }).part.stock, null);
assert.equal(L.makePart({ style: 'LSW', number: '022', color: '', stock: '0', full: '' }).part.stock, 0);
// Part of a code is enough, e.g. colour only
assert.deepEqual(L.makePart({ style: '', number: '', color: 'kk', stock: '3', full: '' }, L.emptyPart(), T0).part, P('', '', 'KK', 3, { counts: [{ at: T0.toISOString(), stock: 3 }] }));
assert.deepEqual(L.makePart({ style: '', number: '022', color: '', stock: '', full: '' }).part, P('', '022', '', null));
const kk = P('', '', 'KK', 3);
assert.ok(L.matches(kk, { style: '', number: '', color: 'kk' }));
assert.ok(!L.matches(kk, { style: 'LSW', number: '', color: 'KK' }));
assert.ok(L.makePart({ style: 'LSW', number: '022', color: '', stock: '-1', full: '' }).error);
assert.ok(L.makePart({ style: 'LSW', number: '022', color: '', stock: '2.5', full: '' }).error);
assert.deepEqual(L.makePart({ style: '', number: ' ', color: '', stock: '5', full: '' }).part, L.emptyPart());

// --- "Full" number, last counted and the previous 10 counts (kept with the partition, shown only as "last counted")
{
  const T1 = new Date('2026-10-16T14:05:00Z');
  const fields = (stock, full = '40') => ({ style: 'LSW', number: '022', color: 'BR', stock, full });
  const first = L.makePart(fields('38'), L.emptyPart(), T0).part;
  assert.equal(first.full, 40);
  assert.deepEqual(first.counts, [{ at: T0.toISOString(), stock: 38 }]);
  assert.equal(L.lastCounted(first), T0.toISOString());
  const recount = L.makePart(fields('38'), first, T1).part;              // same number is still a count
  assert.equal(L.lastCounted(recount), T1.toISOString());
  assert.equal(recount.counts.length, 2);
  const blank = L.makePart(fields(''), recount, T1).part;                // blank stock = not counted: no date, history kept
  assert.equal(L.lastCounted(blank), null);
  assert.equal(blank.counts.length, 2);
  let many = first;
  for (let i = 0; i < 15; i++) many = L.makePart(fields(String(i), ''), many, new Date(T0.getTime() + i * 1000)).part;
  assert.equal(many.counts.length, 10);                                  // only the last 10 are kept
  assert.equal(many.counts.at(-1).stock, 14);
  assert.equal(many.full, null);
  assert.match(L.makePart(fields('1', '0')).error, /Full/);              // full must be a whole number of 1 or more
  assert.match(L.makePart(fields('1', '2.5')).error, /Full/);
  assert.deepEqual(L.makePart({ style: '', number: '', color: '', stock: '5', full: '40' }, recount).part, L.emptyPart()); // clearing resets all
  // Usage: stock ÷ full, capped at full; unknown without a count or a full number
  assert.equal(L.usage({ ...first, stock: 20 }), 0.5);
  assert.equal(L.usage({ ...first, stock: 50 }), 1);
  assert.equal(L.usage({ ...first, stock: 0 }), 0);
  assert.equal(L.usage({ ...first, full: null }), null);
  assert.equal(L.usage({ ...first, stock: null }), null);
}

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
  L.setLayout(one.levels[0], '4');
  one.levels[0].parts[0] = P('LSW', '022', 'BR', 10, { full: 40, counts: [{ at: T0.toISOString(), stock: 10 }] });
  one.levels[0].parts[3] = P('LSW', '021', 'WKOF', null);
  L.setLayout(two.levels[1], '2II');
  L.setLayout(two.levels[2], '3II');
  two.levels[2].parts[2] = P('LAN', '911', 'FG', 0);
  back.columns[0].levels[0].parts[0] = P('LAN', '042', 'W', 7);
  mid.columns[0].name = 'Fire exit';
  L.setLayout(back.columns[0].levels[1], '16');
  back.columns[0].levels[1].parts[15] = P('', '', 'KK', 11); // colour-only code
  L.setLayout(back.columns[0].levels[2], '3T');
  return s;
}

// Each stillage is a 12×12 block of Excel cells per level (the layout grid); table fields line up with stillage columns.
const G = 12, col = i => 2 + i * G;
const opened = async buf => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf); return wb; };
const headerRow = ws => { let h = 0; ws.eachRow((r, n) => { if (!h && r.getCell(col(0)).text === 'Row' && r.getCell(col(4)).text === 'Stock') h = n; }); return h; };

// --- Map layout: row name, column numbers, then levels a/b/c of 12 Excel rows each
{
  const wb = await opened(await exportWorkbook(ExcelJS, sample(), { date: new Date(2026, 9, 15) }));
  const ws = wb.getWorksheet('Warehouse');
  const at = (r, c) => ws.getCell(r, c).text;
  assert.equal(wb.getWorksheet('_layout').state, 'hidden');
  // Mid Row: name row 1, numbers row 2, level a rows 3–14, b 15–26, c 27–38
  assert.equal(at(1, 1), 'Mid Row');
  assert.equal(at(2, col(1)), '1');
  assert.equal(at(2, col(2)), '2');
  assert.equal(at(3, 1), 'a');
  assert.equal(at(15, 1), 'b');
  assert.equal(at(3, col(0)), 'Fire exit');                          // named space covers all three levels
  assert.equal(at(38, col(0) + 11), 'Fire exit');
  assert.equal(at(3, col(1)), 'LSW022BR');                           // '4': top-left is 6×6
  assert.equal(at(8, col(1) + 5), 'LSW022BR');
  assert.equal(at(3, col(1) + 6), 'NA');                             // top-right
  assert.equal(at(9, col(1) + 6), 'LSW021WKOF');                     // bottom-right
  assert.equal(at(15, col(2)), 'NA');                                // '2II', level b, left half
  assert.equal(at(27, col(2) + 8), 'LAN911FG');                      // '3II', level c, third block (4×12)
  assert.equal(at(38, col(2) + 11), 'LAN911FG');
  // Back Row starts after Mid Row's 36 level rows + 1 blank: name row 40, numbers 41, level a 42–53, b 54–65, c 66–77
  assert.equal(at(40, 1), 'Back Row');
  assert.equal(at(42, col(0)), 'LAN042W');
  assert.equal(at(63, col(0) + 9), 'KK');                            // '16', level b, partition 16 (3×3, bottom-right)
  assert.equal(at(65, col(0) + 11), 'KK');
  assert.equal(at(62, col(0) + 9), 'NA');                            // partition 12, just above
  assert.equal(at(71, col(0) + 11), 'NA');                           // '3T', level c: top block is 12×6
  assert.equal(at(72, col(0) + 6), 'NA');                            // bottom-right block
  // --- Count table: walking order, blank stock for uncounted
  const h = headerRow(ws);
  assert.ok(h > 0, 'count table header found');
  assert.ok(ws.getCell(h - 1, 1).text.includes('2026-10-15'));
  const line = n => [0, 1, 2, 3, 4].map(i => ws.getCell(n, col(i)).text);
  assert.deepEqual(line(h + 1), ['Mid Row', 'a1', '1', 'LSW022BR', '10']);
  assert.deepEqual(line(h + 2), ['Mid Row', 'a1', '4', 'LSW021WKOF', '']);
  assert.deepEqual(line(h + 3), ['Mid Row', 'c2', '3', 'LAN911FG', '0']);
  assert.deepEqual(line(h + 4), ['Back Row', 'a1', '1', 'LAN042W', '7']);
  assert.deepEqual(line(h + 5), ['Back Row', 'b1', '16', 'KK', '11']);
  assert.equal(ws.getCell(h + 6, col(0)).text, '');
  // Full numbers and count dates stay out of the visible sheet: no extra column, map shows codes only
  assert.equal(ws.getCell(h, col(5)).text, '');
  assert.equal(ws.getCell(h + 1, col(5)).text, '');
  ws.eachRow(r => r.eachCell(c => { // (ExcelJS's .text throws on an empty merged cell, so skip empties)
    if (c.value != null) assert.ok(!/2026-10-16|\/ 40/.test(c.text), `visible cell ${c.address} holds usage data`);
  }));
}

// --- Export options: map only, or count only; both files still import in full (layout travels in the hidden sheet)
{
  const mapOnly = await opened(await exportWorkbook(ExcelJS, sample(), { count: false }));
  assert.equal(headerRow(mapOnly.getWorksheet('Warehouse')), 0, 'map-only file has no count table');
  assert.equal(mapOnly.getWorksheet('Warehouse').getCell(1, 1).text, 'Mid Row');
  const countOnly = await opened(await exportWorkbook(ExcelJS, sample(), { map: false }));
  const ws = countOnly.getWorksheet('Warehouse');
  assert.match(ws.getCell(1, 1).text, /Inventory count/, 'count-only file starts with the count');
  assert.equal(headerRow(ws), 2);
  for (const wb of [mapOnly, countOnly]) {
    assert.deepEqual(await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer()), { state: sample(), skipped: [] });
  }
}

// --- Files from the first release (6-column table, no layout names or space names) still import
{
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Warehouse');
  ['Row', 'Box', 'Partition', 'Code', 'Stock'].forEach((h, i) => { ws.getCell(9, 2 + i * 6).value = h; });
  ['Mid Row', 'a1', 1, 'LSW022BR', 42].forEach((v, i) => { ws.getCell(10, 2 + i * 6).value = v; });
  const meta = wb.addWorksheet('_layout', { state: 'hidden' });
  meta.getCell(1, 1).value = 'INVENTORY-APP-V1';
  const oldEmpty = { style: '', number: '', color: '', stock: null };
  meta.getCell(2, 1).value = JSON.stringify({ styles: ['LSW'], rows: [{ name: 'Mid Row', columns: [{ type: 'space' }, { type: 'stillage', levels: [
    { parts: [{ style: 'LSW', number: '022', color: 'BR', stock: 10 }, oldEmpty] }, { parts: [oldEmpty] }, { parts: [oldEmpty] }] }] }] });
  const out = await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer());
  assert.deepEqual(out.skipped, []);
  assert.equal(out.state.rows[0].columns[0].name, 'Walkway');
  assert.deepEqual(out.state.rows[0].columns[1].levels.map(x => x.layout), ['2II', '1', '1']);
  assert.equal(out.state.rows[0].columns[1].levels[0].parts[0].stock, 42); // read from the old table columns
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
  ws.getCell(h + 6, col(0)).value = 'Mid Row';          // a line for a box that doesn't exist
  ws.getCell(h + 6, col(1)).value = 'b9';
  ws.getCell(h + 6, col(2)).value = 1;
  ws.getCell(h + 6, col(4)).value = 5;
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
  big.rows[0].columns[0].levels[0].parts[0] = P('LSW', '022', 'BR', 1);
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
  assert.match(await tamper(s => { lsw(s).parts.push(L.emptyPart()); }), /damaged/); // 5 partitions for layout '4'
  assert.match(await tamper(s => { lsw(s).layout = '5'; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).layout = 'constructor'; }), /damaged/);
  assert.match(await tamper(s => { s.rows[0].columns[0].name = ['<b>x</b>']; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].full = 0; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].full = '<b>40</b>'; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].counts = 'x'; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].counts[0].stock = -1; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].counts[0].at = '<img src=x>'; }), /damaged/);
  assert.match(await tamper(s => { lsw(s).parts[0].counts = Array(11).fill(lsw(s).parts[0].counts[0]); }), /damaged/);
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
    L.setLayout(lvl, '4');
    lvl.parts = lvl.parts.map((_, p) => P('LSW', String(c).padStart(3, '0'), 'C' + p, c * 100 + p));
  }));
  const t = Date.now();
  const buf = await exportWorkbook(ExcelJS, big);
  const ms = Date.now() - t;
  assert.ok(ms < 3000, `exporting ${L.summary(big).products} products took ${ms} ms`);
  assert.deepEqual(await importWorkbook(ExcelJS, buf), { state: big, skipped: [] });
}

// ===== Service worker =====
// An update must cache fresh files, not the browser's HTTP-cached ones (GitHub Pages caches for 10 minutes).
{
  const handlers = {}, realRequest = globalThis.Request;
  let cached, installed;
  globalThis.self = { addEventListener: (type, fn) => { handlers[type] = fn; }, skipWaiting() {} };
  globalThis.caches = { open: async () => ({ addAll: async reqs => { cached = reqs; } }) };
  globalThis.Request = class { constructor(url, init = {}) { this.url = url; this.cache = init.cache; } }; // Node's Request rejects relative URLs
  await import('./sw.js');
  handlers.install({ waitUntil: p => { installed = p; } });
  await installed;
  assert.ok(cached.length >= 10, 'all app files cached');
  assert.ok(cached.every(r => r instanceof globalThis.Request && r.cache === 'reload'), 'every file fetched with cache: reload');
  globalThis.Request = realRequest;
  delete globalThis.self;
  delete globalThis.caches;
}

console.log('all tests passed');

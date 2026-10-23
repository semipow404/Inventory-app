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

// --- List partitions: a partition can hold any number of products, each with its own count
{
  const T1 = new Date('2026-10-23T09:00:00Z'), T2 = new Date('2026-10-23T15:00:00Z');
  let items = [];
  ({ items } = L.addToList(items, { style: 'LSW', number: ' 017 ', color: 'wc' }));
  ({ items } = L.addToList(items, { style: 'LSW', number: '022', color: 'WCBF' }));
  assert.deepEqual(items.map(L.code), ['LSW017WC', 'LSW022WCBF']);
  assert.match(L.addToList(items, { style: 'LSW', number: '022', color: 'wcbf' }).error, /already in the list/);
  assert.match(L.addToList(items, { style: '', number: ' ', color: '' }).error, /code/);
  // Saving: the name fields become the partition's default; a typed stock counts that product, a blank box leaves it alone
  const { part } = L.makeList({ style: 'LSW', number: '022', color: 'wcbf' }, items, ['7', ''], T1);
  assert.ok(L.isList(part));
  assert.deepEqual([part.style, part.number, part.color, part.stock, part.full], ['LSW', '022', 'WCBF', null, null]);
  assert.deepEqual(part.list[0], P('LSW', '017', 'WC', 7, { counts: [{ at: T1.toISOString(), stock: 7 }] }));
  assert.deepEqual(part.list[1], P('LSW', '022', 'WCBF', null));
  const again = L.makeList({ style: 'LSW', number: '022', color: 'WCBF' }, part.list, ['', '3'], T2).part;
  assert.deepEqual([again.list[0].stock, again.list[0].counts.length], [7, 1]);
  assert.equal(L.lastCounted(again.list[1]), T2.toISOString());
  assert.match(L.makeList({ style: '', number: '', color: '' }, items, ['x', '']).error, /whole number/);
  // On the map: "LIST" and a product count; no usage colour, graph or own count; search looks inside
  assert.equal(L.productCount(2), '2 products');
  assert.equal(L.productCount(1), '1 product');
  assert.ok(L.hasProduct(again));
  assert.ok(!L.hasProduct({ ...again, list: [] }));
  assert.ok(L.matches(again, { style: '', number: '017', color: '' }));
  assert.ok(!L.matches(again, { style: '', number: '999', color: '' }));
  assert.equal(L.usage(again), null);
  assert.deepEqual(L.stockLine(again), []);
  assert.equal(L.lastCounted(again), null);
  // One count line per product, all carrying the partition's number; a list moves as a whole
  const st = L.newState();
  st.rows.push(L.newRow('R'));
  const lv = st.rows[0].columns[0].levels[0];
  L.setLayout(lv, '2II');
  lv.parts[0] = P('LAN', '042', 'W', 5);
  lv.parts[1] = again;
  const lines = () => L.allParts(st).filter(x => L.hasProduct(x.part)).map(x => `${x.box}/${x.partition}/${L.code(x.part)}`);
  assert.deepEqual(lines(), ['a1/1/LAN042W', 'a1/2/LSW017WC', 'a1/2/LSW022WCBF']);
  assert.deepEqual(L.summary(st), { rows: 1, stillages: 3, products: 3 });
  L.swapParts(st, { r: 0, c: 0, l: 0, p: 0 }, { r: 0, c: 0, l: 0, p: 1 });
  assert.deepEqual(lines(), ['a1/1/LSW017WC', 'a1/1/LSW022WCBF', 'a1/2/LAN042W']);
  // Yesterday's stillage-wide lists become one list partition, keeping their default and products
  const old = { styles: [], rows: [{ name: 'R', columns: [{ type: 'stillage', levels: [
    { layout: 'List', defaults: { style: 'LSW', number: '', color: 'WK' }, parts: [P('LSW', '012', 'WK', 3)] },
    { layout: '1', parts: [L.emptyPart()] }, { layout: '1', parts: [L.emptyPart()] }] }] }] };
  const conv = L.migrate(old).rows[0].columns[0].levels[0];
  assert.equal(conv.layout, '1');
  assert.equal(conv.defaults, undefined);
  assert.deepEqual(conv.parts, [{ ...L.emptyPart(), style: 'LSW', color: 'WK', list: [P('LSW', '012', 'WK', 3)] }]);
}

// --- Yard: White Bulk and Other Bulk, each rows of product cards counted in stillages to the quarter
{
  const s = L.newState();
  assert.deepEqual(s.yard, { white: [[]], other: [[]] });
  assert.deepEqual(L.YARD_SECTIONS, [['white', 'White Bulk'], ['other', 'Other Bulk']]);
  const { card } = L.makeYardCard({ style: 'LSW', number: ' 017 ', color: 'w' });
  assert.deepEqual(card, { style: 'LSW', number: '017', color: 'W', count: 0 });
  assert.match(L.makeYardCard({ style: '', number: ' ', color: '' }).error, /code/);
  let n = 0;
  for (const d of [1, 1, 0.5, -0.25, 0.25, -1]) n = L.stepCount(n, d);
  assert.equal(n, 1.5);
  assert.equal(L.stepCount(0.25, -0.5), 0);                    // never below 0
  let q = 0;
  for (let i = 0; i < 10; i++) q = L.stepCount(q, 0.25);
  assert.equal(q, 2.5);                                         // quarters add up exactly
  // Lines for the Excel Yard tab: white first, rows numbered from 1
  s.yard.white[0].push({ ...card, count: 3 }, { style: 'LSW', number: '022', color: 'W', count: 2 });
  s.yard.other = [[], [{ style: 'LSW', number: '033', color: 'KK', count: 1.5 }]];
  assert.deepEqual(L.yardLines(s), [
    { section: 'White Bulk', row: 1, product: 'LSW017W', stillages: 3 },
    { section: 'White Bulk', row: 1, product: 'LSW022W', stillages: 2 },
    { section: 'Other Bulk', row: 2, product: 'LSW033KK', stillages: 1.5 },
  ]);
  // Saves from before the yard existed get an empty one
  assert.deepEqual(L.migrate({ styles: [], rows: [] }).yard, { white: [[]], other: [[]] });
}

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
  for (let i = 0; i < 35; i++) many = L.makePart(fields(String(i), ''), many, new Date(T0.getTime() + i * 1000)).part;
  assert.equal(many.counts.length, 30);                                  // only the last 30 are kept
  assert.equal(many.counts[0].stock, 5);
  assert.equal(many.counts.at(-1).stock, 34);
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

// --- Count mode: a quick count keeps the product and its Full number, and stamps a count
{
  const before = P('LSW', '017', 'MC', 16, { full: 20, counts: [{ at: T0.toISOString(), stock: 16 }] });
  const T1 = new Date('2026-10-17T08:00:00Z');
  const { part } = L.countPart(before, ' 12 ', T1);
  assert.deepEqual(part, { ...before, stock: 12, counts: [...before.counts, { at: T1.toISOString(), stock: 12 }] });
  assert.equal(L.lastCounted(part), T1.toISOString());
  assert.match(L.countPart(before, '').error, /whole number/);
  assert.match(L.countPart(before, '1.5').error, /whole number/);
  assert.match(L.countPart(before, '-2').error, /whole number/);
}

// --- 30-day stock line: x runs from 0 (30 days ago) to 1 (now), with a dot at each count and flat after the latest
{
  const now = new Date('2026-10-31T12:00:00Z');
  const ago = d => new Date(now.getTime() - d * 864e5).toISOString();
  const round = line => line.map(p => ({ ...p, x: Math.round(p.x * 1000) / 1000 }));
  // Your example: 16, still 16 ten days later, about 15 ten days after that, 0 at the latest count
  const mc = P('LSW', '017', 'MC', 0, { full: 20, counts: [
    { at: ago(40), stock: 16 }, { at: ago(20), stock: 16 }, { at: ago(10), stock: 15 }, { at: ago(0), stock: 0 }] });
  assert.deepEqual(round(L.stockLine(mc, now)), [
    { x: 0, stock: 16, dot: false },                 // the stock when the 30 days began (counted 40 days ago)
    { x: 0.333, stock: 16, dot: true },
    { x: 0.667, stock: 15, dot: true },
    { x: 1, stock: 0, dot: true },
    { x: 1, stock: 0, dot: false },
  ]);
  // First ever count inside the window: the line starts there, not at the left edge
  const fresh = P('LAN', '042', 'W', 7, { counts: [{ at: ago(15), stock: 9 }, { at: ago(5), stock: 7 }] });
  assert.deepEqual(round(L.stockLine(fresh, now)), [
    { x: 0.5, stock: 9, dot: true }, { x: 0.833, stock: 7, dot: true }, { x: 1, stock: 7, dot: false }]);
  // Only counted long ago: a flat line at that level; never counted: no line
  assert.deepEqual(round(L.stockLine(P('', '', 'KK', 3, { counts: [{ at: ago(60), stock: 3 }] }), now)),
    [{ x: 0, stock: 3, dot: false }, { x: 1, stock: 3, dot: false }]);
  assert.deepEqual(L.stockLine(P('', '', 'KK', null), now), []);
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

// --- Moving: two partitions swap everything (code, stock, full, counts); into an empty spot it's simply a move
{
  const m = L.newState();
  m.rows.push(L.newRow('A'), L.newRow('B'));
  L.setLayout(m.rows[0].columns[0].levels[0], '2II');
  const prod = P('LSW', '022', 'BR', 38, { full: 40, counts: [{ at: T0.toISOString(), stock: 38 }] });
  const other = P('LAN', '042', 'W', 5);
  m.rows[0].columns[0].levels[0].parts[0] = prod;
  m.rows[1].columns[0].levels[2].parts[0] = other;
  const a = { r: 0, c: 0, l: 0, p: 0 }, b = { r: 1, c: 0, l: 2, p: 0 }, empty = { r: 0, c: 0, l: 0, p: 1 };
  L.swapParts(m, a, b);                                          // across rows
  assert.deepEqual(m.rows[0].columns[0].levels[0].parts[0], other);
  assert.deepEqual(m.rows[1].columns[0].levels[2].parts[0], prod);
  L.swapParts(m, a, b);                                          // swapping the same two back undoes it
  assert.deepEqual(m.rows[0].columns[0].levels[0].parts[0], prod);
  L.swapParts(m, a, empty);
  assert.deepEqual(m.rows[0].columns[0].levels[0].parts, [L.emptyPart(), prod]);
  // Whole stillages swap too: layout and every partition, across rows
  const midA = m.rows[0].columns[0].levels[0], backC = m.rows[1].columns[0].levels[2];
  L.swapLevels(m, { r: 0, c: 0, l: 0 }, { r: 1, c: 0, l: 2 });
  assert.equal(m.rows[0].columns[0].levels[0], backC);
  assert.equal(m.rows[1].columns[0].levels[2], midA);
  assert.equal(m.rows[1].columns[0].levels[2].layout, '2II');
  L.swapLevels(m, { r: 0, c: 0, l: 0 }, { r: 1, c: 0, l: 2 }); // swapping back undoes it
  assert.equal(m.rows[0].columns[0].levels[0], midA);
}

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

// Map sheet: each stillage is a 12×12 block of Excel cells per level (the layout grid), from column B.
// Stock count sheet: title in row 1, header in row 2, one line per product from row 3, in columns A–E.
const G = 12, col = i => 2 + i * G;
const opened = async buf => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf); return wb; };
const line = (ws, n) => [1, 2, 3, 4, 5].map(c => ws.getCell(n, c).text);
const sheets = wb => wb.worksheets.map(s => s.name);

// --- Map sheet: row name, column numbers, then levels a/b/c of 12 Excel rows each
{
  const wb = await opened(await exportWorkbook(ExcelJS, sample(), { date: new Date(2026, 9, 15) }));
  assert.deepEqual(sheets(wb), ['Stillage map', 'Stock count', 'Yard', '_layout']);
  assert.equal(wb.getWorksheet('_layout').state, 'hidden');
  const ws = wb.getWorksheet('Stillage map');
  const at = (r, c) => ws.getCell(r, c).text;
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
  // --- Stock count sheet: walking order, blank stock for uncounted
  const sc = wb.getWorksheet('Stock count');
  assert.ok(sc.getCell(1, 1).text.includes('2026-10-15'));
  assert.deepEqual(line(sc, 2), ['Row', 'Box', 'Partition', 'Code', 'Stock']);
  assert.deepEqual(line(sc, 3), ['Mid Row', 'a1', '1', 'LSW022BR', '10']);
  assert.deepEqual(line(sc, 4), ['Mid Row', 'a1', '4', 'LSW021WKOF', '']);
  assert.deepEqual(line(sc, 5), ['Mid Row', 'c2', '3', 'LAN911FG', '0']);
  assert.deepEqual(line(sc, 6), ['Back Row', 'a1', '1', 'LAN042W', '7']);
  assert.deepEqual(line(sc, 7), ['Back Row', 'b1', '16', 'KK', '11']);
  assert.equal(sc.getCell(8, 1).text, '');
  // Full numbers and count dates stay out of the visible sheets: no extra column, map shows codes only
  assert.equal(sc.getCell(2, 6).text, '');
  assert.equal(sc.getCell(3, 6).text, '');
  for (const sheet of [ws, sc]) sheet.eachRow(r => r.eachCell(c => { // (ExcelJS's .text throws on an empty merged cell)
    if (c.value != null) assert.ok(!/2026-10-16|\/ 40/.test(c.text), `visible cell ${sheet.name}!${c.address} holds usage data`);
  }));
}

// --- Export options: only the ticked sheets are written; either file still imports in full (layout travels hidden)
{
  const mapOnly = await opened(await exportWorkbook(ExcelJS, sample(), { count: false }));
  assert.deepEqual(sheets(mapOnly), ['Stillage map', 'Yard', '_layout']);
  const countOnly = await opened(await exportWorkbook(ExcelJS, sample(), { map: false }));
  assert.deepEqual(sheets(countOnly), ['Stock count', 'Yard', '_layout']);
  for (const wb of [mapOnly, countOnly]) {
    assert.deepEqual(await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer()), { state: sample(), skipped: [] });
  }
}

// --- Round trip with edits a user might make in Excel (Review Focus 2 and 3)
{
  const wb = await opened(await exportWorkbook(ExcelJS, sample(), { chunk: 7 })); // tiny chunks split the JSON across many cells
  const sc = wb.getWorksheet('Stock count');
  sc.getCell(3, 5).value = 99;                 // LSW022BR: new count
  sc.getCell(5, 5).value = 'abc';              // LAN911FG: not a number → skipped, keeps 0
  sc.getCell(6, 1).value = '  back row ';      // LAN042W: different case/spaces still match
  sc.getCell(6, 2).value = 'A1';
  sc.getCell(6, 5).value = null;               // cleared → not counted
  sc.getCell(8, 1).value = 'Mid Row';          // a line for a box that doesn't exist
  sc.getCell(8, 2).value = 'b9';
  sc.getCell(8, 3).value = 1;
  sc.getCell(8, 5).value = 5;
  const out = await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer());
  const expected = sample();
  expected.rows[0].columns[1].levels[0].parts[0].stock = 99;
  expected.rows[1].columns[0].levels[0].parts[0].stock = null;
  assert.deepEqual(out.state, expected);
  assert.equal(out.skipped.length, 2);
  assert.match(out.skipped.join('\n'), /abc/);
  assert.match(out.skipped.join('\n'), /b9/);
}

// --- List partitions in Excel: "LIST · n products" on the map, a line per product, matched by code on import
{
  const s = L.newState();
  s.styles = ['LSW'];
  s.rows.push(L.newRow('List Row'));
  const lv = s.rows[0].columns[0].levels[0];
  L.setLayout(lv, '2II');
  lv.parts[0] = P('LAN', '042', 'W', 5);
  lv.parts[1] = { ...P('LSW', '', 'WC', null), list: [P('LSW', '017', 'WC', 3), P('LSW', '022', 'WCBF', null)] };
  const wb = await opened(await exportWorkbook(ExcelJS, s));
  const map = wb.getWorksheet('Stillage map'), sc = wb.getWorksheet('Stock count');
  assert.equal(map.getCell(3, col(0) + 6).text, 'LIST · 2 products');
  assert.deepEqual([3, 4, 5].map(r => line(sc, r)), [
    ['List Row', 'a1', '1', 'LAN042W', '5'], ['List Row', 'a1', '2', 'LSW017WC', '3'], ['List Row', 'a1', '2', 'LSW022WCBF', '']]);
  sc.getCell(5, 5).value = 8;                      // LSW022WCBF counted in Excel
  sc.getCell(4, 4).value = 'LSW999WC';             // a changed code no longer matches
  const out = await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer());
  const expected = structuredClone(s);
  expected.rows[0].columns[0].levels[0].parts[1].list[1].stock = 8;
  assert.deepEqual(out.state, expected);
  assert.equal(out.skipped.length, 1);
  assert.match(out.skipped[0], /LSW999WC/);
  // Bad list data in the hidden sheet is rejected like any other
  const tamperList = async mutate => {
    const w = await opened(await exportWorkbook(ExcelJS, s));
    const meta = w.getWorksheet('_layout');
    const data = JSON.parse(meta.getCell(2, 1).value);
    mutate(data.rows[0].columns[0].levels[0].parts[1]);
    meta.getCell(2, 1).value = JSON.stringify(data);
    return (await importWorkbook(ExcelJS, await w.xlsx.writeBuffer())).error ?? 'accepted';
  };
  assert.equal(await tamperList(() => {}), 'accepted');
  assert.equal(await tamperList(p => { p.list = []; }), 'accepted');               // an empty list is fine
  assert.match(await tamperList(p => { p.list = 'x'; }), /damaged/);
  assert.match(await tamperList(p => { p.list[0].stock = '<b>1</b>'; }), /damaged/);
  assert.match(await tamperList(p => { p.list[0].list = []; }), /damaged/);        // no lists inside lists
}

// --- Yard tab: one line per product card, white first; the yard itself comes back on import from the hidden data
{
  const s = sample();
  s.yard.white[0].push({ style: 'LSW', number: '017', color: 'W', count: 3 });
  s.yard.other[0].push({ style: 'LSW', number: '033', color: 'KK', count: 1.5 });
  const wb = await opened(await exportWorkbook(ExcelJS, s, { date: new Date(2026, 9, 24) }));
  assert.deepEqual(sheets(wb), ['Stillage map', 'Stock count', 'Yard', '_layout']);
  const y = wb.getWorksheet('Yard');
  assert.ok(y.getCell(1, 1).text.includes('2026-10-24'));
  assert.deepEqual([2, 3, 4, 5].map(r => [1, 2, 3, 4].map(c => y.getCell(r, c).text)), [
    ['Section', 'Row', 'Product', 'Stillages'], ['White Bulk', '1', 'LSW017W', '3'], ['Other Bulk', '1', 'LSW033KK', '1.5'], ['', '', '', '']]);
  assert.deepEqual(await importWorkbook(ExcelJS, await wb.xlsx.writeBuffer()), { state: s, skipped: [] });
  assert.deepEqual(sheets(await opened(await exportWorkbook(ExcelJS, s, { yard: false }))), ['Stillage map', 'Stock count', '_layout']);
  // Bad yard data in the hidden sheet is rejected like any other
  const tamperYard = async mutate => {
    const w = await opened(await exportWorkbook(ExcelJS, s));
    const meta = w.getWorksheet('_layout');
    const data = JSON.parse(meta.getCell(2, 1).value);
    mutate(data.yard);
    meta.getCell(2, 1).value = JSON.stringify(data);
    return (await importWorkbook(ExcelJS, await w.xlsx.writeBuffer())).error ?? 'accepted';
  };
  assert.equal(await tamperYard(() => {}), 'accepted');
  assert.match(await tamperYard(y => { y.white[0][0].count = -1; }), /damaged/);
  assert.match(await tamperYard(y => { y.white[0][0].count = 0.3; }), /damaged/);
  assert.match(await tamperYard(y => { y.other = 'x'; }), /damaged/);
  assert.match(await tamperYard(y => { y.white[0][0].number = ['<b>'] ; }), /damaged/);
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
  assert.match(await tamper(s => { lsw(s).parts[0].counts = Array(31).fill(lsw(s).parts[0].counts[0]); }), /damaged/);
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

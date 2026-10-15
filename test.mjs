import assert from 'node:assert/strict';
import * as L from './logic.js';

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

console.log('logic tests passed');

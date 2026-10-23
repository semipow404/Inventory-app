// Pure warehouse logic: no DOM, no storage. Used by app.js, excel.js and test.mjs.
export const LEVELS = ['a', 'b', 'c'];

// Partition layouts in dropdown order. Each block is [x, y, w, h] on a 12×12 grid (12 divides by 2, 3 and 4),
// numbered left→right then top→bottom. The app and the Excel map both draw from this.
const grid = (cols, rows) => Array.from({ length: cols * rows }, (_, i) =>
  [(i % cols) * (12 / cols), Math.floor(i / cols) * (12 / rows), 12 / cols, 12 / rows]);
export const LAYOUTS = new Map([
  ['1', grid(1, 1)],
  ['2II', grid(2, 1)],
  ['2=', grid(1, 2)],
  ['3II', grid(3, 1)],
  ['3T', [[0, 0, 12, 6], [0, 6, 6, 6], [6, 6, 6, 6]]],
  ['4', grid(2, 2)],
  ['4T', [[0, 0, 12, 6], [0, 6, 4, 6], [4, 6, 4, 6], [8, 6, 4, 6]]],
  ['6', grid(3, 2)],
  ['9', grid(3, 3)],
  ['12', grid(4, 3)],
  ['16', grid(4, 4)],
]);

export const newState = () => ({ styles: [], rows: [], yard: emptyYard() });

// The yard: bulk stillages waiting to come into the warehouse. Two fixed sections, each a list of rows of cards;
// a card is a product code and how many stillages of it there are, in quarters.
export const YARD_SECTIONS = [['white', 'White Bulk'], ['other', 'Other Bulk']];
const emptyYard = () => ({ white: [[]], other: [[]] });

export function makeYardCard(fields) {
  const card = { ...cleanCode(fields), count: 0 };
  return code(card) ? { card } : { error: 'Enter a code.' };
}

// +1 / −1 / +½ / −½ / +¼ / −¼, kept to exact quarters and never below 0.
export const stepCount = (count, delta) => Math.max(0, Math.round((count + delta) * 4) / 4);

export const yardLines = state => YARD_SECTIONS.flatMap(([key, section]) => state.yard[key].flatMap((row, r) =>
  row.map(card => ({ section, row: r + 1, product: code(card), stillages: card.count }))));
// full: what "full" means for this partition (for the usage colours). counts: the last COUNTS_KEPT counts, newest last.
export const emptyPart = () => ({ style: '', number: '', color: '', stock: null, full: null, counts: [] });
export const COUNTS_KEPT = 30;
const addCount = (counts, stock, now) => [...counts, { at: now.toISOString(), stock }].slice(-COUNTS_KEPT);
export const newStillage = () => ({ type: 'stillage', levels: LEVELS.map(() => ({ layout: '1', parts: [emptyPart()] })) });
export const newRow = name => ({ name: name.trim(), columns: [newStillage()] });

export const code = p => p.style + p.number + p.color;
// A list partition holds any number of products in `list` (each shaped like a partition, without a list of its own);
// its own style / number / colour are the default the pop-up starts from, and it has no stock, Full or counts.
export const isList = p => Array.isArray(p.list);
export const hasProduct = p => (isList(p) ? p.list.length > 0 : code(p) !== '');
export const productCount = n => `${n} product${n === 1 ? '' : 's'}`;

// 1-based stillage number for each column of a row; null for walkways.
export function columnNumbers(row) {
  let n = 0;
  return row.columns.map(c => (c.type === 'stillage' ? ++n : null));
}

// Every partition in walking order: row, column, level (a→c), partition. A list partition gives one entry per
// product instead (with its index `i`), so tables, search and imports see each product.
export function allParts(state) {
  const out = [];
  state.rows.forEach((row, r) => {
    const nums = columnNumbers(row);
    row.columns.forEach((col, c) => {
      if (col.type !== 'stillage') return;
      col.levels.forEach((lvl, l) => lvl.parts.forEach((part, p) => {
        const at = { r, c, l, p, row: row.name, box: LEVELS[l] + nums[c], partition: p + 1 };
        if (isList(part)) part.list.forEach((item, i) => out.push({ ...at, i, part: item }));
        else out.push({ ...at, part });
      }));
    });
  });
  return out;
}

export const isSearching = q => Boolean(q.style || q.number || q.color);

// Each filled search box must equal that part of the code, ignoring case; empty boxes are ignored.
export function matches(part, q) {
  if (isList(part)) return part.list.some(x => matches(x, q));
  const eq = (value, wanted) => !wanted || value.toUpperCase() === wanted.toUpperCase();
  return isSearching(q) && hasProduct(part) && eq(part.style, q.style) && eq(part.number, q.number) && eq(part.color, q.color);
}

// Error message, or null if usable. `except` is the row being renamed. Import matches rows by name, so names must be unique.
export function rowNameError(state, name, except = null) {
  const wanted = name.trim().toUpperCase();
  if (!wanted) return 'Enter a row name.';
  return state.rows.some(r => r !== except && r.name.toUpperCase() === wanted) ? 'A row with that name already exists.' : null;
}

export function addColumn(row, type, side) {
  const col = type === 'space' ? { type: 'space', name: 'Walkway' } : newStillage();
  if (side === 'left') row.columns.unshift(col);
  else row.columns.push(col);
}

export const dropsProducts = (level, n) => level.parts.slice(n).some(hasProduct);

// Moving a product: the two partitions ({r, c, l, p}) swap everything they hold. Into an empty spot, that's a move.
export function swapParts(state, a, b) {
  const parts = ({ r, c, l }) => state.rows[r].columns[c].levels[l].parts;
  [parts(a)[a.p], parts(b)[b.p]] = [parts(b)[b.p], parts(a)[a.p]];
}

// Moving a whole stillage: two levels ({r, c, l}) swap their layout and everything in them.
export function swapLevels(state, a, b) {
  const levels = ({ r, c }) => state.rows[r].columns[c].levels;
  [levels(a)[a.l], levels(b)[b.l]] = [levels(b)[b.l], levels(a)[a.l]];
}

// Keeps partitions by number: switching between layouts with the same count keeps every product.
export function setLayout(level, id) {
  const n = LAYOUTS.get(id).length;
  while (level.parts.length < n) level.parts.push(emptyPart());
  level.parts.length = n;
  level.layout = id;
}

// Fills in fields added after the first release so older saves and exports still load.
// Anything malformed is left as it is for import validation to reject.
const FIRST_LAYOUTS = ['1', '2II', '3II', '4']; // the first release had only these, picked by partition count
export function migrate(state) {
  const list = v => (Array.isArray(v) ? v : []);
  if (state && typeof state === 'object' && state.yard === undefined) state.yard = emptyYard();
  for (const row of list(state?.rows)) for (const col of list(row?.columns)) {
    if (col?.type === 'space' && col.name === undefined) col.name = 'Walkway';
    if (col?.type === 'stillage') for (const lvl of list(col.levels)) {
      if (lvl?.layout === 'List' && Array.isArray(lvl.parts)) { // the short-lived stillage-wide list → one list partition
        Object.assign(lvl, { layout: '1', parts: [{ ...emptyPart(), ...lvl.defaults, list: lvl.parts }] });
        delete lvl.defaults;
      }
      if (lvl && lvl.layout === undefined && Array.isArray(lvl.parts)) lvl.layout = FIRST_LAYOUTS[lvl.parts.length - 1];
      for (const part of list(lvl?.parts)) for (const p of [part, ...list(part?.list)]) {
        if (p && p.full === undefined) p.full = null;
        if (p && p.counts === undefined) p.counts = [];
      }
    }
  }
  return state;
}

// Turns the product pop-up's raw field values into a partition, or an error message. Any part of the code may be blank.
// Saving a stock number is a count, even if the number didn't change: it's added to `prev`'s counts, stamped `now`.
const cleanCode = ({ style, number, color }) => ({ style, number: number.trim().toUpperCase(), color: color.trim().toUpperCase() });

export function makePart({ style, number, color, stock, full }, prev = emptyPart(), now = new Date()) {
  const p = { ...cleanCode({ style, number, color }), stock: stock.trim(), full: full.trim() };
  if (!p.style && !p.number && !p.color) return { part: emptyPart() };
  if (p.stock !== '' && !/^\d+$/.test(p.stock)) return { error: 'Stock must be a whole number, 0 or more.' };
  if (p.full !== '' && !/^[1-9]\d*$/.test(p.full)) return { error: 'Full must be a whole number, 1 or more.' };
  const s = p.stock === '' ? null : Number(p.stock);
  const counts = s === null ? prev.counts : addCount(prev.counts, s, now);
  return { part: { ...p, stock: s, full: p.full === '' ? null : Number(p.full), counts } };
}

// List partitions in the pop-up: the name fields add a product to the list being edited.
export function addToList(items, fields) {
  const item = { ...emptyPart(), ...cleanCode(fields) };
  if (!hasProduct(item)) return { error: 'Enter a code to add.' };
  if (items.some(x => code(x) === code(item))) return { error: `${code(item)} is already in the list.` };
  return { items: [...items, item] };
}

// Saving a list partition: the name fields become its default. typed[i] is items[i]'s stock box: a number counts
// that product (stamped `now`), a blank box leaves it as it was.
export function makeList(fields, items, typed, now = new Date()) {
  const box = i => (typed[i] ?? '').trim();
  if (items.some((_, i) => box(i) !== '' && !/^\d+$/.test(box(i)))) return { error: 'Stock must be a whole number, 0 or more.' };
  const list = items.map((x, i) => (box(i) === '' ? x : { ...x, stock: Number(box(i)), counts: addCount(x.counts, Number(box(i)), now) }));
  return { part: { ...emptyPart(), ...cleanCode(fields), list } };
}

// Count mode: a new stock number for an existing product; everything else stays as it is.
export function countPart(part, stock, now = new Date()) {
  stock = stock.trim();
  if (!/^\d+$/.test(stock)) return { error: 'Enter the count as a whole number, 0 or more.' };
  return { part: { ...part, stock: Number(stock), counts: addCount(part.counts, Number(stock), now) } };
}

// The partition's stock over the last `days`: points { x, stock, dot } with x from 0 (`days` ago) to 1 (now).
// It opens at the last count before the window (if any), has a dot per count inside it, and runs flat to now.
export function stockLine(part, now = new Date(), days = 30) {
  const end = now.getTime(), start = end - days * 864e5;
  const counts = part.counts.map(c => ({ t: Date.parse(c.at), stock: c.stock })).filter(c => c.t <= end);
  const before = counts.filter(c => c.t <= start).at(-1);
  const line = before ? [{ x: 0, stock: before.stock, dot: false }] : [];
  for (const c of counts) if (c.t > start) line.push({ x: (c.t - start) / (end - start), stock: c.stock, dot: true });
  if (line.length) line.push({ x: 1, stock: line.at(-1).stock, dot: false });
  return line;
}

// When this partition was last counted (ISO string), or null if its stock is blank.
export const lastCounted = p => (p.stock === null ? null : p.counts.at(-1)?.at ?? null);

// How full the partition is, 0–1 (over full counts as 1); null without a count or a "full" number.
export const usage = p => (p.stock === null || !p.full ? null : Math.min(p.stock / p.full, 1));

export function addStyle(state, s) {
  s = s.trim().toUpperCase();
  if (!s) return 'Enter a style.';
  if (state.styles.includes(s)) return 'That style is already in the list.';
  state.styles.push(s);
  return null;
}

// Dropdown options: blank, the Settings list, plus the current style if it has since been removed from the list.
export const styleChoices = (styles, current) =>
  ['', ...styles, ...(current && !styles.includes(current) ? [current] : [])];

export function summary(state) {
  return {
    rows: state.rows.length,
    stillages: state.rows.reduce((n, row) => n + row.columns.filter(c => c.type === 'stillage').length * LEVELS.length, 0),
    products: allParts(state).filter(x => hasProduct(x.part)).length,
  };
}

// Local date as YYYY-MM-DD (toISOString would give the UTC date, off by one near midnight).
export const ymd = (d = new Date()) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

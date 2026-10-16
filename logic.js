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

export const newState = () => ({ styles: [], rows: [] });
// full: what "full" means for this partition (for the usage colours). counts: the last COUNTS_KEPT counts, newest last.
export const emptyPart = () => ({ style: '', number: '', color: '', stock: null, full: null, counts: [] });
export const COUNTS_KEPT = 10;
export const newStillage = () => ({ type: 'stillage', levels: LEVELS.map(() => ({ layout: '1', parts: [emptyPart()] })) });
export const newRow = name => ({ name: name.trim(), columns: [newStillage()] });

export const code = p => p.style + p.number + p.color;
export const hasProduct = p => code(p) !== '';

// 1-based stillage number for each column of a row; null for walkways.
export function columnNumbers(row) {
  let n = 0;
  return row.columns.map(c => (c.type === 'stillage' ? ++n : null));
}

// Every partition in walking order: row, column, level (a→c), partition.
export function allParts(state) {
  const out = [];
  state.rows.forEach((row, r) => {
    const nums = columnNumbers(row);
    row.columns.forEach((col, c) => {
      if (col.type !== 'stillage') return;
      col.levels.forEach((lvl, l) => lvl.parts.forEach((part, p) =>
        out.push({ r, c, l, p, row: row.name, box: LEVELS[l] + nums[c], partition: p + 1, part })));
    });
  });
  return out;
}

export const isSearching = q => Boolean(q.style || q.number || q.color);

// Each filled search box must equal that part of the code, ignoring case; empty boxes are ignored.
export function matches(part, q) {
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
  for (const row of list(state?.rows)) for (const col of list(row?.columns)) {
    if (col?.type === 'space' && col.name === undefined) col.name = 'Walkway';
    if (col?.type === 'stillage') for (const lvl of list(col.levels)) {
      if (lvl && lvl.layout === undefined && Array.isArray(lvl.parts)) lvl.layout = FIRST_LAYOUTS[lvl.parts.length - 1];
      for (const part of list(lvl?.parts)) {
        if (part && part.full === undefined) part.full = null;
        if (part && part.counts === undefined) part.counts = [];
      }
    }
  }
  return state;
}

// Turns the product pop-up's raw field values into a partition, or an error message. Any part of the code may be blank.
// Saving a stock number is a count, even if the number didn't change: it's added to `prev`'s counts, stamped `now`.
export function makePart({ style, number, color, stock, full }, prev = emptyPart(), now = new Date()) {
  const p = { style, number: number.trim().toUpperCase(), color: color.trim().toUpperCase(), stock: stock.trim(), full: full.trim() };
  if (!p.style && !p.number && !p.color) return { part: emptyPart() };
  if (p.stock !== '' && !/^\d+$/.test(p.stock)) return { error: 'Stock must be a whole number, 0 or more.' };
  if (p.full !== '' && !/^[1-9]\d*$/.test(p.full)) return { error: 'Full must be a whole number, 1 or more.' };
  const s = p.stock === '' ? null : Number(p.stock);
  const counts = s === null ? prev.counts : [...prev.counts, { at: now.toISOString(), stock: s }].slice(-COUNTS_KEPT);
  return { part: { ...p, stock: s, full: p.full === '' ? null : Number(p.full), counts } };
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
  const parts = allParts(state);
  return {
    rows: state.rows.length,
    stillages: parts.filter(x => x.p === 0).length,
    products: parts.filter(x => hasProduct(x.part)).length,
  };
}

// Local date as YYYY-MM-DD (toISOString would give the UTC date, off by one near midnight).
export const ymd = (d = new Date()) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

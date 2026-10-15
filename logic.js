// Pure warehouse logic: no DOM, no storage. Used by app.js, excel.js and test.mjs.
export const LEVELS = ['a', 'b', 'c'];

export const newState = () => ({ styles: [], rows: [] });
export const emptyPart = () => ({ style: '', number: '', color: '', stock: null });
export const newStillage = () => ({ type: 'stillage', levels: LEVELS.map(() => ({ parts: [emptyPart()] })) });
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
  const col = type === 'space' ? { type: 'space' } : newStillage();
  if (side === 'left') row.columns.unshift(col);
  else row.columns.push(col);
}

export const dropsProducts = (level, n) => level.parts.slice(n).some(hasProduct);

export function setPartitions(level, n) {
  while (level.parts.length < n) level.parts.push(emptyPart());
  level.parts.length = n;
}

// Turns the product pop-up's raw field values into a partition, or an error message.
export function makePart({ style, number, color, stock }) {
  const p = { style, number: number.trim().toUpperCase(), color: color.trim().toUpperCase(), stock: stock.trim() };
  if (!p.style && !p.number && !p.color) return { part: emptyPart() };
  if (!p.style) return { error: 'Pick a style.' };
  if (p.stock !== '' && !/^\d+$/.test(p.stock)) return { error: 'Stock must be a whole number, 0 or more.' };
  return { part: { ...p, stock: p.stock === '' ? null : Number(p.stock) } };
}

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

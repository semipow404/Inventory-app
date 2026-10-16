// UI: renders the view chosen by the URL hash into #view and routes taps to `actions`.
// Rules live in logic.js; Excel lives in excel.js.
import * as L from './logic.js';
import { exportWorkbook, importWorkbook } from './excel.js';

const KEY = 'inventory-app-v1';
const $ = sel => document.querySelector(sel);
const esc = s => String(s).replace(/[&<>"']/g, ch => `&#${ch.charCodeAt(0)};`);

let state = load();

function load() {
  try { return L.migrate(JSON.parse(localStorage.getItem(KEY))) ?? L.newState(); } catch { return L.newState(); }
}

function save() {
  try { localStorage.setItem(KEY, JSON.stringify(state)); $('#banner').hidden = true; }
  catch { $('#banner').hidden = false; }
}

function commit() { save(); render(); }

// views: name → (...args) => { title, html, search }. actions: data-action name → (dataset, element).
const views = {};
const actions = { back: () => history.back() };

// '' = home, #warehouse, #stillage/r/c/l, #settings. Using the hash keeps the phone's back button working.
function route() {
  const [name, ...args] = location.hash.slice(1).split('/');
  return views[name] ? { name, args: args.map(Number) } : { name: 'home', args: [] };
}

const styleOptions = selected => L.styleChoices(state.styles, selected)
  .map(s => `<option${s === selected ? ' selected' : ''}>${esc(s)}</option>`).join('');

// Scroll positions survive re-renders and trips to the close-up and back.
const scrollMemory = new Map();

function render() {
  document.querySelectorAll('[data-scroll]').forEach(el => scrollMemory.set(el.dataset.scroll, [el.scrollLeft, el.scrollTop]));
  const { name, args } = route();
  const v = views[name](...args);
  $('#title').textContent = v.title;
  $('#back').hidden = name === 'home';
  $('#search').hidden = !v.search;
  if (v.search) $('#q-style').innerHTML = styleOptions($('#q-style').value);
  $('#view').innerHTML = v.html;
  document.querySelectorAll('[data-scroll]').forEach(el => {
    const saved = scrollMemory.get(el.dataset.scroll);
    if (saved) [el.scrollLeft, el.scrollTop] = saved;
  });
}

document.addEventListener('click', e => {
  const el = e.target.closest('[data-action]');
  if (el && el.tagName !== 'FORM') actions[el.dataset.action]?.(el.dataset, el);
});
document.addEventListener('submit', e => {
  const form = e.target;
  if (!form.dataset.action) return; // dialog forms handle themselves
  e.preventDefault();
  actions[form.dataset.action](form.dataset, form);
});
window.addEventListener('hashchange', render);
// Another tab or the installed app saved: reload, so this page's older copy can't overwrite it on its next save.
window.addEventListener('storage', e => { if (e.key === KEY) { state = load(); render(); } });

// ---- Home and Settings
Object.assign(views, {
  home: () => ({ title: 'Home', html: `
    <nav class="home">
      <a class="big warehouse" href="#warehouse">Warehouse</a>
      <button class="big yard" disabled>Yard</button>
      <button class="big store" disabled>Store</button>
      <div class="pair">
        <button class="io" data-action="import">Import</button>
        <button class="io" data-action="export">Export</button>
      </div>
      <a class="big settings-btn" href="#settings">Settings</a>
    </nav>` }),

  settings: () => ({ title: 'Settings', html: `
    <section class="settings">
      <h2>Product styles</h2>
      <ul>${state.styles.map((s, i) => `
        <li>${esc(s)} <button data-action="remove-style" data-i="${i}" aria-label="Remove ${esc(s)}">✕</button></li>`).join('')
        || '<li class="muted">No styles yet</li>'}</ul>
      <form data-action="add-style">
        <input name="code" placeholder="e.g. LSW" autocomplete="off" autocapitalize="characters">
        <button>Add</button>
      </form>
    </section>` }),
});

Object.assign(actions, {
  'add-style'(_, form) {
    const error = L.addStyle(state, form.elements.code.value);
    if (error) return alert(error);
    commit();
  },
  'remove-style'({ i }) { state.styles.splice(Number(i), 1); commit(); },
});

// ---- Dialogs
// Shows a <dialog>; resolves with the pressed button's value ('save', 'delete' or 'cancel').
// `check` runs on Save; returning a message keeps the dialog open and shows it.
function ask(dialog, check = () => null) {
  const error = dialog.querySelector('.error');
  error.textContent = '';
  return new Promise(resolve => {
    // Resolve from the button press itself: some browsers never fire the dialog's 'close' event.
    dialog.querySelector('form').onsubmit = e => {
      e.preventDefault();
      const choice = e.submitter?.value || 'save';
      const msg = choice === 'save' ? check() : null;
      if (msg) return void (error.textContent = msg);
      dialog.close(choice);
      resolve(choice);
    };
    dialog.onclose = () => resolve('cancel'); // Escape key / back gesture; ignored if a button already resolved
    dialog.showModal();
  });
}

// ---- Warehouse map
const query = () => ({ style: $('#q-style').value, number: $('#q-number').value.trim(), color: $('#q-color').value.trim() });
$('#search').addEventListener('input', render);

// Stillage usage colours: a per-phone setting, so it lives outside the warehouse data (and out of exports).
const USAGE_KEY = 'inventory-app-usage';
let usageOn = (() => { try { return localStorage.getItem(USAGE_KEY) === '1'; } catch { return false; } })();
// Red (empty) → yellow → green (full); grey when not counted or no "full" number.
const usageColour = p => { const u = L.usage(p); return u === null ? '#2a2f45' : `hsl(${Math.round(120 * u)} 75% 40%)`; };

const shortDate = iso => (iso ? new Date(iso).toLocaleDateString(undefined, { day: 'numeric', month: 'short' }) : '');
const longDate = iso => new Date(iso).toLocaleString(undefined, { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

// The 30-day stock line drawn across a partition: 0 at the bottom, Full (or the highest count) at the top.
function graphHtml(p) {
  const line = L.stockLine(p);
  if (!line.length) return '';
  const top = p.full || Math.max(1, ...line.map(pt => pt.stock));
  const y = s => (96 - Math.min(s / top, 1) * 92).toFixed(1); // keeps 0 and full just inside the edges
  const x = pt => (pt.x * 100).toFixed(1);
  return `<svg class="graph" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><polyline points="${
    line.map(pt => `${x(pt)},${y(pt.stock)}`).join(' ')}"/></svg>${
    line.filter(pt => pt.dot).map(pt => `<i class="dot" style="left:${x(pt)}%;top:${y(pt.stock)}%"></i>`).join('')}`;
}

// Partitions sit on the same 12×12 grid the Excel map uses (L.LAYOUTS), as CSS grid areas.
// `graphs` adds the 30-day stock lines (the close-up, with usage on).
function boxHtml(lvl, label, q, boxAttrs, partAttrs = () => '', graphs = false) {
  const blocks = L.LAYOUTS.get(lvl.layout);
  return `<div class="box" ${boxAttrs}>
    <span class="label">${label}</span>
    <div class="parts${usageOn ? ' usage' : ''}" data-n="${lvl.parts.length}">${lvl.parts.map((p, i) => {
      const [x, y, w, h] = blocks[i];
      const area = `grid-area:${y + 1}/${x + 1}/span ${h}/span ${w}`;
      const body = usageOn && L.hasProduct(p)
        ? `${esc(L.code(p))}<b>${esc(p.stock ?? '–')}${p.full ? ` / ${esc(p.full)}` : ''}</b>`
        : esc(L.code(p) || 'NA');
      return `<div class="part${L.matches(p, q) ? ' hit' : ''}" style="${area}${usageOn ? `;background:${usageColour(p)}` : ''}" ${partAttrs(i)}>${
        graphs ? graphHtml(p) : ''}<span class="txt">${body}</span></div>`;
    }).join('')}</div>
  </div>`;
}

function listHtml(items, scrollKey) {
  return `<div class="list" data-scroll="${scrollKey}"><table>
    <thead><tr><th>Code</th><th>Location</th><th>Box</th><th>Stock</th><th>Counted</th></tr></thead>
    <tbody>${items.map(x => `<tr><td>${esc(L.code(x.part))}</td><td>${esc(x.row)}</td><td>${x.box}</td><td>${esc(x.part.stock ?? '')}</td><td>${shortDate(L.lastCounted(x.part))}</td></tr>`).join('')}</tbody>
  </table></div>`;
}

// Map modes. Move: first tap picks a product, the second swaps. Count: a tap opens the quick count pop-up.
// Edit: shows the controls that change the layout; off at every start so nobody changes it by accident.
let moving = false, counting = false, editing = false, picked = null; // picked: { r, c, l, p }
const samePart = (a, b) => a.r === b.r && a.c === b.c && a.l === b.l && a.p === b.p;
const partAt = ({ r, c, l, p }) => state.rows[r].columns[c].levels[l].parts[p];

function rowHtml(row, r, q) {
  const nums = L.columnNumbers(row);
  // What a tap does: open the close-up, or (in Move / Count mode) act on the partition itself.
  const taps = (c, l) => {
    const where = `data-r="${r}" data-c="${c}" data-l="${l}"`;
    if (moving) return ['', p => `role="button" data-action="pick" ${where} data-p="${p}"${picked && samePart(picked, { r, c, l, p }) ? ' data-picked' : ''}`];
    if (counting) return ['', p => `role="button" data-action="quick-count" ${where} data-p="${p}"`];
    return [`role="button" data-action="open" ${where}`, undefined];
  };
  const adders = side => (editing ? `<div class="adders">
      <button data-action="add-col" data-r="${r}" data-type="stillage" data-side="${side}">Add Stillage</button>
      <button data-action="add-col" data-r="${r}" data-type="space" data-side="${side}">Add Space</button>
    </div>` : '');
  const columns = row.columns.map((col, c) => col.type === 'space'
    ? (editing ? `<button class="walkway" data-action="edit-space" data-r="${r}" data-c="${c}">${esc(col.name)}</button>`
      : `<div class="walkway">${esc(col.name)}</div>`)
    : `<div class="col">${col.levels.map((lvl, l) => boxHtml(lvl, L.LEVELS[l] + nums[c], q, ...taps(c, l))).join('')}</div>`).join('');
  return `<section class="row">
    <h2>${editing ? `<button class="row-name" data-action="edit-row" data-r="${r}">${esc(row.name)}</button>` : esc(row.name)}</h2>
    <div class="row-scroll" data-scroll="row-${r}">${adders('left')}${columns}${adders('right')}</div>
  </section>`;
}

const modeButton = (action, label, on, cls) =>
  `<button class="mode-btn ${cls}${on ? ' on' : ''}" data-action="${action}" aria-pressed="${on}">${label}</button>`;

Object.assign(views, {
  warehouse() {
    const q = query();
    const items = L.allParts(state).filter(x => (L.isSearching(q) ? L.matches(x.part, q) : L.hasProduct(x.part)));
    const hint = moving ? (picked ? `Now tap where ${esc(L.code(partAt(picked)))} should go (tap it again to cancel)` : 'Tap a product to move')
      : counting ? 'Tap a product to count it' : '';
    const add = editing ? '<button class="add-row" data-action="new-row" aria-label="Add row">+</button>'
      : state.rows.length ? '' : '<p class="muted">Turn on Edit to add a row.</p>';
    return { title: 'Warehouse', search: true, html: `
      ${hint ? `<p class="move-hint">${hint}</p>` : ''}
      <div class="map${state.rows.length ? '' : ' empty'}" data-scroll="map">
        ${state.rows.map((row, r) => rowHtml(row, r, q)).join('')}
        ${add}
      </div>
      <div class="usage-bar">
        ${modeButton('toggle-usage', 'Usage', usageOn, 'green')}${modeButton('toggle-move', 'Move', moving, 'amber')}${
          modeButton('toggle-edit', 'Edit', editing, 'blue')}${modeButton('toggle-count', 'Count', counting, 'teal')}
      </div>${listHtml(items, 'list')}` };
  },
});

// One name pop-up for rows and spaces; a null deleteLabel hides its delete button.
const nameDialog = $('#name-dialog'), nameInput = $('#name-input'), countDialog = $('#count-dialog');
function askName(title, value, deleteLabel, check) {
  $('#name-title').textContent = title;
  const del = nameDialog.querySelector('[value=delete]');
  del.hidden = !deleteLabel;
  del.textContent = deleteLabel ?? '';
  nameInput.value = value;
  return ask(nameDialog, check);
}

Object.assign(actions, {
  async 'new-row'() {
    if (await askName('Row name', '', null, () => L.rowNameError(state, nameInput.value)) !== 'save') return;
    state.rows.push(L.newRow(nameInput.value));
    commit();
  },
  async 'edit-row'({ r }) {
    const row = state.rows[r];
    const choice = await askName('Row name', row.name, 'Delete row', () => L.rowNameError(state, nameInput.value, row));
    if (choice === 'save') row.name = nameInput.value.trim();
    else if (choice === 'delete' && confirm(`Delete "${row.name}" and everything in it?`)) state.rows.splice(Number(r), 1);
    else return;
    commit();
  },
  async 'edit-space'({ r, c }) {
    const space = state.rows[r].columns[c];
    const choice = await askName('Space name', space.name, 'Remove space');
    if (choice === 'save') space.name = nameInput.value.trim() || 'Walkway';
    else if (choice === 'delete') state.rows[r].columns.splice(Number(c), 1); // a space holds nothing, so no extra confirm
    else return;
    commit();
  },
  'add-col'({ r, type, side }) { L.addColumn(state.rows[r], type, side); commit(); },
  'toggle-usage'() {
    usageOn = !usageOn;
    try { localStorage.setItem(USAGE_KEY, usageOn ? '1' : '0'); } catch { /* still works for this visit */ }
    render();
  },
  open({ r, c, l }) { location.hash = `#stillage/${r}/${c}/${l}`; },
  'toggle-move'() { moving = !moving; if (moving) counting = false; picked = null; render(); },
  'toggle-count'() { counting = !counting; if (counting) moving = false; picked = null; render(); },
  'toggle-edit'() { editing = !editing; render(); },
  async 'quick-count'({ r, c, l, p }) {
    const at = { r: +r, c: +c, l: +l, p: +p };
    const part = partAt(at);
    if (!L.hasProduct(part)) return;
    $('#c-code').textContent = L.code(part);
    $('#c-current').textContent = part.stock ?? 'not counted';
    $('#c-stock').value = '';
    let made;
    const answer = ask(countDialog, () => (made = L.countPart(part, $('#c-stock').value)).error);
    $('#c-stock').focus(); // brings up the number keypad straight away
    if (await answer !== 'save') return;
    state.rows[at.r].columns[at.c].levels[at.l].parts[at.p] = made.part;
    commit();
  },
  pick({ r, c, l, p }) {
    const at = { r: +r, c: +c, l: +l, p: +p };
    if (picked && samePart(picked, at)) picked = null;                    // tapped again: un-pick
    else if (picked) { L.swapParts(state, picked, at); picked = null; return commit(); }
    else if (L.hasProduct(partAt(at))) picked = at;                        // only a product can be picked up
    render();
  },
});

// ---- Stillage close-up
const partDialog = $('#part-dialog');
const here = () => {
  const [r, c, l] = route().args;
  return { r, c, l, lvl: state.rows[r]?.columns[c]?.levels?.[l] };
};

Object.assign(views, {
  stillage(r, c, l) {
    const lvl = state.rows[r]?.columns[c]?.levels?.[l];
    if (!lvl) return views.warehouse(); // stale link, e.g. after an import
    const label = L.LEVELS[l] + L.columnNumbers(state.rows[r])[c];
    const items = L.allParts(state).filter(x => x.r === r && x.c === c && x.l === l && L.hasProduct(x.part));
    return { title: 'Warehouse', search: true, html: `
      <div class="closeup">
        ${boxHtml(lvl, label, query(), '', i => `role="button" data-action="edit-part" data-p="${i}"`, usageOn)}
        ${editing ? `<label class="partitions">Partition
          <select id="partitions">${[...L.LAYOUTS.keys()].map(id => `<option${id === lvl.layout ? ' selected' : ''}>${id}</option>`).join('')}</select>
        </label>
        <button class="danger" data-action="delete-col">Delete column</button>` : ''}
      </div>${listHtml(items, 'closeup-list')}` };
  },
});

Object.assign(actions, {
  async 'edit-part'({ p }) {
    const { lvl } = here();
    const part = lvl.parts[p];
    $('#p-style').innerHTML = styleOptions(part.style);
    $('#p-number').value = part.number;
    $('#p-color').value = part.color;
    $('#p-stock').value = part.stock ?? '';
    $('#p-full').value = part.full ?? '';
    const counted = L.lastCounted(part);
    $('#p-counted').textContent = counted ? `Last counted ${longDate(counted)}` : 'Not counted yet';
    let made;
    const check = () => (made = L.makePart({
      style: $('#p-style').value, number: $('#p-number').value, color: $('#p-color').value,
      stock: $('#p-stock').value, full: $('#p-full').value,
    }, part)).error;
    if (await ask(partDialog, check) !== 'save') return;
    lvl.parts[p] = made.part;
    commit();
  },
  'delete-col'() {
    const { r, c } = here();
    if (!confirm('Delete this whole column (levels a, b and c) and everything in it?')) return;
    state.rows[r].columns.splice(c, 1);
    save();
    history.back();
  },
});

document.addEventListener('change', e => {
  if (e.target.id !== 'partitions') return;
  const { lvl } = here();
  const id = e.target.value;
  if (L.dropsProducts(lvl, L.LAYOUTS.get(id).length) && !confirm('Fewer partitions will remove the products in the ones dropped. Continue?')) return render();
  L.setLayout(lvl, id);
  commit();
});

// ---- Excel import / export
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const exportDialog = $('#export-dialog');

Object.assign(actions, {
  async export() {
    const parts = () => ({ map: $('#x-map').checked, count: $('#x-count').checked });
    if (await ask(exportDialog, () => (parts().map || parts().count ? null : 'Tick at least one.')) !== 'save') return;
    const name = `warehouse-${L.ymd()}.xlsx`;
    const file = new File([await exportWorkbook(window.ExcelJS, state, parts())], name, { type: XLSX_TYPE });
    if (navigator.canShare?.({ files: [file] })) {
      try { return await navigator.share({ files: [file] }); }
      catch (e) { if (e.name === 'AbortError') return; } // user cancelled; any other share failure falls back to a download
    }
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(file), download: name });
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  },
  import() { $('#file').click(); },
});

$('#file').addEventListener('change', async e => {
  const file = e.target.files[0];
  e.target.value = ''; // lets the same file be picked again later
  if (!file) return;
  const res = await importWorkbook(window.ExcelJS, await file.arrayBuffer());
  if (res.error) return alert(res.error);
  const s = L.summary(res.state);
  const skipped = res.skipped.length ? `\n\nSkipped ${res.skipped.length} line(s):\n${res.skipped.slice(0, 10).join('\n')}` : '';
  if (!confirm(`Load ${s.rows} rows, ${s.stillages} stillages and ${s.products} products?${skipped}\n\nThis replaces everything on this device.`)) return;
  state = res.state;
  commit();
  alert('Import complete.');
});

// ---- Start
navigator.storage?.persist?.();
// Skipped on localhost so a cached copy never hides edits during development; test offline via 127.0.0.1.
if ('serviceWorker' in navigator && location.hostname !== 'localhost') navigator.serviceWorker.register('sw.js');
// Always open at Home: opened at a deeper page (a saved link or icon), Back would have no history to return to.
if (location.hash) history.replaceState(null, '', location.pathname + location.search);
render();

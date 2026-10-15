// UI: renders the view chosen by the URL hash into #view and routes taps to `actions`.
// Rules live in logic.js; Excel lives in excel.js.
import * as L from './logic.js';
import { exportWorkbook, importWorkbook } from './excel.js';

const KEY = 'inventory-app-v1';
const $ = sel => document.querySelector(sel);
const esc = s => String(s).replace(/[&<>"']/g, ch => `&#${ch.charCodeAt(0)};`);

let state = load();

function load() {
  try { return JSON.parse(localStorage.getItem(KEY)) ?? L.newState(); } catch { return L.newState(); }
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

function boxHtml(lvl, label, q, boxAttrs, partAttrs = () => '') {
  return `<div class="box" ${boxAttrs}>
    <span class="label">${label}</span>
    <div class="parts p${lvl.parts.length}">${lvl.parts.map((p, i) =>
      `<div class="part${L.matches(p, q) ? ' hit' : ''}" ${partAttrs(i)}>${esc(L.code(p) || 'NA')}</div>`).join('')}</div>
  </div>`;
}

function listHtml(items, scrollKey) {
  return `<div class="list" data-scroll="${scrollKey}"><table>
    <thead><tr><th>Code</th><th>Location</th><th>Box</th><th>Stock</th></tr></thead>
    <tbody>${items.map(x => `<tr><td>${esc(L.code(x.part))}</td><td>${esc(x.row)}</td><td>${x.box}</td><td>${x.part.stock ?? ''}</td></tr>`).join('')}</tbody>
  </table></div>`;
}

function rowHtml(row, r, q) {
  const nums = L.columnNumbers(row);
  const adders = side => `<div class="adders">
      <button data-action="add-col" data-r="${r}" data-type="stillage" data-side="${side}">Add Stillage</button>
      <button data-action="add-col" data-r="${r}" data-type="space" data-side="${side}">Add Space</button>
    </div>`;
  const columns = row.columns.map((col, c) => col.type === 'space'
    ? `<button class="walkway" data-action="remove-space" data-r="${r}" data-c="${c}">Walk<br>way</button>`
    : `<div class="col">${col.levels.map((lvl, l) =>
        boxHtml(lvl, L.LEVELS[l] + nums[c], q, `role="button" data-action="open" data-r="${r}" data-c="${c}" data-l="${l}"`)).join('')}</div>`).join('');
  return `<section class="row">
    <h2><button class="row-name" data-action="edit-row" data-r="${r}">${esc(row.name)}</button></h2>
    <div class="row-scroll" data-scroll="row-${r}">${adders('left')}${columns}${adders('right')}</div>
  </section>`;
}

Object.assign(views, {
  warehouse() {
    const q = query();
    const items = L.allParts(state).filter(x => (L.isSearching(q) ? L.matches(x.part, q) : L.hasProduct(x.part)));
    return { title: 'Warehouse', search: true, html: `
      <div class="map${state.rows.length ? '' : ' empty'}" data-scroll="map">
        ${state.rows.map((row, r) => rowHtml(row, r, q)).join('')}
        <button class="add-row" data-action="new-row" aria-label="Add row">+</button>
      </div>${listHtml(items, 'list')}` };
  },
});

const rowDialog = $('#row-dialog'), rowName = $('#row-name');

Object.assign(actions, {
  async 'new-row'() {
    rowDialog.querySelector('[value=delete]').hidden = true;
    rowName.value = '';
    if (await ask(rowDialog, () => L.rowNameError(state, rowName.value)) !== 'save') return;
    state.rows.push(L.newRow(rowName.value));
    commit();
  },
  async 'edit-row'({ r }) {
    const row = state.rows[r];
    rowDialog.querySelector('[value=delete]').hidden = false;
    rowName.value = row.name;
    const choice = await ask(rowDialog, () => L.rowNameError(state, rowName.value, row));
    if (choice === 'save') row.name = rowName.value.trim();
    else if (choice === 'delete' && confirm(`Delete "${row.name}" and everything in it?`)) state.rows.splice(Number(r), 1);
    else return;
    commit();
  },
  'add-col'({ r, type, side }) { L.addColumn(state.rows[r], type, side); commit(); },
  'remove-space'({ r, c }) {
    if (!confirm('Remove this walkway?')) return;
    state.rows[r].columns.splice(Number(c), 1);
    commit();
  },
  open({ r, c, l }) { location.hash = `#stillage/${r}/${c}/${l}`; },
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
        ${boxHtml(lvl, label, query(), '', i => `role="button" data-action="edit-part" data-p="${i}"`)}
        <label class="partitions">Partition
          <select id="partitions">${[1, 2, 3, 4].map(n => `<option${n === lvl.parts.length ? ' selected' : ''}>${n}</option>`).join('')}</select>
        </label>
        <button class="danger" data-action="delete-col">Delete column</button>
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
    let made;
    const check = () => (made = L.makePart({
      style: $('#p-style').value, number: $('#p-number').value, color: $('#p-color').value, stock: $('#p-stock').value,
    })).error;
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
  const n = Number(e.target.value);
  if (L.dropsProducts(lvl, n) && !confirm('Fewer partitions will remove the products in the ones dropped. Continue?')) return render();
  L.setPartitions(lvl, n);
  commit();
});

// ---- Excel import / export
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

Object.assign(actions, {
  async export() {
    const name = `warehouse-${L.ymd()}.xlsx`;
    const file = new File([await exportWorkbook(window.ExcelJS, state)], name, { type: XLSX_TYPE });
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
render();

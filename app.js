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
      <a class="big yard" href="#yard">Yard</a>
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
      <h2>Bulk types</h2>
      <ul>${state.yard.map(({ name }, i) => `
        <li>${esc(name)} <button data-action="remove-bulk" data-i="${i}" aria-label="Remove ${esc(name)}">✕</button></li>`).join('')
        || '<li class="muted">No bulk types yet</li>'}</ul>
      <form data-action="add-bulk">
        <input name="code" placeholder="e.g. Black Bulk" autocomplete="off">
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
  'add-bulk'(_, form) {
    const error = L.addBulkType(state, form.elements.code.value);
    if (error) return alert(error);
    commit();
  },
  'remove-bulk'({ i }) { // its cards go with it, so ask first if it has any
    const { name, rows } = state.yard[i], n = rows.flat().length;
    if (n && !confirm(`Remove ${name} and its ${L.productCount(n)} from the yard?`)) return;
    state.yard.splice(Number(i), 1);
    commit();
  },
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
// `graphs` adds the 30-day stock lines (the close-up, with usage on). A list partition shows "LIST" and its product
// count, with no usage colour or graph.
function boxHtml(lvl, label, q, boxAttrs, partAttrs = () => '', graphs = false) {
  const blocks = L.LAYOUTS.get(lvl.layout);
  return `<div class="box" ${boxAttrs}>
    <span class="label">${label}</span>
    <div class="parts${usageOn ? ' usage' : ''}" data-n="${lvl.parts.length}">${lvl.parts.map((p, i) => {
      const [x, y, w, h] = blocks[i];
      const area = `grid-area:${y + 1}/${x + 1}/span ${h}/span ${w}`;
      const body = L.isList(p) ? `LIST<small>${L.productCount(p.list.length)}</small>`
        : usageOn && L.hasProduct(p) ? `${esc(L.code(p))}<b>${esc(p.stock ?? '–')}${p.full ? ` / ${esc(p.full)}` : ''}</b>`
        : esc(L.code(p) || 'NA');
      return `<div class="part${L.isList(p) ? ' list-part' : ''}${L.matches(p, q) ? ' hit' : ''}" style="${area}${usageOn ? `;background:${usageColour(p)}` : ''}" ${partAttrs(i)}>${
        graphs ? graphHtml(p) : ''}<span class="txt">${body}</span></div>`;
    }).join('')}</div>
  </div>`;
}

function listHtml(items, scrollKey, rowAttrs = () => '') {
  return `<div class="list" data-scroll="${scrollKey}"><table>
    <thead><tr><th>Code</th><th>Location</th><th>Box</th><th>Stock</th><th>Counted</th></tr></thead>
    <tbody>${items.map(x => `<tr ${rowAttrs(x)}><td>${esc(L.code(x.part))}</td><td>${esc(x.row)}</td><td>${x.box}</td><td>${esc(x.part.stock ?? '')}</td><td>${shortDate(L.lastCounted(x.part))}</td></tr>`).join('')}</tbody>
  </table></div>`;
}

// Map modes. Move: tap a product to pick it, tap it again to pick its whole stillage, then tap where it goes.
// Count: a tap opens the quick count pop-up. Edit: shows the controls that change the layout; off at every start.
let moving = false, counting = false, editing = false, picked = null; // picked: { r, c, l, p, whole }
const sameBox = (a, b) => a.r === b.r && a.c === b.c && a.l === b.l;
const samePart = (a, b) => sameBox(a, b) && a.p === b.p;
const partAt = ({ r, c, l, p }) => state.rows[r].columns[c].levels[l].parts[p];

function rowHtml(row, r, q) {
  const nums = L.columnNumbers(row);
  // What a tap does: open the close-up, or (in Move / Count mode) act on the partition itself.
  const taps = (c, l) => {
    const where = `data-r="${r}" data-c="${c}" data-l="${l}"`;
    if (moving) return [picked?.whole && sameBox(picked, { r, c, l }) ? 'data-picked' : '',
      p => `role="button" data-action="pick" ${where} data-p="${p}"${picked && !picked.whole && samePart(picked, { r, c, l, p }) ? ' data-picked' : ''}`];
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
    const box = picked && `${esc(state.rows[picked.r].name)} ${L.LEVELS[picked.l]}${L.columnNumbers(state.rows[picked.r])[picked.c]}`;
    const hint = moving ? (picked?.whole ? `Now tap the stillage to swap ${box} with (tap ${box} again to cancel)`
      : picked ? `Now tap where ${esc(L.code(partAt(picked)))} should go, or tap it again to move the whole stillage`
      : 'Tap a product to move')
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
    if (L.isList(part)) return editPart(at); // a list is counted in its pop-up, one stock box per product
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
    if (picked?.whole) {
      if (sameBox(picked, at)) picked = null;                              // third tap on it: cancel
      else { L.swapLevels(state, picked, at); picked = null; return commit(); }
    } else if (picked) {
      if (samePart(picked, at)) picked = { ...picked, whole: true };       // second tap: the whole stillage
      else { L.swapParts(state, picked, at); picked = null; return commit(); }
    } else if (L.hasProduct(partAt(at))) picked = at;                      // only a product can be picked up
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

// ---- Product pop-up: one product, or with List ticked, a list of products with a stock box each.
// Nothing changes until Save, so Cancel undoes adds and removals.
let draft = null; // the list being edited: { items, typed }, typed[i] being what's in items[i]'s stock box
const nameFields = () => ({ style: $('#p-style').value, number: $('#p-number').value, color: $('#p-color').value });
const listOn = () => $('#p-is-list').checked;

function drawDraft() {
  $('#p-single').hidden = listOn();
  $('#p-list-section').hidden = !listOn();
  $('#p-list').innerHTML = draft.items.map((x, i) => `<div class="list-line">
      <span>${esc(L.code(x))}</span>
      <label>Stock <input class="list-stock" data-i="${i}" inputmode="numeric" autocomplete="off" placeholder="${esc(x.stock ?? '')}" value="${esc(draft.typed[i])}"></label>
      <button type="button" class="list-remove" data-action="list-remove" data-i="${i}" aria-label="Remove ${esc(L.code(x))}">✕</button>
    </div>`).join('') || '<p class="muted">No products yet</p>';
}

$('#p-list').addEventListener('input', e => { if (e.target.dataset.i) draft.typed[e.target.dataset.i] = e.target.value; });
$('#p-is-list').addEventListener('change', e => {
  if (!e.target.checked && draft.items.length && !confirm('Turn off List? Only the code in the name fields will be kept.')) e.target.checked = true;
  drawDraft();
});
// In a list, Enter in a name field adds that product rather than saving the pop-up.
partDialog.querySelector('.code-fields').addEventListener('keydown', e => {
  if (e.key === 'Enter' && listOn()) { e.preventDefault(); actions['list-add'](); }
});

async function editPart(at) {
  const part = partAt(at), wasList = L.isList(part);
  $('#p-style').innerHTML = styleOptions(part.style);
  $('#p-number').value = part.number;
  $('#p-color').value = part.color;
  $('#p-stock').value = part.stock ?? '';
  $('#p-full').value = part.full ?? '';
  const counted = L.lastCounted(part);
  $('#p-counted').textContent = counted ? `Last counted ${longDate(counted)}` : 'Not counted yet';
  $('#p-is-list').checked = wasList;
  // Ticking List on a partition with a product makes that product the list's first.
  const items = wasList ? part.list : L.hasProduct(part) ? [part] : [];
  draft = { items, typed: items.map(() => '') };
  drawDraft();
  let made;
  const check = () => (made = listOn()
    ? L.makeList(nameFields(), draft.items, draft.typed)
    : L.makePart({ ...nameFields(), stock: $('#p-stock').value, full: $('#p-full').value }, wasList ? L.emptyPart() : part)).error;
  if (await ask(partDialog, check) !== 'save') return;
  state.rows[at.r].columns[at.c].levels[at.l].parts[at.p] = made.part;
  commit();
}

Object.assign(actions, {
  'edit-part'({ p }) { const { r, c, l } = here(); return editPart({ r, c, l, p: +p }); },
  'list-add'() {
    const res = L.addToList(draft.items, nameFields());
    partDialog.querySelector('.error').textContent = res.error ?? '';
    if (res.error) return;
    draft.items = res.items;
    draft.typed.push('');
    drawDraft();
  },
  'list-remove'({ i }) {
    draft.items = draft.items.filter((_, k) => k !== +i);
    draft.typed.splice(+i, 1);
    drawDraft();
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

// ---- Yard: bulk stillages waiting to come into the warehouse. Each section is rows of product cards, and each
// card's count changes in quarter stillages with its six buttons.
const yardDialog = $('#yard-dialog');
const STEPS = [[1, '+1'], [-1, '−1'], [0.5, '+½'], [-0.5, '−½'], [0.25, '+¼'], [-0.25, '−¼']];

function yardCardHtml(s, r, c, card) {
  const at = `data-s="${s}" data-r="${r}" data-c="${c}"`;
  return `<div class="yard-card">
    <button class="yard-code" data-action="yard-edit" ${at}>${esc(L.code(card))} <span>(${card.count})</span></button>
    <div class="yard-steps">${STEPS.map(([d, label]) =>
      `<button class="${d > 0 ? 'up' : 'down'}" data-action="yard-step" ${at} data-d="${d}" aria-label="${label} ${esc(L.code(card))}">${label}</button>`).join('')}</div>
  </div>`;
}

Object.assign(views, {
  yard: () => ({ title: 'Yard', html: `<div class="yard-page" data-scroll="yard">${state.yard.map(({ name, rows }, s) => `
    <section class="yard-section">
      <h2>${esc(name)}</h2>
      ${rows.map((row, r) => `<div class="yard-row" data-scroll="yard-${s}-${r}">
        ${row.map((card, c) => yardCardHtml(s, r, c, card)).join('')}
        <button class="yard-add" data-action="yard-add-card" data-s="${s}" data-r="${r}" aria-label="Add a product to this row">+</button>
      </div>`).join('')}
      <button class="yard-add" data-action="yard-add-row" data-s="${s}" aria-label="Add a row to ${esc(name)}">+</button>
    </section>`).join('') || '<p class="muted">Add a bulk type in Settings to start.</p>'}</div>` }),
});

// The yard card pop-up: resolves with the new card's code fields, 'delete', or null when cancelled.
async function askYardCard(card, deletable) {
  $('#y-style').innerHTML = styleOptions(card.style);
  $('#y-number').value = card.number;
  $('#y-color').value = card.color;
  yardDialog.querySelector('[value=delete]').hidden = !deletable;
  let made;
  const choice = await ask(yardDialog, () =>
    (made = L.makeYardCard({ style: $('#y-style').value, number: $('#y-number').value, color: $('#y-color').value })).error);
  return choice === 'save' ? made.card : choice === 'delete' ? 'delete' : null;
}

Object.assign(actions, {
  'yard-step'({ s, r, c, d }) {
    const card = state.yard[s].rows[r][c];
    card.count = L.stepCount(card.count, Number(d));
    commit();
  },
  async 'yard-add-card'({ s, r }) {
    const card = await askYardCard({ style: '', number: '', color: '' }, false);
    if (!card || card === 'delete') return;
    state.yard[s].rows[r].push(card);
    commit();
  },
  async 'yard-add-row'({ s }) { // a new row starts with its first product, so no empty rows
    const card = await askYardCard({ style: '', number: '', color: '' }, false);
    if (!card || card === 'delete') return;
    state.yard[s].rows.push([card]);
    commit();
  },
  async 'yard-edit'({ s, r, c }) {
    const rows = state.yard[s].rows, row = rows[r], card = row[c];
    const result = await askYardCard(card, true);
    if (result === 'delete') {
      if (!confirm(`Delete ${L.code(card)} from the yard?`)) return;
      row.splice(Number(c), 1);
      if (!row.length) rows.splice(Number(r), 1); // a row left empty goes too
    } else if (result) Object.assign(card, { style: result.style, number: result.number, color: result.color });
    else return;
    commit();
  },
});

// ---- Excel import / export
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const exportDialog = $('#export-dialog');

Object.assign(actions, {
  async export() {
    const parts = () => ({ map: $('#x-map').checked, count: $('#x-count').checked, yard: $('#x-yard').checked });
    if (await ask(exportDialog, () => (Object.values(parts()).some(Boolean) ? null : 'Tick at least one.')) !== 'save') return;
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

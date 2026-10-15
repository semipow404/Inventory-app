// UI: renders the view chosen by the URL hash into #view and routes taps to `actions`.
// Rules live in logic.js; Excel lives in excel.js.
import * as L from './logic.js';

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

// ---- Start
navigator.storage?.persist?.();
render();

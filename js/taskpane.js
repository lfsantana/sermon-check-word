// Panel de Sermon Check dentro de Word.
import { DEFAULTS, sanitize, sections, formatDuration } from './timing.js';
import * as word from './word.js';

const $ = (id) => document.getElementById(id);
const SETTINGS_KEY = 'sc.settings';
const REHEARSAL_KEY = 'sc.rehearsal';

const store = {
  get(key, fallback) {
    try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; }
  },
  set(key, value) { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* sin almacenamiento */ } },
};

const state = {
  api: null,
  settings: { ...DEFAULTS, ...store.get(SETTINGS_KEY, {}) },
  paragraphs: [],
  view: null,
  busy: false,
};

// ---------- avisos ----------
let noticeTimer = null;
function notify(text, { sticky = false } = {}) {
  $('noticeText').textContent = text;
  $('notice').hidden = false;
  clearTimeout(noticeTimer);
  if (!sticky) noticeTimer = setTimeout(() => { $('notice').hidden = true; }, 6000);
}
$('noticeClose').addEventListener('click', () => { $('notice').hidden = true; });

const plural = (n, one, many) => `${n} ${n === 1 ? one : many}`;

function explain(err) {
  const code = err?.code || '';
  if (code === 'AccessDenied' || /protected|read-?only/i.test(err?.message || '')) {
    return 'Word no permite cambiar este documento (está protegido o en modo de solo lectura).';
  }
  return `Word no pudo completar la acción. ${err?.message || ''}`.trim();
}

// ---------- vista ----------
const el = (tag, cls, text) => {
  const node = document.createElement(tag);
  if (cls) node.className = cls;
  if (text != null) node.textContent = text;
  return node;
};

function compute() {
  state.view = word.analyze(state.paragraphs, state.settings);
  return state.view;
}

function jump(i) {
  state.api.goToBlock(i, state.settings).catch(() => {});
}

function render() {
  const { tl, planned, existing } = compute();
  $('statTotal').textContent = formatDuration(tl.total);
  $('statWords').textContent = tl.words.toLocaleString('es');
  $('statInDoc').textContent = String(existing.length);
  $('statInDocLabel').textContent = existing.length === 1 ? 'etiqueta en el documento' : 'etiquetas en el documento';

  const empty = tl.words === 0;
  $('btnPlace').disabled = state.busy || empty || !planned.length;
  $('btnPlace').textContent = existing.length ? 'Volver a colocar etiquetas' : 'Colocar etiquetas';
  $('btnCursor').disabled = state.busy || empty;
  $('btnRehearse').disabled = state.busy || empty;
  $('btnRecalc').disabled = state.busy || !existing.length;
  $('btnRemove').disabled = state.busy || !existing.length;
  $('planHint').textContent = empty
    ? 'El documento no tiene texto todavía.'
    : planned.length
      ? `Con estos ajustes se colocarán ${plural(planned.length, 'etiqueta', 'etiquetas')}${existing.length ? ', en lugar de las que hay ahora' : ''}.`
      : 'Activa alguna regla de «Etiquetas de tiempo» para colocarlas automáticamente.';

  const secs = $('sections');
  secs.replaceChildren();
  for (const r of sections(tl, $('showSub').checked ? 2 : 1)) {
    const li = el('li', r.level === 2 ? 'sub' : '');
    li.title = `${r.words.toLocaleString('es')} palabras · empieza en ${formatDuration(r.start)}`;
    li.append(el('span', 'li-start', formatDuration(r.start)), el('span', 'li-title', r.title), el('span', 'li-end', formatDuration(r.end - r.start)));
    li.addEventListener('click', () => jump(r.i));
    secs.append(li);
  }

  const list = $('planned');
  list.replaceChildren();
  $('plannedCount').textContent = planned.length ? `(${planned.length})` : '';
  if (!planned.length) list.append(el('li', 'empty', 'Ninguna con los ajustes actuales.'));
  for (const m of planned) {
    const block = tl.blocks[m.after];
    const words = block.text.split(' ');
    const li = el('li');
    li.title = 'Ir a ese párrafo';
    li.append(el('span', 'li-start', formatDuration(m.seconds)), el('span', 'li-title', `…${words.slice(-7).join(' ')}`), el('span', 'li-label', m.text));
    li.addEventListener('click', () => jump(m.after));
    list.append(li);
  }
}

// ---------- documento ----------
let reading = null;
async function refresh({ quiet = false } = {}) {
  if (reading) return reading;
  reading = (async () => {
    try {
      state.paragraphs = await state.api.readParagraphs();
      render();
    } catch (err) {
      if (!quiet) notify(explain(err), { sticky: true });
    } finally {
      reading = null;
    }
  })();
  return reading;
}

async function act(fn, done) {
  if (state.busy) return;
  state.busy = true;
  render();
  try {
    if (reading) await reading;
    const result = await fn();
    const msg = done(result);
    if (msg) notify(msg);
  } catch (err) {
    notify(explain(err), { sticky: true });
  } finally {
    state.busy = false;
    await refresh({ quiet: true });
    render();
  }
}

$('btnRefresh').addEventListener('click', () => refresh());
$('btnPlace').addEventListener('click', () => act(
  () => state.api.placeLabels(state.settings),
  (r) => `${plural(r.placed, 'etiqueta colocada', 'etiquetas colocadas')}${r.removed ? ` (se quitaron ${r.removed} anteriores)` : ''}. Puedes deshacerlo con Ctrl+Z.`,
));
$('btnRemove').addEventListener('click', () => act(
  () => state.api.removeLabels(state.settings),
  (r) => `Se ${r.removed === 1 ? 'quitó' : 'quitaron'} ${plural(r.removed, 'etiqueta', 'etiquetas')}.`,
));
$('btnRecalc').addEventListener('click', () => act(
  () => state.api.refreshLabels(state.settings),
  (r) => (r.changed
    ? `Se ${r.changed === 1 ? 'corrigió' : 'corrigieron'} ${plural(r.changed, 'etiqueta', 'etiquetas')} de ${r.total}.`
    : 'Las etiquetas ya tenían el tiempo correcto.'),
));
$('btnCursor').addEventListener('click', () => act(
  () => state.api.labelAtCursor(state.settings),
  (r) => (r.text
    ? `Etiqueta ${r.text} añadida después del párrafo del cursor.`
    : r.already
      ? `Ese párrafo ya tiene la etiqueta ${r.already}.`
      : 'Coloca el cursor dentro de un párrafo con texto del documento.'),
));

// ---------- ajustes ----------
function syncInputs() {
  const s = state.settings;
  document.querySelectorAll('[data-setting]').forEach((input) => {
    const v = s[input.dataset.setting];
    if (input.type === 'checkbox') input.checked = !!v;
    else if (document.activeElement !== input) input.value = Array.isArray(v) ? v.join(', ') : v;
  });
  const clean = sanitize(s);
  document.querySelectorAll('#percents input').forEach((input) => { input.checked = clean.percents.includes(Number(input.value)); });
  document.querySelectorAll('[data-wpm]').forEach((b) => b.setAttribute('aria-pressed', String(Number(s.wpm) === Number(b.dataset.wpm))));
  document.querySelectorAll('[data-format]').forEach((b) => b.setAttribute('aria-pressed', String(clean.format === b.dataset.format)));
  document.querySelectorAll('.rule').forEach((rule) => {
    const box = rule.querySelector('input[type="checkbox"][data-setting]');
    rule.classList.toggle('offed', !!box && !box.checked);
  });
}

function setSetting(key, value) {
  state.settings = { ...state.settings, [key]: value };
  store.set(SETTINGS_KEY, state.settings);
  syncInputs();
  render();
}

document.querySelectorAll('[data-setting]').forEach((input) => {
  const onChange = () => {
    const key = input.dataset.setting;
    if (input.type === 'checkbox') { setSetting(key, input.checked); return; }
    // Escribir un valor en una regla la activa
    const box = input.closest('.rule')?.querySelector('input[type="checkbox"][data-setting]');
    if (box && !box.checked && input.value.trim()) state.settings[box.dataset.setting] = true;
    setSetting(key, input.value);
  };
  input.addEventListener(input.type === 'checkbox' || input.tagName === 'SELECT' ? 'change' : 'input', onChange);
  if (input.type !== 'checkbox') {
    input.addEventListener('blur', () => {
      const clean = sanitize(state.settings)[input.dataset.setting];
      input.value = Array.isArray(clean) ? clean.join(', ') : clean;
    });
  }
});
$('percents').addEventListener('change', () => {
  const values = [...document.querySelectorAll('#percents input:checked')].map((input) => Number(input.value));
  if (values.length) state.settings.usePercents = true;
  setSetting('percents', values);
});
document.querySelectorAll('[data-wpm]').forEach((b) => b.addEventListener('click', () => setSetting('wpm', Number(b.dataset.wpm))));
document.querySelectorAll('[data-format]').forEach((b) => b.addEventListener('click', () => setSetting('format', b.dataset.format)));
$('showSub').addEventListener('change', render);

// ---------- ensayo ----------
$('btnRehearse').addEventListener('click', async () => {
  await refresh({ quiet: true });
  const { tl, planned, settings } = compute();
  const inDoc = word.existingAsMarkers(tl);
  const payload = JSON.stringify({ at: Date.now(), tl, settings, markers: inDoc.length ? inDoc : planned });
  try { localStorage.setItem(REHEARSAL_KEY, payload); } catch { /* se enviará por mensaje */ }
  try {
    await state.api.openRehearsal(payload, (wpm) => {
      setSetting('wpm', wpm);
      notify(`Ritmo base cambiado a ${wpm} palabras por minuto.`);
    });
  } catch (err) {
    notify(err?.message || 'No se pudo abrir la ventana de ensayo.', { sticky: true });
  }
});

// ---------- arranque ----------
export function start(api) {
  state.api = api;
  $('outside').hidden = true;
  $('app').hidden = false;
  syncInputs();
  render();
  let timer = null;
  api.onDocumentActivity?.(() => {
    clearTimeout(timer);
    timer = setTimeout(() => { if (!state.busy) refresh({ quiet: true }); }, 1200);
  });
  window.addEventListener('focus', () => { if (!state.busy) refresh({ quiet: true }); });
  return refresh();
}

/* global Office */
function officeApi() {
  return {
    ...word,
    onDocumentActivity(handler) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
    },
    openRehearsal(payload, onWpm) {
      return new Promise((resolve, reject) => {
        const url = new URL('ensayo.html', window.location.href).href;
        Office.context.ui.displayDialogAsync(url, { height: 88, width: 88, displayInIframe: false }, (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error('No se pudo abrir la ventana de ensayo. Si ya hay una abierta, ciérrala primero.'));
            return;
          }
          const dialog = res.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            let msg = {};
            try { msg = JSON.parse(arg.message); } catch { /* mensaje desconocido */ }
            if (msg.type === 'ready') {
              try { dialog.messageChild(payload); } catch { /* la ventana ya lo leyó del almacenamiento */ }
            } else if (msg.type === 'wpm') {
              dialog.close();
              onWpm(Number(msg.wpm));
            } else if (msg.type === 'close') {
              dialog.close();
            }
          });
          resolve();
        });
      });
    },
  };
}

function outside() {
  $('app').hidden = true;
  $('outside').hidden = false;
}

if (window.Office?.onReady) {
  Office.onReady((info) => {
    if (info?.host === Office.HostType.Word) start(officeApi());
    else if (!state.api) outside();
  });
} else {
  outside();
}

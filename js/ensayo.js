// Ventana de ensayo: recibe el plan desde el panel y devuelve el ritmo medido.
import { startRehearsal } from './rehearsal.js';

const REHEARSAL_KEY = 'sc.rehearsal';
const MAX_AGE = 2 * 60 * 1000; // el plan guardado solo vale recién abierto

let started = false;

export function begin(payload, send) {
  if (started) return false;
  let data;
  try { data = typeof payload === 'string' ? JSON.parse(payload) : payload; } catch { return false; }
  if (!data?.tl?.blocks?.length) return false;
  started = true;
  document.getElementById('waiting').hidden = true;
  let answered = false;
  startRehearsal({
    root: document.getElementById('rehearsal'),
    tl: data.tl,
    markers: data.markers || [],
    settings: data.settings,
    onFinish: (wpm) => { answered = true; send({ type: 'wpm', wpm }); },
    onClose: () => setTimeout(() => { if (!answered) send({ type: 'close' }); }, 0),
  });
  return true;
}

function fromStorage() {
  try {
    const raw = localStorage.getItem(REHEARSAL_KEY);
    if (!raw) return null;
    localStorage.removeItem(REHEARSAL_KEY);
    const data = JSON.parse(raw);
    return Date.now() - Number(data.at) < MAX_AGE ? data : null;
  } catch { return null; }
}

/* global Office */
if (window.Office?.onReady) {
  Office.onReady(() => {
    const ui = Office.context?.ui;
    if (!ui?.messageParent) {
      document.getElementById('waiting').textContent = 'Esta ventana se abre desde el panel de Sermon Check en Word.';
      return;
    }
    const send = (msg) => ui.messageParent(JSON.stringify(msg));
    const stored = fromStorage();
    if (stored) begin(stored, send);
    try {
      ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => begin(arg.message, send));
    } catch { /* versiones sin mensajes del panel: se usa el plan guardado */ }
    send({ type: 'ready' });
    setTimeout(() => {
      if (!started) document.getElementById('waiting').textContent = 'No llegó el plan del ensayo. Cierra esta ventana y vuelve a pulsar «Modo ensayo».';
    }, 5000);
  });
}

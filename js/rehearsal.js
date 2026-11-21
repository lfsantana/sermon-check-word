// Modo ensayo: cronómetro con el texto y el desfase respecto al plan.
import { formatDuration } from './timing.js';

// Segundos de desfase respecto al párrafo actual: positivo = atrasado (ya debería
// haberlo terminado), negativo = adelantado (aún no tocaba empezarlo), 0 = dentro.
export function drift(elapsedSeconds, block) {
  if (elapsedSeconds > block.end) return elapsedSeconds - block.end;
  if (elapsedSeconds < block.start) return elapsedSeconds - block.start;
  return 0;
}

// Ritmo base que habría hecho coincidir el plan con el tiempo real.
export function measuredWpm(predictedSeconds, actualSeconds, settings) {
  if (!(predictedSeconds > 0) || !(actualSeconds > 0)) return settings.wpm;
  return Math.min(300, Math.max(40, Math.round((settings.wpm * predictedSeconds) / actualSeconds)));
}

const TOLERANCE = 15; // segundos que cuentan como «a tiempo»

export function driftText(seconds) {
  if (Math.abs(seconds) <= TOLERANCE) return { text: 'A tiempo', state: 'ok' };
  return seconds > 0
    ? { text: `+${formatDuration(seconds)} atrasado`, state: 'late' }
    : { text: `−${formatDuration(-seconds)} adelantado`, state: 'early' };
}

export function startRehearsal({ root, tl, markers, settings, onFinish, onClose }) {
  const doc = root.ownerDocument;
  const content = tl.blocks.filter((b) => b.kind !== 'marker' && b.kind !== 'empty');
  if (!content.length) return;
  const markerAfter = new Map(markers.map((m) => [m.after, m]));

  let elapsed = 0;      // segundos acumulados
  let startedAt = null; // performance.now() si está corriendo
  let current = 0;      // índice dentro de content
  let timer = null;

  root.innerHTML = `
    <div class="rh-bar">
      <div class="rh-clock" data-rh="clock">0:00</div>
      <div class="rh-drift" data-rh="drift" data-state="ok">Listo</div>
      <div class="rh-plan" data-rh="plan"></div>
      <div class="rh-actions">
        <button type="button" class="btn primary" data-rh="toggle">Iniciar</button>
        <button type="button" class="btn" data-rh="reset">Reiniciar</button>
        <button type="button" class="btn" data-rh="end">Terminar</button>
      </div>
    </div>
    <div class="rh-text" data-rh="text" tabindex="0"></div>
    <div class="rh-help">Toca el párrafo por donde vas, o usa ↓ / espacio para avanzar y ↑ para volver. Esc para terminar.</div>
    <div class="rh-result" data-rh="result" hidden></div>`;
  const $ = (name) => root.querySelector(`[data-rh="${name}"]`);
  const textEl = $('text');

  content.forEach((b, k) => {
    const p = doc.createElement(b.kind === 'heading' ? 'h3' : 'p');
    p.className = `rh-block k-${b.kind}`;
    p.dataset.k = String(k);
    p.textContent = b.text;
    textEl.appendChild(p);
    const m = markerAfter.get(b.i);
    if (m) {
      const tag = doc.createElement('div');
      tag.className = 'rh-marker';
      tag.textContent = m.text;
      textEl.appendChild(tag);
    }
  });

  const now = () => elapsed + (startedAt === null ? 0 : (performance.now() - startedAt) / 1000);

  const render = () => {
    const t = now();
    $('clock').textContent = formatDuration(t);
    const b = content[current];
    const started = t > 0;
    const d = driftText(started ? drift(t, b) : 0);
    $('drift').textContent = started ? d.text : 'Listo';
    $('drift').dataset.state = started ? d.state : 'ok';
    $('plan').textContent = `Previsto aquí: ${formatDuration(b.start)}–${formatDuration(b.end)} · Total: ${formatDuration(tl.total)}`;
    $('toggle').textContent = startedAt === null ? (t > 0 ? 'Continuar' : 'Iniciar') : 'Pausar';
  };

  const select = (k, scroll = true) => {
    current = Math.min(content.length - 1, Math.max(0, k));
    textEl.querySelectorAll('.rh-block').forEach((el) => {
      const idx = Number(el.dataset.k);
      el.classList.toggle('current', idx === current);
      el.classList.toggle('done', idx < current);
    });
    if (scroll) textEl.querySelector('.rh-block.current')?.scrollIntoView({ block: 'center', behavior: 'smooth' });
    render();
  };

  const pause = () => {
    if (startedAt !== null) { elapsed = now(); startedAt = null; }
    clearInterval(timer); timer = null;
  };
  const toggle = () => {
    if (startedAt === null) {
      startedAt = performance.now();
      timer = setInterval(render, 250);
    } else pause();
    render();
  };
  const reset = () => { pause(); elapsed = 0; $('result').hidden = true; select(0); };

  const close = () => {
    pause();
    doc.removeEventListener('keydown', onKey);
    root.hidden = true;
    root.innerHTML = '';
    onClose?.();
  };

  const finish = () => {
    pause();
    const t = now();
    const b = content[current];
    const atEnd = current === content.length - 1;
    const predicted = atEnd ? tl.total : b.end;
    if (t < 20 || predicted < 20) { close(); return; }
    const wpm = measuredWpm(predicted, t, settings);
    const d = driftText(t - predicted);
    const result = $('result');
    result.hidden = false;
    result.innerHTML = `
      <div class="rh-card">
        <h2>Resumen del ensayo</h2>
        <dl>
          <div><dt>Tiempo real</dt><dd>${formatDuration(t)}</dd></div>
          <div><dt>Previsto hasta ${atEnd ? 'el final' : 'este punto'}</dt><dd>${formatDuration(predicted)}</dd></div>
          <div><dt>Diferencia</dt><dd>${d.text}</dd></div>
          <div><dt>Tu ritmo medido</dt><dd>${wpm} palabras/min <span class="muted">(ahora ${settings.wpm})</span></dd></div>
        </dl>
        <div class="rh-card-actions">
          <button type="button" class="btn primary" data-rh="use">Usar ${wpm} como ritmo base</button>
          <button type="button" class="btn" data-rh="back">Seguir ensayando</button>
          <button type="button" class="btn" data-rh="close">Cerrar</button>
        </div>
      </div>`;
    $('use').addEventListener('click', () => { close(); onFinish?.(wpm); });
    $('back').addEventListener('click', () => { result.hidden = true; });
    $('close').addEventListener('click', close);
  };

  function onKey(e) {
    if (e.key === 'Escape') { e.preventDefault(); if ($('result').hidden) finish(); else close(); return; }
    if (!$('result').hidden) return;
    const onButton = e.target instanceof HTMLButtonElement;
    if (e.key === 'ArrowDown' || e.key === ' ' || e.key === 'PageDown') { e.preventDefault(); select(current + 1); }
    else if (e.key === 'ArrowUp' || e.key === 'PageUp') { e.preventDefault(); select(current - 1); }
    else if (e.key === 'Enter' && !onButton) { e.preventDefault(); toggle(); }
  }

  textEl.addEventListener('click', (e) => {
    const el = e.target.closest('.rh-block');
    if (el) select(Number(el.dataset.k), false);
  });
  $('toggle').addEventListener('click', toggle);
  $('reset').addEventListener('click', reset);
  $('end').addEventListener('click', finish);
  // El espacio no debe «pulsar» el botón enfocado al soltar la tecla
  root.addEventListener('keyup', (e) => { if (e.key === ' ') e.preventDefault(); });
  doc.addEventListener('keydown', onKey);

  root.hidden = false;
  select(0);
  $('toggle').focus();
}

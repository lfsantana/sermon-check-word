// Cálculo de tiempos y colocación de etiquetas. Funciones puras, sin DOM.

export const DEFAULTS = Object.freeze({
  wpm: 120,          // palabras por minuto
  quoteWpm: 110,     // ritmo al leer citas
  headingPause: 3,   // segundos de pausa por título
  margin: 0,         // % extra para improvisación
  useEveryMin: true, everyMin: 5,
  useEveryWords: false, everyWords: 500,
  useEndOfLevel: false, endOfLevel: 1,
  usePercents: false, percents: [50],
  useAtMinutes: false, atMinutes: [],
  format: 'min',     // 'min' → [10 MIN], 'clock' → [10:30]
});

const num = (v, fallback, { zeroOk = false } = {}) => {
  const n = typeof v === 'string' ? parseFloat(v.replace(',', '.')) : v;
  if (typeof n !== 'number' || !Number.isFinite(n)) return fallback;
  if (n < 0 || (n === 0 && !zeroOk)) return fallback;
  return n;
};

const numList = (v) => {
  const items = Array.isArray(v) ? v : String(v ?? '').split(/[\s,;]+/);
  const out = [];
  for (const item of items) {
    const n = num(item, null);
    if (n !== null && !out.includes(n)) out.push(n);
  }
  return out.sort((a, b) => a - b);
};

export function sanitize(input = {}) {
  const s = { ...DEFAULTS, ...input };
  return {
    wpm: num(s.wpm, DEFAULTS.wpm),
    quoteWpm: num(s.quoteWpm, DEFAULTS.quoteWpm),
    headingPause: num(s.headingPause, DEFAULTS.headingPause, { zeroOk: true }),
    margin: num(s.margin, DEFAULTS.margin, { zeroOk: true }),
    useEveryMin: !!s.useEveryMin,
    everyMin: num(s.everyMin, DEFAULTS.everyMin),
    useEveryWords: !!s.useEveryWords,
    everyWords: Math.round(num(s.everyWords, DEFAULTS.everyWords)) || DEFAULTS.everyWords,
    useEndOfLevel: !!s.useEndOfLevel,
    endOfLevel: [1, 2, 3].includes(Number(s.endOfLevel)) ? Number(s.endOfLevel) : 1,
    usePercents: !!s.usePercents,
    percents: numList(s.percents).filter((p) => p > 0 && p < 100),
    useAtMinutes: !!s.useAtMinutes,
    atMinutes: numList(s.atMinutes),
    format: s.format === 'clock' ? 'clock' : 'min',
  };
}

export const countWords = (text) =>
  String(text || '').split(/\s+/).filter((w) => /[\p{L}\p{N}]/u.test(w)).length;

const REF = /\d+\s?[:.]\s?\d+/;

// Añade `kind`, `words`, `group` (id de cita) a los bloques crudos.
export function classify(rawBlocks, overrides = {}) {
  const blocks = rawBlocks.map((b, idx) => ({ ...b, i: idx, words: countWords(b.text), group: null }));
  let ref = null;       // índice de la referencia abierta
  let quoteOpen = false; // ya hay al menos un párrafo de esa cita
  let prev = null;       // bloque con contenido anterior
  for (const b of blocks) {
    let kind;
    if (b.marker) kind = 'marker';
    else if (b.words === 0) kind = 'empty';
    else if (b.level > 0) kind = 'heading';
    else if (b.bold && b.words <= 12 && REF.test(b.text)) kind = 'ref';
    else if (ref !== null && (!quoteOpen || b.verse)) kind = 'quote';
    else if (b.indent) kind = 'quote';
    else kind = 'normal';

    const forced = overrides[b.i];
    if (forced && kind !== 'marker' && kind !== 'empty') kind = forced;
    if (kind === 'heading' && !b.level) b.level = 2;

    if (kind === 'marker' || kind === 'empty') { b.kind = kind; continue; }
    if (kind === 'ref') { ref = b.i; quoteOpen = false; }
    else if (kind === 'quote') {
      if (forced) b.group = prev?.kind === 'quote' ? prev.group : `q${b.i}`;
      else if (ref !== null) { b.group = `r${ref}`; quoteOpen = true; }
      else b.group = prev?.kind === 'quote' && prev.group.startsWith('q') ? prev.group : `q${b.i}`;
      if (!b.group.startsWith('r')) { ref = null; quoteOpen = false; }
    } else { ref = null; quoteOpen = false; }
    b.kind = kind;
    prev = b;
  }
  return blocks;
}

export function timeline(blocks, settingsIn) {
  const s = sanitize(settingsIn);
  const factor = 1 + s.margin / 100;
  let t = 0;
  let w = 0;
  const out = blocks.map((b) => {
    let dur = 0;
    if (b.kind !== 'marker' && b.kind !== 'empty') {
      dur = (b.words / (b.kind === 'quote' ? s.quoteWpm : s.wpm)) * 60;
      if (b.kind === 'heading') dur += s.headingPause;
      dur *= factor;
      w += b.words;
    }
    const block = { ...b, start: t, end: t + dur, wordsEnd: w };
    t += dur;
    return block;
  });
  return { blocks: out, total: t, words: w };
}

const isContent = (b) => b.kind !== 'marker' && b.kind !== 'empty';

export function sections(tl, level = 1) {
  const content = tl.blocks.filter(isContent);
  const heads = content.filter((b) => b.kind === 'heading' && b.level <= level);
  if (!heads.length) {
    return [{ title: 'Todo el discurso', start: 0, end: tl.total, words: tl.words, i: content[0]?.i ?? 0 }];
  }
  const out = [];
  const first = heads[0];
  if (first.start > 0 && first.wordsEnd - first.words > 0) {
    out.push({ title: 'Inicio', start: 0, end: first.start, words: first.wordsEnd - first.words, i: content[0].i });
  }
  heads.forEach((h, k) => {
    const next = heads[k + 1];
    const end = next ? next.start : tl.total;
    const wordsEnd = next ? next.wordsEnd - next.words : tl.words;
    out.push({ title: h.text, level: h.level, start: h.start, end, words: wordsEnd - (h.wordsEnd - h.words), i: h.i });
  });
  return out;
}

// Límites donde puede ir una etiqueta: después de un bloque con contenido.
export function boundaries(tl) {
  const content = tl.blocks.filter(isContent);
  const out = [];
  for (let k = 0; k < content.length - 1; k++) {
    const b = content[k];
    const next = content[k + 1];
    const insideQuote = b.kind === 'quote' && next.kind === 'quote' && b.group !== null && b.group === next.group;
    const valid = b.kind !== 'heading' && b.kind !== 'ref' && !insideQuote;
    out.push({ after: b.i, end: b.end, wordsEnd: b.wordsEnd, valid, nextI: next.i });
  }
  return out;
}

const pad = (n) => String(n).padStart(2, '0');

export function formatDuration(seconds) {
  const s = Math.max(0, Math.round(seconds));
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  return h ? `${h}:${pad(m)}:${pad(s % 60)}` : `${m}:${pad(s % 60)}`;
}

export function formatLabel(seconds, format) {
  const s = Math.max(0, Math.round(seconds));
  if (format !== 'clock' && s % 60 === 0) return `[${s / 60} MIN]`;
  return `[${Math.floor(s / 60)}:${pad(s % 60)}]`;
}

// Etiqueta con el tiempo real de un punto (redondeado al minuto en formato «min»).
export const realLabel = (seconds, format) =>
  format === 'clock'
    ? formatLabel(seconds, 'clock')
    : formatLabel(Math.max(1, Math.round(seconds / 60)) * 60, 'min');

const PRIORITY = { at: 0, every: 1, percent: 2, section: 3, words: 4, manual: 5 };

export function placeMarkers(tl, settingsIn, manual = []) {
  const s = sanitize(settingsIn);
  const all = boundaries(tl);
  const valid = all.filter((b) => b.valid);
  const chosen = new Map(); // after → candidato

  const offer = (b, cand) => {
    const prev = chosen.get(b.after);
    const better = !prev
      || PRIORITY[cand.source] < PRIORITY[prev.source]
      || (cand.source === prev.source && cand.dist < prev.dist);
    if (better) chosen.set(b.after, { after: b.after, seconds: b.end, ...cand });
  };
  const nearest = (value, key) => {
    let best = null;
    let bestDist = Infinity;
    for (const b of valid) {
      const d = Math.abs(b[key] - value);
      if (d < bestDist) { best = b; bestDist = d; }
    }
    return best && { b: best, dist: bestDist };
  };
  const timeTarget = (T, source, shown = T) => {
    if (!(T > 0) || T >= tl.total) return;
    const hit = nearest(T, 'end');
    if (hit) offer(hit.b, { source, dist: hit.dist, text: formatLabel(shown, s.format) });
  };

  if (valid.length) {
    if (s.useAtMinutes) for (const m of s.atMinutes) timeTarget(m * 60, 'at');
    if (s.useEveryMin) {
      const step = s.everyMin * 60;
      for (let T = step, n = 0; T < tl.total && n < 2000; T += step, n++) timeTarget(T, 'every');
    }
    if (s.usePercents) {
      for (const p of s.percents) {
        const T = (tl.total * p) / 100;
        timeTarget(T, 'percent', s.format === 'clock' ? T : Math.max(1, Math.round(T / 60)) * 60);
      }
    }
    if (s.useEndOfLevel) {
      const content = tl.blocks.filter(isContent);
      content.forEach((h, k) => {
        if (h.kind !== 'heading' || h.level > s.endOfLevel || k === 0) return;
        // último límite válido antes de este título
        let hit = null;
        for (const b of valid) if (b.end <= h.start + 1e-9) hit = b;
        // un «punto» de menos de medio minuto (p. ej. el título del sermón) no merece etiqueta
        if (hit && hit.end >= 30) offer(hit, { source: 'section', dist: 0, text: realLabel(hit.end, s.format) });
      });
    }
    if (s.useEveryWords) {
      for (let W = s.everyWords, n = 0; W < tl.words && n < 2000; W += s.everyWords, n++) {
        const hit = nearest(W, 'wordsEnd');
        if (hit) offer(hit.b, { source: 'words', dist: hit.dist, text: realLabel(hit.b.end, s.format) });
      }
    }
  }

  for (const m of manual) {
    if (m.off) { chosen.delete(m.after); continue; }
    const b = all.find((x) => x.after === m.after);
    if (b && !chosen.has(b.after)) {
      chosen.set(b.after, { after: b.after, seconds: b.end, source: 'manual', dist: 0, text: realLabel(b.end, s.format) });
    }
  }
  // «off» gana siempre, aunque esté antes en la lista
  for (const m of manual) if (m.off) chosen.delete(m.after);

  return [...chosen.values()]
    .sort((a, b) => a.after - b.after)
    .map(({ after, seconds, text, source }) => ({ after, seconds, text, source }));
}

// Lectura y escritura del documento abierto en Word (Office.js).
// Las funciones de la primera mitad son puras y se prueban con Node.
import { sanitize, classify, timeline, placeMarkers, labelFor, MARKER } from './timing.js';

export { MARKER };
const QUOTE_STYLE = /cita|quote/i;
const VERSE = /^\d{1,3}\s*\p{Lu}/u;

// Formato de las etiquetas, igual que en el documento de ejemplo.
export const LABEL_FORMAT = Object.freeze({ color: '#A60000', size: 23, bold: true });

const clean = (t) => String(t || '').replace(/[\u0000-\u001f ]+/g, ' ').replace(/\s+/g, ' ').trim();

function headingLevel(p) {
  if (p.styleBuiltIn === 'Title') return 1;
  const n = Number(p.outlineLevel);
  return n >= 1 && n <= 9 ? Math.min(3, n) : 0;
}

// Convierte los párrafos de Word en bloques crudos. `anchors[i]` dice en qué
// párrafo (índice dentro del documento) termina el bloque i y si es una tabla.
export function paragraphsToBlocks(paragraphs) {
  const blocks = [];
  const anchors = [];
  let table = null; // bloque de tabla abierto
  paragraphs.forEach((p, idx) => {
    const text = clean(p.text);
    if (Number(p.tableNestingLevel) > 0) {
      if (!table) {
        table = { text: '', level: 0, indent: false, bold: false, verse: false, table: true, marker: false };
        blocks.push(table);
        anchors.push({ p: idx, table: true });
      }
      if (text) table.text = table.text ? `${table.text} ${text}` : text;
      // el ancla es el último párrafo que pertenece a la tabla exterior
      if (Number(p.tableNestingLevel) === 1) anchors[anchors.length - 1].p = idx;
      return;
    }
    table = null;
    const level = headingLevel(p);
    blocks.push({
      text,
      level,
      indent: !level && !p.isListItem && (Number(p.leftIndent) > 0 || QUOTE_STYLE.test(String(p.style || ''))),
      bold: !level && p.font?.bold === true,
      verse: !level && VERSE.test(text),
      table: false,
      marker: MARKER.test(text),
    });
    anchors.push({ p: idx, table: false, list: !!p.isListItem });
  });
  return { blocks, anchors };
}

// Todo lo que el panel necesita saber del documento.
export function analyze(paragraphs, settingsIn) {
  const settings = sanitize(settingsIn);
  const { blocks, anchors } = paragraphsToBlocks(paragraphs);
  const tl = timeline(classify(blocks), settings);
  const planned = placeMarkers(tl, settings);
  const existing = tl.blocks.filter((b) => b.kind === 'marker');
  return { settings, tl, anchors, planned, existing };
}

// Etiquetas que ya están en el documento, en el formato que usa el modo ensayo.
export function existingAsMarkers(tl) {
  const out = [];
  let prev = null;
  for (const b of tl.blocks) {
    if (b.kind === 'marker') { if (prev) out.push({ after: prev.i, text: b.text }); } else if (b.kind !== 'empty') prev = b;
  }
  return out;
}

// Texto que debería tener cada etiqueta existente según el tiempo real en ese punto.
export function recalculated(tl, settings) {
  return tl.blocks
    .filter((b) => b.kind === 'marker')
    .map((b) => ({ i: b.i, old: b.text, text: labelFor(b.start, settings, tl.total) }))
    .filter((m) => m.text !== m.old);
}

// ---------- Word ----------
const LOAD = 'text,style,styleBuiltIn,outlineLevel,leftIndent,isListItem,tableNestingLevel,font/bold';

async function loadParagraphs(context) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load(LOAD);
  await context.sync();
  return paragraphs.items;
}

function insertLabel(items, anchor, text) {
  const target = items[anchor.p];
  const para = anchor.table
    ? target.parentTable.insertParagraph(text, 'After')
    : target.insertParagraph(text, 'After');
  if (anchor.list) para.detachFromList();
  para.styleBuiltIn = 'Normal';
  para.leftIndent = 0;
  para.firstLineIndent = 0;
  para.alignment = 'Left';
  para.font.bold = LABEL_FORMAT.bold;
  para.font.italic = false;
  para.font.superscript = false;
  para.font.color = LABEL_FORMAT.color;
  para.font.size = LABEL_FORMAT.size;
  return para;
}

/* global Word */
// Copia sencilla de los párrafos, para poder recalcular sin volver a Word.
export async function readParagraphs() {
  return Word.run(async (context) => (await loadParagraphs(context)).map((p) => ({
    text: p.text,
    style: p.style,
    styleBuiltIn: p.styleBuiltIn,
    outlineLevel: p.outlineLevel,
    leftIndent: p.leftIndent,
    isListItem: p.isListItem,
    tableNestingLevel: p.tableNestingLevel,
    font: { bold: p.font.bold },
  })));
}

// Quita las etiquetas que haya y coloca las que dictan las reglas.
export async function placeLabels(settings) {
  return Word.run(async (context) => {
    const items = await loadParagraphs(context);
    const a = analyze(items, settings);
    for (const m of a.planned) insertLabel(items, a.anchors[m.after], m.text);
    for (const b of a.existing) items[a.anchors[b.i].p].delete();
    await context.sync();
    return { placed: a.planned.length, removed: a.existing.length };
  });
}

export async function removeLabels(settings) {
  return Word.run(async (context) => {
    const items = await loadParagraphs(context);
    const a = analyze(items, settings);
    for (const b of a.existing) items[a.anchors[b.i].p].delete();
    await context.sync();
    return { removed: a.existing.length };
  });
}

// Mantiene cada etiqueta donde está y corrige su tiempo.
export async function refreshLabels(settings) {
  return Word.run(async (context) => {
    const items = await loadParagraphs(context);
    const a = analyze(items, settings);
    const changes = recalculated(a.tl, a.settings);
    for (const c of changes) items[a.anchors[c.i].p].insertText(c.text, 'Replace');
    await context.sync();
    return { changed: changes.length, total: a.existing.length };
  });
}

// Añade una etiqueta después del párrafo donde está el cursor.
export async function labelAtCursor(settings) {
  return Word.run(async (context) => {
    const items = await loadParagraphs(context);
    const a = analyze(items, settings);
    const first = context.document.getSelection().paragraphs.getFirst();
    const before = context.document.body.getRange('Start').expandTo(first.getRange('Whole')).paragraphs;
    before.load('items');
    await context.sync();
    const p = before.items.length - 1;
    const i = a.anchors.findIndex((x, k) => p <= x.p && (k === 0 || p > a.anchors[k - 1].p));
    const block = a.tl.blocks[i];
    if (!block || block.kind === 'marker') return { text: null };
    if (!(block.end > 0)) return { text: null };
    const next = a.tl.blocks[i + 1];
    if (next?.kind === 'marker') return { text: null, already: next.text };
    const text = labelFor(block.end, a.settings, a.tl.total);
    insertLabel(items, a.anchors[i], text);
    await context.sync();
    return { text };
  });
}

// Lleva el cursor de Word al bloque i (para saltar desde el panel).
export async function goToBlock(i, settings) {
  return Word.run(async (context) => {
    const items = await loadParagraphs(context);
    const a = analyze(items, settings);
    const anchor = a.anchors[i];
    if (!anchor) return;
    items[anchor.p].getRange('Start').select();
    await context.sync();
  });
}

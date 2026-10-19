/* Armador de propuestas y anexos — ERA SOCIAL
   Todo corre en el navegador: catálogo (catalog.json) + plantilla Word (plantilla.docx). */
'use strict';

// ---------- Configuración ----------
// Para cambiar la contraseña: calcula el SHA-256 de la nueva contraseña y pégalo aquí.
const PASS_HASH = '0b6fbc9050fcab82c4ef5989a77be261cc55ed9f95194c22e887a69d0024801f';
const CATS = [['ASISTENCIAS', 'Asistencias'], ['BENEFICIOS', 'Beneficios'], ['SEGUROS', 'Seguros']];
const MEMS = ['Clásica', 'Tradicional', 'Prestige'];
const BENEF_LABEL = { 'boletos-de-cine': ['Boletos de Cine', 'Boletos de Cine'], 'plataforma-de-aprendizaje-uvelop': ['Plataforma de Aprendizaje Uvelop', 'Uvelop'], 'red-de-descuentos-epic-club': ['Red de Descuentos Epic Club', 'la Red de Descuentos'] };
const BENEF_ORDER = ['boletos-de-cine', 'plataforma-de-aprendizaje-uvelop', 'red-de-descuentos-epic-club'];
const STORE_KEY = 'era-armador-v1';

// ---------- Utilidades XML ----------
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unesc = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const BLK = /<w:p[ >](?:(?!<w:p[ >])[\s\S])*?<\/w:p>|<w:tbl>[\s\S]*?<\/w:tbl>/g;
const RUN = /<w:r(?: [^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g;
const ptext = b => unesc((b.match(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join('')).trim();
const isTable = b => b.startsWith('<w:tbl');
const EMPTY = '<w:p><w:pPr><w:pStyle w:val="Sinespaciado"/><w:jc w:val="both"/></w:pPr></w:p>';
const HL_BEFORE = ['w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath'];

function hlRun(run, on = true) {
  if (!on) return run.replace(/<w:highlight w:val="[^"]*"\/>/g, '');
  if (run.includes('<w:highlight')) return run.replace(/<w:highlight w:val="[^"]*"\/>/, '<w:highlight w:val="yellow"/>');
  const m = run.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
  if (!m) return run.replace(/^(<w:r(?: [^>]*)?>)/, '$1<w:rPr><w:highlight w:val="yellow"/></w:rPr>');
  let inner = m[1], pos = inner.length;
  for (const t of HL_BEFORE) { const i = inner.search(new RegExp('<' + t + '[ />]')); if (i >= 0 && i < pos) pos = i; }
  inner = inner.slice(0, pos) + '<w:highlight w:val="yellow"/>' + inner.slice(pos);
  return run.replace(m[0], '<w:rPr>' + inner + '</w:rPr>');
}
// Reemplaza el texto completo de un párrafo (conserva formato del primer run).
function setParaText(p, text, hl = false) {
  p = p.replace(/<w:bookmark(?:Start|End)\b[^>]*\/>/g, '');
  let first = true;
  p = p.replace(RUN, run => {
    if (!/<w:t(?: [^>]*)?>/.test(run)) return run;
    if (first) {
      first = false;
      let ft = true;
      const r = run.replace(/<w:t(?: [^>]*)?>[^<]*<\/w:t>/g, () => { if (ft) { ft = false; return '<w:t xml:space="preserve">' + esc(text) + '</w:t>'; } return ''; });
      return hl ? hlRun(r) : r;
    }
    return '';
  });
  return p;
}
// Reemplaza una cadena literal dentro de los runs, marcando en amarillo lo nuevo.
function replaceInPara(p, find, repl, hl = true) {
  const fe = esc(find); let count = 0;
  const out = p.replace(RUN, run => {
    const tm = run.match(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/);
    if (!tm || !tm[1].includes(fe)) return run;
    const parts = tm[1].split(fe); count += parts.length - 1;
    const mk = (t, h) => { const r = run.replace(tm[0], '<w:t xml:space="preserve">' + t + '</w:t>'); return h ? hlRun(r) : r; };
    let s = '';
    parts.forEach((pt, i) => { if (pt) s += mk(pt, false); if (i < parts.length - 1) s += mk(esc(repl), hl); });
    return s;
  });
  if (count === 0 && ptext(p).includes(find)) return [setParaText(p, ptext(p).split(find).join(repl), hl), 1];
  return [out, count];
}

// ---------- Números en letra ----------
const U = ['', 'uno', 'dos', 'tres', 'cuatro', 'cinco', 'seis', 'siete', 'ocho', 'nueve', 'diez', 'once', 'doce', 'trece', 'catorce', 'quince', 'dieciséis', 'diecisiete', 'dieciocho', 'diecinueve', 'veinte', 'veintiuno', 'veintidós', 'veintitrés', 'veinticuatro', 'veinticinco', 'veintiséis', 'veintisiete', 'veintiocho', 'veintinueve'];
const DEC = ['', '', '', 'treinta', 'cuarenta', 'cincuenta', 'sesenta', 'setenta', 'ochenta', 'noventa'];
const CEN = ['', 'ciento', 'doscientos', 'trescientos', 'cuatrocientos', 'quinientos', 'seiscientos', 'setecientos', 'ochocientos', 'novecientos'];
const apoc = s => s.replace(/veintiuno$/, 'veintiún').replace(/uno$/, 'un');
function n999(n) { if (!n) return ''; if (n === 100) return 'cien'; const r = n % 100; let t = r < 30 ? U[r] : DEC[Math.floor(r / 10)] + (r % 10 ? ' y ' + U[r % 10] : ''); return (CEN[Math.floor(n / 100)] + ' ' + t).trim(); }
function nWords(n) {
  if (!n) return 'cero';
  const mi = Math.floor(n / 1e6), th = Math.floor((n % 1e6) / 1000), r = n % 1000, o = [];
  if (mi) o.push(mi === 1 ? 'un millón' : apoc(nWords(mi)) + ' millones');
  if (th) o.push(th === 1 ? 'mil' : apoc(n999(th)) + ' mil');
  if (r) o.push(n999(r));
  return o.join(' ');
}
function pesosWords(v, cap = false) {
  const ent = Math.floor(v + 1e-9), c = Math.round((v - ent) * 100);
  let w = apoc(nWords(ent)) + (ent >= 1e6 && ent % 1e6 === 0 ? ' de' : '') + ' pesos ' + String(c).padStart(2, '0') + '/100 M.N.';
  return cap ? w[0].toUpperCase() + w.slice(1) : w;
}
const fmtMoney = v => '$' + Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const parseMoney = s => { const v = parseFloat(String(s).replace(/[$,\s]/g, '')); return isFinite(v) && v >= 0 ? v : null; };

// ---------- Parámetros editables (montos y eventos) ----------
const NOUNS = { evento: 'eventos', consulta: 'consultas', 'envío': 'envíos', 'cupón': 'cupones', estudio: 'estudios', mascota: 'mascotas', 'sesión': 'sesiones' };
const PLUR2S = Object.fromEntries(Object.entries(NOUNS).map(([s, p]) => [p, s]));
const MONEY_RE = /\$(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)( MXN)?( \([^()]*?M\.N\.\))?/g;
const COUNT_RE = new RegExp('(?<![\\d\\w])(\\d{1,2}) (' + [...Object.keys(NOUNS), ...Object.values(NOUNS)].join('|') + ')(?![\\wáéíóúñ])', 'g');

function detectParams(v) {
  const P = {};
  v.xml.forEach((b, i) => {
    if (i === 0) return;
    const t = ptext(b);
    for (const m of t.matchAll(MONEY_RE)) {
      const k = 'm:' + m[1];
      (P[k] = P[k] || { key: k, type: 'money', value: parseMoney(m[1]), label: fmtMoney(parseMoney(m[1])), paras: new Set() }).paras.add(i);
    }
    for (const m of t.matchAll(COUNT_RE)) {
      const sing = PLUR2S[m[2]] || m[2], k = 'c:' + m[1] + ':' + sing;
      (P[k] = P[k] || { key: k, type: 'count', value: +m[1], noun: sing, label: m[1] + ' ' + (m[1] === '1' ? sing : NOUNS[sing]), paras: new Set() }).paras.add(i);
    }
  });
  return Object.values(P);
}
// Devuelve [nuevoTextoPlano, lista de [literalViejo, literalNuevo]] para un párrafo
function paramReplacements(text, prm, nv) {
  const pairs = [];
  if (prm.type === 'money') {
    for (const m of text.matchAll(MONEY_RE)) {
      if ('m:' + m[1] !== prm.key) continue;
      let s = fmtMoney(nv) + (m[2] || '') + (m[3] ? ' (' + pesosWords(nv) + ')' : '');
      pairs.push([m[0], s]);
    }
  } else {
    for (const m of text.matchAll(COUNT_RE)) {
      const sing = PLUR2S[m[2]] || m[2];
      if ('c:' + m[1] + ':' + sing !== prm.key) continue;
      let a = m[0], c = nv + ' ' + (+nv === 1 ? sing : NOUNS[sing]);
      const pre = text.slice(0, m.index).match(/(Se cubren?|Incluye|Incluyen) $/);
      if (pre) { a = pre[0] + a; c = (/^Se cubre/.test(pre[1]) ? (+nv === 1 ? 'Se cubre ' : 'Se cubren ') : pre[0]) + c; }
      pairs.push([a, c]);
    }
  }
  const uniq = [...new Map(pairs.map(p => [p[0], p])).values()];
  let nt = text; uniq.forEach(([a, b]) => { nt = nt.split(a).join(b); });
  return [nt, uniq];
}

// ---------- Estado ----------
let CAT = null, TPL = null;
const state = { datos: { nombre: 'Contigo+ ', cliente: '', rep: '', fecha: '', precio: '', vigencia: '30' }, sel: {}, edits: {} };
const byId = {}; const vById = {};

function save() { try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); } catch (e) { } }
function load() { try { const s = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); if (s && s.sel) Object.assign(state, s); } catch (e) { } }
const edOf = vid => (state.edits[vid] = state.edits[vid] || { params: {}, paras: {} });
const isModified = vid => { const e = state.edits[vid]; return !!e && (Object.keys(e.params).length + Object.keys(e.paras).length) > 0; };

// ---------- Construcción del XML de un servicio con sus cambios ----------
function buildVariant(v) {
  const e = state.edits[v.id] || { params: {}, paras: {} };
  let xml = v.xml.slice();
  for (const [i, t] of Object.entries(e.paras)) xml[+i] = setParaText(xml[+i], t, true);
  for (const prm of v.params) {
    if (!(prm.key in e.params)) continue;
    const nv = e.params[prm.key];
    xml = xml.map((b, i) => {
      if (i === 0 || isTable(b)) return b;
      const [, pairs] = paramReplacements(ptext(b), prm, nv);
      for (const [a, c] of pairs) b = replaceInPara(b, a, c)[0];
      return b;
    });
  }
  return xml;
}
function sectionLines(xml, label, orig) {
  const out = []; let on = false;
  xml.forEach((b, i) => {
    if (i === 0) return;
    const t = ptext(b);
    if (t.endsWith(':') && t.length < 60 && !b.includes('<w:numPr>')) { on = t.startsWith(label); return; }
    if (on && t) out.push({ t, hl: !!orig && b !== orig[i] });
  });
  return out;
}

// ---------- Selección ----------
function selectedList() {
  const L = [];
  for (const [cat] of CATS) for (const s of CAT.services.filter(s => s.cat === cat)) if (state.sel[s.id]) L.push({ s, v: vById[state.sel[s.id]] });
  return L;
}
function warnings() {
  const W = []; const L = selectedList(); const names = new Set(L.map(x => x.s.name));
  for (const { s, v } of L) {
    const miss = v.refs.filter(r => !names.has(r));
    if (miss.length) W.push({ s: s.id, msg: `${v.heading} menciona ${miss.join(', ')}, que no está incluido. Revisa el texto o agrega el servicio.` });
  }
  if (!L.length) W.push({ msg: 'No has elegido servicios.' });
  if (!parseMoney(state.datos.precio)) W.push({ msg: 'Falta el precio mensual (quedará como $xxxx en amarillo).' });
  const segs = L.filter(x => x.s.cat === 'SEGUROS');
  if (segs.length && !(L.find(x => x.s.id === 'seguro-de-muerte-accidental') && L.find(x => x.s.id === 'perdida-organica-escala-b')))
    W.push({ msg: 'Las Disposiciones Generales de los Seguros se incluyen completas; se omite la regla de Pérdida Orgánica + Muerte porque no van ambos seguros.' , info: true });
  return W;
}

// ---------- UI ----------
const $ = (s, r = document) => r.querySelector(s);
const h = (tag, attrs = {}, ...kids) => {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') el.className = v; else if (k.startsWith('on')) el.addEventListener(k.slice(2), v); else if (v !== false && v != null) el.setAttribute(k, v === true ? '' : v);
  }
  for (const k of kids.flat()) if (k != null && k !== false) el.append(k.nodeType ? k : document.createTextNode(k));
  return el;
};
const open = new Set();

function render() {
  renderDatos(); renderServicios(); renderResumen(); save();
}
function renderDatos() {
  for (const k of Object.keys(state.datos)) { const el = $('#f-' + k); if (el && document.activeElement !== el) el.value = state.datos[k]; }
  const v = parseMoney(state.datos.precio);
  $('#precio-letra').textContent = v ? `${fmtMoney(v)} (${pesosWords(v, true)}) + IVA` : 'Sin precio: quedará $xxxx en amarillo';
}
function renderServicios() {
  const root = $('#servicios'); root.innerHTML = '';
  for (const [cat, label] of CATS) {
    const list = CAT.services.filter(s => s.cat === cat);
    root.append(h('h3', { class: 'cat' }, label));
    for (const s of list) root.append(serviceRow(s));
  }
}
function serviceRow(s) {
  const vid = state.sel[s.id]; const on = !!vid; const v = on ? vById[vid] : null;
  const row = h('div', { class: 'svc' + (on ? ' on' : '') });
  const head = h('div', { class: 'svc-head' },
    h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: on, onchange: e => { if (e.target.checked) state.sel[s.id] = s.variants[s.variants.length - 1].id; else { delete state.sel[s.id]; open.delete(s.id); } render(); } }), h('span', { class: 'svc-name' }, s.name)),
    h('div', { class: 'seg' }, s.variants.map(x => h('button', {
      type: 'button', class: 'seg-b' + (vid === x.id ? ' act' : ''), title: 'Versión de ' + x.mem.join(', '),
      onclick: () => { state.sel[s.id] = x.id; render(); }
    }, x.label))),
    on ? h('span', { class: 'tags' }, isModified(vid) ? h('span', { class: 'tag mod' }, 'Modificado') : null,
      warnings().some(w => w.s === s.id) ? h('span', { class: 'tag warn', title: 'Menciona un servicio no incluido' }, 'Revisar') : null) : null,
    on ? h('button', { type: 'button', class: 'link', onclick: () => { open.has(s.id) ? open.delete(s.id) : open.add(s.id); render(); } }, open.has(s.id) ? 'Ocultar' : 'Detalle') : null
  );
  row.append(head);
  if (on && open.has(s.id)) row.append(detailPanel(s, v));
  return row;
}
function detailPanel(s, v) {
  const e = edOf(v.id); const xml = buildVariant(v);
  const box = h('div', { class: 'detail' });
  const miss = warnings().filter(w => w.s === s.id);
  miss.forEach(w => box.append(h('div', { class: 'alert' }, w.msg)));
  const lim = sectionLines(xml, 'Límites Variables', v.xml);
  if (lim.length) box.append(h('div', { class: 'lim' }, h('div', { class: 'lbl' }, 'Límites Variables'), lim.map(l => h('div', { class: l.hl ? 'hlt' : '' }, l.t))));
  if (v.params.length) {
    box.append(h('div', { class: 'lbl' }, 'Montos y eventos'));
    const grid = h('div', { class: 'params' });
    for (const prm of v.params) {
      const cur = prm.key in e.params ? e.params[prm.key] : prm.value;
      const inp = h('input', { type: 'text', inputmode: 'decimal', value: prm.type === 'money' ? fmtMoney(cur) : cur, class: prm.key in e.params ? 'chg' : '' });
      const apply = () => {
        const raw = inp.value; const nv = prm.type === 'money' ? parseMoney(raw) : parseInt(raw, 10);
        if (nv == null || isNaN(nv) || (prm.type === 'count' && (nv < 1 || nv > 99))) { inp.value = prm.type === 'money' ? fmtMoney(cur) : cur; return; }
        if (nv === cur) { inp.value = prm.type === 'money' ? fmtMoney(cur) : cur; return; }
        confirmParam(s, v, prm, nv, () => { if (nv === prm.value) delete e.params[prm.key]; else e.params[prm.key] = nv; render(); }, () => { inp.value = prm.type === 'money' ? fmtMoney(cur) : cur; });
      };
      inp.addEventListener('change', apply);
      inp.addEventListener('keydown', ev => { if (ev.key === 'Enter') inp.blur(); });
      grid.append(h('label', { class: 'prm' }, h('span', {}, (prm.type === 'money' ? 'Monto ' : 'Cantidad: ') + prm.label + ` · en ${prm.paras.size} ${prm.paras.size === 1 ? 'línea' : 'líneas'}`), inp,
        prm.key in e.params ? h('button', { type: 'button', class: 'link', onclick: () => { delete e.params[prm.key]; render(); } }, 'Deshacer') : null));
    }
    box.append(grid);
  }
  box.append(h('div', { class: 'row-actions' },
    h('button', { type: 'button', class: 'btn sec', onclick: () => textEditor(s, v) }, 'Ver y editar el texto completo'),
    isModified(v.id) ? h('button', { type: 'button', class: 'btn ghost', onclick: () => { if (confirm('¿Quitar todos los cambios de este servicio?')) { delete state.edits[v.id]; render(); } } }, 'Quitar todos los cambios') : null));
  return box;
}
function diffHtml(oldT, newT) {
  // marca las partes nuevas
  let a = 0; while (a < oldT.length && a < newT.length && oldT[a] === newT[a]) a++;
  let b = 0; while (b < oldT.length - a && b < newT.length - a && oldT[oldT.length - 1 - b] === newT[newT.length - 1 - b]) b++;
  while (a > 0 && !/\s/.test(oldT[a - 1])) a--;
  while (b > 0 && !/\s/.test(oldT[oldT.length - b] || ' ')) b--;
  const e2 = x => esc(x);
  return { old: e2(oldT.slice(0, a)) + '<del>' + e2(oldT.slice(a, oldT.length - b)) + '</del>' + e2(oldT.slice(oldT.length - b)), nu: e2(newT.slice(0, a)) + '<mark>' + e2(newT.slice(a, newT.length - b)) + '</mark>' + e2(newT.slice(newT.length - b)) };
}
function modal(title, bodyEl, okText, onOk, onCancel) {
  const m = $('#modal'); $('#m-title').textContent = title; const b = $('#m-body'); b.innerHTML = ''; b.append(bodyEl);
  $('#m-ok').textContent = okText; m.hidden = false;
  const close = () => { m.hidden = true; $('#m-ok').onclick = null; $('#m-cancel').onclick = null; };
  $('#m-ok').onclick = () => { if (onOk() !== false) close(); };
  $('#m-cancel').onclick = () => { close(); onCancel && onCancel(); };
}
function confirmParam(s, v, prm, nv, ok, cancel) {
  const xml = buildVariant(v); const body = h('div');
  body.append(h('p', {}, `Vas a cambiar ${prm.label} por ${prm.type === 'money' ? fmtMoney(nv) : nv + ' ' + (nv === 1 ? prm.noun : NOUNS[prm.noun])} en ${v.heading} (versión ${v.label}). Estas líneas cambian:`));
  for (const i of prm.paras) {
    const t = ptext(xml[i]); const [nt] = paramReplacements(t, prm, nv); if (nt === t) continue;
    const d = diffHtml(t, nt);
    const it = h('div', { class: 'diff' }); it.innerHTML = `<div class="o">${d.old}</div><div class="n">${d.nu}</div>`; body.append(it);
  }
  body.append(h('p', { class: 'muted' }, 'En el Word, lo modificado queda marcado en amarillo para revisión.'));
  modal('Confirmar cambio', body, 'Sí, aplicar el cambio', ok, cancel);
}
function textEditor(s, v) {
  const e = edOf(v.id); const body = h('div', { class: 'editor' });
  body.append(h('p', { class: 'muted' }, 'Toca el lápiz de una línea para editarla. Cada cambio pide confirmación y queda en amarillo en el Word.'));
  const xml = buildVariant(v);
  v.xml.forEach((b, i) => {
    if (i === 0) return;
    const orig = ptext(b); if (!orig && !isTable(b)) return;
    if (isTable(b)) { body.append(h('div', { class: 'ln tbl' }, '[Tabla: se conserva igual]')); return; }
    const cur = ptext(xml[i]);
    const isLabel = /:$/.test(orig) && orig.length < 50 && !b.includes('<w:numPr>');
    const line = h('div', { class: 'ln' + (isLabel ? ' lab' : '') + (i in e.paras ? ' hlt' : '') + (b.includes('<w:ilvl w:val="1"/>') ? ' sub' : b.includes('<w:numPr>') ? ' bul' : '') });
    const txt = h('span', {}, cur);
    line.append(txt);
    if (!isLabel) line.append(h('button', {
      type: 'button', class: 'pen', title: 'Editar', onclick: () => {
        const ta = h('textarea', { rows: Math.min(8, Math.ceil(cur.length / 70) + 1) }); ta.value = cur;
        const save = h('button', { type: 'button', class: 'btn' }, 'Revisar cambio');
        const undo = i in e.paras ? h('button', { type: 'button', class: 'btn ghost' }, 'Volver al original') : null;
        line.innerHTML = ''; line.append(ta, h('div', { class: 'row-actions' }, save, undo));
        ta.focus();
        save.onclick = () => {
          const nt = ta.value.replace(/\s+/g, ' ').trim(); if (!nt || nt === cur) { textEditor(s, v); return; }
          const d = diffHtml(cur, nt); const bb = h('div'); bb.append(h('p', {}, 'Confirma el cambio de esta línea:'));
          const it = h('div', { class: 'diff' }); it.innerHTML = `<div class="o">${d.old}</div><div class="n">${d.nu}</div>`; bb.append(it);
          modal('Confirmar cambio de texto', bb, 'Sí, aplicar', () => { e.paras[i] = nt; render(); setTimeout(() => textEditor(s, v), 0); }, () => textEditor(s, v));
        };
        if (undo) undo.onclick = () => { delete e.paras[i]; render(); textEditor(s, v); };
      }
    }, '✎'));
    body.append(line);
  });
  modal(v.heading + ' · ' + v.label, body, 'Cerrar', () => { render(); });
  $('#m-cancel').style.display = 'none';
  const restore = () => { $('#m-cancel').style.display = ''; };
  $('#m-ok').addEventListener('click', restore, { once: true });
}
function renderResumen() {
  const L = selectedList(); const ol = $('#res-list'); ol.innerHTML = '';
  let n = 0;
  for (const [cat, label] of CATS) {
    const items = L.filter(x => x.s.cat === cat); if (!items.length) continue;
    ol.append(h('li', { class: 'rcat' }, label));
    for (const { s, v } of items) ol.append(h('li', {}, h('span', { class: 'num' }, ++n + '.'), ' ' + v.heading, h('span', { class: 'muted' }, ' · ' + v.label), isModified(v.id) ? h('span', { class: 'dot', title: 'Modificado' }) : null));
  }
  if (!L.length) ol.append(h('li', { class: 'muted' }, 'Aún no hay servicios.'));
  const W = warnings(); const wl = $('#res-warn'); wl.innerHTML = '';
  W.forEach(w => wl.append(h('li', { class: w.info ? 'info' : '' }, w.msg)));
  $('#res-count').textContent = L.length ? `${L.length} servicios` : '';
}
function base(mem) {
  state.sel = {};
  for (const s of CAT.services) { const v = s.variants.find(x => x.mem.includes(mem)); if (v) state.sel[s.id] = v.id; }
  state.datos.nombre = 'Contigo+ ' + mem;
  open.clear(); render();
}

// ---------- Generación del anexo ----------
async function loadTpl() { if (!TPL) TPL = await (await fetch('plantilla.docx')).arrayBuffer(); return await JSZip.loadAsync(TPL); }
function splitDoc(x) {
  const bs = x.indexOf('<w:body>') + 8; const be = x.lastIndexOf('<w:sectPr');
  const blocks = x.slice(bs, be).match(BLK) || [];
  return { head: x.slice(0, bs), blocks, tail: x.slice(be) };
}
const findIdx = (B, pred, from = 0) => { for (let i = from; i < B.length; i++) if (pred(ptext(B[i]), B[i])) return i; return -1; };
function listEs(arr) { return arr.length <= 1 ? arr.join('') : arr.slice(0, -1).join(', ') + ' y ' + arr[arr.length - 1]; }

async function makeAnexo() {
  const zip = await loadTpl(); let x = await zip.file('word/document.xml').async('string');
  const D = splitDoc(x); let B = D.blocks; const d = state.datos;
  const L = selectedList();
  const hasSeg = L.some(i => i.s.cat === 'SEGUROS');
  const ben = BENEF_ORDER.filter(id => state.sel[id]);
  const nombre = (d.nombre || '').trim() || 'xxxx';
  // Encabezado
  B[0] = B[0].replace(RUN, run => {
    const tm = run.match(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/); const t = tm && tm[1];
    const put = (r, nt) => r.replace(tm[0], '<w:t xml:space="preserve">' + nt + '</w:t>');
    if (t === '_________,' && d.fecha) return hlRun(put(run, esc(fechaLarga(d.fecha).toUpperCase()) + ','), false);
    if (t === '______' && d.cliente.trim()) return hlRun(put(run, esc(d.cliente.trim().toUpperCase())), false);
    if (t === '__________' && d.rep.trim()) return hlRun(put(run, ' ' + esc(d.rep.trim().toUpperCase())), false);
    return run;
  });
  // Título
  let k = findIdx(B, t => t.startsWith('“MEMBRESÍA'));
  B[k] = replaceInPara(B[k], 'MEMBRESÍA CONTIGO+ PRESTIGE', 'MEMBRESÍA ' + nombre.toUpperCase(), false)[0];
  // Rubros
  k = findIdx(B, t => t.startsWith('Los Servicios se dividen en los siguientes rubros'));
  const rub = [];
  if (L.some(i => i.s.cat === 'ASISTENCIAS')) rub.push('Asistencias');
  if (ben.length) rub.push('Beneficios (' + listEs(ben.map(id => BENEF_LABEL[id][0])) + ')');
  if (hasSeg) rub.push('Seguros');
  B[k] = setParaText(B[k], rub.length > 1 ? 'Los Servicios se dividen en los siguientes rubros: ' + listEs(rub) + '.' : 'Los Servicios corresponden al rubro de ' + (rub[0] || 'xxxx') + '.');
  // Definición Cabina de Atención a Clientes
  k = findIdx(B, t => t.startsWith('Cabina de Atención a Clientes:'));
  const tailSeg = ', así como el reporte y seguimiento de los siniestros de los Seguros.';
  const kill = new Set();
  if (!ben.length && !hasSeg) kill.add(k);
  else if (!hasSeg) B[k] = replaceInPara(B[k], tailSeg, '.', false)[0];
  else if (!ben.length) B[k] = replaceInPara(B[k], ' área que atiende las dudas o problemas de acceso de los Usuarios a los beneficios digitales' + tailSeg, ' área que atiende el reporte y seguimiento de los siniestros de los Seguros.', false)[0];
  // Párrafo de proveedores de Beneficios y de Seguros (con su renglón vacío)
  k = findIdx(B, t => t.startsWith('En caso de que algún proveedor de los Beneficios'));
  if (!ben.length && k >= 0) { kill.add(k); if (!ptext(B[k + 1])) kill.add(k + 1); }
  k = findIdx(B, t => t.startsWith('Los Seguros incluidos en la Membresía son otorgados'));
  if (!hasSeg && k >= 0) { kill.add(k); if (!ptext(B[k + 1])) kill.add(k + 1); }
  if (!ben.length) { k = findIdx(B, t => t.startsWith('Aplicación:')); if (k >= 0) B[k] = replaceInPara(B[k], 'realizar la Activación, solicitar los Servicios y acceder a los beneficios digitales.', 'realizar la Activación y solicitar los Servicios.', false)[0]; }
  // Cláusula de la App: beneficios digitales
  k = findIdx(B, t => t.startsWith('Acceso a los beneficios digitales ('));
  if (k >= 0) { if (!ben.length) kill.add(k); else B[k] = setParaText(B[k], 'Acceso a los beneficios digitales (' + listEs(ben.map(id => BENEF_LABEL[id][0])) + ') directamente desde la Aplicación.'); }
  // Precio
  k = findIdx(B, t => /^\$61\.91/.test(t));
  const pv = parseMoney(d.precio);
  B[k] = pv ? setParaText(B[k], `${fmtMoney(pv)} (${pesosWords(pv, true)})`) : setParaText(B[k], '$xxxx (xxxx pesos xx/100 M.N.)', true);

  // Índice
  const iL = findIdx(B, t => t === 'Listado de coberturas');
  const iS = findIdx(B, t => t === 'ASISTENCIAS', iL);
  let iE = iS; while (iE + 1 < B.length && (B[iE + 1].includes('<w:numId w:val="25"/>') || ['BENEFICIOS', 'SEGUROS'].includes(ptext(B[iE + 1])))) iE++;
  const subT = B[iS], itemT = B.slice(iS, iE + 1).find(b => b.includes('<w:numId w:val="25"/>'));
  const idx = [];
  for (const [cat] of CATS) {
    const items = L.filter(i => i.s.cat === cat); if (!items.length) continue;
    idx.push(setParaText(subT, cat)); items.forEach(i => idx.push(setParaText(itemT, i.v.heading)));
  }
  // Detalle
  const dD = findIdx(B, t => t === 'Detalle de las coberturas');
  const dP = findIdx(B, t => t.startsWith('Para dudas o problemas de acceso a los beneficios digitales'), dD);
  const tel = ', el Usuario puede comunicarse a la Cabina de Atención a Clientes, en el mismo teléfono de Cabina de Asistencia.';
  if (!ben.length && !hasSeg) kill.add(dP);
  else {
    const bl = ben.map(id => BENEF_LABEL[id][1]);
    let t = ben.length ? 'Para dudas o problemas de acceso a los beneficios digitales (como ' + (bl.length > 1 ? bl.slice(0, -1).join(', ') + ' o ' + bl[bl.length - 1] : bl[0]) + ')' + (hasSeg ? ', así como para el reporte y seguimiento de siniestros de los Seguros' : '') : 'Para el reporte y seguimiento de siniestros de los Seguros';
    B[dP] = setParaText(B[dP], t + tel);
  }
  const dS = findIdx(B, t => t === 'ASISTENCIAS', dD);
  const dE = findIdx(B, t => t.startsWith('TERCERA.'), dS);
  const subD = B[dS];
  const det = [];
  for (const [cat] of CATS) {
    const items = L.filter(i => i.s.cat === cat); if (!items.length) continue;
    det.push(setParaText(subD, cat), EMPTY);
    for (const i of items) { det.push(...buildVariant(i.v), EMPTY); }
    if (cat === 'SEGUROS') {
      const both = L.find(i => i.s.id === 'seguro-de-muerte-accidental') && L.find(i => i.s.id === 'perdida-organica-escala-b');
      det.push(...CAT.disposiciones.filter(b => both || !ptext(b).startsWith('Si como consecuencia de un mismo accidente se paga una indemnización por Pérdida Orgánica')), EMPTY);
    }
  }
  // Ensamblar
  const out = [];
  for (let i = 0; i < B.length; i++) {
    if (i === iS) { out.push(...idx); i = iE; continue; }
    if (i === dS) { out.push(...det); i = dE - 1; continue; }
    if (kill.has(i)) continue;
    out.push(B[i]);
  }
  zip.file('word/document.xml', D.head + out.join('') + D.tail);
  let core = await zip.file('docProps/core.xml').async('string');
  core = core.replace(/<dc:title>[^<]*<\/dc:title>/, '<dc:title>' + esc('Anexo A - Membresía ' + nombre) + '</dc:title>').replace(/<dc:creator>[^<]*<\/dc:creator>/, '<dc:creator>ERA SOCIAL</dc:creator>');
  zip.file('docProps/core.xml', core);
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}
function fechaLarga(iso) {
  if (!iso) return '';
  const [y, m, dd] = iso.split('-').map(Number);
  const M = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  return `${dd} de ${M[m - 1]} de ${y}`;
}

// ---------- Propuesta comercial ----------
const TEAL = '178374';
function rr(text, o = {}) {
  return `<w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:cstheme="majorHAnsi"/>${o.b ? '<w:b/>' : ''}<w:color w:val="${o.color || '000000'}"/><w:sz w:val="${o.sz || 21}"/><w:szCs w:val="${o.sz || 21}"/>${o.hl ? '<w:highlight w:val="yellow"/>' : ''}</w:rPr><w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
}
function pp(runs, o = {}) {
  return `<w:p><w:pPr>${o.keep ? '<w:keepNext/>' : ''}${o.bul ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="11"/></w:numPr>' : ''}${o.border ? `<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="${TEAL}"/></w:pBdr>` : ''}<w:spacing w:before="${o.before || 0}" w:after="${o.after == null ? 120 : o.after}"/>${o.ind ? `<w:ind w:left="${o.ind}"/>` : ''}<w:jc w:val="${o.jc || 'both'}"/></w:pPr>${Array.isArray(runs) ? runs.join('') : runs}</w:p>`;
}
function cell(content, w, o = {}) {
  return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${o.span ? `<w:gridSpan w:val="${o.span}"/>` : ''}${o.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>` : ''}</w:tcPr>${content}</w:tc>`;
}
async function makePropuesta() {
  const zip = await loadTpl(); const x = await zip.file('word/document.xml').async('string');
  const D = splitDoc(x); const d = state.datos; const L = selectedList();
  const nombre = (d.nombre || '').trim() || 'xxxx'; const cli = d.cliente.trim();
  const pv = parseMoney(d.precio);
  const body = [];
  body.push(pp(rr('ERA SOCIAL', { b: true, color: TEAL, sz: 22 }), { after: 0 }));
  body.push(pp(rr('Propuesta comercial', { b: true, sz: 36 }), { after: 60 }));
  body.push(pp(rr('Membresía ' + nombre, { b: true, color: TEAL, sz: 28 }), { border: true, after: 240 }));
  body.push(pp([rr('Para: ', { b: true }), rr(cli || 'xxxx', { hl: !cli })], { after: 40 }));
  body.push(pp([rr('Fecha: ', { b: true }), rr(d.fecha ? fechaLarga(d.fecha) : 'xxxx', { hl: !d.fecha })], { after: 240 }));
  body.push(pp(rr(`ERA SOCIAL presenta ${cli ? 'a ' + cli : 'a usted'} la propuesta de la Membresía ${nombre}, que incluye los servicios que se describen a continuación. Las condiciones, exclusiones y procedimientos de cada servicio se detallan en el Anexo A del contrato.`), { after: 240 }));
  // Tabla
  const W1 = 3400, W2 = 5438;
  const bd = `<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders>`;
  let rows = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(pp(rr('Servicio', { b: true, color: 'FFFFFF' }), { after: 0, jc: 'left' }), W1, { fill: TEAL })}${cell(pp(rr('Cobertura', { b: true, color: 'FFFFFF' }), { after: 0, jc: 'left' }), W2, { fill: TEAL })}</w:tr>`;
  let n = 0;
  for (const [cat, label] of CATS) {
    const items = L.filter(i => i.s.cat === cat); if (!items.length) continue;
    rows += `<w:tr>${cell(pp(rr(label.toUpperCase(), { b: true, sz: 19 }), { after: 0, jc: 'left' }), W1 + W2, { span: 2, fill: 'E8F3F1' })}</w:tr>`;
    for (const i of items) {
      const xml = buildVariant(i.v);
      const desc = sectionLines(xml, 'Descripción').map(l => l.t).join(' ');
      let lim = sectionLines(xml, 'Límites Variables', i.v.xml);
      if (!lim.length) lim = xml.map((b, j) => ({ t: ptext(b), hl: b !== i.v.xml[j] })).filter(l => /^(Suma asegurada|Asegurado)/.test(l.t));
      const c1 = pp(rr(++n + '. ' + i.v.heading, { b: true }), { after: 60, jc: 'left' }) + (desc ? pp(rr(desc, { sz: 18, color: '595959' }), { after: 40, jc: 'left' }) : '');
      const c2 = (lim.length ? lim : [{ t: 'Según el Anexo A.' }]).map(l => pp(rr(l.t, { hl: l.hl }), { after: 60, jc: 'left' })).join('');
      rows += `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(c1, W1)}${cell(c2, W2)}</w:tr>`;
    }
  }
  body.push(`<w:tbl><w:tblPr><w:tblW w:w="${W1 + W2}" w:type="dxa"/>${bd}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="${W1}"/><w:gridCol w:w="${W2}"/></w:tblGrid>${rows}</w:tbl>`);
  body.push(pp(rr('Inversión', { b: true, color: TEAL, sz: 26 }), { before: 360, after: 120, border: true, keep: true }));
  body.push(pp([rr('Costo mensual por Membresía: ', { b: true }), pv ? rr(`${fmtMoney(pv)} (${pesosWords(pv, true)}) más IVA.`) : rr('$xxxx más IVA.', { hl: true })], { after: 120 }));
  body.push(pp(rr('Condiciones generales', { b: true, color: TEAL, sz: 26 }), { before: 240, after: 120, border: true, keep: true }));
  const vig = parseInt(d.vigencia, 10);
  [
    'Los eventos y límites de cada servicio se renuevan al iniciar cada año de vigencia de la Membresía, contado a partir de su Activación; los eventos no utilizados no son acumulables.',
    'Los Servicios se solicitan a través de Cabina de Asistencia y se prestan dentro de la República Mexicana, salvo que en cada servicio se indique otra cosa.',
    'Las condiciones, exclusiones y procedimientos completos de cada servicio se establecen en el Anexo A del contrato.',
  ].forEach(t => body.push(pp(rr(t), { bul: true, after: 60 })));
  body.push(pp([rr('Esta propuesta tiene una vigencia de '), rr(vig > 0 ? `${vig} días naturales` : 'xx días naturales', { hl: !(vig > 0) }), rr(' a partir de su fecha.')], { bul: true, after: 60 }));
  if (L.some(i => isModified(i.v.id))) body.push(pp(rr('Los datos marcados en amarillo son condiciones ajustadas para esta propuesta y están sujetos a validación.', { sz: 18, color: '595959' }), { before: 240 }));
  zip.file('word/document.xml', D.head + body.join('') + D.tail);
  let core = await zip.file('docProps/core.xml').async('string');
  core = core.replace(/<dc:title>[^<]*<\/dc:title>/, '<dc:title>' + esc('Propuesta comercial - Membresía ' + nombre) + '</dc:title>').replace(/<dc:creator>[^<]*<\/dc:creator>/, '<dc:creator>ERA SOCIAL</dc:creator>');
  zip.file('docProps/core.xml', core);
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

// ---------- Descargas y archivos de configuración ----------
function download(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
const slug = () => ((state.datos.nombre || 'membresia').trim() + (state.datos.cliente ? ' ' + state.datos.cliente.trim() : '')).replace(/[^\wáéíóúñÁÉÍÓÚÑ+]+/g, '_').replace(/^_|_$/g, '');
async function guard(btn, fn) {
  const t = btn.textContent; btn.disabled = true; btn.textContent = 'Generando…';
  try { await fn(); } catch (e) { console.error(e); alert('No se pudo generar el archivo: ' + e.message); }
  btn.disabled = false; btn.textContent = t;
}

// ---------- Inicio ----------
async function sha256(s) { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s)); return [...new Uint8Array(b)].map(x => x.toString(16).padStart(2, '0')).join(''); }
async function boot() {
  CAT = await (await fetch('catalog.json')).json();
  for (const s of CAT.services) { byId[s.id] = s; for (const v of s.variants) { v.params = detectParams(v); vById[v.id] = v; } }
  load();
  for (const k of Object.keys(state.sel)) if (!vById[state.sel[k]]) delete state.sel[k];
  for (const k of Object.keys(state.datos)) {
    const el = $('#f-' + k); if (!el) continue;
    el.addEventListener('input', () => { state.datos[k] = el.value; renderDatos(); renderResumen(); save(); });
  }
  document.querySelectorAll('[data-base]').forEach(b => b.addEventListener('click', () => {
    if (Object.keys(state.sel).length && !confirm('Esto reemplaza la selección actual por los servicios de ' + b.dataset.base + '. ¿Continuar?')) return; base(b.dataset.base);
  }));
  $('#b-limpiar').addEventListener('click', () => { if (confirm('¿Quitar todos los servicios y cambios?')) { state.sel = {}; state.edits = {}; open.clear(); render(); } });
  $('#b-anexo').addEventListener('click', e => guard(e.target, async () => download(await makeAnexo(), 'Anexo_A_' + slug() + '.docx')));
  $('#b-prop').addEventListener('click', e => guard(e.target, async () => download(await makePropuesta(), 'Propuesta_' + slug() + '.docx')));
  $('#b-guardar').addEventListener('click', () => download(new Blob([JSON.stringify({ app: 'era-armador', v: 1, ...state }, null, 1)], { type: 'application/json' }), 'Configuracion_' + slug() + '.json'));
  $('#f-abrir').addEventListener('change', async e => {
    const f = e.target.files[0]; if (!f) return;
    try { const s = JSON.parse(await f.text()); if (s.app !== 'era-armador') throw new Error('archivo no reconocido'); state.datos = s.datos; state.sel = s.sel; state.edits = s.edits || {}; open.clear(); render(); }
    catch (err) { alert('No se pudo abrir: ' + err.message); }
    e.target.value = '';
  });
  render();
}
document.addEventListener('DOMContentLoaded', () => {
  const ok = () => { $('#login').hidden = true; $('#app').hidden = false; boot().catch(e => { console.error(e); alert('No se pudo cargar el catálogo. Abre la página desde GitHub Pages o un servidor (no como archivo local).'); }); };
  let pre = false; try { pre = sessionStorage.getItem('era-ok') === PASS_HASH; } catch (e) { }
  if (pre) return ok();
  $('#login-form').addEventListener('submit', async e => {
    e.preventDefault();
    if (await sha256($('#pw').value) === PASS_HASH) { try { sessionStorage.setItem('era-ok', PASS_HASH); } catch (er) { } ok(); }
    else { $('#pw-err').hidden = false; $('#pw').select(); }
  });
});

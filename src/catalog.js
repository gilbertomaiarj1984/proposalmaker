/* =====================================================================
   CATÁLOGO DE PERFIS E TAXAS (Excel) — um arquivo com UMA ABA POR CLIENTE.
   Cada aba pode ter o formato do próprio cliente (o padrão é
   Role | Level | Location | Cost/h | Rate net/h; abas com outros títulos também são lidas); as colunas são detectadas
   pelo cabeçalho. Alterações são aplicadas NA PLANILHA ORIGINAL (formatação,
   filtros e cores preservados): atualiza células, acrescenta linhas, remove linhas
   e cria abas de clientes novos (clonando a aba de outro cliente ou em branco).
   Sem DOM: roda também no Node (testes e geração do arquivo modelo).
===================================================================== */
const Catalog = (() => {
'use strict';
const E = Engine;
let seq = 0;
const newId = () => ++seq;   // identifica cada perfil/nível na memória (as linhas da planilha mudam ao gravar)

/* ---------------- leitura ---------------- */
const hnorm = v => E.norm(String(v ?? '').replace(/\s+/g, ' '));
const isRole = h => /^(roles?|perfis|perfil|funcao|cargo|resource|profile)\b/.test(h);
const isLevel = h => /^(experience level|level|nivel|senioridade|seniority|grade|band)\b/.test(h);
const isLoc = h => /^(location|local|localidade|country|pais)\b/.test(h);
const isCost = h => /^(cost|custo|costo)\b/.test(h);
const isRate = h => !isCost(h) && /(rate|without taxes|sem impostos|tarifa|price|preco|valor hora)/.test(h);
const CUR = /\b(BRL|USD|EUR|MXN|COP|ARS)\b/i;

function text(c){
  const v = c && c.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object'){
    if (v.richText) return v.richText.map(t => t.text).join('');
    if ('result' in v) return String(v.result ?? '');
    if (v.text) return String(v.text);
    if (v instanceof Date) return v.toISOString();
  }
  return String(v);
}
function numOf(c){
  const v = c && c.value;
  if (typeof v === 'number') return v || '';
  if (v && typeof v === 'object' && typeof v.result === 'number') return v.result || '';
  const t = text(c).replace(/R\$|US\$|€|\s/g, '');
  return t === '' ? '' : (E.num(t) || '');
}

// procura a linha de cabeçalho (até a linha 25) e as colunas Perfil/Nível/Local/Custo/Rate
function detect(ws){
  const lim = Math.min(25, ws.rowCount);
  for (let r = 1; r <= lim; r++){
    const cols = {}; let cur = '';
    ws.getRow(r).eachCell({includeEmpty: false}, (cell, n) => {
      const h = hnorm(text(cell));
      if (!h) return;
      if (isRole(h) && !cols.role) cols.role = n;
      else if (isLevel(h) && !cols.level) cols.level = n;
      else if (isLoc(h) && !cols.location) cols.location = n;
      else if (isCost(h) && !cols.cost) { cols.cost = n; const m = text(cell).match(CUR); if (m && !cur) cur = m[1].toUpperCase(); }
      else if (isRate(h) && !cols.rate) { cols.rate = n; const m = text(cell).match(CUR); if (m) cur = m[1].toUpperCase(); }
    });
    if (cols.role && (cols.rate || cols.cost)) return {hdrRow: r, cols, currency: cur};
  }
  return null;
}
function parseSheet(ws){
  const d = detect(ws); if (!d) return null;
  const {hdrRow, cols} = d, entries = [];
  for (let r = hdrRow + 1; r <= ws.rowCount; r++){
    const row = ws.getRow(r), role = text(row.getCell(cols.role)).trim();
    if (!role) continue;
    entries.push({
      role, level: cols.level ? text(row.getCell(cols.level)).trim() : '', location: cols.location ? text(row.getCell(cols.location)).trim() : '',
      cost: cols.cost ? numOf(row.getCell(cols.cost)) : '', rate: cols.rate ? numOf(row.getCell(cols.rate)) : '', row: r, id: newId()
    });
  }
  return {name: ws.name, sheet: ws.name, hdrRow, cols, currency: d.currency, entries};
}
// buffer -> {clients, skipped}; abas sem colunas reconhecíveis (ex.: Instructions) são ignoradas
async function parse(ExcelJS, buffer){
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const clients = [], skipped = [];
  wb.worksheets.forEach(ws => { if (ws.state === 'hidden' || ws.state === 'veryHidden') return; const c = parseSheet(ws); if (c) clients.push(c); else skipped.push(ws.name); });
  return {clients, skipped};
}

/* ---------------- escrita (aplica as mudanças na planilha original) ---------------- */
const BAD = /[\\/?*[\]:]/g;
function sheetName(name, existing){
  let base = String(name || 'Cliente').replace(BAD, ' ').replace(/\s+/g, ' ').trim().slice(0, 31) || 'Cliente';
  const taken = new Set(existing.map(n => n.toLowerCase()));
  let n = base, k = 2;
  while (taken.has(n.toLowerCase())){ const suf = ` ${k++}`; n = base.slice(0, 31 - suf.length) + suf; }
  return n;
}
const STD_COLS = {role: 1, level: 2, location: 3, cost: 4, rate: 5};
const HEAD_FILL = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFD9E5F3'}};
function standardSheet(wb, name){
  const ws = wb.addWorksheet(name, {views: [{state: 'frozen', ySplit: 1}]});
  ws.addRow(['Role', 'Level', 'Location', 'Cost/h', 'Rate net/h']);
  ws.getRow(1).eachCell(c => { c.font = {bold: true}; c.fill = HEAD_FILL; });
  [30, 10, 14, 12, 12].forEach((w, i) => ws.getColumn(i + 1).width = w);
  ws.getColumn(4).numFmt = '#,##0.00'; ws.getColumn(5).numFmt = '#,##0.00';
  return ws;
}
function cloneSheet(wb, base, name){
  const ws = wb.addWorksheet(name);
  // o modelo carrega nome e id da aba base: sobrescreve com os da nova aba antes de atribuir
  ws.model = Object.assign({}, base.model, {mergeCells: base.model.merges, name, id: ws.id, orderNo: ws.orderNo});
  return ws;
}
function setText(cell, v){ if (text(cell).trim() !== String(v ?? '').trim()) cell.value = v === '' ? null : v; }
function writeEntry(ws, cl, e, r){
  const row = ws.getRow(r), c = cl.cols;
  setText(row.getCell(c.role), e.role);
  if (c.level) setText(row.getCell(c.level), e.level);
  if (c.location) setText(row.getCell(c.location), e.location);
  if (c.cost) row.getCell(c.cost).value = e.cost === '' || e.cost === undefined ? null : e.cost;
  if (c.rate) row.getCell(c.rate).value = e.rate === '' || e.rate === undefined ? null : e.rate;
  row.commit && row.commit();
}
// state = {clients:[{name, sheet, isNew?, base?, hdrRow, cols, entries:[{id, ..., row?, deleted?}]}]}
// Devolve {buffer, rows:{id: linha}, removed:[ids], sheets:{cliente: nomeDaAba}} — o chamador atualiza a memória com isso.
async function serialize(ExcelJS, buffer, state){
  const st = JSON.parse(JSON.stringify(state));   // trabalha numa cópia: edições feitas durante a gravação não se perdem
  const wb = new ExcelJS.Workbook();
  if (buffer) await wb.xlsx.load(buffer);
  const rows = {}, removed = [], sheets = {};
  for (const cl of st.clients){
    let ws = wb.getWorksheet(cl.sheet);
    if (!ws){
      const nm = sheetName(cl.sheet, wb.worksheets.map(w => w.name));
      const baseWs = cl.base ? wb.getWorksheet(cl.base) : null;
      if (baseWs){
        // clona cabeçalho/formatação e refaz as linhas a partir da memória (a aba base pode ter mudado nesta mesma gravação)
        const sample = {};
        Object.values(cl.cols).forEach(cn => { const c = baseWs.getRow(cl.hdrRow + 1).getCell(cn); sample[cn] = c && c.style ? {...c.style} : null; });
        ws = cloneSheet(wb, baseWs, nm);
        ws.spliceRows(cl.hdrRow + 1, Math.max(0, ws.rowCount - cl.hdrRow));
        cl.entries.forEach(e => { delete e.row; });
        cl.styleSample = sample;
      } else ws = standardSheet(wb, nm);
      cl.sheet = nm;
    }
    sheets[state.clients[st.clients.indexOf(cl)].name] = cl.sheet;
    const lastRow = cl.entries.reduce((m, e) => Math.max(m, e.row || 0), cl.hdrRow);
    let next = lastRow + 1;
    for (const e of cl.entries){
      if (e.deleted){ if (e.id) removed.push(e.id); continue; }
      if (e.row) writeEntry(ws, cl, e, e.row);
      else {
        const r = next++;
        const prev = ws.getRow(Math.max(cl.hdrRow + 1, r - 1)); // herda o estilo da linha de cima (ou o da aba base)
        Object.values(cl.cols).forEach(cn => { const st0 = cl.styleSample ? cl.styleSample[cn] : (prev.getCell(cn) || {}).style; if (st0) ws.getRow(r).getCell(cn).style = {...st0}; });
        writeEntry(ws, cl, e, r);
        e.row = r;
      }
      rows[e.id] = e.row;
    }
    const gone = cl.entries.filter(e => e.deleted && e.row).map(e => e.row).sort((a, b) => b - a);
    gone.forEach(r => { ws.spliceRows(r, 1); Object.keys(rows).forEach(id => { if (rows[id] > r && cl.entries.some(e => String(e.id) === id)) rows[id]--; }); });
  }
  return {buffer: await wb.xlsx.writeBuffer(), rows, removed, sheets};
}

/* ---------------- arquivo modelo (Roche + JnJ) ---------------- */
const JNJ_ROWS = [
  ['Developer: SAP BASIS', 'Senior', 142.91], ['Developer: SAP BASIS', 'Expert', 178.08],
  ['Functional Analyst: SAP functional analyst', 'Junior', 109.93], ['Functional Analyst: SAP functional analyst', 'Senior', 167.09], ['Functional Analyst: SAP functional analyst', 'Expert', 198.97],
  ['SAP Architect', 'Junior'], ['SAP Architect', 'Senior'], ['SAP Architect', 'Expert'],
  ['Change Management/ Training Lead', 'Expert'], ['Change Management/ Training Lead', 'Junior'], ['Change Management/ Training Lead', 'Senior'],
  ['Enterprise Architect', 'Expert'], ['Enterprise Architect', 'Junior'], ['Enterprise Architect', 'Senior'],
  ['Microsoft Dynamics Developer', 'Expert'], ['Microsoft Dynamics Developer', 'Junior'], ['Microsoft Dynamics Developer', 'Senior'],
  ['Platform Engineer', 'Expert'], ['Platform Engineer', 'Junior'], ['Platform Engineer', 'Senior'],
  ['Product Lead Engineer', 'Expert'], ['Product Lead Engineer', 'Junior']
];
function seed(ExcelJS){
  const wb = new ExcelJS.Workbook();
  wb.creator = 'ProposalMaker'; wb.created = new Date();
  // Roche: formato com Local, Custo/h e Rate/h (valores ilustrativos — substituir pelos oficiais)
  const ro = standardSheet(wb, 'Roche');
  Workbook.exampleData().rates.forEach(r => ro.addRow([r[0], r[1], r[2], r[3], r[4]]));
  ro.getColumn(6).width = 46; ro.getCell('F1').value = 'Notes'; ro.getCell('F1').font = {bold: true}; ro.getCell('F1').fill = HEAD_FILL;
  ro.getCell('F2').value = 'Valores ilustrativos — substitua pelos oficiais da Roche';
  ro.autoFilter = 'A1:F1';
  // JnJ: mesmo layout padrão de todos os clientes (Role | Level | Location | Cost/h | Rate net/h); valores em BRL sem impostos
  const jn = standardSheet(wb, 'JnJ');
  JNJ_ROWS.forEach(r => jn.addRow([r[0], r[1], null, null, r[2] === undefined ? null : r[2]]));
  jn.autoFilter = 'A1:E1';
  const ins = wb.addWorksheet('Instructions');
  ins.getColumn(1).width = 110;
  ['Catálogo de perfis e taxas — uma aba por cliente',
   'Layout padrão de todos os clientes: Role | Level | Location | Cost/h | Rate net/h. Os valores da JnJ são uma amostra (Rate net/h em BRL, sem impostos): substitua pelos oficiais.',
   'Cada aba nova (cliente) é criada pelo app a partir de outra aba ou em branco; o app lê as colunas pelo cabeçalho.',
   'Perfis sem taxa ficam em branco e aparecem como "sem taxa" nas propostas.'].forEach((t, i) => { const r = ins.addRow([t]); if (!i) r.font = {bold: true, size: 13}; });
  return wb;
}

return { parse, serialize, seed, sheetName, detect, newId };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = Catalog;

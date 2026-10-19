/* =====================================================================
   APP — interface (DOM). Toda a regra de negócio está em engine.js / doc.js.
===================================================================== */
(() => {
'use strict';
const E = Engine;
const $ = id => document.getElementById(id);
const esc = E.esc;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/* ---------------- estado ---------------- */
const S = {
  rows: [], nPeriods: 0, unit: 'SP', unitAuto: 'SP', rates: {}, fileName: '',
  taxes: [{name: 'PIS/COFINS/ISS', pct: 5.65}], edits: {},
  result: null, handle: null, autoFile: false, dirty: false
};
const LIB = {dir: null, files: [], cache: new Map()};
const charts = {};

/* ---------------- utilidades de interface ---------------- */
const nfPT = (v, d = 0) => new Intl.NumberFormat('pt-BR', {minimumFractionDigits: d, maximumFractionDigits: d}).format(v);
const curPT = (v, cur, d = 0) => { try { return new Intl.NumberFormat('pt-BR', {style: 'currency', currency: cur, minimumFractionDigits: d, maximumFractionDigits: d}).format(v); } catch(e){ return cur + ' ' + nfPT(v, d); } };
const pad = n => String(n).padStart(2, '0');
const hhmm = (d = new Date()) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const slug = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^\w]+/g, '_').replace(/^_+|_+$/g, '');
const fileBase = () => `Proposta_${slug($('p_client').value) || 'Cliente'}_${slug($('p_project').value) || 'Projeto'}`;
function msg(el, type, html){ const n = $(el); if (n) n.innerHTML = html ? `<div class="msg ${type}">${html}</div>` : ''; }

function toast(text, type = 'info', ms = 6000, actions = []){
  const t = document.createElement('div');
  t.className = 'toast ' + (type === 'info' ? '' : type);
  const span = document.createElement('span'); span.textContent = text; t.appendChild(span);
  actions.forEach(a => { const b = document.createElement('button'); b.className = 'btn light sm'; b.textContent = a.label; b.onclick = () => { try { a.fn(); } finally { t.remove(); } }; t.appendChild(b); });
  $('toasts').appendChild(t);
  if (ms) setTimeout(() => t.remove(), ms);
  return t;
}
// nada pode travar a aplicação: erros inesperados viram aviso
window.addEventListener('error', e => { console.error(e.error || e.message); toast('Erro inesperado: ' + (e.message || 'veja o console'), 'err', 8000); });
window.addEventListener('unhandledrejection', e => { console.error(e.reason); toast('Falha: ' + ((e.reason && e.reason.message) || e.reason || 'erro desconhecido'), 'err', 8000); });

/* ---------------- bibliotecas embutidas (carregadas sob demanda) ---------------- */
const LIBS = {xlsx: 'XLSX', chart: 'Chart', exceljs: 'ExcelJS', docx: 'docx', pptx: 'PptxGenJS'};
const libPromises = {};
function loadLib(name){
  const g = LIBS[name];
  if (window[g]) return Promise.resolve(window[g]);
  if (libPromises[name]) return libPromises[name];
  libPromises[name] = new Promise((res, rej) => {
    const holder = document.querySelector(`script[data-lib="${name}"]`);
    if (!holder) return rej(new Error(`biblioteca "${name}" ausente neste arquivo`));
    try { const s = document.createElement('script'); s.textContent = holder.textContent; document.head.appendChild(s); s.remove(); }
    catch(e){ return rej(e); }
    window[g] ? res(window[g]) : rej(new Error(`biblioteca "${name}" não inicializou`));
  });
  libPromises[name].catch(() => { delete libPromises[name]; });
  return libPromises[name];
}

/* ---------------- parâmetros (DOM <-> estado) ---------------- */
const PARAM_IDS = Workbook.PARAM_DEFS.map(d => d[0]);
const BOOL = Workbook.BOOLS;
function raw(){
  const o = {};
  PARAM_IDS.forEach(id => { const el = $(id); if (el) o[id] = BOOL.has(id) ? el.checked : el.value; });
  return o;
}
function setRaw(o){
  PARAM_IDS.forEach(id => {
    const el = $(id); if (!el || o[id] === undefined) return;
    if (BOOL.has(id)) el.checked = !!o[id]; else el.value = o[id];
    if (el.tagName === 'SELECT' && el.selectedIndex < 0) el.selectedIndex = 0;
  });
  syncStartUI();
}
const P = () => E.parseParams(raw(), {taxes: S.taxes});
const model = () => ({rows: S.rows, nPeriods: S.nPeriods, unit: S.unit, rates: S.rates});
const snapshot = () => ({raw: raw(), taxes: JSON.parse(JSON.stringify(S.taxes)), rows: S.rows, nPeriods: S.nPeriods, unit: S.unit, unitAuto: S.unitAuto, rates: S.rates, edits: S.edits, fileName: S.fileName});

function syncStartUI(){
  const m = $('p_startmode').value;
  $('lbl_start').style.display = m === 'date' ? '' : 'none';
  $('lbl_startm').style.display = m === 'month' ? '' : 'none';
}

/* ---------------- abas ---------------- */
const TABS = ['tl', 'par', 'ana', 'doc'];
function goTab(t){
  if ((t === 'ana' || t === 'doc') && !S.result) recalc(false);
  TABS.forEach(x => {
    const b = $('tabbtn-' + x), on = x === t;
    b.setAttribute('aria-selected', on ? 'true' : 'false'); b.tabIndex = on ? 0 : -1;
    $('tab-' + x).classList.toggle('active', on);
  });
  if (t === 'ana') drawCharts();
  window.scrollTo(0, 0);
}
function enableTabs(n){ TABS.forEach((x, i) => { $('tabbtn-' + x).disabled = i > n; }); }
$('tabbtn-tl').closest('nav').addEventListener('keydown', e => {
  if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
  const en = TABS.filter(x => !$('tabbtn-' + x).disabled), cur = TABS.find(x => $('tabbtn-' + x) === document.activeElement);
  if (!cur) return;
  const next = en[(en.indexOf(cur) + (e.key === 'ArrowRight' ? 1 : en.length - 1)) % en.length];
  goTab(next); $('tabbtn-' + next).focus();
});
document.querySelectorAll('nav.tabs button').forEach(b => b.addEventListener('click', () => !b.disabled && goTab(b.dataset.tab)));

/* ---------------- conexão (só afeta os extras online) ---------------- */
function renderNet(){
  const pill = $('netpill'), bar = $('netbar'), st = Net.status;
  if (!st.online){
    pill.className = 'pill off'; pill.textContent = 'Sem internet';
    bar.textContent = '⚠ Sem conexão com a internet: câmbio automático e conferência online de feriados estão indisponíveis. Todo o resto funciona normalmente (cálculo, proposta, Excel, Word, PDF).';
    bar.classList.add('show');
  } else if (st.lastError){
    pill.className = 'pill off'; pill.textContent = 'Conexão limitada';
    bar.textContent = `⚠ Acesso à internet limitado (${st.lastError}). Os recursos online (câmbio, feriados) podem falhar; digite os valores manualmente. O restante do app não é afetado.`;
    bar.classList.add('show');
  } else {
    pill.className = st.lastOkAt ? 'pill on' : 'pill chk'; pill.textContent = st.lastOkAt ? 'Online' : 'Conexão: verificar';
    bar.classList.remove('show');
  }
}
Net.onChange(renderNet);
async function netCheck(){
  const pill = $('netpill'); pill.textContent = 'Verificando…';
  const r = await Net.probe();
  const lab = {fx: 'câmbio', holidays: 'feriados'};
  const bad = Object.keys(r).filter(k => !r[k].ok);
  if (!bad.length){ Net.status.lastError = null; Net.status.lastOkAt = new Date(); renderNet(); toast('Conexão OK: câmbio e feriados online acessíveis.', 'ok'); return; }
  Net.status.lastError = bad.map(k => `${lab[k]}: ${r[k].msg}`).join('; ');
  renderNet();
  toast('Recursos online indisponíveis: ' + bad.map(k => lab[k]).join(', ') + '. O app continua funcionando offline.', 'warn', 9000);
}

/* ---------------- persistência: rascunho, arquivo .xlsx e biblioteca ---------------- */
let saveTimer = null;
function setSaveStatus(t){ $('savebar').textContent = t; }
function markDirty(){
  S.dirty = true; setSaveStatus('Alterações não salvas');
  clearTimeout(saveTimer); saveTimer = setTimeout(autosave, 1500);
}
async function autosave(){
  if (!S.rows.length) return;
  let text;
  const ok = await Store.kv.set('draft', {snap: snapshot(), at: Date.now()});
  text = ok ? `Rascunho salvo ${hhmm()}` : 'Rascunho local indisponível neste navegador';
  if (S.handle && S.autoFile){
    try {
      if (await Store.fs.granted(S.handle, true)){ await writeFile(S.handle); S.dirty = false; text = `Salvo em ${S.handle.name} · ${hhmm()}`; }
      else text += ' · clique em 💾 Salvar para reconectar o arquivo';
    } catch(e){ text += ' · falha ao gravar o arquivo: ' + e.message; }
  }
  setSaveStatus(text);
}
async function buildBlob(){
  const R = S.result || recalc(false);
  if (!R) throw new Error('Carregue um timeline antes de salvar.');
  const Ex = await loadLib('exceljs');
  const wb = await Workbook.buildProposal(Ex, snapshot(), R);
  return new Blob([await wb.xlsx.writeBuffer()], {type: XLSX_MIME});
}
async function writeFile(handle){ await Store.fs.write(handle, await buildBlob()); }
const aborted = e => e && (e.name === 'AbortError' || /abort/i.test(e.message || ''));

async function actSave(){
  if (!S.rows.length) return toast('Carregue um timeline primeiro.', 'warn');
  try {
    if (S.handle && Store.fs.supported() && await Store.fs.perm(S.handle, true)){
      await writeFile(S.handle); S.dirty = false; S.autoFile = true;
      setSaveStatus(`Salvo em ${S.handle.name} · ${hhmm()}`); toast('Proposta salva em ' + S.handle.name, 'ok'); return;
    }
    return actSaveAs();
  } catch(e){ if (!aborted(e)) toast('Não foi possível salvar: ' + e.message, 'err', 9000); }
}
async function actSaveAs(){
  if (!S.rows.length) return toast('Carregue um timeline primeiro.', 'warn');
  try {
    const name = fileBase() + '.xlsx';
    if (Store.fs.supported()){
      const h = await Store.fs.pickSave(name);
      await writeFile(h); S.handle = h; S.autoFile = true; S.dirty = false;
      setSaveStatus(`Salvo em ${h.name} · ${hhmm()}`); toast('Proposta salva. As próximas alterações serão gravadas neste arquivo automaticamente.', 'ok', 7000);
    } else {
      Store.download(await buildBlob(), name); S.dirty = false;
      setSaveStatus(`Arquivo baixado · ${hhmm()}`);
      toast('Seu navegador não permite gravar direto no arquivo: o .xlsx foi baixado. Use “Abrir proposta” para continuar depois.', 'info', 9000);
    }
  } catch(e){ if (!aborted(e)) toast('Não foi possível salvar: ' + e.message, 'err', 9000); }
}
async function actOpen(){
  try {
    if (!Store.fs.supported()) return $('tlFile').click();
    const h = await Store.fs.pickOpen();
    await loadBuffer(await Store.fs.read(h), h.name, h);
  } catch(e){ if (!aborted(e)) toast('Não foi possível abrir: ' + e.message, 'err', 9000); }
}

/* ---------------- carregamento de arquivos ---------------- */
async function handleFile(file){
  if (!file) return;
  try { await loadBuffer(await file.arrayBuffer(), file.name, null); }
  catch(e){ msg('tlMsg', 'err', '❌ ' + esc(e.message)); }
}
async function loadBuffer(buf, name, handle){
  const X = await loadLib('xlsx');
  let wb;
  try { wb = X.read(new Uint8Array(buf), {type: 'array', cellDates: true}); }
  catch(e){ throw new Error('arquivo ilegível (' + e.message + ')'); }
  const aoa = n => X.utils.sheet_to_json(wb.Sheets[n], {header: 1, defval: '', raw: true, blankrows: false});

  if (Workbook.isProposalWorkbook(X, wb)){
    const st = Workbook.read(X, wb);
    st.fileName = name; applyState(st);
    S.handle = handle || null; S.autoFile = !!handle; S.dirty = false;
    msg('tlMsg', 'ok', `✅ Proposta <b>${esc(name)}</b> aberta: ${S.rows.length} linhas, ${S.nPeriods} ${S.unit === 'SP' ? 'sprints' : 'semanas'}, ${Object.keys(S.rates).length} perfis, ${Object.keys(S.edits).length} texto(s) editado(s).` + (st.warnings.length ? '<br>⚠ ' + st.warnings.map(esc).join('<br>⚠ ') : ''));
    setSaveStatus(handle ? `Arquivo: ${name} (alterações são gravadas nele)` : `Aberta de ${name} (use Salvar para gravar)`);
    return;
  }

  let tlName = wb.SheetNames.find(n => /timeline|cronograma|gantt|plan/i.test(n)) || null;
  let parsed = null; const errors = [];
  const candidates = tlName ? [tlName, ...wb.SheetNames.filter(n => n !== tlName)] : wb.SheetNames;
  for (const n of candidates){
    if (/rate|instruc|alt format/i.test(n)) continue;
    try { parsed = E.parseTimeline(aoa(n)); parsed.sheet = n; break; }
    catch(e){ errors.push(`${esc(n)}: ${esc(e.message)}`); }
  }
  if (!parsed) throw new Error('Nenhuma aba reconhecida como timeline.<br>' + errors.join('<br>'));
  S.rows = parsed.rows; S.nPeriods = parsed.nPeriods; S.unitAuto = parsed.unit;
  const sel = $('p_unit').value; S.unit = sel === 'auto' ? parsed.unit : sel;
  S.fileName = name; S.handle = null; S.autoFile = false; S.edits = {};
  E.ensureRates(S.rates, S.rows);
  const rcName = wb.SheetNames.find(n => /rate/i.test(n));
  let rcInfo = '';
  if (rcName){
    try { const n = E.applyRateTable(S.rates, E.parseRateRows(aoa(rcName))); rcInfo = ` Rate card importado da aba "${esc(rcName)}" (${n} de ${Object.keys(S.rates).length} perfis).`; }
    catch(e){ rcInfo = ` ⚠ Rate card não importado: ${esc(e.message)}.`; }
  }
  afterTimelineLoaded(`Arquivo <b>${esc(name)}</b> (aba "${esc(parsed.sheet)}"): ${S.rows.length} linhas, ${S.nPeriods} ${S.unit === 'SP' ? 'sprints' : 'semanas'}.${rcInfo}` + (parsed.warnings.length ? '<br>⚠ ' + parsed.warnings.join('<br>⚠ ') : ''));
}
function applyState(st){
  S.rows = st.rows; S.nPeriods = st.nPeriods; S.unit = st.unit; S.unitAuto = st.unitAuto || st.unit; S.rates = st.rates || {};
  S.taxes = st.taxes && st.taxes.length ? st.taxes : [{name: 'Impostos', pct: 5.65}];
  S.edits = st.edits || {}; S.fileName = st.fileName || '';
  E.ensureRates(S.rates, S.rows);
  setRaw(st.raw || {});
  $('p_unit').value = S.unit === S.unitAuto ? 'auto' : S.unit;
  renderTaxes(); S.result = null; renderGantt(); renderRateTable(); enableTabs(1);
  $('tlView').style.display = 'block'; refreshFxCachedBtn();
}
function afterTimelineLoaded(info){
  msg('tlMsg', 'ok', '✅ ' + info);
  if (!$('p_project').value && S.fileName) $('p_project').value = S.fileName.replace(/\.(xlsx|xlsm|xls|csv)$/i, '').replace(/[_-]+/g, ' ');
  S.result = null; renderGantt(); renderRateTable(); enableTabs(1);
  $('tlView').style.display = 'block'; markDirty();
}
function loadExample(){
  const ex = Workbook.exampleData();
  S.rows = ex.rows.map(r => ({phase: r[0], activity: r[1], role: r[2], level: r[3], location: r[4], optional: r[5] === 'Y', alloc: r[6]}));
  S.nPeriods = ex.N; S.unit = S.unitAuto = 'SP'; S.fileName = ''; S.rates = {}; S.edits = {}; S.handle = null; S.autoFile = false;
  E.ensureRates(S.rates, S.rows);
  E.applyRateTable(S.rates, ex.rates.map(r => ({role: r[0], level: r[1], location: r[2], cost: r[3], rate: r[4]})));
  if (!$('p_client').value) $('p_client').value = 'Client';
  if (!$('p_project').value) $('p_project').value = 'Data & Analytics Modernization';
  $('p_unit').value = 'auto';
  afterTimelineLoaded('Exemplo carregado (12 sprints, rates ilustrativos — substitua pelo rate card oficial).');
}
async function importRateCard(file){
  if (!file) return;
  try {
    const X = await loadLib('xlsx');
    const wb = X.read(new Uint8Array(await file.arrayBuffer()), {type: 'array'});
    const n = wb.SheetNames.find(x => /rate/i.test(x)) || wb.SheetNames[0];
    const c = E.applyRateTable(S.rates, E.parseRateRows(X.utils.sheet_to_json(wb.Sheets[n], {header: 1, defval: '', raw: true, blankrows: false})));
    renderRateTable(); S.result = null; markDirty();
    msg('rcMsg', c ? 'ok' : 'warn', c ? `✅ ${c} perfis do timeline atualizados a partir de "${esc(file.name)}".` : `Nenhum perfil do rate card "${esc(file.name)}" corresponde aos perfis do timeline.`);
  } catch(e){ msg('rcMsg', 'err', '❌ ' + esc(e.message)); }
}

/* ---------------- render: timeline importado ---------------- */
const PALETTE = ['#1B4F8C','#13A89E','#6C4AB6','#E07A1F','#2E86C1','#C0392B','#7D8C2E','#AD5A8C','#4A6B82','#D4A017'];
function phaseNames(){ const out = []; S.rows.forEach(r => { if (!out.includes(r.phase)) out.push(r.phase); }); return out; }
function renderGantt(){
  const U = S.unit, N = S.nPeriods, phases = phaseNames();
  let h = `<table class="gantt"><thead><tr><th class="stk" style="left:0">Fase / Perfil</th><th>Nível</th><th>Local</th>`;
  for (let p = 1; p <= N; p++) h += `<th class="c">${U}${p}</th>`;
  h += `<th class="c">Total</th></tr></thead><tbody>`;
  phases.forEach((ph, pi) => {
    const color = PALETTE[pi % PALETTE.length], rows = S.rows.filter(r => r.phase === ph);
    const span = new Array(N).fill(0); rows.forEach(r => r.alloc.forEach((v, i) => span[i] += v));
    h += `<tr class="ph"><td class="stk" style="left:0;background:#F3F6FA"><span style="display:inline-block;width:10px;height:10px;background:${color};border-radius:2px;margin-right:6px"></span>${esc(ph)}${rows.every(r => r.optional) ? ' <i style="color:var(--warn)">(opcional)</i>' : ''}</td><td></td><td></td>`;
    span.forEach(v => h += `<td class="c" style="${v ? `background:${color};color:#fff` : ''}">${v ? nfPT(v, 1) : ''}</td>`);
    h += `<td class="c">${nfPT(span.reduce((a, b) => a + b, 0), 1)}</td></tr>`;
    rows.forEach(r => {
      const tot = r.alloc.reduce((a, b) => a + b, 0);
      h += `<tr style="${r.optional ? 'color:var(--warn);font-style:italic' : ''}"><td class="stk" style="left:0;padding-left:22px">${esc(r.role)}${r.activity ? ` <span class="hint">· ${esc(r.activity)}</span>` : ''}</td><td>${esc(r.level)}</td><td>${esc(r.location)}</td>`;
      r.alloc.forEach(v => { const a = Math.min(1, v / 1.5); h += `<td class="c" style="${v ? `background:rgba(27,79,140,${0.12 + 0.5 * a})` : ''}">${v ? nfPT(v, v % 1 ? 1 : 0) : ''}</td>`; });
      h += `<td class="c">${nfPT(tot, 1)}</td></tr>`;
    });
  });
  $('gantt').innerHTML = h + '</tbody></table>';
  $('tlTitle').textContent = `Timeline importado — ${S.rows.length} linhas · ${phases.length} fases · ${N} ${U === 'SP' ? 'sprints' : 'semanas'}`;
}

/* ---------------- render: rate card ---------------- */
function renderRateTable(){
  const p = P(), keys = Object.keys(S.rates).sort();
  const hp = (S.unit === 'W' ? 1 : p.pw) * p.dpw * p.hpd;
  const hrsBy = {}; S.rows.forEach(r => { const k = E.rkey(r); hrsBy[k] = (hrsBy[k] || 0) + r.alloc.reduce((a, b) => a + b, 0) * hp; });
  let h = `<thead><tr><th>Perfil</th><th>Nível</th><th>Local</th><th class="n">Horas (timeline)</th><th class="n">Custo/h</th><th class="n">Rate/h</th><th class="n">GM%</th></tr></thead><tbody>`;
  keys.forEach(k => {
    const r = S.rates[k], gm = r.rate && r.cost ? (1 - r.cost / r.rate) * 100 : null;
    const miss = v => (v === '' || v === 0 || v === null) ? 'style="border-color:var(--err);background:#FFF5F5"' : '';
    h += `<tr><td>${esc(r.role)}</td><td>${esc(r.level)}</td><td>${esc(r.location)}</td><td class="n">${nfPT(hrsBy[k] || 0)}</td>
      <td class="n"><input type="number" step="0.01" aria-label="Custo/h de ${esc(r.role)}" value="${esc(r.cost)}" data-k="${esc(k)}" data-f="cost" ${p.mode === 'gm' ? miss(r.cost) : ''}></td>
      <td class="n"><input type="number" step="0.01" aria-label="Rate/h de ${esc(r.role)}" value="${esc(r.rate)}" data-k="${esc(k)}" data-f="rate" ${p.mode === 'rate' ? miss(r.rate) : ''}></td>
      <td class="n" style="${gm !== null && gm < 0 ? 'color:var(--err)' : ''}">${gm === null ? '–' : nfPT(gm, 1) + '%'}</td></tr>`;
  });
  $('rcTable').innerHTML = h + '</tbody>';
}
$('rcTable').addEventListener('change', e => {
  const i = e.target.closest('input[data-k]'); if (!i) return;
  S.rates[i.dataset.k][i.dataset.f] = i.value === '' ? '' : E.num(i.value);
  S.result = null; markDirty(); renderRateTable();
});

/* ---------------- impostos ---------------- */
function renderTaxes(){
  $('taxList').innerHTML = S.taxes.map((t, i) => `<div class="taxrow">
    <input data-i="${i}" data-f="name" value="${esc(t.name)}" aria-label="Nome do imposto ${i + 1}">
    <input data-i="${i}" data-f="pct" type="number" step="0.01" min="0" value="${esc(t.pct)}" aria-label="Percentual do imposto ${i + 1}">
    <button class="btn sec sm" data-act="tax-del" data-i="${i}" aria-label="Remover imposto ${i + 1}" ${S.taxes.length < 2 ? 'disabled' : ''}>Remover</button></div>`).join('');
  const tot = S.taxes.reduce((a, t) => a + E.num(t.pct), 0);
  $('taxTotal').textContent = `Total: ${nfPT(tot, 2)}%`;
}
$('taxList').addEventListener('change', e => {
  const i = e.target.closest('input[data-i]'); if (!i) return;
  const t = S.taxes[+i.dataset.i]; if (!t) return;
  if (i.dataset.f === 'name') t.name = i.value; else t.pct = Math.max(0, E.num(i.value));
  S.result = null; renderTaxes(); markDirty();
});

/* ---------------- cálculo ---------------- */
function recalc(go){
  if (!S.rows.length) return null;
  try {
    const R = E.compute(model(), P(), {today: new Date()});
    S.result = R; enableTabs(3);
    renderAnalysis(); renderDoc();
    msg('rcMsg', R.missing.length ? 'warn' : '', R.missing.length ? `⚠ Sem ${R.p.mode === 'rate' ? 'Rate/h' : 'Custo/h'} para: ${esc(R.missing.join(', '))}. Esses perfis entram com valor zero.` : '');
    const inPeriod = R.holidays.filter(h => h.date >= E.iso(R.T.start) && h.date <= E.iso(R.T.end));
    $('holInfo').textContent = R.relative ? 'Sem data de início: feriados não se aplicam.' : (R.p.useHol ? `${inPeriod.length} feriado(s) dentro do período do projeto.` : 'Feriados desativados.');
    if (go) goTab('ana');
    return R;
  } catch(e){ console.error(e); toast('Erro ao calcular a proposta: ' + e.message, 'err', 9000); return null; }
}

/* ---------------- render: análise ---------------- */
function issueText(i, R){
  const d = i.data, cur = R.p.cur;
  switch (i.code){
    case 'missing_rates': return `Sem ${d.mode === 'rate' ? 'Rate/h' : 'Custo/h'} para: <b>${esc(d.list.join(', '))}</b>. Esses perfis entram com valor zero — preencha o rate card antes de enviar a proposta.`;
    case 'gm_negative': return `GM% negativo (${nfPT(d.gm * 100, 1)}%): o preço está abaixo do custo.`;
    case 'gm_below_min': return `GM% de ${nfPT(d.gm * 100, 1)}% está abaixo do mínimo configurado (${nfPT(d.min * 100, 1)}%).`;
    case 'no_cost': return 'Sem Custo/h informado: o GM% não pode ser calculado.';
    case 'price_below_cost': return `Preço abaixo do custo em: ${esc(d.list.join(', '))}.`;
    case 'start_adjusted': return `${esc(d.from)} não é dia útil (fim de semana ou feriado): o projeto começa em <b>${esc(d.to)}</b>.`;
    case 'relative_timeline': return 'Início sem data: a proposta mostra apenas períodos relativos (SP1, SP2…), sem datas de calendário.';
    case 'fx_missing': return `A moeda da proposta (${esc(d.to)}) difere da do rate card (${esc(d.from)}) e o câmbio está em 1. Informe ou busque a cotação.`;
    case 'optional_excluded': return `Há itens opcionais fora do total (${curPT(d.gross, cur)} com impostos).`;
    case 'no_active_scope': return 'Nenhum item no escopo (todos opcionais e excluídos, ou sem alocação).';
    default: return esc(i.code);
  }
}
function renderAnalysis(){
  const R = S.result, T = R.T, p = R.p, U = R.unit, cur = p.cur, C = (v, d = 0) => curPT(v, cur, d);
  const icon = {err: '⛔', warn: '⚠', info: 'ℹ', ok: '✅'};
  const items = R.issues.length ? R.issues : [];
  const hasBad = R.issues.some(i => i.level === 'err' || i.level === 'warn');
  $('checklist').innerHTML = (items.map(i => `<div class="it ${i.level}"><span aria-hidden="true">${icon[i.level]}</span><span>${issueText(i, R)}</span></div>`).join('')) +
    (hasBad ? '' : `<div class="it ok"><span aria-hidden="true">✅</span><span>Nenhum alerta crítico: rate card completo, margem dentro do esperado.</span></div>`);

  const k = (l, v, s = '') => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;
  $('kpis').innerHTML =
    k('Preço total (c/ impostos)', C(T.gross), `Líquido: ${C(T.net)}`) +
    k('Duração', `${T.nPer} ${U === 'SP' ? 'sprints' : 'semanas'}`, R.relative ? `${T.weeks} semanas · sem data de início` : `${T.weeks} semanas · ${E.iso(T.start)} → ${E.iso(T.end)}`) +
    k('Esforço', `${nfPT(T.H)} h`, `${nfPT(T.H / (p.hpd * p.dpw * 52 / 12 || 1), 1)} FTE-mês`) +
    k('Time', `pico ${nfPT(T.peak, 1)} FTE`, `média ${nfPT(T.avgFte, 1)} FTE`) +
    k('Rate médio (blended)', C(T.blended, 2) + '/h', p.mode === 'rate' ? 'rate card' : 'custo + GM') +
    k('GM% estimado', T.cost ? nfPT(T.gm * 100, 1) + '%' : '–', T.cost ? `Custo: ${C(T.cost)}` : 'preencha Custo/h') +
    (T.netOpt ? k('Opcionais', C(T.grossOpt), p.inclOpt ? 'incluídos no total' : 'fora do total') : '');

  let h = `<table class="t"><thead><tr><th>Fase</th><th>Período</th><th>Datas</th><th class="n">Horas</th><th class="n">Líquido</th><th class="n">Com impostos</th><th class="n">% total</th></tr></thead><tbody>`;
  R.phases.forEach(ph => {
    h += `<tr class="${ph.optional ? 'opt' : ''}"><td><span style="display:inline-block;width:9px;height:9px;background:${ph.color};border-radius:2px;margin-right:6px"></span>${esc(ph.name)}${ph.optional ? ' (opcional)' : ''}</td><td>${ph.first === ph.last ? U + (ph.first + 1) : U + (ph.first + 1) + '–' + U + (ph.last + 1)}</td><td>${R.relative ? '–' : E.iso(ph.start) + ' → ' + E.iso(ph.end)}</td><td class="n">${nfPT(ph.H)}</td><td class="n">${C(ph.net)}</td><td class="n">${C(ph.gross)}</td><td class="n">${(!ph.optional || p.inclOpt) && T.net ? nfPT(ph.net / T.net * 100, 1) + '%' : '–'}</td></tr>`;
  });
  h += `<tr class="tot"><td colspan="3">Total ${p.inclOpt ? '(com opcionais)' : '(escopo base)'}</td><td class="n">${nfPT(T.H)}</td><td class="n">${C(T.net)}</td><td class="n">${C(T.gross)}</td><td class="n">100%</td></tr></tbody></table>`;
  $('tbPhase').innerHTML = h;

  h = `<table class="t"><thead><tr><th>Perfil</th><th>Nível</th><th>Local</th><th class="n">FTE médio</th><th class="n">Pico</th><th class="n">Horas</th><th class="n">Rate/h aplicado</th><th class="n">Líquido</th><th class="n">GM%</th></tr></thead><tbody>`;
  R.roles.forEach(r => {
    h += `<tr class="${r.optional ? 'opt' : ''}"><td>${esc(r.role)}${r.optional ? ' (opc.)' : ''}</td><td>${esc(r.level)}</td><td>${esc(r.location)}</td><td class="n">${nfPT(r.avgFte, 2)}</td><td class="n">${nfPT(r.peak, 1)}</td><td class="n">${nfPT(r.H)}</td><td class="n">${C(r.rateH, 2)}</td><td class="n">${C(r.net)}</td><td class="n">${r.cost && r.net ? nfPT((1 - r.cost / r.net) * 100, 1) + '%' : '–'}</td></tr>`;
  });
  $('tbRole').innerHTML = h + '</tbody></table>';

  h = `<table class="t"><thead><tr><th>Mês</th><th class="n">Horas</th><th class="n">Líquido</th><th class="n">Com impostos</th></tr></thead><tbody>`;
  R.months.forEach(m => h += `<tr><td>${R.relative ? 'Mês ' + m.idx : m.key}</td><td class="n">${nfPT(m.H)}</td><td class="n">${C(m.net)}</td><td class="n">${C(m.gross)}</td></tr>`);
  $('tbMonth').innerHTML = h + `<tr class="tot"><td>Total</td><td class="n">${nfPT(T.H)}</td><td class="n">${C(T.net)}</td><td class="n">${C(T.gross)}</td></tr></tbody></table>`;

  h = `<table class="t"><thead><tr><th>#</th><th>Evento</th><th>Faturamento</th><th>Vencimento</th><th class="n">%</th><th class="n">Com impostos</th></tr></thead><tbody>`;
  R.pays.forEach((x, i) => h += `<tr><td>${i + 1}</td><td>${esc(x.trigger === 'phase' ? 'Conclusão: ' + x.label : x.trigger === 'month' ? 'Medição ' + (R.relative ? 'mês ' + x.idx : x.label) : 'Parcela ' + x.label)}</td><td>${R.relative ? (x.trigger === 'phase' ? 'Fim de ' + U + (x.period + 1) : 'Fim do mês ' + x.idx) : E.iso(x.when)}</td><td>${R.relative ? '+' + p.terms + ' dias' : E.iso(x.due)}</td><td class="n">${nfPT(x.pct * 100, 1)}%</td><td class="n">${C(x.gross)}</td></tr>`);
  $('tbPay').innerHTML = h + `<tr class="tot"><td colspan="5">Total</td><td class="n">${C(R.pays.reduce((a, x) => a + x.gross, 0))}</td></tr></tbody></table>`;

  // sensibilidade
  const discs = [0, .025, .05, .075, .10, .15, .20], conts = [0, .05, .10];
  const grid = E.sensitivity(model(), p, {today: new Date()}, discs, conts);
  h = `<table class="t sens"><thead><tr><th>Desconto ↓ / Contingência →</th>${conts.map(c => `<th class="n">${nfPT(c * 100, 1)}%</th>`).join('')}</tr></thead><tbody>`;
  grid.forEach(r => {
    h += `<tr><td><b>${nfPT(r.disc * 100, 1)}%</b></td>` + r.cells.map(c => {
      const cls = c.hasCost && c.gm < 0 ? 'bad' : (c.hasCost && p.gmMin > 0 && c.gm < p.gmMin ? 'low' : '');
      const cur_ = Math.abs(r.disc - p.disc) < 1e-9 && Math.abs(c.cont - p.cont) < 1e-9 ? ' style="outline:2px solid var(--blue)"' : '';
      return `<td class="cell ${cls}"${cur_}>${C(c.gross)}<br><span class="hint">${c.hasCost ? nfPT(c.gm * 100, 1) + '% GM' : '–'}</span></td>`;
    }).join('') + '</tr>';
  });
  $('tbSens').innerHTML = h + '</tbody></table>';
}
async function drawCharts(){
  if (!S.result) return;
  let Chart;
  try { Chart = await loadLib('chart'); }
  catch(e){ ['chFteMsg', 'chMonthMsg'].forEach(i => msg(i, 'warn', 'Gráficos indisponíveis: ' + esc(e.message))); return; }
  ['chFteMsg', 'chMonthMsg'].forEach(i => msg(i, '', ''));
  const R = S.result, U = R.unit;
  Object.values(charts).forEach(c => c.destroy());
  const labels = R.periods.map((x, i) => U + (i + 1));
  const ds = R.phases.map(ph => ({label: ph.name + (ph.optional ? ' (opc.)' : ''),
    data: R.periods.map((_, j) => R.rows.filter(r => r.phase === ph.name).reduce((a, r) => a + r.alloc[j], 0)),
    backgroundColor: ph.optional ? ph.color + '88' : ph.color, stack: 's'}));
  charts.fte = new Chart($('chFte'), {type: 'bar', data: {labels, datasets: ds}, options: {maintainAspectRatio: false, plugins: {legend: {position: 'bottom', labels: {boxWidth: 10, font: {size: 11}}}}, scales: {x: {stacked: true}, y: {stacked: true, title: {display: true, text: 'FTE'}}}}});
  let acc = 0;
  charts.m = new Chart($('chMonth'), {type: 'bar', data: {labels: R.months.map(m => R.relative ? 'M' + m.idx : m.key), datasets: [
    {label: 'Faturamento mensal', data: R.months.map(m => Math.round(m.gross)), backgroundColor: '#1B4F8C', yAxisID: 'y'},
    {label: 'Acumulado', type: 'line', data: R.months.map(m => Math.round(acc += m.gross)), borderColor: '#13A89E', backgroundColor: '#13A89E', yAxisID: 'y1', tension: .25}]},
    options: {maintainAspectRatio: false, plugins: {legend: {position: 'bottom', labels: {boxWidth: 10}}}, scales: {y: {position: 'left'}, y1: {position: 'right', grid: {drawOnChartArea: false}}}}});
}

/* ---------------- render: proposta ---------------- */
function renderDoc(){
  const R = S.result; if (!R) return;
  $('doc').innerHTML = DocModel.toHtml(DocModel.build(R), S.edits);
  const lang = R.p.lang;
  $('doc').querySelectorAll('[data-ek]').forEach(el => el.classList.toggle('edited', S.edits[lang + ':' + el.dataset.ek] !== undefined));
}
$('doc').addEventListener('input', e => {
  const el = e.target.closest('[data-ek]'); if (!el) return;
  S.edits[$('p_lang').value + ':' + el.dataset.ek] = DocModel.sanitizeHtml(el.innerHTML);
  el.classList.add('edited'); markDirty();
});
$('doc').addEventListener('paste', e => {
  if (!e.target.closest('[data-ek]')) return;
  e.preventDefault();
  document.execCommand('insertText', false, (e.clipboardData || window.clipboardData).getData('text/plain'));
});
function restoreTexts(){
  const lang = $('p_lang').value, keys = Object.keys(S.edits).filter(k => k.startsWith(lang + ':'));
  if (!keys.length) return toast('Nenhum texto editado neste idioma.', 'info');
  if (!confirm(`Restaurar os ${keys.length} texto(s) editado(s) para o original gerado pelo app?`)) return;
  keys.forEach(k => delete S.edits[k]); renderDoc(); markDirty();
}

/* ---------------- exportações ---------------- */
async function exportWord(){
  const R = S.result || recalc(false); if (!R) return toast('Carregue um timeline primeiro.', 'warn');
  try {
    const d = await loadLib('docx');
    const blob = await d.Packer.toBlob(DocxExport.build(d, DocModel.build(R), S.edits));
    Store.download(blob, fileBase() + '.docx');
  } catch(e){ toast('Não foi possível gerar o Word: ' + e.message, 'err', 9000); }
}
async function exportPptx(){
  const R = S.result || recalc(false); if (!R) return toast('Carregue um timeline primeiro.', 'warn');
  try {
    const G = await loadLib('pptx');
    const pres = PptxExport.build(G, DocModel.build(R), R, S.edits);
    Store.download(await pres.write({outputType: 'blob'}), fileBase() + '.pptx');
  } catch(e){ toast('Não foi possível gerar o PowerPoint: ' + e.message, 'err', 9000); }
}
async function downloadTemplate(){
  try {
    const Ex = await loadLib('exceljs');
    const wb = await Workbook.buildTemplate(Ex);
    Store.download(new Blob([await wb.xlsx.writeBuffer()], {type: XLSX_MIME}), 'Timeline_Template.xlsx');
  } catch(e){ toast('Não foi possível gerar o template: ' + e.message, 'err', 9000); }
}

/* ---------------- extras online (sempre opcionais, com aviso e alternativa manual) ---------------- */
const fxKey = (a, b) => `pm_fx_${a}_${b}`;
function cachedFx(){ try { return JSON.parse(Store.ls.get(fxKey($('p_rcur').value, $('p_cur').value)) || 'null'); } catch(e){ return null; } }
function refreshFxCachedBtn(){ const c = cachedFx(); const b = $('btnFxCached'); b.style.display = c ? '' : 'none'; if (c) b.textContent = `Usar última cotação salva (${nfPT(c.rate, 4)} em ${c.date})`; }
async function fxOnline(){
  const from = $('p_rcur').value, to = $('p_cur').value;
  if (from === to){ $('p_fx').value = 1; msg('fxMsg', 'ok', 'Mesma moeda: câmbio = 1.'); S.result = null; return markDirty(); }
  msg('fxMsg', 'info', 'Buscando cotação…');
  try {
    const r = await Net.fetchFx(from, to);
    $('p_fx').value = +r.rate.toFixed(6);
    Store.ls.set(fxKey(from, to), JSON.stringify({rate: r.rate, date: r.date || new Date().toISOString().slice(0, 10), source: r.source}));
    msg('fxMsg', 'ok', `✅ 1 ${from} = ${nfPT(r.rate, 4)} ${to} (${esc(r.date)} · ${esc(r.source)}).`);
    S.result = null; refreshFxCachedBtn(); markDirty(); renderRateTable();
  } catch(e){
    const c = cachedFx();
    msg('fxMsg', 'warn', `⚠ Não foi possível obter a cotação: ${esc(e.message)}. Digite o câmbio manualmente${c ? ' ou use a última cotação salva' : ''}. O restante do app não é afetado.`);
  }
}
function fxCached(){
  const c = cachedFx(); if (!c) return;
  $('p_fx').value = c.rate; S.result = null; markDirty();
  msg('fxMsg', 'info', `Usando a cotação salva: 1 ${esc($('p_rcur').value)} = ${nfPT(c.rate, 4)} ${esc($('p_cur').value)} (${esc(c.date)}). Confira se ainda é adequada.`);
}
async function holidaysOnline(){
  const R = S.result || recalc(false); if (!R || R.relative) return msg('holMsg', 'info', 'Defina uma data ou mês de início para conferir os feriados do período.');
  const y0 = R.startBase.getUTCFullYear(), y1 = Math.max(y0, R.T.end.getUTCFullYear());
  msg('holMsg', 'info', 'Consultando BrasilAPI…');
  try {
    const known = new Set(R.holidays.map(h => h.date)), added = [];
    for (let y = y0; y <= y1; y++) (await Net.fetchHolidays(y)).forEach(h => { if (!known.has(h.date)) added.push(h); });
    if (!added.length) return msg('holMsg', 'ok', `✅ Conferido com a BrasilAPI (${y0}${y1 > y0 ? '–' + y1 : ''}): nenhuma diferença em relação ao cálculo local.`);
    const cur = $('p_holx').value.trim();
    $('p_holx').value = (cur ? cur + '\n' : '') + added.map(h => `${h.date} ${h.name}`).join('\n');
    S.result = null; markDirty();
    msg('holMsg', 'ok', `✅ ${added.length} feriado(s) novo(s) adicionado(s) aos feriados adicionais: ${esc(added.map(h => h.date + ' ' + h.name).join('; '))}.`);
  } catch(e){
    msg('holMsg', 'warn', `⚠ Não foi possível consultar a BrasilAPI: ${esc(e.message)}. O cálculo local dos feriados nacionais continua valendo; use “Feriados adicionais” para incluir datas manualmente.`);
  }
}

/* ---------------- biblioteca de propostas (pasta com arquivos .xlsx) ---------------- */
async function libSummary(f){
  const key = f.name + '|' + f.lastModified;
  if (LIB.cache.has(key)) return LIB.cache.get(key);
  let out;
  try {
    const X = await loadLib('xlsx');
    const wb = X.read(new Uint8Array(await Store.fs.read(f.handle)), {type: 'array'});
    if (!Workbook.isProposalWorkbook(X, wb)) out = {other: true};
    else {
      const st = Workbook.read(X, wb);
      const R = E.compute({rows: st.rows, nPeriods: st.nPeriods, unit: st.unit, rates: st.rates}, E.parseParams(st.raw, {taxes: st.taxes}), {today: new Date()});
      out = {client: R.p.client, project: R.p.project, cur: R.p.cur, gross: R.T.gross, gm: R.T.gm, hasCost: R.T.cost > 0, H: R.T.H, nPer: R.T.nPer, unit: R.unit, issues: R.issues.filter(i => i.level === 'err').length};
    }
  } catch(e){ out = {error: e.message}; }
  LIB.cache.set(key, out);
  return out;
}
async function renderLibrary(){
  const body = $('libBody');
  if (!Store.fs.supported()){
    body.innerHTML = `<div class="msg warn">Este navegador não permite acessar pastas do computador (use Chrome ou Edge para isso). Alternativa: <b>💾 Salvar</b> baixa o arquivo .xlsx da proposta e <b>📂 Abrir proposta</b> carrega de volta — o conteúdo é o mesmo.</div>`;
    return;
  }
  if (!LIB.dir) LIB.dir = (await Store.kv.get('dir')) || null;
  if (!LIB.dir){
    body.innerHTML = `<p>Escolha uma pasta do computador (pode ser OneDrive/rede) para guardar e listar as propostas. Cada proposta é um arquivo <code>.xlsx</code>.</p><button class="btn" data-act="lib-pick">Escolher pasta</button>`;
    return;
  }
  body.innerHTML = `<div class="row" style="margin-bottom:10px"><b>📁 ${esc(LIB.dir.name)}</b>
    <button class="btn sec sm" data-act="lib-pick">Trocar pasta</button>
    <button class="btn sm" data-act="lib-here">💾 Salvar a proposta atual aqui</button>
    <button class="btn sec sm" data-act="lib-refresh">Atualizar</button></div><div id="libList">Carregando…</div><div id="libCmp"></div>`;
  if (!(await Store.fs.perm(LIB.dir, true))){
    $('libList').innerHTML = `<div class="msg warn">O navegador precisa da sua permissão para acessar esta pasta. <button class="btn sm" data-act="lib-refresh">Permitir acesso</button></div>`;
    return;
  }
  try { LIB.files = await Store.fs.list(LIB.dir); } catch(e){ $('libList').innerHTML = `<div class="msg err">Não foi possível ler a pasta: ${esc(e.message)}</div>`; return; }
  if (!LIB.files.length){ $('libList').innerHTML = '<p class="hint">Nenhum arquivo .xlsx nesta pasta ainda. Use “Salvar a proposta atual aqui”.</p>'; return; }
  $('libList').innerHTML = `<table class="t"><thead><tr><th></th><th>Arquivo</th><th>Cliente / Projeto</th><th class="n">Preço c/ impostos</th><th class="n">GM%</th><th>Modificado</th><th></th></tr></thead><tbody>${
    LIB.files.map((f, i) => `<tr><td><input type="checkbox" data-cmp="${i}" aria-label="Comparar ${esc(f.name)}"></td><td>${esc(f.name)}</td><td id="lc${i}" class="hint">…</td><td class="n" id="lp${i}"></td><td class="n" id="lg${i}"></td><td>${new Date(f.lastModified).toLocaleString('pt-BR')}</td><td><button class="btn sm" data-act="lib-open" data-i="${i}">Abrir</button></td></tr>`).join('')
  }</tbody></table><div class="row" style="margin-top:10px"><button class="btn sec sm" data-act="lib-compare">Comparar selecionadas</button></div>`;
  LIB.files.forEach(async (f, i) => {
    const s = await libSummary(f);
    if (!$('lc' + i)) return;
    if (s.error){ $('lc' + i).textContent = 'ilegível: ' + s.error; return; }
    if (s.other){ $('lc' + i).textContent = 'não é uma proposta do app'; return; }
    $('lc' + i).textContent = `${s.client} — ${s.project}`;
    $('lp' + i).textContent = curPT(s.gross, s.cur);
    $('lg' + i).textContent = s.hasCost ? nfPT(s.gm * 100, 1) + '%' : '–';
  });
}
async function libCompare(){
  const idx = [...document.querySelectorAll('#libList input[data-cmp]:checked')].map(i => +i.dataset.cmp);
  if (idx.length < 2) return toast('Marque pelo menos duas propostas para comparar.', 'info');
  const sums = await Promise.all(idx.map(i => libSummary(LIB.files[i]).then(s => ({f: LIB.files[i], s}))));
  const ok = sums.filter(x => !x.s.other && !x.s.error);
  const base = ok[0];
  const mixed = new Set(ok.map(x => x.s.cur)).size > 1;
  const d = (v, b) => base && b ? ((v - b) / b * 100) : 0;
  $('libCmp').innerHTML = `<h3>Comparação</h3>${mixed ? '<div class="msg warn">As propostas usam moedas diferentes: os valores não são diretamente comparáveis.</div>' : ''}
    <table class="t"><thead><tr><th>Proposta</th><th>Moeda</th><th class="n">Duração</th><th class="n">Horas</th><th class="n">Preço c/ impostos</th><th class="n">Δ preço</th><th class="n">GM%</th></tr></thead><tbody>${
      ok.map(x => `<tr><td>${esc(x.s.client)} — ${esc(x.s.project)}<br><span class="hint">${esc(x.f.name)}</span></td><td>${esc(x.s.cur)}</td><td class="n">${x.s.nPer} ${x.s.unit === 'SP' ? 'sprints' : 'sem.'}</td><td class="n">${nfPT(x.s.H)}</td><td class="n">${curPT(x.s.gross, x.s.cur)}</td><td class="n">${x === base || mixed ? '–' : nfPT(d(x.s.gross, base.s.gross), 1) + '%'}</td><td class="n">${x.s.hasCost ? nfPT(x.s.gm * 100, 1) + '%' : '–'}</td></tr>`).join('')
    }</tbody></table>`;
}
async function libPick(){
  try { LIB.dir = await Store.fs.pickDir(); await Store.kv.set('dir', LIB.dir); LIB.cache.clear(); renderLibrary(); }
  catch(e){ if (!aborted(e)) toast('Não foi possível escolher a pasta: ' + e.message, 'err'); }
}
async function libHere(){
  if (!S.rows.length) return toast('Carregue um timeline primeiro.', 'warn');
  try {
    if (!(await Store.fs.perm(LIB.dir, true))) return toast('Sem permissão de escrita na pasta.', 'warn');
    const name = fileBase() + '.xlsx';
    if (LIB.files.some(f => f.name === name) && !confirm(`Já existe "${name}" nesta pasta. Sobrescrever?`)) return;
    const h = await Store.fs.create(LIB.dir, name, await buildBlob());
    S.handle = h; S.autoFile = true; S.dirty = false; setSaveStatus(`Salvo em ${h.name} · ${hhmm()}`);
    toast('Proposta salva na pasta.', 'ok'); LIB.cache.clear(); renderLibrary();
  } catch(e){ if (!aborted(e)) toast('Não foi possível salvar na pasta: ' + e.message, 'err', 9000); }
}
async function libOpen(i){
  const f = LIB.files[i]; if (!f) return;
  try { await loadBuffer(await Store.fs.read(f.handle), f.name, f.handle); $('dlgLib').close(); toast('Proposta aberta.', 'ok'); }
  catch(e){ toast('Não foi possível abrir: ' + e.message, 'err', 9000); }
}

/* ---------------- ações (botões com data-act) ---------------- */
const actions = {
  template: downloadTemplate, example: loadExample, open: actOpen, save: actSave, saveas: actSaveAs, netcheck: netCheck,
  library: () => { const d = $('dlgLib'); if (!d.open) d.showModal(); renderLibrary(); },
  'lib-close': () => $('dlgLib').close(), 'lib-pick': libPick, 'lib-here': libHere, 'lib-refresh': () => { LIB.cache.clear(); renderLibrary(); },
  'lib-open': b => libOpen(+b.dataset.i), 'lib-compare': libCompare,
  'go-par': () => goTab('par'), 'go-doc': () => goTab('doc'), calc: () => recalc(true),
  'rc-import': () => $('rcFile').click(),
  'rc-gm': () => {
    const gm = P().gm; let c = 0;
    Object.values(S.rates).forEach(r => { if (r.cost){ r.rate = Math.round(r.cost / (1 - gm) * 100) / 100; c++; } });
    renderRateTable(); S.result = null; markDirty();
    msg('rcMsg', c ? 'ok' : 'warn', c ? `Rate/h recalculado para ${c} perfis com GM ${nfPT(gm * 100, 1)}%.` : 'Preencha o Custo/h primeiro.');
  },
  'tax-add': () => { S.taxes.push({name: 'Imposto', pct: 0}); renderTaxes(); markDirty(); },
  'tax-del': b => { if (S.taxes.length > 1){ S.taxes.splice(+b.dataset.i, 1); renderTaxes(); S.result = null; markDirty(); } },
  'fx-online': fxOnline, 'fx-cached': fxCached, 'holidays-online': holidaysOnline,
  'restore-texts': restoreTexts, word: exportWord, pptx: exportPptx,
  print: () => { if (!S.result) recalc(false); window.print(); }
};
document.addEventListener('click', e => {
  const b = e.target.closest('[data-act]'); if (!b || b.disabled) return;
  const fn = actions[b.dataset.act];
  if (fn) Promise.resolve().then(() => fn(b)).catch(err => toast('Falha: ' + err.message, 'err', 8000));
});

/* ---------------- mudança de parâmetros ---------------- */
const DOC_ONLY = new Set(['p_lang', 'p_tlmode', 'p_tlgran', 'p_tlfte', 'p_tlmiles']);
document.querySelector('main').addEventListener('change', e => {
  const id = e.target.id; if (!id || !id.startsWith('p_')) return;
  S.result = null;
  if (id === 'p_startmode') syncStartUI();
  if (id === 'p_unit' && S.rows.length){ S.unit = e.target.value === 'auto' ? S.unitAuto : e.target.value; renderGantt(); }
  if (id === 'p_context') Object.keys(S.edits).filter(k => k.endsWith(':exec.ctx')).forEach(k => delete S.edits[k]);
  if (id === 'p_assump') Object.keys(S.edits).filter(k => k.endsWith(':assump')).forEach(k => delete S.edits[k]);
  if (id === 'p_cur' || id === 'p_rcur') refreshFxCachedBtn();
  if (S.rows.length && ['p_mode', 'p_pw', 'p_hpd', 'p_dpw', 'p_unit'].includes(id)) renderRateTable();
  // se a análise ou a proposta estiver na tela, atualiza na hora; senão recalcula ao entrar na aba
  const visible = TABS.find(x => $('tab-' + x).classList.contains('active'));
  if (S.rows.length && (visible === 'doc' || visible === 'ana' || (DOC_ONLY.has(id) && !$('tabbtn-doc').disabled))) recalc(false);
  markDirty();
});

/* ---------------- upload ---------------- */
const drop = $('drop');
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => handleFile(e.dataTransfer.files[0]));
drop.addEventListener('click', () => $('tlFile').click());
drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' '){ e.preventDefault(); $('tlFile').click(); } });
$('tlFile').addEventListener('change', e => { handleFile(e.target.files[0]); e.target.value = ''; });
$('rcFile').addEventListener('change', e => { importRateCard(e.target.files[0]); e.target.value = ''; });

/* ---------------- inicialização ---------------- */
function init(){
  if (!$('p_start').value) $('p_start').value = E.iso(E.nextMonday());
  if (!$('p_startm').value){ const d = E.nextMonday(); $('p_startm').value = E.iso(d).slice(0, 7); }
  renderTaxes(); syncStartUI(); renderNet(); refreshFxCachedBtn();
  if (!Store.fs.supported()) $('savebar').textContent = 'Salvar baixa o .xlsx (este navegador não grava direto em arquivo)';
  Store.kv.get('draft').then(d => {
    if (!d || !d.snap || !d.snap.rows || !d.snap.rows.length) return;
    const at = new Date(d.at);
    toast(`Há um rascunho salvo em ${at.toLocaleDateString('pt-BR')} ${hhmm(at)}.`, 'info', 0, [
      {label: 'Restaurar', fn: () => { applyState(d.snap); S.handle = null; S.dirty = true; toast('Rascunho restaurado.', 'ok'); }},
      {label: 'Descartar', fn: () => Store.kv.del('draft')}
    ]);
  });
}
init();
window.__pm = {S, recalc, goTab, loadBuffer, applyState, snapshot, raw, setRaw, buildBlob, actions, renderLibrary}; // ganchos para testes
})();

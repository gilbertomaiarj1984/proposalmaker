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
  rows: [], nPeriods: 0, unit: 'SP', rates: {}, fileName: '', fileLabel: '', saveText: '', saveKind: 'dirty',
  taxes: [{name: 'PIS/COFINS/ISS', pct: 5.65}], phaseDisc: {}, edits: {},
  // catálogo de perfis e taxas (arquivo Excel, uma aba por cliente); snapRates = taxas gravadas na proposta (detecção de mudança)
  cat: {loaded: false, name: '', handle: null, buffer: null, clients: [], skipped: [], dirty: false, saving: false, version: 0, status: '', kind: 'dirty', needsPerm: false},
  catClient: '', snapRates: {}, drift: null,
  result: null, handle: null, autoFile: false, dirty: false
};
const LIB = {dir: null, files: [], cache: new Map()};
const charts = {};
let ED = null; // editor de timeline (criado no init)

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
const P = () => E.parseParams(raw(), {taxes: S.taxes, phaseDisc: S.phaseDisc});
const model = () => ({rows: S.rows, nPeriods: S.nPeriods, unit: S.unit, rates: S.rates});
const snapshot = () => ({raw: raw(), taxes: JSON.parse(JSON.stringify(S.taxes)), phaseDisc: JSON.parse(JSON.stringify(S.phaseDisc)), rows: S.rows, nPeriods: S.nPeriods, unit: S.unit, rates: S.rates, edits: S.edits, fileName: S.fileName, catalogName: S.cat.name});

function syncStartUI(){
  const m = $('p_startmode').value;
  $('lbl_start').style.display = m === 'date' ? '' : 'none';
  $('lbl_startm').style.display = m === 'month' ? '' : 'none';
}

/* ---------------- abas ---------------- */
const TABS = ['tl', 'par', 'ana', 'doc'];
function goTab(t){
  if ((t === 'ana' || t === 'doc') && !S.result && !recalc(false)) return;
  TABS.forEach(x => {
    const b = $('tabbtn-' + x), on = x === t;
    b.setAttribute('aria-selected', on ? 'true' : 'false'); b.tabIndex = on ? 0 : -1;
    $('tab-' + x).classList.toggle('active', on);
  });
  if (t === 'ana') drawCharts();
  if (t === 'par') renderDiscount();
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

/* ---------------- identificação da proposta (barra superior) ---------------- */
function safeCompute(){ try { return S.rows.length ? E.compute(model(), P(), {today: new Date()}) : null; } catch(e){ return null; } }
function updateIdent(){
  const has = S.rows.length > 0;
  const client = $('p_client').value.trim(), project = $('p_project').value.trim(), ref = $('p_ref').value.trim(), vendor = $('p_vendor').value.trim();
  const title = !has ? 'Nenhuma proposta aberta' : ([client, project].filter(Boolean).join(' — ') || 'Proposta sem nome');
  $('pbTitle').textContent = title;
  document.title = has ? `${title} · Proposta` : 'Gerador de Proposta Comercial';
  const parts = [];
  if (has){
    if (ref) parts.push(`<span>Ref. ${esc(ref)}</span>`);
    if (vendor) parts.push(`<span>${esc(vendor)}</span>`);
    parts.push(S.handle ? `<span title="Arquivo onde a proposta é gravada">📄 ${esc(S.handle.name)}</span>` : (S.fileLabel ? `<span title="O arquivo original não é alterado">📂 origem: ${esc(S.fileLabel)}</span>` : '<span>📄 sem arquivo</span>'));
    const R = safeCompute();
    if (R && R.T.gross > 0) parts.push(`<b title="Preço total com impostos">${esc(curPT(R.T.gross, R.p.cur))}</b>${R.T.cost ? `<span title="Margem bruta estimada">GM ${nfPT(R.T.gm * 100, 1)}%</span>` : ''}`);
    if (R && R.issues.some(i => i.level === 'err')) parts.push('<span class="pbwarn" title="Há alertas críticos na análise financeira">⛔ alertas</span>');
  }
  const st = has && S.saveText ? `<span class="st ${S.saveKind}">${esc(S.saveText)}</span>` : '';
  $('pbSub').innerHTML = parts.join('<span class="dotsep">·</span>') + (parts.length && st ? '<span class="dotsep">·</span>' : '') + st;
}
function updateFileBar(){
  $('fname').textContent = S.handle ? S.handle.name : (S.fileLabel ? `${S.fileLabel} (original preservado)` : 'Sem arquivo');
}

/* ---------------- persistência: rascunho, arquivo .xlsx e biblioteca ---------------- */
let saveTimer = null;
function setSaveStatus(t, kind){
  S.saveText = t; S.saveKind = ['saved', 'saving', 'dirty'].includes(kind) ? kind : 'dirty';
  const f = $('fstatus'); f.textContent = t; f.className = 'st ' + S.saveKind;
  $('saveNow').style.display = (S.rows.length && S.dirty && (!S.handle || !S.autoFile)) ? '' : 'none';
  updateFileBar(); updateIdent();
}
// liga a proposta a um arquivo (a partir daqui, as alterações são gravadas nele)
function setHandle(h){ S.handle = h; S.autoFile = !!h; $('autoFile').checked = !!h; if (h) S.fileLabel = h.name; }
function markDirty(){
  S.dirty = true; setSaveStatus('Alterações não salvas', 'dirty');
  clearTimeout(saveTimer); saveTimer = setTimeout(autosave, 1500);
}
async function autosave(){
  if (!S.rows.length) return;
  const ok = await Store.kv.set('draft', {snap: snapshot(), at: Date.now()});
  let text = ok ? `Rascunho salvo ${hhmm()} (só neste navegador)` : 'Rascunho local indisponível neste navegador', kind = 'dirty';
  if (S.handle && S.autoFile){
    setSaveStatus('⟳ Gravando no Excel…', 'saving');
    try {
      if (await Store.fs.granted(S.handle, true)){ await writeFile(S.handle); S.dirty = false; text = `✓ Salvo em ${S.handle.name} · ${hhmm()}`; kind = 'saved'; }
      else text += ' · clique em 💾 Salvar para reconectar o arquivo';
    } catch(e){ text += ' · falha ao gravar o arquivo: ' + e.message; }
  } else if (!S.handle) text += ' · use “Salvar como…” para gravar em Excel';
  setSaveStatus(text, kind);
}
async function buildBlob(){
  const R = E.compute(model(), P(), {today: new Date()}) || Workbook.emptyResult(snapshot()); // também grava propostas ainda sem alocação
  const Ex = await loadLib('exceljs');
  const wb = await Workbook.buildProposal(Ex, snapshot(), R);
  return new Blob([await wb.xlsx.writeBuffer()], {type: XLSX_MIME});
}
async function writeFile(handle){ await Store.fs.write(handle, await buildBlob()); }
const aborted = e => e && (e.name === 'AbortError' || /abort/i.test(e.message || ''));

async function actSave(){
  if (!S.rows.length) return toast('Abra ou crie uma proposta primeiro.', 'warn');
  try {
    if (S.handle && Store.fs.supported() && await Store.fs.perm(S.handle, true)){
      await writeFile(S.handle); S.dirty = false; S.autoFile = true; $('autoFile').checked = true;
      setSaveStatus(`✓ Salvo em ${S.handle.name} · ${hhmm()}`, 'saved'); toast('Proposta salva em ' + S.handle.name, 'ok'); return;
    }
    return actSaveAs();
  } catch(e){ if (!aborted(e)) toast('Não foi possível salvar: ' + e.message, 'err', 9000); }
}
async function actSaveAs(){
  if (!S.rows.length) return toast('Abra ou crie uma proposta primeiro.', 'warn');
  try {
    const name = fileBase() + '.xlsx';
    if (Store.fs.supported()){
      const h = await Store.fs.pickSave(name);
      await writeFile(h); setHandle(h); S.dirty = false;
      setSaveStatus(`✓ Salvo em ${h.name} · ${hhmm()}`, 'saved'); toast('Proposta salva. As próximas alterações serão gravadas neste arquivo automaticamente.', 'ok', 7000);
    } else {
      Store.download(await buildBlob(), name); S.dirty = false; S.fileLabel = name;
      setSaveStatus(`Arquivo baixado · ${hhmm()}`, 'saved');
      toast('Seu navegador não permite gravar direto no arquivo: o .xlsx foi baixado. Use “Abrir do Excel” para continuar depois e Salvar para baixar de novo.', 'info', 9000);
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
function actNew(){
  if (S.rows.length && S.dirty && !confirm('Há alterações ainda não gravadas em Excel. Começar uma proposta nova mesmo assim?')) return;
  ED.startEmpty(); S.phaseDisc = {};
  S.handle = null; S.autoFile = false; $('autoFile').checked = false; S.fileName = ''; S.fileLabel = ''; S.result = null; S.dirty = false;
  S.rates = {}; S.snapRates = {}; S.catClient = ''; S.drift = null;
  ['p_client', 'p_project', 'p_ref', 'p_context', 'p_assump'].forEach(id => { $(id).value = ''; });
  msg('tlMsg', 'info', S.cat.loaded ? 'Proposta nova. Selecione o <b>cliente</b> (aba 2, “Dados da proposta”, ou o seletor acima) para listar os perfis do catálogo; depois escolha perfil e nível em cada linha e preencha o FTE.' : 'Proposta nova. Carregue o catálogo de perfis (<b>📚 Catálogo</b>) ou cadastre os perfis pela combo de cada linha (<i>➕ Cadastrar novo perfil…</i>); depois preencha o FTE. Use <b>Salvar como…</b> para criar o arquivo Excel e começar a gravar nele.');
  afterLoad(true); markDirty();
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
    S.fileLabel = name; S.dirty = false;
    if (handle) setHandle(handle); else { S.handle = null; S.autoFile = false; $('autoFile').checked = false; }
    msg('tlMsg', 'ok', `✅ Proposta <b>${esc(name)}</b> aberta: ${S.rows.length} linhas, ${S.nPeriods} ${S.unit === 'SP' ? 'sprints' : 'semanas'}, ${Object.keys(S.rates).length} perfis, ${Object.keys(S.edits).length} texto(s) editado(s).` + (st.warnings.length ? '<br>⚠ ' + st.warnings.map(esc).join('<br>⚠ ') : ''));
    setSaveStatus(handle ? `✓ Aberta de ${name} — as alterações são gravadas nele` : `Aberta de ${name} — use Salvar para gravar`, handle ? 'saved' : 'dirty');
    return;
  }

  let tlName = wb.SheetNames.find(n => /timeline|cronograma|gantt|plan/i.test(n)) || null;
  let parsed = null; const errors = [];
  const candidates = tlName ? [tlName, ...wb.SheetNames.filter(n => n !== tlName)] : wb.SheetNames;
  for (const n of candidates){
    if (/rate|perfi|instruc|alt format/i.test(n)) continue;
    try { parsed = E.parseTimeline(aoa(n)); parsed.sheet = n; break; }
    catch(e){ errors.push(`${esc(n)}: ${esc(e.message)}`); }
  }
  if (!parsed) throw new Error('Nenhuma aba reconhecida como timeline.<br>' + errors.join('<br>'));
  S.rows = parsed.rows; S.nPeriods = parsed.nPeriods; S.unit = parsed.unit; $('p_unit').value = S.unit;
  // um timeline "bruto" nunca é sobrescrito: a proposta só passa a ser gravada depois de "Salvar como…"
  S.fileName = name; S.fileLabel = name; S.handle = null; S.autoFile = false; $('autoFile').checked = false; S.edits = {}; S.phaseDisc = {}; S.snapRates = {}; S.catClient = '';
  E.ensureRates(S.rates, S.rows);
  const rcName = wb.SheetNames.find(n => /rate|perfi/i.test(n));
  let rcInfo = '';
  if (rcName){
    try { const n = E.applyRateTable(S.rates, E.parseRateRows(aoa(rcName))); rcInfo = ` Perfis e taxas importados da aba "${esc(rcName)}" (${n} de ${Object.keys(S.rates).length} perfis/níveis).`; }
    catch(e){ rcInfo = ` ⚠ Perfis/taxas não importados: ${esc(e.message)}.`; }
  }
  afterTimelineLoaded(`Arquivo <b>${esc(name)}</b> (aba "${esc(parsed.sheet)}"): ${S.rows.length} linhas, ${S.nPeriods} ${S.unit === 'SP' ? 'sprints' : 'semanas'}.${rcInfo} <b>O arquivo original não será alterado:</b> use “Salvar como…” para criar a proposta e começar a gravar nela.` + (parsed.warnings.length ? '<br>⚠ ' + parsed.warnings.join('<br>⚠ ') : ''));
  setSaveStatus('Timeline importado — ainda não gravado (o original é preservado)', 'dirty');
}
// mostra o editor (e, se for uma carga nova, reinicia histórico e abas)
function afterLoad(reset){
  if (reset){ ED.resetHistory(); enableTabs(1); }
  const has = S.rows.length > 0;
  $('tlEmpty').style.display = has ? 'none' : ''; $('tlEditor').style.display = has ? 'block' : 'none';
  if (reset) bindCatalog(); else { renderClientUI(); renderRateTable(); renderCatalogBar(); }
  if (has) ED.render(); renderDiscount(); updateFileBar(); updateIdent();
}
function applyState(st){
  S.rows = st.rows; S.nPeriods = st.nPeriods; S.unit = st.unit; S.rates = st.rates || {};
  S.snapRates = JSON.parse(JSON.stringify(S.rates)); S.catClient = ''; S.drift = null;   // taxas gravadas na proposta (cópia) — o catálogo, se disponível, prevalece
  S.taxes = st.taxes && st.taxes.length ? st.taxes : [{name: 'Impostos', pct: 5.65}];
  S.phaseDisc = st.phaseDisc || {};
  S.edits = st.edits || {}; S.fileName = st.fileName || '';
  E.ensureRates(S.rates, S.rows);
  setRaw(st.raw || {});
  $('p_unit').value = S.unit;
  renderTaxes(); S.result = null; refreshFxCachedBtn(); afterLoad(true);
  // as taxas passam a vir do catálogo do cliente (se carregado); senão, valem as gravadas na proposta
  if (S.cat.loaded) bindCatalog({drift: true, warnMissing: true});
  else if (Object.keys(S.rates).length) toast(`Catálogo de perfis não carregado: usando as taxas gravadas nesta proposta${st.catalogName ? ` (catálogo: ${st.catalogName})` : ''}.`, 'info', 10000, [{label: 'Abrir catálogo…', fn: () => actions.catalog()}]);
}
function afterTimelineLoaded(info){
  msg('tlMsg', 'ok', '✅ ' + info);
  if (!$('p_project').value && S.fileName) $('p_project').value = S.fileName.replace(/\.(xlsx|xlsm|xls|csv)$/i, '').replace(/[_-]+/g, ' ');
  S.result = null; afterLoad(true); markDirty();
}
function loadExample(){
  const ex = Workbook.exampleData();
  S.rows = ex.rows.map(r => ({phase: r[0], activity: r[1], role: r[2], level: r[3], location: r[4], optional: r[5] === 'Y', alloc: r[6]}));
  S.nPeriods = ex.N; S.unit = 'SP'; S.fileName = ''; S.fileLabel = ''; S.rates = {}; S.snapRates = {}; S.catClient = ''; S.edits = {}; S.phaseDisc = {}; S.handle = null; S.autoFile = false; $('autoFile').checked = false;
  E.ensureRates(S.rates, S.rows);
  E.applyRateTable(S.rates, ex.rates.map(r => ({role: r[0], level: r[1], location: r[2], cost: r[3], rate: r[4]})));
  if (!$('p_client').value) $('p_client').value = 'Client';
  if (!$('p_project').value) $('p_project').value = 'Data & Analytics Modernization';
  $('p_unit').value = 'SP';
  afterTimelineLoaded('Exemplo carregado (12 sprints, taxas ilustrativas — substitua pelas taxas oficiais dos perfis).');
}
async function importRateCard(file){
  if (!file) return;
  try {
    const X = await loadLib('xlsx');
    const wb = X.read(new Uint8Array(await file.arrayBuffer()), {type: 'array'});
    const n = wb.SheetNames.find(x => /rate|perfi/i.test(x)) || wb.SheetNames[0];
    const c = E.applyRateTable(S.rates, E.parseRateRows(X.utils.sheet_to_json(wb.Sheets[n], {header: 1, defval: '', raw: true, blankrows: false})));
    renderRateTable(); S.result = null; markDirty();
    msg('rcMsg', c ? 'ok' : 'warn', c ? `✅ ${c} perfis atualizados a partir de "${esc(file.name)}".` : `Nenhum perfil do arquivo "${esc(file.name)}" corresponde aos perfis do timeline.`);
  } catch(e){ msg('rcMsg', 'err', '❌ ' + esc(e.message)); }
}

/* ---------------- desconto (aba 2) ---------------- */
let discNames = [];
function renderDiscount(){
  const mode = $('p_discmode').value;
  $('l_disc').style.display = mode === 'pct' ? '' : 'none';
  $('l_discamt').style.display = mode === 'amount' ? '' : 'none';
  $('l_disctarget').style.display = mode === 'target' ? '' : 'none';
  $('discPhaseBox').style.display = mode === 'pct' ? '' : 'none';
  const R = safeCompute();
  if (!R){
    $('discResult').innerHTML = ''; $('discSuggest').innerHTML = '';
    $('discPhase').innerHTML = '<tbody><tr><td class="hint">Escolha perfis e preencha o FTE no timeline para configurar o desconto.</td></tr></tbody>';
    return;
  }
  const p = R.p, T = R.T, C = (v, d = 0) => curPT(v, p.cur, d), kp = (b, t) => `<div class="kp"><b>${b}</b>${t}</div>`;
  $('discResult').innerHTML =
    kp(C(T.listGross), 'preço de lista (c/ impostos)') + kp((T.discNet > 0.005 ? '−' : '') + C(T.discGross) + ` <small>(${nfPT(T.discPct * 100, 1)}%)</small>`, 'desconto efetivo') +
    kp(C(T.gross), 'preço final (c/ impostos)') + kp(T.cost ? nfPT(T.gm * 100, 1) + '%' : '–', 'GM% após o desconto');

  // por fase (modo percentual)
  const phases = R.phases.filter(ph => !ph.optional || p.inclOpt); discNames = phases.map(ph => ph.name);
  const gross = v => p.tax < 1 ? v / (1 - p.tax) : v;
  $('discPhase').innerHTML = `<thead><tr><th>Fase</th><th class="n">Preço de lista (c/ impostos)</th><th class="n" style="width:150px">Desconto da fase (%)</th><th class="n">Preço final (c/ impostos)</th></tr></thead><tbody>` +
    phases.map((ph, i) => {
      const rr = R.rows.filter(r => r.phase === ph.name && (!r.optional || p.inclOpt)), list = rr.reduce((a, r) => a + r.listNet, 0), net = rr.reduce((a, r) => a + r.net, 0);
      const ov = S.phaseDisc[ph.name];
      return `<tr><td>${esc(ph.name)}</td><td class="n">${C(gross(list))}</td><td class="n"><input type="number" min="0" max="100" step="0.5" data-i="${i}" value="${ov !== undefined ? esc(ov) : ''}" placeholder="geral ${nfPT(R.dGlobal * 100, 1)}" aria-label="Desconto da fase ${esc(ph.name)}"></td><td class="n">${C(gross(net))}</td></tr>`;
    }).join('') + '</tbody>';

  // sugestões de como aplicar
  const gmMin = p.gmMin, floorTxt = gmMin > 0 ? `GM ≥ ${nfPT(gmMin * 100, 1)}%` : 'GM ≥ 0% (ponto de equilíbrio)';
  const maxPct = R.maxDisc === null ? null : Math.floor(R.maxDisc * 1000) / 10;
  const tryPct = d => { const x = E.compute(model(), {...p, discMode: 'pct', discPhase: {}, disc: d}, {today: new Date()}); return x ? {gross: x.T.gross, gm: x.T.gm, hasCost: x.T.cost > 0} : null; };
  const ladder = [2.5, 5, 7.5, 10].map(d => { const t = tryPct(d / 100); return t ? `<button class="chip ${maxPct !== null && d > maxPct ? 'bad' : ''}" data-act="disc-apply" data-v="${d}" title="Aplicar ${d}% sobre o total">${nfPT(d, 1)}% → ${C(t.gross)}${t.hasCost ? ` · GM ${nfPT(t.gm * 100, 1)}%` : ''}</button>` : ''; }).join(' ');
  $('discSuggest').innerHTML = `<div class="sugg"><h3>💡 Sugestões de como aplicar</h3><ul>
    <li><b>Limite de segurança:</b> ${maxPct === null ? 'informe o Custo/h dos perfis para calcular o desconto máximo.' : `o maior desconto geral que mantém ${floorTxt} é <b>${nfPT(maxPct, 1)}%</b>. <button class="btn sec sm" data-act="disc-apply" data-v="${maxPct}">Aplicar ${nfPT(maxPct, 1)}%</button>` + (gmMin > 0 ? '' : ' Defina o <i>GM% mínimo</i> em “Premissas de preço” para proteger sua margem.')}</li>
    <li><b>Escada de negociação</b> (preço final e margem de cada degrau): ${ladder}</li>
    <li><b>Mostrar o desconto como linha</b> (“preço de lista → desconto → preço final”) ancora o valor e permite trocar desconto por escopo, prazo ou pagamento antecipado na negociação. Use “embutido” só quando não quiser expor a lista.</li>
    <li><b>Concentre o desconto por fase</b> onde o risco é menor ou a entrada é estratégica (ex.: assessment/mobilização) e preserve a margem nas fases de execução.</li>
    <li><b>Já tem orçamento do cliente?</b> Use <i>Preço final alvo</i>: o app calcula o % equivalente e avisa se a margem cair abaixo do mínimo.</li></ul></div>`;
}
$('discPhase').addEventListener('change', e => {
  const i = e.target.closest('input[data-i]'); if (!i) return;
  const name = discNames[+i.dataset.i], v = i.value === '' ? null : E.num(i.value);
  if (v === null || v <= 0) delete S.phaseDisc[name]; else S.phaseDisc[name] = Math.min(100, v);
  S.result = null; markDirty(); renderDiscount();
});

/* ---------------- catálogo de perfis e taxas (Excel, uma aba por cliente) ---------------- */
const NAT = (a, b) => String(a).localeCompare(String(b), 'pt', {numeric: true});
const normName = x => E.norm(String(x || ''));
const boundClient = () => S.catClient ? S.cat.clients.find(c => c.name === S.catClient) : null;
const clientByName = n => (S.cat.loaded && String(n || '').trim()) ? S.cat.clients.find(c => normName(c.name) === normName(n)) : null;
const hasLoc = () => { const c = boundClient(); return !c || !!c.cols.location; };
const hasCost = () => { const c = boundClient(); return !c || !!c.cols.cost; };
const liveEntries = cl => cl.entries.filter(e => !e.deleted);
let catTimer = null;

// taxas efetivas = catálogo do cliente + perfis usados que não estão no catálogo (mantêm a cópia gravada na proposta)
function rebuildRates(){
  const cl = boundClient(); if (!cl) return;
  const rates = {};
  liveEntries(cl).forEach(e => { const k = E.rkey(e); if (!rates[k]) rates[k] = {role: e.role, level: e.level, location: e.location, cost: e.cost, rate: e.rate}; });
  S.rows.filter(r => r.role).forEach(r => { const k = E.rkey(r); if (!rates[k]) rates[k] = {...(S.snapRates[k] || {role: r.role, level: r.level, location: r.location, cost: '', rate: ''}), notInCatalog: true}; });
  S.rates = rates;
}
function refreshAfterRates(){
  S.result = null; renderRateTable(); renderCatalogBar();
  if (S.rows.length){ ED.render(); renderDiscount(); updateIdent(); }
}
// vincula a proposta à aba do cliente (pelo nome em "Cliente"); sem aba correspondente, segue com as taxas gravadas
function bindCatalog(opts){
  opts = opts || {};
  const cl = clientByName($('p_client').value);
  S.catClient = cl ? cl.name : '';
  if (cl){
    if (cl.currency && [...$('p_rcur').options].some(o => o.value === cl.currency)) $('p_rcur').value = cl.currency;
    rebuildRates();
    if (opts.drift) checkDrift();
    if (opts.switched){
      const miss = new Set(S.rows.filter(r => r.role && S.rates[E.rkey(r)] && S.rates[E.rkey(r)].notInCatalog).map(r => `${r.role} ${r.level}`.trim()));
      if (miss.size) toast(`${miss.size} perfil(is) do timeline não existem no catálogo de ${cl.name}: escolha os perfis pelas combos da aba 1.`, 'warn', 9000);
    }
  } else if (S.cat.loaded && opts.warnMissing && $('p_client').value.trim()){
    toast(`O cliente “${$('p_client').value.trim()}” não existe no catálogo: usando as taxas gravadas na proposta.`, 'warn', 12000, [{label: 'Criar cliente', fn: () => openNewClient($('p_client').value.trim())}]);
  }
  renderClientUI(); refreshAfterRates();
}
// compara as taxas gravadas na proposta com as atuais do catálogo e informa o impacto no preço
function checkDrift(){
  S.drift = null;
  const snap = S.snapRates, cl = boundClient();
  if (!cl || !Object.keys(snap).length) return;
  const changed = [];
  Object.keys(snap).forEach(k => { const cur = S.rates[k]; if (!cur || cur.notInCatalog) return; const a = snap[k]; if (E.num(a.rate) !== E.num(cur.rate) || E.num(a.cost) !== E.num(cur.cost)) changed.push(`${cur.role} ${cur.level}`.trim()); });
  if (!changed.length) return;
  const old = {...S.rates}; Object.keys(snap).forEach(k => { if (old[k]) old[k] = {...old[k], cost: snap[k].cost, rate: snap[k].rate}; });
  const p = P(), calc = rates => E.compute({rows: S.rows, nPeriods: S.nPeriods, unit: S.unit, rates}, p, {today: new Date()});
  const b = calc(old), n = calc(S.rates);
  S.drift = {client: cl.name, changed, before: b ? b.T.gross : null, after: n ? n.T.gross : null, cur: p.cur};
  const money = v => v === null ? '—' : curPT(v, p.cur);
  toast(`As taxas do catálogo de ${cl.name} mudaram em ${changed.length} perfil(is) desde que esta proposta foi gravada. Preço com impostos: ${money(S.drift.before)} → ${money(S.drift.after)}.`, 'warn', 15000);
}

/* --- persistência do catálogo no Excel (mesmo princípio da proposta: abre de lá, grava de volta) --- */
function setCatStatus(text, kind){ S.cat.status = text; S.cat.kind = ['saved', 'saving', 'dirty'].includes(kind) ? kind : 'dirty'; renderCatalogBar(); updateCatBtn(); }
function markCatDirty(){
  S.cat.dirty = true; S.cat.version++;
  setCatStatus('Alterações no catálogo não gravadas', 'dirty');
  clearTimeout(catTimer); catTimer = setTimeout(saveCatalog, 1200);
}
async function saveCatalog(){
  const C = S.cat; if (!C.loaded || !C.dirty || C.saving) return;
  if (!C.handle || !(await Store.fs.granted(C.handle, true))){
    setCatStatus(C.handle ? 'Catálogo: falta permissão de escrita — use “Salvar catálogo”' : 'Alterações no catálogo não gravadas — use “Salvar catálogo” (baixa o .xlsx)', 'dirty'); return;
  }
  await writeCatalog(C.handle);
}
function applySaved(r){
  S.cat.clients.forEach(cl => {
    if (r.sheets[cl.name] !== undefined){ if (cl.isNew){ cl.isNew = false; delete cl.base; } cl.sheet = r.sheets[cl.name]; }
    cl.entries = cl.entries.filter(e => !(e.deleted && r.removed.includes(e.id)));
    cl.entries.forEach(e => { if (r.rows[e.id] !== undefined) e.row = r.rows[e.id]; });
  });
}
async function writeCatalog(handle){
  const C = S.cat; C.saving = true; const v = C.version;
  setCatStatus('⟳ Gravando o catálogo…', 'saving');
  try {
    const Ex = await loadLib('exceljs');
    const r = await Catalog.serialize(Ex, C.buffer, {clients: C.clients});
    const blob = new Blob([r.buffer], {type: XLSX_MIME});
    if (handle) await Store.fs.write(handle, blob); else Store.download(blob, C.name || 'Perfis_e_Taxas.xlsx');
    C.buffer = r.buffer; applySaved(r);
    if (C.version === v) C.dirty = false;
    setCatStatus(C.dirty ? 'Alterações no catálogo não gravadas' : `✓ Catálogo gravado · ${hhmm()}`, C.dirty ? 'dirty' : 'saved');
  } catch(e){
    setCatStatus('Falha ao gravar o catálogo: ' + e.message, 'dirty'); toast('Não foi possível gravar o catálogo: ' + e.message, 'err', 9000);
  } finally { C.saving = false; }
  if (S.cat.dirty && S.cat.handle){ clearTimeout(catTimer); catTimer = setTimeout(saveCatalog, 400); }
}
async function catalogLoad(buf, name, handle){
  const Ex = await loadLib('exceljs');
  const parsed = await Catalog.parse(Ex, buf);
  if (!parsed.clients.length) throw new Error('nenhuma aba de cliente reconhecida (esperado: cabeçalho com “Role/Perfil” e a coluna de taxa)');
  Object.assign(S.cat, {loaded: true, name, handle: handle || null, buffer: buf, clients: parsed.clients, skipped: parsed.skipped, dirty: false, saving: false, needsPerm: false, version: 0});
  if (handle) Store.kv.set('catalog', handle);
  setCatStatus(handle ? `✓ ${name} conectado — as alterações são gravadas nele` : `${name} carregado (neste navegador, “Salvar catálogo” baixa o .xlsx)`, handle ? 'saved' : 'dirty');
  bindCatalog({drift: Object.keys(S.snapRates).length > 0, warnMissing: S.rows.length > 0});
  renderCatalogDialog();
}
async function catalogOpen(){
  try {
    if (!Store.fs.supported()) return $('catFile').click();
    const h = await Store.fs.pickOpen();
    await catalogLoad(await Store.fs.read(h), h.name, h);
  } catch(e){ if (!aborted(e)) toast('Não foi possível abrir o catálogo: ' + e.message, 'err', 9000); }
}
async function catalogRestore(){
  const h = await Store.kv.get('catalog'); if (!h) return;
  S.cat.handle = h; S.cat.name = h.name;
  if (await Store.fs.granted(h, false)){
    try { await catalogLoad(await Store.fs.read(h), h.name, h); } catch(e){ setCatStatus('Não foi possível reabrir o catálogo ' + h.name + ': ' + e.message, 'dirty'); }
  } else { S.cat.needsPerm = true; setCatStatus(`Catálogo ${h.name}: clique em “Reconectar” para liberar o acesso`, 'dirty'); }
  renderCatalogDialog();
}
async function catalogReconnect(){
  const h = S.cat.handle; if (!h) return;
  try { if (await Store.fs.perm(h, true)) await catalogLoad(await Store.fs.read(h), h.name, h); else toast('Acesso ao catálogo negado.', 'warn'); }
  catch(e){ toast('Não foi possível reconectar: ' + e.message, 'err', 8000); }
}
async function catalogCreate(){
  try {
    const Ex = await loadLib('exceljs');
    const buf = await Catalog.seed(Ex).xlsx.writeBuffer();
    let handle = null;
    if (Store.fs.supported()){
      handle = await Store.fs.pickSave('Perfis_e_Taxas.xlsx');
      await Store.fs.write(handle, new Blob([buf], {type: XLSX_MIME}));
    } else Store.download(new Blob([buf], {type: XLSX_MIME}), 'Perfis_e_Taxas.xlsx');
    await catalogLoad(buf, handle ? handle.name : 'Perfis_e_Taxas.xlsx', handle);
    toast('Catálogo criado com as abas Roche e JnJ (dados de exemplo): substitua pelos oficiais.', 'ok', 8000);
  } catch(e){ if (!aborted(e)) toast('Não foi possível criar o catálogo: ' + e.message, 'err', 9000); }
}
async function catalogSave(){
  const C = S.cat; if (!C.loaded) return;
  try {
    if (C.handle && Store.fs.supported() && await Store.fs.perm(C.handle, true)) await writeCatalog(C.handle);
    else await writeCatalog(null);
  } catch(e){ if (!aborted(e)) toast('Não foi possível salvar o catálogo: ' + e.message, 'err', 9000); }
}

/* --- clientes (uma aba por cliente) --- */
function openNewClient(prefill, baseName){
  if (!S.cat.loaded){
    toast('Carregue ou crie o catálogo de perfis para cadastrar clientes.', 'warn', 8000, [{label: 'Abrir catálogo…', fn: () => actions.catalog()}]); return;
  }
  $('ncName').value = prefill || '';
  const def = baseName || (boundClient() && boundClient().name) || (S.cat.clients[0] && S.cat.clients[0].name) || '';
  $('ncBase').innerHTML = S.cat.clients.map(c => `<option value="${esc(c.name)}" ${c.name === def ? 'selected' : ''}>${esc(c.name)} (${liveEntries(c).length} perfis)</option>`).join('') + '<option value="">(em branco — só o cabeçalho)</option>';
  const d = $('dlgClient'); if (!d.open) d.showModal(); $('ncName').focus();
}
function createClient(name, baseName){
  const base = baseName ? S.cat.clients.find(c => c.name === baseName) : null;
  const sheet = Catalog.sheetName(name, S.cat.clients.map(c => c.sheet));
  const cl = base
    ? {name: sheet, sheet, isNew: true, base: base.sheet, hdrRow: base.hdrRow, cols: {...base.cols}, currency: base.currency, skipped: [],
       entries: liveEntries(base).map(e => ({...e, id: Catalog.newId()}))}
    : {name: sheet, sheet, isNew: true, hdrRow: 1, cols: {role: 1, level: 2, location: 3, cost: 4, rate: 5}, currency: '', entries: []};
  S.cat.clients.push(cl); markCatDirty();
  return cl;
}
$('ncForm').addEventListener('submit', e => {
  e.preventDefault();
  const name = $('ncName').value.trim(); if (!name) return;
  if (clientByName(name)) return toast('Já existe um cliente com esse nome no catálogo.', 'warn');
  const cl = createClient(name, $('ncBase').value);
  $('dlgClient').close();
  useClient(cl.name, true);
  toast(`Cliente “${cl.name}” criado no catálogo${$('ncBase').value ? ` com base em ${$('ncBase').value}` : ''}. Ajuste as taxas em “Perfis e taxas”.`, 'ok', 8000);
  renderCatalogDialog();
});
function useClient(name, scroll){
  $('p_client').value = name; S.snapRates = {}; S.drift = null;
  bindCatalog({switched: true}); markDirty();
  if (scroll) $('rcCard').scrollIntoView({block: 'nearest'});
}
function renderClientUI(){
  const loaded = S.cat.loaded, inp = $('p_client'), cur = inp.value.trim();
  const opts = loaded ? S.cat.clients.map(c => `<option value="${esc(c.name)}" ${normName(c.name) === normName(cur) ? 'selected' : ''}>${esc(c.name)}</option>`) : [];
  if (loaded && cur && !clientByName(cur)) opts.unshift(`<option value="${esc(cur)}" selected>${esc(cur)} (fora do catálogo)</option>`);
  if (loaded && !cur) opts.unshift('<option value="" selected>— selecione o cliente —</option>');
  ['p_clientsel', 'p_clientsel1'].forEach(id => { $(id).innerHTML = opts.join(''); });
  $('p_clientsel').style.display = loaded ? '' : 'none'; inp.style.display = loaded ? 'none' : '';
  $('clientBox1').style.display = loaded ? '' : 'none';
  const cl = boundClient();
  $('clientHint').textContent = !loaded ? 'Sem catálogo: carregue o Excel de perfis e taxas em “📚 Catálogo” para escolher o cliente.'
    : cl ? `Perfis e taxas do catálogo “${S.cat.name}” — aba ${cl.sheet}.` : 'Selecione um cliente do catálogo ou crie um novo (＋ Novo).';
}
['p_clientsel', 'p_clientsel1'].forEach(id => $(id).addEventListener('change', e => { if (e.target.value) useClient(e.target.value); }));

/* --- tabela "Perfis e taxas" (dados do cliente selecionado; edição grava no catálogo) --- */
function rateEdit(key, field, value){
  const cl = boundClient();
  if (!cl){ if (S.rates[key]) S.rates[key][field] = value; return; }
  let e = liveEntries(cl).find(x => E.rkey(x) === key);
  if (!e){ const r = S.rates[key]; e = {id: Catalog.newId(), role: r.role, level: r.level, location: r.location, cost: r.cost, rate: r.rate}; cl.entries.push(e); }
  e[field] = value; markCatDirty(); rebuildRates();
}
// cadastra um perfil/nível (no catálogo do cliente, ou localmente se não houver catálogo); devolve {entry, existed}
function addProfile(entry){
  const cl = boundClient();
  if (cl){
    if (!cl.cols.location) entry.location = '';
    if (!cl.cols.level) entry.level = '';
    const k = E.rkey(entry), found = liveEntries(cl).find(x => E.rkey(x) === k);
    if (!found){ cl.entries.push({id: Catalog.newId(), role: entry.role, level: entry.level, location: entry.location, cost: entry.cost === undefined ? '' : entry.cost, rate: entry.rate === undefined ? '' : entry.rate}); markCatDirty(); }
    rebuildRates(); refreshAfterRatesLite();
    return {entry: S.rates[k], existed: !!found};
  }
  const k = E.rkey(entry), existed = !!S.rates[k];
  if (!existed) S.rates[k] = entry;
  return {entry: S.rates[k], existed};
}
function refreshAfterRatesLite(){ S.result = null; renderRateTable(); renderCatalogBar(); }
function renderRateTable(){
  const p = P(), cl = boundClient();
  const q = $('rcFilter').value.trim().toLowerCase();
  const keys = Object.keys(S.rates).sort((a, b) => NAT(S.rates[a].role, S.rates[b].role) || NAT(S.rates[a].level, S.rates[b].level));
  const hp = (S.unit === 'W' ? 1 : p.pw) * p.dpw * p.hpd;
  const hrsBy = {}; S.rows.filter(r => r.role).forEach(r => { const k = E.rkey(r); hrsBy[k] = (hrsBy[k] || 0) + r.alloc.reduce((a, b) => a + b, 0) * hp; });
  const showLoc = !cl || !!cl.cols.location, showCost = !cl || !!cl.cols.cost, showLvl = !cl || !!cl.cols.level;
  $('rcTitle').textContent = cl ? `Perfis e taxas — ${cl.name}` : 'Perfis e taxas';
  $('rcLegacy').style.display = cl ? 'none' : '';
  $('btnRcGm').style.display = showCost ? '' : 'none';
  $('rcaLvlBox').style.display = showLvl ? '' : 'none'; $('rcaLocBox').style.display = showLoc ? '' : 'none'; $('rcaCostBox').style.display = showCost ? '' : 'none';
  const shown = keys.filter(k => { const r = S.rates[k]; return !q || `${r.role} ${r.level} ${r.location}`.toLowerCase().includes(q); });
  let h = `<thead><tr><th>Perfil</th>${showLvl ? '<th>Nível</th>' : ''}${showLoc ? '<th>Local</th>' : ''}<th class="n">Horas (timeline)</th>${showCost ? '<th class="n">Custo/h</th>' : ''}<th class="n">Rate/h (${esc(p.rcur)})</th>${showCost ? '<th class="n">GM%</th>' : ''}${cl ? '<th></th>' : ''}</tr></thead><tbody>`;
  shown.forEach(k => {
    const r = S.rates[k], gm = r.rate && r.cost ? (1 - r.cost / r.rate) * 100 : null;
    const miss = v => (v === '' || v === 0 || v === null) ? 'style="border-color:var(--err);background:#FFF5F5"' : '';
    const used = hrsBy[k] > 0;
    h += `<tr><td>${esc(r.role)}${r.notInCatalog ? ' <span class="tag warn" title="Está no timeline mas não no catálogo deste cliente">fora do catálogo</span>' : ''}</td>${showLvl ? `<td>${esc(r.level)}</td>` : ''}${showLoc ? `<td>${esc(r.location)}</td>` : ''}<td class="n">${nfPT(hrsBy[k] || 0)}</td>
      ${showCost ? `<td class="n"><input type="number" step="0.01" aria-label="Custo/h de ${esc(r.role)} ${esc(r.level)}" value="${esc(r.cost)}" data-k="${esc(k)}" data-f="cost" ${p.mode === 'gm' && used ? miss(r.cost) : ''}></td>` : ''}
      <td class="n"><input type="number" step="0.01" aria-label="Rate/h de ${esc(r.role)} ${esc(r.level)}" value="${esc(r.rate)}" data-k="${esc(k)}" data-f="rate" ${p.mode === 'rate' && used ? miss(r.rate) : ''}></td>
      ${showCost ? `<td class="n" style="${gm !== null && gm < 0 ? 'color:var(--err)' : ''}">${gm === null ? '–' : nfPT(gm, 1) + '%'}</td>` : ''}
      ${cl ? `<td class="n">${r.notInCatalog ? `<button class="btn sec sm" data-act="rc-addcat" data-k="${esc(k)}" title="Grava este perfil no catálogo do cliente">＋ catálogo</button>` : `<button class="xbtn" data-act="rc-del" data-k="${esc(k)}" title="Remover do catálogo" aria-label="Remover ${esc(r.role)} ${esc(r.level)} do catálogo">🗑</button>`}</td>` : ''}</tr>`;
  });
  $('rcTable').innerHTML = h + '</tbody>';
  $('rcCount').textContent = `${shown.length} de ${keys.length} perfis/níveis`;
}
$('rcFilter').addEventListener('input', () => renderRateTable());
$('rcTable').addEventListener('change', e => {
  const i = e.target.closest('input[data-k]'); if (!i) return;
  rateEdit(i.dataset.k, i.dataset.f, i.value === '' ? '' : E.num(i.value));
  S.result = null; markDirty(); renderRateTable();
  if (S.rows.length){ ED.render(); updateIdent(); }
});
function rcAdd(){
  const role = $('rcaRole').value.trim(); if (!role) return $('rcaRole').focus();
  const r = addProfile({role, level: $('rcaLevel').value.trim(), location: $('rcaLoc').value.trim(), cost: E.num($('rcaCost').value) || '', rate: E.num($('rcaRate').value) || ''});
  ['rcaRole', 'rcaLevel', 'rcaLoc', 'rcaCost', 'rcaRate'].forEach(id => { $(id).value = ''; });
  toast(r.existed ? 'Esse perfil/nível já existe — nada foi duplicado.' : 'Perfil adicionado.', r.existed ? 'info' : 'ok', 3500);
  refreshAfterRates();
}
function rcDelete(key){
  const cl = boundClient(); if (!cl) return;
  const e = liveEntries(cl).find(x => E.rkey(x) === key); if (!e) return;
  const used = S.rows.some(r => r.role && E.rkey(r) === key);
  if (!confirm(`Remover ${e.role}${e.level ? ' — ' + e.level : ''} do catálogo de ${cl.name}?${used ? '\nEle é usado no timeline desta proposta.' : ''}`)) return;
  e.deleted = true; markCatDirty(); rebuildRates(); refreshAfterRates(); markDirty();
}

/* --- barra e diálogo do catálogo --- */
function updateCatBtn(){
  const b = $('btnCat'), C = S.cat;
  b.className = 'btn light' + (C.loaded ? (C.dirty ? ' dirty' : ' loaded') : '');
  b.textContent = C.loaded ? `📚 ${C.name}` : (C.needsPerm ? '📚 Reconectar catálogo' : '📚 Catálogo');
  b.title = C.loaded ? `${C.status} — ${S.cat.clients.length} cliente(s)` : 'Catálogo de perfis e taxas (Excel, uma aba por cliente)';
}
function renderCatalogBar(){
  const C = S.cat, cl = boundClient(), chip = C.status ? `<span class="st ${C.kind}">${esc(C.status)}</span>` : '';
  const need = C.dirty && (!C.handle || !Store.fs.supported());
  $('rcBar').innerHTML = !C.loaded
    ? `<span>📚 <b>Sem catálogo carregado</b> — as taxas abaixo são locais a esta proposta.</span>${C.needsPerm ? '<button class="btn sm" data-act="cat-reconnect">Reconectar catálogo</button>' : ''}<button class="btn sec sm" data-act="catalog">Abrir/criar catálogo…</button>`
    : `<span>📚 <b>${esc(C.name)}</b></span>${chip}<span>· Cliente: <b>${cl ? esc(cl.name) : 'nenhum selecionado'}</b>${cl && cl.currency ? ` · moeda: ${esc(cl.currency)}` : ''}</span>
       <button class="btn sec sm" data-act="catalog">Catálogo e clientes…</button>${C.dirty || !C.handle ? `<button class="btn sm" data-act="cat-save">💾 Salvar catálogo</button>` : ''}${need ? '' : ''}`;
  updateCatBtn();
}
function renderCatalogDialog(){
  const body = $('catBody'), C = S.cat; if (!body) return;
  const chip = C.status ? `<span class="st ${C.kind}">${esc(C.status)}</span>` : '';
  let h = `<div class="catbar"><span>📚 <b>${C.loaded ? esc(C.name) : (C.handle ? esc(C.handle.name) : 'Nenhum catálogo')}</b></span>${chip}</div>
    <div class="row" style="margin:10px 0"><button class="btn" data-act="cat-open">📂 Abrir catálogo (Excel)…</button>
      <button class="btn sec" data-act="cat-create">＋ Criar catálogo com Roche e JnJ…</button>
      ${C.needsPerm ? '<button class="btn sm" data-act="cat-reconnect">Reconectar</button>' : ''}
      ${C.loaded ? '<button class="btn sec" data-act="cat-save">💾 Salvar catálogo</button>' : ''}</div>
    <p class="hint">O catálogo é um Excel com <b>uma aba por cliente</b> (perfis, níveis e taxas sem impostos). O arquivo da proposta guarda só o cliente escolhido e uma cópia das taxas usadas.</p>`;
  if (C.loaded){
    h += `<table class="t"><thead><tr><th>Cliente (aba)</th><th class="n">Perfis/níveis</th><th>Moeda</th><th>Colunas</th><th></th></tr></thead><tbody>` +
      C.clients.map(c => {
        const cols = [c.cols.role && 'Perfil', c.cols.level && 'Nível', c.cols.location && 'Local', c.cols.cost && 'Custo/h', c.cols.rate && 'Rate/h'].filter(Boolean).join(', ');
        return `<tr><td><b>${esc(c.name)}</b>${S.catClient === c.name ? ' <span class="st saved">nesta proposta</span>' : ''}${c.isNew ? ' <span class="st dirty">nova aba</span>' : ''}</td><td class="n">${liveEntries(c).length}</td><td>${esc(c.currency || '—')}</td><td class="hint">${esc(cols)}</td>
          <td class="n"><button class="btn sm" data-act="client-use" data-n="${esc(c.name)}">Usar nesta proposta</button> <button class="btn sec sm" data-act="client-new" data-base="${esc(c.name)}">Novo cliente a partir deste…</button></td></tr>`;
      }).join('') + `</tbody></table><div class="row" style="margin-top:10px"><button class="btn" data-act="client-new">＋ Novo cliente…</button></div>`;
    if (C.skipped.length) h += `<p class="hint">Abas ignoradas (sem colunas de perfil/taxa): ${esc(C.skipped.join(', '))}.</p>`;
  }
  body.innerHTML = h;
}

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
    if (!R){ toast('Escolha ao menos um perfil e preencha o FTE no timeline antes de calcular.', 'warn', 6000); return null; }
    S.result = R; enableTabs(3);
    renderAnalysis(); renderDoc();
    updateIdent(); renderDiscount();
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
    case 'missing_rates': return `Sem ${d.mode === 'rate' ? 'Rate/h' : 'Custo/h'} para: <b>${esc(d.list.join(', '))}</b>. Esses perfis entram com valor zero — preencha as taxas na aba 2 antes de enviar a proposta.`;
    case 'gm_negative': return `GM% negativo (${nfPT(d.gm * 100, 1)}%): o preço está abaixo do custo.`;
    case 'gm_below_min': return `GM% de ${nfPT(d.gm * 100, 1)}% está abaixo do mínimo configurado (${nfPT(d.min * 100, 1)}%).`;
    case 'no_cost': return 'Sem Custo/h informado: o GM% não pode ser calculado.';
    case 'price_below_cost': return `Preço abaixo do custo em: ${esc(d.list.join(', '))}.`;
    case 'start_adjusted': return `${esc(d.from)} não é dia útil (fim de semana ou feriado): o projeto começa em <b>${esc(d.to)}</b>.`;
    case 'relative_timeline': return 'Início sem data: a proposta mostra apenas períodos relativos (SP1, SP2…), sem datas de calendário.';
    case 'fx_missing': return `A moeda da proposta (${esc(d.to)}) difere da das taxas (${esc(d.from)}) e o câmbio está em 1. Informe ou busque a cotação.`;
    case 'target_above_list': return `O preço final alvo (${curPT(d.target, cur)}) é maior que o preço de lista (${curPT(d.list, cur)}): nenhum desconto foi aplicado.`;
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
    (hasBad ? '' : `<div class="it ok"><span aria-hidden="true">✅</span><span>Nenhum alerta crítico: taxas completas, margem dentro do esperado.</span></div>`);

  const k = (l, v, s = '') => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`;
  $('kpis').innerHTML =
    k('Preço total (c/ impostos)', C(T.gross), `Líquido: ${C(T.net)}` + (T.discNet > 0.5 ? ` · desconto ${nfPT(T.discPct * 100, 1)}% (lista ${C(T.listGross)})` : '')) +
    k('Duração', `${T.nPer} ${U === 'SP' ? 'sprints' : 'semanas'}`, R.relative ? `${T.weeks} semanas · sem data de início` : `${T.weeks} semanas · ${E.iso(T.start)} → ${E.iso(T.end)}`) +
    k('Esforço', `${nfPT(T.H)} h`, `${nfPT(T.H / (p.hpd * p.dpw * 52 / 12 || 1), 1)} FTE-mês`) +
    k('Time', `pico ${nfPT(T.peak, 1)} FTE`, `média ${nfPT(T.avgFte, 1)} FTE`) +
    k('Rate médio (blended)', C(T.blended, 2) + '/h', p.mode === 'rate' ? 'taxa do perfil' : 'custo + GM') +
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
      const R = E.compute({rows: st.rows, nPeriods: st.nPeriods, unit: st.unit, rates: st.rates}, E.parseParams(st.raw, {taxes: st.taxes, phaseDisc: st.phaseDisc}), {today: new Date()});
      if (!R) out = {error: 'sem alocação no timeline'}; else out = {client: R.p.client, project: R.p.project, cur: R.p.cur, gross: R.T.gross, gm: R.T.gm, hasCost: R.T.cost > 0, H: R.T.H, nPer: R.T.nPer, unit: R.unit, issues: R.issues.filter(i => i.level === 'err').length};
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
    setHandle(h); S.dirty = false; setSaveStatus(`✓ Salvo em ${h.name} · ${hhmm()}`, 'saved');
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
  template: downloadTemplate, example: loadExample, new: actNew, open: actOpen, save: actSave, saveas: actSaveAs, netcheck: netCheck,
  library: () => { const d = $('dlgLib'); if (!d.open) d.showModal(); renderLibrary(); },
  'lib-close': () => $('dlgLib').close(), 'lib-pick': libPick, 'lib-here': libHere, 'lib-refresh': () => { LIB.cache.clear(); renderLibrary(); },
  'lib-open': b => libOpen(+b.dataset.i), 'lib-compare': libCompare,
  'go-par': () => goTab('par'), 'go-doc': () => goTab('doc'), calc: () => recalc(true),
  'disc-apply': b => { $('p_discmode').value = 'pct'; $('p_disc').value = b.dataset.v; S.result = null; markDirty(); renderDiscount(); toast(`Desconto geral de ${nfPT(+b.dataset.v, 1)}% aplicado.`, 'ok', 3000); },
  catalog: () => { const d = $('dlgCat'); if (!d.open) d.showModal(); renderCatalogDialog(); },
  'cat-close': () => $('dlgCat').close(), 'cat-open': catalogOpen, 'cat-create': catalogCreate, 'cat-save': catalogSave, 'cat-reconnect': catalogReconnect,
  'client-new': b => openNewClient('', b && b.dataset ? b.dataset.base : ''), 'nc-close': () => $('dlgClient').close(),
  'client-use': b => { useClient(b.dataset.n); $('dlgCat').close(); toast(`Cliente “${b.dataset.n}” selecionado.`, 'ok', 3000); },
  'rc-add': rcAdd, 'rc-del': b => rcDelete(b.dataset.k),
  'rc-addcat': b => { const r = S.rates[b.dataset.k]; if (!r) return; const cl = boundClient(); if (!cl) return; cl.entries.push({id: Catalog.newId(), role: r.role, level: cl.cols.level ? r.level : '', location: cl.cols.location ? r.location : '', cost: r.cost, rate: r.rate}); markCatDirty(); rebuildRates(); refreshAfterRates(); },
  'rc-import': () => $('rcFile').click(),
  'rc-gm': () => {
    const gm = P().gm; let c = 0;
    Object.keys(S.rates).forEach(k => { const r = S.rates[k]; if (r.cost){ rateEdit(k, 'rate', Math.round(r.cost / (1 - gm) * 100) / 100); c++; } });
    renderRateTable(); S.result = null; markDirty(); if (S.rows.length) ED.render();
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
  const id = e.target.id; if (!id || !id.startsWith('p_') || id === 'p_clientsel' || id === 'p_clientsel1') return;
  S.result = null;
  if (id === 'p_client'){ S.snapRates = S.catClient || !S.cat.loaded ? S.snapRates : {}; bindCatalog({switched: true}); }
  if (id === 'p_startmode') syncStartUI();
  if (id === 'p_unit' && S.rows.length) S.unit = e.target.value === 'SP' ? 'SP' : 'W';
  if (id === 'p_context') Object.keys(S.edits).filter(k => k.endsWith(':exec.ctx')).forEach(k => delete S.edits[k]);
  if (id === 'p_assump') Object.keys(S.edits).filter(k => k.endsWith(':assump')).forEach(k => delete S.edits[k]);
  if (id === 'p_cur' || id === 'p_rcur') refreshFxCachedBtn();
  if (S.rows.length && ['p_mode', 'p_pw', 'p_hpd', 'p_dpw', 'p_unit'].includes(id)) renderRateTable();
  if (S.rows.length && ED.TL_IDS.includes(id)) ED.render();
  if (S.rows.length) renderDiscount();
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
  ED = TimelineEditor({S, E, $, esc, nfPT, toast, P, raw, setRaw, addProfile, hasLocation: hasLoc, catalogHint: () => (S.cat.loaded && !S.catClient) ? 'Selecione o cliente da proposta (aba 2 ou o seletor da aba 1) para listar os perfis do catálogo' : '',
    compute: () => E.compute(model(), P(), {today: new Date()}),
    changed: () => { S.result = null; markDirty(); }});
  renderTaxes(); syncStartUI(); renderNet(); refreshFxCachedBtn(); updateIdent(); renderClientUI(); renderRateTable(); renderCatalogBar();
  $('catFile').addEventListener('change', async e => { const f = e.target.files[0]; e.target.value = ''; if (!f) return; try { await catalogLoad(await f.arrayBuffer(), f.name, null); } catch(err){ toast('Não foi possível abrir o catálogo: ' + err.message, 'err', 9000); } });
  catalogRestore();
  if (!Store.fs.supported()){
    $('autoFile').disabled = true; $('autoFile').checked = false;
    $('fstatus').textContent = 'Este navegador não grava direto no arquivo: “Salvar” baixa o .xlsx';
  }
  $('autoFile').addEventListener('change', e => {
    S.autoFile = e.target.checked;
    if (S.autoFile && S.handle) autosave(); else setSaveStatus(S.saveText || 'Alterações não salvas', S.saveKind);
  });
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
window.__pm = {S, catalogLoad, boundClient, saveCatalog, createClient, rebuildRates, recalc, goTab, loadBuffer, applyState, snapshot, raw, setRaw, buildBlob, actions, renderLibrary}; // ganchos para testes
})();

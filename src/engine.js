/* =====================================================================
   ENGINE — lógica pura (sem DOM, sem rede). Funciona no navegador e no Node
   (testes em tests/engine.test.mjs). Nada aqui lê ou escreve na tela.
===================================================================== */
const Engine = (() => {
'use strict';

/* ---------------- utilidades ---------------- */
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g,'').toLowerCase().trim();
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function num(v){
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).trim(); if (!s) return 0;
  const pct = s.endsWith('%');
  s = s.replace('%','').replace(/\s/g,'');
  if (/,\d+$/.test(s) && s.indexOf('.') > -1) s = s.replace(/\./g,'').replace(',', '.'); // 1.234,5
  else s = s.replace(',', '.');
  const n = parseFloat(s);
  if (isNaN(n)) return 0;
  return pct ? n/100 : n;
}
function periodNum(v){
  if (typeof v === 'number') return Math.round(v);
  const m = String(v ?? '').match(/(\d+)/);
  return m ? parseInt(m[1],10) : 0;
}
const isYes = v => ['y','yes','s','sim','si','x','true','1','optional','opcional'].includes(norm(v));
const rkey = r => [r.role, r.level || '-', r.location || '-'].join('|');

// datas em UTC
function parseDate(s){ const [y,m,d] = String(s).split('-').map(Number); return new Date(Date.UTC(y, m-1, d)); }
const addDays = (d, n) => new Date(d.getTime() + n*86400000);
const iso = d => d.toISOString().slice(0,10);
const monthKey = d => d.toISOString().slice(0,7);
const isValidIso = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s||'')) && !isNaN(parseDate(s).getTime());
function nextMonday(from){
  const d = from || new Date();
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dow = t.getUTCDay();
  return addDays(t, ((8 - dow) % 7) || 7);
}

/* ---------------- feriados nacionais do Brasil (cálculo local, funciona offline) ---------------- */
function easter(y){ // algoritmo gregoriano anônimo
  const a = y%19, b = Math.floor(y/100), c = y%100, d = Math.floor(b/4), e = b%4, f = Math.floor((b+8)/25),
        g = Math.floor((b-f+1)/3), h = (19*a+b-d-g+15)%30, i = Math.floor(c/4), k = c%4,
        l = (32+2*e+2*i-h-k)%7, m = Math.floor((a+11*h+22*l)/451),
        mo = Math.floor((h+l-7*m+114)/31), da = ((h+l-7*m+114)%31)+1;
  return new Date(Date.UTC(y, mo-1, da));
}
function holidaysBR(y){
  const fx = (m,d,name) => ({date: iso(new Date(Date.UTC(y,m-1,d))), name});
  const E = easter(y);
  const list = [
    fx(1,1,'Confraternização Universal'), fx(4,21,'Tiradentes'), fx(5,1,'Dia do Trabalho'),
    fx(9,7,'Independência do Brasil'), fx(10,12,'Nossa Senhora Aparecida'), fx(11,2,'Finados'),
    fx(11,15,'Proclamação da República'), fx(12,25,'Natal'),
    {date: iso(addDays(E,-48)), name:'Carnaval'}, {date: iso(addDays(E,-47)), name:'Carnaval'},
    {date: iso(addDays(E,-2)), name:'Sexta-feira Santa'}, {date: iso(addDays(E,60)), name:'Corpus Christi'}
  ];
  if (y >= 2024) list.push(fx(11,20,'Consciência Negra'));
  return list.sort((a,b) => a.date.localeCompare(b.date));
}

/* ---------------- leitura do timeline ---------------- */
const ALIAS = {
  phase:    ['phase','fase','workstream','frente','wave','onda','stage','etapa','track'],
  activity: ['activity','atividade','actividad','deliverable','entregavel','entregable','task','tarefa'],
  role:     ['role','perfil','papel','funcao','rol','profile','resource','recurso','cargo'],
  level:    ['level','nivel','senioridade','seniority','grade','band','q level','qlevel'],
  location: ['location','local','localidade','localizacao','ubicacion','country','pais','delivery'],
  optional: ['optional','opcional','opt'],
  start:    ['start','inicio','start sprint','sprint inicio','from','de'],
  end:      ['end','fim','fin','end sprint','sprint fim','to','ate','hasta'],
  fte:      ['fte','alocacao','allocation','asignacion','qty','quantidade','hc','headcount']
};
const findCol = (headers, key) => headers.findIndex(h => ALIAS[key].includes(h));

function mkRow(line, col, phase, role, alloc){
  return {
    phase: phase || 'General',
    activity: col.activity >= 0 ? String(line[col.activity] ?? '').trim() : '',
    role,
    level: col.level >= 0 ? String(line[col.level] ?? '').trim() : '',
    location: col.location >= 0 ? String(line[col.location] ?? '').trim() : '',
    optional: col.optional >= 0 ? isYes(line[col.optional]) : /opcional|optional/i.test(phase),
    alloc
  };
}

function parseTimeline(aoa, opts){
  opts = opts || {};
  if (!aoa || !aoa.length) throw new Error('aba vazia');
  let hr = -1;
  for (let i = 0; i < Math.min(20, aoa.length); i++){
    if (aoa[i].some(c => ALIAS.role.includes(norm(c)))) { hr = i; break; }
  }
  if (hr < 0) throw new Error('coluna "Role/Perfil" não encontrada');
  const headers = aoa[hr].map(norm);
  const col = {}; Object.keys(ALIAS).forEach(k => col[k] = findCol(headers, k));
  const warnings = [];

  const pcols = [];
  headers.forEach((h, i) => {
    let m = h.match(/^(sp|sprint|s)\s*[-_#]?\s*(\d+)$/);
    if (m) { pcols.push({i, p:+m[2], u:'SP'}); return; }
    m = h.match(/^(w|wk|week|semana|sem)\s*[-_#]?\s*(\d+)$/);
    if (m) pcols.push({i, p:+m[2], u:'W'});
  });

  const rows = []; let lastPhase = '';
  let unit = 'SP', nPeriods = 0;

  if (pcols.length){
    unit = pcols.filter(c => c.u === 'W').length > pcols.length/2 ? 'W' : 'SP';
    nPeriods = Math.max(...pcols.map(c => c.p));
    for (let r = hr+1; r < aoa.length; r++){
      const line = aoa[r];
      const role = String(line[col.role] ?? '').trim();
      const ph = col.phase >= 0 ? String(line[col.phase] ?? '').trim() : '';
      if (ph) lastPhase = ph;
      if (/^total/i.test(role)) continue;
      if (!role && !(opts.keepEmpty && ph)) continue; // linha sem perfil só se mantém no modo editor e com fase explícita
      const alloc = new Array(nPeriods).fill(0);
      pcols.forEach(c => alloc[c.p-1] += num(line[c.i]));
      if (alloc.every(v => v === 0) && !opts.keepEmpty) { warnings.push(`Linha ${r+1} (${esc(role)}) sem alocação – ignorada.`); continue; }
      const row = mkRow(line, col, lastPhase, role, alloc); row.srcLine = r;
      rows.push(row);
    }
  } else if (col.start >= 0 && col.end >= 0){
    const tmp = [];
    for (let r = hr+1; r < aoa.length; r++){
      const line = aoa[r];
      const role = String(line[col.role] ?? '').trim();
      const ph = col.phase >= 0 ? String(line[col.phase] ?? '').trim() : '';
      if (ph) lastPhase = ph;
      if (!role || /^total/i.test(role)) continue;
      const s = periodNum(line[col.start]), e = periodNum(line[col.end]);
      const f = col.fte >= 0 ? num(line[col.fte]) : 1;
      if (!s || !e || e < s) { warnings.push(`Linha ${r+1}: início/fim inválidos – ignorada.`); continue; }
      tmp.push({line, role, phase:lastPhase, s, e, f});
      nPeriods = Math.max(nPeriods, e);
      const txt = String(line[col.start]).toLowerCase();
      if (/w|sem/.test(txt) && !/sp/.test(txt)) unit = 'W';
    }
    tmp.forEach(t => {
      const alloc = new Array(nPeriods).fill(0);
      for (let p = t.s; p <= t.e; p++) alloc[p-1] = t.f || 1;
      rows.push(mkRow(t.line, col, t.phase, t.role, alloc));
    });
  } else throw new Error('não encontrei colunas de período (SP1, SP2… / W1, W2…) nem Start/End');

  if (!rows.length) throw new Error('nenhuma linha com alocação');
  if (opts.unit === 'SP' || opts.unit === 'W') unit = opts.unit;
  return {rows, nPeriods, unit, warnings};
}

/* ---------------- rate card ---------------- */
function parseRateRows(aoa){
  const hr = aoa.findIndex(r => r.some(c => ALIAS.role.includes(norm(c))));
  if (hr < 0) throw new Error('Rate card sem coluna Role/Perfil');
  const h = aoa[hr].map(norm);
  const ci = {
    role: findCol(h,'role'), level: findCol(h,'level'), location: findCol(h,'location'),
    cost: h.findIndex(x => /^(cost|custo|costo)/.test(x)),
    rate: h.findIndex(x => /^(rate|price|preco|tarifa|taxa|bill)/.test(x))
  };
  const table = [];
  for (let r = hr+1; r < aoa.length; r++){
    const L = aoa[r]; const role = String(L[ci.role] ?? '').trim(); if (!role) continue;
    table.push({role, level: ci.level>=0 ? String(L[ci.level]??'').trim() : '', location: ci.location>=0 ? String(L[ci.location]??'').trim() : '',
      cost: ci.cost>=0 ? num(L[ci.cost]) : 0, rate: ci.rate>=0 ? num(L[ci.rate]) : 0});
  }
  return table;
}
const rateMatch = (t, r) => norm(t.role) === norm(r.role) && (!t.level || !r.level || norm(t.level) === norm(r.level)) && (!t.location || !r.location || norm(t.location) === norm(r.location));

// garante uma entrada de rate por perfil usado no timeline; o catálogo (Rate Card) é mantido,
// inclusive perfis ainda não usados. Com prune=true, remove os que não são usados.
function ensureRates(rates, rows, prune){
  const used = rows.filter(r => r.role);
  used.forEach(r => { const k = rkey(r); if (!rates[k]) rates[k] = {role:r.role, level:r.level, location:r.location, cost:'', rate:''}; });
  if (prune){ const keys = new Set(used.map(rkey)); Object.keys(rates).forEach(k => { if (!keys.has(k)) delete rates[k]; }); }
  return rates;
}
// aplica a tabela do rate card sobre as entradas existentes; devolve quantas foram preenchidas
function applyRateTable(rates, table){
  let count = 0;
  Object.values(rates).forEach(r => {
    const t = table.find(t => rateMatch(t, r));
    if (t){ if (t.cost) r.cost = t.cost; if (t.rate) r.rate = t.rate; count++; }
  });
  return count;
}

/* ---------------- parâmetros ---------------- */
const CONTMODES = ['price','both'];
function parseExtraHolidays(txt){
  const out = [];
  String(txt||'').split('\n').forEach(l => {
    const m = l.match(/(\d{4}-\d{2}-\d{2})/);
    if (m && isValidIso(m[1])) out.push({date:m[1], name: l.replace(m[1],'').replace(/^[\s;,|-]+/,'').trim() || 'Feriado'});
  });
  return out;
}
function parseMilestoneLines(txt){
  const out = [];
  String(txt||'').split('\n').forEach(l => {
    const parts = l.split('|').map(s => s.trim());
    if (parts.length < 2 || !parts[0] || !parts[1]) return;
    const ref = parts[1];
    let m;
    if (isValidIso(ref)) out.push({name: parts[0], date: ref});
    else if ((m = ref.match(/^(sp|sprint|s)\s*(\d+)$/i))) out.push({name: parts[0], unit:'SP', n:+m[2]});
    else if ((m = ref.match(/^(w|wk|sem|semana|week)\s*(\d+)$/i))) out.push({name: parts[0], unit:'W', n:+m[2]});
  });
  return out;
}
// raw = valores crus (strings/booleans) dos campos; extra = { taxes:[{name,pct}] }
function parseParams(raw, extra){
  raw = raw || {}; extra = extra || {};
  const str = (k, d='') => String(raw[k] ?? d).trim();
  const n = (k, d=0) => (raw[k] === undefined || raw[k] === '' || raw[k] === null) ? d : num(raw[k]);
  const taxes = (extra.taxes && extra.taxes.length ? extra.taxes : (raw.p_tax !== undefined ? [{name:'Impostos', pct:num(raw.p_tax)}] : [{name:'Impostos', pct:5.65}]))
    .map(t => ({name: String(t.name||'').trim() || 'Imposto', pct: Math.max(0, num(t.pct))}));
  const taxPct = taxes.reduce((a,t) => a + t.pct, 0) / 100;
  const startMode = ['date','month','none'].includes(raw.p_startmode) ? raw.p_startmode : 'date';
  return {
    vendor: str('p_vendor') || 'TCS', client: str('p_client') || 'Client', project: str('p_project') || 'Project', ref: str('p_ref'),
    startMode, start: str('p_start'), startMonth: str('p_startm'),
    lang: ['en','pt','es'].includes(raw.p_lang) ? raw.p_lang : 'en',
    model: raw.p_model === 'tm' ? 'tm' : 'fp',
    cur: str('p_cur') || 'BRL', rcur: str('p_rcur') || str('p_cur') || 'BRL', fx: n('p_fx', 1) > 0 ? n('p_fx', 1) : 1,
    context: str('p_context'),
    pw: Math.max(1, n('p_pw', 2) || 2), hpd: n('p_hpd', 8) > 0 ? n('p_hpd', 8) : 8, dpw: Math.min(7, Math.max(1, Math.round(n('p_dpw', 5)))),
    unitSel: ['SP','W'].includes(raw.p_unit) ? raw.p_unit : 'auto',
    mode: raw.p_mode === 'gm' ? 'gm' : 'rate', gm: Math.min(0.95, Math.max(0, n('p_gm', 35)/100)), gmMin: Math.max(0, n('p_gmmin', 0)/100),
    cont: Math.max(0, n('p_cont', 0)/100), contMode: CONTMODES.includes(raw.p_contmode) ? raw.p_contmode : 'price',
    disc: Math.min(0.99, Math.max(0, n('p_disc', 0)/100)),
    discMode: ['pct','amount','target'].includes(raw.p_discmode) ? raw.p_discmode : 'pct',
    discAmount: Math.max(0, n('p_discamt', 0)), discTarget: Math.max(0, n('p_disctarget', 0)),   // valores COM impostos (o que o cliente vê)
    discShow: raw.p_discshow === 'line' ? 'line' : 'embedded',
    discPhase: Object.fromEntries(Object.entries(extra.phaseDisc || {}).map(([k, v]) => [k, Math.min(1, Math.max(0, num(v) / 100))]).filter(([, v]) => v > 0)),
    taxes, tax: Math.min(0.9, taxPct), taxShow: !!raw.p_taxshow,
    inclOpt: !!raw.p_inclOpt,
    pay: ['milestone','monthly','install'].includes(raw.p_pay) ? raw.p_pay : 'milestone',
    inst: Math.max(1, Math.round(n('p_inst', 6))), terms: Math.max(0, n('p_terms', 30)), valid: Math.max(1, n('p_valid', 30)),
    assump: str('p_assump').split('\n').map(s => s.trim()).filter(Boolean),
    useHol: raw.p_hol === undefined ? true : !!raw.p_hol, holx: parseExtraHolidays(raw.p_holx),
    tl: {
      mode: ['bars','roles','milestones'].includes(raw.p_tlmode) ? raw.p_tlmode : 'bars',
      gran: raw.p_tlgran === 'week' ? 'week' : 'period',
      fte: raw.p_tlfte === undefined ? true : !!raw.p_tlfte,
      miles: parseMilestoneLines(raw.p_tlmiles)
    }
  };
}

/* ---------------- calendário (independe das linhas do timeline) ---------------- */
// feriados nacionais calculados localmente + extras manuais/online; início ajustado para dia útil;
// cada período termina em dia útil. Usado pelo cálculo e pela grade do editor.
function buildCalendar(p, N, unit, ctx){
  ctx = ctx || {};
  const relative = p.startMode === 'none';
  const pw = unit === 'W' ? 1 : p.pw;
  const hol = new Set();
  const holNames = {};
  const baseDate = (() => {
    if (relative) return nextMonday(ctx.today);
    if (p.startMode === 'month' && /^\d{4}-\d{2}$/.test(p.startMonth)) return parseDate(p.startMonth + '-01');
    if (isValidIso(p.start)) return parseDate(p.start);
    return nextMonday(ctx.today);
  })();
  if (p.useHol && !relative){
    const y0 = baseDate.getUTCFullYear(), y1 = y0 + Math.ceil((N * pw) / 52) + 1;
    for (let y = y0; y <= y1; y++) holidaysBR(y).forEach(h => { hol.add(h.date); holNames[h.date] = h.name; });
    p.holx.forEach(h => { hol.add(h.date); holNames[h.date] = h.name; });
    if (ctx.holidays) ctx.holidays.forEach(d => hol.add(d));
  }
  const isWork = d => { const dow = d.getUTCDay(), wd = dow === 0 ? 7 : dow; return wd <= p.dpw && !hol.has(iso(d)); };
  const nextWork = d => { let x = d; for (let i = 0; i < 400 && !isWork(x); i++) x = addDays(x, 1); return x; };
  const prevWork = d => { let x = d; for (let i = 0; i < 400 && !isWork(x); i++) x = addDays(x, -1); return x; };

  const effStart = nextWork(baseDate);
  const startAdjusted = !relative && iso(effStart) !== iso(baseDate);
  const pStart = i => nextWork(addDays(effStart, i * pw * 7));
  const pEnd = i => { const s = pStart(i), e = prevWork(addDays(effStart, (i + 1) * pw * 7 - 1)); return e < s ? s : e; };

  return {relative, pw, hol, holNames, baseDate, isWork, nextWork, prevWork, effStart, startAdjusted, pStart, pEnd};
}

/* ---------------- cálculo principal ---------------- */
// model = {rows, nPeriods, unit, rates}; ctx = {today?: Date, holidays?: Set<string> (extras online/manuais)}
function compute(model, p, ctx){
  ctx = ctx || {};
  if (!model || !model.rows || !model.rows.length) return null;
  const idxRows = model.rows.map((r, i) => ({r, i})).filter(x => x.r.role); // linhas sem perfil (editor) não entram no cálculo
  if (!idxRows.length) return null;
  const rows0 = idxRows.map(x => x.r), N = model.nPeriods, unit = model.unit, rates = model.rates || {};
  const relative = p.startMode === 'none';
  const pw = unit === 'W' ? 1 : p.pw;
  const hpp = pw * p.dpw * p.hpd;
  const priceF = (1 + p.cont); // fator de lista (contingência); o desconto é aplicado depois, por linha
  const costF = p.contMode === 'both' ? (1 + p.cont) : 1;
  const gross = v => p.tax < 1 ? v / (1 - p.tax) : v;

  const cal = buildCalendar(p, N, unit, ctx);
  const {hol, holNames, baseDate, isWork, nextWork, prevWork, effStart, startAdjusted, pStart, pEnd} = cal;

  // taxas (rate/h e custo/h já convertidos para a moeda da proposta)
  const rc = r => rates[rkey(r)] || {};
  const basePriceH = r => p.mode === 'rate' ? num(rc(r).rate) : (num(rc(r).cost) / (1 - p.gm));
  const priceH = r => basePriceH(r) * p.fx;
  const costH = r => num(rc(r).cost) * p.fx * costF;

  const missing = [];
  const usedKeys = new Set(rows0.filter(r => r.alloc.some(v => v > 0)).map(rkey)); // só perfis realmente alocados exigem taxa
  Object.values(rates).filter(r => usedKeys.has(rkey(r))).forEach(r => {
    const ok = p.mode === 'rate' ? num(r.rate) > 0 : num(r.cost) > 0;
    if (!ok) missing.push(`${r.role}${r.level?' '+r.level:''}${r.location?' ('+r.location+')':''}`);
  });

  const inScope = r => !r.optional || p.inclOpt;
  // 1) preço de lista por linha (taxa × horas × contingência), sem desconto
  const base = rows0.map((r, k) => {
    const hrs = r.alloc.map(v => v * hpp), H = hrs.reduce((a, b) => a + b, 0);
    return {r, k, hrs, H, pb: priceH(r), listNet: H * priceH(r) * priceF, cost: H * costH(r)};
  });
  // 2) desconto efetivo: % (geral + exceções por fase), valor fixo ou preço-alvo (ambos COM impostos → convertidos para líquido)
  const scopeList = base.filter(x => inScope(x.r)).reduce((a, x) => a + x.listNet, 0);
  let dGlobal = p.disc, targetAboveList = false;
  if (p.discMode === 'amount') dGlobal = scopeList > 0 ? p.discAmount * (1 - p.tax) / scopeList : 0;
  else if (p.discMode === 'target' && p.discTarget > 0){
    dGlobal = scopeList > 0 ? 1 - (p.discTarget * (1 - p.tax)) / scopeList : 0;
    if (dGlobal < 0){ dGlobal = 0; targetAboveList = true; }
  }
  dGlobal = Math.min(0.99, Math.max(0, dGlobal));
  const discOf = r => (p.discMode === 'pct' && p.discPhase[r.phase] !== undefined) ? p.discPhase[r.phase] : dGlobal;
  const rows = base.map(x => {
    const d = discOf(x.r), f = x.pb * priceF * (1 - d);
    return {...x.r, srcIndex: idxRows[x.k].i, hrs: x.hrs, H: x.H, priceBase: x.pb, rateList: x.pb * priceF, discPct: d, listNet: x.listNet,
      rateH: f, net: x.H * f, cost: x.cost, netP: x.hrs.map(h => h * f), costP: x.hrs.map(h => h * costH(x.r))};
  });
  const PALETTE = ['#1B4F8C','#13A89E','#6C4AB6','#E07A1F','#2E86C1','#C0392B','#7D8C2E','#AD5A8C','#4A6B82','#D4A017'];

  const phaseNames = []; rows0.forEach(r => { if (!phaseNames.includes(r.phase) && rows0.some(x => x.phase === r.phase && x.alloc.some(v => v > 0))) phaseNames.push(r.phase); }); // fase sem alocação não entra na proposta
  const phases = phaseNames.map((name, i) => {
    const rr = rows.filter(r => r.phase === name);
    const act = new Array(N).fill(0); rr.forEach(r => r.alloc.forEach((v,j) => act[j] += v));
    let first = act.findIndex(v => v > 0); if (first < 0) first = 0;
    let last = act.length - 1 - [...act].reverse().findIndex(v => v > 0); if (last >= act.length || last < first) last = first;
    const H = rr.reduce((a,r)=>a+r.H,0), net = rr.reduce((a,r)=>a+r.net,0), cost = rr.reduce((a,r)=>a+r.cost,0);
    const optional = rr.every(r => r.optional);
    return {name, color: PALETTE[i % PALETTE.length], first, last, start: pStart(first), end: pEnd(last), H, net, gross: gross(net), cost, optional,
      partialOpt: !optional && rr.some(r => r.optional), fte: act, roles: [...new Set(rr.map(r => r.role))]};
  });

  const roleMap = {};
  rows.forEach(r => {
    const k = rkey(r) + '|' + (r.optional ? 'opt' : 'core');
    if (!roleMap[k]) roleMap[k] = {role:r.role, level:r.level, location:r.location, optional:r.optional, H:0, net:0, cost:0, rateH:r.rateH, rateList:r.rateList, act:new Array(N).fill(0)};
    const m = roleMap[k]; m.H += r.H; m.net += r.net; m.cost += r.cost; r.alloc.forEach((v,j) => m.act[j] += v);
  });
  const roles = Object.values(roleMap).filter(m => m.H > 0).map(m => {
    const activeP = m.act.filter(v => v > 0).length;
    return {...m, gross: gross(m.net), avgFte: activeP ? m.act.reduce((a,b)=>a+b,0)/activeP : 0, peak: Math.max(...m.act), periods: activeP};
  }).sort((a,b) => (a.optional - b.optional) || b.net - a.net);

  const scRows = rows.filter(inScope);
  const periods = [];
  for (let j = 0; j < N; j++){
    periods.push({idx:j, start: pStart(j), end: pEnd(j),
      fte: scRows.reduce((a,r)=>a+r.alloc[j],0), fteAll: rows.reduce((a,r)=>a+r.alloc[j],0),
      H: scRows.reduce((a,r)=>a+r.hrs[j],0), net: scRows.reduce((a,r)=>a+r.netP[j],0), cost: scRows.reduce((a,r)=>a+r.costP[j],0)});
  }

  // distribuição mensal: cada período é rateado pelos seus dias úteis
  const months = {};
  periods.forEach(pd => {
    let days = [];
    for (let d = pd.start; d <= pd.end; d = addDays(d, 1)) if (isWork(d)) days.push(d);
    if (!days.length) days = [pd.start];
    days.forEach(dt => {
      const k = monthKey(dt); if (!months[k]) months[k] = {key:k, H:0, net:0, cost:0};
      months[k].H += pd.H/days.length; months[k].net += pd.net/days.length; months[k].cost += pd.cost/days.length;
    });
  });
  const monthArr = Object.values(months).sort((a,b) => a.key.localeCompare(b.key)).map((m, i) => ({...m, idx: i+1, gross: gross(m.net)}));

  const sum = (arr, k) => arr.reduce((a,x)=>a+x[k],0);
  const core = rows.filter(r => !r.optional), opt = rows.filter(r => r.optional);
  const T = {H: sum(scRows,'H'), net: sum(scRows,'net'), cost: sum(scRows,'cost'),
    Hcore: sum(core,'H'), netCore: sum(core,'net'), Hopt: sum(opt,'H'), netOpt: sum(opt,'net')};
  T.gross = gross(T.net); T.grossCore = gross(T.netCore); T.grossOpt = gross(T.netOpt);
  T.listNet = sum(scRows, 'listNet'); T.discNet = T.listNet - T.net; T.listGross = gross(T.listNet); T.discGross = gross(T.discNet);
  T.discPct = T.listNet ? T.discNet / T.listNet : 0;
  // desconto máximo (uniforme) que ainda preserva o GM mínimo (ou o ponto de equilíbrio, se não houver mínimo)
  const gmFloor = Math.max(0, p.gmMin);
  const maxDisc = (T.listNet > 0 && T.cost > 0) ? Math.min(1, Math.max(0, 1 - T.cost / (T.listNet * (1 - gmFloor)))) : null;
  T.gm = T.net ? (T.net - T.cost)/T.net : 0;
  const activeP = periods.filter(x => x.fte > 0);
  T.peak = Math.max(0, ...periods.map(x => x.fte));
  T.avgFte = activeP.length ? sum(activeP,'fte')/activeP.length : 0;
  const scPhases = phases.filter(ph => !ph.optional || p.inclOpt);
  const firstP = periods.findIndex(x => x.fte > 0);
  const lastP = firstP < 0 ? -1 : periods.length - 1 - [...periods].reverse().findIndex(x => x.fte > 0);
  if (firstP < 0){ T.start = effStart; T.end = effStart; T.nPer = 0; T.weeks = 0; }
  else { T.start = pStart(firstP); T.end = pEnd(lastP); T.nPer = lastP - firstP + 1; T.weeks = T.nPer * pw; }
  T.blended = T.H ? T.net / T.H : 0;

  // pagamentos
  let pays = [];
  const eom = d => new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth()+1, 0));
  if (p.pay === 'milestone'){
    scPhases.filter(ph => ph.net > 0).sort((a,b) => a.last - b.last || a.first - b.first).forEach(ph => {
      const netIn = rows.filter(r => r.phase === ph.name && inScope(r)).reduce((a,r)=>a+r.net,0);
      pays.push({label: ph.name, trigger:'phase', when: ph.end, net: netIn, period: ph.last});
    });
  } else if (p.pay === 'monthly'){
    monthArr.forEach(m => pays.push({label: m.key, trigger:'month', when: prevWork(eom(parseDate(m.key + '-01'))), net: m.net, idx: m.idx}));
  } else {
    for (let i = 0; i < p.inst; i++) pays.push({label: i+1, trigger:'inst', when: prevWork(eom(new Date(Date.UTC(T.start.getUTCFullYear(), T.start.getUTCMonth()+i, 1)))), net: T.net/p.inst, idx: i+1});
  }
  pays = pays.map(x => ({...x, gross: gross(x.net), pct: T.net ? x.net/T.net : 0, due: nextWork(addDays(x.when, p.terms))}));

  // alertas (códigos neutros de idioma; o app converte em texto)
  const issues = [];
  if (missing.length) issues.push({level:'err', code:'missing_rates', data:{list: missing, mode: p.mode}});
  if (T.cost > 0){
    if (T.gm < 0) issues.push({level:'err', code:'gm_negative', data:{gm: T.gm}});
    else if (p.gmMin > 0 && T.gm < p.gmMin) issues.push({level:'warn', code:'gm_below_min', data:{gm: T.gm, min: p.gmMin}});
  } else if (!missing.length) issues.push({level:'info', code:'no_cost', data:{}});
  const below = roles.filter(r => !r.optional && r.cost > 0 && r.net > 0 && r.net < r.cost).map(r => r.role + (r.level ? ' ' + r.level : ''));
  if (below.length) issues.push({level:'warn', code:'price_below_cost', data:{list: below}});
  if (startAdjusted) issues.push({level:'warn', code:'start_adjusted', data:{from: iso(baseDate), to: iso(effStart)}});
  if (relative) issues.push({level:'info', code:'relative_timeline', data:{}});
  if (p.cur !== p.rcur && p.fx === 1) issues.push({level:'warn', code:'fx_missing', data:{from: p.rcur, to: p.cur}});
  if (targetAboveList) issues.push({level:'info', code:'target_above_list', data:{list: T.listGross, target: p.discTarget}});
  if (T.netOpt && !p.inclOpt) issues.push({level:'info', code:'optional_excluded', data:{gross: T.grossOpt}});
  if (firstP < 0) issues.push({level:'err', code:'no_active_scope', data:{}});

  return {p, relative, unit, nPeriods: N, pw, hpp, f: priceF, maxDisc, gmFloor, dGlobal, rows, rowByIndex: new Map(rows.map(r => [r.srcIndex, r])), phases, roles, periods, months: monthArr, T, pays, missing, issues,
    start: effStart, startBase: baseDate, startAdjusted, holidays: [...hol].sort().map(d => ({date:d, name: holNames[d] || ''})), isWork};
}

/* ---------------- sensibilidade (desconto x contingência) ---------------- */
function sensitivity(model, p, ctx, discs, conts){
  return discs.map(d => ({disc: d, cells: conts.map(c => {
    const R = compute(model, {...p, discMode: 'pct', discPhase: {}, disc: d, cont: c}, ctx);
    return {cont: c, gross: R ? R.T.gross : 0, gm: R ? R.T.gm : 0, hasCost: !!(R && R.T.cost)};
  })}));
}

/* ---------------- visão de timeline (sprint/semana, com/sem FTE, barras/perfis) ---------------- */
function timelineView(R, o){
  o = o || {};
  const U = R.unit, pw = R.pw;
  const expand = (o.gran === 'week' && U === 'SP' && pw > 1) ? pw : 1;
  const cols = [];
  R.periods.forEach((pd, j) => {
    for (let k = 0; k < expand; k++){
      cols.push(U === 'SP'
        ? {label: expand > 1 ? 'W' + (j*pw + k + 1) : 'SP' + (j+1), group: 'SP' + (j+1), period: j}
        : {label: 'W' + (j+1), group: null, period: j});
    }
  });
  const spans = (cells) => { const f = cells.findIndex(v => v > 0); if (f < 0) return {first:-1,last:-1}; return {first:f, last: cells.length - 1 - [...cells].reverse().findIndex(v => v > 0)}; };
  const rows = [];
  R.phases.forEach(ph => {
    const cells = cols.map(c => ph.fte[c.period] || 0);
    rows.push({kind:'phase', label: ph.name, optional: ph.optional, color: ph.color, cells, ...spans(cells)});
    if (o.mode === 'roles'){
      R.rows.filter(r => r.phase === ph.name).forEach(r => {
        const rc = cols.map(c => r.alloc[c.period] || 0);
        rows.push({kind:'role', label: r.role + (r.level ? ' ' + r.level : ''), optional: r.optional, color: ph.color, cells: rc, ...spans(rc)});
      });
    }
  });
  return {cols, rows, showFte: o.fte !== false, groups: U === 'SP' && expand > 1};
}

/* ---------------- marcos ---------------- */
// devolve [{name, kind, when, periodLabel, order}] ordenado por data/ordem
function milestones(R, extraLines, labels){
  labels = labels || {kick:'Project kickoff', done:'completed'};
  const U = R.unit, out = [];
  const lab = (idx) => U + (idx + 1);
  if (R.T.nPer > 0){
    const firstIdx = R.periods.findIndex(x => x.fte > 0);
    out.push({name: labels.kick, kind:'kick', when: R.T.start, periodLabel: lab(firstIdx), order: firstIdx * 2});
  }
  R.phases.filter(ph => !ph.optional || R.p.inclOpt).forEach(ph => {
    out.push({name: `${ph.name} — ${labels.done}`, kind:'phase', when: ph.end, periodLabel: lab(ph.last), order: ph.last * 2 + 1, optional: false});
  });
  R.phases.filter(ph => ph.optional && !R.p.inclOpt).forEach(ph => {
    out.push({name: `${ph.name} — ${labels.done} *`, kind:'phase', when: ph.end, periodLabel: lab(ph.last), order: ph.last * 2 + 1, optional: true});
  });
  (extraLines || []).forEach(m => {
    if (m.date){
      const d = parseDate(m.date);
      const idx = R.periods.findIndex(x => d >= x.start && d <= addDays(x.end, 1));
      out.push({name: m.name, kind:'custom', when: d, periodLabel: idx >= 0 ? lab(idx) : '', order: (idx >= 0 ? idx : 0) * 2 + 1});
    } else if (m.n >= 1 && m.n <= R.periods.length){
      // SPn/Wn: se a unidade do pedido difere da do timeline, converte via semanas
      let idx = m.n - 1;
      if (m.unit !== U){
        const weekN = m.unit === 'W' ? m.n : (m.n - 1) * R.pw + 1; // semana de referência
        idx = U === 'SP' ? Math.min(R.periods.length - 1, Math.floor((weekN - 1) / R.pw)) : weekN - 1;
      }
      const endsAtPeriodEnd = m.unit === U || m.unit === 'SP';
      out.push({name: m.name, kind:'custom', when: endsAtPeriodEnd ? R.periods[idx].end : R.periods[idx].start, periodLabel: lab(idx), order: idx * 2 + 1});
    }
  });
  return out.sort((a,b) => a.order - b.order || a.when - b.when);
}

return { norm, esc, num, periodNum, isYes, rkey, parseDate, addDays, iso, monthKey, isValidIso, nextMonday,
  easter, holidaysBR, ALIAS, parseTimeline, parseRateRows, rateMatch, ensureRates, applyRateTable,
  parseParams, parseExtraHolidays, parseMilestoneLines, buildCalendar, compute, sensitivity, timelineView, milestones };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = Engine;

/* =====================================================================
   ARQUIVO DA PROPOSTA (.xlsx) — o Excel é o "banco" de cada proposta.
   O navegador é o editor; o Excel guarda tudo (parâmetros, timeline, rate card,
   impostos, textos editados) em abas legíveis, com fórmulas para auditoria.
   write: buildProposal(ExcelJS, st, R)   read: read(XLSX, workbook)
   Sem DOM: roda também no Node (build do template e testes).
===================================================================== */
const Workbook = (() => {
'use strict';
const E = Engine;
const APP_ID = 'ProposalMaker', FORMAT = 3;

// ordem e rótulos dos parâmetros gravados na aba "Parametros"
const PARAM_DEFS = [
  ['p_vendor','Fornecedor','text'], ['p_client','Cliente','text'], ['p_project','Projeto','text'], ['p_ref','Referência / CRM ID','text'],
  ['p_lang','Idioma da proposta (en/pt/es)','text'], ['p_model','Modelo de contratação (fp/tm)','text'],
  ['p_startmode','Início do projeto (date/month/none)','text'], ['p_start','Data de início (AAAA-MM-DD)','text'], ['p_startm','Mês de início (AAAA-MM)','text'],
  ['p_hol','Considerar feriados nacionais BR','bool'], ['p_holx','Feriados adicionais (AAAA-MM-DD [nome])','long'],
  ['p_cur','Moeda da proposta','text'], ['p_rcur','Moeda do rate card','text'], ['p_fx','Câmbio (1 moeda do rate card = X moeda da proposta)','num'],
  ['p_unit','Unidade do período (auto/SP/W)','text'], ['p_pw','Semanas por período (sprint)','num'], ['p_hpd','Horas faturáveis / dia','num'], ['p_dpw','Dias úteis / semana','num'],
  ['p_mode','Modo de precificação (rate/gm)','text'], ['p_gm','GM% alvo (modo custo)','num'], ['p_gmmin','GM% mínimo (alerta)','num'],
  ['p_cont','Contingência / risco %','num'], ['p_contmode','Contingência entra em (price/both)','text'],
  ['p_discmode','Tipo de desconto (pct/amount/target)','text'], ['p_disc','Desconto comercial % (modo pct)','num'], ['p_discamt','Desconto em valor fixo, com impostos (modo amount)','num'],
  ['p_disctarget','Preço final alvo, com impostos (modo target)','num'], ['p_discshow','Exibição do desconto (embedded/line)','text'],
  ['p_inclOpt','Incluir itens opcionais no total','bool'], ['p_taxshow','Detalhar impostos na proposta','bool'],
  ['p_pay','Modelo de pagamento (milestone/monthly/install)','text'], ['p_inst','Nº de parcelas (se iguais)','num'], ['p_terms','Prazo de pagamento (dias)','num'], ['p_valid','Validade da proposta (dias)','num'],
  ['p_tlmode','Timeline: modo (bars/roles/milestones)','text'], ['p_tlgran','Timeline: granularidade (period/week)','text'], ['p_tlfte','Timeline: mostrar FTE','bool'], ['p_tldeliv','Timeline: mostrar entregáveis','bool'],
  ['p_tlmiles','Timeline: marcos adicionais (Nome | SP6, W12 ou AAAA-MM-DD)','long'],
  ['p_context','Contexto / objetivo','long'], ['p_assump','Premissas adicionais (uma por linha)','long']
];
const BOOLS = new Set(PARAM_DEFS.filter(d => d[2] === 'bool').map(d => d[0]));

const colName = n => { let s = ''; while (n > 0){ const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

/* ---------- dados do exemplo / template ---------- */
function exampleData(){
  const N = 12, z = () => new Array(N).fill(0);
  const fill = (s, e, v) => { const a = z(); for (let i = s; i <= e; i++) a[i-1] = v; return a; };
  const rows = [
    ['Mobilization & Assessment','','Project Manager','Q4','Brazil','',fill(1,2,1)],
    ['Mobilization & Assessment','','Data Architect','Q4','Brazil','',fill(1,2,1)],
    ['Mobilization & Assessment','','Business Analyst','Q3','Brazil','',fill(1,2,1)],
    ['Ingestion Migration','','Data Engineer','Q3','Brazil','',fill(3,8,2)],
    ['Ingestion Migration','','Data Architect','Q4','Brazil','',fill(3,8,0.5)],
    ['Transformation (dbt / Medallion)','','Analytics Engineer','Q3','Brazil','',fill(5,10,2)],
    ['Transformation (dbt / Medallion)','','Data Engineer','Q3','Brazil','',fill(9,10,1)],
    ['Transformation (dbt / Medallion)','','Business Analyst','Q3','Brazil','',fill(5,10,0.5)],
    ['UAT & Go-Live','','Tester','Q2','Brazil','',fill(9,11,1)],
    ['UAT & Go-Live','','Data Engineer','Q3','Brazil','',fill(11,11,1)],
    ['Hypercare','','Data Engineer','Q3','Brazil','',fill(12,12,1)],
    ['Project Management','','Project Manager','Q4','Brazil','',fill(3,12,0.5)],
    ['Dashboards Squad','','Data Visualization Analyst','Q3','Brazil','Y',fill(11,12,2)]
  ];
  const rates = [
    ['Project Manager','Q4','Brazil',170,260], ['Data Architect','Q4','Brazil',185,280], ['Business Analyst','Q3','Brazil',125,190],
    ['Data Engineer','Q3','Brazil',130,200], ['Analytics Engineer','Q3','Brazil',130,200], ['Tester','Q2','Brazil',90,140],
    ['Data Visualization Analyst','Q3','Brazil',125,190]
  ];
  return {N, rows, rates};
}

const HEAD_STYLE = { font:{bold:true, color:{argb:'FFFFFFFF'}}, fill:{type:'pattern', pattern:'solid', fgColor:{argb:'FF0B2545'}}, alignment:{vertical:'middle'} };
function styleHeader(ws, rowNum = 1){ const r = ws.getRow(rowNum); r.eachCell(c => { c.font = HEAD_STYLE.font; c.fill = HEAD_STYLE.fill; c.alignment = HEAD_STYLE.alignment; }); r.height = 20; }
const money = '#,##0.00', num2 = '#,##0.00', pct = '0.0%';

/* =====================================================================
   TEMPLATE (única fonte: o botão "Template Excel" do app e o arquivo do repositório)
===================================================================== */
async function buildTemplate(ExcelJS){
  const ex = exampleData(), wb = new ExcelJS.Workbook();
  wb.creator = APP_ID; wb.created = new Date();

  const ws = wb.addWorksheet('Timeline', {views:[{state:'frozen', xSplit:6, ySplit:1}]});
  const hdr = ['Phase','Activity','Role','Level','Location','Optional', ...Array.from({length:ex.N}, (_, i) => 'SP' + (i+1))];
  ws.addRow(hdr);
  ex.rows.forEach(r => ws.addRow([r[0], r[1], r[2], r[3], r[4], r[5], ...r[6].map(v => v || null)]));
  const last = 1 + ex.rows.length;
  const tot = ws.addRow(['', '', 'Total FTE', '', '', '', ...Array.from({length:ex.N}, (_, i) => ({formula:`SUM(${colName(7+i)}2:${colName(7+i)}${last})`}))]);
  tot.font = {bold:true}; tot.eachCell(c => { c.fill = {type:'pattern', pattern:'solid', fgColor:{argb:'FFEEF3F9'}}; });
  styleHeader(ws);
  [30,14,26,8,12,9].forEach((w, i) => ws.getColumn(i+1).width = w);
  for (let i = 0; i < ex.N; i++) ws.getColumn(7+i).width = 6;
  const lastCol = colName(6 + ex.N);
  ws.dataValidations.add(`F2:F${last + 40}`, {type:'list', allowBlank:true, formulae:['"Y,N"'], showErrorMessage:true, errorTitle:'Optional', error:'Use Y (opcional) ou deixe vazio.'});
  ws.dataValidations.add(`G2:${lastCol}${last + 40}`, {type:'decimal', operator:'between', allowBlank:true, formulae:[0, 50], showErrorMessage:true, errorTitle:'FTE', error:'Informe a alocação em FTE (ex.: 1, 0,5).'});
  ws.addConditionalFormatting({ref:`G2:${lastCol}${last}`, rules:[{type:'colorScale', cfvo:[{type:'num', value:0}, {type:'num', value:2}], color:[{argb:'FFFFFFFF'}, {argb:'FF7FA8D8'}]}]});
  ws.getCell('A' + (last + 3)).value = 'Insira novas linhas ACIMA da linha "Total FTE" (a linha Total é ignorada na importação).';
  ws.getCell('A' + (last + 3)).font = {italic:true, color:{argb:'FF5F6B7A'}};

  const rc = wb.addWorksheet('Rate Card', {views:[{state:'frozen', ySplit:1}]});
  rc.addRow(['Role','Level','Location','Cost/h','Rate/h','Notes']);
  ex.rates.forEach(r => rc.addRow([...r, 'Valor ilustrativo — substitua pelo rate card oficial']));
  styleHeader(rc);
  [28,8,12,10,10,48].forEach((w, i) => rc.getColumn(i+1).width = w);
  rc.getColumn(4).numFmt = money; rc.getColumn(5).numFmt = money;

  const lf = wb.addWorksheet('Alt format (list)');
  lf.addRow(['Phase','Role','Level','Location','Start','End','FTE','Optional']);
  lf.addRow(['Assessment','Project Manager','Q4','Brazil','SP1','SP2',1,'']);
  lf.addRow(['Build','Data Engineer','Q3','Brazil','SP3','SP8',2,'']);
  styleHeader(lf); [18,22,8,12,8,8,8,10].forEach((w, i) => lf.getColumn(i+1).width = w);

  const ins = wb.addWorksheet('Instructions');
  ins.getColumn(1).width = 120;
  ['How to use / Como usar',
   '1. Timeline: uma linha por Fase + Perfil; FTE por sprint (SP1, SP2…). Use W1, W2… para planos semanais.',
   '2. Marque o escopo opcional com "Y" na coluna Optional.',
   '3. Rate Card: preencha Cost/h e/ou Rate/h (líquidos de impostos). A aba é importada automaticamente.',
   '4. Carregue este arquivo no app (aba 1). As propostas geradas pelo app são salvas em outro .xlsx, que o app também reabre.',
   'Formato alternativo: lista com Start / End / FTE (aba "Alt format (list)").'
  ].forEach((t, i) => { const r = ins.addRow([t]); if (!i) r.font = {bold:true, size:13}; });
  return wb;
}

// resultado "vazio" para gravar propostas que ainda não têm perfis alocados
function emptyResult(st){
  const p = E.parseParams(st.raw, {taxes: st.taxes});
  const cal = E.buildCalendar(p, st.nPeriods, st.unit, {today: new Date()});
  const z = new Date(cal.effStart);
  return {p, relative: cal.relative, unit: st.unit, nPeriods: st.nPeriods, pw: cal.pw, hpp: cal.pw * p.dpw * p.hpd, f: (1 + p.cont) * (1 - p.disc),
    rows: [], rowByIndex: new Map(), phases: [], roles: [], periods: [], months: [], pays: [], holidays: [], issues: [], missing: [],
    T: {H: 0, net: 0, cost: 0, gross: 0, gm: 0, blended: 0, grossOpt: 0, netOpt: 0, listNet: 0, discNet: 0, listGross: 0, discGross: 0, discPct: 0, nPer: 0, weeks: 0, peak: 0, avgFte: 0, start: z, end: z}};
}

/* =====================================================================
   ARQUIVO DA PROPOSTA — gravação
   st = { rows, nPeriods, unit, rates, taxes, edits, raw (param id -> valor), fileName }
   R  = resultado de Engine.compute (para valores em cache das fórmulas)
===================================================================== */
async function buildProposal(ExcelJS, st, R){
  const wb = new ExcelJS.Workbook();
  wb.creator = APP_ID; wb.created = new Date(); wb.modified = new Date();
  const p = R.p, N = st.nPeriods;

  /* ----- Parametros ----- */
  const ps = wb.addWorksheet('Parametros', {views:[{state:'frozen', ySplit:1}]});
  ps.addRow(['Chave','Descrição','Valor']);
  styleHeader(ps);
  const rowOf = {};
  PARAM_DEFS.forEach(([id, label, kind]) => {
    let v = st.raw[id];
    if (kind === 'bool') v = (v === true || v === 'true' || E.isYes(v)) ? 'Sim' : 'Não';
    else if (kind === 'num') v = (v === '' || v === undefined || v === null) ? null : E.num(v);
    else v = v === undefined || v === null ? '' : String(v);
    const r = ps.addRow([id, label, v]);
    rowOf[id] = r.number;
    if (kind === 'long') r.getCell(3).alignment = {wrapText:true, vertical:'top'};
  });
  const extraRow = (k, label, v) => { const r = ps.addRow([k, label, v]); rowOf[k] = r.number; return r; };
  extraRow('_unit', 'Unidade do timeline (SP/W) — gerado', st.unit);
  extraRow('_hpp', 'Horas por período por FTE — calculado', {formula:'IF(C' + rowOf._unit + '="W",1,C' + rowOf.p_pw + ')*C' + rowOf.p_dpw + '*C' + rowOf.p_hpd, result:R.hpp});
  extraRow('_tax', 'Impostos totais (%) — soma da aba Impostos', {formula:'SUM(Impostos!$B$2:$B$60)', result:p.tax * 100});
  ps.getColumn(1).width = 14; ps.getColumn(2).width = 62; ps.getColumn(3).width = 46;
  const nm = (name, key) => wb.definedNames.add(`Parametros!$C$${rowOf[key]}`, name);
  nm('PW','p_pw'); nm('DPW','p_dpw'); nm('HPD','p_hpd'); nm('MODE','p_mode'); nm('GMP','p_gm'); nm('CONTP','p_cont'); nm('CONTMODE','p_contmode');
  nm('FX','p_fx'); nm('INCLOPT','p_inclOpt'); nm('HPP','_hpp'); nm('TAXP','_tax');
  ps.addRow([]);
  const note = ps.addRow(['', 'Este arquivo é gravado e lido pelo app. A edição é feita no navegador; alterações feitas diretamente aqui podem ser sobrescritas no próximo salvamento.', '']);
  note.getCell(2).font = {italic:true, color:{argb:'FF5F6B7A'}}; note.getCell(2).alignment = {wrapText:true};

  /* ----- Impostos ----- */
  const ts = wb.addWorksheet('Impostos');
  ts.addRow(['Nome', '%']); styleHeader(ts);
  p.taxes.forEach(t => ts.addRow([t.name, t.pct]));
  ts.getColumn(1).width = 34; ts.getColumn(2).width = 10;

  /* ----- Descontos por fase ----- */
  const ds = wb.addWorksheet('Descontos');
  ds.addRow(['Fase','%']); styleHeader(ds);
  Object.entries(st.phaseDisc || {}).forEach(([k, v]) => ds.addRow([k, E.num(v)]));
  ds.getColumn(1).width = 34; ds.getColumn(2).width = 10;

  /* ----- Modelo (timeline + cálculo) ----- */
  const ms = wb.addWorksheet('Modelo', {views:[{state:'frozen', xSplit:7, ySplit:1}]});
  const U = st.unit;
  const c0 = 8, cCostS = c0 + N, cRateS = c0 + N + 1, cTot = c0 + N + 2, cHrs = cTot + 1, cPrice = cTot + 2, cDisc = cTot + 3, cNet = cTot + 4, cCost = cTot + 5, cScope = cTot + 6;
  ms.addRow(['Phase','Activity','Role','Level','Location','Optional','Key',
    ...Array.from({length:N}, (_, i) => U + (i+1)), 'Cost/h (catálogo)', 'Rate/h (catálogo)', 'Total FTE', 'Hours', 'Price/h (calc.)', 'Desconto efetivo (calc. pelo app)', 'Net (calc.)', 'Cost (calc.)', 'In scope']);
  styleHeader(ms);
  const n1 = 2, nLast = Math.max(2, st.rows.length + 1);
  st.rows.forEach((src, i) => {
    const r = R.rowByIndex.get(i) || {...src, alloc: src.alloc.slice(), H: 0, priceBase: 0, discPct: 0, net: 0, cost: 0}; // linha sem perfil: zeros
    const n = i + 2, first = colName(c0), lastc = colName(c0 + N - 1);
    const rt = (src.role && st.rates[E.rkey(src)]) || {};   // taxa do catálogo usada nesta linha (cópia para auditoria/recálculo)
    const cS = `${colName(cCostS)}${n}`, rS = `${colName(cRateS)}${n}`;
    ms.addRow([r.phase, r.activity, r.role, r.level || '', r.location || '', r.optional ? 'Y' : '',
      {formula:`C${n}&"|"&IF(D${n}="","-",D${n})&"|"&IF(E${n}="","-",E${n})`, result: E.rkey(r)},
      ...r.alloc.map(v => v || null),
      E.num(rt.cost) || null, E.num(rt.rate) || null,
      {formula:`SUM(${first}${n}:${lastc}${n})`, result: r.alloc.reduce((a,b)=>a+b,0)},
      {formula:`${colName(cTot)}${n}*HPP`, result: r.H},
      {formula:`IF(MODE="gm",IF(GMP<100,N(${cS})/(1-GMP/100),0),N(${rS}))*FX`, result: r.priceBase},
      r.discPct || 0,
      {formula:`${colName(cHrs)}${n}*${colName(cPrice)}${n}*(1+CONTP/100)*(1-${colName(cDisc)}${n})`, result: r.net},
      {formula:`${colName(cHrs)}${n}*N(${cS})*FX*IF(CONTMODE="both",1+CONTP/100,1)`, result: r.cost},
      {formula:`IF(OR(F${n}<>"Y",INCLOPT="Sim"),1,0)`, result: (!r.optional || p.inclOpt) ? 1 : 0}]);
  });
  [26,14,24,7,10,9,30].forEach((w, i) => ms.getColumn(i+1).width = w);
  for (let i = 0; i < N; i++) ms.getColumn(c0 + i).width = 6;
  [cTot, cHrs].forEach(c => ms.getColumn(c).numFmt = '#,##0.0');
  [cCostS, cRateS].forEach(c => { ms.getColumn(c).numFmt = money; ms.getColumn(c).width = 14; });
  [cPrice, cNet, cCost].forEach(c => { ms.getColumn(c).numFmt = money; ms.getColumn(c).width = 14; });
  ms.getColumn(cDisc).numFmt = pct; ms.getColumn(cDisc).width = 16;
  ms.addConditionalFormatting({ref:`${colName(c0)}2:${colName(c0 + N - 1)}${nLast}`, rules:[{type:'colorScale', cfvo:[{type:'num', value:0}, {type:'num', value:2}], color:[{argb:'FFFFFFFF'}, {argb:'FF7FA8D8'}]}]});
  const rng = c => `Modelo!$${colName(c)}$${n1}:$${colName(c)}$${nLast}`;

  /* ----- Resumo ----- */
  const sm = wb.addWorksheet('Resumo');
  sm.getColumn(1).width = 38; sm.getColumn(2).width = 22;
  const T = R.T;
  const add = (label, v, fmt) => { const r = sm.addRow([label, v]); if (fmt) r.getCell(2).numFmt = fmt; return r; };
  sm.addRow(['Proposta comercial — resumo']).font = {bold:true, size:14};
  sm.addRow([]);
  add('Cliente', p.client); add('Projeto', p.project); add('Fornecedor', p.vendor); add('Referência', p.ref);
  add('Moeda', p.cur); add('Salvo em', new Date().toISOString().slice(0,16).replace('T',' '));
  add('Início', R.relative ? '(relativo)' : E.iso(T.start)); add('Fim', R.relative ? '(relativo)' : E.iso(T.end));
  add('Duração (períodos)', T.nPer);
  sm.addRow([]);
  const sp = (col) => `SUMPRODUCT(${rng(cScope)},${rng(col)})`;
  const rH = add('Horas totais', {formula: sp(cHrs), result: T.H}, '#,##0');
  const rC = add('Custo total', {formula: sp(cCost), result: T.cost}, money);
  const rN = add('Preço líquido de impostos', {formula: sp(cNet), result: T.net}, money);
  add('Preço com impostos', {formula: `IF(TAXP<100,B${rN.number}/(1-TAXP/100),B${rN.number})`, result: T.gross}, money);
  add('GM%', {formula: `IF(B${rN.number}=0,0,(B${rN.number}-B${rC.number})/B${rN.number})`, result: T.gm}, pct);
  add('Rate médio (blended) /h', {formula: `IF(B${rH.number}=0,0,B${rN.number}/B${rH.number})`, result: T.blended}, money);
  add('Preço de lista (com impostos)', T.listGross, money);
  add('Desconto (com impostos)', T.discGross, money);
  add('Desconto efetivo %', T.discPct, pct);
  add('Opcionais (com impostos, fora do total)', T.grossOpt, money);

  /* ----- PorFase (fórmulas) ----- */
  const fs = wb.addWorksheet('PorFase');
  fs.addRow(['Phase','Opcional','Início','Fim','Horas','Líquido','Com impostos','Custo']); styleHeader(fs);
  R.phases.forEach((ph, i) => {
    const n = i + 2;
    const sif = col => `SUMIFS(${rng(col)},${rng(1)},$A${n},${rng(cScope)},1)`;
    fs.addRow([ph.name, ph.optional ? 'Y' : '', R.relative ? '' : E.iso(ph.start), R.relative ? '' : E.iso(ph.end),
      {formula: sif(cHrs), result: R.rows.filter(r => r.phase === ph.name && (!r.optional || p.inclOpt)).reduce((a,r)=>a+r.H,0)},
      {formula: sif(cNet), result: R.rows.filter(r => r.phase === ph.name && (!r.optional || p.inclOpt)).reduce((a,r)=>a+r.net,0)},
      {formula: `IF(TAXP<100,F${n}/(1-TAXP/100),F${n})`, result: ph.optional && !p.inclOpt ? 0 : (p.tax < 1 ? R.rows.filter(r => r.phase === ph.name && (!r.optional || p.inclOpt)).reduce((a,r)=>a+r.net,0) / (1 - p.tax) : 0)},
      {formula: sif(cCost), result: R.rows.filter(r => r.phase === ph.name && (!r.optional || p.inclOpt)).reduce((a,r)=>a+r.cost,0)}]);
  });
  [34,9,12,12,10,16,16,16].forEach((w, i) => fs.getColumn(i+1).width = w);
  [6,7,8].forEach(c => fs.getColumn(c).numFmt = money); fs.getColumn(5).numFmt = '#,##0';

  /* ----- abas de referência (valores) ----- */
  const rp = wb.addWorksheet('PorPerfil');
  rp.addRow(['Role','Level','Location','Opcional','FTE médio','Pico','Horas','Rate/h aplicado','Líquido','Com impostos','Custo']); styleHeader(rp);
  R.roles.forEach(r => rp.addRow([r.role, r.level, r.location, r.optional ? 'Y' : '', +r.avgFte.toFixed(3), r.peak, +r.H.toFixed(2), +r.rateH.toFixed(2), +r.net.toFixed(2), +r.gross.toFixed(2), +r.cost.toFixed(2)]));
  [26,8,12,9,10,8,10,14,14,14,14].forEach((w, i) => rp.getColumn(i+1).width = w);

  const mn = wb.addWorksheet('Mensal');
  mn.addRow([R.relative ? 'Mês (relativo)' : 'Mês','Horas','Líquido','Com impostos','Custo']); styleHeader(mn);
  R.months.forEach(m => mn.addRow([R.relative ? 'M' + m.idx : m.key, +m.H.toFixed(2), +m.net.toFixed(2), +m.gross.toFixed(2), +m.cost.toFixed(2)]));
  [14,10,14,14,14].forEach((w, i) => mn.getColumn(i+1).width = w);

  const pg = wb.addWorksheet('Pagamentos');
  pg.addRow(['#','Evento','Faturamento','Vencimento','%','Líquido','Com impostos']); styleHeader(pg);
  R.pays.forEach((x, i) => pg.addRow([i+1, String(x.label), R.relative ? '' : E.iso(x.when), R.relative ? '' : E.iso(x.due), +(x.pct*100).toFixed(2), +x.net.toFixed(2), +x.gross.toFixed(2)]));
  [5,36,14,14,8,14,14].forEach((w, i) => pg.getColumn(i+1).width = w);

  if (R.holidays.length){
    const hs = wb.addWorksheet('Feriados');
    hs.addRow(['Data','Nome']); styleHeader(hs);
    R.holidays.forEach(h => hs.addRow([h.date, h.name])); hs.getColumn(1).width = 12; hs.getColumn(2).width = 34;
  }

  /* ----- Entregáveis e cenários ----- */
  const scs = st.scenarios && st.scenarios.length ? st.scenarios : null;
  const activeId = scs ? scs[st.active || 0].id : 1;
  const dl = [];   // [cenário, fase, entregável, início(1-based), fim(1-based), manual]
  const dlist = scs ? scs.map((sc, i) => ({id: sc.id, list: i === (st.active || 0) ? (st.deliverables || []) : (sc.deliverables || [])})) : [{id: 1, list: st.deliverables || []}];
  dlist.forEach(x => x.list.forEach(d => dl.push([x.id, d.phase, d.name, (d.start | 0) + 1, (d.end | 0) + 1, d.manual ? 'Y' : ''])));
  if (dl.length){
    const dsh = wb.addWorksheet('Entregaveis', {views:[{state:'frozen', ySplit:1}]});
    dsh.addRow(['Cenário','Fase','Entregável','Início (período)','Fim (período)','Manual']); styleHeader(dsh);
    dl.forEach(r => dsh.addRow(r));
    [10,28,40,16,14,9].forEach((w, i) => dsh.getColumn(i+1).width = w);
  }
  if (scs && scs.length > 1){
    const cs = wb.addWorksheet('Cenarios', {views:[{state:'frozen', ySplit:1}]});
    cs.addRow(['Id','Nome','Ativo','Unidade','Períodos','p_unit','p_pw','p_hpd','p_dpw','p_startmode','p_start','p_startm','Aba']); styleHeader(cs);
    scs.forEach((sc, i) => {
      const act = i === (st.active || 0), tl = sc.tl || {};
      cs.addRow([sc.id, sc.name, act ? 'Y' : '', act ? st.unit : sc.unit, act ? st.nPeriods : sc.nPeriods, ...['p_unit','p_pw','p_hpd','p_dpw','p_startmode','p_start','p_startm'].map(k => tl[k] === undefined ? '' : tl[k]), act ? 'Modelo' : 'Cen_' + sc.id]);
      if (act) return;
      const ws = wb.addWorksheet('Cen_' + sc.id, {views:[{state:'frozen', xSplit:6, ySplit:1}]});
      ws.addRow(['Phase','Activity','Role','Level','Location','Optional', ...Array.from({length: sc.nPeriods}, (_, k) => sc.unit + (k+1)), 'Cost/h (catálogo)', 'Rate/h (catálogo)']); styleHeader(ws);
      sc.rows.forEach(r => { const rt = (r.role && st.rates[E.rkey(r)]) || {}; ws.addRow([r.phase, r.activity, r.role, r.level || '', r.location || '', r.optional ? 'Y' : '', ...r.alloc.slice(0, sc.nPeriods).map(v => v || null), E.num(rt.cost) || null, E.num(rt.rate) || null]); });
      [26,14,24,7,10,9].forEach((w, k) => ws.getColumn(k+1).width = w);
    });
    [6,22,7,9,10,10,8,8,8,12,12,10,10].forEach((w, i) => cs.getColumn(i+1).width = w);
  }

  /* ----- Textos editados ----- */
  const xs = wb.addWorksheet('Textos');
  xs.addRow(['Chave','Texto (HTML simples)']); styleHeader(xs);
  Object.entries(st.edits || {}).forEach(([k, v]) => xs.addRow([k, v]));
  xs.getColumn(1).width = 24; xs.getColumn(2).width = 100;

  /* ----- meta (oculta) ----- */
  const meta = wb.addWorksheet('_meta', {state:'hidden'});
  meta.addRow([APP_ID, FORMAT, new Date().toISOString(), st.fileName || '', st.catalogName || '']);

  wb.calcProperties = {fullCalcOnLoad: true};
  return wb;
}

/* =====================================================================
   ARQUIVO DA PROPOSTA — leitura (SheetJS)
===================================================================== */
function isProposalWorkbook(XLSX, wb){
  const m = wb.Sheets && wb.Sheets['_meta'];
  if (!m) return false;
  const a1 = m['A1'] && m['A1'].v;
  return a1 === APP_ID;
}
function read(XLSX, wb){
  const aoa = name => wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:'', raw:true, blankrows:false}) : null;
  const par = aoa('Parametros');
  if (!par) throw new Error('aba "Parametros" não encontrada — este não é um arquivo de proposta do app');
  const kv = {};
  par.slice(1).forEach(r => { if (r[0]) kv[String(r[0])] = r[2]; });
  const raw = {};
  PARAM_DEFS.forEach(([id,, kind]) => {
    let v = kv[id];
    if (v === undefined) return;
    if (kind === 'bool') raw[id] = E.isYes(v);
    else if (kind === 'num') raw[id] = v === '' ? '' : E.num(v);
    else raw[id] = v === null || v === undefined ? '' : String(v);
  });
  const unit = kv['_unit'] === 'W' ? 'W' : 'SP';

  const taxRows = aoa('Impostos');
  const taxes = taxRows ? taxRows.slice(1).filter(r => r[0] !== '' || r[1] !== '').map(r => ({name: String(r[0] || 'Imposto'), pct: E.num(r[1])})) : [];

  const modelo = aoa('Modelo');
  if (!modelo) throw new Error('aba "Modelo" não encontrada');
  const t = E.parseTimeline(modelo, {unit, keepEmpty: true});

  // taxas usadas por linha (cópia gravada na aba Modelo); arquivos antigos tinham a aba Perfis/RateCard
  const rates = {};
  const rcName = wb.SheetNames.find(n => /^(perfis|rate ?card)$/i.test(n));
  if (rcName) E.parseRateRows(aoa(rcName)).forEach(t => { rates[E.rkey(t)] = {role: t.role, level: t.level, location: t.location, cost: t.cost || '', rate: t.rate || ''}; });
  const applySnap = (sheet, parsed) => {
    const hdr = sheet.find(r => r.some(c => E.norm(c) === 'role')) || [];
    const iCost = hdr.findIndex(c => E.norm(c) === 'cost/h (catalogo)'), iRate = hdr.findIndex(c => E.norm(c) === 'rate/h (catalogo)');
    if (iCost >= 0 || iRate >= 0) parsed.rows.forEach(r => {
      if (!r.role || r.srcLine === undefined) return;
      const line = sheet[r.srcLine] || [], c = iCost >= 0 ? E.num(line[iCost]) : 0, rt = iRate >= 0 ? E.num(line[iRate]) : 0, k = E.rkey(r);
      if (!rates[k] || c || rt) rates[k] = {role: r.role, level: r.level, location: r.location, cost: c || (rates[k] && rates[k].cost) || '', rate: rt || (rates[k] && rates[k].rate) || ''};
    });
  };
  applySnap(modelo, t);
  E.ensureRates(rates, t.rows);

  const phaseDisc = {};
  const dsr = aoa('Descontos');
  if (dsr) dsr.slice(1).forEach(r => { const v = E.num(r[1]); if (r[0] && v > 0) phaseDisc[String(r[0])] = v; });

  const edits = {};
  const tx = aoa('Textos');
  if (tx) tx.slice(1).forEach(r => { if (r[0]) edits[String(r[0])] = String(r[1] ?? ''); });

  // cenários (opcional) e entregáveis
  let scenarios = null, active = 0, deliverables = [];
  const dsh = aoa('Entregaveis'), allD = [];
  if (dsh) dsh.slice(1).forEach((r, i) => { if (r[1] !== '' && r[2] !== '') allD.push({sc: Number(r[0]) || 1, d: {id: i + 1, phase: String(r[1]), name: String(r[2]), start: Math.max(0, (E.num(r[3]) || 1) - 1), end: Math.max(0, (E.num(r[4]) || 1) - 1), manual: E.isYes(r[5])}}); });
  const csr = aoa('Cenarios');
  if (csr && csr.length > 1){
    const h = csr[0].map(x => String(x));
    scenarios = csr.slice(1).filter(r => r[0] !== '').map((r, i) => {
      const o = {}; h.forEach((k, j) => { o[k] = r[j]; });
      const id = Number(o.Id) || i + 1, isAct = E.isYes(o.Ativo);
      const tl = {}; ['p_unit','p_pw','p_hpd','p_dpw','p_startmode','p_start','p_startm'].forEach(k => { if (o[k] !== '' && o[k] !== undefined) tl[k] = String(o[k] instanceof Date ? E.iso(o[k]) : o[k]); });
      const sc = {id, name: String(o.Nome || 'Cenário ' + (i + 1)), tl, deliverables: allD.filter(x => x.sc === id).map(x => x.d)};
      if (isAct){ active = i; sc.rows = t.rows; sc.nPeriods = t.nPeriods; sc.unit = t.unit; deliverables = sc.deliverables; }
      else {
        const m = aoa('Cen_' + id);
        let pt = null; try { pt = m ? E.parseTimeline(m, {unit: o.Unidade === 'W' ? 'W' : 'SP', keepEmpty: true}) : null; } catch(e){ pt = null; }
        if (pt){ applySnap(m, pt); sc.rows = pt.rows; sc.nPeriods = pt.nPeriods; sc.unit = pt.unit; E.ensureRates(rates, pt.rows); }
        else { sc.rows = []; sc.nPeriods = 12; sc.unit = 'SP'; }
      }
      return sc;
    });
  } else deliverables = allD.map(x => x.d);

  return {scenarios, active, deliverables, raw, taxes, phaseDisc, catalogName: (wb.Sheets['_meta'] && wb.Sheets['_meta']['E1']) ? String(wb.Sheets['_meta']['E1'].v || '') : '', rows: t.rows, nPeriods: t.nPeriods, unit: t.unit, rates, edits, warnings: t.warnings};
}

return { APP_ID, FORMAT, PARAM_DEFS, BOOLS, colName, exampleData, emptyResult, buildTemplate, buildProposal, isProposalWorkbook, read };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = Workbook;

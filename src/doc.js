/* =====================================================================
   DOCUMENTO DA PROPOSTA — modelo de blocos (puro, sem DOM).
   build(R, edits) -> {lang, blocks}. A mesma lista de blocos alimenta a tela/PDF
   (toHtml) e o Word (.docx, em docx-export.js).
===================================================================== */
const DocModel = (() => {
'use strict';
const E = Engine;
const LOC = {en:'en-US', pt:'pt-BR', es:'es-ES'};

function mkFmt(lang, cur){
  const loc = LOC[lang] || 'en-US';
  const nf = (v, dec=0) => new Intl.NumberFormat(loc, {minimumFractionDigits:dec, maximumFractionDigits:dec}).format(v);
  const money = (v, dec=0) => { try { return new Intl.NumberFormat(loc, {style:'currency', currency:cur, minimumFractionDigits:dec, maximumFractionDigits:dec}).format(v); } catch(e){ return cur + ' ' + nf(v, dec); } };
  const date = d => d.toLocaleDateString(loc, {timeZone:'UTC', day:'2-digit', month:'short', year:'numeric'});
  const month = k => E.parseDate(k + '-01').toLocaleDateString(loc, {timeZone:'UTC', month:'short', year:'numeric'});
  return {nf, money, date, month, loc};
}

const I18N = {
  en: {
    tag:'Commercial Proposal', prepared:'Prepared for', date:'Date', ref:'Reference', validity:'Valid until',
    s1:'Executive Summary', s2:'Delivery Timeline', s3:'Team Composition', s4:'Investment', s5:'Payment Schedule', s6:'Commercial Assumptions', s7:'Optional Scope',
    duration:'Duration', peak:'Peak team', effort:'Total effort', invest:'Total investment',
    phase:'Phase', period:'Sprints', dates:'Dates', hours:'Hours', net:'Net of taxes', gross:'Taxes included', share:'% of total', total:'Total',
    role:'Role', level:'Level', loc:'Location', avgFte:'Avg. FTE', rateH:'Rate / hour', amount:'Amount',
    event:'Billing event', invDate:'Invoice date', due:'Due date',
    sp:'Sprint', wk:'Week', weeks:'weeks', sprints:'sprints', fte:'FTE', weekWord:'Week', monthWord:'Month',
    milestone:'Milestone', target:'Target', endOf:'End of', plusDays:d=>`+${d} days`, kick:'Project kickoff', done:'completed',
    model:{fp:'Fixed Price', tm:'Time & Material'},
    payTxt:{milestone:'Invoices are issued upon completion and acceptance of each phase.', monthly:'Invoices are issued monthly based on the effort consumed in the period.', install:'The total investment is invoiced in equal monthly installments.'},
    trig:{phase:'Completion of', month:'Monthly measurement', inst:'Installment'},
    monthlyBilling:'Monthly billing (taxes included)', subtotal:'Subtotal (list price)', discountRow:'Commercial discount',
    taxes:'Taxes', taxLine:(list,tot)=>`Taxes included in the price: ${list} (total ${tot}).`,
    relNote:'The timeline is expressed in relative periods; calendar dates will be defined at project kickoff.',
    exec:(p,T,U,f,rel)=>`${p.vendor} is pleased to present this commercial proposal to ${p.client} for the <b>${p.project}</b> initiative. The engagement is planned over <b>${T.nPer} ${U==='SP'?'sprints':'weeks'} (${T.weeks} weeks)</b>${rel?'':`, from ${f.date(T.start)} to ${f.date(T.end)}`}, with a peak team of ${f.nf(T.peak,1)} FTE and a total effort of ${f.nf(T.H)} hours.`,
    execInv:(p,T,f)=>`The total investment is <b>${f.money(T.gross)}</b> (taxes included), under a <b>${I18N.en.model[p.model]}</b> commercial model.`,
    optTxt:(T,f)=>`The items below are not included in the base investment and can be activated upon formal approval. Additional investment: <b>${f.money(T.grossOpt)}</b> (taxes included).`,
    optIncl:(T,f)=>`Optional items are included in the total investment above, representing ${f.money(T.grossOpt)} (taxes included).`,
    a:(p,R,f)=>[
      `Commercial model: ${I18N.en.model[p.model]}. Prices in ${p.cur}${p.tax?`, including local taxes (${f.nf(p.tax*100,2)}% gross-up)`:''}.`,
      `Effort estimated on ${f.nf(p.hpd,1)} billable hours per day, ${p.dpw} working days per week${R.unit==='SP'?` and ${R.pw}-week sprints`:''}.`,
      R.relative ? `Project start date to be agreed; the timeline is expressed in relative ${R.unit==='SP'?'sprints':'weeks'}.` : `Project start on ${f.date(R.start)}; any change in the start date may require re-planning of the timeline.`,
      `Payment terms: ${p.terms} days from invoice date.`,
      `Scope, timeline and investment are based on the information available at the time of this proposal; changes will be handled through a formal change request process.`,
      `This proposal is valid for ${p.valid} days.`
    ],
    foot:p=>`${p.vendor} · Commercial Proposal · ${p.client} — ${p.project}`
  },
  pt: {
    tag:'Proposta Comercial', prepared:'Preparado para', date:'Data', ref:'Referência', validity:'Válida até',
    s1:'Sumário Executivo', s2:'Cronograma de Entrega', s3:'Composição do Time', s4:'Investimento', s5:'Cronograma de Pagamentos', s6:'Premissas Comerciais', s7:'Escopo Opcional',
    duration:'Duração', peak:'Pico do time', effort:'Esforço total', invest:'Investimento total',
    phase:'Fase', period:'Sprints', dates:'Datas', hours:'Horas', net:'Líquido de impostos', gross:'Com impostos', share:'% do total', total:'Total',
    role:'Perfil', level:'Nível', loc:'Local', avgFte:'FTE médio', rateH:'Valor / hora', amount:'Valor',
    event:'Evento de faturamento', invDate:'Faturamento', due:'Vencimento',
    sp:'Sprint', wk:'Semana', weeks:'semanas', sprints:'sprints', fte:'FTE', weekWord:'Semana', monthWord:'Mês',
    milestone:'Marco', target:'Previsão', endOf:'Fim do', plusDays:d=>`+${d} dias`, kick:'Início do projeto', done:'concluída',
    model:{fp:'Preço Fixo', tm:'Time & Material'},
    payTxt:{milestone:'O faturamento ocorre na conclusão e aceite de cada fase.', monthly:'O faturamento é mensal, conforme o esforço consumido no período.', install:'O investimento total é faturado em parcelas mensais iguais.'},
    trig:{phase:'Conclusão de', month:'Medição mensal', inst:'Parcela'},
    monthlyBilling:'Faturamento mensal (com impostos)', subtotal:'Subtotal (preço de lista)', discountRow:'Desconto comercial',
    taxes:'Impostos', taxLine:(list,tot)=>`Impostos incluídos no preço: ${list} (total ${tot}).`,
    relNote:'O cronograma está expresso em períodos relativos; as datas de calendário serão definidas no início do projeto.',
    exec:(p,T,U,f,rel)=>`A ${p.vendor} tem a satisfação de apresentar à ${p.client} esta proposta comercial para a iniciativa <b>${p.project}</b>. O projeto está planejado em <b>${T.nPer} ${U==='SP'?'sprints':'semanas'} (${T.weeks} semanas)</b>${rel?'':`, de ${f.date(T.start)} a ${f.date(T.end)}`}, com pico de ${f.nf(T.peak,1)} FTE e esforço total de ${f.nf(T.H)} horas.`,
    execInv:(p,T,f)=>`O investimento total é de <b>${f.money(T.gross)}</b> (impostos inclusos), no modelo comercial <b>${I18N.pt.model[p.model]}</b>.`,
    optTxt:(T,f)=>`Os itens abaixo não estão incluídos no investimento base e podem ser ativados mediante aprovação formal. Investimento adicional: <b>${f.money(T.grossOpt)}</b> (impostos inclusos).`,
    optIncl:(T,f)=>`Os itens opcionais estão incluídos no investimento total acima, representando ${f.money(T.grossOpt)} (impostos inclusos).`,
    a:(p,R,f)=>[
      `Modelo comercial: ${I18N.pt.model[p.model]}. Valores em ${p.cur}${p.tax?`, com impostos locais inclusos (gross-up de ${f.nf(p.tax*100,2)}%)`:''}.`,
      `Esforço estimado com base em ${f.nf(p.hpd,1)} horas faturáveis por dia, ${p.dpw} dias úteis por semana${R.unit==='SP'?` e sprints de ${R.pw} semanas`:''}.`,
      R.relative ? `Data de início a combinar; o cronograma está expresso em ${R.unit==='SP'?'sprints':'semanas'} relativas.` : `Início do projeto em ${f.date(R.start)}; alterações na data de início podem exigir replanejamento do cronograma.`,
      `Prazo de pagamento: ${p.terms} dias a partir da emissão da nota fiscal.`,
      `Escopo, prazo e investimento baseiam-se nas informações disponíveis nesta data; mudanças serão tratadas via processo formal de change request.`,
      `Esta proposta é válida por ${p.valid} dias.`
    ],
    foot:p=>`${p.vendor} · Proposta Comercial · ${p.client} — ${p.project}`
  },
  es: {
    tag:'Propuesta Comercial', prepared:'Preparado para', date:'Fecha', ref:'Referencia', validity:'Válida hasta',
    s1:'Resumen Ejecutivo', s2:'Cronograma de Entrega', s3:'Composición del Equipo', s4:'Inversión', s5:'Cronograma de Pagos', s6:'Supuestos Comerciales', s7:'Alcance Opcional',
    duration:'Duración', peak:'Pico del equipo', effort:'Esfuerzo total', invest:'Inversión total',
    phase:'Fase', period:'Sprints', dates:'Fechas', hours:'Horas', net:'Neto de impuestos', gross:'Impuestos incluidos', share:'% del total', total:'Total',
    role:'Rol', level:'Nivel', loc:'Ubicación', avgFte:'FTE promedio', rateH:'Tarifa / hora', amount:'Monto',
    event:'Evento de facturación', invDate:'Facturación', due:'Vencimiento',
    sp:'Sprint', wk:'Semana', weeks:'semanas', sprints:'sprints', fte:'FTE', weekWord:'Semana', monthWord:'Mes',
    milestone:'Hito', target:'Previsión', endOf:'Fin de', plusDays:d=>`+${d} días`, kick:'Inicio del proyecto', done:'concluida',
    model:{fp:'Precio Fijo', tm:'Time & Material'},
    payTxt:{milestone:'La facturación se realiza al concluir y aceptar cada fase.', monthly:'La facturación es mensual, según el esfuerzo consumido en el período.', install:'La inversión total se factura en cuotas mensuales iguales.'},
    trig:{phase:'Conclusión de', month:'Medición mensual', inst:'Cuota'},
    monthlyBilling:'Facturación mensual (impuestos incluidos)', subtotal:'Subtotal (precio de lista)', discountRow:'Descuento comercial',
    taxes:'Impuestos', taxLine:(list,tot)=>`Impuestos incluidos en el precio: ${list} (total ${tot}).`,
    relNote:'El cronograma se expresa en períodos relativos; las fechas de calendario se definirán al inicio del proyecto.',
    exec:(p,T,U,f,rel)=>`${p.vendor} se complace en presentar a ${p.client} esta propuesta comercial para la iniciativa <b>${p.project}</b>. El proyecto está planificado en <b>${T.nPer} ${U==='SP'?'sprints':'semanas'} (${T.weeks} semanas)</b>${rel?'':`, del ${f.date(T.start)} al ${f.date(T.end)}`}, con un pico de ${f.nf(T.peak,1)} FTE y un esfuerzo total de ${f.nf(T.H)} horas.`,
    execInv:(p,T,f)=>`La inversión total es de <b>${f.money(T.gross)}</b> (impuestos incluidos), bajo el modelo comercial <b>${I18N.es.model[p.model]}</b>.`,
    optTxt:(T,f)=>`Los ítems a continuación no están incluidos en la inversión base y pueden activarse mediante aprobación formal. Inversión adicional: <b>${f.money(T.grossOpt)}</b> (impuestos incluidos).`,
    optIncl:(T,f)=>`Los ítems opcionales están incluidos en la inversión total, representando ${f.money(T.grossOpt)} (impuestos incluidos).`,
    a:(p,R,f)=>[
      `Modelo comercial: ${I18N.es.model[p.model]}. Valores en ${p.cur}${p.tax?`, con impuestos locales incluidos (gross-up de ${f.nf(p.tax*100,2)}%)`:''}.`,
      `Esfuerzo estimado con base en ${f.nf(p.hpd,1)} horas facturables por día, ${p.dpw} días hábiles por semana${R.unit==='SP'?` y sprints de ${R.pw} semanas`:''}.`,
      R.relative ? `Fecha de inicio por acordar; el cronograma se expresa en ${R.unit==='SP'?'sprints':'semanas'} relativas.` : `Inicio del proyecto el ${f.date(R.start)}; cambios en la fecha de inicio pueden requerir replanificación.`,
      `Plazo de pago: ${p.terms} días a partir de la fecha de factura.`,
      `El alcance, plazo e inversión se basan en la información disponible a la fecha; los cambios se gestionarán mediante un proceso formal de control de cambios.`,
      `Esta propuesta es válida por ${p.valid} días.`
    ],
    foot:p=>`${p.vendor} · Propuesta Comercial · ${p.client} — ${p.project}`
  }
};

// agrupa perfis iguais (núcleo + opcional) para a tabela de time
function mergeRoles(R){
  const map = new Map();
  R.roles.filter(r => !r.optional || R.p.inclOpt).forEach(r => {
    const k = E.rkey(r);
    if (!map.has(k)) map.set(k, {role:r.role, level:r.level, location:r.location, H:0, net:0, rateH:r.rateH, rateList:r.rateList, act:new Array(R.nPeriods).fill(0)});
    const m = map.get(k); m.H += r.H; m.net += r.net; r.act.forEach((v,j) => m.act[j] += v);
  });
  return [...map.values()].map(m => { const a = m.act.filter(v => v > 0).length; return {...m, avgFte: a ? m.act.reduce((x,y)=>x+y,0)/a : 0}; });
}

function build(R, opts){
  opts = opts || {};
  const p = R.p, T = R.T, L = I18N[p.lang], U = R.unit, rel = R.relative;
  const f = mkFmt(p.lang, p.cur);
  const today = opts.today || new Date();
  const tdy = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()));
  const gross = v => p.tax < 1 ? v / (1 - p.tax) : v;
  const per = ph => ph.first === ph.last ? `${U}${ph.first+1}` : `${U}${ph.first+1}–${U}${ph.last+1}`;
  const inScopePh = R.phases.filter(ph => !ph.optional || p.inclOpt);
  const optPh = R.phases.filter(ph => ph.optional);
  const weekOf = idx => (U === 'SP' ? (idx + 1) * R.pw : idx + 1); // semana em que o período termina
  const B = [];

  B.push({t:'cover', tag:L.tag, title:p.project, client:p.client, titleKey:'cover.title', clientKey:'cover.client',
    meta: [{k:L.prepared, v:p.client}, {k:L.date, v:f.date(tdy)}, {k:L.validity, v:f.date(E.addDays(tdy, p.valid))}].concat(p.ref ? [{k:L.ref, v:p.ref}] : [])});

  // 1. Resumo
  B.push({t:'h', n:1, text:L.s1});
  if (p.context) B.push({t:'p', key:'exec.ctx', html:E.esc(p.context)});
  const pe = {...p, vendor:E.esc(p.vendor), client:E.esc(p.client), project:E.esc(p.project)};
  B.push({t:'p', key:'exec.main', html:`${L.exec(pe,T,U,f,rel)} ${L.execInv(pe,T,f)}`});
  B.push({t:'kpis', items:[
    {l:L.duration, v:`${T.nPer} ${U==='SP'?L.sprints:L.weeks}`}, {l:L.peak, v:`${f.nf(T.peak,1)} ${L.fte}`},
    {l:L.effort, v:`${f.nf(T.H)} h`}, {l:L.invest, v:f.money(T.gross)}]});

  // 2. Timeline
  B.push({t:'h', n:2, text:L.s2});
  const tl = p.tl;
  if (rel) B.push({t:'note', text:L.relNote});
  if (tl.mode === 'milestones'){
    const ms = E.milestones(R, tl.miles, {kick:L.kick, done:L.done});
    const head = [{text:L.milestone}, {text:U==='SP'?L.sp:L.wk}, {text:L.wk}].concat(rel ? [] : [{text:L.target}]);
    const rows = ms.map(m => {
      const idx = parseInt((m.periodLabel.match(/\d+/)||[1])[0], 10) - 1;
      const wk = m.kind === 'kick' ? 1 : weekOf(idx);
      return {cells:[m.name, m.periodLabel, `${L.weekWord} ${wk}`].concat(rel ? [] : [f.date(m.when)]), cls: m.optional ? 'opt' : ''};
    });
    B.push({t:'table', head, rows, isTimeline:true});
  } else {
    const view = E.timelineView(R, {gran: tl.gran, mode: tl.mode, fte: tl.fte});
    B.push({t:'timeline', view});
    const head = [{text:L.phase}, {text:U==='SP'?L.period:L.weeks}].concat(rel ? [] : [{text:L.dates}]).concat(tl.fte ? [{text:L.fte, num:true}] : []);
    const rows = R.phases.map(ph => {
      const act = ph.fte.filter(v => v > 0), avg = act.length ? act.reduce((a,b)=>a+b,0)/act.length : 0;
      return {cells:[ph.name + (ph.optional ? ' *' : ''), per(ph)].concat(rel ? [] : [`${f.date(ph.start)} – ${f.date(ph.end)}`]).concat(tl.fte ? [f.nf(avg,1)] : []), cls: ph.optional ? 'opt' : ''};
    });
    B.push({t:'table', head, rows});
  }
  if (optPh.length) B.push({t:'note', text:`* ${L.s7}`});

  // 3. Time
  B.push({t:'h', n:3, text:L.s3});
  const team = mergeRoles(R);
  B.push({t:'table', head:[{text:L.role},{text:L.level},{text:L.loc},{text:L.avgFte,num:true},{text:L.hours,num:true}],
    rows: team.map(r => ({cells:[r.role, r.level, r.location, f.nf(r.avgFte,2), f.nf(r.H)]})),
    total: {cells:[L.total,'','','',f.nf(T.H)]}});

  // 4. Investimento
  B.push({t:'h', n:4, text:L.s4});
  const showDisc = p.discShow === 'line' && T.discNet > 0.005;   // lista → desconto → preço final
  const invRows = [];
  inScopePh.forEach(ph => {
    const rr = R.rows.filter(r => r.phase === ph.name && (!r.optional || p.inclOpt));
    const H = rr.reduce((a,r)=>a+r.H,0), net = showDisc ? rr.reduce((a,r)=>a+r.listNet,0) : rr.reduce((a,r)=>a+r.net,0);
    if (!H) return;
    const base = showDisc ? T.listNet : T.net;
    invRows.push({cells:[ph.name, f.nf(H), f.money(net), f.money(gross(net)), `${base ? f.nf(net/base*100,1) : 0}%`]});
  });
  if (showDisc){
    invRows.push({cells:[L.subtotal, f.nf(T.H), f.money(T.listNet), f.money(T.listGross), '100%'], cls:'sub'});
    invRows.push({cells:[`${L.discountRow} (${f.nf(T.discPct*100,1)}%)`, '', '−' + f.money(T.discNet), '−' + f.money(T.discGross), ''], cls:'disc'});
  }
  B.push({t:'table', head:[{text:L.phase},{text:L.hours,num:true},{text:L.net,num:true},{text:L.gross,num:true},{text:L.share,num:true}],
    rows: invRows, total: {cells:[L.total, f.nf(T.H), f.money(T.net), f.money(T.gross), showDisc ? '' : '100%']}});
  if (p.taxShow && p.taxes.length){
    B.push({t:'note', text: L.taxLine(p.taxes.map(t => `${t.name} ${f.nf(t.pct,2)}%`).join(', '), f.nf(p.tax*100,2) + '%')});
  }
  if (p.model === 'tm'){
    B.push({t:'sub', text:L.rateH});
    B.push({t:'table', head:[{text:L.role},{text:L.level},{text:L.loc},{text:`${L.rateH} (${L.net})`,num:true},{text:`${L.rateH} (${L.gross})`,num:true}],
      rows: team.map(r => { const rate = showDisc ? r.rateList : r.rateH; return {cells:[r.role, r.level, r.location, f.money(rate,2), f.money(gross(rate),2)]}; })});
  }

  // 5. Pagamentos
  B.push({t:'h', n:5, text:L.s5});
  B.push({t:'p', key:'pay.txt', html:E.esc(L.payTxt[p.pay])});
  const payHead = [{text:'#'},{text:L.event}].concat(rel ? [{text:L.invDate},{text:L.due}] : [{text:L.invDate},{text:L.due}]).concat([{text:'%',num:true},{text:L.gross,num:true}]);
  const payRows = R.pays.map((x, i) => {
    let ev, when, due;
    if (x.trigger === 'phase') ev = `${L.trig.phase} ${x.label}`;
    else if (x.trigger === 'month') ev = `${L.trig.month} – ${rel ? `${L.monthWord} ${x.idx}` : f.month(x.label)}`;
    else ev = `${L.trig.inst} ${x.label}/${p.inst}`;
    if (rel){
      when = x.trigger === 'phase' ? `${L.endOf} ${U}${x.period + 1}` : `${L.endOf} ${L.monthWord.toLowerCase()} ${x.idx}`;
      due = L.plusDays(p.terms);
    } else { when = f.date(x.when); due = f.date(x.due); }
    return {cells:[String(i+1), ev, when, due, `${f.nf(x.pct*100,1)}%`, f.money(x.gross)]};
  });
  B.push({t:'table', head: payHead, rows: payRows, total: {cells:[L.total,'','','','',f.money(T.gross)], span:5}});

  // 6. Opcionais
  let sec = 6;
  if (optPh.length || T.netOpt){
    B.push({t:'h', n:sec++, text:L.s7});
    B.push({t:'p', key:'opt.txt', html: p.inclOpt ? L.optIncl(T,f) : L.optTxt(T,f)});
    const rows = [];
    R.phases.forEach(ph => {
      const rr = R.rows.filter(r => r.phase === ph.name && r.optional); if (!rr.length) return;
      const H = rr.reduce((a,r)=>a+r.H,0), net = rr.reduce((a,r)=>a+r.net,0);
      rows.push({cells:[ph.name, per(ph), f.nf(H), f.money(gross(net))]});
    });
    B.push({t:'table', head:[{text:L.phase},{text:L.period},{text:L.hours,num:true},{text:L.gross,num:true}], rows});
  }

  // 7. Premissas
  B.push({t:'h', n:sec, text:L.s6});
  B.push({t:'ul', key:'assump', items:[...L.a(p,R,f), ...p.assump].map(a => E.esc(a))});
  B.push({t:'foot', text:L.foot(p)});
  return {lang:p.lang, blocks:B, fmt:f, L};
}

/* ---------------- sanitização de textos editados ---------------- */
// mantém só <b>, <strong>, <i>, <em>, <u>, <br>, <li>, <ul>, <p>; remove atributos e o resto
function sanitizeHtml(html){
  const allowed = new Set(['B','STRONG','I','EM','U','BR','LI','UL','P']);
  if (typeof document === 'undefined') return E.esc(String(html).replace(/<[^>]*>/g, ''));
  const tpl = document.createElement('template');
  tpl.innerHTML = String(html ?? '');
  const walk = node => {
    [...node.childNodes].forEach(ch => {
      if (ch.nodeType === 3) return;
      if (ch.nodeType !== 1){ ch.remove(); return; }
      if (!allowed.has(ch.tagName)){ // descarta a tag mas mantém o texto
        walk(ch);
        const frag = document.createDocumentFragment();
        while (ch.firstChild) frag.appendChild(ch.firstChild);
        ch.replaceWith(frag);
        return;
      }
      [...ch.attributes].forEach(a => ch.removeAttribute(a.name));
      walk(ch);
    });
  };
  walk(tpl.content);
  return tpl.innerHTML;
}

/* ---------------- HTML (tela e impressão) ---------------- */
function toHtml(doc, edits){
  const e = E.esc;
  const ed = (key, dflt) => { const v = edits && edits[doc.lang + ':' + key]; return v !== undefined && v !== null ? v : dflt; };
  let out = '', open = false;
  const page = () => { if (!open){ out += '<div class="page">'; open = true; } };
  for (const b of doc.blocks){
    switch (b.t){
      case 'cover':
        out += `<div class="cover"><div class="tag">${e(b.tag)}</div>
          <h1 contenteditable data-ek="${b.titleKey}">${ed(b.titleKey, e(b.title))}</h1>
          <h2 contenteditable data-ek="${b.clientKey}">${ed(b.clientKey, e(b.client))}</h2>
          <div class="meta">${b.meta.map(m => `<div><span>${e(m.k)}</span>${e(m.v)}</div>`).join('')}</div></div>`;
        break;
      case 'h': page(); out += `<h3 class="sec">${b.n}. ${e(b.text)}</h3>`; break;
      case 'p': page(); out += `<p contenteditable data-ek="${b.key}">${ed(b.key, b.html)}</p>`; break;
      case 'kpis': out += `<div class="hl">${b.items.map(i => `<div><span>${e(i.l)}</span><b>${e(i.v)}</b></div>`).join('')}</div>`; break;
      case 'note': out += `<p class="note">${e(b.text)}</p>`; break;
      case 'sub': out += `<h4 class="sub">${e(b.text)}</h4>`; break;
      case 'ul': out += `<ul contenteditable data-ek="${b.key}">${ed(b.key, b.items.map(i => `<li>${i}</li>`).join(''))}</ul>`; break;
      case 'foot': out += `<div class="foot">${e(b.text)}</div>`; break;
      case 'table': {
        out += `<table class="t"><thead><tr>${b.head.map(h => `<th class="${h.num?'n':''}">${e(h.text)}</th>`).join('')}</tr></thead><tbody>`;
        b.rows.forEach(r => out += `<tr class="${r.cls||''}">${r.cells.map((c,i) => `<td class="${b.head[i] && b.head[i].num ? 'n' : ''}">${e(c)}</td>`).join('')}</tr>`);
        if (b.total){
          const cs = b.total.cells, span = b.total.span;
          if (span) out += `<tr class="tot"><td colspan="${span}">${e(cs[0])}</td><td class="n">${e(cs[cs.length-1])}</td></tr>`;
          else out += `<tr class="tot">${cs.map((c,i) => `<td class="${b.head[i] && b.head[i].num ? 'n' : ''}">${e(c)}</td>`).join('')}</tr>`;
        }
        out += '</tbody></table>';
        break;
      }
      case 'timeline': {
        const v = b.view, n = v.cols.length;
        let h = `<table class="tlt" style="--n:${n}"><thead>`;
        if (v.groups){
          h += '<tr><th class="lab"></th>';
          let i = 0; while (i < n){ let j = i; while (j < n && v.cols[j].group === v.cols[i].group) j++; h += `<th colspan="${j-i}" class="grp">${e(v.cols[i].group)}</th>`; i = j; }
          h += '</tr>';
        }
        h += `<tr><th class="lab"></th>${v.cols.map(c => `<th>${e(c.label)}</th>`).join('')}</tr></thead><tbody>`;
        v.rows.forEach(r => {
          h += `<tr class="${r.kind}${r.optional ? ' optrow' : ''}"><td class="lab">${e(r.label)}${r.optional && r.kind==='phase' ? ' *' : ''}</td>`;
          r.cells.forEach(val => {
            const on = val > 0;
            const style = on ? `background:${r.color}${r.kind==='role' ? 'AA' : ''};${r.optional ? 'background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.35) 0 3px,transparent 3px 6px);' : ''}` : '';
            h += `<td class="${on ? 'on' : ''}" style="${style}">${on && v.showFte ? e(String(Math.round(val*10)/10).replace('.', doc.lang==='en' ? '.' : ',')) : ''}</td>`;
          });
          h += '</tr>';
        });
        out += h + '</tbody></table>';
        break;
      }
    }
  }
  if (open) out += '</div>';
  return out;
}

return { I18N, mkFmt, build, toHtml, sanitizeHtml, mergeRoles };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = DocModel;

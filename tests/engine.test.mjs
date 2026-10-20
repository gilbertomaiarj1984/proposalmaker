import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const Engine = new Function(fs.readFileSync(path.join(here, '../src/engine.js'), 'utf8') + '\n;return Engine;')();

const TODAY = new Date(2026, 9, 19); // 2026-10-19 (segunda-feira)
const mkModel = (extra = {}) => {
  const z = n => new Array(n).fill(0);
  const fill = (s, e, v, n = 6) => { const a = z(n); for (let i = s; i <= e; i++) a[i-1] = v; return a; };
  const rows = [
    {phase:'Assessment', activity:'', role:'PM', level:'Q4', location:'Brazil', optional:false, alloc: fill(1,2,1)},
    {phase:'Build', activity:'', role:'Dev', level:'Q3', location:'Brazil', optional:false, alloc: fill(3,5,2)},
    {phase:'Extra', activity:'', role:'Dev', level:'Q3', location:'Brazil', optional:true, alloc: fill(6,6,1)}
  ];
  const rates = {};
  rows.forEach(r => { rates[Engine.rkey(r)] = {role:r.role, level:r.level, location:r.location, cost:100, rate:200}; });
  return {rows, nPeriods: 6, unit: 'SP', rates, ...extra};
};
const params = (over = {}) => Engine.parseParams({p_start:'2026-11-02', p_pw:2, p_hpd:8, p_dpw:5, p_tax:0, p_hol:false, ...over});

test('num entende formatos brasileiros e percentuais', () => {
  assert.equal(Engine.num('1.234,5'), 1234.5);
  assert.equal(Engine.num('50%'), 0.5);
  assert.equal(Engine.num('0,5'), 0.5);
  assert.equal(Engine.num(''), 0);
});

test('esc escapa aspas simples e HTML', () => {
  assert.equal(Engine.esc(`<b onclick='x'>"`), '&lt;b onclick=&#39;x&#39;&gt;&quot;');
});

test('feriados BR: Páscoa 2026 e feriados móveis', () => {
  assert.equal(Engine.iso(Engine.easter(2026)), '2026-04-05');
  const h = Object.fromEntries(Engine.holidaysBR(2026).map(x => [x.name + x.date, 1]));
  assert.ok(h['Sexta-feira Santa2026-04-03']);
  assert.ok(h['Corpus Christi2026-06-04']);
  assert.ok(h['Consciência Negra2026-11-20']);
  assert.ok(!Engine.holidaysBR(2023).some(x => x.name === 'Consciência Negra'));
});

test('parseTimeline: matriz e lista; ignora linha Total', () => {
  const aoa = [['Phase','Role','Level','Location','Optional','SP1','SP2'], ['A','PM','Q4','BR','','1',0.5], ['','Total FTE','','','',1,0.5]];
  const t = Engine.parseTimeline(aoa);
  assert.equal(t.rows.length, 1); assert.equal(t.nPeriods, 2); assert.equal(t.unit, 'SP');
  assert.deepEqual(t.rows[0].alloc, [1, 0.5]);
  const l = Engine.parseTimeline([['Phase','Role','Level','Start','End','FTE'], ['A','PM','Q4','W2','W4',2]]);
  assert.equal(l.unit, 'W'); assert.deepEqual(l.rows[0].alloc, [0,2,2,2]);
  const forced = Engine.parseTimeline([['Phase','Role','Start','End'], ['A','PM',1,3]], {unit:'W'});
  assert.equal(forced.unit, 'W');
});

test('rate card é aplicado mesmo depois de ensureRates (bug de ordem de importação)', () => {
  const rows = [{phase:'A', role:'PM', level:'Q4', location:'Brazil', alloc:[1]}];
  const rates = Engine.ensureRates({}, rows);
  const table = Engine.parseRateRows([['Role','Level','Location','Cost/h','Rate/h'], ['PM','Q4','Brazil',170,260]]);
  assert.equal(Engine.applyRateTable(rates, table), 1);
  assert.equal(rates['PM|Q4|Brazil'].rate, 260);
});

test('horas e preço básicos', () => {
  const R = Engine.compute(mkModel(), params(), {today: TODAY});
  // 2*1*80 + 3*2*80 = 640h no escopo base; hpp = 2*5*8 = 80
  assert.equal(R.hpp, 80);
  assert.equal(R.T.H, 640);
  assert.equal(R.T.net, 640 * 200);
  assert.equal(R.T.cost, 640 * 100);
  assert.ok(Math.abs(R.T.gm - 0.5) < 1e-9);
});

test('opcionais entram só quando marcados', () => {
  const m = mkModel();
  assert.equal(Engine.compute(m, params(), {}).T.H, 640);
  assert.equal(Engine.compute(m, params({p_inclOpt:true}), {}).T.H, 720);
});

test('gross-up de impostos múltiplos', () => {
  const p = Engine.parseParams({p_start:'2026-11-02', p_hol:false}, {taxes:[{name:'ISS',pct:2},{name:'PIS/COFINS',pct:3.65}]});
  assert.ok(Math.abs(p.tax - 0.0565) < 1e-12);
  const R = Engine.compute(mkModel(), p, {});
  assert.ok(Math.abs(R.T.gross - R.T.net / (1 - 0.0565)) < 1e-6);
});

test('desconto e contingência; contingência em "both" não altera o GM%', () => {
  const m = mkModel();
  const base = Engine.compute(m, params(), {});
  const disc = Engine.compute(m, params({p_disc:10}), {});
  assert.ok(Math.abs(disc.T.net - base.T.net * 0.9) < 1e-6);
  const contPrice = Engine.compute(m, params({p_cont:10}), {});
  assert.ok(contPrice.T.gm > base.T.gm);
  const contBoth = Engine.compute(m, params({p_cont:10, p_contmode:'both'}), {});
  assert.ok(Math.abs(contBoth.T.gm - base.T.gm) < 1e-9);
});

test('modo custo + GM% e câmbio', () => {
  const m = mkModel();
  const R = Engine.compute(m, params({p_mode:'gm', p_gm:50}), {});
  assert.ok(Math.abs(R.T.net - 640 * 200) < 1e-6); // 100/(1-0,5) = 200
  const fx = Engine.compute(m, params({p_fx:0.2, p_cur:'USD', p_rcur:'BRL'}), {});
  assert.ok(Math.abs(fx.T.net - 640 * 200 * 0.2) < 1e-6);
  assert.ok(!fx.issues.some(i => i.code === 'fx_missing'));
  assert.ok(Engine.compute(m, params({p_cur:'USD', p_rcur:'BRL'}), {}).issues.some(i => i.code === 'fx_missing'));
});

test('perfis sem taxa geram erro visível e preço zero', () => {
  const m = mkModel(); Object.values(m.rates).forEach(r => { r.rate = ''; });
  const R = Engine.compute(m, params(), {});
  assert.equal(R.T.net, 0);
  assert.ok(R.issues.some(i => i.code === 'missing_rates' && i.level === 'err'));
});

test('GM% abaixo do mínimo e negativo', () => {
  const m = mkModel();
  assert.ok(Engine.compute(m, params({p_gmmin:60}), {}).issues.some(i => i.code === 'gm_below_min'));
  Object.values(m.rates).forEach(r => { r.rate = 50; });
  assert.ok(Engine.compute(m, params(), {}).issues.some(i => i.code === 'gm_negative'));
});

test('datas: sprint de 2 semanas começando na segunda termina na sexta', () => {
  const R = Engine.compute(mkModel(), params(), {});
  assert.equal(Engine.iso(R.periods[0].start), '2026-11-02');
  assert.equal(Engine.iso(R.periods[0].end), '2026-11-13');
  assert.equal(Engine.iso(R.periods[1].start), '2026-11-16');
});

test('datas: início numa quarta termina numa terça (dia útil), não no domingo', () => {
  const R = Engine.compute(mkModel(), params({p_start:'2026-11-04'}), {});
  assert.equal(Engine.iso(R.periods[0].start), '2026-11-04');
  assert.equal(Engine.iso(R.periods[0].end), '2026-11-17');
});

test('datas: início em fim de semana avança para o próximo dia útil e avisa', () => {
  const R = Engine.compute(mkModel(), params({p_start:'2026-11-07'}), {});
  assert.equal(Engine.iso(R.start), '2026-11-09');
  assert.ok(R.startAdjusted && R.issues.some(i => i.code === 'start_adjusted'));
});

test('feriados: início em feriado avança; fim de sprint recua para dia útil', () => {
  const R = Engine.compute(mkModel(), params({p_start:'2026-11-20', p_hol:true}), {});
  assert.equal(Engine.iso(R.start), '2026-11-23');
  const R2 = Engine.compute(mkModel(), params({p_start:'2026-11-06', p_hol:true}), {});
  // 2026-11-20 (Consciência Negra) cai no fim do 1º sprint (06/11 + 14 - 1 = 19/11 → quinta)
  assert.equal(Engine.iso(R2.periods[0].end), '2026-11-19');
  const R3 = Engine.compute(mkModel(), params({p_start:'2026-11-09', p_hol:true}), {});
  assert.equal(Engine.iso(R3.periods[0].end), '2026-11-19'); // sexta 20/11 é feriado
});

test('feriados extras manuais', () => {
  const p = params({p_start:'2026-12-14', p_hol:true, p_holx:'2026-12-14 Recesso\nlixo'});
  assert.equal(p.holx.length, 1);
  assert.equal(Engine.iso(Engine.compute(mkModel(), p, {}).start), '2026-12-15');
});

test('modo mês: começa no primeiro dia útil do mês', () => {
  const R = Engine.compute(mkModel(), params({p_startmode:'month', p_startm:'2027-01', p_hol:true}), {});
  assert.equal(Engine.iso(R.start), '2027-01-04'); // 01/01 feriado, 02-03 fim de semana
});

test('modo sem data: cronograma relativo, sem ajuste de início', () => {
  const R = Engine.compute(mkModel(), params({p_startmode:'none'}), {today: TODAY});
  assert.equal(R.relative, true);
  assert.equal(Engine.iso(R.start), '2026-10-26');
  assert.ok(R.issues.some(i => i.code === 'relative_timeline'));
  assert.equal(R.months[0].idx, 1);
});

test('todos os itens opcionais e excluídos: sem escopo, sem datas absurdas', () => {
  const m = mkModel(); m.rows.forEach(r => { r.optional = true; });
  const R = Engine.compute(m, params(), {});
  assert.equal(R.T.nPer, 0);
  assert.equal(Engine.iso(R.T.end), Engine.iso(R.T.start));
  assert.ok(R.issues.some(i => i.code === 'no_active_scope'));
});

test('pagamentos somam o total e vencimento cai em dia útil', () => {
  for (const pay of ['milestone','monthly','install']) {
    const R = Engine.compute(mkModel(), params({p_pay:pay, p_inst:3, p_hol:true}), {});
    const tot = R.pays.reduce((a,x) => a + x.net, 0);
    assert.ok(Math.abs(tot - R.T.net) < 1e-6, pay);
    R.pays.forEach(x => assert.ok(R.isWork(x.due), pay + ' ' + Engine.iso(x.due)));
  }
});

test('distribuição mensal soma o total', () => {
  const R = Engine.compute(mkModel(), params({p_hol:true}), {});
  const tot = R.months.reduce((a,m) => a + m.net, 0);
  assert.ok(Math.abs(tot - R.T.net) < 1e-6);
});

test('timelineView: sprint vs semana, modo perfis', () => {
  const R = Engine.compute(mkModel(), params(), {});
  const sp = Engine.timelineView(R, {gran:'period', mode:'bars', fte:true});
  assert.equal(sp.cols.length, 6); assert.equal(sp.cols[0].label, 'SP1');
  const wk = Engine.timelineView(R, {gran:'week', mode:'bars', fte:false});
  assert.equal(wk.cols.length, 12); assert.equal(wk.cols[3].label, 'W4'); assert.equal(wk.cols[3].group, 'SP2');
  assert.equal(wk.showFte, false); assert.equal(wk.groups, true);
  const roles = Engine.timelineView(R, {gran:'period', mode:'roles'});
  assert.equal(roles.rows.length, 3 + 3); // 3 fases + 3 linhas de perfil
  assert.equal(roles.rows[0].first, 0); assert.equal(roles.rows[0].last, 1);
});

test('marcos: kickoff, fim de fase e marcos extras (SPn, Wn, data)', () => {
  const R = Engine.compute(mkModel(), params(), {});
  const extra = Engine.parseMilestoneLines('Go-live | SP5\nCorte | W3\nAssinatura | 2026-11-10\nruim');
  assert.equal(extra.length, 3);
  const ms = Engine.milestones(R, extra);
  assert.equal(ms[0].kind, 'kick');
  assert.ok(ms.some(m => m.name === 'Assessment — completed'));
  assert.ok(ms.find(m => m.name === 'Go-live').periodLabel === 'SP5');
  assert.ok(ms.find(m => m.name === 'Corte').periodLabel === 'SP2');
  assert.ok(ms.some(m => m.name === 'Extra — completed *')); // opcional fora do escopo
});

test('sensibilidade: preço cai com desconto', () => {
  const g = Engine.sensitivity(mkModel(), params(), {}, [0, 0.1], [0]);
  assert.ok(g[1].cells[0].gross < g[0].cells[0].gross);
});

test('parseParams é tolerante a lixo', () => {
  const p = Engine.parseParams({p_pw:'abc', p_dpw:'99', p_gm:'200', p_startmode:'x', p_lang:'xx'});
  assert.equal(p.pw, 2); assert.equal(p.dpw, 7); assert.ok(p.gm <= 0.95); assert.equal(p.startMode, 'date'); assert.equal(p.lang, 'en');
});

test('desconto em %: aplicado sobre o preço de lista; contingência vem antes', () => {
  const m = mkModel();
  const R = Engine.compute(m, params({p_cont:10, p_disc:10}), {});
  assert.ok(Math.abs(R.T.listNet - 640 * 200 * 1.10) < 1e-6);
  assert.ok(Math.abs(R.T.net - R.T.listNet * 0.9) < 1e-6);
  assert.ok(Math.abs(R.T.discPct - 0.10) < 1e-9);
});

test('desconto por fase: exceção ignora o geral e entra no total', () => {
  const m = mkModel();
  const p = Engine.parseParams({p_start:'2026-11-02', p_hol:false, p_disc:10}, {phaseDisc:{Assessment:100}});
  const R = Engine.compute(m, p, {});
  assert.equal(R.rows.find(r => r.phase === 'Assessment').net, 0);
  const build = R.rows.find(r => r.phase === 'Build');
  assert.ok(Math.abs(build.net - build.listNet * 0.9) < 1e-6);
  assert.ok(R.T.discPct > 0.1);
});

test('desconto em valor fixo e preço-alvo (valores COM impostos)', () => {
  const m = mkModel(), tax = [{name:'I', pct:10}];
  const base = Engine.compute(m, Engine.parseParams({p_start:'2026-11-02', p_hol:false}, {taxes:tax}), {});
  const amt = Engine.compute(m, Engine.parseParams({p_start:'2026-11-02', p_hol:false, p_discmode:'amount', p_discamt:10000}, {taxes:tax}), {});
  assert.ok(Math.abs((base.T.gross - amt.T.gross) - 10000) < 1e-6, 'desconto em valor');
  const tgt = Engine.compute(m, Engine.parseParams({p_start:'2026-11-02', p_hol:false, p_discmode:'target', p_disctarget:100000}, {taxes:tax}), {});
  assert.ok(Math.abs(tgt.T.gross - 100000) < 1e-6, 'preço-alvo');
  const above = Engine.compute(m, Engine.parseParams({p_start:'2026-11-02', p_hol:false, p_discmode:'target', p_disctarget:9e9}, {taxes:tax}), {});
  assert.equal(above.T.discNet, 0); assert.ok(above.issues.some(i => i.code === 'target_above_list'));
});

test('desconto máximo preserva o GM mínimo', () => {
  const m = mkModel();
  const R = Engine.compute(m, params({p_gmmin:30}), {});
  assert.ok(R.maxDisc > 0 && R.maxDisc < 1);
  const at = Engine.compute(m, params({p_gmmin:30, p_disc: R.maxDisc * 100}), {});
  assert.ok(Math.abs(at.T.gm - 0.30) < 1e-6, 'GM no limite: ' + at.T.gm);
  const over = Engine.compute(m, params({p_gmmin:30, p_disc: R.maxDisc * 100 + 3}), {});
  assert.ok(over.issues.some(i => i.code === 'gm_below_min'));
});

test('sensibilidade ignora o modo de desconto configurado', () => {
  const m = mkModel();
  const g = Engine.sensitivity(m, params({p_discmode:'amount', p_discamt:99999}), {}, [0, 0.1], [0]);
  assert.ok(g[1].cells[0].gross < g[0].cells[0].gross);
});

test('linhas sem perfil e fases sem alocação não entram no cálculo; perfis não usados não exigem taxa', () => {
  const m = mkModel();
  m.rows.push({phase:'Vazia', activity:'', role:'', level:'', location:'', optional:false, alloc:new Array(6).fill(0)});
  m.rows.push({phase:'Zerada', activity:'', role:'Arq', level:'Q5', location:'Brazil', optional:false, alloc:new Array(6).fill(0)});
  m.rates['Sobra|Q1|Brazil'] = {role:'Sobra', level:'Q1', location:'Brazil', cost:'', rate:''};
  const R = Engine.compute(m, params(), {});
  assert.equal(R.T.H, 640);
  assert.ok(!R.phases.some(f => f.name === 'Vazia' || f.name === 'Zerada'));
  assert.equal(R.missing.length, 0);
  assert.equal(R.rowByIndex.size, 4); // 3 originais + a linha com perfil e sem alocação
});

test('timeline só com linhas sem perfil: compute devolve null', () => {
  const m = mkModel(); m.rows = [{phase:'A', activity:'', role:'', level:'', location:'', optional:false, alloc:new Array(6).fill(0)}];
  assert.equal(Engine.compute(m, params(), {}), null);
});

test('parseTimeline com keepEmpty mantém linhas sem perfil e sem alocação', () => {
  const aoa = [['Phase','Role','Level','Location','Optional','SP1','SP2'], ['A','PM','Q4','BR','',0,0], ['B','','','','',0,0]];
  const t = Engine.parseTimeline(aoa, {keepEmpty:true});
  assert.equal(t.rows.length, 2); assert.equal(t.rows[1].role, '');
  assert.throws(() => Engine.parseTimeline(aoa), /nenhuma linha/);
});

test('ensureRates preserva o catálogo (prune só se pedido)', () => {
  const rates = {'Sobra|Q1|Brazil': {role:'Sobra', level:'Q1', location:'Brazil', cost:1, rate:2}};
  Engine.ensureRates(rates, [{role:'PM', level:'Q4', location:'BR'}, {role:'', level:'', location:''}]);
  assert.ok(rates['Sobra|Q1|Brazil'] && rates['PM|Q4|BR'] && Object.keys(rates).length === 2);
  Engine.ensureRates(rates, [{role:'PM', level:'Q4', location:'BR'}], true);
  assert.equal(Object.keys(rates).length, 1);
});

test('entregáveis: não manuais se distribuem igualmente na janela da fase; manuais ficam e são limitados à janela', () => {
  const dl = [
    {id:1, phase:'Build', name:'A', start:0, end:0, manual:false},
    {id:2, phase:'Build', name:'B', start:0, end:0, manual:false},
    {id:3, phase:'Build', name:'C', start:0, end:0, manual:false},
    {id:4, phase:'Assessment', name:'Rel', start:0, end:99, manual:true},
    {id:5, phase:'Fantasma', name:'X', start:0, end:0}
  ];
  const R = Engine.compute(mkModel({deliverables: dl}), params(), {today: TODAY});
  // Build ocupa SP3–SP5 (índices 2–4): 3 entregáveis, um por sprint
  const b = R.deliverables.filter(d => d.phase === 'Build');
  assert.deepEqual(b.map(d => [d.start, d.end]), [[2,2],[3,3],[4,4]]);
  // manual é limitado à janela da fase (Assessment = SP1–SP2)
  const a = R.deliverables.find(d => d.name === 'Rel');
  assert.deepEqual([a.start, a.end], [0, 1]);
  assert.ok(a.manual && a.endDate instanceof Date);
  // entregável de fase inexistente/sem alocação é ignorado
  assert.ok(!R.deliverables.some(d => d.name === 'X'));
  // 5 entregáveis em 2 sprints: ninguém sai da janela e todos ficam ≥ início
  const many = Engine.distributeDeliv(Array.from({length: 5}, (_, i) => ({phase:'P', name:'n'+i, manual:false})), 'P', 3, 4);
  assert.ok(many.every(d => d.start >= 3 && d.end <= 4 && d.end >= d.start));
});

test('timeline e marcos incluem os entregáveis', () => {
  const dl = [{id:1, phase:'Build', name:'Carga inicial', start:0, end:0, manual:false}, {id:2, phase:'Build', name:'Go-live', start:0, end:0, manual:false}];
  const R = Engine.compute(mkModel({deliverables: dl}), params(), {today: TODAY});
  const v = Engine.timelineView(R, {mode:'bars'});
  const rows = v.rows.filter(r => r.kind === 'deliv');
  assert.equal(rows.length, 2);
  assert.ok(rows.every(r => r.mark >= r.first && r.mark === r.last), 'o losango fica no fim do entregável');
  assert.equal(Engine.timelineView(R, {mode:'bars', deliv:false}).rows.filter(r => r.kind === 'deliv').length, 0);
  assert.equal(Engine.timelineView(R, {mode:'milestones'}).rows.filter(r => r.kind === 'deliv').length, 0);
  assert.ok(Engine.milestones(R, []).some(m => m.kind === 'deliv' && m.name.startsWith('Go-live')));
  // por semana (sprint de 2 semanas): o losango vai para a última semana do sprint
  const w = Engine.timelineView(R, {mode:'bars', gran:'week'});
  const d = w.rows.filter(r => r.kind === 'deliv').at(-1);   // o último entregável termina no SP5 (índice 4)
  assert.equal(w.cols[d.mark].period, 4);
  assert.equal(w.cols[d.mark + 1].period, 5, 'o losango fica na última semana do sprint');
});

test('comparação de cenários: métricas e melhor valor por indicador', () => {
  const a = Engine.compute(mkModel(), params(), {today: TODAY});
  const m2 = mkModel(); m2.rows = m2.rows.map(r => ({...r, alloc: r.alloc.map(v => v * 2)}));
  const b = Engine.compute(m2, params(), {today: TODAY});
  const M = [Engine.scenarioMetrics(a), Engine.scenarioMetrics(b), null];
  assert.ok(M[1].gross > M[0].gross && M[1].peak > M[0].peak);
  const best = Engine.compareScenarios(M);
  assert.deepEqual(best.gross, [0]); assert.deepEqual(best.peak, [0]);
  assert.equal(best.weeks, null, 'mesmo prazo = sem destaque');
  assert.equal(Engine.compareScenarios([M[0], null]).gross, null, 'um só cenário válido não tem comparação');
});

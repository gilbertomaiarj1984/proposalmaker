import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const src = f => fs.readFileSync(path.join(here, '../src', f), 'utf8');
const { Engine, Workbook } = new Function(src('engine.js') + '\n' + src('workbook.js') + '\n;return {Engine, Workbook};')();
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

function exampleState(over = {}){
  const ex = Workbook.exampleData();
  const rows = ex.rows.map(r => ({phase:r[0], activity:r[1], role:r[2], level:r[3], location:r[4], optional:r[5]==='Y', alloc:r[6]}));
  const rates = Engine.ensureRates({}, rows);
  Engine.applyRateTable(rates, ex.rates.map(r => ({role:r[0], level:r[1], location:r[2], cost:r[3], rate:r[4]})));
  const raw = {p_vendor:'ACME', p_client:'Cliente <X>', p_project:'Projeto', p_start:'2026-11-02', p_startmode:'date', p_hol:true, p_cur:'BRL', p_rcur:'BRL', p_fx:1,
    p_pw:2, p_hpd:8, p_dpw:5, p_mode:'rate', p_gm:35, p_cont:5, p_contmode:'price', p_disc:3, p_lang:'pt', p_inclOpt:false, p_pay:'milestone',
    p_assump:'Linha 1\nLinha 2', p_tlmiles:'Go-live | SP11', ...over};
  return {rows, nPeriods:12, unit:'SP', rates, taxes:[{name:'ISS',pct:2},{name:'PIS/COFINS',pct:3.65}], phaseDisc:{'Hypercare':50, 'Ingestion Migration':10}, edits:{'pt:exec.ctx':'<b>Texto</b> editado'}, raw, fileName:'x.xlsx'};
}
const run = (st) => Engine.compute({rows:st.rows, nPeriods:st.nPeriods, unit:st.unit, rates:st.rates}, Engine.parseParams(st.raw, {taxes:st.taxes, phaseDisc:st.phaseDisc}), {today:new Date(2026,9,19)});

test('ida e volta: gravar e ler devolve o mesmo estado', async () => {
  const st = exampleState();
  const R = run(st);
  const wb = await Workbook.buildProposal(ExcelJS, st, R);
  const buf = await wb.xlsx.writeBuffer();
  const rd = XLSX.read(new Uint8Array(buf), {type:'array'});
  assert.ok(Workbook.isProposalWorkbook(XLSX, rd));
  const back = Workbook.read(XLSX, rd);
  assert.equal(back.unit, 'SP'); assert.equal(back.nPeriods, 12);
  assert.equal(back.rows.length, st.rows.length);
  assert.deepEqual(back.rows[3].alloc, st.rows[3].alloc);
  assert.equal(back.raw.p_client, 'Cliente <X>');
  assert.equal(back.raw.p_cont, 5); assert.equal(back.raw.p_hol, true); assert.equal(back.raw.p_inclOpt, false);
  assert.deepEqual(back.taxes, st.taxes);
  assert.deepEqual(back.phaseDisc, st.phaseDisc);
  assert.equal(back.raw.p_discshow, undefined === back.raw.p_discshow ? undefined : back.raw.p_discshow);
  assert.equal(back.edits['pt:exec.ctx'], '<b>Texto</b> editado');
  assert.equal(Object.values(back.rates).find(r => r.role === 'Tester').rate, 140);
  assert.ok(!rd.SheetNames.includes('Perfis') && !rd.SheetNames.includes('RateCard'), 'a proposta não traz o cadastro de perfis');
  const R2 = run({...back, raw:back.raw});
  assert.ok(Math.abs(R2.T.gross - R.T.gross) < 1e-6);
  assert.ok(Math.abs(R2.T.H - R.T.H) < 1e-9);
});

test('arquivo que não é do app é rejeitado', async () => {
  const wb = new ExcelJS.Workbook(); wb.addWorksheet('Timeline').addRow(['a']);
  const rd = XLSX.read(new Uint8Array(await wb.xlsx.writeBuffer()), {type:'array'});
  assert.equal(Workbook.isProposalWorkbook(XLSX, rd), false);
});

test('template é lido de volta pelo parser do app (matriz + rate card)', async () => {
  const wb = await Workbook.buildTemplate(ExcelJS);
  const rd = XLSX.read(new Uint8Array(await wb.xlsx.writeBuffer()), {type:'array'});
  const aoa = n => XLSX.utils.sheet_to_json(rd.Sheets[n], {header:1, defval:'', raw:true, blankrows:false});
  const t = Engine.parseTimeline(aoa('Timeline'));
  assert.equal(t.rows.length, 13); assert.equal(t.nPeriods, 12);
  assert.equal(Engine.parseRateRows(aoa('Rate Card')).length, 7);
});

test('fórmulas recalculadas pelo LibreOffice batem com o motor', async (t) => {
  try { execFileSync('soffice', ['--version'], {stdio:'ignore'}); } catch(e){ return t.skip('LibreOffice indisponível'); }
  for (const over of [{}, {p_mode:'gm', p_gm:40, p_contmode:'both', p_inclOpt:true, p_fx:0.2, p_cur:'USD'}, {p_discmode:'amount', p_discamt:25000}, {p_discmode:'target', p_disctarget:700000}, {p_ovh:8, p_comm:1.5, p_fin:1, p_risk:2, p_terms:45, p_gpmmin:20}]) {
    const st = exampleState(over), R = run(st);
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pm-'));
    const f = path.join(dir, 'p.xlsx');
    fs.writeFileSync(f, Buffer.from(await (await Workbook.buildProposal(ExcelJS, st, R)).xlsx.writeBuffer()));
    // recalcula e converte para xlsx (LibreOffice recalcula ao abrir por fullCalcOnLoad)
    execFileSync('soffice', ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', path.join(dir, 'out'), f], {stdio:'ignore', timeout:120000});
    const rd = XLSX.read(fs.readFileSync(path.join(dir, 'out', 'p.xlsx')), {type:'buffer'});
    const rows = XLSX.utils.sheet_to_json(rd.Sheets['Resumo'], {header:1, defval:''});
    const get = label => rows.find(r => r[0] === label)[1];
    assert.ok(Math.abs(get('Horas totais') - R.T.H) < 1e-6, 'horas');
    assert.ok(Math.abs(get('Custo total') - R.T.cost) < 0.01, 'custo ' + get('Custo total') + ' vs ' + R.T.cost);
    assert.ok(Math.abs(get('Preço líquido de impostos') - R.T.net) < 0.01, 'net');
    assert.ok(Math.abs(get('Preço com impostos') - R.T.gross) < 0.01, 'gross');
    assert.ok(Math.abs(get('GM%') - R.T.gm) < 1e-6, 'gm');
    assert.ok(Math.abs(get('Desconto (com impostos)') - R.T.discGross) < 0.01, 'desconto');
    assert.ok(Math.abs(get('GPM% (após custos indiretos)') - R.T.gpm) < 1e-6, 'gpm ' + get('GPM% (após custos indiretos)') + ' vs ' + R.T.gpm);
    assert.ok(Math.abs(get('Custo financeiro (prazo de pagamento)') - R.T.ind.fin) < 0.01, 'custo financeiro');
  }
});

test('proposta ainda sem perfis alocados também é gravada e reaberta (linhas vazias preservadas)', async () => {
  const st = exampleState();
  st.rows.push({phase:'Fase nova', activity:'', role:'', level:'', location:'', optional:false, alloc:new Array(12).fill(0)});
  st.rows.push({phase:'Fase nova', activity:'', role:'Arquiteto', level:'Q5', location:'Brazil', optional:false, alloc:new Array(12).fill(0)});
  st.rates['Arquiteto|Q5|Brazil'] = {role:'Arquiteto', level:'Q5', location:'Brazil', cost:'', rate:''};
  st.rates['Sobra|Q1|Brazil'] = {role:'Sobra', level:'Q1', location:'Brazil', cost:50, rate:90};
  const R = run(st);
  const buf = await (await Workbook.buildProposal(ExcelJS, st, R)).xlsx.writeBuffer();
  const back = Workbook.read(XLSX, XLSX.read(new Uint8Array(buf), {type:'array'}));
  assert.equal(back.rows.length, st.rows.length);
  assert.equal(back.rows[back.rows.length - 2].role, '');
  assert.equal(back.rows[back.rows.length - 1].role, 'Arquiteto');
  assert.ok(!back.rates['Sobra|Q1|Brazil'], 'perfis não usados ficam só no catálogo, não na proposta');
  assert.equal(back.rates['Arquiteto|Q5|Brazil'].rate, '', 'perfil usado sem taxa volta sem taxa');
  // sem nenhuma alocação
  const empty = {...st, rows:[{phase:'Fase 1', activity:'', role:'', level:'', location:'', optional:false, alloc:new Array(12).fill(0)}]};
  const R0 = Engine.compute({rows:empty.rows, nPeriods:12, unit:'SP', rates:empty.rates}, Engine.parseParams(empty.raw, {taxes:empty.taxes}), {});
  assert.equal(R0, null);
  const wb0 = await Workbook.buildProposal(ExcelJS, empty, Workbook.emptyResult(empty));
  const back0 = Workbook.read(XLSX, XLSX.read(new Uint8Array(await wb0.xlsx.writeBuffer()), {type:'array'}));
  assert.equal(back0.rows.length, 1); assert.equal(back0.rows[0].phase, 'Fase 1');
});

test('cenários e entregáveis: ida e volta no Excel (Modelo = ativo; demais em Cen_n)', async () => {
  const st = exampleState();
  const rowsB = st.rows.map(r => ({...r, alloc: r.alloc.map(v => v * 2)}));   // cenário B: equipe dobrada (outra alocação)
  const delA = [{id:1, phase:st.rows[0].phase, name:'Plano do projeto', start:0, end:0, manual:false}, {id:2, phase:st.rows[0].phase, name:'Relatório final', start:1, end:1, manual:true}];
  const delB = [{id:1, phase:st.rows[0].phase, name:'Plano acelerado', start:0, end:0, manual:true}];
  const scen = [
    {id:1, name:'Base', rows:st.rows, nPeriods:12, unit:'SP', deliverables:delA, tl:{p_pw:'2', p_unit:'SP'}},
    {id:3, name:'Acelerado', rows:rowsB, nPeriods:10, unit:'W', deliverables:delB, tl:{p_pw:'1', p_unit:'W'}}
  ];
  const full = {...st, deliverables: delA, scenarios: scen, active: 0};
  const R = run({...full, deliverables: delA});
  R.deliverables = R.deliverables || [];
  const buf = await (await Workbook.buildProposal(ExcelJS, full, R)).xlsx.writeBuffer();
  const rd = XLSX.read(new Uint8Array(buf), {type:'array'});
  assert.ok(rd.SheetNames.includes('Cenarios') && rd.SheetNames.includes('Cen_3') && !rd.SheetNames.includes('Cen_1'), rd.SheetNames.join());
  const back = Workbook.read(XLSX, rd);
  assert.equal(back.scenarios.length, 2); assert.equal(back.active, 0);
  const [a, b] = back.scenarios;
  assert.equal(a.name, 'Base'); assert.equal(b.name, 'Acelerado'); assert.equal(b.id, 3);
  assert.deepEqual(a.rows[3].alloc, st.rows[3].alloc);
  assert.deepEqual(b.rows[3].alloc.slice(0, 10), rowsB[3].alloc.slice(0, 10));
  assert.equal(b.unit, 'W'); assert.equal(b.nPeriods, 10); assert.equal(b.tl.p_pw, '1');
  assert.deepEqual(back.deliverables.map(d => [d.name, d.start, d.end, d.manual]), [['Plano do projeto', 0, 0, false], ['Relatório final', 1, 1, true]]);
  assert.deepEqual(b.deliverables.map(d => [d.name, d.start, d.end, d.manual]), [['Plano acelerado', 0, 0, true]]);
  assert.equal(Object.values(back.rates).find(r => r.role === 'Tester').rate, 140, 'taxas do cenário inativo também voltam');
  // ativo = B: Modelo traz B e Cen_1 traz A
  const full2 = {...st, rows: rowsB, nPeriods: 10, unit: 'W', deliverables: delB, scenarios: scen, active: 1};
  const back2 = Workbook.read(XLSX, XLSX.read(new Uint8Array(await (await Workbook.buildProposal(ExcelJS, full2, run(full2))).xlsx.writeBuffer()), {type:'array'}));
  assert.equal(back2.active, 1); assert.equal(back2.scenarios[0].rows.length, st.rows.length);
  assert.deepEqual(back2.rows[3].alloc.slice(0, 10), rowsB[3].alloc.slice(0, 10));
});

test('proposta com um único cenário não ganha abas de cenário e mantém os entregáveis', async () => {
  const st = exampleState();
  const dl = [{id:1, phase:st.rows[0].phase, name:'Entrega única', start:0, end:1, manual:true}];
  const full = {...st, deliverables: dl};
  const rd = XLSX.read(new Uint8Array(await (await Workbook.buildProposal(ExcelJS, full, run(full))).xlsx.writeBuffer()), {type:'array'});
  assert.ok(!rd.SheetNames.includes('Cenarios') && rd.SheetNames.includes('Entregaveis'));
  const back = Workbook.read(XLSX, rd);
  assert.equal(back.scenarios, null);
  assert.deepEqual(back.deliverables.map(d => [d.name, d.start, d.end]), [['Entrega única', 0, 1]]);
});

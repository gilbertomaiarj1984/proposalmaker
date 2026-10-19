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
  return {rows, nPeriods:12, unit:'SP', rates, taxes:[{name:'ISS',pct:2},{name:'PIS/COFINS',pct:3.65}], edits:{'pt:exec.ctx':'<b>Texto</b> editado'}, raw, fileName:'x.xlsx'};
}
const run = (st) => Engine.compute({rows:st.rows, nPeriods:st.nPeriods, unit:st.unit, rates:st.rates}, Engine.parseParams(st.raw, {taxes:st.taxes}), {today:new Date(2026,9,19)});

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
  assert.equal(back.edits['pt:exec.ctx'], '<b>Texto</b> editado');
  assert.equal(Object.values(back.rates).find(r => r.role === 'Tester').rate, 140);
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
  for (const over of [{}, {p_mode:'gm', p_gm:40, p_contmode:'both', p_inclOpt:true, p_fx:0.2, p_cur:'USD'}]) {
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
  }
});

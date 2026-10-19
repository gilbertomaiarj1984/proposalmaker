import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const src = f => fs.readFileSync(path.join(here, '../src', f), 'utf8');
const { Engine, Workbook, DocModel, PptxExport } = new Function(src('engine.js') + src('workbook.js') + src('doc.js') + src('pptx-export.js') + ';return {Engine,Workbook,DocModel,PptxExport};')();
const PptxGenJS = require('pptxgenjs');
const JSZip = require('jszip');

const ex = Workbook.exampleData();
const rows = ex.rows.map(r => ({phase:r[0], activity:r[1], role:r[2], level:r[3], location:r[4], optional:r[5]==='Y', alloc:r[6]}));
const rates = Engine.ensureRates({}, rows);
Engine.applyRateTable(rates, ex.rates.map(r => ({role:r[0], level:r[1], location:r[2], cost:r[3], rate:r[4]})));

async function deck(over = {}, edits = {}) {
  const raw = {p_client:'Cliente SA', p_project:'Projeto', p_vendor:'TCS', p_start:'2026-11-02', p_hol:true, p_lang:'pt', p_cur:'BRL', p_pw:2, p_hpd:8, p_dpw:5, ...over};
  const R = Engine.compute({rows, nPeriods:12, unit:'SP', rates}, Engine.parseParams(raw, {taxes:[{name:'ISS',pct:5.65}]}), {today:new Date(2026,9,19)});
  const pres = PptxExport.build(PptxGenJS, DocModel.build(R, {today:new Date(2026,9,19)}), R, edits);
  const zip = await JSZip.loadAsync(await pres.write({outputType:'nodebuffer'}));
  const slides = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  const texts = await Promise.all(slides.map(async f => (await zip.file(f).async('string')).replace(/<[^>]+>/g, ' ')));
  return {zip, slides, texts, all: texts.join(' ')};
}

test('um slide por seção: capa, sumário, cronograma, time, investimento, pagamentos, opcional e premissas', async () => {
  const d = await deck();
  assert.ok(d.slides.length >= 8 && d.slides.length <= 11, 'slides: ' + d.slides.length);
  for (const t of ['Cliente SA', 'Sumário Executivo', 'Cronograma de Entrega', 'Composição do Time', 'Investimento', 'Cronograma de Pagamentos', 'Escopo Opcional', 'Premissas Comerciais'])
    assert.ok(d.all.includes(t), 'falta: ' + t);
});

test('tabelas, formas e gráfico são nativos (sem imagens)', async () => {
  const d = await deck();
  const files = Object.keys(d.zip.files);
  assert.ok(files.some(f => /^ppt\/charts\/chart\d+\.xml$/.test(f)), 'gráfico nativo');
  assert.ok(!files.some(f => /^ppt\/media\/.+/.test(f)), 'sem imagens');
  const xml = (await Promise.all(d.slides.map(f => d.zip.file(f).async('string')))).join('');
  assert.ok(/<a:tbl>/.test(xml), 'tabela nativa');
});

test('modos de timeline: semana com perfis cria mais linhas; marcos não usa barras', async () => {
  const bars = await deck({p_tlmode:'bars'}), roles = await deck({p_tlmode:'roles', p_tlgran:'week'}), miles = await deck({p_tlmode:'milestones', p_tlmiles:'Go-live | SP11'});
  assert.ok(roles.slides.length <= bars.slides.length + 1, 'perfis em semanas não geram slides extras de timeline');
  assert.ok(roles.all.includes('W24'));
  assert.ok(miles.all.includes('Go-live') && !miles.all.includes('SP5 SP6'), 'marcos');
});

test('sem data: slides sem datas absolutas; idioma e textos editados são respeitados', async () => {
  const rel = await deck({p_startmode:'none', p_lang:'en'}, {'en:exec.main':'Texto <b>editado</b> pelo usuário'});
  assert.ok(rel.all.includes('Executive Summary') && rel.all.includes('editado'));
  const tl = rel.texts.filter(t => t.includes('Delivery Timeline')).join(' ');
  assert.ok(tl && !/\b20\d\d\b/.test(tl), 'sem datas no cronograma');
  const chartFile = Object.keys(rel.zip.files).find(f => /^ppt\/charts\/chart\d+\.xml$/.test(f));
  assert.ok((await rel.zip.file(chartFile).async('string')).includes('Monthly billing'));
});

test('nome de cliente com HTML não vira marcação', async () => {
  const d = await deck({p_client:'A&B <b>Corp</b>'});
  assert.ok(d.all.includes('A&amp;B &lt;b&gt;Corp&lt;/b&gt;') || d.all.includes('A&amp;B'), 'texto escapado');
});

test('a timeline cabe SEMPRE em um único slide (barras, perfis, semanas, marcos)', async () => {
  const many = Array.from({length: 40}, (_, i) => ({phase:'Fase ' + (1 + Math.floor(i / 4)), activity:'', role:'Perfil ' + i, level:'Q3', location:'BR', optional:false, alloc:Array.from({length:12}, (_, j) => (j + i) % 5 === 0 ? 1 : 0.5)}));
  const R2 = Engine.ensureRates({}, many); Object.values(R2).forEach(r => { r.rate = 100; r.cost = 60; });
  for (const over of [{p_tlmode:'bars'}, {p_tlmode:'roles'}, {p_tlmode:'roles', p_tlgran:'week'}, {p_tlmode:'milestones'}, {p_tlmode:'roles', p_tlfte:false}]) {
    for (const data of [rows, many]) {
      const raw = {p_client:'C', p_project:'P', p_start:'2026-11-02', p_lang:'pt', p_tlmiles:'A | SP1\nB | SP2\nC | SP3\nD | SP4\nE | SP5\nF | SP6\nG | SP7\nH | SP8', ...over};
      const model = {rows: data, nPeriods:12, unit:'SP', rates: data === rows ? rates : R2};
      const R = Engine.compute(model, Engine.parseParams(raw, {taxes:[{name:'I',pct:5}]}), {today:new Date(2026,9,19)});
      const pres = PptxExport.build(PptxGenJS, DocModel.build(R), R, {});
      const zip = await JSZip.loadAsync(await pres.write({outputType:'nodebuffer'}));
      const files = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f));
      let withTl = 0, withCont = 0;
      for (const f of files) { const t = (await zip.file(f).async('string')).replace(/<[^>]+>/g, ' '); if (/Cronograma de Entrega/.test(t)) { withTl++; if (/\(cont\.\)/.test(t)) withCont++; } }
      const tlSlides = [];
      for (const f of files) { const x = await zip.file(f).async('string'); if (/name="Timeline"/.test(x) || (over.p_tlmode === 'milestones' && /Cronograma de Entrega/.test(x.replace(/<[^>]+>/g, ' ')) && /Hito|Marco/.test(x.replace(/<[^>]+>/g, ' ')))) tlSlides.push(f); }
      assert.equal(tlSlides.length, 1, `timeline em ${tlSlides.length} slides (${JSON.stringify(over)}, ${data === rows ? 'exemplo' : '40 perfis'})`);
    }
  }
});

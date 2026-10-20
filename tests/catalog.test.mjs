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
const { Engine, Workbook, Catalog } = new Function(src('engine.js') + src('workbook.js') + src('catalog.js') + ';return {Engine,Workbook,Catalog};')();
const ExcelJS = require('exceljs');

const seedBuf = async () => Buffer.from(await Catalog.seed(ExcelJS).xlsx.writeBuffer());
const get = (clients, n) => clients.find(c => c.name === n);

test('modelo: abas Roche e JnJ são lidas pelo cabeçalho de cada cliente', async () => {
  const {clients, skipped} = await Catalog.parse(ExcelJS, await seedBuf());
  assert.deepEqual(clients.map(c => c.name), ['Roche', 'JnJ']);
  assert.deepEqual(skipped, ['Instructions']);
  const jnj = get(clients, 'JnJ');
  assert.deepEqual(jnj.cols, {role: 1, level: 2, location: 3, cost: 4, rate: 5}, 'JnJ usa o layout padrão Role | Level | Location | Cost/h | Rate/h');
  assert.equal(jnj.entries.length, 22);
  assert.deepEqual([jnj.entries[0].role, jnj.entries[0].level, jnj.entries[0].rate], ['Developer: SAP BASIS', 'Senior', 142.91]);
  assert.equal(jnj.entries.find(e => e.role === 'SAP Architect' && e.level === 'Senior').rate, '', 'perfil sem taxa fica vazio');
  const roche = get(clients, 'Roche');
  assert.ok(roche.cols.cost && roche.cols.rate && roche.cols.location);
  assert.equal(roche.entries.length, 7);
  assert.deepEqual([roche.entries[0].role, roche.entries[0].level, roche.entries[0].location, roche.entries[0].cost, roche.entries[0].rate], ['Project Manager', 'Q4', 'Brazil', 170, 260]);
});

test('cabeçalhos variados: "Perfil / Nível / Rate/h" e taxa "sem impostos"', async () => {
  const wb = new ExcelJS.Workbook();
  const a = wb.addWorksheet('Acme'); a.addRow(['Título solto']); a.addRow(['Perfil', 'Nível', 'Valor hora sem impostos (USD)']); a.addRow(['Dev', 'Q3', 'US$ 1.234,50']);
  const b = wb.addWorksheet('Outra'); b.addRow(['Nada', 'aqui']);
  const {clients, skipped} = await Catalog.parse(ExcelJS, Buffer.from(await wb.xlsx.writeBuffer()));
  assert.equal(clients.length, 1); assert.deepEqual(skipped, ['Outra']);
  assert.equal(clients[0].hdrRow, 2); assert.equal(clients[0].currency, 'USD');
  assert.equal(clients[0].entries[0].rate, 1234.5);
});

test('edição aplicada na planilha original: atualiza, acrescenta e remove linhas preservando formatação', async () => {
  const buf = await seedBuf();
  const {clients} = await Catalog.parse(ExcelJS, buf);
  const jnj = get(clients, 'JnJ');
  jnj.entries.find(e => e.role === 'SAP Architect' && e.level === 'Senior').rate = 250;
  jnj.entries.find(e => e.role === 'Platform Engineer' && e.level === 'Junior').deleted = true;
  jnj.entries.push({role: 'QA Lead', level: 'Senior', location: '', cost: '', rate: 99.9});
  get(clients, 'Roche').entries[0].rate = 275;
  const out = Buffer.from((await Catalog.serialize(ExcelJS, buf, {clients})).buffer);
  const back = await Catalog.parse(ExcelJS, out);
  const j2 = get(back.clients, 'JnJ');
  assert.equal(j2.entries.length, 22);                                   // -1 removida, +1 nova
  assert.equal(j2.entries.find(e => e.role === 'SAP Architect' && e.level === 'Senior').rate, 250);
  assert.ok(!j2.entries.some(e => e.role === 'Platform Engineer' && e.level === 'Junior'));
  assert.equal(j2.entries.find(e => e.role === 'QA Lead').rate, 99.9);
  assert.equal(get(back.clients, 'Roche').entries[0].rate, 275);
  // formatação do cliente preservada
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(out);
  const ws = wb.getWorksheet('JnJ');
  assert.equal(ws.getCell('A1').fill.fgColor.argb, 'FFD9E5F3');
  assert.equal(ws.getCell('E1').fill.fgColor.argb, 'FFD9E5F3');
  assert.ok(ws.autoFilter, 'filtro preservado');
  assert.match(ws.getCell('E2').numFmt, /#,##0/);
  const added = j2.entries.find(e => e.role === 'QA Lead');
  assert.match(ws.getCell(`E${added.row}`).numFmt, /#,##0/, 'linha nova herda o formato');
});

test('cliente novo baseado em outro: a aba é clonada com perfis, taxas e formato', async () => {
  const buf = await seedBuf();
  const {clients} = await Catalog.parse(ExcelJS, buf);
  const jnj = get(clients, 'JnJ');
  const novo = {name: 'Acme', sheet: 'Acme', isNew: true, base: 'JnJ', hdrRow: jnj.hdrRow, cols: {...jnj.cols}, currency: jnj.currency, entries: jnj.entries.map(e => ({...e}))};
  novo.entries[0].rate = 150;
  const out = Buffer.from((await Catalog.serialize(ExcelJS, buf, {clients: [...clients, novo]})).buffer);
  const back = await Catalog.parse(ExcelJS, out);
  assert.deepEqual(back.clients.map(c => c.name), ['Roche', 'JnJ', 'Acme']);
  const a = get(back.clients, 'Acme');
  assert.equal(a.entries.length, 22); assert.equal(a.entries[0].rate, 150);
  assert.equal(get(back.clients, 'JnJ').entries[0].rate, 142.91, 'a aba base não é afetada');
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(out);
  assert.equal(wb.getWorksheet('Acme').getCell('A1').fill.fgColor.argb, 'FFD9E5F3', 'cabeçalho clonado');
  assert.equal(wb.worksheets.length, 4);
});

test('cliente novo em branco e nomes de aba válidos e únicos', async () => {
  const buf = await seedBuf();
  const {clients} = await Catalog.parse(ExcelJS, buf);
  const blank = {name: 'Cliente: Novo/Teste?', sheet: 'Cliente: Novo/Teste?', isNew: true, hdrRow: 1, cols: {role: 1, level: 2, location: 3, cost: 4, rate: 5}, entries: [{role: 'Dev', level: 'Q3', location: 'Brazil', cost: 100, rate: 180}]};
  const out = Buffer.from((await Catalog.serialize(ExcelJS, buf, {clients: [...clients, blank]})).buffer);
  const back = await Catalog.parse(ExcelJS, out);
  const c = back.clients.at(-1);
  assert.match(c.name, /^[^\\/?*[\]:]{1,31}$/); assert.equal(c.entries[0].rate, 180);
  assert.equal(Catalog.sheetName('Roche', ['Roche', 'JnJ']), 'Roche 2');
  assert.equal(Catalog.sheetName('X'.repeat(40), []).length, 31);
});

test('arquivo gravado abre no LibreOffice (formato válido)', async (t) => {
  try { execFileSync('soffice', ['--version'], {stdio: 'ignore'}); } catch(e){ return t.skip('LibreOffice indisponível'); }
  const buf = await seedBuf();
  const {clients} = await Catalog.parse(ExcelJS, buf);
  const jnj = get(clients, 'JnJ');
  const novo = {name: 'Acme', sheet: 'Acme', isNew: true, base: 'JnJ', hdrRow: jnj.hdrRow, cols: {...jnj.cols}, entries: jnj.entries.map(e => ({...e}))};
  const out = (await Catalog.serialize(ExcelJS, buf, {clients: [...clients, novo]})).buffer;
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cat-'));
  fs.writeFileSync(path.join(dir, 'c.xlsx'), Buffer.from(out));
  execFileSync('soffice', ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', path.join(dir, 'o'), path.join(dir, 'c.xlsx')], {stdio: 'ignore', timeout: 120000});
  const back = await Catalog.parse(ExcelJS, fs.readFileSync(path.join(dir, 'o', 'c.xlsx')));
  assert.deepEqual(back.clients.map(c => c.name), ['Roche', 'JnJ', 'Acme']);
  assert.equal(get(back.clients, 'Acme').entries.length, 22);
});

test('gravação devolve o mapa de linhas: novas, removidas e deslocadas', async () => {
  const buf = await seedBuf();
  const {clients} = await Catalog.parse(ExcelJS, buf);
  const jnj = get(clients, 'JnJ');
  const first = jnj.entries[0], third = jnj.entries[2], last = jnj.entries.at(-1);
  first.deleted = true;
  const fresh = {id: Catalog.newId(), role: 'Novo', level: 'Senior', location: '', cost: '', rate: 10};
  jnj.entries.push(fresh);
  const r = await Catalog.serialize(ExcelJS, buf, {clients});
  assert.deepEqual(r.removed, [first.id]);
  assert.equal(r.rows[third.id], third.row - 1, 'linhas abaixo da removida sobem uma');
  assert.equal(r.rows[fresh.id], last.row, 'a nova ocupa o fim (já deslocado)');
  const back = await Catalog.parse(ExcelJS, Buffer.from(r.buffer));
  const j2 = get(back.clients, 'JnJ');
  assert.equal(j2.entries.find(e => e.role === 'Novo').row, r.rows[fresh.id]);
  assert.equal(j2.entries.find(e => e.role === 'Functional Analyst: SAP functional analyst' && e.level === 'Junior').row, r.rows[third.id]);
  assert.equal(r.sheets.JnJ, 'JnJ');
});

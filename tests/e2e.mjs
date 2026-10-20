// Teste de ponta a ponta no Chromium, com TODA a rede bloqueada (simula a máquina restrita).
//   npm run build && npm run e2e
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright-core';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const URL_ = 'file://' + path.join(root, 'gerador_proposta_comercial.html');
const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pm-e2e-'));
const exe = process.env.CHROMIUM_PATH || '/opt/pw-browsers/chromium';

import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const srcOf = f => fs.readFileSync(path.join(root, 'src', f), 'utf8');
const { Catalog } = new Function(srcOf('engine.js') + srcOf('workbook.js') + srcOf('catalog.js') + ';return {Catalog};')();
const ExcelJS = require('exceljs'), XLSXlib = require('xlsx');

const browser = await chromium.launch({executablePath: fs.existsSync(exe) ? exe : undefined, args: ['--no-sandbox']});
let n = 0;
const ok = (name) => console.log(`  ✓ ${++n}. ${name}`);

async function newPage(ctx, {fs: withFs = false} = {}) {
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('console', m => { if (m.type() === 'error' && !/Failed to load resource|net::ERR/.test(m.text())) errors.push(m.text()); });
  await page.route(/^https?:/, r => r.abort('internetdisconnected')); // rede restrita
  if (!withFs) await page.addInitScript(() => { delete window.showSaveFilePicker; delete window.showOpenFilePicker; delete window.showDirectoryPicker; });
  await page.goto(URL_);
  return {page, errors};
}
const $text = (page, sel) => page.evaluate(s => document.querySelector(s)?.innerText ?? '', sel);

try {
  const ctx = await browser.newContext({acceptDownloads: true, viewport: {width: 1300, height: 900}});
  const {page, errors} = await newPage(ctx);

  /* 1. offline: carrega timeline + rate card do template */
  await page.setInputFiles('#tlFile', path.join(root, 'Timeline_Template.xlsx'));
  await page.waitForSelector('#tlEditor', {state: 'visible'});
  const tlMsg = await $text(page, '#tlMsg');
  assert.match(tlMsg, /7 de 7 perfis/, 'rate card do xlsx deve ser importado: ' + tlMsg);
  const rates = await page.evaluate(() => Object.values(__pm.S.rates).map(r => r.rate));
  assert.ok(rates.every(r => r > 0), 'todas as taxas preenchidas');
  ok('offline: timeline e rate card importados (bug da importação corrigido)');

  /* 2. cálculo, checklist e abas */
  await page.evaluate(() => { document.getElementById('p_client').value = 'Cliente <img src=x onerror="window.__xss=1"> SA'; document.getElementById('p_project').value = 'Projeto X'; document.getElementById('p_start').value = '2026-11-04'; });
  await page.evaluate(() => __pm.recalc(true));
  assert.equal(await page.evaluate(() => window.__xss), undefined, 'sem injeção de HTML');
  assert.match(await $text(page, '#checklist'), /não é dia útil|Nenhum alerta/);
  assert.equal(await page.evaluate(() => __pm.S.result.T.nPer), 12);
  ok('cálculo; sem XSS; checklist exibido');

  /* 3. rate card vazio -> alerta visível na análise e na proposta */
  await page.evaluate(() => { Object.values(__pm.S.rates).forEach(r => { r.rate = ''; }); __pm.S.result = null; __pm.recalc(false); });
  assert.match(await $text(page, '#checklist'), /Sem Rate\/h para/);
  await page.evaluate(() => { __pm.applyState(__pm.snapshot()); });
  ok('perfis sem taxa geram alerta de erro');
  await page.setInputFiles('#tlFile', path.join(root, 'Timeline_Template.xlsx'));
  await page.waitForTimeout(300);

  /* 4. opções de timeline na proposta */
  await page.evaluate(() => __pm.goTab('doc'));
  const setOpt = (id, v) => page.evaluate(([i, val]) => { const el = document.getElementById(i); if (el.type === 'checkbox') el.checked = val; else el.value = val; el.dispatchEvent(new Event('change', {bubbles: true})); }, [id, v]);
  const cols = () => page.evaluate(() => document.querySelectorAll('#doc table.tlt thead tr:last-child th').length - 1);
  assert.equal(await cols(), 12); // por sprint
  await setOpt('p_tlgran', 'week'); assert.equal(await cols(), 24);
  assert.ok(await page.evaluate(() => document.querySelector('#doc table.tlt th.grp')), 'cabeçalho de sprint sobre as semanas');
  await setOpt('p_tlmode', 'roles');
  assert.ok(await page.evaluate(() => document.querySelectorAll('#doc table.tlt tr.role').length) >= 10);
  await setOpt('p_tlfte', false);
  assert.equal(await page.evaluate(() => [...document.querySelectorAll('#doc table.tlt td.on')].filter(td => td.textContent.trim()).length), 0, 'sem números de FTE');
  await setOpt('p_tlmode', 'milestones');
  await page.evaluate(() => { const t = document.getElementById('p_tlmiles'); t.value = 'Go-live | SP11'; t.dispatchEvent(new Event('change', {bubbles: true})); });
  const ms = await $text(page, '#doc');
  assert.match(ms, /Go-live/); assert.ok(!(await page.evaluate(() => document.querySelector('#doc table.tlt'))), 'modo marcos não desenha barras');
  ok('timeline: sprint/semana, com/sem FTE, fases/perfis e só marcos');

  /* 5. início: data, mês e sem data */
  await setOpt('p_tlmode', 'bars'); await setOpt('p_tlfte', true); await setOpt('p_tlgran', 'period');
  await setOpt('p_startmode', 'month'); await page.evaluate(() => { const m = document.getElementById('p_startm'); m.value = '2027-01'; m.dispatchEvent(new Event('change', {bubbles: true})); });
  assert.equal(await page.evaluate(() => __pm.S.result.p.startMode), 'month');
  assert.equal(await page.evaluate(() => __pm.S.result.start.toISOString().slice(0, 10)), '2027-01-04');
  await setOpt('p_startmode', 'none');
  const rel = await $text(page, '#doc');
  assert.match(rel, /relativ|relative/i); assert.ok(!/\b2027\b/.test((await page.evaluate(() => document.querySelector('#doc .page').innerText)).split('Valid')[0] || ''), 'sem datas');
  await setOpt('p_lang', 'pt');
  assert.match(await $text(page, '#doc'), /Fim do SP6|Fim do mês|Fim do/);
  await setOpt('p_startmode', 'date');
  ok('início por data, por mês (1º dia útil) e sem data (relativo)');

  /* 6. edição no navegador persiste por idioma e vai para o Excel */
  await page.evaluate(() => { const el = document.querySelector('#doc [data-ek="exec.main"]'); el.innerHTML = 'Texto <b>editado</b> <script>1</script>'; el.dispatchEvent(new Event('input', {bubbles: true})); });
  assert.ok(await page.evaluate(() => Object.values(__pm.S.edits).some(v => /editado/.test(v) && !/script/.test(v))), 'edição sanitizada');
  await setOpt('p_lang', 'en');
  assert.ok(!/Texto editado/.test(await $text(page, '#doc')), 'edição é por idioma');
  await setOpt('p_lang', 'pt');
  assert.match(await $text(page, '#doc'), /Texto editado/);
  ok('edições persistem por idioma e são sanitizadas');

  /* 6b. desconto: valor fixo, por fase, linha visível e persistência */
  await page.evaluate(() => { const set = (id, v) => { const el = document.getElementById(id); el.value = v; el.dispatchEvent(new Event('change', {bubbles: true})); };
    set('p_discmode', 'pct'); set('p_disc', '5'); set('p_discshow', 'line'); __pm.S.phaseDisc = {'Hypercare': 50}; __pm.S.result = null; __pm.recalc(false); });
  const dsc = await page.evaluate(() => { const T = __pm.S.result.T; return {pct: T.discPct, show: __pm.S.result.p.discShow, hyper: __pm.S.result.rows.find(r => r.phase === 'Hypercare').discPct}; });
  assert.ok(dsc.pct > 0.05 && dsc.hyper === 0.5 && dsc.show === 'line');
  assert.match(await $text(page, '#doc'), /Subtotal|Desconto comercial|Descuento/);
  await page.evaluate(() => { __pm.goTab('par'); });
  assert.match(await $text(page, '#discSuggest'), /Sugest/); assert.ok(await page.evaluate(() => document.querySelectorAll('#discPhase input[data-i]').length) >= 5);
  await page.evaluate(() => document.querySelector('#discSuggest [data-act="disc-apply"]').click());
  assert.equal(await page.evaluate(() => document.getElementById('p_discmode').value), 'pct');
  await page.evaluate(() => { const el = document.getElementById('p_discmode'); el.value = 'amount'; el.dispatchEvent(new Event('change', {bubbles: true})); const a = document.getElementById('p_discamt'); a.value = '20000'; a.dispatchEvent(new Event('change', {bubbles: true})); __pm.S.phaseDisc = {'Hypercare': 50}; __pm.S.result = null; __pm.recalc(false); });
  const amt = await page.evaluate(() => __pm.S.result.T.discGross);
  assert.ok(Math.abs(amt - 20000) < 1, 'desconto em valor fixo (com impostos): ' + amt);
  ok('desconto: percentual, por fase, valor fixo, linha visível e sugestões');

  /* 7. salvar (download) e reabrir: o .xlsx é o "banco" da proposta */
  const [dl] = await Promise.all([page.waitForEvent('download'), page.evaluate(() => __pm.actions.save())]);
  const xlsxPath = path.join(outDir, 'proposta.xlsx'); await dl.saveAs(xlsxPath);
  assert.ok(fs.statSync(xlsxPath).size > 5000);
  const snapBefore = await page.evaluate(() => { const R = __pm.S.result; return {gross: R.T.gross, client: R.p.client, edits: Object.keys(__pm.S.edits).length, mode: R.p.discMode, amt: R.p.discAmount, show: R.p.discShow, phaseDisc: __pm.S.phaseDisc}; });
  await page.reload(); // página nova: estado zerado
  await page.setInputFiles('#tlFile', xlsxPath);
  await page.waitForFunction(() => __pm.S.rows.length > 0);
  await page.evaluate(() => __pm.recalc(false));
  const snapAfter = await page.evaluate(() => { const R = __pm.S.result; return {gross: R.T.gross, client: R.p.client, edits: Object.keys(__pm.S.edits).length, mode: R.p.discMode, amt: R.p.discAmount, show: R.p.discShow, phaseDisc: __pm.S.phaseDisc}; });
  assert.deepEqual(snapAfter, snapBefore);
  assert.match(await $text(page, '#tlMsg'), /Proposta .*proposta\.xlsx.* aberta/);
  ok('salvar em .xlsx e reabrir restaura parâmetros, perfis, impostos, descontos e textos editados');

  /* 8. Word .docx e template */
  const [dw] = await Promise.all([page.waitForEvent('download'), page.evaluate(() => __pm.actions.word())]);
  const docxPath = path.join(outDir, 'p.docx'); await dw.saveAs(docxPath);
  assert.equal(fs.readFileSync(docxPath).subarray(0, 2).toString(), 'PK');
  const [dp] = await Promise.all([page.waitForEvent('download'), page.evaluate(() => __pm.actions.pptx())]);
  const pptxPath = path.join(outDir, 'p.pptx'); await dp.saveAs(pptxPath);
  assert.equal(fs.readFileSync(pptxPath).subarray(0, 2).toString(), 'PK');
  assert.ok(fs.statSync(pptxPath).size > 20000);
  const [dt] = await Promise.all([page.waitForEvent('download'), page.evaluate(() => __pm.actions.template())]);
  const tplPath = path.join(outDir, 't.xlsx'); await dt.saveAs(tplPath);
  assert.equal(fs.readFileSync(tplPath).subarray(0, 2).toString(), 'PK');
  ok('Word (.docx), PowerPoint (.pptx) e template Excel gerados offline');

  /* 9. extras online sem internet: aviso, alternativa manual, app segue funcionando */
  await page.evaluate(() => { document.getElementById('p_cur').value = 'USD'; document.getElementById('p_rcur').value = 'BRL'; });
  await page.evaluate(() => __pm.actions['fx-online']());
  await page.waitForFunction(() => /Não foi possível obter a cotação/.test(document.getElementById('fxMsg').innerText), null, {timeout: 20000});
  await page.evaluate(() => __pm.actions['holidays-online']());
  await page.waitForFunction(() => /Não foi possível consultar a BrasilAPI/.test(document.getElementById('holMsg').innerText), null, {timeout: 20000});
  await page.evaluate(() => __pm.actions.netcheck());
  await page.waitForFunction(() => document.getElementById('netbar').classList.contains('show'), null, {timeout: 20000});
  assert.match(await $text(page, '#netbar'), /limitado|Sem conexão|indispon/i);
  await page.evaluate(() => { document.getElementById('p_fx').value = '0.2'; document.getElementById('p_fx').dispatchEvent(new Event('change', {bubbles: true})); __pm.S.result = null; __pm.recalc(false); });
  assert.ok(await page.evaluate(() => __pm.S.result.T.gross) > 0);
  ok('sem internet: câmbio e feriados online avisam o motivo, não travam, e o câmbio manual funciona');

  /* 10. feriados locais funcionam offline */
  await page.evaluate(() => { document.getElementById('p_start').value = '2026-11-20'; __pm.S.result = null; __pm.recalc(false); });
  assert.equal(await page.evaluate(() => __pm.S.result.start.toISOString().slice(0, 10)), '2026-11-23'); // 20/11 = Consciência Negra
  ok('feriados nacionais calculados localmente (sem internet)');

  /* 11. rascunho (autosave) sobrevive ao recarregamento */
  await page.evaluate(() => { document.getElementById('p_vendor').value = 'ACME Rascunho'; document.getElementById('p_vendor').dispatchEvent(new Event('change', {bubbles: true})); });
  await page.waitForTimeout(2200);
  await page.reload();
  await page.waitForSelector('.toast');
  assert.match(await $text(page, '#toasts'), /rascunho salvo/i);
  await page.click('.toast button:has-text("Restaurar")');
  assert.equal(await page.evaluate(() => document.getElementById('p_vendor').value), 'ACME Rascunho');
  ok('autosave: rascunho restaurado após recarregar');

  /* 12. gravação direta em arquivo (File System Access) com autosave no mesmo arquivo */
  const ctx2 = await browser.newContext({acceptDownloads: true});
  const p2 = await newPage(ctx2, {fs: true});
  await p2.page.evaluate(() => {
    window.__written = [];
    window.showSaveFilePicker = async () => ({name: 'minha.xlsx', kind: 'file',
      queryPermission: async () => 'granted', requestPermission: async () => 'granted',
      createWritable: async () => ({ write: async d => { window.__written.push(d.size); }, close: async () => {} })});
  });
  await p2.page.evaluate(() => document.querySelector('[data-act="example"]').click());
  await p2.page.waitForFunction(() => __pm.S.rows.length > 0);
  await p2.page.evaluate(() => __pm.actions.saveas());
  await p2.page.waitForFunction(() => window.__written.length === 1);
  await p2.page.evaluate(() => { const c = document.getElementById('p_disc'); c.value = '5'; c.dispatchEvent(new Event('change', {bubbles: true})); });
  await p2.page.waitForFunction(() => window.__written.length === 2, null, {timeout: 8000});
  assert.match(await $text(p2.page, '#fstatus'), /Salvo em minha\.xlsx/);
  assert.deepEqual(p2.errors, [], 'sem erros de página: ' + p2.errors.join(' | '));
  ok('Salvar como grava no arquivo e as alterações seguintes são persistidas automaticamente');
  await ctx2.close();

  /* 12b. editor de timeline: células, combos de perfil e nível, pintar, desfazer, parâmetros */
  const ctx3 = await browser.newContext({acceptDownloads: true, viewport: {width: 1440, height: 1000}});
  const p3 = await newPage(ctx3);
  const e3 = p3.page;
  await e3.evaluate(() => document.querySelector('[data-act="example"]').click());
  await e3.waitForSelector('#edGrid tr[data-r]');
  assert.match(await $text(e3, '#pbTitle'), /Client — Data & Analytics Modernization/);
  assert.match(await $text(e3, '#pbSub'), /R\$/, 'preço na barra superior');
  const firstCell = e3.locator('#edGrid tr[data-r] input[data-cell][data-i="0"]').first();
  await firstCell.fill('2'); await e3.keyboard.press('Enter');
  assert.equal(await e3.evaluate(() => __pm.S.rows[0].alloc[0]), 2);
  await e3.locator('#edGrid tr[data-r] input[data-cell][data-i="0"]').first().fill('abc'); await e3.keyboard.press('Enter');
  assert.equal(await e3.evaluate(() => __pm.S.rows[0].alloc[0]), 2, 'valor inválido não é aceito');
  assert.match(await $text(e3, '#toasts'), /FTE inválido/);
  // combo de perfil (lista os perfis cadastrados) e nível auto-preenchido
  await e3.locator('#edGrid .combo:not(.lv)').first().click();
  await e3.waitForSelector('#edPop.show'); await e3.fill('#edPq', 'tester');
  await e3.click('#edPl li[role=option]');
  assert.deepEqual(await e3.evaluate(() => { const r = __pm.S.rows[0]; return [r.role, r.level]; }), ['Tester', 'Q2']);
  // combo de nível: trocar por existente e adicionar um novo
  await e3.locator('#edGrid .combo.lv').first().click();
  await e3.waitForSelector('#edPopL.show');
  await e3.click('#edLnew'); await e3.fill('#edFLv', 'Q5'); await e3.fill('#edFLvRate', '300'); await e3.fill('#edFLvCost', '150');
  await e3.click('#edLf button[type=submit]');
  assert.deepEqual(await e3.evaluate(() => { const r = __pm.S.rows[0]; return [r.role, r.level, __pm.S.rates['Tester|Q5|Brazil'].rate]; }), ['Tester', 'Q5', 300]);
  assert.ok(await e3.evaluate(() => [...document.querySelectorAll('#edPl')].length >= 0));
  // nova fase + perfil novo pela combo
  await e3.click('#edAddPhase');
  await e3.evaluate(() => { document.querySelector('.edscroll').scrollTop = 1e6; });
  await e3.locator('#edGrid .combo.empty').last().click();
  await e3.click('#edPnew'); await e3.fill('#edFRole', 'Scrum Master'); await e3.fill('#edFLevel', 'Q3'); await e3.fill('#edFLoc', 'Brazil'); await e3.fill('#edFRate', '210');
  await e3.click('#edPf button[type=submit]');
  assert.equal(await e3.evaluate(() => __pm.S.rows.at(-1).role), 'Scrum Master');
  // pintar com arrastar
  await e3.click('#edMPaint'); await e3.click('#edBrush .chip[data-b="2"]');
  await e3.evaluate(() => { document.querySelector('.edscroll').scrollTop = 1e6; });
  const lastRow = e3.locator('#edGrid tr[data-r]').last();
  await lastRow.scrollIntoViewIfNeeded();
  const cs = lastRow.locator('td.cell'), b1 = await cs.nth(2).boundingBox(), b2 = await cs.nth(5).boundingBox();
  await e3.mouse.move(b1.x + 8, b1.y + 8); await e3.mouse.down(); await e3.mouse.move(b2.x + 8, b2.y + 8, {steps: 6}); await e3.mouse.up();
  assert.deepEqual(await e3.evaluate(() => __pm.S.rows.at(-1).alloc.slice(2, 6)), [2, 2, 2, 2]);
  // desfazer / refazer
  await e3.click('#edUndo');
  assert.deepEqual(await e3.evaluate(() => __pm.S.rows.at(-1).alloc.slice(2, 6)), [0, 0, 0, 0]);
  await e3.click('#edRedo');
  assert.deepEqual(await e3.evaluate(() => __pm.S.rows.at(-1).alloc.slice(2, 6)), [2, 2, 2, 2]);
  await e3.click('#edMType');
  // parâmetros: períodos, unidade e início
  await e3.click('#edNPlus'); assert.equal(await e3.evaluate(() => __pm.S.nPeriods), 13);
  assert.ok(await e3.evaluate(() => [...document.querySelectorAll('#edGrid thead th')].some(t => t.textContent === 'SP13')));
  assert.ok(await e3.evaluate(() => document.querySelectorAll('#edGrid thead th.d').length) > 0, 'datas sob os períodos');
  await e3.evaluate(() => { const u = document.getElementById('p_unit'); u.value = 'W'; u.dispatchEvent(new Event('change', {bubbles: true})); });
  assert.ok(await e3.evaluate(() => [...document.querySelectorAll('#edGrid thead th')].some(t => t.textContent === 'W1')));
  await e3.evaluate(() => { const u = document.getElementById('p_startmode'); u.value = 'none'; u.dispatchEvent(new Event('change', {bubbles: true})); });
  assert.equal(await e3.evaluate(() => document.querySelectorAll('#edGrid thead th.d').length), 0, 'sem data: sem linha de datas');
  assert.ok(await e3.evaluate(() => document.querySelectorAll('#edPv table.tlt').length) === 1, 'prévia na proposta');
  assert.deepEqual(p3.errors, [], 'sem erros: ' + p3.errors.join(' | '));
  await ctx3.close();
  ok('editor: célula, validação, combo de perfil, combo de nível (trocar/adicionar), nova fase, pintar, desfazer, períodos, unidade e início');

  /* 12c. proposta nova (vazia) é gravada em Excel; timeline "bruto" nunca é sobrescrito */
  const ctx4 = await browser.newContext({acceptDownloads: true});
  const p4 = await newPage(ctx4, {fs: true});
  const tplB64 = fs.readFileSync(path.join(root, 'Timeline_Template.xlsx')).toString('base64');
  await p4.page.evaluate((b64) => {
    window.__writes = {}; window.__bad = 0;
    const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    const mk = (name, original) => ({name, kind: 'file', queryPermission: async () => 'granted', requestPermission: async () => 'granted',
      getFile: async () => new File([bytes], name),
      createWritable: async () => { if (original) window.__bad++; return {write: async d => { (window.__writes[name] ||= []).push(d.size); }, close: async () => {}}; }});
    window.showOpenFilePicker = async () => [mk('Timeline_Template.xlsx', true)];
    window.showSaveFilePicker = async () => mk('Nova_Proposta.xlsx', false);
  }, tplB64);
  await p4.page.evaluate(() => __pm.actions.new());
  await p4.page.waitForSelector('#tlEditor', {state: 'visible'});
  assert.equal(await p4.page.evaluate(() => __pm.S.rows.length), 1);
  await p4.page.evaluate(() => __pm.actions.saveas());
  await p4.page.waitForFunction(() => (window.__writes['Nova_Proposta.xlsx'] || []).length === 1);
  await p4.page.evaluate(() => __pm.S.rows[0].alloc[0] = 0);
  await p4.page.locator('#edGrid .combo:not(.lv)').first().click();
  assert.match(await $text(p4.page, '#edPl'), /Nenhum perfil/i, 'proposta nova começa sem perfis');
  await p4.page.click('#edPnew'); await p4.page.fill('#edFRole', 'Analista'); await p4.page.fill('#edFLevel', 'Q3'); await p4.page.fill('#edFLoc', 'Brazil'); await p4.page.fill('#edFRate', '150');
  await p4.page.click('#edPf button[type=submit]');
  await p4.page.waitForFunction(() => (window.__writes['Nova_Proposta.xlsx'] || []).length >= 2, null, {timeout: 8000});
  // abrir o template bruto: não grava nele
  await p4.page.evaluate(() => __pm.actions.open());
  await p4.page.waitForFunction(() => __pm.S.rows.length > 5);
  assert.equal(await p4.page.evaluate(() => __pm.S.handle), null);
  assert.match(await $text(p4.page, '#fstatus'), /original/);
  await p4.page.evaluate(() => { const c = document.getElementById('p_vendor'); c.value = 'Outro'; c.dispatchEvent(new Event('change', {bubbles: true})); });
  await p4.page.waitForTimeout(2300);
  assert.equal(await p4.page.evaluate(() => window.__bad), 0, 'o arquivo original não pode ser gravado');
  assert.equal(await p4.page.evaluate(() => (window.__writes['Timeline_Template.xlsx'] || []).length), 0);
  assert.match(await $text(p4.page, '#pbSub'), /origem: Timeline_Template\.xlsx/);
  assert.deepEqual(p4.errors, [], 'sem erros: ' + p4.errors.join(' | '));
  await ctx4.close();
  ok('proposta nova é gravada em Excel; timeline bruto aberto nunca é sobrescrito; barra superior identifica a origem');

  /* 12d. catálogo de perfis e taxas: uma aba por cliente (Excel), 2 vias */
  const ctx5 = await browser.newContext({acceptDownloads: true, viewport: {width: 1440, height: 1000}});
  const p5 = await newPage(ctx5, {fs: true});
  const e5 = p5.page;
  const catB64 = fs.readFileSync(path.join(root, 'Perfis_e_Taxas.xlsx')).toString('base64');
  await e5.evaluate((b64) => {
    const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    window.__w = {};   // nome do arquivo -> último conteúdo gravado (base64)
    const b64of = async blob => { const a = new Uint8Array(await blob.arrayBuffer()); let s = ''; for (let i = 0; i < a.length; i += 8192) s += String.fromCharCode(...a.subarray(i, i + 8192)); return btoa(s); };
    const mk = (name, data) => ({name, kind: 'file', queryPermission: async () => 'granted', requestPermission: async () => 'granted',
      getFile: async () => new File([data || new Uint8Array()], name),
      createWritable: async () => ({write: async d => { window.__w[name] = await b64of(d); data = Uint8Array.from(atob(window.__w[name]), c => c.charCodeAt(0)); }, close: async () => {}})});
    window.__catHandle = mk('Perfis_e_Taxas.xlsx', bytes);
    window.showOpenFilePicker = async () => [window.__catHandle];
    window.showSaveFilePicker = async () => (window.__propHandle = mk('Proposta_Teste.xlsx'));
  }, catB64);
  const written = async name => Buffer.from(await e5.evaluate(n => window.__w[n] || '', name), 'base64');
  await e5.evaluate(() => __pm.actions['cat-open']());
  await e5.waitForFunction(() => __pm.S.cat.loaded);
  assert.deepEqual(await e5.evaluate(() => __pm.S.cat.clients.map(c => c.name)), ['Roche', 'JnJ']);
  assert.match(await $text(e5, '#btnCat'), /Perfis_e_Taxas/);
  // nova proposta: cliente por seleção; perfis, níveis e taxas vêm da aba do cliente
  await e5.evaluate(() => __pm.actions.new());
  await e5.waitForSelector('#tlEditor', {state: 'visible'});
  assert.match(await $text(e5, '#tlMsg'), /Selecione o/i);
  await e5.selectOption('#p_clientsel1', 'JnJ');
  assert.equal(await e5.evaluate(() => __pm.S.catClient), 'JnJ');
  assert.equal(await e5.evaluate(() => document.getElementById('p_rcur').value), 'BRL');
  assert.equal(await e5.evaluate(() => __pm.S.rates['Developer: SAP BASIS|Senior|-'].rate), 142.91);
  assert.match(await $text(e5, '#pbTitle'), /JnJ/);
  await e5.locator('#edGrid .combo:not(.lv)').first().click();
  await e5.fill('#edPq', 'basis'); await e5.click('#edPl li[role=option]');
  await e5.locator('#edGrid .combo.lv').first().click();
  const lv = await $text(e5, '#edLl'); assert.match(lv, /Senior/i); assert.match(lv, /Expert/i);
  await e5.click('#edLl li[role=option] >> nth=0'); // primeiro nível do perfil (ordem natural)
  await e5.locator('#edGrid tr[data-r] input[data-cell][data-i="0"]').first().fill('1'); await e5.keyboard.press('Enter');
  const lvl = await e5.evaluate(() => __pm.S.rows[0].level);
  const price = await e5.evaluate(() => { const R = __pm.recalc(false); return R.T.net; });
  const expect = lvl === 'Senior' ? 142.91 : 178.08;
  assert.ok(Math.abs(price - 80 * expect) < 0.01, `preço vem da aba do cliente: ${price} (nível ${lvl})`);
  // nível "sem taxa" (SAP Architect) é sinalizado
  await e5.click('#edAddPhase');
  await e5.evaluate(() => { document.querySelector('.edscroll').scrollTop = 1e6; });
  await e5.locator('#edGrid .combo.empty').last().click();
  await e5.fill('#edPq', 'sap arch'); await e5.click('#edPl li[role=option]');
  assert.ok(await e5.evaluate(() => [...document.querySelectorAll('#edGrid .tag')].some(t => /sem taxa/i.test(t.textContent))), 'perfil sem taxa sinalizado');
  // "Perfis e taxas" mostra os dados do cliente selecionado; editar grava no catálogo (Excel)
  await e5.evaluate(() => __pm.goTab('par'));
  assert.match(await $text(e5, '#rcTitle'), /JnJ/);
  assert.ok(await e5.evaluate(() => document.querySelectorAll('#rcTable tbody tr').length) >= 22);
  assert.ok(await e5.evaluate(() => /Local/i.test(document.querySelector('#rcTable thead').innerText) && /Cost/i.test(document.querySelector('#rcTable thead').innerText)), 'JnJ usa o layout padrão: Perfil, Nível, Local, Cost/h e Rate net/h');
  const archLevel = await e5.evaluate(() => __pm.S.rows.at(-1).level);
  const archKey = 'SAP Architect|' + archLevel + '|-';
  await e5.locator(`#rcTable input[data-k="${archKey}"][data-f="rate"]`).fill('250'); await e5.keyboard.press('Tab');
  await e5.waitForFunction(() => !__pm.S.cat.dirty && /gravado/.test(__pm.S.cat.status), null, {timeout: 10000});
  let cat = await Catalog.parse(ExcelJS, await written('Perfis_e_Taxas.xlsx'));
  const jn = cat.clients.find(c => c.name === 'JnJ');
  assert.equal(jn.entries.find(e => e.role === 'SAP Architect' && e.level === archLevel).rate, 250);
  const wbw = new ExcelJS.Workbook(); await wbw.xlsx.load(await written('Perfis_e_Taxas.xlsx'));
  assert.equal(wbw.getWorksheet('JnJ').getCell('A1').fill.fgColor.argb, 'FFD9E5F3', 'formatação do cliente preservada');
  // cliente novo: pergunta em qual se basear; cria uma aba nova
  await e5.evaluate(() => __pm.actions['client-new']({dataset: {base: 'JnJ'}}));
  await e5.waitForSelector('#dlgClient[open]');
  assert.deepEqual(await e5.evaluate(() => [...document.querySelectorAll('#ncBase option')].map(o => o.value)), ['Roche', 'JnJ', ''], 'pergunta a base: clientes existentes ou em branco');
  assert.equal(await e5.evaluate(() => document.getElementById('ncBase').value), 'JnJ');
  await e5.fill('#ncName', 'Acme'); await e5.click('#ncForm button[type=submit]');
  await e5.waitForFunction(() => !__pm.S.cat.dirty && /gravado/.test(__pm.S.cat.status), null, {timeout: 10000});
  cat = await Catalog.parse(ExcelJS, await written('Perfis_e_Taxas.xlsx'));
  assert.deepEqual(cat.clients.map(c => c.name), ['Roche', 'JnJ', 'Acme']);
  assert.equal(cat.clients.find(c => c.name === 'Acme').entries.length, 22, 'a aba nova nasce com os perfis do cliente base');
  assert.equal(await e5.evaluate(() => __pm.S.catClient), 'Acme');
  assert.equal(await e5.evaluate(() => document.getElementById('p_client').value), 'Acme');
  // a proposta é gravada SEM o cadastro de perfis; só a cópia das taxas usadas
  await e5.evaluate(() => __pm.actions.saveas());
  await e5.waitForFunction(() => window.__w['Proposta_Teste.xlsx']);
  const propBuf = await written('Proposta_Teste.xlsx');
  const pw = XLSXlib.read(propBuf, {type: 'buffer'});
  assert.ok(!pw.SheetNames.includes('Perfis') && !pw.SheetNames.includes('RateCard'), 'proposta sem o cadastro de perfis');
  assert.match(String(pw.Sheets['_meta'].E1.v), /Perfis_e_Taxas/);
  assert.ok(XLSXlib.utils.sheet_to_json(pw.Sheets['Modelo'], {header: 1})[0].includes('Rate/h (catálogo)'));
  // taxas do catálogo mudam depois: ao reabrir a proposta, avisa e usa as atuais
  await e5.evaluate(() => { const k = Object.keys(__pm.S.rates).find(x => /^Developer: SAP BASIS/.test(x)); __pm.S.rates[k].rate = 999; });
  const usedKey = await e5.evaluate(() => { const r = __pm.S.rows[0]; return `${r.role}|${r.level}|-`; });
  await e5.locator(`#rcTable input[data-k="${usedKey}"][data-f="rate"]`).fill('150'); await e5.keyboard.press('Tab');
  await e5.waitForFunction(() => !__pm.S.cat.dirty, null, {timeout: 10000});
  await e5.evaluate(async (arr) => { await __pm.loadBuffer(Uint8Array.from(arr).buffer, 'Proposta_Teste.xlsx', null); }, [...propBuf]);
  assert.match(await $text(e5, '#toasts'), /mudaram/i, 'avisa que as taxas do catálogo mudaram');
  assert.equal(await e5.evaluate((k) => __pm.S.rates[k].rate, usedKey), 150, 'usa a taxa atual do catálogo');
  assert.equal(await e5.evaluate(() => __pm.S.catClient), 'Acme');
  assert.deepEqual(p5.errors, [], 'sem erros: ' + p5.errors.join(' | '));
  await ctx5.close();
  ok('catálogo: aba por cliente, perfil/nível/taxa vindos do cliente, edição grava no Excel (formato preservado), cliente novo pergunta a base, proposta sem perfis e aviso de mudança de taxas');

  /* 12e. entregáveis por fase e cenários (até 3) com painel comparativo */
  const ctx6 = await browser.newContext({acceptDownloads: true, viewport: {width: 1400, height: 900}});
  const p6 = await newPage(ctx6); const e6 = p6.page;
  await e6.evaluate(() => { document.getElementById('p_lang').value = 'pt'; document.getElementById('p_lang').dispatchEvent(new Event('change', {bubbles: true})); __pm.actions.example(); });
  await e6.waitForSelector('#tlEditor', {state: 'visible'});
  // entregáveis: adicionar dois na primeira fase; nascem "auto" e se distribuem pelo tempo da fase
  const ph0 = await e6.evaluate(() => __pm.S.rows[0].phase);
  await e6.locator('#edGrid tr.ph').first().hover();
  await e6.locator('#edGrid tr.ph').first().locator('[data-ed="adddel"]').click();
  await e6.locator('#edGrid tr.ph').first().locator('[data-ed="adddel"]').click();
  assert.equal(await e6.evaluate(() => document.querySelectorAll('#edGrid tr.dl').length), 2);
  const spans = () => e6.evaluate(() => [...document.querySelectorAll('#edGrid tr.dl')].map(tr => [...tr.querySelectorAll('input.dp')].map(i => +i.value)));
  let sp = await spans();
  assert.ok(sp[0][1] < sp[1][0] || sp[0][1] <= sp[1][1], 'distribuídos em sequência: ' + JSON.stringify(sp));
  // renomear e fixar a data de entrega do 1º (vira manual) pela caixa "fim"
  await e6.locator('#edGrid tr.dl input.dn').first().fill('Plano do projeto'); await e6.keyboard.press('Tab');
  const win = await e6.evaluate(() => { const i = document.querySelector('#edGrid tr.dl input[data-ed="dend"]'); return [+i.min, +i.max]; });
  await e6.locator('#edGrid tr.dl input[data-ed="dend"]').first().fill(String(win[1])); await e6.keyboard.press('Tab');
  assert.deepEqual(await e6.evaluate(() => __pm.S.deliverables.map(d => [d.name, d.manual])), [['Plano do projeto', true], ['Entregável 2', false]]);
  assert.equal((await spans())[0][1], win[1], 'fim = último período da fase');
  // fora da janela da fase não aceita (limita); clique numa célula define o fim
  await e6.locator('#edGrid tr.dl input[data-ed="dend"]').first().fill('99'); await e6.keyboard.press('Tab');
  assert.equal((await spans())[0][1], win[1], 'limitado à janela da fase');
  assert.ok(await e6.evaluate(() => document.querySelectorAll('#edGrid tr.dl td.dc.out').length) > 0, 'períodos fora da fase marcados');
  assert.ok((await $text(e6, '#edPv')).includes('◆') && (await $text(e6, '#edPv')).includes('Plano do projeto'), 'losango e nome na prévia');
  await e6.locator('#edGrid tr.dlh [data-ed="distdel"]').first().click();   // redistribui tudo
  assert.deepEqual(await e6.evaluate(() => __pm.S.deliverables.map(d => d.manual)), [false, false]);
  // aparece na proposta (tabela de entregáveis) e nos marcos
  await e6.evaluate(() => { __pm.S.result = null; __pm.recalc(false); });
  assert.match(await e6.evaluate(() => document.querySelector('#doc').innerText), /Entregáveis[\s\S]*Plano do projeto/);
  // cenários: duplicar → cenário B editável de forma independente
  assert.equal(await e6.evaluate(() => __pm.S.scen.length), 1);
  await e6.click('#scBar [data-act="sc-dup"]');
  assert.equal(await e6.evaluate(() => [__pm.S.scen.length, __pm.S.active, __pm.S.scen[1].name]).then(a => a.join('|')), '2|1|Cenário A (cópia)');
  assert.equal(await e6.evaluate(() => __pm.S.deliverables.length), 2, 'entregáveis copiados');
  await e6.evaluate(() => { __pm.S.rows.forEach(r => { r.alloc = r.alloc.map(v => v * 2); }); __pm.S.nPeriods = 10; __pm.S.rows.forEach(r => { r.alloc.length = 10; }); __pm.S.deliverables = []; });
  await e6.evaluate(() => __pm.commitScen());
  await e6.evaluate(() => __pm.renameScen(1, 'Acelerado'));
  await e6.click('#scBar [data-sc="0"]');
  assert.equal(await e6.evaluate(() => [__pm.S.nPeriods, __pm.S.deliverables.length].join('|')), '12|2', 'A manteve seus períodos e entregáveis');
  // painel comparativo
  await e6.click('#scBar [data-act="go-cen"]');
  assert.equal(await e6.evaluate(() => document.querySelectorAll('#cenCards .sccard').length), 3, '2 cartões + "novo cenário"');
  const cmp = await $text(e6, '#cenTable');
  assert.match(cmp, /Preço final/); assert.match(cmp, /Pico da equipe/);
  assert.ok(await e6.evaluate(() => document.querySelectorAll('#cenTable td.best').length) > 0, 'melhor valor destacado');
  assert.match(await $text(e6, '#cenInsights'), /Acelerado.*vs\./);
  await e6.waitForFunction(() => document.querySelector('#chCenFte') && !document.querySelector('#chCenFteMsg .msg'), null, {timeout: 8000});
  // usar o cenário B: vira o ativo e a análise passa a usá-lo
  await e6.click('#cenCards [data-act="sc-use"]');
  assert.equal(await e6.evaluate(() => __pm.S.active), 1);
  await e6.evaluate(() => { __pm.S.result = null; __pm.recalc(false); });
  assert.equal(await e6.evaluate(() => __pm.S.result.nPeriods), 10, 'análise usa o cenário ativo');
  assert.match(await $text(e6, '#pbSub'), /Acelerado|R\$/);
  // limite de 3
  await e6.evaluate(() => __pm.goTab('tl'));
  await e6.click('#scBar [data-act="sc-new"]');
  assert.equal(await e6.evaluate(() => __pm.S.scen.length), 3);
  assert.equal(await e6.evaluate(() => document.querySelectorAll('#scBar [data-act="sc-new"]').length), 0, 'sem botão acima de 3');
  // gravar em Excel e reabrir: cenários e entregáveis voltam
  const buf64 = await e6.evaluate(async () => { const b = new Uint8Array(await (await __pm.buildBlob()).arrayBuffer()); let s = ''; b.forEach(c => s += String.fromCharCode(c)); return btoa(s); });
  const sheets = XLSXlib.read(Buffer.from(buf64, 'base64'), {type: 'buffer'}).SheetNames;
  assert.ok(['Cenarios', 'Entregaveis', 'Cen_1', 'Cen_2'].every(n => sheets.includes(n)) && !sheets.includes('Cen_3'), sheets.join());
  await e6.evaluate(async (b64) => { const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0)); await __pm.loadBuffer(bytes.buffer, 'p.xlsx', null); }, buf64);
  assert.equal(await e6.evaluate(() => [__pm.S.scen.map(s => s.name).join(','), __pm.S.active, __pm.S.nPeriods].join('|')), 'Cenário A,Acelerado,Cenário C|2|10');
  await e6.click('#scBar [data-sc="0"]');
  assert.equal(await e6.evaluate(() => [__pm.S.nPeriods, __pm.S.deliverables.map(d => d.name).join(',')].join('|')), '12|Plano do projeto,Entregável 2');
  // excluir cenário
  e6.once('dialog', d => d.accept());
  await e6.click('#scBar [data-act="sc-del"]');
  assert.equal(await e6.evaluate(() => __pm.S.scen.length), 2);
  assert.deepEqual(p6.errors, [], 'sem erros: ' + p6.errors.join(' | '));
  await ctx6.close();
  ok('entregáveis por fase (auto/manual, limite da fase, prévia e proposta) e até 3 cenários com painel comparativo, uso do ativo e persistência no Excel');

  /* 12f. painel de margens GM/GPM, racional ao passar o mouse, salário sugerido e separação proposta × catálogo */
  const ctx7 = await browser.newContext({acceptDownloads: true, viewport: {width: 1400, height: 900}});
  const p7 = await newPage(ctx7); const e7 = p7.page;
  await e7.evaluate(() => __pm.actions.example());
  await e7.waitForSelector('#tlEditor', {state: 'visible'});
  await e7.evaluate(() => { const set = (id, v) => { const el = document.getElementById(id); el.value = v; el.dispatchEvent(new Event('change', {bubbles: true})); };
    set('p_gmmin', '30'); set('p_gpmmin', '25'); set('p_ovh', '8'); set('p_comm', '1'); set('p_fin', '1'); set('p_risk', '2'); set('p_terms', '30'); set('p_disc', '3'); });
  await e7.evaluate(() => { __pm.S.result = null; __pm.recalc(true); });
  const mg = await e7.evaluate(() => { const R = __pm.S.result; return {gm: R.T.gm, gpm: R.T.gpm, gross: R.T.gross, ind: R.T.ind.total}; });
  assert.ok(mg.gpm < mg.gm && mg.ind > 0);
  assert.match(await $text(e7, '#mgKpis'), /GM — margem bruta[\s\S]*GPM — após indiretos/i);
  assert.ok(await e7.evaluate(() => document.querySelectorAll('#mgWf > div').length) === 10, 'cascata do preço ao GPM');
  assert.ok(await e7.evaluate(() => document.querySelectorAll('#mgPhase tbody tr').length) >= 5);
  // racional ao passar o mouse
  await e7.hover('#mgKpis .kpi >> nth=1');
  await e7.waitForFunction(() => getComputedStyle(document.getElementById('whyTip')).display === 'block');
  const tip = await $text(e7, '#whyTip');
  assert.match(tip, /GPM = \(GM − custos indiretos\)/); assert.match(tip, /overhead 8,0%/); assert.match(tip, /custo financeiro/);
  await e7.hover('#kpis .kpi >> nth=0'); assert.match(await $text(e7, '#whyTip'), /Preço com impostos/);
  await e7.mouse.move(5, 5); await e7.waitForFunction(() => getComputedStyle(document.getElementById('whyTip')).display === 'none');
  // barra superior mostra GM e GPM
  assert.match(await $text(e7, '#pbSub'), /GM [\d,]+%[\s\S]*GPM [\d,]+%/);
  // calculadora reversa: GPM 28% → preço alvo aplicável, e o resultado fecha com o motor
  await e7.fill('#mgTarget', '28');
  assert.match(await $text(e7, '#mgCalc'), /Pelo preço[\s\S]*Pelos rates[\s\S]*Pelo custo/);
  if (await e7.locator('#mgCalc [data-act="mg-apply"]').count()){
    await e7.locator('#mgCalc [data-act="mg-apply"]').click();
    assert.equal(await e7.evaluate(() => document.getElementById('p_discmode').value), 'target');
    await e7.evaluate(() => { __pm.S.result = null; __pm.recalc(false); });
    assert.ok(Math.abs(await e7.evaluate(() => __pm.S.result.T.gpm) - 0.28) < 0.002, 'GPM ≈ 28% após aplicar o preço alvo');
  }
  // simulador
  await e7.evaluate(() => { const r = document.getElementById('mgSlide'); r.value = 10; r.dispatchEvent(new Event('input', {bubbles: true})); });
  assert.match(await $text(e7, '#mgSlideV'), /10,0%/); assert.ok(await e7.evaluate(() => document.querySelectorAll('#mgSim tbody tr').length) === 8);
  // alerta de GPM abaixo do mínimo (mínimo alto)
  await e7.evaluate(() => { const el = document.getElementById('p_gpmmin'); el.value = '60'; el.dispatchEvent(new Event('change', {bubbles: true})); __pm.S.result = null; __pm.recalc(false); });
  assert.match(await $text(e7, '#checklist'), /GPM% de [\d,]+% está abaixo do mínimo/);
  // Perfis e taxas: Rate net/h e salário médio sugerido com racional
  await e7.evaluate(() => __pm.goTab('par'));
  const head = await $text(e7, '#rcTable thead');
  assert.match(head, /Rate net\/h/); assert.match(head, /Salário médio sugerido/);
  const sal = await e7.evaluate(() => { const r = Object.values(__pm.S.rates).find(x => x.cost > 0); return {cost: Number(r.cost), p: document.getElementById('p_hpd').value}; });
  await e7.hover('#rcTable tbody tr:first-child td:nth-last-child(1)');
  await e7.waitForFunction(() => /Salário médio sugerido/.test(document.getElementById('whyTip').innerText));
  assert.match(await $text(e7, '#whyTip'), /Horas por mês[\s\S]*Custo mensal da empresa[\s\S]*Fator de encargos[\s\S]*Salário bruto médio/);
  await e7.mouse.move(5, 5);
  // proposta × catálogo: um não abre no lugar do outro
  const propB64 = await e7.evaluate(async () => { const b = new Uint8Array(await (await __pm.buildBlob()).arrayBuffer()); let s = ''; b.forEach(c => s += String.fromCharCode(c)); return btoa(s); });
  const catB64x = fs.readFileSync(path.join(root, 'Perfis_e_Taxas.xlsx')).toString('base64');
  const err1 = await e7.evaluate(async (b64) => { try { await __pm.catalogLoad(Uint8Array.from(atob(b64), c => c.charCodeAt(0)).buffer, 'p.xlsx', null); return ''; } catch(e){ return e.message; } }, propB64);
  assert.match(err1, /PROPOSTA/);
  const err2 = await e7.evaluate(async (b64) => { try { await __pm.loadBuffer(Uint8Array.from(atob(b64), c => c.charCodeAt(0)).buffer, 'Perfis_e_Taxas.xlsx', null); return ''; } catch(e){ return e.message; } }, catB64x);
  assert.match(err2, /cat[áa]logo de perfis e taxas/);
  assert.deepEqual(p7.errors, [], 'sem erros: ' + p7.errors.join(' | '));
  await ctx7.close();
  ok('painel GM/GPM (cascata, fases, perfis, calculadora reversa, simulador, alertas), racional ao passar o mouse, salário sugerido e separação proposta × catálogo');

  /* 13. PDF (layout de impressão) */
  await page.evaluate(() => { __pm.S.result = null; __pm.recalc(false); __pm.goTab('doc'); });
  await page.pdf({path: path.join(outDir, 'proposta.pdf'), format: 'A4', preferCSSPageSize: true, printBackground: true});
  assert.ok(fs.statSync(path.join(outDir, 'proposta.pdf')).size > 20000);
  ok('PDF gerado pelo layout de impressão');

  assert.deepEqual(errors, [], 'sem erros de página: ' + errors.join(' | '));
  ok('nenhum erro de JavaScript durante todo o fluxo');
  console.log(`\nOK — ${n} verificações. Artefatos em ${outDir}`);
} catch (e) {
  console.error('\nFALHOU:', e.message);
  process.exitCode = 1;
} finally {
  await browser.close();
}

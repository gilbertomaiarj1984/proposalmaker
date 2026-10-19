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
  await page.waitForSelector('#tlView', {state: 'visible'});
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

  /* 7. salvar (download) e reabrir: o .xlsx é o "banco" da proposta */
  const [dl] = await Promise.all([page.waitForEvent('download'), page.evaluate(() => __pm.actions.save())]);
  const xlsxPath = path.join(outDir, 'proposta.xlsx'); await dl.saveAs(xlsxPath);
  assert.ok(fs.statSync(xlsxPath).size > 5000);
  const snapBefore = await page.evaluate(() => { const R = __pm.S.result; return {gross: R.T.gross, client: R.p.client, edits: Object.keys(__pm.S.edits).length}; });
  await page.reload(); // página nova: estado zerado
  await page.setInputFiles('#tlFile', xlsxPath);
  await page.waitForFunction(() => __pm.S.rows.length > 0);
  await page.evaluate(() => __pm.recalc(false));
  const snapAfter = await page.evaluate(() => { const R = __pm.S.result; return {gross: R.T.gross, client: R.p.client, edits: Object.keys(__pm.S.edits).length}; });
  assert.deepEqual(snapAfter, snapBefore);
  assert.match(await $text(page, '#tlMsg'), /Proposta .*proposta\.xlsx.* aberta/);
  ok('salvar em .xlsx e reabrir restaura parâmetros, rate card, impostos e textos editados');

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
  assert.match(await $text(p2.page, '#savebar'), /Salvo em minha\.xlsx/);
  assert.deepEqual(p2.errors, [], 'sem erros de página: ' + p2.errors.join(' | '));
  ok('Salvar como grava no arquivo e as alterações seguintes são persistidas automaticamente');
  await ctx2.close();

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

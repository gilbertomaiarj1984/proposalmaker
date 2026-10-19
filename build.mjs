// Monta gerador_proposta_comercial.html (arquivo único, sem dependências externas em tempo de execução).
//   node build.mjs             -> gera o HTML
//   node build.mjs --template  -> gera Timeline_Template.xlsx (mesma função do botão do app)
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { build as esbuild } from 'esbuild';

const root = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const read = f => fs.readFileSync(path.join(root, f), 'utf8');
const src = f => read(path.join('src', f));

if (process.argv.includes('--template')) {
  const { Workbook } = new Function(src('engine.js') + '\n' + src('workbook.js') + '\n;return {Workbook};')();
  const wb = await Workbook.buildTemplate(require('exceljs'));
  await wb.xlsx.writeFile(path.join(root, 'Timeline_Template.xlsx'));
  console.log('Timeline_Template.xlsx gerado');
  process.exit(0);
}

// texto de biblioteca seguro para ficar dentro de <script type="text/plain">
const safe = code => code
  .replace(/\/\/# sourceMappingURL=.*$/gm, '')
  .replace(/<\/script/gi, '<\\/script')
  .replace(/<!--/g, '<\\!--');

const docxBundle = (await esbuild({
  stdin: { contents: "export * from 'docx';", resolveDir: root, loader: 'js' },
  bundle: true, format: 'iife', globalName: 'docx', minify: true, platform: 'browser', write: false, legalComments: 'none'
})).outputFiles[0].text;

const libs = {
  xlsx:    fs.readFileSync(path.join(root, 'node_modules/xlsx/dist/xlsx.full.min.js'), 'utf8'),
  chart:   fs.readFileSync(path.join(root, 'node_modules/chart.js/dist/chart.umd.js'), 'utf8'),
  exceljs: fs.readFileSync(path.join(root, 'node_modules/exceljs/dist/exceljs.min.js'), 'utf8'),
  docx:    docxBundle,
  pptx:    fs.readFileSync(path.join(root, 'node_modules/pptxgenjs/dist/pptxgen.bundle.js'), 'utf8')
};
const libTags = Object.entries(libs).map(([k, v]) => `<script type="text/plain" data-lib="${k}">${safe(v)}</script>`).join('\n');

const scripts = ['engine.js', 'doc.js', 'workbook.js', 'docx-export.js', 'pptx-export.js', 'net.js', 'storage.js', 'editor.js', 'app.js'].map(f => `/* ===== ${f} ===== */\n${src(f)}`).join('\n');
const replace = (s, token, value) => { const i = s.indexOf(token); if (i < 0) throw new Error('token ausente: ' + token); return s.slice(0, i) + value + s.slice(i + token.length); };

let html = src('index.template.html');
html = replace(html, '/*STYLES*/', src('styles.css'));
html = replace(html, '<!--LIBS-->', libTags);
html = replace(html, '/*SCRIPTS*/', scripts);
fs.writeFileSync(path.join(root, 'gerador_proposta_comercial.html'), html);
console.log(`gerador_proposta_comercial.html: ${(Buffer.byteLength(html) / 1048576).toFixed(2)} MB`);

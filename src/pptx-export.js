/* =====================================================================
   EXPORT POWERPOINT (.pptx editável) — um slide por seção do documento.
   Tudo nativo do PowerPoint: texto, tabelas, formas e gráfico (sem imagens),
   a partir do mesmo modelo de blocos da tela/Word.
   build(PptxGenJS, doc, R, edits) -> instância; use pres.write({outputType:'blob'}).
===================================================================== */
const PptxExport = (() => {
'use strict';

const NAVY = '0B2545', BLUE = '1B4F8C', TEAL = '13A89E', LIGHT = 'EEF3F9', MUTED = '5F6B7A', TEXT = '1D2733', LINE = 'D7DFEA', WARN = 'B7791F', ICE = 'BFD2EA';
const FONT = 'Calibri';
const SW = 13.333, SH = 7.5, MX = 0.6, CW = SW - 2 * MX, TOP = 1.4, BOTTOM = 6.8;

const decode = s => String(s).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&amp;/g,'&');
const plain = html => decode(String(html ?? '').replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]*>/g, ''));
// mistura a cor com branco (f = 0..1 de branco)
const tint = (hex, f) => { const c = [0,2,4].map(i => parseInt(hex.slice(i, i + 2), 16)); return c.map(v => Math.round(v + (255 - v) * f).toString(16).padStart(2, '0')).join('').toUpperCase(); };

// HTML simples (b/i/u/br) -> runs do PptxGenJS
function runs(html, base){
  const out = []; let b = false, i = false, u = false;
  String(html ?? '').split(/(<\/?(?:b|strong|i|em|u|br)\s*\/?>)/gi).forEach(tok => {
    const m = tok.match(/^<(\/?)(b|strong|i|em|u|br)\s*\/?>$/i);
    if (m){
      const close = !!m[1], tag = m[2].toLowerCase();
      if (tag === 'br') out.push({text: '\n', options: {...base, breakLine: false}});
      else if (tag === 'b' || tag === 'strong') b = !close;
      else if (tag === 'i' || tag === 'em') i = !close;
      else u = !close;
      return;
    }
    const text = decode(tok.replace(/<[^>]*>/g, ''));
    if (text) out.push({text, options: {...base, bold: base.bold || b, italic: base.italic || i, underline: u ? {style: 'sng'} : undefined}});
  });
  return out.length ? out : [{text: '', options: {...base}}];
}

const estLines = (text, widthIn, fontPt) => Math.max(1, Math.ceil(String(text).length * (fontPt / 72) * 0.5 / Math.max(0.5, widthIn)));

function build(PptxGenJS, doc, R, edits){
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.theme = {headFontFace: FONT, bodyFontFace: FONT};
  pres.author = 'ProposalMaker';
  const cover0 = doc.blocks.find(b => b.t === 'cover');
  pres.title = cover0 ? plain(cover0.title) : 'Proposta';
  const foot = doc.blocks.find(b => b.t === 'foot');
  const ed = (key, dflt) => { const v = edits && edits[doc.lang + ':' + key]; return v !== undefined && v !== null ? v : dflt; };

  pres.defineSlideMaster({
    title: 'COVER', background: {color: NAVY}, objects: []
  });
  pres.defineSlideMaster({
    title: 'CONTENT', background: {color: 'FFFFFF'},
    objects: [
      {text: {text: foot ? foot.text : '', options: {x: MX, y: 7.0, w: 10.5, h: 0.3, fontSize: 10, color: MUTED, fontFace: FONT, margin: 0, objectName: 'Rodapé'}}},
      {placeholder: {options: {name: 'title', type: 'title', x: MX, y: 0.45, w: CW, h: 0.8, fontSize: 32, bold: true, color: NAVY, fontFace: FONT, margin: 0, valign: 'middle', align: 'left', fit: 'shrink'}, text: ''}}
    ],
    slideNumber: {x: SW - MX - 0.8, y: 7.0, w: 0.8, h: 0.3, fontSize: 10, color: MUTED, fontFace: FONT, align: 'right'}
  });

  /* ---------- capa ---------- */
  if (cover0){
    const s = pres.addSlide({masterName: 'COVER'});
    const title = plain(ed(cover0.titleKey, cover0.title)), client = plain(ed(cover0.clientKey, cover0.client));
    s.addText(cover0.tag.toUpperCase(), {x: 0.9, y: 1.6, w: 11.5, h: 0.4, fontSize: 14, color: ICE, fontFace: FONT, charSpacing: 6, margin: 0, isTextBox: true, objectName: 'Tipo do documento'});
    s.addText(title, {x: 0.9, y: 2.1, w: 11.5, h: 1.5, fontSize: title.length > 45 ? 36 : 44, bold: true, color: 'FFFFFF', fontFace: FONT, margin: 0, valign: 'top', fit: 'shrink', isTextBox: true, objectName: 'Título'});
    s.addText(client, {x: 0.9, y: 3.7, w: 11.5, h: 0.7, fontSize: 28, color: 'FFFFFF', fontFace: FONT, margin: 0, fit: 'shrink', isTextBox: true, objectName: 'Cliente'});
    const n = cover0.meta.length, mw = Math.min(3.3, 11.5 / n);
    cover0.meta.forEach((m, i) => {
      s.addText(m.k.toUpperCase(), {x: 0.9 + i * mw, y: 5.3, w: mw - 0.2, h: 0.3, fontSize: 11, color: ICE, fontFace: FONT, margin: 0, isTextBox: true, objectName: 'Meta ' + (i + 1) + ' rótulo'});
      s.addText(m.v, {x: 0.9 + i * mw, y: 5.62, w: mw - 0.2, h: 0.4, fontSize: 16, color: 'FFFFFF', fontFace: FONT, margin: 0, fit: 'shrink', isTextBox: true, objectName: 'Meta ' + (i + 1) + ' valor'});
    });
  }

  /* ---------- seções ---------- */
  let ctx = null, secTitle = '', secCount = 0;
  const newSlide = (cont) => {
    const s = pres.addSlide({masterName: 'CONTENT', sectionTitle: secTitle});
    s.addText(secTitle + (cont ? ' (cont.)' : ''), {placeholder: 'title', align: 'left'});
    ctx = {s, y: TOP}; return ctx;
  };
  const ensure = (h) => { if (!ctx || ctx.y + h > BOTTOM) newSlide(!!ctx); return ctx; };

  function textBlock(html, key){
    const size = 16, plainTxt = plain(html), lines = estLines(plainTxt, CW, size);
    const h = Math.min(BOTTOM - TOP, lines * 0.3 + 0.2);
    const c = ensure(h);
    c.s.addText(runs(html, {fontSize: size, color: TEXT, fontFace: FONT}), {x: MX, y: c.y, w: CW, h, margin: 0, valign: 'top', paraSpaceAfter: 6, fit: 'shrink', isTextBox: true, objectName: 'Texto'});
    c.y += h + 0.15;
  }
  function kpis(items){
    const gap = 0.25, w = (CW - gap * (items.length - 1)) / items.length, h = 1.3;
    const c = ensure(h);
    items.forEach((it, i) => {
      const x = MX + i * (w + gap);
      c.s.addShape('roundRect', {x, y: c.y, w, h, fill: {color: LIGHT}, line: {type: 'none'}, rectRadius: 0.08, objectName: 'Indicador ' + (i + 1)});
      c.s.addText(it.l.toUpperCase(), {x: x + 0.2, y: c.y + 0.15, w: w - 0.4, h: 0.3, fontSize: 11, color: MUTED, fontFace: FONT, margin: 0, isTextBox: true, objectName: 'Indicador ' + (i + 1) + ' rótulo'});
      c.s.addText(it.v, {x: x + 0.2, y: c.y + 0.5, w: w - 0.4, h: 0.65, fontSize: 24, bold: true, color: NAVY, fontFace: FONT, margin: 0, fit: 'shrink', isTextBox: true, objectName: 'Indicador ' + (i + 1) + ' valor'});
    });
    c.y += h + 0.3;
  }
  function note(text){
    const h = 0.4, c = ensure(h);
    c.s.addText(text, {x: MX, y: c.y, w: CW, h, fontSize: 12, italic: true, color: MUTED, fontFace: FONT, margin: 0, valign: 'top', isTextBox: true, objectName: 'Nota'});
    c.y += h + 0.05;
  }
  function sub(text){
    const h = 0.5, c = ensure(h + 0.8);
    c.s.addText(text, {x: MX, y: c.y, w: CW, h, fontSize: 18, bold: true, color: BLUE, fontFace: FONT, margin: 0, isTextBox: true, objectName: 'Subtítulo'});
    c.y += h + 0.05;
  }
  function bullets(html, key){
    const items = [...String(html).matchAll(/<li>([\s\S]*?)<\/li>/gi)].map(m => m[1]);
    const list = items.length ? items : [html];
    let batch = [], used = 0;
    const flush = () => {
      if (!batch.length) return;
      const c = ensure(used);
      const paras = [];
      batch.forEach((it, i) => runs(it, {fontSize: 16, color: TEXT, fontFace: FONT}).forEach((r, j, arr) =>
        paras.push({text: r.text, options: {...r.options, bullet: j === 0 ? {indent: 18} : undefined, paraSpaceAfter: 8, breakLine: j === arr.length - 1 && i < batch.length - 1}})));
      c.s.addText(paras, {x: MX, y: c.y, w: CW, h: used, margin: 0, valign: 'top', fit: 'shrink', isTextBox: true, objectName: 'Lista'});
      c.y += used + 0.1; batch = []; used = 0;
    };
    list.forEach(it => {
      const h = estLines(plain(it), CW - 0.3, 16) * 0.3 + 0.14;
      if (ctx && ctx.y + used + h > BOTTOM && batch.length){ flush(); ctx = null; }
      if (!ctx && used + h > BOTTOM - TOP && batch.length){ flush(); }
      batch.push(it); used += h;
    });
    flush();
  }

  // tabela nativa com paginação própria (cabeçalho repetido) e altura de linha estimada
  function table(b, opts){
    opts = opts || {};
    const head = b.head, weights = head.map(h => h.num ? 1.2 : (h.text.length <= 2 ? 0.4 : 2));
    const sumW = weights.reduce((a, c) => a + c, 0), W = opts.width || CW;
    const colW = weights.map(w => W * w / sumW);
    let fs = opts.fs || 12, minRow = opts.minRow || 0.4, k = fs / 12;
    const rowsData = b.rows.map(r => ({cells: r.cells, opt: r.cls === 'opt'}));
    if (b.total){
      const tc = b.total.cells;
      if (b.total.span){ const first = tc[0], last = tc[tc.length - 1]; rowsData.push({cells: [first, ...new Array(head.length - 2).fill(''), last], total: true, span: head.length - 1}); }
      else rowsData.push({cells: tc, total: true});
    }
    const hOf = r => Math.max(minRow, (Math.max(...r.cells.map((c, i) => estLines(c, colW[i] - 0.2, fs))) * 0.22 + 0.16) * k);
    let headH = Math.max(minRow + 0.02, (Math.max(...head.map((h, i) => estLines(h.text, colW[i] - 0.2, fs))) * 0.22 + 0.16) * k);
    if (opts.oneSlide){ // deve caber em um único slide: reduz fonte e altura das linhas até caber
      const availH = BOTTOM - (ctx ? ctx.y : TOP);
      const tot = () => headH + rowsData.reduce((a, r) => a + hOf(r), 0);
      for (let guard = 0; guard < 12 && tot() > availH && fs > 7; guard++){ fs -= 1; k = fs / 12; minRow = Math.max(0.2, minRow * 0.88); headH = Math.max(minRow + 0.02, headH * 0.9); }
    }
    const mkCell = (text, o) => ({text: String(text ?? ''), options: {fontFace: FONT, fontSize: fs, valign: 'middle', margin: [0.04, 0.1, 0.04, 0.1], ...o}});
    const totalH = headH + rowsData.reduce((a, r) => a + hOf(r), 0);
    if (!opts.fixedSlide && !opts.oneSlide && ctx && ctx.y + totalH > BOTTOM && totalH <= BOTTOM - TOP) newSlide(true);
    let i = 0;
    while (i < rowsData.length){
      const minNeeded = headH + hOf(rowsData[i]);
      const c = (opts.fixedSlide || (ctx && ctx.y + minNeeded <= BOTTOM)) ? ctx : newSlide(!!ctx);
      let y = opts.fixedTop != null ? opts.fixedTop : c.y, used = headH, batch = [];
      while (i < rowsData.length && (opts.oneSlide || y + used + hOf(rowsData[i]) <= BOTTOM || !batch.length)){ used += hOf(rowsData[i]); batch.push(rowsData[i]); i++; }
      const rows = [head.map((h, k) => mkCell(h.text, {bold: true, color: 'FFFFFF', fill: {color: NAVY}, align: h.num ? 'right' : 'left'}))];
      const heights = [headH];
      batch.forEach(r => {
        heights.push(hOf(r));
        if (r.span){
          rows.push([mkCell(r.cells[0], {bold: true, fill: {color: LIGHT}, color: NAVY, colspan: r.span}), mkCell(r.cells[r.cells.length - 1], {bold: true, fill: {color: LIGHT}, color: NAVY, align: 'right'})]);
        } else {
          rows.push(r.cells.map((cell, k) => mkCell(cell, {
            align: head[k] && head[k].num ? 'right' : 'left', bold: !!r.total, italic: r.opt, color: r.opt ? WARN : (r.total ? NAVY : TEXT),
            fill: r.total ? {color: LIGHT} : undefined,
            border: [{type: 'none'}, {type: 'none'}, {pt: 0.75, color: LINE}, {type: 'none'}]
          })));
        }
      });
      c.s.addTable(rows, {x: MX, y, w: W, colW, rowH: heights, autoPage: false, objectName: opts.name || 'Tabela'});
      if (opts.fixedSlide){ c.y = Math.max(c.y, y + used + 0.15); } else c.y = y + used + 0.2;
      opts.out && (opts.out.bottom = y + used);
    }
  }

  // timeline: tabela colorida (cada célula editável). SEMPRE em um único slide:
  // se houver muitas linhas, a altura da linha e a fonte diminuem até caber.
  // `companion` = tabela de fases/datas que vem logo depois; entra no mesmo slide se couber.
  function timeline(v, companion){
    const c = ctx || newSlide(false);
    const n = v.cols.length, labW = 2.9, cw = Math.min(0.85, (CW - labW) / n), W = labW + cw * n;
    const baseFs = n > 18 ? 10 : 11, border = {pt: 1, color: 'FFFFFF'};
    const nHdr = v.groups ? 2 : 1;
    const lines = v.rows.map(r => estLines(r.label + (r.optional && r.kind === 'phase' ? ' *' : ''), labW - 0.2, 11));
    const avail = BOTTOM - c.y;
    const prefH = lines.map(l => Math.max(0.32, l * 0.2 + 0.1)), prefHdr = 0.32 * nHdr;
    const need = prefH.reduce((a, b) => a + b, 0) + prefHdr;

    // a tabela de fases/datas entra no mesmo slide se, com a timeline compactada (linhas de 0.28), couber tudo
    let withTable = false;
    if (companion){
      const compactNeed = (v.rows.length + nHdr) * 0.28 + 0.2 + (companion.rows.length + 1 + (companion.total ? 1 : 0)) * 0.3 + 0.1;
      withTable = compactNeed <= avail;
    }
    const target = withTable ? (v.rows.length + nHdr) * 0.28 : avail;
    const scale = Math.min(1, target / need);
    const rowAvg = (scale >= 1 ? need : target) / (v.rows.length + nHdr);
    const clampN = (x, lo, hi) => Math.max(lo, Math.min(hi, x));
    const fs = scale >= 1 ? baseFs : clampN(Math.floor(rowAvg * 72 * 0.78), 6, baseFs);
    const labFs = scale >= 1 ? 11 : clampN(Math.floor(rowAvg * 72 * 0.78), 6, 11);
    const rowH = prefH.map(h => Math.max(0.1, h * scale)), hdrH = Math.max(0.12, 0.32 * scale);

    const hdr = [];
    if (v.groups){
      const g = [{text: '', options: {fill: {color: 'FFFFFF'}, fontSize: fs, fontFace: FONT, margin: 0, border: [{type: 'none'}, {type: 'none'}, {type: 'none'}, {type: 'none'}]}}];
      let i = 0; while (i < n){ let j = i; while (j < n && v.cols[j].group === v.cols[i].group) j++; g.push({text: v.cols[i].group, options: {colspan: j - i, align: 'center', valign: 'middle', bold: true, color: 'FFFFFF', fill: {color: NAVY}, fontSize: fs, fontFace: FONT, margin: 0, border}}); i = j; }
      hdr.push(g);
    }
    hdr.push([{text: '', options: {fill: {color: LIGHT}, fontSize: fs, fontFace: FONT, margin: 0}}, ...v.cols.map(col => ({text: col.label, options: {align: 'center', valign: 'middle', bold: true, color: NAVY, fill: {color: LIGHT}, fontSize: fs, fontFace: FONT, margin: 0, border}}))]);
    const data = v.rows.map((r, k) => {
      const base = r.color.replace('#', ''), fill = r.kind === 'role' ? tint(base, 0.35) : base;
      const lab = r.label + (r.optional && r.kind === 'phase' ? ' *' : '');
      return [
        {text: lab, options: {bold: r.kind === 'phase', italic: r.optional, color: r.kind === 'phase' ? NAVY : MUTED, fontSize: labFs, fontFace: FONT, valign: 'middle', margin: [0, 0.08, 0, r.kind === 'role' ? 0.25 : 0.08], border: [{type: 'none'}, {type: 'none'}, {pt: 0.5, color: LINE}, {type: 'none'}]}},
        ...r.cells.map(val => ({text: val > 0 && v.showFte ? String(Math.round(val * 10) / 10).replace('.', doc.lang === 'en' ? '.' : ',') : '', options: {align: 'center', valign: 'middle', bold: true, color: 'FFFFFF', fontSize: fs, fontFace: FONT, margin: 0, fill: {color: val > 0 ? fill : 'F0F3F7'}, border}}))
      ];
    });
    c.s.addTable([...hdr, ...data], {x: MX, y: c.y, w: W, colW: [labW, ...new Array(n).fill(cw)], rowH: [...new Array(nHdr).fill(hdrH), ...rowH], autoPage: false, objectName: 'Timeline'});
    c.y += rowH.reduce((a, b) => a + b, 0) + hdrH * nHdr + 0.2;
    return withTable;
  }

  // gráfico nativo de faturamento mensal (barras)
  function monthlyChart(s, x, y, w, h){
    if (!R || !R.months.length) return false;
    const labels = R.months.map(m => R.relative ? 'M' + m.idx : m.key);
    s.addChart('bar', [{name: doc.L.invest || 'Total', labels, values: R.months.map(m => Math.round(m.gross))}], {
      x, y, w, h, barDir: 'col', chartColors: [BLUE], showTitle: true, title: doc.L.monthlyBilling, titleFontSize: 14, titleColor: NAVY, titleFontFace: '+mn-lt',
      showLegend: false, valAxisLabelFontSize: 10, catAxisLabelFontSize: 10, valAxisLabelColor: MUTED, catAxisLabelColor: MUTED,
      valAxisLabelFontFace: '+mn-lt', catAxisLabelFontFace: '+mn-lt', valAxisLabelFormatCode: '#,##0',
      valGridLine: {color: LINE, size: 0.5}, catGridLine: {style: 'none'}, showValue: false, objectName: 'Faturamento mensal'
    });
    return true;
  }

  /* ---------- percorre os blocos: um slide por seção (continua se não couber) ---------- */
  const blocks = doc.blocks.filter(b => b.t !== 'cover' && b.t !== 'foot');
  let chartDone = false;
  for (let bi = 0; bi < blocks.length; bi++){
    const b = blocks[bi];
    switch (b.t){
      case 'h': secCount++; secTitle = `${b.n}. ${b.text}`; pres.addSection({title: secTitle}); ctx = null; newSlide(false); break;
      case 'p': textBlock(ed(b.key, b.html), b.key); break;
      case 'kpis': kpis(b.items); break;
      case 'note': note(b.text); break;
      case 'sub': sub(b.text); break;
      case 'ul': bullets(ed(b.key, b.items.map(i => `<li>${i}</li>`).join('')), b.key); break;
      case 'timeline': {
        const nxt = blocks[bi + 1], comp = nxt && nxt.t === 'table' && !nxt.isTimeline ? nxt : null;
        if (timeline(b.view, comp) && comp){ table(comp, {fs: 11, minRow: 0.3, fixedSlide: true, fixedTop: ctx.y, name: 'Fases e datas'}); bi++; }
        break;
      }
      case 'table': {
        // seção de investimento: tabela de fases + gráfico mensal lado a lado, se couber
        const isInvest = !chartDone && /^4\./.test(secTitle) && b.head.length === 5 && R && R.months.length && b.rows.length <= 8;
        if (isInvest){
          chartDone = true;
          const tw = 7.6, c = ensure(3.4), top = c.y;
          table(b, {width: tw, fixedSlide: true, fixedTop: top, name: 'Investimento por fase'});
          monthlyChart(c.s, MX + tw + 0.3, top - 0.05, CW - tw - 0.3, 3.9);
          c.y = Math.max(c.y, top + 3.9);
        } else table(b, b.isTimeline ? {oneSlide: true, name: 'Timeline (marcos)'} : undefined);
        break;
      }
    }
  }
  return pres;
}

return { build };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = PptxExport;

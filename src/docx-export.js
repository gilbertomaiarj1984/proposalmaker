/* =====================================================================
   EXPORT WORD (.docx de verdade) a partir do mesmo modelo de blocos da tela.
   build(docxLib, doc, edits) -> Document (empacote com docxLib.Packer).
===================================================================== */
const DocxExport = (() => {
'use strict';

const NAVY = '0B2545', BLUE = '1B4F8C', TEAL = '13A89E', LIGHT = 'EEF3F9', WARN = 'B7791F', MUTED = '5F6B7A';
const PAGE_W = 11906, PAGE_H = 16838, MARGIN = 1134, CONTENT_W = PAGE_W - 2 * MARGIN;

const decode = s => String(s).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#39;/g,"'").replace(/&amp;/g,'&');

const tintHex = h => { const n = [0, 2, 4].map(i => parseInt(h.substr(i, 2), 16)); return n.map(c => Math.round(c + (255 - c) * 0.65).toString(16).padStart(2, '0')).join('').toUpperCase(); };
function build(d, doc, edits){
  const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, Footer, PageNumber,
          BorderStyle, PageBreak, VerticalAlign, HeightRule, TableLayoutType } = d;
  const ed = (key, dflt) => { const v = edits && edits[doc.lang + ':' + key]; return v !== undefined && v !== null ? v : dflt; };

  // HTML simples (b/i/u/br) -> TextRun[]
  const runs = (html, base = {}) => {
    const out = []; let b = false, i = false, u = false;
    String(html ?? '').split(/(<\/?(?:b|strong|i|em|u|br)\s*\/?>)/gi).forEach(tok => {
      const m = tok.match(/^<(\/?)(b|strong|i|em|u|br)\s*\/?>$/i);
      if (m){
        const close = !!m[1], tag = m[2].toLowerCase();
        if (tag === 'br') out.push(new TextRun({break:1}));
        else if (tag === 'b' || tag === 'strong') b = !close;
        else if (tag === 'i' || tag === 'em') i = !close;
        else u = !close;
        return;
      }
      const text = decode(tok.replace(/<[^>]*>/g, ''));
      if (text) out.push(new TextRun({...base, text, bold: base.bold || b, italics: base.italics || i, underline: u ? {} : undefined}));
    });
    return out;
  };
  const para = (html, o = {}) => new Paragraph({children: runs(html, {size: o.size || 21, color: o.color, bold: o.bold, italics: o.italics}),
    spacing: {after: o.after ?? 120, line: 300}, alignment: o.align});
  const shade = hex => ({type: ShadingType.CLEAR, fill: hex, color: 'auto'});
  const none = {style: BorderStyle.NONE, size: 0, color: 'FFFFFF'};
  const hair = {style: BorderStyle.SINGLE, size: 4, color: 'D7DFEA'};
  const borders = {top: hair, bottom: hair, left: none, right: none};

  const cell = (text, w, o = {}) => new TableCell({
    width: {size: w, type: WidthType.DXA}, verticalAlign: VerticalAlign.CENTER, columnSpan: o.span,
    shading: o.fill ? shade(o.fill) : undefined, borders: o.borders || borders,
    margins: o.tight ? {top: 20, bottom: 20, left: 10, right: 10} : {top: 50, bottom: 50, left: 80, right: 80},
    children: [new Paragraph({alignment: o.right ? AlignmentType.RIGHT : (o.center ? AlignmentType.CENTER : AlignmentType.LEFT),
      children: [new TextRun({text: String(text ?? ''), bold: o.bold, italics: o.italics, color: o.color, size: o.size || 18})]})]
  });

  const kids = [];
  let first = true;

  for (const b of doc.blocks){
    switch (b.t){
      case 'cover': {
        const title = decode(ed(b.titleKey, b.title).replace(/<[^>]*>/g, '')), client = decode(ed(b.clientKey, b.client).replace(/<[^>]*>/g, ''));
        const white = {color: 'FFFFFF'};
        const metaPar = b.meta.map(m => new Paragraph({spacing:{after:80}, children:[new TextRun({text: m.k.toUpperCase() + '  ', size: 15, ...white, color: 'BFD2EA'}), new TextRun({text: m.v, size: 20, ...white})]}));
        kids.push(new Table({
          width: {size: CONTENT_W, type: WidthType.DXA}, columnWidths: [CONTENT_W], layout: TableLayoutType.FIXED,
          rows: [new TableRow({height: {value: 9000, rule: HeightRule.ATLEAST}, children: [new TableCell({
            width: {size: CONTENT_W, type: WidthType.DXA}, shading: shade(NAVY), verticalAlign: VerticalAlign.CENTER,
            borders: {top: none, bottom: none, left: none, right: none}, margins: {top: 600, bottom: 600, left: 600, right: 600},
            children: [
              new Paragraph({spacing:{after:200}, children:[new TextRun({text: b.tag.toUpperCase(), size: 20, characterSpacing: 60, color: 'BFD2EA'})]}),
              new Paragraph({spacing:{after:120}, children:[new TextRun({text: title, bold: true, size: 64, ...white})]}),
              new Paragraph({spacing:{after:600}, children:[new TextRun({text: client, size: 36, ...white})]}),
              ...metaPar]
          })]})]
        }));
        kids.push(new Paragraph({children: [new PageBreak()]}));
        break;
      }
      case 'h':
        kids.push(new Paragraph({spacing: {before: first ? 0 : 360, after: 160}, keepNext: true,
          border: {bottom: {style: BorderStyle.SINGLE, size: 12, color: TEAL, space: 4}},
          children: [new TextRun({text: `${b.n}. ${b.text}`, bold: true, size: 32, color: NAVY})]}));
        first = false;
        break;
      case 'p': kids.push(para(ed(b.key, b.html))); break;
      case 'note': kids.push(para(b.text, {size: 16, color: MUTED, italics: true})); break;
      case 'sub': kids.push(new Paragraph({spacing:{before: 200, after: 100}, keepNext: true, children:[new TextRun({text: b.text, bold: true, size: 24, color: BLUE})]})); break;
      case 'ul': {
        const html = ed(b.key, b.items.map(i => `<li>${i}</li>`).join(''));
        const items = [...String(html).matchAll(/<li>([\s\S]*?)<\/li>/gi)].map(m => m[1]);
        (items.length ? items : [html]).forEach(it => kids.push(new Paragraph({bullet: {level: 0}, spacing:{after: 60, line: 300}, children: runs(it, {size: 21})})));
        break;
      }
      case 'foot': break; // vai no rodapé
      case 'kpis': {
        const w = Math.floor(CONTENT_W / b.items.length);
        kids.push(new Table({width: {size: CONTENT_W, type: WidthType.DXA}, columnWidths: b.items.map(() => w), layout: TableLayoutType.FIXED,
          rows: [new TableRow({children: b.items.map(it => new TableCell({
            width: {size: w, type: WidthType.DXA}, shading: shade(LIGHT), margins: {top: 100, bottom: 100, left: 140, right: 100},
            borders: {top: none, bottom: none, left: {style: BorderStyle.SINGLE, size: 24, color: 'FFFFFF'}, right: none},
            children: [new Paragraph({children:[new TextRun({text: it.l.toUpperCase(), size: 14, color: MUTED})]}),
                       new Paragraph({children:[new TextRun({text: it.v, size: 26, bold: true, color: NAVY})]})]
          }))})]}));
        kids.push(new Paragraph({spacing:{after: 120}, children: []}));
        break;
      }
      case 'table': {
        const wts = b.head.map(h => h.num ? 1.2 : (h.text.length <= 2 ? 0.4 : 2));
        const sum = wts.reduce((a, c) => a + c, 0);
        const ws = wts.map(x => Math.floor(CONTENT_W * x / sum));
        const rows = [new TableRow({tableHeader: true, children: b.head.map((h, i) => cell(h.text, ws[i], {fill: NAVY, color: 'FFFFFF', bold: true, right: h.num, borders: {top: none, bottom: none, left: none, right: none}}))})];
        b.rows.forEach(r => rows.push(new TableRow({cantSplit: true, children: r.cells.map((c, i) => cell(c, ws[i], {right: b.head[i] && b.head[i].num, color: r.cls === 'opt' ? WARN : undefined, italics: r.cls === 'opt'}))})));
        if (b.total){
          const cs = b.total.cells;
          if (b.total.span){
            const wSpan = ws.slice(0, b.total.span).reduce((a, c) => a + c, 0);
            rows.push(new TableRow({children: [cell(cs[0], wSpan, {fill: LIGHT, bold: true, span: b.total.span}), cell(cs[cs.length - 1], ws[ws.length - 1], {fill: LIGHT, bold: true, right: true})]}));
          } else rows.push(new TableRow({children: cs.map((c, i) => cell(c, ws[i], {fill: LIGHT, bold: true, right: b.head[i] && b.head[i].num}))}));
        }
        kids.push(new Table({width: {size: CONTENT_W, type: WidthType.DXA}, columnWidths: ws, layout: TableLayoutType.FIXED, rows}));
        kids.push(new Paragraph({spacing:{after: 120}, children: []}));
        break;
      }
      case 'timeline': {
        const v = b.view, n = v.cols.length;
        const lab = 2300, cw = Math.max(200, Math.floor((CONTENT_W - lab) / n));
        const widths = [lab, ...v.cols.map(() => cw)];
        const small = n > 16 ? 12 : 14;
        const rows = [];
        if (v.groups){
          const cells = [cell('', lab, {fill: 'FFFFFF'})];
          let i = 0; while (i < n){ let j = i; while (j < n && v.cols[j].group === v.cols[i].group) j++; cells.push(cell(v.cols[i].group, cw * (j - i), {fill: NAVY, color: 'FFFFFF', bold: true, center: true, span: j - i, size: small, tight: true})); i = j; }
          rows.push(new TableRow({children: cells}));
        }
        rows.push(new TableRow({children: [cell('', lab, {fill: LIGHT}), ...v.cols.map(c => cell(c.label, cw, {fill: LIGHT, color: NAVY, bold: true, center: true, size: small, tight: true}))]}));
        v.rows.forEach(r => {
          const colr = r.color.replace('#', '');
          rows.push(new TableRow({cantSplit: true, children: [
            cell(r.label + (r.optional && r.kind === 'phase' ? ' *' : ''), lab, {bold: r.kind === 'phase', color: r.kind === 'phase' ? NAVY : MUTED, italics: r.optional, size: 16}),
            ...r.cells.map((val, ci) => r.kind === 'deliv' ? cell(ci === r.mark ? '◆' : '', cw, {fill: val > 0 ? tintHex(colr) : 'F0F3F7', color: colr, bold: true, center: true, size: small, tight: true, borders: {top: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}, bottom: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}, left: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}, right: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}}}) : cell(val > 0 && v.showFte ? String(Math.round(val * 10) / 10).replace('.', doc.lang === 'en' ? '.' : ',') : '', cw,
              {fill: val > 0 ? colr : 'F0F3F7', color: 'FFFFFF', bold: true, center: true, size: small, tight: true, borders: {top: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}, bottom: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}, left: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}, right: {style: BorderStyle.SINGLE, size: 6, color: 'FFFFFF'}}}))
          ]}));
        });
        kids.push(new Table({width: {size: lab + cw * n, type: WidthType.DXA}, columnWidths: widths, layout: TableLayoutType.FIXED, rows}));
        kids.push(new Paragraph({spacing:{after: 120}, children: []}));
        break;
      }
    }
  }

  const foot = doc.blocks.find(b => b.t === 'foot');
  return new Document({
    creator: 'ProposalMaker', title: doc.blocks[0] && doc.blocks[0].title,
    styles: {default: {document: {run: {font: 'Calibri', size: 21}}}},
    sections: [{
      properties: {page: {size: {width: PAGE_W, height: PAGE_H}, margin: {top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN}}},
      footers: {default: new Footer({children: [new Paragraph({alignment: AlignmentType.CENTER, children: [
        new TextRun({text: (foot ? foot.text + '  ·  ' : ''), size: 16, color: MUTED}),
        new TextRun({children: [PageNumber.CURRENT], size: 16, color: MUTED})]})]})},
      children: kids
    }]
  });
}

return { build };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = DocxExport;

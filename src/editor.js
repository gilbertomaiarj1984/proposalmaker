/* =====================================================================
   EDITOR DE TIMELINE — grade editável no navegador (fases, perfis do Rate Card
   por combo, FTE por período), com desfazer/refazer, modo pintar e prévia da
   proposta. Edita direto S.rows; a persistência (Excel) fica no app.js.
===================================================================== */
function TimelineEditor(ctx){
  'use strict';
  const {S, E, $, esc, nfPT, toast} = ctx;
  const PAL = ['#1B4F8C','#13A89E','#6C4AB6','#E07A1F','#2E86C1','#C0392B','#7D8C2E','#AD5A8C','#4A6B82','#D4A017'];
  const TL_IDS = ['p_unit', 'p_pw', 'p_hpd', 'p_dpw', 'p_startmode', 'p_start', 'p_startm'];
  let mode = 'type', brush = 1, painting = null, drag = null;
  const collapsed = new Set();
  const hist = [], fut = [];

  /* ---------------- utilidades ---------------- */
  const idOf = r => r._id || (r._id = S.rows.reduce((m, x) => Math.max(m, x._id || 0), 0) + 1);
  const fmtV = v => v ? String(Math.round(v * 100) / 100).replace('.', ',') : '';
  const fmtD = d => d.toLocaleDateString('pt-BR', {timeZone: 'UTC', day: '2-digit', month: '2-digit'});
  const fmtDY = d => d.toLocaleDateString('pt-BR', {timeZone: 'UTC', day: '2-digit', month: 'short', year: 'numeric'});
  const blank = (phase) => ({phase, activity: '', role: '', level: '', location: '', optional: false, alloc: new Array(S.nPeriods).fill(0)});
  function fit(){ S.rows.forEach(r => { while (r.alloc.length < S.nPeriods) r.alloc.push(0); if (r.alloc.length > S.nPeriods) r.alloc.length = S.nPeriods; idOf(r); }); }
  function groups(){
    const out = [], map = new Map();
    S.rows.forEach(r => { let g = map.get(r.phase); if (!g){ g = {name: r.phase, rows: []}; map.set(r.phase, g); out.push(g); } g.rows.push(r); });
    return out;
  }
  // janela ativa da fase (primeiro/último período com alocação) e entregáveis já posicionados (os não manuais se distribuem sozinhos)
  const phaseWin = rows => { const a = new Array(S.nPeriods).fill(0); rows.forEach(r => r.alloc.forEach((v, i) => { a[i] += v; })); const f = a.findIndex(v => v > 0); return f < 0 ? null : {first: f, last: a.length - 1 - [...a].reverse().findIndex(v => v > 0)}; };
  function placed(g){
    const w = phaseWin(g.rows); if (!w) return {w: null, list: []};
    let list = (S.deliverables || []).filter(d => d.phase === g.name);
    list = E.distributeDeliv(list, g.name, w.first, w.last).map(d => ({...d, ...E.clampDeliv(d, w.first, w.last)}));
    return {w, list};
  }
  const nextDelId = () => (S.deliverables || []).reduce((m, d) => Math.max(m, d.id || 0), 0) + 1;
  const flatten = gs => { S.rows = gs.flatMap(g => g.rows); };
  const catalog = () => Object.values(S.rates).filter(r => r.role).sort((a, b) => a.role.localeCompare(b.role) || a.level.localeCompare(b.level));
  const missingRate = (r, p) => p.mode === 'rate' ? !(E.num(r.rate) > 0) : !(E.num(r.cost) > 0);
  const rateOf = (row) => S.rates[E.rkey(row)];

  /* ---------------- histórico ---------------- */
  const tlRaw = () => { const o = {}; TL_IDS.forEach(id => { const el = $(id); if (el) o[id] = el.value; }); return o; };
  const snap = () => JSON.stringify({rows: S.rows, nPeriods: S.nPeriods, unit: S.unit, deliverables: S.deliverables, raw: tlRaw()});   // as taxas vivem no catálogo: desfazer não as altera
  function pushHist(){ hist.push(snap()); if (hist.length > 100) hist.shift(); fut.length = 0; }
  function restore(str){
    const o = JSON.parse(str);
    S.rows = o.rows; S.nPeriods = o.nPeriods; S.unit = o.unit; S.deliverables = o.deliverables || []; ctx.setRaw(o.raw);
    changed(); render();
  }
  function undo(){ if (!hist.length) return; fut.push(snap()); restore(hist.pop()); }
  function redo(){ if (!fut.length) return; hist.push(snap()); restore(fut.pop()); }
  function resetHistory(){ hist.length = 0; fut.length = 0; collapsed.clear(); }
  function changed(){ ctx.changed(); }
  function mutate(fn, keepFocus){ pushHist(); fn(); fit(); changed(); render(keepFocus); }

  /* ---------------- render ---------------- */
  function render(keepFocus){
    if (!S.rows.length) return;
    fit();
    const f = keepFocus && document.activeElement && document.activeElement.dataset && document.activeElement.dataset.cell ? {...document.activeElement.dataset} : null;
    const p = ctx.P(), N = S.nPeriods, U = S.unit;
    const cal = E.buildCalendar(p, N, U, {today: new Date()});
    const gs = groups(), colTot = new Array(N).fill(0);
    const hasDates = !cal.relative;

    let h = `<thead><tr><th class="l c1">Fase / Perfil</th><th class="l c2">Nível</th><th class="l c3">Local</th><th class="l c4" title="Item opcional">Opc.</th>`;
    for (let i = 1; i <= N; i++) h += `<th>${U}${i}</th>`;
    h += `<th>Total</th><th></th></tr>`;
    if (hasDates) h += `<tr><th class="l c1 d" style="text-align:left;padding-left:10px">início do período</th><th class="l c2 d"></th><th class="l c3 d"></th><th class="l c4 d"></th>${Array.from({length: N}, (_, i) => `<th class="d">${fmtD(cal.pStart(i))}</th>`).join('')}<th class="d"></th><th class="d"></th></tr>`;
    h += `</thead><tbody>`;

    gs.forEach((g, gi) => {
      const color = PAL[gi % PAL.length], sum = new Array(N).fill(0);
      g.rows.forEach(r => r.alloc.forEach((v, i) => { sum[i] += v; colTot[i] += v; }));
      const isCol = collapsed.has(g.name), allOpt = g.rows.length && g.rows.every(r => r.optional);
      h += `<tr class="ph"><td class="l c1"><div class="name">
        <button class="caret" data-ed="toggle" data-g="${esc(g.name)}" aria-label="${isCol ? 'Expandir' : 'Recolher'} fase">${isCol ? '▶' : '▼'}</button>
        <span class="dot" style="background:${color}"></span>
        <input class="pn" value="${esc(g.name)}" data-ed="pname" data-g="${esc(g.name)}" aria-label="Nome da fase" title="${esc(g.name)}">
      </div></td><td class="l c2"></td><td class="l c3"></td><td class="l c4"><span class="hint">${allOpt ? 'opc.' : ''}</span></td>`;
      sum.forEach(v => h += `<td class="cell" style="${v ? `background:${color}` : ''}">${v ? nfPT(v, 1) : ''}</td>`);
      h += `<td class="sum">${nfPT(sum.reduce((a, b) => a + b, 0), 1)}</td>
        <td><div class="row-actions"><button class="xbtn" title="Adicionar perfil à fase" aria-label="Adicionar perfil" data-ed="addrow" data-g="${esc(g.name)}">＋</button><button class="xbtn" title="Adicionar entregável à fase" aria-label="Adicionar entregável" data-ed="adddel" data-g="${esc(g.name)}">◆</button><button class="xbtn" title="Excluir fase" aria-label="Excluir fase" data-ed="delphase" data-g="${esc(g.name)}">🗑</button></div></td></tr>`;
      if (isCol) return;
      g.rows.forEach(r => {
        const rid = idOf(r), has = !!r.role, rt = has ? rateOf(r) : null, noRate = has && (!rt || missingRate(rt, p));
        const tot = r.alloc.reduce((a, b) => a + b, 0);
        h += `<tr data-r="${rid}"><td class="l c1 ${r.optional ? 'opt-on' : ''}" style="padding-left:18px">
          <button class="combo ${has ? '' : 'empty'}" data-ed="combo" data-r="${rid}" aria-haspopup="listbox" aria-label="Perfil da linha: ${has ? esc(r.role) : 'não escolhido'}">
            <span class="nm">${has ? esc(r.role) : 'Escolher perfil…'}</span>${noRate ? '<span class="tag" title="Sem taxa no Rate Card">sem taxa</span>' : ''}<span class="arrow">▼</span></button></td>
          <td class="l c2" style="padding:0">${has ? `<button class="combo lv" data-ed="lvcombo" data-r="${rid}" aria-haspopup="listbox" aria-label="Nível da linha: ${esc(r.level || 'sem nível')}"><span class="nm">${esc(r.level || '—')}</span><span class="arrow">▼</span></button>` : '<span style="padding:0 8px;color:var(--muted)">—</span>'}</td>
          <td class="l c3" style="padding:0 6px;color:var(--muted)">${has ? esc(r.location) : '—'}</td>
          <td class="l c4"><input class="opt-box" type="checkbox" ${r.optional ? 'checked' : ''} data-ed="opt" data-r="${rid}" aria-label="Item opcional"></td>`;
        r.alloc.forEach((v, i) => {
          h += `<td class="cell ${v ? '' : 'empty'}" data-r="${rid}" data-i="${i}" style="${v ? `background:${r.optional ? color + '99' : color}` : ''}"><input value="${fmtV(v)}" data-cell="1" data-r="${rid}" data-i="${i}" ${mode === 'paint' ? 'readonly tabindex="-1"' : ''} aria-label="FTE ${U}${i + 1}" inputmode="decimal"></td>`;
        });
        h += `<td class="sum">${nfPT(tot, 1)}</td>
          <td><div class="row-actions"><button class="xbtn" title="Duplicar linha" aria-label="Duplicar linha" data-ed="dup" data-r="${rid}">⧉</button><button class="xbtn" title="Excluir linha" aria-label="Excluir linha" data-ed="delrow" data-r="${rid}">×</button></div></td></tr>`;
      });
      // entregáveis da fase: posicionados dentro da janela ativa da fase
      const pl = placed(g);
      if (pl.w && pl.list.length){
        h += `<tr class="dlh"><td class="l c1" colspan="4" style="padding:2px 8px 2px 30px;font-size:11px;color:var(--muted)">Entregáveis da fase · ${pl.list.some(d => d.manual) ? 'alguns com datas manuais' : 'distribuídos igualmente'} <button class="btn sec sm" style="padding:1px 8px;font-size:11px" data-ed="distdel" data-g="${esc(g.name)}" title="Redistribui todos os entregáveis igualmente no tempo da fase">⇄ distribuir igualmente</button></td>${'<td></td>'.repeat(N)}<td></td><td></td></tr>`;
        pl.list.forEach(d => {
          h += `<tr class="dl" data-d="${d.id}"><td class="l c1" style="padding-left:30px"><div class="name" style="width:240px"><span class="dm" style="color:${color}">◆</span><input class="dn" value="${esc(d.name)}" data-ed="dname" data-d="${d.id}" placeholder="Nome do entregável" aria-label="Nome do entregável"></div></td>
            <td class="l c2" style="padding:0 2px"><input class="dp" type="number" min="${pl.w.first + 1}" max="${pl.w.last + 1}" value="${d.start + 1}" data-ed="dstart" data-d="${d.id}" aria-label="Início do entregável (${U})" title="Início (${U}${pl.w.first + 1}–${U}${pl.w.last + 1})"></td>
            <td class="l c3" style="padding:0 2px"><input class="dp" type="number" min="${pl.w.first + 1}" max="${pl.w.last + 1}" value="${d.end + 1}" data-ed="dend" data-d="${d.id}" aria-label="Fim do entregável (${U})" title="Fim (${U}${pl.w.first + 1}–${U}${pl.w.last + 1}) — a data de entrega"></td>
            <td class="l c4"><span class="hint" title="${d.manual ? 'Datas ajustadas manualmente' : 'Distribuído automaticamente'}">${d.manual ? 'manual' : 'auto'}</span></td>`;
          for (let i = 0; i < N; i++){
            const inW = i >= pl.w.first && i <= pl.w.last, on = i >= d.start && i <= d.end;
            h += `<td class="cell dc ${inW ? '' : 'out'}" ${inW ? `data-d="${d.id}" data-i="${i}"` : ''} style="${on ? `background:${color}55;` : ''}color:${color}" title="${inW ? 'Clique: define o fim · Shift+clique: define o início' : 'Fora do período da fase'}">${i === d.end ? '◆' : ''}</td>`;
          }
          h += `<td class="sum"></td><td><div class="row-actions"><button class="xbtn" title="Excluir entregável" aria-label="Excluir entregável" data-ed="deldel" data-d="${d.id}">×</button></div></td></tr>`;
        });
      } else if (!pl.w && (S.deliverables || []).some(d => d.phase === g.name)){
        h += `<tr class="dlh"><td class="l c1" colspan="4" style="padding:2px 8px 2px 30px;font-size:11px;color:var(--err)">Há entregáveis nesta fase, mas ela ainda não tem alocação.</td>${'<td></td>'.repeat(N)}<td></td><td></td></tr>`;
      }
    });
    const peak = Math.max(0, ...colTot);
    h += `<tr class="foot"><td class="l c1">Total FTE por período</td><td class="l c2"></td><td class="l c3"></td><td class="l c4"></td>${colTot.map(v => `<td>${v ? nfPT(v, 1) : '–'}</td>`).join('')}<td>${nfPT(peak, 1)} pico</td><td></td></tr></tbody>`;
    $('edGrid').innerHTML = h;
    $('edN').textContent = N;
    $('edMType').setAttribute('aria-pressed', mode === 'type'); $('edMPaint').setAttribute('aria-pressed', mode === 'paint');
    $('edBrush').style.display = mode === 'paint' ? 'inline-flex' : 'none';
    $('edUndo').disabled = !hist.length; $('edRedo').disabled = !fut.length;
    const nCat = catalog().length, nMiss = catalog().filter(r => missingRate(r, p)).length;
    $('edRc').innerHTML = `Perfis do cliente: <b>${nCat}</b>` + (nMiss ? ` · <span style="color:var(--err)">${nMiss} sem taxa</span>` : '');
    $('p_pw').disabled = U === 'W';
    $('lbl_start').style.display = $('p_startmode').value === 'date' ? '' : 'none';
    $('lbl_startm').style.display = $('p_startmode').value === 'month' ? '' : 'none';
    summary(p, cal, colTot);
    preview();
    if (f){ const el = document.querySelector(`#edGrid input[data-cell][data-r="${f.r}"][data-i="${f.i}"]`); if (el){ el.focus(); } }
  }

  function summary(p, cal, colTot){
    const N = S.nPeriods, U = S.unit, hpp = cal.pw * p.dpw * p.hpd;
    let last = -1, first = -1; colTot.forEach((v, i) => { if (v > 0){ if (first < 0) first = i; last = i; } });
    const hours = S.rows.filter(r => r.role).reduce((a, r) => a + r.alloc.reduce((x, y) => x + y, 0), 0) * hpp;
    const active = last < 0 ? 0 : last - first + 1;
    const kp = (b, s, w) => `<div class="kp"${w ? ` data-why="${w.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/\n/g, '&#10;')}"` : ''}><b>${b}</b>${s}</div>`;
    $('edSum').innerHTML =
      kp(`${active || N} ${U === 'SP' ? 'sprints' : 'semanas'}`, active ? `${active * cal.pw} semanas com alocação · ${N} no grid` : `${N * cal.pw} semanas no grid`, `**Duração**\nPeríodos entre o primeiro e o último com alocação (${active || N}) × ${cal.pw} semana(s) = ${(active || N) * cal.pw} semanas.`) +
      kp(cal.relative ? 'relativo' : fmtDY(last < 0 ? cal.effStart : cal.pStart(first)), 'início' + (cal.startAdjusted ? ' (ajustado p/ dia útil)' : '')) +
      kp(cal.relative ? '—' : fmtDY(cal.pEnd(last < 0 ? N - 1 : last)), 'fim previsto') +
      kp(`${nfPT(Math.round(hours))} h`, 'esforço no timeline', `**Esforço**\nΣ FTE (${nfPT(hours / hpp, 1)}) × horas por período por FTE (${cal.pw} sem × ${p.dpw} d × ${p.hpd} h = ${hpp} h) = ${nfPT(Math.round(hours))} h.`) +
      kp(`${nfPT(Math.max(0, ...colTot), 1)} FTE`, 'pico do time', '**Pico**\nMaior soma de FTE em um único período (linha “Total FTE por período”).');
  }
  function preview(){
    const box = $('edPv'), R = ctx.compute();
    if (!R){ box.innerHTML = '<p class="hint" style="margin:6px 0">Escolha um perfil e preencha o FTE para ver a prévia da proposta.</p>'; return; }
    const tl = R.p.tl, view = E.timelineView(R, {gran: tl.gran, mode: tl.mode === 'milestones' ? 'bars' : tl.mode, fte: tl.fte});
    box.innerHTML = DocModel.toHtml({lang: R.p.lang, blocks: [{t: 'timeline', view}]}, {});
  }

  /* ---------------- ações da grade ---------------- */
  const rowById = id => S.rows.find(r => String(idOf(r)) === String(id));
  const groupRows = name => S.rows.filter(r => r.phase === name);
  $('edGrid').addEventListener('click', e => {
    const b = e.target.closest('[data-ed]'); if (!b) return;
    const a = b.dataset.ed, g = b.dataset.g, r = b.dataset.r ? rowById(b.dataset.r) : null;
    if (a === 'toggle'){ collapsed.has(g) ? collapsed.delete(g) : collapsed.add(g); render(); }
    else if (a === 'addrow') mutate(() => { collapsed.delete(g); const gs = groups(); gs.find(x => x.name === g).rows.push(blank(g)); flatten(gs); });
    else if (a === 'delphase'){
      const rows = groupRows(g);
      if (rows.some(x => x.alloc.some(v => v > 0)) && !confirm(`Excluir a fase "${g}" e suas ${rows.length} linha(s)?`)) return;
      mutate(() => { S.rows = S.rows.filter(x => x.phase !== g); S.deliverables = (S.deliverables || []).filter(d => d.phase !== g); });
    }
    else if (a === 'delrow') mutate(() => { S.rows.splice(S.rows.indexOf(r), 1); });
    else if (a === 'dup') mutate(() => { S.rows.splice(S.rows.indexOf(r) + 1, 0, {...r, _id: undefined, alloc: r.alloc.slice()}); });
    else if (a === 'adddel') mutate(() => { collapsed.delete(g); const n = (S.deliverables || []).filter(d => d.phase === g).length + 1; S.deliverables = [...(S.deliverables || []), {id: nextDelId(), phase: g, name: `Entregável ${n}`, start: 0, end: 0, manual: false}]; });
    else if (a === 'distdel') mutate(() => { S.deliverables = S.deliverables.map(d => d.phase === g ? {...d, manual: false} : d); });
    else if (a === 'deldel') mutate(() => { S.deliverables = S.deliverables.filter(d => String(d.id) !== b.dataset.d); });
    else if (a === 'combo') openCombo(b, r);
    else if (a === 'lvcombo') openLevel(b, r);
  });
  $('edGrid').addEventListener('change', e => {
    const t = e.target, a = t.dataset.ed;
    if (a === 'opt'){ const r = rowById(t.dataset.r); mutate(() => { r.optional = t.checked; }); }
    else if (a === 'pname'){
      const old = t.dataset.g, name = t.value.trim();
      if (!name || name === old){ t.value = old; return; }
      mutate(() => { S.rows.forEach(r => { if (r.phase === old) r.phase = name; }); if (collapsed.delete(old)) collapsed.add(name); S.deliverables = (S.deliverables || []).map(d => d.phase === old ? {...d, phase: name} : d); if (S.phaseDisc && S.phaseDisc[old] !== undefined){ S.phaseDisc[name] = S.phaseDisc[old]; delete S.phaseDisc[old]; } });
    }
    else if (a === 'dname'){ const nm = t.value.trim() || 'Entregável'; mutate(() => { S.deliverables = S.deliverables.map(d => String(d.id) === t.dataset.d ? {...d, name: nm} : d); }, true); }
    else if (a === 'dstart' || a === 'dend') setDelSpan(t.dataset.d, a === 'dstart' ? 'start' : 'end', Math.round(E.num(t.value)) - 1);
    else if (t.dataset.cell){
      const r = rowById(t.dataset.r), i = +t.dataset.i, txt = String(t.value).trim();
      let v = 0;
      if (txt !== ''){
        v = E.num(txt);
        if (!/^[\d.,\s%]+$/.test(txt) || !isFinite(v) || v < 0 || v > 99){ toast('FTE inválido: use um número entre 0 e 99 (ex.: 1, 0,5 ou 50%).', 'warn', 4500); render(true); return; }
      }
      v = Math.round(v * 100) / 100;
      if ((r.alloc[i] || 0) === v){ t.value = fmtV(v); return; }
      mutate(() => { r.alloc[i] = v; }, true);
    }
  });
  $('edGrid').addEventListener('keydown', e => {
    const t = e.target; if (!t.dataset || !t.dataset.cell) return;
    const rows = [...document.querySelectorAll('#edGrid tr[data-r]')], ri = rows.findIndex(r => r.dataset.r === t.dataset.r), i = +t.dataset.i;
    const go = (r, c) => { const row = rows[r]; if (!row) return; const el = row.querySelector(`input[data-cell][data-i="${c}"]`); if (el){ el.focus(); el.select(); } };
    if (e.key === 'Enter'){ e.preventDefault(); const next = Math.min(rows.length - 1, ri + 1); t.blur(); setTimeout(() => go(next, i), 0); }
    else if (e.key === 'ArrowDown'){ e.preventDefault(); go(ri + 1, i); }
    else if (e.key === 'ArrowUp'){ e.preventDefault(); go(ri - 1, i); }
    else if (e.key === 'ArrowRight' && t.selectionStart === t.value.length){ e.preventDefault(); go(ri, i + 1); }
    else if (e.key === 'ArrowLeft' && t.selectionStart === 0){ e.preventDefault(); go(ri, i - 1); }
    else if (e.key === 'Delete'){ t.value = ''; t.dispatchEvent(new Event('change', {bubbles: true})); }
    else if (e.key === 'Escape'){ render(true); }
  });
  $('edGrid').addEventListener('focusin', e => { if (e.target.dataset && e.target.dataset.cell) e.target.select(); });

  // ajusta início/fim de um entregável (vira "manual"); mantém início ≤ fim e dentro da janela da fase
  function setDelSpan(id, which, idx){
    const g = groups().find(x => (S.deliverables || []).some(d => String(d.id) === String(id) && d.phase === x.name)); if (!g) return;
    const pl = placed(g), cur = pl.list.find(d => String(d.id) === String(id)); if (!cur || !pl.w) return;
    let s0 = cur.start, e0 = cur.end; idx = Math.max(pl.w.first, Math.min(pl.w.last, idx));
    if (which === 'start'){ s0 = idx; if (e0 < s0) e0 = s0; } else { e0 = idx; if (s0 > e0) s0 = e0; }
    mutate(() => { S.deliverables = S.deliverables.map(d => String(d.id) === String(id) ? {...d, start: s0, end: e0, manual: true} : d); });
  }
  $('edGrid').addEventListener('click', e => {
    const td = e.target.closest('td.dc[data-d]'); if (!td) return;
    setDelSpan(td.dataset.d, e.shiftKey ? 'start' : 'end', +td.dataset.i);
  });

  /* ---------------- modo pintar ---------------- */
  function paintCell(td, r, i){
    const gi = groups().findIndex(g => g.name === r.phase), color = PAL[gi % PAL.length], v = r.alloc[i];
    td.style.background = v ? (r.optional ? color + '99' : color) : ''; td.classList.toggle('empty', !v);
    td.querySelector('input').value = fmtV(v);
  }
  $('edGrid').addEventListener('pointerdown', e => {
    if (mode !== 'paint') return;
    const td = e.target.closest('td.cell[data-r]'); if (!td) return;
    e.preventDefault();
    const r = rowById(td.dataset.r), i = +td.dataset.i;
    pushHist(); drag = {rid: td.dataset.r, value: (r.alloc[i] || 0) === brush ? 0 : brush};
    r.alloc[i] = drag.value; paintCell(td, r, i); painting = true;
  });
  $('edGrid').addEventListener('pointerover', e => {
    if (!painting) return;
    const td = e.target.closest('td.cell[data-r]'); if (!td || td.dataset.r !== drag.rid) return;
    const r = rowById(td.dataset.r); r.alloc[+td.dataset.i] = drag.value; paintCell(td, r, +td.dataset.i);
  });
  window.addEventListener('pointerup', () => { if (painting){ painting = false; drag = null; changed(); render(); } });
  const setMode = m => { mode = m; render(); };
  $('edMType').onclick = () => setMode('type'); $('edMPaint').onclick = () => setMode('paint');
  document.querySelectorAll('#edBrush .chip[data-b]').forEach(c => c.onclick = () => { brush = +c.dataset.b; $('edBrushVal').value = brush; document.querySelectorAll('#edBrush .chip[data-b]').forEach(x => x.classList.toggle('on', x === c)); });
  $('edBrushVal').onchange = e => { brush = Math.max(0, E.num(e.target.value)); document.querySelectorAll('#edBrush .chip[data-b]').forEach(x => x.classList.toggle('on', +x.dataset.b === brush)); };

  /* ---------------- barra de ferramentas e parâmetros ---------------- */
  function uniquePhase(base){ const names = new Set(S.rows.map(r => r.phase)); let n = base, k = 2; while (names.has(n)) n = `${base} ${k++}`; return n; }
  $('edAddPhase').onclick = () => mutate(() => { S.rows.push(blank(uniquePhase('Nova fase'))); });
  $('edUndo').onclick = undo; $('edRedo').onclick = redo;
  $('edNPlus').onclick = () => mutate(() => { S.nPeriods++; });
  $('edNMinus').onclick = () => {
    if (S.nPeriods <= 1) return;
    const i = S.nPeriods - 1;
    if (S.rows.some(r => (r.alloc[i] || 0) > 0) && !confirm(`O período ${S.unit}${S.nPeriods} tem alocação. Remover mesmo assim?`)) return;
    mutate(() => { S.nPeriods--; });
  };
  window.addEventListener('keydown', e => {
    if (!(e.ctrlKey || e.metaKey) || !$('tab-tl').classList.contains('active')) return;
    if (document.activeElement && ['TEXTAREA', 'SELECT'].includes(document.activeElement.tagName)) return;
    if (document.activeElement && document.activeElement.tagName === 'INPUT' && !document.activeElement.dataset.cell) return;
    if (e.key === 'z' && !e.shiftKey){ e.preventDefault(); undo(); }
    else if (e.key === 'y' || (e.key === 'z' && e.shiftKey)){ e.preventDefault(); redo(); }
  });

  /* ---------------- combos: Perfil e Nível (vêm do cadastro de perfis) ---------------- */
  const pop = $('edPop'), popL = $('edPopL');
  let cur = null, active = 0, shown = [], activeL = 0, shownL = [];
  // mantém o popup inteiro dentro da janela (o formulário de cadastro aumenta a altura)
  const fitPop = el => { const r = el.getBoundingClientRect(); if (r.bottom > innerHeight - 8) el.style.top = Math.max(8, innerHeight - 8 - r.height) + 'px'; };
  const closePops = () => { pop.classList.remove('show'); popL.classList.remove('show'); };
  const place = (el, btn) => {
    const rc = btn.getBoundingClientRect();
    el.style.left = Math.max(8, Math.min(rc.left, innerWidth - 400)) + 'px';
    el.style.top = Math.max(8, Math.min(rc.bottom + 4, innerHeight - 430)) + 'px';
  };
  const natural = (a, b) => String(a).localeCompare(String(b), 'pt', {numeric: true});
  const priceOf = (r, p) => p.mode === 'rate' ? E.num(r.rate) : (E.num(r.cost) / (1 - p.gm));
  const priceChip = (entries, p) => {
    const ok = entries.filter(r => !missingRate(r, p));
    if (!ok.length) return '<span class="p no">sem taxa</span>';
    const v = ok.map(r => priceOf(r, p)), lo = Math.min(...v), hi = Math.max(...v);
    return `<span class="p">${esc(p.rcur)} ${lo === hi ? nfPT(lo, 0) : nfPT(lo, 0) + '–' + nfPT(hi, 0)}/h</span>`;
  };

  /* Perfil */
  function openCombo(btn, row){
    cur = row; closePops(); $('edPf').classList.remove('show'); $('edPq').value = ''; active = 0;
    $('edFLoc').style.display = ctx.hasLocation() ? '' : 'none';
    place(pop, btn); pop.classList.add('show'); listProfiles(); fitPop(pop); $('edPq').focus();
  }
  function listProfiles(){
    const p = ctx.P(), q = $('edPq').value.trim().toLowerCase();
    const usedRoles = new Set(S.rows.filter(r => r.role).map(r => r.role));
    const byRole = new Map(); catalog().forEach(e => { if (!byRole.has(e.role)) byRole.set(e.role, []); byRole.get(e.role).push(e); });
    shown = [...byRole.entries()].map(([role, entries]) => ({role, entries}))
      .filter(x => !q || `${x.role} ${x.entries.map(e => e.level + ' ' + e.location).join(' ')}`.toLowerCase().includes(q));
    const item = x => {
      const i = shown.indexOf(x), levels = [...new Set(x.entries.map(e => e.level || '—'))].sort(natural).join(', ');
      return `<li role="option" data-i="${i}" class="${i === active ? 'act' : ''}"><span class="r"><b>${esc(x.role)}</b><span>Níveis: ${esc(levels)}</span></span>${priceChip(x.entries, p)}</li>`;
    };
    const grp = (t, arr) => arr.length ? `<li class="grp" role="presentation">${t}</li>` + arr.map(item).join('') : '';
    $('edPl').innerHTML = grp('Já usados neste timeline', shown.filter(x => usedRoles.has(x.role))) + grp('Demais perfis cadastrados', shown.filter(x => !usedRoles.has(x.role)))
      || `<li class="grp" role="presentation">${esc(ctx.catalogHint() || 'Nenhum perfil encontrado — cadastre um novo abaixo')}</li>`;
  }
  function applyEntry(entry){
    pushHist(); cur.role = entry.role; cur.level = entry.level; cur.location = entry.location;
    closePops(); fit(); changed(); render();
  }
  function chooseRole(x){
    // mantém o nível/local atuais se o novo perfil também os tiver; senão usa o primeiro nível (em ordem natural)
    const sorted = x.entries.slice().sort((a, b) => natural(a.level, b.level));
    applyEntry(sorted.find(e => e.level === cur.level && e.location === cur.location) || sorted.find(e => e.level === cur.level) || sorted[0]);
  }
  $('edPq').oninput = () => { active = 0; listProfiles(); };
  $('edPq').onkeydown = e => {
    if (e.key === 'ArrowDown'){ active = Math.min(shown.length - 1, active + 1); listProfiles(); e.preventDefault(); }
    else if (e.key === 'ArrowUp'){ active = Math.max(0, active - 1); listProfiles(); e.preventDefault(); }
    else if (e.key === 'Enter' && shown[active]){ e.preventDefault(); chooseRole(shown[active]); }
    else if (e.key === 'Escape') closePops();
  };
  $('edPl').onclick = e => { const li = e.target.closest('li[data-i]'); if (li) chooseRole(shown[+li.dataset.i]); };
  $('edPnew').onclick = () => { $('edPf').classList.add('show'); fitPop(pop); $('edFRole').value = $('edPq').value; $('edFRole').focus(); };
  $('edPcancel').onclick = () => $('edPf').classList.remove('show');
  $('edPf').onsubmit = e => {
    e.preventDefault();
    const role = $('edFRole').value.trim();
    if (!role) return $('edFRole').focus();
    const entry = {role, level: $('edFLevel').value.trim(), location: $('edFLoc').value.trim(), cost: E.num($('edFCost').value) || '', rate: E.num($('edFRate').value) || ''};
    const res = ctx.addProfile(entry);
    if (res.existed) toast('Esse perfil/nível já está cadastrado — usando o existente.', 'info', 3500);
    ['edFRole', 'edFLevel', 'edFLoc', 'edFCost', 'edFRate'].forEach(id => $(id).value = '');
    applyEntry(res.entry);
  };

  /* Nível: troca entre os níveis do perfil, usa um nível de outro perfil ou adiciona um novo */
  function openLevel(btn, row){
    cur = row; closePops(); $('edLf').classList.remove('show'); $('edLq').value = ''; activeL = 0;
    place(popL, btn); popL.classList.add('show'); listLevels(); fitPop(popL); $('edLq').focus();
  }
  function listLevels(){
    const p = ctx.P(), q = $('edLq').value.trim().toLowerCase();
    const own = catalog().filter(e => e.role === cur.role);
    const ownLevels = new Set(own.map(e => e.level));
    const others = [...new Set(catalog().map(e => e.level).filter(l => l && !ownLevels.has(l)))].sort(natural);
    shownL = [
      ...own.sort((a, b) => natural(a.level, b.level)).map(e => ({kind: 'own', level: e.level, entry: e})),
      ...others.map(l => ({kind: 'other', level: l}))
    ].filter(x => !q || x.level.toLowerCase().includes(q) || (x.entry && x.entry.location.toLowerCase().includes(q)));
    const item = x => {
      const i = shownL.indexOf(x);
      const right = x.kind === 'own' ? priceChip([x.entry], p) : '<span class="p no">sem taxa</span>';
      const sub = x.kind === 'own' ? `${esc(x.entry.location || '-')}${x.entry.level === cur.level && x.entry.location === cur.location ? ' · atual' : ''}` : 'usar neste perfil (cria o nível sem taxa)';
      return `<li role="option" data-i="${i}" class="${i === activeL ? 'act' : ''}"><span class="r"><b>${esc(x.level || '—')}</b><span>${sub}</span></span>${right}</li>`;
    };
    const grp = (t, arr) => arr.length ? `<li class="grp" role="presentation">${t}</li>` + arr.map(item).join('') : '';
    $('edLl').innerHTML = grp(`Níveis de ${esc(cur.role)}`, shownL.filter(x => x.kind === 'own')) + grp('Níveis de outros perfis', shownL.filter(x => x.kind === 'other'))
      || '<li class="grp" role="presentation">Nenhum nível — adicione um novo abaixo</li>';
  }
  function chooseLevel(x){
    if (x.kind === 'own') return applyEntry(x.entry);
    applyEntry(ctx.addProfile({role: cur.role, level: x.level, location: cur.location, cost: '', rate: ''}).entry);
  }
  $('edLq').oninput = () => { activeL = 0; listLevels(); };
  $('edLq').onkeydown = e => {
    if (e.key === 'ArrowDown'){ activeL = Math.min(shownL.length - 1, activeL + 1); listLevels(); e.preventDefault(); }
    else if (e.key === 'ArrowUp'){ activeL = Math.max(0, activeL - 1); listLevels(); e.preventDefault(); }
    else if (e.key === 'Enter' && shownL[activeL]){ e.preventDefault(); chooseLevel(shownL[activeL]); }
    else if (e.key === 'Escape') closePops();
  };
  $('edLl').onclick = e => { const li = e.target.closest('li[data-i]'); if (li) chooseLevel(shownL[+li.dataset.i]); };
  $('edLnew').onclick = () => { $('edLf').classList.add('show'); fitPop(popL); $('edFLv').value = $('edLq').value; $('edFLv').focus(); };
  $('edLcancel').onclick = () => $('edLf').classList.remove('show');
  $('edLf').onsubmit = e => {
    e.preventDefault();
    const level = $('edFLv').value.trim();
    if (!level) return $('edFLv').focus();
    const entry = {role: cur.role, level, location: cur.location, cost: E.num($('edFLvCost').value) || '', rate: E.num($('edFLvRate').value) || ''};
    const res = ctx.addProfile(entry);
    if (res.existed) toast(`O nível ${level} já existe para ${cur.role} — usando o existente.`, 'info', 3500);
    ['edFLv', 'edFLvCost', 'edFLvRate'].forEach(id => $(id).value = '');
    applyEntry(res.entry);
  };
  document.addEventListener('mousedown', e => {
    if ((pop.classList.contains('show') || popL.classList.contains('show')) && !pop.contains(e.target) && !popL.contains(e.target) && !e.target.closest('[data-ed="combo"],[data-ed="lvcombo"]')) closePops();
  });

  /* ---------------- novo timeline vazio ---------------- */
  function startEmpty(){
    S.rows = [blank('Fase 1')]; S.deliverables = []; S.nPeriods = 12; S.unit = 'SP'; S.edits = {};
    $('p_unit').value = 'SP'; resetHistory(); fit();
  }

  return {render, resetHistory, startEmpty, TL_IDS, undo, redo};
}

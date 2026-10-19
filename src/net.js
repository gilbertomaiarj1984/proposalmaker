/* =====================================================================
   REDE — só para os extras (câmbio e feriados online). Nada aqui é necessário
   para o app funcionar. Toda chamada tem timeout e devolve erro com motivo legível;
   o app mostra aviso e oferece alternativa manual — nunca trava.
===================================================================== */
const Net = (() => {
'use strict';

const REASONS = {
  offline: 'sem conexão com a internet',
  timeout: 'tempo esgotado (a rede pode estar restrita)',
  blocked: 'acesso bloqueado (firewall/proxy da rede)',
  http: 'o serviço respondeu com erro',
  parse: 'resposta inesperada do serviço'
};
class NetError extends Error { constructor(code, detail){ super(REASONS[code] + (detail ? ` (${detail})` : '')); this.code = code; } }

const status = {online: typeof navigator !== 'undefined' ? navigator.onLine !== false : true, lastOkAt: null, lastError: null, listeners: []};
const notify = () => status.listeners.forEach(fn => { try { fn(status); } catch(e){} });
function onChange(fn){ status.listeners.push(fn); }
if (typeof window !== 'undefined'){
  window.addEventListener('online',  () => { status.online = true; notify(); });
  window.addEventListener('offline', () => { status.online = false; notify(); });
}

async function getJson(url, timeout = 8000){
  if (typeof navigator !== 'undefined' && navigator.onLine === false){ const e = new NetError('offline'); status.lastError = e.message; notify(); throw e; }
  const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), timeout);
  try {
    const r = await fetch(url, {signal: ctl.signal, headers: {'Accept': 'application/json'}});
    if (!r.ok) throw new NetError('http', 'HTTP ' + r.status);
    let j; try { j = await r.json(); } catch(e){ throw new NetError('parse'); }
    status.lastOkAt = new Date(); status.lastError = null; notify();
    return j;
  } catch(e){
    let err = e;
    if (!(e instanceof NetError)) err = e && e.name === 'AbortError' ? new NetError('timeout') : new NetError('blocked');
    status.lastError = err.message; notify();
    throw err;
  } finally { clearTimeout(t); }
}

/* ---------- câmbio: Frankfurter (BCE) -> open.er-api.com (cobre COP/ARS etc.) ---------- */
async function fetchFx(from, to){
  if (from === to) return {rate: 1, date: new Date().toISOString().slice(0,10), source: '—'};
  const errors = [];
  try {
    const j = await getJson(`https://api.frankfurter.dev/v1/latest?base=${encodeURIComponent(from)}&symbols=${encodeURIComponent(to)}`);
    if (j && j.rates && typeof j.rates[to] === 'number') return {rate: j.rates[to], date: j.date, source: 'Frankfurter (BCE)'};
  } catch(e){ errors.push(e); if (e.code === 'offline') throw e; }
  try {
    const j = await getJson(`https://open.er-api.com/v6/latest/${encodeURIComponent(from)}`);
    if (j && j.rates && typeof j.rates[to] === 'number') return {rate: j.rates[to], date: (j.time_last_update_utc ? new Date(j.time_last_update_utc).toISOString().slice(0,10) : ''), source: 'open.er-api.com'};
  } catch(e){ errors.push(e); }
  throw errors[errors.length - 1] || new NetError('parse');
}

/* ---------- feriados nacionais: BrasilAPI (confere o cálculo local) ---------- */
async function fetchHolidays(year){
  const j = await getJson(`https://brasilapi.com.br/api/feriados/v1/${year}`);
  if (!Array.isArray(j)) throw new NetError('parse');
  return j.filter(h => h && /^\d{4}-\d{2}-\d{2}$/.test(h.date)).map(h => ({date: h.date, name: String(h.name || 'Feriado')}));
}

/* ---------- teste de conexão (manual, pelo botão) ---------- */
async function probe(){
  const out = {};
  const test = async (key, fn) => { try { await fn(); out[key] = {ok: true}; } catch(e){ out[key] = {ok: false, msg: e.message}; } };
  await Promise.all([
    test('fx', () => getJson('https://api.frankfurter.dev/v1/latest?base=USD&symbols=BRL', 6000)),
    test('holidays', () => getJson('https://brasilapi.com.br/api/feriados/v1/' + new Date().getFullYear(), 6000))
  ]);
  return out;
}

return { NetError, status, onChange, getJson, fetchFx, fetchHolidays, probe };
})();

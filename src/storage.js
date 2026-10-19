/* =====================================================================
   ARMAZENAMENTO LOCAL — tudo opcional e à prova de falha:
   localStorage/sessionStorage, IndexedDB (rascunho) e File System Access API
   (pasta/arquivo de propostas .xlsx). Se algo não existir ou for bloqueado,
   as funções devolvem null/false e o app segue funcionando.
===================================================================== */
const Store = (() => {
'use strict';

const wrap = (area) => ({
  get(k){ try { return window[area].getItem(k); } catch(e){ return null; } },
  set(k, v){ try { window[area].setItem(k, v); return true; } catch(e){ return false; } },
  del(k){ try { window[area].removeItem(k); } catch(e){} }
});
const ls = wrap('localStorage'), ss = wrap('sessionStorage');

/* ---------- IndexedDB (chave/valor) ---------- */
let dbp = null;
function openDb(){
  if (dbp) return dbp;
  dbp = new Promise((res, rej) => {
    try {
      if (!window.indexedDB) return rej(new Error('IndexedDB indisponível'));
      const r = indexedDB.open('proposalmaker', 1);
      r.onupgradeneeded = () => { r.result.createObjectStore('kv'); };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error || new Error('Falha ao abrir IndexedDB'));
      r.onblocked = () => rej(new Error('IndexedDB bloqueado'));
    } catch(e){ rej(e); }
  });
  dbp.catch(() => { dbp = null; });
  return dbp;
}
async function tx(mode, fn){
  const db = await openDb();
  return new Promise((res, rej) => {
    try {
      const t = db.transaction('kv', mode), s = t.objectStore('kv');
      const out = fn(s);
      t.oncomplete = () => res(out && out.result !== undefined ? out.result : undefined);
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error || new Error('transação abortada'));
    } catch(e){ rej(e); }
  });
}
const kv = {
  async get(k){ try { return await tx('readonly', s => s.get(k)); } catch(e){ return undefined; } },
  async set(k, v){ try { await tx('readwrite', s => s.put(v, k)); return true; } catch(e){ return false; } },
  async del(k){ try { await tx('readwrite', s => s.delete(k)); return true; } catch(e){ return false; } },
  async keys(){ try { return await tx('readonly', s => s.getAllKeys()); } catch(e){ return []; } }
};

/* ---------- File System Access API ---------- */
const fs = {
  supported(){ return typeof window !== 'undefined' && !!(window.showDirectoryPicker && window.showSaveFilePicker && window.showOpenFilePicker); },
  async perm(handle, write){
    if (!handle) return false;
    const opts = {mode: write ? 'readwrite' : 'read'};
    try {
      if ((await handle.queryPermission(opts)) === 'granted') return true;
      return (await handle.requestPermission(opts)) === 'granted';
    } catch(e){ return false; }
  },
  // só consulta (não pede): seguro de chamar sem gesto do usuário, ex.: no autosave
  async granted(handle, write){
    try { return !!handle && (await handle.queryPermission({mode: write ? 'readwrite' : 'read'})) === 'granted'; } catch(e){ return false; }
  },
  async pickDir(){ return window.showDirectoryPicker({id:'proposalmaker', mode:'readwrite'}); },
  async pickOpen(){
    const [h] = await window.showOpenFilePicker({id:'proposalmaker-open', multiple:false, types:[{description:'Proposta (Excel)', accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]});
    return h;
  },
  async pickSave(name){
    return window.showSaveFilePicker({id:'proposalmaker-save', suggestedName:name, types:[{description:'Proposta (Excel)', accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]});
  },
  async list(dir){
    const out = [];
    for await (const [name, h] of dir.entries()){
      if (h.kind !== 'file' || !/\.xlsx$/i.test(name) || name.startsWith('~$')) continue;
      try { const f = await h.getFile(); out.push({name, handle:h, lastModified:f.lastModified, size:f.size}); } catch(e){}
    }
    return out.sort((a,b) => b.lastModified - a.lastModified);
  },
  async read(handle){ const f = await handle.getFile(); return f.arrayBuffer(); },
  async write(handle, data){
    const w = await handle.createWritable();
    try { await w.write(data); } finally { await w.close(); }
  },
  async create(dir, name, data){
    const h = await dir.getFileHandle(name, {create:true});
    await fs.write(h, data);
    return h;
  }
};

/* download de arquivo (fallback universal) */
function download(blob, name){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

return { ls, ss, kv, fs, download };
})();

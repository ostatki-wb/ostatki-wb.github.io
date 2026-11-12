// Хранилище в браузере пользователя. Магазины и настройки — localStorage (мало данных),
// загруженные карточки/остатки и снимки для отката — IndexedDB (может быть несколько мегабайт).
// Ничего не уходит никуда, кроме WB.

const LS_KEY = "ostatki-wb:v1";

function lsRead() {
  try { return JSON.parse(localStorage.getItem(LS_KEY) || "null") || {}; } catch (e) { return {}; }
}
function lsWrite(obj) {
  try { localStorage.setItem(LS_KEY, JSON.stringify(obj)); return true; } catch (e) { return false; }
}

export function getState() {
  const s = lsRead();
  return { shops: s.shops || [], current: s.current || null, ui: s.ui || {} };
}
export function saveState(st) { return lsWrite(st); }

export function newId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// ---------- IndexedDB ----------

let dbp = null;
function db() {
  if (!dbp) {
    dbp = new Promise((resolve, reject) => {
      const req = indexedDB.open("ostatki-wb", 1);
      req.onupgradeneeded = () => {
        const d = req.result;
        if (!d.objectStoreNames.contains("kv")) d.createObjectStore("kv");
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  return dbp;
}

async function tx(mode, fn) {
  const d = await db();
  return new Promise((resolve, reject) => {
    const t = d.transaction("kv", mode);
    const store = t.objectStore("kv");
    const r = fn(store);
    t.oncomplete = () => resolve(r && "result" in r ? r.result : undefined);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error);
  });
}

export const kvGet = key => tx("readonly", s => s.get(key));
export const kvSet = (key, val) => tx("readwrite", s => s.put(val, key));
export const kvDel = key => tx("readwrite", s => s.delete(key));
export const kvKeys = () => tx("readonly", s => s.getAllKeys());

export const cacheKey = shopId => `cache:${shopId}`;
export const snapPrefix = shopId => `snap:${shopId}:`;

export async function listSnapshots(shopId) {
  const keys = (await kvKeys()).filter(k => String(k).startsWith(snapPrefix(shopId))).sort().reverse();
  const out = [];
  for (const k of keys) out.push(Object.assign({ key: k }, await kvGet(k)));
  return out;
}

export async function deleteShopData(shopId) {
  for (const k of await kvKeys()) {
    if (k === cacheKey(shopId) || String(k).startsWith(snapPrefix(shopId))) await kvDel(k);
  }
}

// Клиент WB API (перенос app/wb.py). Работает прямо из браузера: WB разрешает такие запросы (CORS).
// Токен уходит только в WB.

export const CONTENT = "https://content-api.wildberries.ru";
export const ANALYTICS = "https://seller-analytics-api.wildberries.ru";
export const COMMON = "https://common-api.wildberries.ru";

export const sleep = ms => new Promise(r => setTimeout(r, ms));

export class WBError extends Error {
  constructor(msg, status = null, body = "") {
    super(msg);
    this.status = status;
    this.body = body || "";
  }
}

function human(code, url, text) {
  const host = url.split("/")[2];
  if (code === 401) return `WB не принял токен (401, ${host}). Проверьте, что токен скопирован полностью и не истёк.`;
  if (code === 403) return `У токена нет доступа к этому разделу (403, ${host}). Проверьте категории доступа токена.`;
  if (code === 404) return `Метод не найден (404, ${host}). Возможно, WB изменил API.`;
  if (code === 429) return `WB просит подождать: слишком частые запросы (429, ${host}). Попробуйте через минуту.`;
  return `Ошибка WB ${code} (${host}): ${text}`;
}

export async function request(method, url, token, body, { retries = 3, progress = null, fetchImpl = globalThis.fetch } = {}) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    let resp;
    try {
      resp = await fetchImpl(url, {
        method,
        headers: Object.assign({ Authorization: token }, body !== undefined ? { "Content-Type": "application/json" } : {}),
        body: body !== undefined ? JSON.stringify(body) : undefined,
      });
    } catch (e) {
      if (attempt < retries) { await sleep(3000); continue; }
      throw new WBError(`Нет связи с ${url.split("/")[2]}: ${e.message || e}. Проверьте интернет.`);
    }
    const text = await resp.text();
    if (resp.ok) return text ? JSON.parse(text) : null;
    if (resp.status === 429 && attempt < retries) {
      // браузер не даёт прочитать X-Ratelimit-Retry, поэтому ждём с запасом
      let wait = parseInt(resp.headers.get("X-Ratelimit-Retry") || resp.headers.get("Retry-After") || "21", 10);
      wait = Math.min(Math.max(isNaN(wait) ? 21 : wait, 2), 70);
      if (progress) progress(`WB просит подождать ${wait} с…`);
      await sleep(wait * 1000);
      continue;
    }
    if (resp.status >= 500 && attempt < retries) { await sleep(3000); continue; }
    throw new WBError(human(resp.status, url, text.slice(0, 300)), resp.status, text.slice(0, 300));
  }
}

// ---------- карточки и ярлыки (раздел «Товары и цены») ----------

/** Все карточки продавца, постранично по 100. subjectIds — только эти предметы. */
export async function fetchCards(token, { progress = () => {}, subjectIds = null, pause = 700 } = {}) {
  const cards = [], seen = new Set();
  let cursor = { limit: 100 };
  const filter = { withPhoto: -1 };
  if (subjectIds && subjectIds.length) filter.objectIDs = subjectIds.map(Number);
  for (;;) {
    const res = (await request("POST", CONTENT + "/content/v2/get/cards/list", token,
      { settings: { cursor, filter } }, { progress })) || {};
    const batch = res.cards || [];
    for (const c of batch) {
      if (seen.has(c.nmID)) continue;     // курсор WB иногда повторяет карточки на стыке страниц
      seen.add(c.nmID);
      cards.push({
        nmID: c.nmID, imtID: c.imtID, vendorCode: c.vendorCode || "",
        subjectID: c.subjectID, subjectName: c.subjectName || "Без категории",
        title: c.title || "", tags: (c.tags || []).map(t => t.name).filter(Boolean),
      });
    }
    progress(`Загружаю карточки «Товары и цены»: ${cards.length}`);
    const cur = res.cursor || {};
    if (batch.length < 100 || !cur.updatedAt || (cur.total || 0) < 100) break;
    cursor = { limit: 100, updatedAt: cur.updatedAt, nmID: cur.nmID };
    await sleep(pause);                    // лимит Content API: 100 запросов в минуту
  }
  return cards;
}

export async function fetchTags(token) {
  const res = (await request("GET", CONTENT + "/content/v2/tags", token)) || {};
  return (res.data || []).filter(t => t.name).map(t => ({ id: t.id, name: t.name, color: t.color || "" }));
}

/** nmID товаров, которые сейчас лежат в карточке imtId. */
export async function imtMembers(token, imtId) {
  const res = (await request("POST", CONTENT + "/content/v2/get/cards/list", token,
    { settings: { cursor: { limit: 100 }, filter: { withPhoto: -1, imtID: Number(imtId) } } })) || {};
  return (res.cards || []).filter(c => c.imtID === Number(imtId)).map(c => c.nmID);
}

/** Объединение/разъединение. С targetImt — перенос в карточку; без — отсоединение
 *  (несколько товаров одним запросом окажутся ВМЕСТЕ). На несуществующий nmID WB отвечает успехом. */
export async function moveNm(token, nmIds, targetImt = null) {
  const body = targetImt != null ? { targetIMT: Number(targetImt), nmIDs: nmIds.map(Number) } : { nmIDs: nmIds.map(Number) };
  return request("POST", CONTENT + "/content/v2/cards/moveNm", token, body);
}

export async function sellerName(token) {
  try {
    const res = (await request("GET", COMMON + "/api/v1/seller-info", token, undefined, { retries: 1 })) || {};
    return res.tradeMark || res.name || "";
  } catch (e) { return ""; }
}

// ---------- остатки ----------
// Колонка «Остаток» в «Товары и цены» = отчёт «Остатки» в Аналитике (stocks-report), stockType="" —
// склады WB + свои склады. Проверено: совпадает с кабинетом.

export const STOCK_TYPES = ["", "wb", "mp"];

function today() {
  const d = new Date();
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

/** Текущие остатки: { nmID (строкой): шт }. Товаров без остатков WB в отчёте не возвращает. */
export async function fetchStocks(token, stockType = "", { progress = () => {}, pause = 21000 } = {}) {
  if (!STOCK_TYPES.includes(stockType)) stockType = "";
  const day = today(), out = {};
  for (let offset = 0; ; offset += 1000) {
    const body = {
      currentPeriod: { start: day, end: day }, stockType, skipDeletedNm: true,
      availabilityFilters: ["deficient", "actual", "balanced", "nonActual", "nonLiquid", "invalidData"],
      orderBy: { field: "stockCount", mode: "desc" }, limit: 1000, offset,
    };
    const res = (await request("POST", ANALYTICS + "/api/v2/stocks-report/products/products", token, body, { progress })) || {};
    const items = (res.data || {}).items;
    if (!Array.isArray(items)) throw new WBError("WB вернул остатки в неожиданном формате (нет data.items).");
    for (const it of items) {
      if (it.nmID == null) continue;
      const q = parseInt(((it.metrics || {}).stockCount) || 0, 10);
      out[String(it.nmID)] = (out[String(it.nmID)] || 0) + q;
    }
    progress(`Загружаю остатки: ${Object.keys(out).length} товаров`);
    if (items.length < 1000) break;
    progress(`Загружаю остатки: ${Object.keys(out).length} товаров, жду ${pause / 1000} с (лимит WB)…`);
    await sleep(pause);
  }
  return out;
}

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
export async function fetchCards(token, {
  progress = () => {}, subjectIds = null, pause = 750, fetchImpl = globalThis.fetch,
} = {}) {
  const cards = [], seen = new Set(), seenCursors = new Set();
  const limit = 100;
  let cursor = { limit };
  const filter = { withPhoto: -1 };
  if (subjectIds && subjectIds.length) filter.objectIDs = subjectIds.map(Number);
  for (;;) {
    const res = (await request("POST", CONTENT + "/content/v2/get/cards/list", token,
      { settings: { cursor, filter } }, { progress, fetchImpl })) || {};
    const batch = res.cards;
    if (!Array.isArray(batch)) throw new WBError("WB вернул каталог в неожиданном формате (нет списка карточек).");
    for (const c of batch) {
      if (seen.has(c.nmID)) continue;     // курсор WB иногда повторяет карточки на стыке страниц
      seen.add(c.nmID);
      const photo = c.photos?.[0];
      cards.push({
        nmID: c.nmID, imtID: c.imtID, vendorCode: c.vendorCode || "",
        subjectID: c.subjectID, subjectName: c.subjectName || "Без категории",
        title: c.title || "", tags: (c.tags || []).map(t => t.name).filter(Boolean),
        photoUrl: photo?.c516x688 || photo?.big || photo?.c246x328 || "",
      });
    }
    progress(`Загружаю карточки «Товары и цены»: ${cards.length}`);
    const next = res.cursor;
    const total = next?.total ?? batch.length;
    if (!Number.isInteger(total) || total < 0) throw new WBError("WB вернул некорректный размер страницы каталога.");
    if (total < limit) break;
    if (!next?.updatedAt || !Number.isSafeInteger(next.nmID) || next.nmID <= 0) {
      throw new WBError("WB не указал, с какой карточки продолжить загрузку каталога. Повторите обновление.");
    }
    const cursorKey = JSON.stringify([next.updatedAt, next.nmID]);
    if (seenCursors.has(cursorKey)) throw new WBError("WB повторил страницу каталога. Повторите обновление.");
    seenCursors.add(cursorKey);
    cursor = { limit, updatedAt: next.updatedAt, nmID: next.nmID };
    await sleep(pause);
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

// ---------- заказы ----------

export const MAX_ORDER_DAYS = 365;
const DAY_MS = 24 * 60 * 60 * 1000;
const moscowDate = new Intl.DateTimeFormat("sv-SE", {
  timeZone: "Europe/Moscow", year: "numeric", month: "2-digit", day: "2-digit",
});

function dateValue(date) {
  const value = /^\d{4}-\d{2}-\d{2}$/.test(date) ? Date.parse(date) : NaN;
  return Number.isFinite(value) && new Date(value).toISOString().slice(0, 10) === date ? value : NaN;
}

/** Границы и число календарных дней, включая обе даты. */
export function orderPeriod(selection = { days: 30 }, now = new Date()) {
  const today = moscowDate.format(now);
  const todayValue = dateValue(today);
  const earliest = todayValue - (MAX_ORDER_DAYS - 1) * DAY_MS;
  if (!("start" in selection || "end" in selection)) {
    const { days } = selection;
    if (!Number.isInteger(days) || days < 1 || days > MAX_ORDER_DAYS) throw new RangeError("Выберите период в пределах последних 365 дней.");
    const start = new Date(todayValue - (days - 1) * DAY_MS).toISOString().slice(0, 10);
    return { start, end: today, days };
  }
  const { start, end } = selection;
  if (!start || !end) throw new RangeError("Укажите начало и конец периода.");
  const first = dateValue(start), last = dateValue(end);
  if (!Number.isFinite(first) || !Number.isFinite(last)) throw new RangeError("Укажите существующие даты.");
  if (first > last) throw new RangeError("Дата начала должна быть раньше даты окончания или совпадать с ней.");
  if (last > todayValue) throw new RangeError("Конец периода не может быть позже сегодняшней даты по Москве.");
  if (first < earliest) throw new RangeError("Выберите даты в пределах последних 365 дней.");
  return { start, end, days: (last - first) / DAY_MS + 1 };
}

/** Заказанные единицы по артикулу WB за выбранный период, по всем моделям продаж. */
export async function fetchOrders(token, {
  period, progress = () => {}, pause = 21000, now = new Date(), fetchImpl = globalThis.fetch,
} = {}) {
  const { start, end, days } = orderPeriod(period, now);
  const limit = 1000;
  const counts = {};
  for (let offset = 0; ; offset += limit) {
    const body = {
      selectedPeriod: { start, end }, skipDeletedNm: true,
      orderBy: { field: "orderCount", mode: "desc" }, limit, offset,
    };
    const res = await request("POST", ANALYTICS + "/api/analytics/v3/sales-funnel/products", token, body, { progress, fetchImpl });
    const products = res?.data?.products;
    if (!Array.isArray(products)) throw new WBError("WB вернул заказы в неожиданном формате (нет data.products).");
    for (const item of products) {
      const nmId = item.product?.nmId;
      const count = item.statistic?.selected?.orderCount;
      if (nmId != null && Number.isSafeInteger(count) && count >= 0) counts[nmId] = count;
    }
    if (products.length < limit) break;
    progress(`Загружаю заказы: ${Object.keys(counts).length} товаров, жду ${pause / 1000} с (лимит WB)…`);
    await sleep(pause);
  }
  return { start, end, days, counts };
}

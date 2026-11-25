// Сборка отчёта по категориям и ярлыкам (перенос app/report.py из приложения «Остатки WB»).

export const NO_TAG = "без ярлыка";
export const PREFERRED_ORDER = ["животные", "люди", "город", "кошка", "море",
  "натюрморт", "природа", "птицы", "собака", "цветы"];

// ключ -> stockType метода stocks-report
export const SOURCES = {
  all: "Все склады: WB + свои (как «Остаток» в «Товары и цены»)",
  wb: "Только склады WB",
  mp: "Только свои склады (FBS)",
};
export const STOCK_TYPE = { all: "", wb: "wb", mp: "mp" };

// категории (предметы), у которых в отчёт добавляется колонка «Название» — для всех ярлыков
export const NAMED_SUBJECTS = new Set(["алмазная мозаика", "картины по номерам", "наборы для вышивания"]);

// В JavaScript \b не понимает кириллицу — границы слова задаём явно.
const B = "(?<![\\p{L}\\p{N}_])";   // начало слова
const E = "(?![\\p{L}\\p{N}_])";    // конец слова
// служебные слова из наименования карточки, которые не относятся к названию картины
const NOISE = new RegExp(
  "алмазн(ая|ой)\\s+мозаик[аи]|для\\s+детей|ч[её]рн(ый|ом)\\s+холст[е]?|картин[аы]\\s+по\\s+номерам" +
  `|вышивка\\s+крестиком|${B}набор${E}|${B}детская${E}|(на|без)\\s+подрамник[аеи]?|${B}икон[аы]${E}` +
  `|\\d+\\s*[хxХX×]\\s*\\d+(\\s*[сc][мm]\\.?)?|${B}[сc][мm]${E}\\.?`, "giu");

function strip(s, chars) {
  let a = 0, b = s.length;
  while (a < b && chars.includes(s[a])) a++;
  while (b > a && chars.includes(s[b - 1])) b--;
  return s.slice(a, b);
}

/** Название картины из наименования карточки:
 *  «Алмазная мозаика 35х50 cм Икона Николай Чудотворец KC007» → «Николай Чудотворец». */
export function shortName(title, article) {
  let s = String(title || "");
  if (article) s = s.split(article).join(" ");
  s = s.replace(NOISE, " ");
  s = s.replace(/\s*,(\s*,)+/gu, ",");
  s = strip(s.replace(/\s+/gu, " "), " ,.-–—");
  s = s.replace(/^на\s+/iu, "");      // «Алмазная мозаика на 27х33 см Вера Надежда Любовь»
  return s.slice(0, 1).toUpperCase() + s.slice(1);
}

export function isNamed(subject) {
  return NAMED_SUBJECTS.has(String(subject || "").trim().toLowerCase());
}

// сравнение строк по кодам символов — как sorted() в Python
export function cmp(a, b) { return a < b ? -1 : a > b ? 1 : 0; }

export function orderTags(names) {
  const key = n => {
    const low = n.toLowerCase();
    if (low === NO_TAG) return [2, 0, low];
    const i = PREFERRED_ORDER.indexOf(low);
    if (i >= 0) return [0, i, low];
    return [1, 0, low];
  };
  return [...names].sort((x, y) => {
    const a = key(x), b = key(y);
    return (a[0] - b[0]) || (a[1] - b[1]) || cmp(a[2], b[2]);
  });
}

export function stockMap(cache, source) {
  return ((cache && cache.stocks) || {})[source] || {};
}

export function categories(cache) {
  const counts = new Map();
  for (const c of (cache && cache.cards) || []) counts.set(c.subjectName, (counts.get(c.subjectName) || 0) + 1);
  return [...counts].map(([name, count]) => ({ name, count }))
    .sort((a, b) => (b.count - a.count) || cmp(a.name, b.name));
}

export function tagsForCategory(cache, subject) {
  const colors = {};
  for (const t of (cache && cache.tags) || []) colors[t.name] = t.color || "";
  const counts = new Map();
  for (const c of (cache && cache.cards) || []) {
    if (c.subjectName !== subject) continue;
    for (const n of cardTags(c)) counts.set(n, (counts.get(n) || 0) + 1);
  }
  return orderTags([...counts.keys()]).map(n => ({ name: n, count: counts.get(n), color: colors[n] || "" }));
}

function cardTags(card) {
  return card.tags?.length ? card.tags : [NO_TAG];
}

function stockRow(card, stocks, orders) {
  const article = String(card.vendorCode ?? "");
  const count = orders?.counts?.[card.nmID];
  return {
    article: article.trim() ? article : String(card.nmID),
    nmID: card.nmID,
    qty: parseInt(stocks[String(card.nmID)] || 0, 10),
    ordersPerDay: count == null ? undefined : count / orders.days,
    photoUrl: card.photoUrl,
  };
}

function compareRows(sortBy) {
  return (a, b) => {
    if (sortBy === "orders") {
      const difference = (b.ordersPerDay ?? -1) - (a.ordersPerDay ?? -1);
      if (difference) return difference;
    }
    return (b.qty - a.qty) || cmp(a.article, b.article);
  };
}

export function build(cache, subject, tags, source, includeZero, sortBy = "stock") {
  const stocks = stockMap(cache, source);
  const orders = cache?.orders;
  const wanted = orderTags(tags);
  const groups = new Map(wanted.map(t => [t, []]));
  const named = isNamed(subject);
  for (const c of (cache && cache.cards) || []) {
    if (c.subjectName !== subject) continue;
    const row = stockRow(c, stocks, orders);
    if (row.qty <= 0 && !includeZero) continue;
    if (named) row.name = shortName(c.title, c.vendorCode);
    for (const t of cardTags(c)) groups.get(t)?.push(row);
  }
  const compare = compareRows(sortBy);
  return wanted.map(t => {
    const rows = groups.get(t).sort(compare);
    return { tag: t, rows, sku: rows.length, total: rows.reduce((s, r) => s + r.qty, 0), named };
  });
}

export function summarize(groups) {
  const quantities = new Map();
  for (const group of groups) {
    for (const row of group.rows) quantities.set(row.nmID, row.qty);
  }
  let total = 0;
  for (const qty of quantities.values()) total += qty;
  return { sku: quantities.size, total };
}

export function buildCardGroups(cache, subject, tags, source, includeZero, sortBy = "stock") {
  const wanted = new Set(tags);
  const stocks = stockMap(cache, source);
  const members = new Map();
  const seen = new Set();
  for (const card of cache?.cards || []) {
    const imtID = Number(card.imtID), nmID = Number(card.nmID);
    if (card.subjectName !== subject || !Number.isSafeInteger(imtID) || imtID <= 0
        || !Number.isSafeInteger(nmID) || nmID <= 0 || seen.has(nmID)) continue;
    seen.add(nmID);
    if (!members.has(imtID)) members.set(imtID, []);
    members.get(imtID).push(card);
  }

  const groups = [];
  const compare = compareRows(sortBy);
  for (const [, cards] of [...members].sort(([a], [b]) => a - b)) {
    // Фильтры отбирают карточки целиком. В состав каждой выбранной карточки входят и товары без остатка.
    if (cards.length < 2 || !cards.some(c => cardTags(c).some(t => wanted.has(t)))) continue;
    const rows = cards.map(c => stockRow(c, stocks, cache.orders));
    if (!includeZero && !rows.some(row => row.qty > 0)) continue;
    groups.push({ name: `Карточка ${groups.length + 1}`, rows: rows.sort(compare) });
  }
  return groups;
}

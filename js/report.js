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
    for (const n of (c.tags && c.tags.length ? c.tags : [NO_TAG])) counts.set(n, (counts.get(n) || 0) + 1);
  }
  return orderTags([...counts.keys()]).map(n => ({ name: n, count: counts.get(n), color: colors[n] || "" }));
}

export function build(cache, subject, tags, source, includeZero, sortBy = "stock") {
  const stocks = stockMap(cache, source);
  const orders = cache?.orders;
  const wanted = orderTags(tags);
  const groups = new Map(wanted.map(t => [t, []]));
  const named = isNamed(subject);
  for (const c of (cache && cache.cards) || []) {
    if (c.subjectName !== subject) continue;
    const qty = parseInt(stocks[String(c.nmID)] || 0, 10);
    if (qty <= 0 && !includeZero) continue;
    const count = orders?.counts?.[c.nmID];
    const ordersPerDay = count == null ? undefined : count / orders.days;
    for (const t of (c.tags && c.tags.length ? c.tags : [NO_TAG])) {
      if (!groups.has(t)) continue;
      const row = { article: c.vendorCode, nmID: c.nmID, qty, ordersPerDay };
      if (named) row.name = shortName(c.title, c.vendorCode);
      groups.get(t).push(row);
    }
  }
  return wanted.map(t => {
    const rows = groups.get(t).sort((a, b) => {
      if (sortBy === "orders") {
        const difference = (b.ordersPerDay ?? -1) - (a.ordersPerDay ?? -1);
        if (difference) return difference;
      }
      return (b.qty - a.qty) || cmp(a.article, b.article);
    });
    return { tag: t, rows, sku: rows.length, total: rows.reduce((s, r) => s + r.qty, 0), named };
  });
}

import assert from "node:assert/strict";
import test from "node:test";
import ExcelJS from "exceljs";
import { fetchCards, fetchOrders, orderPeriod } from "../js/wb.js";
import { build, buildCardGroups, summarize } from "../js/report.js";
import { parseXlsx, toXlsx } from "../js/excel.js";

test("WB: 30 дней по Москве включительно, пагинация воронки, ноль и пропуск заказов", async () => {
  const requests = [];
  const product = (nmId, orderCount) => ({ product: { nmId }, statistic: { selected: { orderCount } } });
  const firstPage = [
    product(1, 221), product(2, 0), product(3),
    ...Array.from({ length: 997 }, (_, i) => product(i + 4, 3)),
  ];
  const res = await fetchOrders("test-token", {
    now: new Date("2026-11-19T21:00:00Z"),
    pause: 0,
    fetchImpl: async (url, req) => {
      assert.equal(url, "https://seller-analytics-api.wildberries.ru/api/analytics/v3/sales-funnel/products");
      const body = JSON.parse(req.body);
      requests.push(body);
      const products = body.offset === 0 ? firstPage : [product(1001, 30)];
      return new Response(JSON.stringify({ data: { products } }));
    },
  });
  assert.deepEqual(requests.map(r => r.offset), [0, 1000]);
  for (const { selectedPeriod } of requests) {
    assert.deepEqual(selectedPeriod, { start: "2026-10-22", end: "2026-11-20" });
  }
  assert.equal(res.days, 30);
  assert.equal(res.counts[1], 221);
  assert.equal(res.counts[2], 0);
  assert.equal(res.counts[3], undefined);
  assert.equal(res.counts[1001], 30);
});

test("WB: пресеты, год и свои даты с учётом обеих границ", async () => {
  const now = new Date("2026-11-19T21:00:00Z");
  assert.deepEqual(orderPeriod({ days: 7 }, now), { start: "2026-11-14", end: "2026-11-20", days: 7 });
  assert.deepEqual(orderPeriod({ days: 90 }, now), { start: "2026-08-23", end: "2026-11-20", days: 90 });
  const year = orderPeriod({ days: 365 }, now);
  assert.deepEqual(year, { start: "2025-11-21", end: "2026-11-20", days: 365 });
  assert.deepEqual(orderPeriod({ start: year.start, end: year.start }, now), {
    start: "2025-11-21", end: "2025-11-21", days: 1,
  });
  assert.throws(() => orderPeriod({ start: "2025-11-20", end: "2026-11-20" }, now), /365 дней/);
  assert.throws(() => orderPeriod({ start: "2026-11-20", end: "2026-11-19" }, now), /Дата начала/);
  assert.throws(() => orderPeriod({ start: "2026-11-20", end: "2026-11-21" }, now), /сегодняшней/);
  assert.throws(() => orderPeriod({ start: "2026-02-30", end: "2026-03-01" }, now), /существующие даты/);
  const period = orderPeriod({ start: "2026-10-29", end: "2026-11-02" }, now);
  const data = await fetchOrders("test-token", {
    period, now,
    fetchImpl: async (_url, req) => {
      assert.deepEqual(JSON.parse(req.body).selectedPeriod, { start: "2026-10-29", end: "2026-11-02" });
      return new Response(JSON.stringify({ data: { products: [{ product: { nmId: 1 }, statistic: { selected: { orderCount: 10 } } }] } }));
    },
  });
  assert.deepEqual(data, { start: "2026-10-29", end: "2026-11-02", days: 5, counts: { 1: 10 } });
});

test("заказы по ярлыкам: сортировка, нулевой остаток, отсутствие данных и старый кэш", () => {
  const subject = "Алмазная мозаика";
  const tags = ["Лебеди", "Река"];
  const cards = ["A", "B", "C"].map((vendorCode, idx) => ({
    vendorCode, nmID: idx + 1, tags, subjectName: subject,
  }));
  const stocks = { all: { 1: 0, 2: 5, 3: 19 } };
  const orders = { days: 30, counts: { 1: 221, 2: 0 } };
  const groups = build({ cards, stocks, orders }, subject, tags, "all", true, "orders");
  assert.deepEqual(summarize(groups), { sku: 3, total: 24 });
  for (const { rows } of groups) {
    assert.deepEqual(rows.map(r => [r.article, r.qty, r.ordersPerDay]), [
      ["A", 0, 221 / 30], ["B", 5, 0], ["C", 19, undefined],
    ]);
  }
  const [hiddenZeros] = build({ cards, stocks, orders }, subject, tags, "all", false, "orders");
  assert.deepEqual(hiddenZeros.rows.map(r => r.article), ["B", "C"]);
  const [legacy] = build({ cards, stocks }, subject, tags, "all", false);
  assert.deepEqual(legacy.rows.map(r => [r.article, r.ordersPerDay]), [["C", undefined], ["B", undefined]]);
});

test("Excel: скорость числом, период и формула SKU, повторный импорт игнорирует метрики", async () => {
  const groups = [{
    tag: "Лебеди", named: true,
    rows: [
      { article: "A", name: "Лебеди", qty: 19, ordersPerDay: 221 / 30 },
      { article: "B", name: "Река", qty: 5, ordersPerDay: 1 },
      { article: "C", name: "Озеро", qty: 0, ordersPerDay: 0 },
      { article: "D", name: "Гора", qty: 1 },
    ],
  }];
  const dates = "22.10.2026 – 20.11.2026 (МСК)";
  const xlsx = await toXlsx(groups, ExcelJS, dates);
  const book = await new ExcelJS.Workbook().xlsx.load(xlsx);
  const sheet = book.getWorksheet("Остатки");
  assert.equal(sheet.getCell("D2").value, "Заказов в день");
  assert.match(sheet.getCell("D2").note, /22\.10\.2026.*20\.11\.2026/);
  assert.equal(sheet.getCell("D3").value, 221 / 30);
  assert.equal(sheet.getCell("D3").numFmt, "0.00");
  assert.equal(sheet.getCell("D5").value, 0);
  assert.equal(sheet.getCell("D6").value, null);
  assert.equal(sheet.getCell("C7").formula, "COUNTA(A3:A6)");
  const parsed = await parseXlsx(xlsx, new Set(["A", "B", "C", "D", "19", "1"]), ExcelJS);
  assert.deepEqual(parsed.map(g => g.items), [["A", "B", "C", "D"]]);
});

test("Excel: столбец фото доступен отдельно от заказов в обеих выгрузках", async () => {
  const groups = [{ tag: "Река", name: "Карточка 1", rows: [{ article: "А1", nmID: 1, qty: 8, ordersPerDay: 2 }] }];
  const dates = "14.11.2026 – 20.11.2026 (МСК)";
  for (const byCards of [false, true]) {
    const xlsx = await toXlsx(groups, ExcelJS, dates, { byCards, photos: new Map(), withOrders: false });
    const book = await new ExcelJS.Workbook().xlsx.load(xlsx);
    const sheet = book.worksheets[0];
    assert.deepEqual(sheet.getRow(2).values.slice(1), ["Артикул", "Кол-во", "Фото"]);
    assert.equal(sheet.getCell("B4").result, 1);
    assert.deepEqual(await parseXlsx(xlsx, new Set(["А1", "8", "Фото"]), ExcelJS), [
      { name: byCards ? "Карточка 1" : "Река", items: ["А1"] },
    ]);
  }
});

test("Карточки: все варианты сохраняются, даже если только один есть в наличии", async () => {
  const subject = "Алмазная мозаика";
  const item = (nmID, imtID, vendorCode, tags) => ({ nmID, imtID, vendorCode, tags, subjectName: subject });
  const first = item(1, 50, "KWA0282", ["Река", "Лебеди"]);
  const cards = [first, item(2, 50, "SWA0139", ["Горы"]), first,
    item(3, 90, "", ["Река"]), item(4, 90, "Нет остатков", []),
    item(5, 99, "Одиночный", ["Река"]), item(6, null, "D", ["Река"]), item(7, null, "E", ["Река"])];
  const cache = { cards, stocks: { wb: { 1: 1, 5: 90, 6: 6, 7: 8 } }, orders: { days: 7, counts: { 1: 0, 2: 14 } } };
  const groups = buildCardGroups(cache, subject, ["Река", "Лебеди"], "wb", false);
  assert.equal(groups.length, 1);
  assert.equal(groups[0].name, "Карточка 1");
  assert.deepEqual(groups[0].rows.map(r => [r.article, r.qty, r.ordersPerDay]), [["KWA0282", 1, 0], ["SWA0139", 0, 2]]);
  const all = buildCardGroups(cache, subject, ["Река"], "wb", true, "orders");
  assert.deepEqual(all.map(g => g.rows.map(r => r.article)), [["SWA0139", "KWA0282"], ["3", "Нет остатков"]]);

  const xlsx = await toXlsx(groups, ExcelJS, "14.11.2026 – 20.11.2026 (МСК)", { byCards: true });
  const book = await new ExcelJS.Workbook().xlsx.load(xlsx);
  const sheet = book.getWorksheet("По карточкам");
  assert.equal(sheet.getCell("A1").value, "Карточка 1");
  assert.equal(sheet.getCell("B5").formula, "COUNTA(A3:A4)");
  assert.equal(sheet.getCell("B5").result, 2);
  assert.equal(sheet.getCell("B4").value, 0);
  assert.equal(sheet.getCell("C4").value, 2);
  assert.deepEqual(await parseXlsx(xlsx, new Set(["KWA0282", "SWA0139", "0", "1", "2"]), ExcelJS), [
    { name: "Карточка 1", items: ["KWA0282", "SWA0139"] },
  ]);
});

test("WB: каталог продолжается по курсору, повторная страница и неполный ответ выявляются", async () => {
  const requests = [];
  const first = { nmID: 1, imtID: 5, vendorCode: "KWA0282" };
  const second = { nmID: 2, imtID: 5, vendorCode: "SWA0139" };
  const cursor = { total: 100, nmID: 1, updatedAt: "2026-11-25T16:00:00Z" };
  const cards = await fetchCards("test-token", {
    subjectIds: [44],
    pause: 0,
    fetchImpl: async (_url, req) => {
      requests.push(JSON.parse(req.body).settings);
      return new Response(JSON.stringify(requests.length === 1
        ? { cards: [first], cursor }
        : { cards: [first, second], cursor: { total: 2 } }));
    },
  });
  assert.deepEqual(cards.map(card => card.nmID), [1, 2]);
  assert.deepEqual(requests[1].cursor, { limit: 100, nmID: 1, updatedAt: cursor.updatedAt });
  assert.deepEqual(requests[1].filter.objectIDs, [44]);
  await assert.rejects(fetchCards("test-token", {
    pause: 0,
    fetchImpl: async () => new Response(JSON.stringify({ cards: [first], cursor })),
  }), /повторил страницу/);
  await assert.rejects(fetchCards("test-token", {
    fetchImpl: async () => new Response("{}"),
  }), /нет списка карточек/);
});

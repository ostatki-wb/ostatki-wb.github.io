import assert from "node:assert/strict";
import test from "node:test";
import ExcelJS from "exceljs";
import { fetchOrders } from "../js/wb.js";
import { build } from "../js/report.js";
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

test("заказы по ярлыкам: сортировка, нулевой остаток, отсутствие данных и старый кэш", () => {
  const subject = "Алмазная мозаика";
  const tags = ["Лебеди", "Река"];
  const cards = ["A", "B", "C"].map((vendorCode, idx) => ({
    vendorCode, nmID: idx + 1, tags, subjectName: subject,
  }));
  const stocks = { all: { 1: 0, 2: 5, 3: 19 } };
  const orders = { days: 30, counts: { 1: 221, 2: 0 } };
  const groups = build({ cards, stocks, orders }, subject, tags, "all", true, "orders");
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

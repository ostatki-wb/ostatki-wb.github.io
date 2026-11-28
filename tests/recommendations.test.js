import test from "node:test";
import assert from "node:assert/strict";
import { index } from "../js/cards.js";
import { checkAdditions, saveRecommendations } from "../js/recommendations.js";
import { CONTENT, ANALYTICS, fetchRecommendations, fetchStocks } from "../js/wb.js";

function mockWB(...answers) {
  const calls = [];
  const fetchImpl = async (url, request) => {
    calls.push([url, JSON.parse(request.body)]);
    assert.ok(answers.length, "Неожиданный запрос к WB");
    return new Response(JSON.stringify(answers.shift()));
  };
  return { calls, fetchImpl, retries: 0, verifyDelays: [0, 0] };
}

const list = ids => ({ data: [{ nmId: 101, recomNms: ids, recomCount: ids.length }] });

test("Рекомендации: артикул продавца и WB, дубли, основной товар и лимит", () => {
  const catalog = index([
    { nmID: 101, vendorCode: "ОСНОВНОЙ" },
    { nmID: 102, vendorCode: "PWB0075" },
    { nmID: 103, vendorCode: "ДЛИННЫЙ КОД" },
  ]);
  assert.deepEqual(checkAdditions("pwb0075, 102\nДЛИННЫЙ КОД", 101, [102], catalog), {
    ids: [102, 103], errors: [], missing: false,
  });
  assert.equal(checkAdditions("101, неизвестный", 101, [], catalog).errors.length, 2);
  const full = Array.from({ length: 20 }, (_, i) => i + 200);
  assert.match(checkAdditions("102", 101, full, catalog).errors[0], /до 20/);
});

test("WB: поиск рекомендаций выбирает точный артикул и отличает пустой ответ от неполного", async () => {
  const wb = mockWB({ data: [{ nmId: 105, recomNms: [102] }] }, list([]), {
    data: [{ nmId: 101, recomCount: 2, recomNms: [103] }],
  });
  assert.deepEqual(await fetchRecommendations("test", 101, wb), []);
  assert.deepEqual(await fetchRecommendations("test", 101, wb), []);
  assert.deepEqual(wb.calls[0], [CONTENT + "/api/content/v1/recommendations/list", { search: "101", limit: 20 }]);
  await assert.rejects(fetchRecommendations("test", 101, wb), /неполный список/);
});

test("WB: снять последнюю рекомендацию пустым списком и подтвердить повторным чтением", async () => {
  const wb = mockWB(list([102]), { isError: false, errors: [] }, list([102]), list([]));
  const result = await saveRecommendations("test", 101, [102], [], wb);
  assert.deepEqual(result, { actual: [], draft: [], confirmed: true, error: "" });
  assert.deepEqual(wb.calls[1], [CONTENT + "/api/content/v1/recommendations/set", {
    recList: [{ nmId: 101, recommendations: [] }], replace: true,
  }]);
});

test("WB: при изменении списка в кабинете сначала показать новый план без записи", async () => {
  const wb = mockWB(list([102, 103, 104]));
  const result = await saveRecommendations("test", 101, [102, 103], [102, 105], wb);
  assert.deepEqual(result, { actual: [102, 103, 104], draft: [102, 104, 105], changedOnWB: true });
  assert.equal(wb.calls.length, 1);
});

test("WB: ошибка отдельных товаров в HTTP 200 и фактический список после записи", async () => {
  const wb = mockWB(list([102]), {
    isError: true, errors: [{ mainNm: "101", recomNm: "104", message: "Нет фото" }],
  }, list([102, 103]), list([102, 103]));
  const result = await saveRecommendations("test", 101, [102], [102, 103, 104], wb);
  assert.deepEqual(result, {
    actual: [102, 103], draft: [102, 103, 104], confirmed: false, error: "104 — Нет фото",
  });
  assert.deepEqual(wb.calls[1][1].recList[0].recommendations, [
    { recomNm: 102, sort: 1 }, { recomNm: 103, sort: 2 }, { recomNm: 104, sort: 3 },
  ]);
  assert.equal(wb.calls.filter(([url]) => url.endsWith("/set")).length, 1);
});

test("WB: остатки рекомендаций запрашиваются только для нужных артикулов и по складам магазина", async () => {
  const wb = mockWB({ data: { items: [{ nmID: 102, metrics: { stockCount: 12 } }] } });
  assert.deepEqual(await fetchStocks("test", "mp", { ...wb, nmIDs: [102, 103] }), { 102: 12 });
  const [[url, body]] = wb.calls;
  assert.equal(url, ANALYTICS + "/api/v2/stocks-report/products/products");
  assert.deepEqual(body.nmIDs, [102, 103]);
  assert.equal(body.stockType, "mp");
});

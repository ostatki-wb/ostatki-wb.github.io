// Тесты логики сайта без обращения к WB. Запуск: npm test
import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "../node_modules/exceljs/lib/exceljs.nodejs.js";
import * as cards from "../js/cards.js";
import * as R from "../js/report.js";
import { toXlsx, parseXlsx } from "../js/excel.js";

class FakeErr extends Error { constructor(body) { super(body); this.body = body; } }

/** Имитирует поведение WB, замеченное на тестовом аккаунте:
 *  - несколько товаров без targetIMT отсоединяются ВМЕСТЕ в одну новую карточку;
 *  - новая карточка становится видна (и доступна как цель) через `delay` чтений;
 *  - отсоединение одиночного товара — ошибка «All cards are in the same group»;
 *  - иногда чтение возвращает устаревшее состояние (stale). */
class FakeWB {
  constructor(list, { delay = 2, stale = 0, seed = 1 } = {}) {
    this.c = new Map(list.map(x => [x.nmID, { ...x }]));
    this.nextImt = 9000; this.hidden = new Map(); this.delay = delay; this.stale = stale;
    this.prev = null; this.seed = seed; this.calls = [];
  }
  rnd() { this.seed = (this.seed * 1103515245 + 12345) % 2147483648; return this.seed / 2147483648; }
  membersOf(imt) { return new Set([...this.c.values()].filter(x => x.imtID === imt).map(x => x.nmID)); }
  tick() { for (const [k, v] of this.hidden) { if (v <= 1) this.hidden.delete(k); else this.hidden.set(k, v - 1); } }
  view() {
    this.tick();
    if (this.prev && this.rnd() < this.stale) return this.prev;
    return new Map([...this.c].map(([k, v]) => [k, { ...v }]));
  }
  snapshot() { this.prev = new Map([...this.c].map(([k, v]) => [k, { ...v }])); }
  async move(nms, target = null) {
    this.calls.push(["move", [...nms], target]);
    if (!nms.length) throw new FakeErr("the nmIDs field cannot be empty");
    this.snapshot();
    if (target === null) {
      const first = this.c.get(nms[0]).imtID, m = this.membersOf(first);
      if (nms.every(n => this.c.get(n).imtID === first) && m.size === nms.length && nms.every(n => m.has(n)))
        throw new FakeErr("All cards are in the same group");
      this.nextImt++;
      for (const n of nms) this.c.get(n).imtID = this.nextImt;
      this.hidden.set(this.nextImt, this.delay);
      return;
    }
    if (this.hidden.has(target) || !this.membersOf(target).size) throw new FakeErr("target imt not found");
    const subj = new Set([...this.membersOf(target)].map(n => this.c.get(n).subjectID));
    for (const n of nms) {
      if (!this.c.has(n)) continue;                       // несуществующий — молча игнорируем, как WB
      assert.ok(subj.has(this.c.get(n).subjectID), "разные предметы");
      this.c.get(n).imtID = target;
    }
  }
  async members(imt) {
    const v = this.view();
    if (this.hidden.has(imt)) return [];
    return [...v.values()].filter(x => x.imtID === imt).map(x => x.nmID).sort((a, b) => a - b);
  }
  async cards(subjectIds) {
    const v = this.view();
    return [...v.values()].filter(x => subjectIds.includes(x.subjectID) && !this.hidden.has(x.imtID));
  }
  isError(e, text) { return String(e.body || e.message).toLowerCase().includes(text.toLowerCase()); }
  async sleep() {}
  groups() {
    const g = new Map();
    for (const x of this.c.values()) { if (!g.has(x.imtID)) g.set(x.imtID, []); g.get(x.imtID).push(x.vendorCode); }
    return [...g.values()].map(a => a.sort()).sort((a, b) => R.cmp(a.join(), b.join()));
  }
}

function mk(n, { subject = 1, prefix = "A", merged = [] } = {}) {
  const imt = {};
  merged.forEach((grp, gi) => grp.forEach(i => { imt[i] = 100 + gi; }));
  const out = [];
  for (let i = 1; i <= n; i++) {
    out.push({ nmID: subject * 1000 + i, imtID: imt[i] ?? subject * 1000 + 500 + i,
      vendorCode: `${prefix}${String(i).padStart(3, "0")}`, subjectID: subject, subjectName: `Предмет ${subject}` });
  }
  return out;
}
const sortG = gs => gs.map(a => [...a].sort()).sort((a, b) => R.cmp(a.join(), b.join()));
const run = (fake, plan) => cards.execute(fake, plan, () => {}, { pause: 0, poll: 0, waitAnchor: 5, waitMove: 5, waitVerify: 5 });
const planFor = (fake, groups) => cards.makePlan(groups, [...fake.c.values()]);
const known = new Set([...Array(20)].map((_, i) => `A${String(i + 1).padStart(3, "0")}`).concat(["CWA8630+", "TEST-SHK-001"]));

test("разбор текста", () => {
  const t = cards.parseText(`Карточка 1: A001, A002, A003
Люди 2 - A004 A005
Карточка 3:
A006
A007

A008 A009
A010;A011
Группа 5
CWA8630+ TEST-SHK-001
`, known);
  assert.equal(t.length, 5);
  assert.deepEqual(t[0], { name: "Карточка 1", items: ["A001", "A002", "A003"] });
  assert.deepEqual(t[1], { name: "Люди 2", items: ["A004", "A005"] });
  assert.deepEqual(t[2].items, ["A006", "A007"]);
  assert.deepEqual(t[3], { name: "Карточка 4", items: ["A008", "A009", "A010", "A011"] });
  assert.deepEqual(t[4], { name: "Группа 5", items: ["CWA8630+", "TEST-SHK-001"] });
});

test("Excel: колонка = карточка, отчёт об остатках и отчёт с названиями читаются", async () => {
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet("x");
  ws.addRow(["Люди 1", null, null, "Микс"]); ws.addRow(["A001", null, null, "A004"]);
  ws.addRow(["A002", null, null, "A005"]); ws.addRow(["A003"]);
  assert.deepEqual(await parseXlsx(await wb.xlsx.writeBuffer(), known, ExcelJS),
    [{ name: "Люди 1", items: ["A001", "A002", "A003"] }, { name: "Микс", items: ["A004", "A005"] }]);
  const rep = await toXlsx([
    { tag: "животные", rows: [{ article: "A001", qty: 5 }, { article: "A002", qty: 3 }], sku: 2, total: 8 },
    { tag: "люди", rows: [{ article: "A003", qty: 1 }], sku: 1, total: 1 }], ExcelJS);
  assert.deepEqual(await parseXlsx(rep, known, ExcelJS),
    [{ name: "животные", items: ["A001", "A002"] }, { name: "люди", items: ["A003"] }]);
  const rep2 = await toXlsx([
    { tag: "животные", named: true, rows: [{ article: "A001", name: "Кот", qty: 5 }], sku: 1, total: 5 },
    { tag: "икона", named: true, rows: [{ article: "A004", name: "Николай Чудотворец", qty: 2 }, { article: "A005", name: "", qty: 1 }], sku: 2, total: 3 }], ExcelJS);
  assert.deepEqual(await parseXlsx(rep2, known, ExcelJS),
    [{ name: "животные", items: ["A001"] }, { name: "икона", items: ["A004", "A005"] }]);
});

test("названия картин", () => {
  for (const [t, a, want] of [
    ["Алмазная мозаика 27х33 см Икона Семистрельная KDA034", "KDA034", "Семистрельная"],
    ["Алмазная мозаика 35х50 cм Икона Николай Чудотворец KC007", "KC007", "Николай Чудотворец"],
    ["Алмазная мозаика на 27х33 см Вера Надежда Любовь KDA039", "KDA039", "Вера Надежда Любовь"],
    ["Вышивка крестиком набор 40х50 см Лиса ZA111", "ZA111", "Лиса"],
    ["Картина по номерам черный холст 40х40 см Волк HWC0027", "HWC0027", "Волк"],
    ["Алмазная мозаика 40х50 см, без подрамника Лебеди SWA0004", "SWA0004", "Лебеди"],
    ["Алмазная мозаика 40х50 на подрамнике природа RWA082", "RWA082", "Природа"],
    ["Алмазная мозаика на подрамнике 30х40 см Икона RWB010", "RWB010", ""],
  ]) assert.equal(R.shortName(t, a), want, t);
});

test("план: проверки", () => {
  const base = mk(40);
  let p = cards.makePlan([{ name: "g", items: ["A001", "Z999"] }], base);
  assert.ok(!p.ok && p.errors.join(" ").includes("Z999"));
  p = cards.makePlan([{ name: "g1", items: ["A001", "A002"] }, { name: "g2", items: ["A002", "A003"] }], base);
  assert.ok(!p.ok && p.groups.every(G => G.errors.length));
  p = cards.makePlan([{ name: "big", items: [...Array(31)].map((_, i) => `A${String(i + 1).padStart(3, "0")}`) }], base);
  assert.ok(!p.ok);
  p = cards.makePlan([{ name: "ok30", items: [...Array(30)].map((_, i) => `A${String(i + 1).padStart(3, "0")}`) }], base);
  assert.ok(p.ok);
  p = cards.makePlan([{ name: "mix", items: ["A001", "B001"] }], mk(5).concat(mk(5, { subject: 2, prefix: "B" })));
  assert.ok(!p.ok && p.errors.join(" ").includes("предмет"));
  p = cards.makePlan([{ name: "g", items: ["a001", "A002", "A002"] }], base);
  assert.ok(p.ok && p.groups[0].items.length === 2 && p.groups[0].warnings.length);
  p = cards.makePlan([{ name: "g", items: ["A001", "A002"] }], mk(5, { merged: [[1, 2]] }));
  assert.ok(p.groups[0].action === "skip" && p.todo === 0);
  p = cards.makePlan([{ name: "g", items: ["A001", "A002"] }], mk(5, { merged: [[1, 2, 3]] }));
  assert.ok(p.ok && p.warnings.join(" ").includes("A003"));
  p = cards.makePlan([{ name: "g", items: [1001, 99999] }], mk(3));
  assert.ok(!p.ok, "несуществующий nmID");
});

test("выполнение: сборка из одиночных без отсоединения", async () => {
  const f = new FakeWB(mk(10));
  const res = await run(f, planFor(f, [{ name: "g1", items: ["A001", "A002", "A003"] }, { name: "g2", items: ["A004", "A005"] }]));
  assert.deepEqual(f.groups(), sortG([["A001", "A002", "A003"], ["A004", "A005"], ["A006"], ["A007"], ["A008"], ["A009"], ["A010"]]));
  assert.ok(!f.calls.some(c => c[2] === null));
  assert.ok(res.every(r => r.ok));
});

test("выполнение: перетасовка, отсоединение по одному, откат", async () => {
  const f = new FakeWB(mk(8, { merged: [[1, 2, 3, 4], [5, 6]] }), { delay: 3 });
  const p = planFor(f, [{ name: "x", items: ["A001", "A005"] }, { name: "y", items: ["A002", "A006"] }, { name: "z", items: ["A003", "A004"] }]);
  const res = await run(f, p);
  assert.deepEqual(f.groups(), sortG([["A001", "A005"], ["A002", "A006"], ["A003", "A004"], ["A007"], ["A008"]]));
  assert.ok(res.every(r => r.ok));
  assert.ok(f.calls.filter(c => c[2] === null).every(c => c[1].length === 1));
  await run(f, planFor(f, cards.snapshotGroups(p.before)));
  assert.deepEqual(f.groups(), sortG([["A001", "A002", "A003", "A004"], ["A005", "A006"], ["A007"], ["A008"]]));
});

test("выполнение: один товар — в отдельную карточку", async () => {
  const f = new FakeWB(mk(4, { merged: [[1, 2, 3]] }));
  const res = await run(f, planFor(f, [{ name: "solo", items: ["A002"] }]));
  assert.deepEqual(f.groups(), sortG([["A001", "A003"], ["A002"], ["A004"]]));
  assert.ok(res[0].ok);
});

test("выполнение: устаревшие чтения и долгое появление карточки", async () => {
  for (let seed = 1; seed <= 5; seed++) {
    const f = new FakeWB(mk(30, { merged: [[...Array(15)].map((_, i) => i + 1), [...Array(15)].map((_, i) => i + 16)] }), { delay: 4, stale: 0.4, seed });
    const groups = [1, 2, 3].map(k => ({ name: `g${k}`, items: [...Array(30)].map((_, i) => i + 1).filter(i => (i - k) % 3 === 0).map(i => `A${String(i).padStart(3, "0")}`) }));
    const res = await run(f, planFor(f, groups));
    assert.deepEqual(f.groups(), sortG(groups.map(g => g.items)), `seed ${seed}`);
    assert.ok(res.every(r => r.ok), `seed ${seed}`);
  }
});

test("отчёт: названия в алмазке у всех ярлыков, в других категориях нет", async () => {
  const cache = { cards: [
    { nmID: 1, vendorCode: "KC007", subjectName: "Алмазная мозаика", tags: ["икона"], title: "Алмазная мозаика 35х50 cм Икона Николай Чудотворец KC007" },
    { nmID: 2, vendorCode: "X1", subjectName: "Алмазная мозаика", tags: ["животные"], title: "Алмазная мозаика 40х50 на подрамнике Кот X1" },
    { nmID: 3, vendorCode: "K1", subjectName: "Конфеты", tags: ["животные"], title: "Конфеты K1" }],
    stocks: { all: { 1: 3, 2: 4, 3: 5 } } };
  const g = R.build(cache, "Алмазная мозаика", ["икона", "животные"], "all", false);
  assert.ok(g.every(x => x.named));
  assert.deepEqual(g.map(x => x.rows[0].name), ["Кот", "Николай Чудотворец"]);
  const g2 = R.build(cache, "Конфеты", ["животные"], "all", false);
  assert.ok(!g2[0].named && !("name" in g2[0].rows[0]));
});

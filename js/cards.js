// Пересборка карточек WB (перенос app/cards.py): разбор списка, план «было → станет», выполнение, откат.
//
// Как WB объединяет карточки (проверено на тестовом аккаунте):
// - moveNm с targetIMT переносит товары в карточку, сам вынимая их из старых;
// - moveNm без targetIMT отсоединяет; несколько товаров одним запросом попадут ВМЕСТЕ в новую карточку,
//   поэтому отсоединяем строго по одному;
// - в новую карточку можно переносить, только когда она появилась в списке карточек;
// - на несуществующий товар WB отвечает успехом — результат перепроверяем сами;
// - список карточек после изменения несколько секунд может показывать старое состояние.
//
// Способ сборки (как вручную): для группы берём товар-якорь, отсоединяем его от всех карточек и
// переносим к нему остальные. Если текущая карточка уже состоит только из товаров группы — она основа.

import { cmp } from "./report.js";

export const MAX_GROUP = 30;
const HEADER_RE = /^(карточка|группа|склейка)(?![\p{L}\p{N}_])/iu;
const SEP_RE = /^(.*?)\s*(?::|\s[-–—]\s)\s*(.*)$/u;
const tokens = s => String(s).trim().split(/[\s,;]+/u).filter(Boolean);

// ---------- разбор списка ----------

/** «Карточка 1: A1, A2» / заголовок и артикулы столбиком / пустая строка — новая группа. */
export function parseText(text, known = new Set()) {
  const groups = [];
  let cur = null;
  const add = name => {
    const g = { name: String(name || "").trim() || `Карточка ${groups.length + 1}`, items: [] };
    groups.push(g);
    return g;
  };
  for (const raw of String(text || "").split(/\r\n|\r|\n/)) {
    const line = raw.trim();
    if (!line) { cur = null; continue; }
    const m = SEP_RE.exec(line);
    if (m && m[1] && !known.has(m[1])) {
      cur = add(m[1]);
      cur.items.push(...tokens(m[2]));
      continue;
    }
    if (HEADER_RE.test(line) && !tokens(line).every(t => known.has(t))) {
      cur = add(line);
      continue;
    }
    if (!cur) cur = add("");
    cur.items.push(...tokens(line));
  }
  return groups.filter(g => g.items.length);
}

// ---------- план ----------

export function index(cards) {
  const byVc = new Map(), byLow = new Map(), byNm = new Map();
  for (const c of cards) {
    byNm.set(c.nmID, c);
    byVc.set(c.vendorCode, c);
    const k = c.vendorCode.toLowerCase();
    if (!byLow.has(k)) byLow.set(k, []);
    byLow.get(k).push(c);
  }
  return { byVc, byLow, byNm };
}

export function find(item, { byVc, byLow, byNm }) {
  if (typeof item === "number") return byNm.get(item) || null;
  const c = byVc.get(item);
  if (c) return c;
  const lst = byLow.get(String(item).toLowerCase()) || [];
  return lst.length === 1 ? lst[0] : null;
}

const numSort = arr => [...arr].sort((a, b) => a - b);

/** groupsIn: [{name, items: [артикул или nmID]}]; cards: актуальные карточки нужных предметов. */
export function makePlan(groupsIn, cards) {
  const idx = index(cards);
  const members = new Map();
  for (const c of cards) {
    if (!members.has(c.imtID)) members.set(c.imtID, new Set());
    members.get(c.imtID).add(c.nmID);
  }
  const plan = { groups: [], errors: [], warnings: [], ok: false };
  if (!groupsIn.length) {
    plan.errors.push("Список пуст: не найдено ни одной группы с артикулами.");
    return plan;
  }
  const owner = new Map();  // nmID -> индекс группы
  groupsIn.forEach((g, gi) => {
    const G = { name: g.name, items: [], errors: [], warnings: [], action: null };
    const seen = new Set(), unknown = [], dups = [];
    for (const it of g.items) {
      const c = find(it, idx);
      if (!c) { unknown.push(String(it)); continue; }
      if (seen.has(c.nmID)) { dups.push(c.vendorCode); continue; }
      seen.add(c.nmID);
      G.items.push({ vc: c.vendorCode, nm: c.nmID, imt: c.imtID, subject: c.subjectName, subjectID: c.subjectID });
      if (owner.has(c.nmID) && owner.get(c.nmID) !== gi) {
        G.errors.push(`${c.vendorCode} уже есть в группе «${groupsIn[owner.get(c.nmID)].name}»`);
      }
      if (!owner.has(c.nmID)) owner.set(c.nmID, gi);
    }
    if (unknown.length) G.errors.push("Не найдены артикулы: " + unknown.join(", "));
    if (dups.length) G.warnings.push("Повторы убраны: " + dups.join(", "));
    if (G.items.length > MAX_GROUP) G.errors.push(`${G.items.length} товаров — больше ${MAX_GROUP}, WB столько в одну карточку не объединяет`);
    const subjects = [...new Set(G.items.map(i => i.subject))].sort(cmp);
    G.subject = subjects.join(", ");
    if (subjects.length > 1) G.errors.push("Разные предметы в одной карточке нельзя: " + subjects.join(", "));
    plan.groups.push(G);
  });
  // ошибка «уже в другой группе» должна быть видна в обеих группах
  plan.groups.forEach((G, gi) => {
    for (const it of G.items) {
      if (owner.get(it.nm) !== gi) {
        const first = plan.groups[owner.get(it.nm)];
        const msg = `${it.vc} есть ещё в группе «${G.name}»`;
        if (!first.errors.includes(msg)) first.errors.push(msg);
      }
    }
  });
  const listed = new Set(owner.keys());
  const touched = new Set();
  for (const G of plan.groups) {
    const gset = new Set(G.items.map(i => i.nm));
    G.cardsNow = new Set(G.items.map(i => i.imt)).size;
    G.items.forEach(i => touched.add(i.imt));
    if (G.errors.length || !gset.size) { G.action = "error"; continue; }
    // основа: текущая карточка, где лежат только товары этой группы (самая большая)
    let best = null;
    for (const it of G.items) {
      const m = members.get(it.imt) || new Set();
      const subset = [...m].every(n => gset.has(n));
      if (subset && (best === null || m.size > members.get(best).size)) best = it.imt;
    }
    if (best !== null) {
      G.anchor = null;
      G.target = best;
      G.move = G.items.filter(i => i.imt !== best).map(i => i.nm);
      G.action = G.move.length ? "move" : "skip";
    } else {
      const a = G.items[0];
      G.anchor = { nm: a.nm, vc: a.vc, oldImt: a.imt };
      G.target = null;
      G.move = G.items.slice(1).map(i => i.nm);
      G.action = "anchor";
    }
  }
  const leftovers = [];
  for (const imt of numSort(touched)) {
    const rest = [...(members.get(imt) || [])].filter(n => !listed.has(n)).map(n => idx.byNm.get(n).vendorCode).sort(cmp);
    if (rest.length) leftovers.push(rest);
  }
  if (leftovers.length) {
    plan.warnings.push("Эти товары не указаны в списке и останутся в своих старых карточках: " +
      leftovers.map(r => r.join(", ")).join("; "));
  }
  plan.leftovers = leftovers;
  plan.cardsBefore = touched.size;
  plan.cardsAfter = plan.groups.filter(G => G.items.length).length + leftovers.length;
  for (const G of plan.groups) for (const e of G.errors) plan.errors.push(`«${G.name}»: ${e}`);
  plan.todo = plan.groups.filter(G => G.action === "move" || G.action === "anchor").length;
  plan.ok = !plan.errors.length;
  plan.subjectIDs = numSort(new Set(plan.groups.flatMap(G => G.items.map(i => i.subjectID))));
  // снимок «как было»: все товары затронутых карточек
  plan.before = numSort(touched).map(imt => numSort(members.get(imt)));
  plan.vc = {};
  for (const imt of touched) for (const n of members.get(imt)) plan.vc[String(n)] = idx.byNm.get(n).vendorCode;
  return plan;
}

/** Группы для отката: как товары лежали до выполнения плана. */
export function snapshotGroups(before) {
  return before.map((g, i) => ({ name: `Как было ${i + 1}`, items: [...g] }));
}

// ---------- выполнение ----------

/** Обёртка над wb.js (в тестах подменяется). */
export class Api {
  constructor(wb, token) { this.wb = wb; this.token = token; }
  move(nms, target = null) { return this.wb.moveNm(this.token, nms, target); }
  members(imt) { return this.wb.imtMembers(this.token, imt); }
  cards(subjectIds) { return this.wb.fetchCards(this.token, { subjectIds }); }
  isError(e, text) { return e instanceof this.wb.WBError && (e.body || e.message).toLowerCase().includes(text.toLowerCase()); }
  sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
}

/** Выполняет план. Возвращает результаты по группам. Время — в секундах. */
export async function execute(api, plan, log, { pause = 1, waitAnchor = 180, waitMove = 90, waitVerify = 180, poll = 5 } = {}) {
  const now = () => Date.now() / 1000;
  const groups = plan.groups.filter(G => G.action === "move" || G.action === "anchor");
  const res = new Map(groups.map(G => [G, { name: G.name, ok: null, msg: "" }]));
  for (const G of plan.groups) if (G.action === "skip") log(`«${G.name}»: уже собрана как нужно — пропускаю`);
  const fail = (G, msg) => { Object.assign(res.get(G), { ok: false, msg }); log(`«${G.name}»: ОШИБКА — ${msg}`); };
  const pending = G => res.get(G).ok === null;

  // 1. отсоединяем якоря (строго по одному товару на запрос)
  const anchors = groups.filter(G => G.action === "anchor");
  for (let n = 0; n < anchors.length; n++) {
    const G = anchors[n], a = G.anchor;
    log(`Шаг 1/4 · отсоединяю ${a.vc} — основа для «${G.name}» (${n + 1}/${anchors.length})`);
    try { await api.move([a.nm]); }
    catch (e) {
      if (api.isError(e, "same group")) a.oldImt = null;   // уже был один в карточке
      else fail(G, `не удалось отсоединить ${a.vc}: ${e.message}`);
    }
    await api.sleep(pause * 1000);
  }

  // 2. ждём, пока WB покажет новые карточки-основы
  let wait = anchors.filter(pending);
  if (wait.length) log(`Шаг 2/4 · жду, пока WB создаст ${wait.length} новых карточек…`);
  let t0 = now();
  while (wait.length && now() - t0 < waitAnchor) {
    await api.sleep(poll * 1000);
    let cur;
    try { cur = new Map((await api.cards(plan.subjectIDs)).map(c => [c.nmID, c.imtID])); }
    catch (e) { log(`   не удалось прочитать карточки (${e.message}), пробую ещё раз`); continue; }
    for (const G of [...wait]) {
      const a = G.anchor, imt = cur.get(a.nm);
      if (imt == null || imt === a.oldImt) continue;
      let ok = false;
      try {
        const same = m => m.length === 1 && m[0] === a.nm;
        ok = same(await api.members(imt));
        if (ok) { await api.sleep(2000); ok = same(await api.members(imt)); }   // второе чтение — от устаревших данных
      } catch (e) { ok = false; }
      if (ok) {
        G.target = imt;
        wait = wait.filter(x => x !== G);
        log(`   «${G.name}»: новая карточка ${imt} готова`);
      }
    }
  }
  for (const G of wait) fail(G, `WB не создал отдельную карточку для ${G.anchor.vc} за ${waitAnchor} с`);

  // 3. переносим остальные товары в карточку-основу
  let work = groups.filter(pending);
  for (let n = 0; n < work.length; n++) {
    const G = work[n];
    if (!G.move.length) continue;
    log(`Шаг 3/4 · переношу ${G.move.length} товаров в «${G.name}» (${n + 1}/${work.length})`);
    const t1 = now();
    for (;;) {
      try { await api.move(G.move, G.target); break; }
      catch (e) {
        if (api.isError(e, "target imt not found") && now() - t1 < waitMove) { await api.sleep(poll * 1000); continue; }
        fail(G, `не удалось перенести товары: ${e.message}`);
        break;
      }
    }
    await api.sleep(pause * 1000);
  }

  // 4. проверяем: в каждой карточке ровно нужные товары (два совпавших чтения подряд)
  work = groups.filter(pending);
  if (work.length) log(`Шаг 4/4 · проверяю результат (${work.length} карточек)…`);
  const good = new Map(), last = new Map();
  t0 = now();
  while (work.length && now() - t0 < waitVerify) {
    let byImt;
    try {
      byImt = new Map();
      for (const c of await api.cards(plan.subjectIDs)) {
        if (!byImt.has(c.imtID)) byImt.set(c.imtID, new Set());
        byImt.get(c.imtID).add(c.nmID);
      }
    } catch (e) {
      log(`   не удалось прочитать карточки (${e.message}), пробую ещё раз`);
      await api.sleep(poll * 1000);
      continue;
    }
    for (const G of [...work]) {
      const want = new Set(G.items.map(i => i.nm));
      const got = byImt.get(G.target) || new Set();
      last.set(G, got);
      if (got.size === want.size && [...want].every(n => got.has(n))) {
        good.set(G, (good.get(G) || 0) + 1);
        if (good.get(G) >= 2) {
          Object.assign(res.get(G), { ok: true, msg: `собрана: ${want.size} товаров, карточка ${G.target}` });
          log(`   «${G.name}»: готово ✓`);
          work = work.filter(x => x !== G);
        }
      } else good.set(G, 0);
    }
    if (work.length) await api.sleep(poll * 1000);
  }
  const vc = new Map(groups.flatMap(G => G.items.map(i => [i.nm, i.vc])));
  for (const G of work) {
    const want = new Set(G.items.map(i => i.nm)), got = last.get(G) || new Set();
    const miss = [...want].filter(n => !got.has(n)).map(n => vc.get(n) || String(n)).sort(cmp);
    const extra = [...got].filter(n => !want.has(n)).map(String).sort(cmp);
    let msg = "WB не подтвердил результат";
    if (miss.length) msg += "; не попали: " + miss.join(", ");
    if (extra.length) msg += "; лишние в карточке (nmID): " + extra.join(", ");
    fail(G, msg);
  }
  return groups.map(G => res.get(G));
}

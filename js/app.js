// Интерфейс сайта «Остатки WB». Всё выполняется в браузере пользователя.
import * as W from "./wb.js";
import * as R from "./report.js";
import * as C from "./cards.js";
import * as S from "./store.js";
import { toXlsx, parseXlsx } from "./excel.js";

const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const nowStr = () => new Date().toLocaleString("ru-RU", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit" }).replace(",", "");

let st = S.getState();
let cache = null;          // данные текущего магазина
let selected = new Set();  // выбранные ярлыки
let busy = false;
let job = null;            // {kind, running, message, log, error, result}
let plan = null;           // последний план «Карточек» {…, shopId}
let planInput = null;
let xlsxData = null;

const shop = () => st.shops.find(s => s.id === st.current) || null;
const save = () => S.saveState(st);
const ui = (k, v) => { if (v === undefined) return st.ui[k]; st.ui[k] = v; save(); };

function setStatus(kind, html) { const el = $("#status"); el.className = "status " + (kind || ""); el.innerHTML = html || ""; }
function setBusy(b) {
  busy = b;
  $("#btnRefresh").disabled = b;
  $("#btnPlan").disabled = b;
  $("#shopSelect").disabled = b;
  if (plan) renderPlan();
}
window.addEventListener("beforeunload", e => { if (busy) { e.preventDefault(); e.returnValue = ""; } });

// ---------- магазины ----------

function renderShopSelect() {
  const sel = $("#shopSelect");
  sel.innerHTML = st.shops.map(s => `<option value="${esc(s.id)}">${esc(s.name)}</option>`).join("") +
    `<option value="__add">＋ Добавить магазин…</option>`;
  if (st.current) sel.value = st.current;
  sel.style.display = st.shops.length ? "" : "none";
}

async function selectShop(id) {
  st.current = id; save();
  cache = id ? (await S.kvGet(S.cacheKey(id))) || null : null;
  plan = null; planInput = null; job = null; xlsxData = null;
  renderShopSelect();
  await renderAll();
}

function tokenHint(t) { return t && t.length > 12 ? `…${t.slice(-6)}` : ""; }
function cleanToken(t) {
  t = String(t || "").trim().replace(/^["']|["']$/g, "");
  if (/^bearer\s+/i.test(t)) t = t.replace(/^bearer\s+/i, "");
  return t.trim();
}

let editing = null;  // id редактируемого магазина или null — добавление
function openShops(mode) {
  $("#shopSource").innerHTML = Object.entries(R.SOURCES).map(([k, v]) => `<option value="${k}">${esc(v)}</option>`).join("");
  renderShopList();
  if (mode === "add" || !st.shops.length) startEdit(null); else startEdit(st.current);
  $("#shopsModal").classList.add("open");
}
function closeShops() { $("#shopsModal").classList.remove("open"); }

function renderShopList() {
  $("#shopList").innerHTML = st.shops.length ? st.shops.map(s => `
    <div class="shop"><div class="grow"><b>${esc(s.name)}</b><span class="muted">${esc(s.sellerName || "")} · токен ${esc(tokenHint(s.token))}</span></div>
    <button class="ghost" data-edit="${esc(s.id)}">Изменить</button><button class="ghost" data-del="${esc(s.id)}">Удалить</button></div>`).join("")
    : `<div class="muted">Пока нет ни одного магазина.</div>`;
  $$("#shopList [data-edit]").forEach(b => b.onclick = () => startEdit(b.dataset.edit));
  $$("#shopList [data-del]").forEach(b => b.onclick = async () => {
    const s = st.shops.find(x => x.id === b.dataset.del);
    if (!confirm(`Удалить магазин «${s.name}» из этого браузера?\n\nТокен, загруженные данные и история откатов этого магазина будут удалены. В самом WB ничего не изменится.`)) return;
    await S.deleteShopData(s.id);
    st.shops = st.shops.filter(x => x.id !== s.id);
    if (st.current === s.id) st.current = st.shops[0] ? st.shops[0].id : null;
    save();
    renderShopList(); startEdit(null);
    await selectShop(st.current);
    if (!st.shops.length) closeShops();
  });
}

function startEdit(id) {
  editing = id;
  const s = id ? st.shops.find(x => x.id === id) : null;
  $("#shopFormTitle").textContent = s ? `Изменить «${s.name}»` : "Добавить магазин";
  $("#shopName").value = s ? s.name : "";
  $("#shopToken").value = "";
  $("#shopToken").placeholder = s ? "Оставьте пустым, чтобы не менять" : "Вставьте токен";
  $("#shopTokenHint").textContent = s ? `Сохранён токен ${tokenHint(s.token)}.` : "";
  $("#shopSource").value = s ? (s.source || "all") : "all";
  $("#saveShop").textContent = s ? "Сохранить" : "Добавить";
  $("#cancelShop").style.display = s ? "" : "none";
}

async function saveShop() {
  const btn = $("#saveShop");
  const s = editing ? st.shops.find(x => x.id === editing) : null;
  const token = cleanToken($("#shopToken").value);
  let name = $("#shopName").value.trim();
  const source = $("#shopSource").value;
  const hint = $("#shopTokenHint");
  if (!s && !token) { hint.innerHTML = `<span style="color:var(--err)">Вставьте токен WB.</span>`; return; }
  if (token && token.length < 20) { hint.innerHTML = `<span style="color:var(--err)">Токен слишком короткий — скопируйте его целиком.</span>`; return; }
  let sellerName = s ? s.sellerName : "";
  if (token) {
    btn.disabled = true; hint.innerHTML = `<span class="spin"></span>Проверяю токен в WB…`;
    try {
      await W.request("GET", W.CONTENT + "/content/v2/tags", token, undefined, { retries: 1 });
      sellerName = await W.sellerName(token);
    } catch (e) {
      btn.disabled = false;
      hint.innerHTML = `<span style="color:var(--err)">${esc(e.message)}</span>`;
      return;
    }
    btn.disabled = false;
  }
  if (!name) name = sellerName || `Магазин ${st.shops.length + 1}`;
  const sourceChanged = s && (s.source || "all") !== source;
  if (s) {
    Object.assign(s, { name, source, sellerName });
    if (token) s.token = token;
  } else {
    const id = S.newId();
    st.shops.push({ id, name, token, source, sellerName });
    st.current = id;
  }
  save();
  closeShops();
  await selectShop(st.current);
  if (!s || token || sourceChanged) refresh();
}

// ---------- обновление данных ----------

async function refresh() {
  const s = shop();
  if (!s || busy) return;
  setBusy(true);
  const progress = m => setStatus("info", `<span class="spin"></span>${esc(m)}`);
  const notes = [];
  try {
    progress("Загружаю карточки…");
    const cards = await W.fetchCards(s.token, { progress });
    progress("Загружаю список ярлыков…");
    let tags = [];
    try { tags = await W.fetchTags(s.token); }
    catch (e) { notes.push(`Список ярлыков не загрузился (${e.message}); ярлыки взяты из карточек.`); }
    progress("Загружаю остатки…");
    const source = s.source || "all";
    const stocks = { [source]: await W.fetchStocks(s.token, R.STOCK_TYPE[source], { progress }) };
    if (!s.sellerName) { s.sellerName = await W.sellerName(s.token); save(); }
    const d = new Date(), p = n => String(n).padStart(2, "0");
    cache = { updated: `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`,
      cards, tags, stocks, notes };
    if (shop() && shop().id === s.id) await S.kvSet(S.cacheKey(s.id), cache);
    setStatus("", "");
  } catch (e) {
    setStatus("err", esc(e.message));
  } finally {
    setBusy(false);
    await renderStocks();
  }
}

// ---------- вкладка «Остатки» ----------

function colorOf(c) {
  const m = { D1CFD7: "#9a96a6", FEE0E0: "#f08b8b", ECDAFF: "#b48af5", E4EAFF: "#7f9cf5", DEF1DD: "#6cc36a", FFECC7: "#f2b843" };
  if (!c) return "#bbb";
  c = String(c).replace("#", "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(c) ? (m[c] || "#" + c) : "#bbb";
}

async function renderStocks() {
  const s = shop();
  $("#upd").textContent = cache && cache.updated ? `Данные от ${cache.updated} · карточек: ${cache.cards.length}` : (s ? "Данные ещё не загружены" : "");
  const notes = (cache && cache.notes) || [];
  $("#notes").style.display = notes.length ? "block" : "none";
  $("#notes").innerHTML = notes.map(esc).join("<br>");
  const sel = $("#category");
  const cats = cache ? R.categories(cache) : [];
  const cur = sel.value || ui(`cat:${s && s.id}`) || "";
  sel.innerHTML = cats.length ? cats.map(c => `<option value="${esc(c.name)}">${esc(c.name)} (${c.count})</option>`).join("") : `<option value="">— нет данных —</option>`;
  if (cats.some(c => c.name === cur)) sel.value = cur;
  const pick = cats.find(c => c.name === sel.value);
  $("#catInfo").textContent = pick ? `${pick.count} карточек` : "";
  if (!busy && s && !cache) setStatus("info", "Нажмите «Обновить данные», чтобы загрузить товары и остатки из WB.");
  renderTags();
}

function renderTags() {
  const s = shop(), cat = $("#category").value;
  if (!cat || !cache) { $("#tags").innerHTML = ""; renderReport(null); return; }
  ui(`cat:${s.id}`, cat);
  const tags = R.tagsForCategory(cache, cat);
  const saved = ui(`tags:${s.id}:${cat}`);
  selected = new Set((saved || tags.map(t => t.name)).filter(n => tags.some(t => t.name === n)));
  $("#tags").innerHTML = tags.map(t => `<span class="chip ${selected.has(t.name) ? "on" : ""}" data-n="${esc(t.name)}"><span class="dot" style="background:${colorOf(t.color)}"></span>${esc(t.name)} <span class="n">${t.count}</span></span>`).join("")
    || `<span class="muted">В этой категории нет ярлыков</span>`;
  $$(".chip").forEach(ch => ch.onclick = () => {
    const n = ch.dataset.n;
    selected.has(n) ? selected.delete(n) : selected.add(n);
    ch.classList.toggle("on"); saveTags(); renderReport(currentReport());
  });
  renderReport(currentReport());
}
function saveTags() { const s = shop(); ui(`tags:${s.id}:${$("#category").value}`, [...selected]); }

function currentReport() {
  const s = shop();
  if (!cache || !$("#category").value || !selected.size) return null;
  const stocks = cache.stocks || {};
  const source = (s.source || "all") in stocks ? (s.source || "all") : (Object.keys(stocks)[0] || "all");
  return { groups: R.build(cache, $("#category").value, [...selected], source, $("#zero").checked), source: R.SOURCES[source] || source };
}

function renderReport(rep) {
  const grid = $("#grid"), sum = $("#summary");
  if (!rep || !rep.groups.length) { grid.innerHTML = `<div class="empty">Выберите категорию и ярлыки</div>`; sum.innerHTML = ""; $("#btnExcel").disabled = true; return; }
  const sku = rep.groups.reduce((a, g) => a + g.sku, 0), tot = rep.groups.reduce((a, g) => a + g.total, 0);
  sum.innerHTML = `<div><b>${sku}</b> <span>артикулов</span></div><div><b>${tot}</b> <span>шт. всего</span></div><div><span>Источник: ${esc(rep.source)}</span></div>`;
  grid.innerHTML = rep.groups.map(g => {
    const n = g.named ? 3 : 2;
    return `<table><thead><tr><th colspan="${n}">${esc(g.tag)}</th></tr><tr><th>Артикул</th>${g.named ? "<th>Название</th>" : ""}<th>Кол-во</th></tr></thead>
      <tbody>${g.rows.map(r => `<tr><td>${esc(r.article)}</td>${g.named ? `<td>${esc(r.name || "")}</td>` : ""}<td class="q">${r.qty}</td></tr>`).join("") || `<tr><td colspan="${n}" class="muted">нет товаров</td></tr>`}</tbody>
      <tfoot><tr><td>Итого SKU:</td>${g.named ? "<td></td>" : ""}<td class="q">${g.sku}</td></tr></tfoot></table>`;
  }).join("");
  $("#btnExcel").disabled = false;
}

async function downloadExcel() {
  const rep = currentReport();
  if (!rep) return;
  try {
    const buf = await toXlsx(rep.groups, window.ExcelJS);
    const d = new Date(), p = n => String(n).padStart(2, "0");
    saveFile(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      `Остатки ${p(d.getDate())}.${p(d.getMonth() + 1)} ${$("#category").value}.xlsx`);
  } catch (e) { setStatus("err", "Не удалось собрать Excel: " + esc(e.message)); }
}

function saveFile(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 10000);
}

// ---------- вкладка «Карточки» ----------

function log(msg) {
  const t = new Date().toLocaleTimeString("ru-RU");
  job.message = msg; job.log.push(`${t}  ${msg}`);
  renderJob();
}

async function freshCardsFor(s, groups) {
  let known = (cache && cache.cards) || [];
  const resolve = list => {
    const idx = C.index(list), subj = new Set();
    let missing = 0;
    for (const g of groups) for (const it of g.items) { const c = C.find(it, idx); if (c) subj.add(c.subjectID); else missing++; }
    return { subj, missing };
  };
  let r = resolve(known);
  if (r.missing || !known.length) {
    log("Читаю весь каталог карточек из WB…");
    known = await W.fetchCards(s.token, { progress: m => { job.message = m; renderJob(); } });
    if (cache) { cache.cards = known; await S.kvSet(S.cacheKey(s.id), cache); }
    r = resolve(known);
  }
  if (!r.subj.size) return known;
  log("Читаю актуальное состояние карточек нужных предметов…");
  return W.fetchCards(s.token, { subjectIds: [...r.subj].sort((a, b) => a - b) });
}

async function startJob(kind, fn) {
  if (busy) return;
  job = { kind, running: true, message: "Начинаю…", log: [], error: "", result: null };
  setBusy(true); renderJob();
  try { await fn(); }
  catch (e) { job.error = e.message || String(e); }
  finally { job.running = false; setBusy(false); renderJob(); renderPlan(); renderSnaps(); }
}

async function makePlanFromInput() {
  const s = shop();
  const known = new Set(((cache && cache.cards) || []).map(c => c.vendorCode));
  let groups;
  try {
    groups = xlsxData ? await parseXlsx(xlsxData, known, window.ExcelJS) : C.parseText($("#cardsText").value, known);
  } catch (e) { alert("Не удалось прочитать список: " + e.message); return; }
  if (!groups.length) { alert("В списке не найдено ни одного артикула."); return; }
  plan = null; renderPlan();
  await startJob("plan", async () => {
    const list = await freshCardsFor(s, groups);
    const p = C.makePlan(groups, list);
    p.shopId = s.id;
    plan = p; planInput = groups;
    log("Проверка закончена: " + (p.ok ? "можно выполнять" : "есть ошибки, исправьте список"));
  });
}

async function runGroups(groups, label, rollbackOf = null) {
  const s = shop();
  await startJob(label.startsWith("Откат") ? "rollback" : "run", async () => {
    const list = await freshCardsFor(s, groups);
    const p = C.makePlan(groups, list);
    if (!p.ok) { plan = Object.assign(p, { shopId: s.id }); throw new Error("Состояние карточек изменилось, план больше не годится: " + p.errors.slice(0, 3).join("; ")); }
    const stamp = new Date();
    const key = S.snapPrefix(s.id) + stamp.toISOString();
    const snap = { time: nowStr(), shop: s.name, label, groups: p.groups.length, todo: p.todo, before: p.before, vc: p.vc, rollbackOf };
    await S.kvSet(key, snap);
    log(`Снимок «как было» сохранён (${snap.time}) — по нему можно откатить.`);
    const results = await C.execute(new C.Api(W, s.token), p, log);
    const ok = results.filter(r => r.ok).length, bad = results.filter(r => !r.ok);
    let summary = `Готово: собрано ${ok} из ${results.length}` + (bad.length ? `, ошибок ${bad.length}` : "");
    const skipped = p.groups.filter(G => G.action === "skip").length;
    if (skipped) summary += `; уже были собраны и не менялись: ${skipped}`;
    log(summary);
    Object.assign(snap, { results, summary });
    await S.kvSet(key, snap);
    if (rollbackOf) {
      const old = await S.kvGet(rollbackOf);
      if (old) { old.rolledBack = snap.time; await S.kvSet(rollbackOf, old); }
    }
    job.result = { summary, results };
    plan = null;
  });
}

function renderPlan() {
  const panel = $("#planPanel"), s = shop();
  if (!plan || !s || plan.shopId !== s.id) { panel.style.display = "none"; return; }
  const p = plan;
  panel.style.display = "";
  const items = p.groups.reduce((a, g) => a + g.items.length, 0);
  $("#planSummary").innerHTML = `<div><b>${p.groups.length}</b> <span>групп</span></div><div><b>${items}</b> <span>товаров</span></div>` +
    `<div><b>${p.cardsBefore} → ${p.cardsAfter}</b> <span>карточек сейчас → станет</span></div><div><b>${p.todo}</b> <span>к выполнению</span></div>`;
  $("#planErrors").innerHTML = p.errors.map(e => `<li>${esc(e)}</li>`).join("");
  $("#planWarnings").innerHTML = p.warnings.concat(...p.groups.map(g => g.warnings.map(w => `«${g.name}»: ${w}`))).map(e => `<li>${esc(e)}</li>`).join("");
  const ACT = { skip: ["skip", "уже собрана"], move: ["move", "добавить в готовую"], anchor: ["anchor", "собрать заново"], error: ["error", "ошибка"] };
  $("#planTable").innerHTML = `<thead><tr><th>Карточка</th><th>Товаров</th><th>Предмет</th><th>Сейчас в карточках</th><th>Что будет</th><th>Артикулы</th></tr></thead><tbody>` +
    p.groups.map(g => {
      const [cls, lbl] = ACT[g.action] || ["error", "?"];
      let what = `<span class="tag ${cls}">${lbl}</span>`;
      if (g.action === "anchor") what += `<div class="muted">отсоединить ${esc(g.anchor.vc)}, перенести к нему ${g.move.length}</div>`;
      if (g.action === "move") what += `<div class="muted">перенести ${g.move.length}</div>`;
      if (g.action === "error") what += `<div style="color:var(--err);font-size:13px">${g.errors.map(esc).join("<br>")}</div>`;
      return `<tr><td><b>${esc(g.name)}</b></td><td class="q">${g.items.length}</td><td>${esc(g.subject || "")}</td><td class="q">${g.cardsNow}</td><td>${what}</td><td class="arts">${g.items.map(i => esc(i.vc)).join(", ")}</td></tr>`;
    }).join("") + "</tbody>";
  const run = $("#btnRun");
  run.disabled = !p.ok || !p.todo || busy;
  run.title = !p.ok ? "Сначала исправьте ошибки в списке" : (!p.todo ? "Всё уже собрано как нужно" : "");
}

function renderJob() {
  const panel = $("#jobPanel");
  if (!job) { panel.style.display = "none"; return; }
  panel.style.display = "";
  $("#jobTitle").textContent = { plan: "Проверка списка", run: "Выполнение", rollback: "Откат" }[job.kind] || "";
  $("#jobStatus").innerHTML = job.running ? `<span class="spin"></span>${esc(job.message)}` :
    (job.error ? `<span style="color:var(--err)"><b>Ошибка:</b> ${esc(job.error)}</span>` : "");
  const el = $("#jobLog");
  const atEnd = el.scrollTop + el.clientHeight >= el.scrollHeight - 5;
  el.textContent = job.log.join("\n") || "…";
  if (atEnd) el.scrollTop = el.scrollHeight;
  const r = job.result;
  $("#jobResult").innerHTML = r ? `<p><b>${esc(r.summary)}</b></p><table class="plan"><tbody>` +
    r.results.map(x => `<tr><td>${esc(x.name)}</td><td><span class="tag ${x.ok ? "ok" : "error"}">${x.ok ? "готово" : "ошибка"}</span></td><td>${esc(x.msg)}</td></tr>`).join("") + "</tbody></table>" : "";
}

async function renderSnaps() {
  const s = shop(), el = $("#snaps");
  if (!s) { el.innerHTML = ""; return; }
  const list = (await S.listSnapshots(s.id)).slice(0, 30);
  if (!list.length) { el.innerHTML = `<div class="muted">Пока ничего не выполнялось.</div>`; return; }
  el.innerHTML = list.map(x => `<div class="snap"><div class="grow"><b>${esc(x.time)}</b> · ${esc(x.label || "")}<br>
    <span class="muted">${esc(x.summary || "не закончено")} · снимок: ${(x.before || []).reduce((a, g) => a + g.length, 0)} товаров в ${(x.before || []).length} карточках${x.rolledBack ? " · <b>откачено " + esc(x.rolledBack) + "</b>" : ""}</span></div>
    <button class="ghost" data-dl="${esc(x.key)}">Скачать</button><button data-rb="${esc(x.key)}">Откатить</button></div>`).join("");
  $$("#snaps [data-dl]").forEach(b => b.onclick = async () => {
    const x = await S.kvGet(b.dataset.dl);
    saveFile(new Blob([JSON.stringify(Object.assign({ app: "ostatki-wb", shop: s.name }, x), null, 1)], { type: "application/json" }),
      `Снимок ${s.name} ${x.time.replace(/[:/]/g, "-")}.json`);
  });
  $$("#snaps [data-rb]").forEach(b => b.onclick = async () => {
    const x = await S.kvGet(b.dataset.rb);
    if (!confirm(`Откатить магазин «${s.name}» к состоянию до ${x.time}?`)) return;
    await runGroups(C.snapshotGroups(x.before), `Откат к ${x.time}`, b.dataset.rb);
  });
}

async function rollbackFromFile(file) {
  const s = shop();
  let x;
  try { x = JSON.parse(await file.text()); } catch (e) { alert("Это не файл снимка."); return; }
  if (!x || !Array.isArray(x.before) || !x.before.every(g => Array.isArray(g) && g.every(n => Number.isInteger(n)))) { alert("Это не файл снимка «Остатки WB»."); return; }
  if (x.shop && x.shop !== s.name && !confirm(`Снимок сделан для магазина «${x.shop}», а выбран «${s.name}». Всё равно продолжить?`)) return;
  if (!confirm(`Откатить магазин «${s.name}» к состоянию из файла (${x.time || "без даты"})?`)) return;
  await runGroups(C.snapshotGroups(x.before), `Откат из файла (${x.time || "без даты"})`);
}

// ---------- общее ----------

function showTab(t) {
  $$("nav.tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === t));
  $("#tab-stocks").style.display = t === "stocks" ? "" : "none";
  $("#tab-cards").style.display = t === "cards" ? "" : "none";
  ui("tab", t);
}

async function renderAll() {
  const s = shop();
  $("#welcome").style.display = s ? "none" : "";
  $("#workspace").style.display = s ? "" : "none";
  $("#btnRefresh").style.display = s ? "" : "none";
  $("#btnShops").style.display = st.shops.length ? "" : "none";
  if (!s) { $("#upd").textContent = ""; return; }
  $("#shopBanner").innerHTML = `Магазин <b>«${esc(s.name)}»</b>${s.sellerName && s.sellerName !== s.name ? ` (${esc(s.sellerName)})` : ""} — изменения карточек сразу увидят покупатели.`;
  setStatus("", "");
  renderPlan(); renderJob();
  await renderStocks();
  await renderSnaps();
  try { $("#cardsText").value = ui(`text:${s.id}`) || ""; } catch (e) { /* ignore */ }
}

function wire() {
  $("#shopSelect").onchange = e => {
    if (e.target.value === "__add") { e.target.value = st.current || ""; openShops("add"); return; }
    selectShop(e.target.value);
  };
  $("#btnShops").onclick = () => openShops();
  $("#welcomeAdd").onclick = () => openShops("add");
  $("#closeShops").onclick = closeShops;
  $("#cancelShop").onclick = () => startEdit(null);
  $("#saveShop").onclick = saveShop;
  $("#btnRefresh").onclick = refresh;
  $$("nav.tabs button").forEach(b => b.onclick = () => showTab(b.dataset.tab));
  $("#category").onchange = () => renderStocks();
  $("#zero").onchange = () => renderReport(currentReport());
  $("#allTags").onclick = () => { $$(".chip").forEach(c => { selected.add(c.dataset.n); c.classList.add("on"); }); saveTags(); renderReport(currentReport()); };
  $("#noTags").onclick = () => { selected.clear(); $$(".chip").forEach(c => c.classList.remove("on")); saveTags(); renderReport(currentReport()); };
  $("#btnExcel").onclick = downloadExcel;
  $("#cardsText").oninput = () => { const s = shop(); if (s) ui(`text:${s.id}`, $("#cardsText").value); };
  $("#btnXlsx").onclick = () => $("#xlsxFile").click();
  $("#xlsxFile").onchange = async () => {
    const f = $("#xlsxFile").files[0];
    if (!f) return;
    xlsxData = await f.arrayBuffer();
    $("#xlsxName").innerHTML = `${esc(f.name)} <button class="ghost" id="xlsxClear">убрать</button>`;
    $("#xlsxClear").onclick = () => { xlsxData = null; $("#xlsxName").textContent = ""; };
    $("#xlsxFile").value = "";
    makePlanFromInput();
  };
  $("#btnPlan").onclick = () => {
    if (!xlsxData && !$("#cardsText").value.trim()) { alert("Вставьте список или загрузите Excel."); return; }
    makePlanFromInput();
  };
  $("#btnRun").onclick = () => {
    const s = shop();
    if (!plan || !confirm(`Пересобрать ${plan.todo} карточек в магазине «${s.name}»?\n\nПеред началом сохранится снимок для отката.`)) return;
    runGroups(planInput, "Пересборка");
  };
  $("#btnSnapFile").onclick = () => $("#snapFile").click();
  $("#snapFile").onchange = async () => { const f = $("#snapFile").files[0]; $("#snapFile").value = ""; if (f) await rollbackFromFile(f); };
}

wire();
showTab(ui("tab") || "stocks");
if (st.current && !shop()) st.current = st.shops[0] ? st.shops[0].id : null;
renderShopSelect();
selectShop(st.current).then(() => { if (!st.shops.length) openShops("add"); else if (!cache) refresh(); });

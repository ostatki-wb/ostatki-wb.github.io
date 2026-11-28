import * as W from "./wb.js";
import { find, index } from "./cards.js";
import { SOURCES, STOCK_TYPE } from "./report.js";

const same = (left, right) => left.length === right.length && left.every((id, i) => id === right[i]);

function findArticle(article, catalog) {
  const nmId = /^\d+$/.test(article) ? Number(article) : null;
  return find(article, catalog) || (nmId ? find(nmId, catalog) : null);
}

export function checkAdditions(text, mainId, draft, catalog) {
  const articles = text.trim().split(/[\r\n,;]+/).flatMap(line => {
    const article = line.trim();
    return findArticle(article, catalog) ? [article] : article.split(/\s+/).filter(Boolean);
  });
  const ids = new Set(draft);
  const missing = new Set();
  const errors = [];
  for (const article of articles) {
    const card = findArticle(article, catalog);
    if (!card) missing.add(article);
    else if (card.nmID === mainId) errors.push("Основной товар нельзя рекомендовать в его собственной карточке.");
    else ids.add(card.nmID);
  }
  if (missing.size) errors.push(`В каталоге магазина не найдены: ${[...missing].join(", ")}.`);
  if (ids.size > W.MAX_RECOMMENDATIONS) errors.push(`WB разрешает до ${W.MAX_RECOMMENDATIONS} рекомендаций на товар.`);
  return { ids: [...ids], errors: [...new Set(errors)], missing: missing.size > 0 };
}

export async function saveRecommendations(token, mainId, base, draft, { verifyDelays = [0, 1500, 3500], ...options } = {}) {
  const read = () => W.fetchRecommendations(token, mainId, options);
  const before = await read();
  if (!same(before, base)) {
    const retained = before.filter(id => !base.includes(id) || draft.includes(id));
    const added = draft.filter(id => !base.includes(id));
    return { actual: before, draft: [...new Set([...retained, ...added])], changedOnWB: true };
  }

  let writeError;
  try {
    await W.replaceRecommendations(token, mainId, draft, options);
  } catch (e) {
    if (e.status >= 400 && e.status < 500) throw e;
    writeError = e;
  }

  let actual, readError;
  for (const delay of verifyDelays) {
    if (delay) await W.sleep(delay);
    try {
      actual = await read();
      readError = null;
      if (same(actual, draft)) break;
    } catch (e) {
      readError = e;
    }
  }
  const confirmed = actual != null && same(actual, draft);
  const error = [writeError?.message, readError?.message].filter(Boolean).join(" ");
  return { actual, draft, confirmed, error };
}

export function createRecommendations({ shop, cache, isBusy, setBusy, esc }) {
  const el = id => document.getElementById(id);
  const mainInput = el("recommendationMain");
  const text = el("recommendationText");
  const list = el("recommendationList");
  const editor = el("recommendationEditor");
  const open = el("recommendationOpen");
  const check = el("recommendationCheck");
  const apply = el("recommendationApply");
  const source = () => shop()?.source || "all";
  const stocks = () => cache()?.stocks?.[source()];
  let current = null;
  let cachedCards, cachedIndex, freshIndex;

  function catalog() {
    if (cachedCards !== cache()?.cards || !cachedIndex) {
      cachedCards = cache()?.cards;
      cachedIndex = index(cachedCards || []);
      freshIndex = null;
    }
    return freshIndex || cachedIndex;
  }

  function message(id, value = "", error = false) {
    const node = el(id);
    node.textContent = value;
    node.classList.toggle("error", error);
  }

  const lookupMessage = (value, error) => message("recommendationLookupMessage", value, error);
  const editMessage = (value, error) => message("recommendationEditMessage", value, error);
  const dirty = () => current && !same(current.base, current.draft);

  function sync() {
    const busy = isBusy();
    el("recommendationLookupFields").disabled = busy;
    el("recommendationEditFields").disabled = busy;
    open.disabled = !mainInput.value.trim();
    check.disabled = !text.value.trim();
    apply.disabled = !dirty() || !!text.value.trim() || current.draft.length > W.MAX_RECOMMENDATIONS;
  }

  function stockFor(id) {
    const snapshot = stocks();
    if (Object.hasOwn(snapshot || {}, id)) return snapshot[id];
    if (snapshot && cachedIndex.byNm.has(id)) return 0;
    return current.stocks.get(id) ?? null;
  }

  function explainPending() {
    el("recommendationPending").textContent = text.value.trim()
      ? "Нажмите «Проверить», чтобы включить введённые артикулы в список."
      : (current?.draft.length > W.MAX_RECOMMENDATIONS ? `Оставьте до ${W.MAX_RECOMMENDATIONS} товаров перед применением.` : "");
  }

  function render() {
    editor.hidden = !current;
    sync();
    if (!current) return;
    const { main, base, draft } = current;
    const cards = catalog().byNm;
    el("recommendationProduct").textContent = [main.vendorCode, `WB ${main.nmID}`, main.title].filter(Boolean).join(" · ");
    const added = draft.filter(id => !base.includes(id)).length;
    const removed = base.filter(id => !draft.includes(id)).length;
    el("recommendationCount").textContent = dirty()
      ? `Было ${base.length} · станет ${draft.length} из ${W.MAX_RECOMMENDATIONS} · добавить ${added} · убрать ${removed}`
      : `${base.length} из ${W.MAX_RECOMMENDATIONS}`;
    const ids = [...new Set([...base, ...draft])];
    list.hidden = !ids.length;
    el("recommendationEmpty").hidden = !!ids.length;
    if (ids.every(id => stockFor(id) != null)) message("recommendationStockMessage");
    const snapshotTime = stocks() && cache()?.updated;
    const stockTime = snapshotTime ? ` · данные сайта от ${snapshotTime}` : "";
    el("recommendationStockSource").textContent = ids.length ? `${SOURCES[source()]}${stockTime}` : "";
    list.innerHTML = ids.map(id => {
      const card = cards.get(id);
      const article = card?.vendorCode || `WB ${id}`;
      const detail = [card?.vendorCode ? `WB ${id}` : "", card?.title].filter(Boolean).join(" · ");
      const count = stockFor(id);
      const pendingRemoval = !draft.includes(id);
      const label = pendingRemoval ? "Вернуть" : "Убрать";
      const change = pendingRemoval ? "Будет убран" : (!base.includes(id) ? "Будет добавлен" : "");
      return `<li class="recommendation-item">
        <div class="recommendation-article"><strong>${esc(article)}</strong>${detail ? `<div class="muted">${esc(detail)}</div>` : ""}${change ? `<span class="recommendation-change">${change}</span>` : ""}</div>
        <span class="recommendation-quantity${count === 0 ? " zero" : ""}">${count == null ? "Нет данных" : `${esc(count)} шт.`}</span>
        <button type="button" class="ghost" data-recommendation="${id}" aria-label="${label} ${esc(article)}">${label}</button>
      </li>`;
    }).join("");
    explainPending();
  }

  async function loadStocks() {
    const ids = [...new Set([...current.base, ...current.draft])].filter(id => stockFor(id) == null);
    const status = "recommendationStockMessage";
    message(status);
    if (!ids.length) return;
    message(status, "Загружаю остатки из WB…");
    try {
      const quantities = await W.fetchStocks(shop().token, STOCK_TYPE[source()], { nmIDs: ids, retries: 0 });
      for (const id of ids) current.stocks.set(id, quantities[id] ?? 0);
      message(status);
    } catch (e) {
      message(status, `Не удалось загрузить часть остатков. ${e.message}`, true);
    }
    render();
  }

  async function refreshCatalog(progress) {
    progress("Загружаю каталог магазина из WB…");
    freshIndex = index(await W.fetchCards(shop().token, { progress }));
    return freshIndex;
  }

  async function run(progress, action) {
    if (isBusy() || !shop()) return;
    setBusy(true);
    try { await action(); }
    catch (e) { progress(e.message, true); }
    finally { setBusy(false); }
  }

  el("recommendationLookup").onsubmit = event => {
    event.preventDefault();
    const article = mainInput.value.trim();
    if (!article || isBusy()) return;
    if ((dirty() || text.value.trim()) && !confirm("Открыть список из WB и сбросить несохранённые изменения?")) return;
    run(lookupMessage, async () => {
      const main = findArticle(article, catalog()) || findArticle(article, await refreshCatalog(lookupMessage));
      if (!main) throw new Error(`Артикул «${article}» не найден в каталоге этого магазина.`);
      lookupMessage("Читаю рекомендации из WB…");
      const base = await W.fetchRecommendations(shop().token, main.nmID, { progress: lookupMessage });
      current = { main, base, draft: [...base], stocks: new Map() };
      text.value = "";
      lookupMessage();
      editMessage();
      render();
      await loadStocks();
    });
  };

  el("recommendationEdit").onsubmit = event => {
    event.preventDefault();
    if (!current || !text.value.trim()) return;
    run(editMessage, async () => {
      const { main, draft } = current;
      let review = checkAdditions(text.value, main.nmID, draft, catalog());
      if (review.missing && !freshIndex) {
        review = checkAdditions(text.value, main.nmID, draft, await refreshCatalog(editMessage));
      }
      if (review.errors.length) {
        text.setAttribute("aria-invalid", "true");
        throw new Error(review.errors.join(" "));
      }
      current.draft = [...new Set([...current.base.filter(id => review.ids.includes(id)), ...review.ids])];
      text.value = "";
      text.removeAttribute("aria-invalid");
      editMessage(same(draft, current.draft) ? "Все указанные товары уже в списке." : "Проверьте итоговый список перед применением.");
      render();
      await loadStocks();
    });
  };

  list.onclick = event => {
    const button = event.target.closest("[data-recommendation]");
    if (!button || isBusy()) return;
    const id = Number(button.dataset.recommendation);
    const selected = new Set(current.draft);
    if (!selected.delete(id)) selected.add(id);
    const order = new Set([...current.base, ...current.draft]);
    current.draft = [...order].filter(item => selected.has(item));
    editMessage();
    render();
    (list.querySelector(`[data-recommendation="${id}"]`) || text).focus();
  };

  apply.onclick = () => {
    if (apply.disabled) return;
    run(editMessage, async () => {
      editMessage("Сверяю список с WB и применяю изменения…");
      const { main, base, draft } = current;
      const result = await saveRecommendations(shop().token, main.nmID, base, draft, { progress: editMessage });
      current.base = result.actual ?? base;
      current.draft = result.draft;
      if (result.changedOnWB) {
        editMessage(dirty()
          ? "Список в WB изменился. Проверьте обновлённый список перед применением."
          : "Список в WB уже совпадает с итоговым.");
      } else if (result.confirmed) {
        editMessage(result.error ? `Список в WB совпадает с итоговым. WB вернул сообщение об ошибке. ${result.error}` : "Изменения применены. Список проверен в WB.", !!result.error);
      } else {
        editMessage(`WB пока не подтвердил итоговый список. ${result.error || "Сверьте изменения и повторите применение."}`, true);
      }
      render();
      await loadStocks();
    });
  };

  mainInput.oninput = () => { lookupMessage(); sync(); };
  text.oninput = () => {
    text.removeAttribute("aria-invalid");
    editMessage();
    explainPending();
    sync();
  };

  return {
    sync, render,
    reset() {
      current = cachedCards = cachedIndex = freshIndex = null;
      mainInput.value = text.value = "";
      text.removeAttribute("aria-invalid");
      lookupMessage();
      editMessage();
      message("recommendationStockMessage");
      render();
    },
  };
}

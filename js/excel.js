// Excel в оформлении файла «Остатки 10.11» (перенос report.to_xlsx и cards.parse_xlsx).
// Библиотека ExcelJS передаётся параметром: в браузере — window.ExcelJS, в тестах — из npm.

const THIN = { style: "thin" };
const BORDER = { left: THIN, right: THIN, top: THIN, bottom: THIN };
const F_HEAD = { name: "Arial", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
const F_BOLD = { name: "Arial", size: 11, bold: true };
const F_REG = { name: "Arial", size: 11 };
const FILL_HEAD = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
const FILL_SUB = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };
const CENTER = { horizontal: "center" };
const ORDERS_HEADER = "Заказов в день";
const PHOTO_HEADER = "Фото";
const CARD_SHEET = "По карточкам";
const MAX_COLUMNS = 16384;
const PHOTO_PADDING = 4;
const EMU_PER_PIXEL = 9525;

function addPhoto(ws, imageId, photo, col, row) {
  const scale = Math.min(144 / photo.width, 172 / photo.height);
  const width = Math.max(1, Math.round(photo.width * scale));
  const height = Math.max(1, Math.round(photo.height * scale));
  const sheetRow = ws.getRow(row);
  sheetRow.height = Math.max(sheetRow.height || 0, (height + PHOTO_PADDING * 2) * 0.75);
  const offset = PHOTO_PADDING * EMU_PER_PIXEL;
  ws.addImage(imageId, {
    tl: { nativeCol: col - 1, nativeColOff: offset, nativeRow: row - 1, nativeRowOff: offset },
    ext: { width, height },
    editAs: "oneCell",
  });
}

export function colLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Отчёт по ярлыкам или объединённым карточкам. Каждый блок занимает соседние столбцы.
 *  Возвращает содержимое .xlsx (ArrayBuffer в браузере, Buffer в Node). */
export async function toXlsx(groups, ExcelJS, orderDates = null, {
  byCards = false, photos, withOrders = Boolean(orderDates || byCards),
} = {}) {
  const wb = new ExcelJS.Workbook();
  const imageIds = new Map();
  const sheetName = byCards ? CARD_SHEET : "Остатки";
  let ws = wb.addWorksheet(sheetName);
  let col = 1;
  for (const g of groups) {
    const cols = [["Артикул", "article", 14]].concat(g.named ? [["Название", "name", 30]] : [], [["Кол-во", "qty", 8]]);
    const qtyOffset = cols.length - 1;
    if (withOrders) cols.push([ORDERS_HEADER, "ordersPerDay", 19]);
    if (photos) cols.push([PHOTO_HEADER, "photo", 23]);
    if (col + cols.length - 1 > MAX_COLUMNS) {
      ws = wb.addWorksheet(`${sheetName} ${wb.worksheets.length + 1}`);
      col = 1;
    }
    const qtyCol = col + qtyOffset;
    const first = col, last = col + cols.length - 1;
    const a = colLetter(first);
    cols.forEach(([, , w], k) => { ws.getColumn(first + k).width = w; });
    if (last < MAX_COLUMNS) ws.getColumn(last + 1).width = 2;
    for (let c = first; c <= last; c++) {
      const cell = ws.getCell(1, c);
      cell.font = F_HEAD; cell.fill = FILL_HEAD; cell.alignment = CENTER; cell.border = BORDER;
    }
    ws.getCell(1, first).value = byCards ? g.name : g.tag;
    ws.mergeCells(1, first, 1, last);
    cols.forEach(([title, key], k) => {
      const cell = ws.getCell(2, first + k);
      cell.value = title; cell.font = F_BOLD; cell.fill = FILL_SUB; cell.alignment = CENTER; cell.border = BORDER;
      if (key === "ordersPerDay") cell.note = orderDates
        ? `Период заказов: ${orderDates}. Все модели продаж.`
        : "Данные о заказах недоступны.";
    });
    let r = 3;
    for (const row of g.rows) {
      cols.forEach(([, key], k) => {
        const cell = ws.getCell(r, first + k);
        const v = row[key];
        cell.value = v === "" || v === undefined ? null : v;
        if (key === "ordersPerDay") cell.numFmt = "0.00";
        cell.font = F_REG; cell.border = BORDER;
        if (photos) cell.alignment = { vertical: "middle" };
      });
      const photo = photos?.get(row.nmID);
      if (photo) {
        if (!imageIds.has(row.nmID)) imageIds.set(row.nmID, wb.addImage({ buffer: photo.buffer, extension: "jpeg" }));
        addPhoto(ws, imageIds.get(row.nmID), photo, last, r);
      }
      r++;
    }
    const lastRow = r - 1;
    const sku = g.rows.filter(row => row.article !== "" && row.article != null).length;
    for (let c = first; c <= last; c++) {
      const cell = ws.getCell(r, c);
      if (c === first) cell.value = "Итого SKU:";
      else if (c === qtyCol) cell.value = lastRow >= 3 ? { formula: `COUNTA(${a}3:${a}${lastRow})`, result: sku } : 0;
      cell.font = F_BOLD; cell.fill = FILL_SUB; cell.border = BORDER;
    }
    col = last + 2;
  }
  return wb.xlsx.writeBuffer();
}

// ---------- чтение списка карточек из Excel ----------

const SERVICE_WORDS = new Set(["артикул", "артикулы", "кол-во", "количество", "итого sku:", "итого sku", "итого", "итого:"]);
const tokens = s => String(s).trim().split(/[\s,;]+/u).filter(Boolean);

function cellValue(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "object") {
    if (v instanceof Date) return null;
    if (Array.isArray(v.richText)) return v.richText.map(p => p.text).join("");
    if ("formula" in v || "sharedFormula" in v) return null;   // формулы («=COUNTA…») пропускаем
    if ("text" in v) return String(v.text);
    if ("result" in v) return v.result;
    return null;
  }
  return v;
}

/** Excel: одна колонка — одна карточка. Сверху (необязательно) название, ниже артикулы.
 *  Подходит и файл отчёта об остатках: служебные ячейки, количества и «Название» пропускаются. */
export async function parseXlsx(data, known, ExcelJS) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(data);
  const first = wb.worksheets[0];
  if (!first) return [];
  if (first.name === CARD_SHEET) {
    return wb.worksheets.filter(ws => /^По карточкам(?: \d+)?$/.test(ws.name))
      .flatMap(ws => parseSheet(ws, known, true));
  }
  return parseSheet(first, known);
}

function parseSheet(ws, known, byCards = false) {
  const cols = new Map();
  ws.eachRow({ includeEmpty: false }, row => {
    row.eachCell({ includeEmpty: false }, cell => {
      if (cell.isMerged && cell.master && cell.master.address !== cell.address) return; // хвост объединённой ячейки
      const v = cellValue(cell);
      if (v === null || v === "") return;
      const c = cell.col;
      if (!cols.has(c)) cols.set(c, []);
      cols.get(c).push(v);
    });
  });
  const groups = [];
  for (const c of [...cols.keys()].sort((x, y) => x - y)) {
    const vals = cols.get(c);
    if (byCards && cellValue(ws.getCell(2, c)) !== "Артикул") continue;
    if (ws.name === "Остатки" && ["Кол-во", ORDERS_HEADER, PHOTO_HEADER].includes(cellValue(ws.getCell(2, c)))) continue;
    if (vals.some(v => typeof v === "string" && v.trim().toLowerCase() === "название")) continue;
    let name = "", items = [];
    for (const v of vals) {
      if (typeof v === "boolean") continue;
      if (typeof v === "number") {
        const s = Number.isInteger(v) ? String(v) : String(v);
        if (known.has(s)) items.push(s);
        continue;
      }
      const s = String(v).trim();
      if (!s || s.startsWith("=") || SERVICE_WORDS.has(s.toLowerCase())) continue;
      if (known.has(s)) items.push(s);
      else if (!name && !items.length) name = s;
      else items = items.concat(tokens(s));
    }
    if (items.length) groups.push({ name: name || `Карточка ${groups.length + 1}`, items });
  }
  return groups;
}

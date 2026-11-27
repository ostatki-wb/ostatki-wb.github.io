const MAX_SIDE = 344;
const TIMEOUT_MS = 12000;
const CONCURRENT_DOWNLOADS = 6;

async function loadPhoto(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), TIMEOUT_MS);
  let bitmap;
  try {
    if (new URL(url).protocol !== "https:") return null;
    const response = await fetch(url, { credentials: "omit", referrerPolicy: "no-referrer", signal: controller.signal });
    if (!response.ok || !response.headers.get("Content-Type")?.startsWith("image/")) return null;
    bitmap = await createImageBitmap(await response.blob());

    const scale = Math.min(1, MAX_SIDE / Math.max(bitmap.width, bitmap.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(bitmap.width * scale));
    canvas.height = Math.max(1, Math.round(bitmap.height * scale));
    const context = canvas.getContext("2d", { alpha: false });
    context.fillStyle = "#fff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    const jpeg = await new Promise(resolve => canvas.toBlob(resolve, "image/jpeg", 0.87));
    if (!jpeg || jpeg.type !== "image/jpeg") return null;
    return { buffer: new Uint8Array(await jpeg.arrayBuffer()), width: canvas.width, height: canvas.height };
  } catch {
    return null;
  } finally {
    clearTimeout(timer);
    bitmap?.close();
  }
}

/** Загружает фото с ограниченной параллельностью. В результате только доступные изображения. */
export async function loadPhotos(rows, progress = () => {}) {
  const photos = new Map();
  const queue = rows.filter(row => row.photoUrl);
  let next = 0, completed = rows.length - queue.length;
  progress(completed, rows.length);
  await Promise.all(Array.from({ length: Math.min(queue.length, CONCURRENT_DOWNLOADS) }, async () => {
    while (next < queue.length) {
      const row = queue[next++];
      const photo = await loadPhoto(row.photoUrl);
      if (photo) photos.set(row.nmID, photo);
      progress(++completed, rows.length);
    }
  }));
  return photos;
}

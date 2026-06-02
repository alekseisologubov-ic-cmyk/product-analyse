const cleanImageText = (value) =>
  String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

export const cleanSharedMasterImage = (value) => {
  const text = String(value || "").trim();

  if (!text) return "";

  // Do not save embedded/base64 images into Supabase.
  // They are too large and can cause statement timeout.
  if (text.startsWith("data:image/")) return "";

  // Also protect database from very large accidental values.
  if (text.length > 5000) return "";

  return text;
};

export const getImageMimeFromDataUrl = (dataUrl) => {
  const match = String(dataUrl || "").match(/^data:([^;]+);base64,/);
  return match?.[1] || "image/png";
};

export const getImageExtensionFromMime = (mime) => {
  const value = String(mime || "").toLowerCase();

  if (value.includes("jpeg") || value.includes("jpg")) return "jpg";
  if (value.includes("webp")) return "webp";
  if (value.includes("gif")) return "gif";

  return "png";
};

export const makeStorageSafePart = (value) => {
  const cleaned = String(value || "item")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 90);

  return cleaned || "item";
};

export const getEquipmentImageMatchCode = (value) => {
  const text = String(value || "")
    .trim()
    .replace(/\.0+$/g, "");

  const match = text.match(/\d{4,}/);

  if (match) {
    return match[0].replace(/^0+/, "");
  }

  return cleanImageText(text).replace(/[^A-Z0-9]/g, "");
};

export const normalizeEquipmentPictureCode = (value) => {
  const text = String(value || "")
    .trim()
    .replace(/\.0+$/g, "");

  const match = text.match(/\d{4,}/);
  return match ? match[0].replace(/^0+/, "") : "";
};

export const getZipImageMimeType = (fileName) => {
  const text = String(fileName || "").toLowerCase();

  if (text.endsWith(".jpg") || text.endsWith(".jpeg")) return "image/jpeg";
  if (text.endsWith(".webp")) return "image/webp";
  if (text.endsWith(".gif")) return "image/gif";

  return "image/png";
};

export const getZipImageExtension = (fileName) => {
  const text = String(fileName || "").toLowerCase();

  if (text.endsWith(".jpg") || text.endsWith(".jpeg")) return "jpg";
  if (text.endsWith(".webp")) return "webp";
  if (text.endsWith(".gif")) return "gif";

  return "png";
};

export const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

export const normalizeVenue = (value) =>
  cleanText(value)
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\s*-\s*VV$/g, "")
    .replace(/\s*VV$/g, "")
    .replace(/\bTHE\s+/g, "")
    .replace(/\bSCL\b/g, "")
    .replace(/\bVAL\b/g, "")
    .replace(/\bRES\b/g, "")
    .replace(/\bBRL\b/g, "")
    .replace(/\bROJO\b/g, "")
    .replace(/\bARIYA\b/g, "")
    .replace(/\bONLY\b/g, "")
    .replace(/\bMANNOR\b/g, "MANOR")
    .replace(/\s+/g, " ")
    .trim();

const PRODUCT_MATCH_STOP_WORDS = new Set([
  "FRESH",
  "BABY",
  "LARGE",
  "SMALL",
  "REGULAR",
  "HYDROPONIC",
  "OR",
  "AND",
  "THE",
  "FOR",
  "WITH",
  "WITHOUT",
  "LBS",
  "LB",
  "KG",
  "G",
  "OZ",
  "CS",
  "CASE",
  "BOX",
  "PC",
  "PCS",
  "PK",
  "PACK",
  "CT",
  "EA",
  "EACH",
]);

const singularizeProductToken = (token) => {
  if (!token) return "";

  if (token.length > 4 && token.endsWith("IES")) {
    return `${token.slice(0, -3)}Y`;
  }

  if (token.length > 4 && token.endsWith("ES") && !token.endsWith("SES")) {
    return token.slice(0, -2);
  }

  if (token.length > 3 && token.endsWith("S") && !token.endsWith("SS")) {
    return token.slice(0, -1);
  }

  return token;
};

export const getProductMatchTokens = (value) =>
  cleanText(value)
    .replace(/[^A-Z0-9]+/g, " ")
    .split(" ")
    .map((token) => singularizeProductToken(token.trim()))
    .filter((token) => token && token.length > 2)
    .filter((token) => !/^\d+$/.test(token))
    .filter((token) => !PRODUCT_MATCH_STOP_WORDS.has(token));

export const productNamesMatch = (left, right) => {
  const a = cleanText(left);
  const b = cleanText(right);

  if (!a || !b) return false;
  if (a === b) return true;

  if (a.length > 12 && (a.includes(b) || b.includes(a))) return true;
  if (b.length > 12 && (a.includes(b) || b.includes(a))) return true;

  const aTokens = getProductMatchTokens(a);
  const bTokens = getProductMatchTokens(b);

  if (!aTokens.length || !bTokens.length) return false;

  const shortTokens = aTokens.length <= bTokens.length ? aTokens : bTokens;

  const longTokenSet = new Set(
    aTokens.length <= bTokens.length ? bTokens : aTokens
  );

  const matchedCount = shortTokens.filter((token) =>
    longTokenSet.has(token)
  ).length;

  if (shortTokens.length === 1) {
    const token = shortTokens[0];
    return token.length >= 4 && matchedCount === 1;
  }

  return matchedCount >= Math.ceil(shortTokens.length * 0.75);
};

export const getProductReportKey = (value) => {
  const displayValue = String(value || "").trim();
  if (!displayValue) return "";

  const tokens = [...new Set(getProductMatchTokens(displayValue))].sort();

  return tokens.length ? tokens.join("|") : cleanText(displayValue);
};

export const formatQty = (value) => Number(value || 0).toFixed(2);

export const formatMoney = (value) => "$" + Number(value || 0).toFixed(2);

export const getImageUrl = (url, size = "w800") => {
  const value = String(url || "").trim();
  if (!value) return "";

  if (value.startsWith("data:image/")) {
    return value;
  }

  const googleDriveFileMatch = value.match(
    /drive\.google\.com\/file\/d\/([^/]+)/
  );

  const googleDriveIdMatch = value.match(/[?&]id=([^&]+)/);
  const googleDriveId = googleDriveFileMatch?.[1] || googleDriveIdMatch?.[1];

  if (googleDriveId) {
    return `https://drive.google.com/thumbnail?id=${googleDriveId}&sz=${size}`;
  }

  if (value.includes("sharepoint.com") || value.includes("1drv.ms")) {
    return value.includes("?") ? `${value}&download=1` : `${value}?download=1`;
  }

  return value;
};

export const isUsableImageValue = (value) => {
  const text = String(value || "").trim();

  if (!text) return false;
  if (text.startsWith("data:image/")) return true;
  if (/^https?:\/\//i.test(text)) return true;
  if (text.includes("drive.google.com")) return true;
  if (text.includes("sharepoint.com")) return true;
  if (text.includes("1drv.ms")) return true;

  return false;
};

export const getUsableImageValue = (...values) => {
  for (const value of values) {
    const text = String(value || "").trim();

    if (isUsableImageValue(text)) {
      return text;
    }
  }

  return "";
};

export const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

export const toNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : 0;
};

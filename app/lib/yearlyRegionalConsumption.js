import * as XLSX from "xlsx";

export const YEARLY_REGION_ALL = "ALL_REGIONS";

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const safeText = (value) => String(value || "").replace(/\s+/g, " ").trim();

const toNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
};

export const normalizeOrderCode = (value) => {
  const raw = String(value ?? "").trim();

  if (!raw) return "";

  const cleaned = raw
    .replace(/\.0+$/g, "")
    .replace(/[^A-Z0-9]/gi, "")
    .toUpperCase();

  if (!cleaned) return "";

  if (/^\d+$/.test(cleaned)) {
    return cleaned.replace(/^0+/, "") || "0";
  }

  return cleaned;
};

const PRODUCT_MATCH_STOP_WORDS = new Set([
  "FRESH",
  "FROZEN",
  "CHILLED",
  "BABY",
  "LARGE",
  "SMALL",
  "MEDIUM",
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
  "GR",
  "OZ",
  "CS",
  "CASE",
  "BOX",
  "BAG",
  "BAGS",
  "PC",
  "PCS",
  "PK",
  "PACK",
  "CT",
  "EA",
  "EACH",
  "UOM",
  "UM",
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

export const getProductReportKey = (value) => {
  const displayValue = String(value || "").trim();

  if (!displayValue) return "";

  const tokens = [...new Set(getProductMatchTokens(displayValue))].sort();

  return tokens.length ? tokens.join("|") : cleanText(displayValue);
};

export const normalizeYearlyShipCode = (value) => {
  const text = cleanText(value).replace(/RESILIANT/g, "RESILIENT");

  if (!text) return "";

  if (text === "BR" || text === "BRL" || text.includes("BRILLIANT")) {
    return "BRL";
  }

  if (text === "RL" || text.includes("RESILIENT")) {
    return "RL";
  }

  if (text === "SC" || text === "SCL" || text.includes("SCARLET")) {
    return "SC";
  }

  if (
    text === "VL" ||
    text === "V1" ||
    text === "VAL" ||
    text.includes("VALIANT")
  ) {
    return "VL";
  }

  return "";
};

export const normalizeRegionName = (value) => {
  const originalText = cleanText(value);

  if (!originalText) return "";

  if (
    originalText === "A" ||
    originalText === "NA" ||
    originalText === "N/A" ||
    originalText === "NONE" ||
    originalText.includes("TOTAL") ||
    originalText.includes("NOT IN OPERATION") ||
    originalText.includes("NOT OPERATING") ||
    originalText.includes("NO OPERATION") ||
    originalText.includes("NO SAILING")
  ) {
    return "";
  }

  const text = originalText
    .replace(/\bBRILLIANT LADY\b/g, "")
    .replace(/\bBRILLIANT\b/g, "")
    .replace(/\bBRL\b/g, "")
    .replace(/\bBR\b/g, "")
    .replace(/\bRESILIENT LADY\b/g, "")
    .replace(/\bRESILIENT\b/g, "")
    .replace(/\bRL\b/g, "")
    .replace(/\bSCARLET LADY\b/g, "")
    .replace(/\bSCARLET\b/g, "")
    .replace(/\bSCL\b/g, "")
    .replace(/\bSC\b/g, "")
    .replace(/\bVALIANT LADY\b/g, "")
    .replace(/\bVALIANT\b/g, "")
    .replace(/\bVAL\b/g, "")
    .replace(/\bVL\b/g, "")
    .replace(/\bV1\b/g, "")
    .replace(/\bHOME\s*PORT\b/g, "")
    .replace(/\bPORT\b/g, "")
    .replace(/\bPRICE\b/g, "")
    .replace(/\bQTY\b/g, "")
    .replace(/\bQUANTITY\b/g, "")
    .replace(/\bVALUE\b/g, "")
    .replace(/\bDAYS?\b/g, "")
    .replace(/\bDYAS\b/g, "")
    .replace(/\bDYAS?\b/g, "")
    .replace(/\bDYA\b/g, "")
    .replace(/\bDAYAS\b/g, "")
    .replace(/\b\d+\b/g, "")
    .replace(/[^A-Z ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!text) return "";

  if (
    text === "A" ||
    text === "NA" ||
    text === "N A" ||
    text === "NONE" ||
    text.includes("TOTAL") ||
    text.includes("NOT IN OPERATION") ||
    text.includes("NOT OPERATING")
  ) {
    return "";
  }

  if (text.includes("MIAMI")) return "MIAMI";
  if (text.includes("BARCELONA")) return "BARCELONA";
  if (text.includes("ATHENS")) return "ATHENS";
  if (text.includes("SAN JUAN")) return "SAN JUAN";
  if (text.includes("PORTSMOUTH")) return "PORTSMOUTH";
  if (text.includes("NEW YORK")) return "NEW YORK";
  if (text === "LA" || text.includes("LOS ANGELES")) return "LA";

  return text;
};

const MONTH_INFO = {
  MAY: { monthName: "May", monthKey: "2025-05", days: 31 },
  JUNE: { monthName: "June", monthKey: "2025-06", days: 30 },
  JULY: { monthName: "July", monthKey: "2025-07", days: 31 },
  AUGUST: { monthName: "August", monthKey: "2025-08", days: 31 },
  SEPTEMBER: { monthName: "September", monthKey: "2025-09", days: 30 },
  OCTOBER: { monthName: "October", monthKey: "2025-10", days: 31 },
  NOVEMBER: { monthName: "November", monthKey: "2025-11", days: 30 },
  DECEMBER: { monthName: "December", monthKey: "2025-12", days: 31 },
  JANUARY: { monthName: "January", monthKey: "2026-01", days: 31 },
  FEBRUARY: { monthName: "February", monthKey: "2026-02", days: 28 },
  MARCH: { monthName: "March", monthKey: "2026-03", days: 31 },
  APRIL: { monthName: "April", monthKey: "2026-04", days: 30 },
};

const getMonthInfoFromText = (value) => {
  const text = cleanText(value);

  if (!text) return null;

  const monthKey = Object.keys(MONTH_INFO).find((monthName) =>
    text.includes(monthName)
  );

  return monthKey ? MONTH_INFO[monthKey] : null;
};

const extractDaysOverride = (value) => {
  const text = cleanText(value);

  const match = text.match(/(\d+)\s*(DAY|DAYS|DYAS|DYA|DAYAS)\b/);

  if (!match) return 0;

  const days = Number(match[1]);

  return Number.isFinite(days) && days > 0 ? days : 0;
};

const extractShipPrefix = (value) => {
  const raw = safeText(value);

  if (!raw) return null;

  const match = raw.match(
    /^\s*(BRL|BR|BRILLIANT LADY|BRILLIANT|RL|RESILIENT LADY|RESILIENT|SC|SCL|SCARLET LADY|SCARLET|VL|V1|VAL|VALIANT LADY|VALIANT)\b/i
  );

  if (!match) return null;

  const ship = normalizeYearlyShipCode(match[1]);

  const rest = raw
    .slice(match[0].length)
    .replace(/^\s*[-:/]\s*/, "")
    .trim();

  return {
    ship,
    rest,
  };
};

const parseShipRegionHeader = ({ descriptor, fallbackShip, monthInfo }) => {
  const raw = safeText(descriptor);
  const fallbackRaw = safeText(fallbackShip);
  const combinedText = cleanText([raw, fallbackRaw].filter(Boolean).join(" "));

  if (!combinedText) return null;
  if (combinedText.includes("TOTAL")) return null;

  if (
    /\b(NOT|NO)\s*(IN\s*)?(OPERATION|OPERATING|SAILING)\b/.test(combinedText) ||
    /^(N\/?A|NA|A)$/.test(combinedText)
  ) {
    return null;
  }

  let ship = "";
  let regionText = "";

  const hyphenParts = raw
    .split(/\s*-\s*/g)
    .map((part) => part.trim())
    .filter(Boolean);

  if (hyphenParts.length > 1) {
    const shipFromLeft = normalizeYearlyShipCode(hyphenParts[0]);

    if (shipFromLeft) {
      ship = shipFromLeft;
      regionText = hyphenParts.slice(1).join(" - ");
    }
  }

  if (!ship) {
    const prefix = extractShipPrefix(raw);

    if (prefix?.ship) {
      ship = prefix.ship;
      regionText = prefix.rest;
    }
  }

  if (!ship) {
    ship = normalizeYearlyShipCode(fallbackRaw);
    regionText = raw;
  }

  if (!ship) {
    const fallbackPrefix = extractShipPrefix(fallbackRaw);

    if (fallbackPrefix?.ship) {
      ship = fallbackPrefix.ship;
      regionText = raw || fallbackPrefix.rest;
    }
  }

  if (!regionText) {
    const fallbackPrefix = extractShipPrefix(fallbackRaw);
    regionText = fallbackPrefix?.rest || raw || fallbackRaw;
  }

  const region = normalizeRegionName(regionText);

  if (!region) return null;

  const daysOverride = extractDaysOverride(
    [raw, fallbackRaw, regionText].filter(Boolean).join(" ")
  );

  const days = daysOverride || Number(monthInfo?.days || 0);

  if (!days) return null;

  return {
    ship: ship || "UNKNOWN",
    region,
    days,
    descriptor: raw || fallbackRaw,
  };
};

const fillMergedHeaderRows = (worksheet, rows) => {
  const filledRows = rows.map((row) => [...row]);
  const merges = worksheet?.["!merges"] || [];

  merges.forEach((merge) => {
    const sourceValue = filledRows[merge.s.r]?.[merge.s.c];

    if (sourceValue === undefined || sourceValue === null || sourceValue === "") {
      return;
    }

    for (let r = merge.s.r; r <= merge.e.r; r += 1) {
      if (!filledRows[r]) filledRows[r] = [];

      for (let c = merge.s.c; c <= merge.e.c; c += 1) {
        if (
          filledRows[r][c] === undefined ||
          filledRows[r][c] === null ||
          filledRows[r][c] === ""
        ) {
          filledRows[r][c] = sourceValue;
        }
      }
    }
  });

  return filledRows;
};

const findMetricRowIndex = (rows) => {
  let bestRowIndex = 2;
  let bestScore = 0;

  const maxRowsToSearch = Math.min(rows.length, 12);

  for (let rowIndex = 0; rowIndex < maxRowsToSearch; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    let score = 0;

    for (let colIndex = 0; colIndex < row.length - 2; colIndex += 1) {
      const metric = cleanText(row[colIndex]);
      const nextMetric = cleanText(row[colIndex + 1]);
      const valueMetric = cleanText(row[colIndex + 2]);

      if (
        metric === "PRICE" &&
        (nextMetric === "QTY" || nextMetric === "QUANTITY") &&
        valueMetric === "VALUE"
      ) {
        score += 1;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestRowIndex = rowIndex;
    }
  }

  return bestRowIndex;
};

const findMonthForBlock = ({ rows, metricRowIndex, colIndex, currentMonthInfo }) => {
  const searchValues = [];

  for (let r = 0; r < metricRowIndex; r += 1) {
    for (let c = Math.max(0, colIndex - 2); c <= colIndex + 2; c += 1) {
      searchValues.push(rows[r]?.[c]);
    }
  }

  const found = searchValues
    .map((value) => getMonthInfoFromText(value))
    .find(Boolean);

  return found || currentMonthInfo;
};

const getHeaderCandidatesForBlock = ({ rows, metricRowIndex, colIndex }) => {
  const candidates = [];

  // Important:
  // Each regional block is exactly PRICE / QTY / VALUE.
  // For the block starting at colIndex, only read the header cells directly
  // above those same 3 columns: colIndex, colIndex + 1, colIndex + 2.
  //
  // Do NOT read colIndex - 1 or colIndex - 2.
  // That can accidentally pull the previous ship/region header.
  //
  // Example:
  // BR - New York | RL - Athens | SC - Portsmouth | VL - Miami
  //
  // For RL - Athens, only the RL block's 3 columns should be checked.
  // It should not see BR - New York or SC - Portsmouth.
  for (let r = 0; r < metricRowIndex; r += 1) {
    for (let c = colIndex; c <= colIndex + 2; c += 1) {
      const text = safeText(rows[r]?.[c]);

      if (!text) continue;
      if (cleanText(text).includes("TOTAL")) continue;

      candidates.push(text);
    }
  }

  const unique = [];
  const seen = new Set();

  candidates.forEach((item) => {
    const key = cleanText(item);

    if (!key || seen.has(key)) return;
    seen.add(key);
    unique.push(item);
  });

  return unique;
};

const getParsedBlockFromCandidates = ({ candidates, monthInfo }) => {
  const cleanCandidates = (candidates || []).filter((item) => {
    const text = cleanText(item);

    if (!text) return false;
    if (getMonthInfoFromText(text)) return false;

    if (
      text === "PRICE" ||
      text === "QTY" ||
      text === "QUANTITY" ||
      text === "VALUE"
    ) {
      return false;
    }

    return true;
  });

  // First try each exact header candidate by itself.
  // This prevents one region from mixing with another region.
  for (let i = 0; i < cleanCandidates.length; i += 1) {
    const parsed = parseShipRegionHeader({
      descriptor: cleanCandidates[i],
      fallbackShip: "",
      monthInfo,
    });

    if (parsed) return parsed;
  }

  // Fallback only for unusual files where ship and region are split
  // inside the same 3-column PRICE/QTY/VALUE block.
  for (let i = 0; i < cleanCandidates.length; i += 1) {
    const parsed = parseShipRegionHeader({
      descriptor: cleanCandidates[i],
      fallbackShip: cleanCandidates[i + 1] || "",
      monthInfo,
    });

    if (parsed) return parsed;
  }

  return null;
};

const makeAggregateKey = ({ region, productCode, productKey }) =>
  `${region || ""}__${productCode || ""}__${productKey || ""}`;

const addAggregate = (map, payload) => {
  const key = makeAggregateKey(payload);

  if (!map.has(key)) {
    map.set(key, {
      region: payload.region,
      productCode: payload.productCode,
      productKey: payload.productKey,
      productName: payload.productName,
      unitMeasure: payload.unitMeasure,
      categoryName: payload.categoryName,
      subCategoryName: payload.subCategoryName,
      totalQty: 0,
      totalValue: 0,
      totalDays: 0,
      blocks: [],
      byShip: {},
    });
  }

  const row = map.get(key);

  row.totalQty += Number(payload.qty || 0);
  row.totalValue += Number(payload.value || 0);
  row.totalDays += Number(payload.days || 0);

  if (!row.byShip[payload.ship]) {
    row.byShip[payload.ship] = {
      ship: payload.ship,
      totalQty: 0,
      totalDays: 0,
      blocks: 0,
    };
  }

  row.byShip[payload.ship].totalQty += Number(payload.qty || 0);
  row.byShip[payload.ship].totalDays += Number(payload.days || 0);
  row.byShip[payload.ship].blocks += 1;

  row.blocks.push({
    monthKey: payload.monthKey,
    monthName: payload.monthName,
    ship: payload.ship,
    region: payload.region,
    descriptor: payload.descriptor,
    days: payload.days,
    qty: payload.qty,
    value: payload.value,
    price: payload.price,
  });
};

export const parseYearlyRegionalConsumptionWorkbook = (workbook) => {
  const sheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "EXPORT") ||
    workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    return {
      sourceSheet: sheetName || "",
      rows: [],
      aggregates: [],
      regionOptions: [],
      shipRegionBlocks: [],
      productCount: 0,
    };
  }

  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  const rows = fillMergedHeaderRows(worksheet, rawRows);

  if (rows.length < 4) {
    return {
      sourceSheet: sheetName,
      rows: [],
      aggregates: [],
      regionOptions: [],
      shipRegionBlocks: [],
      productCount: 0,
    };
  }

  const metricRowIndex = findMetricRowIndex(rows);
  const metricRow = rows[metricRowIndex] || [];
  const productRows = rows.slice(metricRowIndex + 1);

  let currentMonthInfo = null;
  const shipRegionBlocks = [];

  for (let colIndex = 8; colIndex < metricRow.length - 2; colIndex += 1) {
    const metric = cleanText(metricRow[colIndex]);
    const nextMetric = cleanText(metricRow[colIndex + 1]);
    const valueMetric = cleanText(metricRow[colIndex + 2]);

    if (
      metric !== "PRICE" ||
      (nextMetric !== "QTY" && nextMetric !== "QUANTITY") ||
      valueMetric !== "VALUE"
    ) {
      continue;
    }

    currentMonthInfo = findMonthForBlock({
      rows,
      metricRowIndex,
      colIndex,
      currentMonthInfo,
    });

    if (!currentMonthInfo) {
      continue;
    }

    const candidates = getHeaderCandidatesForBlock({
      rows,
      metricRowIndex,
      colIndex,
    });

    const parsed = getParsedBlockFromCandidates({
      candidates,
      monthInfo: currentMonthInfo,
    });

    if (!parsed) {
      colIndex += 2;
      continue;
    }

    shipRegionBlocks.push({
      ...parsed,
      monthName: currentMonthInfo.monthName,
      monthKey: currentMonthInfo.monthKey,
      priceColIndex: colIndex,
      qtyColIndex: colIndex + 1,
      valueColIndex: colIndex + 2,
      headerCandidates: candidates,
    });

    colIndex += 2;
  }

  const aggregateMap = new Map();
  const detailRows = [];

  productRows.forEach((row, rowIndex) => {
    const categoryCode = safeText(row[1]);
    const categoryName = safeText(row[2]);
    const subCategoryCode = safeText(row[3]);
    const subCategoryName = safeText(row[4]);
    const productCode = normalizeOrderCode(row[5]);
    const productName = safeText(row[6]);
    const unitMeasure = safeText(row[7]);

    if (!productCode && !productName) return;
    if (cleanText(productName) === "PRODUCT NAME") return;

    const productKey = getProductReportKey(productName);

    shipRegionBlocks.forEach((block) => {
      const qty = toNumber(row[block.qtyColIndex]);
      const value = toNumber(row[block.valueColIndex]);
      const price = toNumber(row[block.priceColIndex]);

      const detail = {
        sourceRow: rowIndex + metricRowIndex + 2,
        categoryCode,
        categoryName,
        subCategoryCode,
        subCategoryName,
        productCode,
        productName,
        productKey,
        unitMeasure,
        monthKey: block.monthKey,
        monthName: block.monthName,
        ship: block.ship,
        region: block.region,
        descriptor: block.descriptor,
        days: block.days,
        qty,
        value,
        price,
      };

      detailRows.push(detail);
      addAggregate(aggregateMap, detail);
    });
  });

  const aggregates = Array.from(aggregateMap.values()).map((item) => {
    const avgDailyQty =
      Number(item.totalDays || 0) > 0
        ? Number(item.totalQty || 0) / Number(item.totalDays || 0)
        : 0;

    return {
      ...item,
      avgDailyQty,
      byShip: Object.values(item.byShip || {}).map((shipRow) => ({
        ...shipRow,
        avgDailyQty:
          Number(shipRow.totalDays || 0) > 0
            ? Number(shipRow.totalQty || 0) / Number(shipRow.totalDays || 0)
            : 0,
      })),
    };
  });

  const regionOptions = Array.from(
    new Set(shipRegionBlocks.map((block) => block.region).filter(Boolean))
  ).sort();

  return {
    sourceSheet: sheetName,
    rows: detailRows,
    aggregates,
    regionOptions,
    shipRegionBlocks,
    productCount: productRows.length,
    metricRowIndex,
  };
};

const combineAggregateRows = (rows = []) => {
  const combined = {
    region: rows.length === 1 ? rows[0].region : YEARLY_REGION_ALL,
    productCode: rows[0]?.productCode || "",
    productKey: rows[0]?.productKey || "",
    productName: rows[0]?.productName || "",
    unitMeasure: rows[0]?.unitMeasure || "",
    categoryName: rows[0]?.categoryName || "",
    subCategoryName: rows[0]?.subCategoryName || "",
    totalQty: 0,
    totalValue: 0,
    totalDays: 0,
    avgDailyQty: 0,
    blocks: [],
    byShip: [],
  };

  const byShipMap = {};

  rows.forEach((row) => {
    combined.totalQty += Number(row.totalQty || 0);
    combined.totalValue += Number(row.totalValue || 0);
    combined.totalDays += Number(row.totalDays || 0);
    combined.blocks.push(...(row.blocks || []));

    (row.byShip || []).forEach((shipRow) => {
      if (!byShipMap[shipRow.ship]) {
        byShipMap[shipRow.ship] = {
          ship: shipRow.ship,
          totalQty: 0,
          totalDays: 0,
          blocks: 0,
          avgDailyQty: 0,
        };
      }

      byShipMap[shipRow.ship].totalQty += Number(shipRow.totalQty || 0);
      byShipMap[shipRow.ship].totalDays += Number(shipRow.totalDays || 0);
      byShipMap[shipRow.ship].blocks += Number(shipRow.blocks || 0);
    });
  });

  combined.avgDailyQty =
    Number(combined.totalDays || 0) > 0
      ? Number(combined.totalQty || 0) / Number(combined.totalDays || 0)
      : 0;

  combined.byShip = Object.values(byShipMap).map((shipRow) => ({
    ...shipRow,
    avgDailyQty:
      Number(shipRow.totalDays || 0) > 0
        ? Number(shipRow.totalQty || 0) / Number(shipRow.totalDays || 0)
        : 0,
  }));

  return combined;
};

export const findRegionalConsumptionForProduct = ({
  yearlyRegionalConsumption,
  productCode,
  productName,
  region,
}) => {
  if (!yearlyRegionalConsumption?.aggregates?.length) {
    return null;
  }

  const selectedRegion = String(region || "").trim();

  if (!selectedRegion) {
    return null;
  }

  const codeKey = normalizeOrderCode(productCode);
  const productKey = getProductReportKey(productName);

  let candidates = yearlyRegionalConsumption.aggregates;

  if (selectedRegion !== YEARLY_REGION_ALL) {
    candidates = candidates.filter((item) => item.region === selectedRegion);
  }

  let matches = [];

  if (codeKey) {
    matches = candidates.filter(
      (item) => normalizeOrderCode(item.productCode) === codeKey
    );
  }

  if (!matches.length && productKey) {
    matches = candidates.filter((item) => item.productKey === productKey);
  }

  if (!matches.length) {
    return null;
  }

  return combineAggregateRows(matches);
};

export const getRegionalParSuggestion = ({
  yearlyRegionalConsumption,
  productCode,
  productName,
  region,
  voyageDays,
  bufferPercent = 0,
}) => {
  const selectedRegion = String(region || "").trim();

  const match = findRegionalConsumptionForProduct({
    yearlyRegionalConsumption,
    productCode,
    productName,
    region: selectedRegion,
  });

  const days = Number(voyageDays || 0);
  const bufferMultiplier = 1 + Number(bufferPercent || 0) / 100;

  if (!match || !days) {
    return {
      hasRegionalData: false,
      region: selectedRegion,
      totalQty: 0,
      totalDays: 0,
      avgDailyQty: 0,
      suggestedParLevel: 0,
      evidenceBlocks: 0,
      matchedProductName: "",
      matchedProductCode: "",
      byShip: [],
    };
  }

  const suggestedParLevel =
    Number(match.avgDailyQty || 0) * days * bufferMultiplier;

  return {
    hasRegionalData: Number(match.totalDays || 0) > 0,
    region: selectedRegion,
    totalQty: Number(match.totalQty || 0),
    totalDays: Number(match.totalDays || 0),
    avgDailyQty: Number(match.avgDailyQty || 0),
    suggestedParLevel,
    evidenceBlocks: Array.isArray(match.blocks) ? match.blocks.length : 0,
    matchedProductName: match.productName || "",
    matchedProductCode: match.productCode || "",
    byShip: match.byShip || [],
  };
};

export const formatRegionalQty = (value) => Number(value || 0).toFixed(2);

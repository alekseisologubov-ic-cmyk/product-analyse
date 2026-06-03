"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

import {
  YEARLY_REGION_ALL,
  parseYearlyRegionalConsumptionWorkbook,
  formatRegionalQty,
} from "../../lib/yearlyRegionalConsumption";

const ORDER_BUFFER_PERCENT = 25;
const ORDER_BUFFER_MULTIPLIER = 1 + ORDER_BUFFER_PERCENT / 100;
const PAR_REPORT_DAYS = 7;
const PAR_REPORT_DISPLAY_LIMIT = 500;

// Excel column Y = current onboard ordered quantity.
// A = 0, B = 1, C = 2, ... Y = 24.
const ONBOARD_ORDERED_QTY_COLUMN_INDEX = 24;

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const safeText = (value) => String(value || "").replace(/\s+/g, " ").trim();

const normalizeCode = (value) => {
  const raw = String(value ?? "").trim();

  if (!raw) return "";

  const numberValue = Number(raw);

  if (
    Number.isFinite(numberValue) &&
    raw.replace(/\.0+$/, "") === String(Math.trunc(numberValue))
  ) {
    return String(Math.trunc(numberValue));
  }

  return cleanText(raw).replace(/\.0+$/, "");
};

const getExactProductNameKey = (value) =>
  cleanText(value)
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const normalizeShipCode = (value) => {
  const text = cleanText(value)
    .replace(/RESILIANT/g, "RESILIENT")
    .replace(/[_\-]+/g, " ");

  if (!text) return "";

  if (
    text === "V1" ||
    text === "V 1" ||
    text.includes(" V1 ") ||
    text.includes(" V 1 ") ||
    text === "SCL" ||
    text === "SC" ||
    text.includes("SCARLET")
  ) {
    return "SC";
  }

  if (text === "VL" || text === "VAL" || text.includes("VALIANT")) {
    return "VL";
  }

  if (text === "BRL" || text === "BR" || text.includes("BRILLIANT")) {
    return "BRL";
  }

  if (text === "RL" || text === "RES" || text.includes("RESILIENT")) {
    return "RL";
  }

  return "";
};

const getShipDisplayName = (shipCode) => {
  const names = {
    SC: "Scarlet",
    VL: "Valiant",
    BRL: "Brilliant",
    RL: "Resilient",
  };

  return names[shipCode] || shipCode || "Ship";
};

const toNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : 0;
};

const formatQty = (value) => Number(value || 0).toFixed(2);

const formatMoney = (value) => "$" + Number(value || 0).toFixed(2);

const excelDateToDate = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
  }

  const text = String(value || "").trim();

  if (!text) return null;

  const parsed = new Date(text);

  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
};

const formatDateCell = (value) => {
  const date = excelDateToDate(value);

  if (!date) return String(value || "").trim();

  return date.toLocaleDateString();
};

const getDaysBetweenCells = (startValue, endValue) => {
  const startDate = excelDateToDate(startValue);
  const endDate = excelDateToDate(endValue);

  if (!startDate || !endDate) return 0;

  const startUtc = Date.UTC(
    startDate.getUTCFullYear(),
    startDate.getUTCMonth(),
    startDate.getUTCDate()
  );

  const endUtc = Date.UTC(
    endDate.getUTCFullYear(),
    endDate.getUTCMonth(),
    endDate.getUTCDate()
  );

  const days = Math.round((endUtc - startUtc) / (24 * 60 * 60 * 1000));

  return Number.isFinite(days) && days > 0 ? days : 0;
};

const getHistoricalSailorDays = (cellA, cellB) => {
  const a = toNumber(cellA);
  const b = toNumber(cellB);

  if (!a && !b) return 0;
  if (a && !b) return a;
  if (!a && b) return b;

  const low = Math.min(Math.abs(a), Math.abs(b));
  const high = Math.max(Math.abs(a), Math.abs(b));

  if (low > 0 && high > low * 1000) return high;

  return a * b;
};

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

const getProductMatchTokens = (value) =>
  cleanText(value)
    .replace(/[^A-Z0-9]+/g, " ")
    .split(" ")
    .map((token) => singularizeProductToken(token.trim()))
    .filter((token) => token && token.length > 2)
    .filter((token) => !/^\d+$/.test(token))
    .filter((token) => !PRODUCT_MATCH_STOP_WORDS.has(token));

const productNamesMatch = (left, right) => {
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

const getProductReportKey = (value) => {
  const displayValue = String(value || "").trim();

  if (!displayValue) return "";

  const tokens = [...new Set(getProductMatchTokens(displayValue))].sort();

  return tokens.length ? tokens.join("|") : cleanText(displayValue);
};

const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

const escapeRegex = (value) =>
  String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const textHasWordOrPhrase = (text, word) => {
  const source = cleanText(text).replace(/[^A-Z0-9 ]/g, " ");
  const target = cleanText(word).replace(/[^A-Z0-9 ]/g, " ");

  if (!source || !target) return false;

  return new RegExp(
    `(^|[^A-Z0-9])${escapeRegex(target)}([^A-Z0-9]|$)`
  ).test(source);
};

const isFreshProduceOrderFile = (...values) => {
  const text = values
    .map((value) => cleanText(value))
    .join(" ")
    .replace(/[_\-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!text) return false;

  return text.includes("FRESH PRODUCE");
};

const FAST_SPOILAGE_PRODUCE_WORDS = [
  "STRAWBERRY",
  "STRAWBERRIES",
  "BLUEBERRY",
  "BLUEBERRIES",
  "RASPBERRY",
  "RASPBERRIES",
  "BLACKBERRY",
  "BLACKBERRIES",
  "BERRY",
  "BERRIES",
  "BANANA",
  "BANANAS",
  "LETTUCE",
  "ROMAINE",
  "ARUGULA",
  "ROCKET",
  "SPINACH",
  "MIXED GREENS",
  "SPRING MIX",
  "HERB",
  "HERBS",
  "BASIL",
  "MINT",
  "PARSLEY",
  "CILANTRO",
  "CORIANDER",
  "DILL",
  "CHIVE",
  "CHIVES",
  "TARRAGON",
  "CHERVIL",
  "WATERCRESS",
  "MICROGREEN",
  "MICRO GREENS",
  "MICROGREENS",
  "AVOCADO",
  "AVOCADOS",
];

const LONG_HOLD_PRODUCE_WORDS = [
  "POTATO",
  "POTATOES",
  "SWEET POTATO",
  "SWEET POTATOES",
  "YAM",
  "YAMS",
  "ONION",
  "ONIONS",
  "RED ONION",
  "WHITE ONION",
  "YELLOW ONION",
  "CARROT",
  "CARROTS",
  "BEET",
  "BEETS",
  "BEETROOT",
  "BEETROOTS",
  "TURNIP",
  "TURNIPS",
  "RUTABAGA",
  "PARSNIP",
  "PARSNIPS",
  "CELERIAC",
  "GARLIC",
  "SHALLOT",
  "SHALLOTS",
  "GINGER",
  "SQUASH",
  "BUTTERNUT",
  "BUTTERNUT SQUASH",
  "ACORN SQUASH",
  "KABOCHA",
  "PUMPKIN",
];

const getStandardProduceRule = () => ({
  type: "standard",
  label: "Standard item",
  orderFullTarget: false,
});

const getFreshProduceOrderRule = (
  productName,
  freshProduceOrderEnabled = false
) => {
  if (!freshProduceOrderEnabled) {
    return getStandardProduceRule();
  }

  const text = cleanText(productName);

  if (!text) {
    return getStandardProduceRule();
  }

  const isFastSpoilage = FAST_SPOILAGE_PRODUCE_WORDS.some((word) =>
    textHasWordOrPhrase(text, word)
  );

  if (isFastSpoilage) {
    return {
      type: "fast",
      label: "Quick-spoil fresh produce",
      orderFullTarget: true,
    };
  }

  const isLongHoldProduce = LONG_HOLD_PRODUCE_WORDS.some((word) =>
    textHasWordOrPhrase(text, word)
  );

  if (isLongHoldProduce) {
    return {
      type: "long",
      label: "Long-hold fresh produce",
      orderFullTarget: false,
    };
  }

  return {
    type: "fresh-standard",
    label: "Fresh produce - standard",
    orderFullTarget: false,
  };
};

const exportRowsToExcel = (rows, sheetName, fileName) => {
  if (!rows.length) {
    window.alert("No rows to export.");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  XLSX.writeFile(workbook, fileName);
};

const printSimpleTable = ({ title, subtitle, rows }) => {
  if (!rows.length) {
    window.alert("No rows to print.");
    return;
  }

  const columns = Object.keys(rows[0] || {});

  const html = `
    <html>
      <head>
        <title>${escapeHtml(title)}</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 24px; color: #111; }
          h1 { margin-bottom: 4px; }
          .subtitle { color: #555; margin-bottom: 18px; font-weight: bold; }
          table { width: 100%; border-collapse: collapse; font-size: 11px; }
          th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
          th { background: #f2f2f2; }
          tr { break-inside: avoid; }
        </style>
      </head>

      <body>
        <h1>${escapeHtml(title)}</h1>
        <div class="subtitle">${escapeHtml(subtitle || "")}</div>

        <table>
          <thead>
            <tr>
              ${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${rows
              .map(
                (row) => `
                  <tr>
                    ${columns
                      .map((column) => `<td>${escapeHtml(row[column])}</td>`)
                      .join("")}
                  </tr>
                `
              )
              .join("")}
          </tbody>
        </table>
      </body>
    </html>
  `;

  const printWindow = window.open("", "_blank");

  if (!printWindow) {
    window.alert("Print window was blocked. Allow popups and try again.");
    return;
  }

  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
};

const getStatusForOrderItem = (item) => {
  if (Number(item.averageConsumptionPerDay || 0) <= 0) {
    return {
      type: "review",
      label: "Review",
      description: "No current consumption/day could be calculated.",
    };
  }

  if (Number(item.suggestedOrder || 0) > 0) {
    return {
      type: "need",
      label: "Order Suggested",
      description:
        Number(item.preArrivalShortage || 0) > 0
          ? "Item is short before arrival. The pre-arrival shortage is highlighted but not added to the next order."
          : "Estimated quantity at arrival is below voyage need plus 25% buffer.",
    };
  }

  return {
    type: "covered",
    label: "Covered",
    description: "Future order / stock covers voyage need plus 25% buffer.",
  };
};

const parseFmlRows = (workbook) => {
  const fmlSheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "FML") ||
    workbook.SheetNames.find((name) => cleanText(name).includes("FML"));

  if (!fmlSheetName) return [];

  const worksheet = workbook.Sheets[fmlSheetName];

  if (!worksheet) return [];

  const decodedRange = XLSX.utils.decode_range(worksheet["!ref"] || "A1:I1");

  const fmlRange = {
    s: { r: decodedRange.s.r, c: 0 },
    e: { r: decodedRange.e.r, c: Math.max(decodedRange.e.c, 8) },
  };

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    range: fmlRange,
  });

  const seen = new Set();

  return rows
    .slice(3)
    .map((row, index) => {
      const excelRow = index + 4;

      const department = safeText(row[0]);
      const category = safeText(row[1]);
      const subCategory = safeText(row[2]);
      const code = safeText(row[3]);
      const product = safeText(row[4]);
      const venueText = safeText(row[5]);
      const uom = safeText(row[8]);

      return {
        excelRow,
        department,
        category,
        subCategory,
        code,
        product,
        venueText,
        venues: venueText
          .split(",")
          .map((venue) => safeText(venue))
          .filter(Boolean),
        uom,
      };
    })
    .filter((item) => {
      if (!item.code && !item.product) return false;
      if (cleanText(item.code) === "PRODUCT") return false;
      if (cleanText(item.product) === "PRODUCT NAME") return false;

      const key =
        normalizeCode(item.code) ||
        `${getProductReportKey(item.product)}|${cleanText(item.venueText)}`;

      if (!key) return false;
      if (seen.has(key)) return false;

      seen.add(key);
      return true;
    });
};

const findOrderItemForFmlItem = (fmlItem, orderRows, orderByCode) => {
  const codeKey = normalizeCode(fmlItem.code);

  if (codeKey && orderByCode.has(codeKey)) {
    return orderByCode.get(codeKey);
  }

  const productKey = getProductReportKey(fmlItem.product);

  if (!productKey) return null;

  return (
    orderRows.find((item) => {
      const orderProductKey = getProductReportKey(item.product);

      return (
        orderProductKey === productKey ||
        productNamesMatch(fmlItem.product, item.product)
      );
    }) || null
  );
};

const buildFmlReports = ({ fmlRows, orderRows }) => {
  const orderByCode = new Map();

  orderRows.forEach((item) => {
    const codeKey = normalizeCode(item.code);

    if (codeKey && !orderByCode.has(codeKey)) {
      orderByCode.set(codeKey, item);
    }
  });

  const missingRows = [];
  const runningLowRows = [];

  fmlRows.forEach((fmlItem) => {
    const orderItem = findOrderItemForFmlItem(fmlItem, orderRows, orderByCode);

    const futureOrders = Number(orderItem?.futureOrders || 0);
    const pastConsumption = Number(orderItem?.pastConsumption || 0);
    const averageConsumptionPerDay = Number(
      orderItem?.averageConsumptionPerDay || 0
    );
    const estimatedQtyAtArrival = Number(orderItem?.estimatedQtyAtArrival || 0);
    const targetQtyForVoyage = Number(orderItem?.targetQtyForVoyage || 0);
    const suggestedOrder = Number(orderItem?.suggestedOrder || 0);

    if (!orderItem || (futureOrders <= 0 && pastConsumption <= 0)) {
      missingRows.push({
        ...fmlItem,
        standardOrderRow: orderItem?.excelRow || "",
        stockOnHand: Number(orderItem?.stockOnHand || 0),
        futureOrders,
        pastConsumption,
        averageConsumptionPerDay,
        estimatedQtyAtArrival,
        targetQtyForVoyage,
        suggestedOrder,
        reason: !orderItem
          ? "FML item was not found in the order sheet."
          : "FML item has no future order and no current/past consumption in the order sheet.",
      });

      return;
    }

    if (
      orderItem &&
      futureOrders <= 0 &&
      pastConsumption > 0 &&
      averageConsumptionPerDay > 0 &&
      suggestedOrder > 0
    ) {
      runningLowRows.push({
        ...fmlItem,
        standardOrderRow: orderItem.excelRow,
        stockOnHand: Number(orderItem.stockOnHand || 0),
        futureOrders,
        pastConsumption,
        averageConsumptionPerDay,
        estimatedQtyAtArrival,
        targetQtyForVoyage,
        suggestedOrder,
        reason:
          "FML item has consumption and no future order. Estimated quantity at arrival is below voyage need plus 25% buffer.",
      });
    }
  });

  return {
    fmlMissingRows: missingRows.sort(
      (a, b) => Number(a.excelRow || 0) - Number(b.excelRow || 0)
    ),
    fmlRunningLowRows: runningLowRows.sort(
      (a, b) => Number(a.excelRow || 0) - Number(b.excelRow || 0)
    ),
  };
};

const parseNextOrderWorkbook = (workbook, sourceFileName = "") => {
  const sheetName = workbook.SheetNames.includes("Standard Order Template")
    ? "Standard Order Template"
    : workbook.SheetNames.includes("Order Sheet")
      ? "Order Sheet"
      : workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    throw new Error("Could not find a usable order sheet.");
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  const orderShipName = safeText(rows[0]?.[1]);
  const orderShipCode = normalizeShipCode(orderShipName);

  const isFreshProduceOrder = isFreshProduceOrderFile(
    sourceFileName,
    sheetName,
    orderShipName
  );

  const rawOrderDate = rows[1]?.[1];
  const rawArrivalDate = rows[2]?.[1];
  const targetSailors = toNumber(rows[4]?.[1]);
  const voyageDays = toNumber(rows[5]?.[1]);
  const daysUntilArrival = getDaysBetweenCells(rawOrderDate, rawArrivalDate);

  const futureOrderColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13]; // F:N
  const pastConsumptionColumns = [34, 35, 36, 37, 38, 39]; // AI:AN

  const historicalSailorDays = pastConsumptionColumns.reduce(
    (sum, colIndex) =>
      sum + getHistoricalSailorDays(rows[4]?.[colIndex], rows[5]?.[colIndex]),
    0
  );

  const parsedRows = [];
  const seen = new Set();

  rows.slice(9).forEach((row, rowOffset) => {
    const excelRow = rowOffset + 10;

    const code = safeText(row[0]);
    const product = safeText(row[1]);
    const uom = safeText(row[2]);

    if (!product || !uom) return;
    if (cleanText(product) === "PRODUCT NAME") return;

    const stockOnHand = toNumber(row[3]);
    const parLevel = toNumber(row[16]);

    // Column Y only.
    // This is used for Order vs Par report comparison.
    // Do not replace this with F:N future orders.
    const onboardOrderedQty = toNumber(row[ONBOARD_ORDERED_QTY_COLUMN_INDEX]);

    // F:N future orders remain used for estimated arrival coverage in main logic.
    const futureOrders = futureOrderColumns.reduce(
      (sum, colIndex) => sum + toNumber(row[colIndex]),
      0
    );

    const pastConsumption = pastConsumptionColumns.reduce(
      (sum, colIndex) => sum + toNumber(row[colIndex]),
      0
    );

    const averageConsumptionPerSailorDay =
      historicalSailorDays > 0 ? pastConsumption / historicalSailorDays : 0;

    const averageConsumptionPerDay =
      averageConsumptionPerSailorDay > 0 && targetSailors > 0
        ? averageConsumptionPerSailorDay * targetSailors
        : voyageDays > 0
          ? pastConsumption / voyageDays
          : 0;

    const consumptionUntilArrival = averageConsumptionPerDay * daysUntilArrival;

    const estimatedQtyAtArrival =
      stockOnHand + futureOrders - consumptionUntilArrival;

    const preArrivalShortage = Math.max(-estimatedQtyAtArrival, 0);

    // If arrival estimate is negative, we highlight the shortage,
    // but we do NOT add the negative amount to the next order.
    const usableQtyAtArrivalForOrder = Math.max(estimatedQtyAtArrival, 0);

    const voyageNeed = averageConsumptionPerDay * voyageDays;
    const orderBufferQty = voyageNeed * (ORDER_BUFFER_PERCENT / 100);
    const targetQtyForVoyage = voyageNeed + orderBufferQty;

    const produceRule = getFreshProduceOrderRule(product, isFreshProduceOrder);

    const normalSuggestedOrder = Math.max(
      targetQtyForVoyage - usableQtyAtArrivalForOrder,
      0
    );

    const suggestedOrder =
      produceRule.orderFullTarget && targetQtyForVoyage > 0
        ? targetQtyForVoyage
        : normalSuggestedOrder;

    const currentUsageSuggestedPar = targetQtyForVoyage;

    // This is for the Order vs Par report only.
    // It does not change the main order calculation.
    const currentUsageSuggestedPar7Days =
      averageConsumptionPerDay * PAR_REPORT_DAYS * ORDER_BUFFER_MULTIPLIER;

    const futureCoverageDifference =
      usableQtyAtArrivalForOrder - targetQtyForVoyage;

    const itemKey =
      normalizeCode(code) ||
      `${getProductReportKey(product)}|${cleanText(uom)}|${excelRow}`;

    if (seen.has(itemKey)) return;
    seen.add(itemKey);

    const status = getStatusForOrderItem({
      averageConsumptionPerDay,
      suggestedOrder,
      preArrivalShortage,
    });

    parsedRows.push({
      itemKey,
      excelRow,
      code,
      product,
      uom,
      stockOnHand,
      parLevel,
      onboardOrderedQty,
      futureOrders,
      pastConsumption,
      historicalSailorDays,
      targetSailors,
      voyageDays,
      daysUntilArrival,
      averageConsumptionPerSailorDay,
      averageConsumptionPerDay,
      consumptionUntilArrival,
      estimatedQtyAtArrival,
      preArrivalShortage,
      usableQtyAtArrivalForOrder,
      voyageNeed,
      orderBufferQty,
      targetQtyForVoyage,
      suggestedOrder,
      currentUsageSuggestedPar,
      currentUsageSuggestedPar7Days,
      futureCoverageDifference,
      isFreshProduceOrder,
      produceRuleType: produceRule.type,
      produceRuleLabel: produceRule.label,
      produceOrderFullTarget: produceRule.orderFullTarget,
      statusType: status.type,
      statusLabel: status.label,
      statusDescription: status.description,
    });
  });

  const fmlRows = parseFmlRows(workbook);

  const { fmlMissingRows, fmlRunningLowRows } = buildFmlReports({
    fmlRows,
    orderRows: parsedRows,
  });

  return {
    rows: parsedRows.sort(
      (a, b) => Number(a.excelRow || 0) - Number(b.excelRow || 0)
    ),
    fmlRows,
    fmlMissingRows,
    fmlRunningLowRows,
    meta: {
      sheetName,
      shipName: orderShipName,
      shipCode: orderShipCode,
      shipDisplayName: getShipDisplayName(orderShipCode),
      isFreshProduceOrder,
      orderDate: formatDateCell(rawOrderDate),
      arrivalDate: formatDateCell(rawArrivalDate),
      targetSailors,
      voyageDays,
      daysUntilArrival,
      historicalSailorDays,
      totalItems: parsedRows.length,
      itemsNeedingOrder: parsedRows.filter((item) => item.suggestedOrder > 0)
        .length,
      coveredItems: parsedRows.filter(
        (item) => item.suggestedOrder <= 0 && item.averageConsumptionPerDay > 0
      ).length,
      reviewItems: parsedRows.filter(
        (item) => item.averageConsumptionPerDay <= 0
      ).length,
      preArrivalShortageItems: parsedRows.filter(
        (item) => item.preArrivalShortage > 0
      ).length,
      fastSpoilageItems: parsedRows.filter(
        (item) => item.produceRuleType === "fast"
      ).length,
      longHoldProduceItems: parsedRows.filter(
        (item) => item.produceRuleType === "long"
      ).length,
      fmlRows: fmlRows.length,
      fmlMissingItems: fmlMissingRows.length,
      fmlRunningLowItems: fmlRunningLowRows.length,
    },
  };
};

const getIngredientHeaderIndexes = (rows = []) => {
  const compactName = (value) => cleanText(value).replace(/[^A-Z0-9]/g, "");

  const headerRowIndex = rows.findIndex((row) => {
    const compactText = (row || []).map(compactName).join("|");

    return (
      compactText.includes("RESTAURANTNAME") &&
      compactText.includes("RECIPENAME")
    );
  });

  const finalHeaderRowIndex = headerRowIndex >= 0 ? headerRowIndex : 0;
  const headers = rows[finalHeaderRowIndex] || [];
  const compactHeaders = headers.map(compactName);

  const findExact = (names, fallback) => {
    const wanted = names.map(compactName).filter(Boolean);
    const index = compactHeaders.findIndex((header) => wanted.includes(header));

    return index >= 0 ? index : fallback;
  };

  return {
    headerRowIndex: finalHeaderRowIndex,
    venue: findExact(["RestaurantName", "Restaurant Name"], 1),
    menuName: findExact(["MenuName", "Menu Name"], 3),
    ingredientCode: findExact(
      [
        "Code",
        "IngredientCode",
        "Ingredient Code",
        "ProductCode",
        "Product Code",
      ],
      6
    ),
    ingredientName: findExact(
      [
        "Name",
        "IngredientName",
        "Ingredient Name",
        "ProductName",
        "Product Name",
      ],
      7
    ),
    assigned: findExact(["Assigned"], 12),
    recipeCode: findExact(["RecipeCode", "Recipe Code"], 15),
    recipeName: findExact(["RecipeName", "Recipe Name"], 16),
  };
};

const getRecipeUsageForItem = (item, recipeRows = []) => {
  const productName = safeText(item?.product || item?.name);
  const productCode = normalizeCode(item?.code);

  if (!productName && !productCode) return [];

  const sourceRows = Array.isArray(recipeRows) ? recipeRows : [];

  if (!sourceRows.length) return [];

  const indexes = getIngredientHeaderIndexes(sourceRows);
  const rows = sourceRows.slice(indexes.headerRowIndex + 1);
  const usageRows = [];

  rows.forEach((row) => {
    const venue = safeText(row[indexes.venue]);
    const ingredientCode = normalizeCode(row[indexes.ingredientCode]);
    const ingredientName = safeText(row[indexes.ingredientName]);
    const assignedProduct = safeText(row[indexes.assigned]);
    const recipeCode = normalizeCode(row[indexes.recipeCode]);
    const recipeName = safeText(row[indexes.recipeName]);
    const menuName = safeText(row[indexes.menuName]);

    if (!venue || !recipeName) return;

    const matchedByCode =
      productCode &&
      [
        ingredientCode,
        normalizeCode(ingredientName),
        normalizeCode(assignedProduct),
      ].includes(productCode);

    const matchedByName =
      productName &&
      (productNamesMatch(productName, ingredientName) ||
        productNamesMatch(productName, assignedProduct));

    if (!matchedByCode && !matchedByName) return;

    usageRows.push({
      venue,
      menuName,
      recipeCode,
      recipeName,
      ingredientCode,
      ingredientName,
      assignedProduct,
    });
  });

  const seen = new Set();

  return usageRows.filter((row) => {
    const key = [
      row.venue,
      row.menuName,
      row.recipeCode,
      row.recipeName,
      row.ingredientCode,
      row.ingredientName,
      row.assignedProduct,
    ].join("|");

    if (seen.has(key)) return false;

    seen.add(key);
    return true;
  });
};

const buildYearlyRegionalLookup = (yearlyRegionalConsumption) => {
  const sourceRows = Array.isArray(yearlyRegionalConsumption?.rows)
    ? yearlyRegionalConsumption.rows
    : [];

  const byCode = new Map();
  const byExactName = new Map();

  const addToMap = (map, key, row) => {
    const cleanKey = String(key || "").trim();

    if (!cleanKey) return;

    if (!map.has(cleanKey)) {
      map.set(cleanKey, []);
    }

    map.get(cleanKey).push(row);
  };

  sourceRows.forEach((row, index) => {
    const productCode = normalizeCode(row.productCode);
    const productName = safeText(row.productName);
    const exactNameKey = getExactProductNameKey(productName);

    const preparedRow = {
      ...row,
      __lookupId: index,
      __region: safeText(row.region),
      __ship: normalizeShipCode(row.ship) || safeText(row.ship),
      __productCode: productCode,
      __productName: productName,
      __exactNameKey: exactNameKey,
    };

    addToMap(byCode, productCode, preparedRow);
    addToMap(byExactName, exactNameKey, preparedRow);
  });

  return {
    rows: sourceRows,
    byCode,
    byExactName,
  };
};

const getRegionalCandidateRowsForItem = (item, regionalLookup) => {
  if (!regionalLookup) return [];

  const itemCode = normalizeCode(item?.code);
  const itemExactNameKey = getExactProductNameKey(item?.product || item?.name);

  // Important:
  // If item has a code, last-year comparison matches by code only.
  // No loose name matching is used.
  if (itemCode) {
    return regionalLookup.byCode.get(itemCode) || [];
  }

  // If item has no code, match only exact normalized product description.
  if (itemExactNameKey) {
    return regionalLookup.byExactName.get(itemExactNameKey) || [];
  }

  return [];
};

const getRegionalStatsForItem = ({
  item,
  yearlyRegionalConsumption,
  regionalLookup,
  regionFilter,
  shipFilter,
  parDays,
  bufferPercent = ORDER_BUFFER_PERCENT,
}) => {
  const usingLookup = Boolean(regionalLookup);

  const sourceRows = usingLookup
    ? getRegionalCandidateRowsForItem(item, regionalLookup)
    : Array.isArray(yearlyRegionalConsumption?.rows)
      ? yearlyRegionalConsumption.rows
      : [];

  const itemCode = normalizeCode(item?.code);
  const itemExactNameKey = getExactProductNameKey(item?.product || item?.name);

  if (!sourceRows.length || !item) {
    return {
      matched: false,
      totalQty: 0,
      totalValue: 0,
      totalDays: 0,
      avgDailyQty: 0,
      suggestedPar: 0,
      blocks: 0,
      matchedProducts: [],
      parDays,
    };
  }

  const wantedRegion = safeText(regionFilter);
  const wantedShip = normalizeShipCode(shipFilter) || safeText(shipFilter);

  const matchedRows = sourceRows.filter((row) => {
    const rowRegion = row.__region ?? safeText(row.region);
    const rowShip =
      row.__ship || normalizeShipCode(row.ship) || safeText(row.ship);

    if (
      wantedRegion &&
      wantedRegion !== YEARLY_REGION_ALL &&
      rowRegion !== wantedRegion
    ) {
      return false;
    }

    if (wantedShip && rowShip !== wantedShip) {
      return false;
    }

    if (usingLookup) {
      return true;
    }

    const rowCode = normalizeCode(row.productCode);
    const rowExactNameKey = getExactProductNameKey(row.productName);

    if (itemCode) {
      return rowCode && itemCode === rowCode;
    }

    return itemExactNameKey && rowExactNameKey === itemExactNameKey;
  });

  const totalQty = matchedRows.reduce(
    (sum, row) => sum + Number(row.qty || 0),
    0
  );

  const totalValue = matchedRows.reduce(
    (sum, row) => sum + Number(row.value || 0),
    0
  );

  const totalDays = matchedRows.reduce(
    (sum, row) => sum + Number(row.days || 0),
    0
  );

  const blocks = matchedRows.length;
  const avgDailyQty = totalDays > 0 ? totalQty / totalDays : 0;

  const suggestedPar =
    avgDailyQty *
    Number(parDays || 0) *
    (1 + Number(bufferPercent || 0) / 100);

  const matchedProducts = [
    ...new Set(
      matchedRows
        .map((row) =>
          `${safeText(row.productCode) || "No code"} - ${safeText(
            row.productName
          )}`
        )
        .filter(Boolean)
    ),
  ].slice(0, 8);

  return {
    matched: matchedRows.length > 0,
    totalQty,
    totalValue,
    totalDays,
    avgDailyQty,
    suggestedPar,
    blocks,
    matchedProducts,
    parDays,
  };
};

export default function GenerateNextOrder({
  styles,
  userShip,
  onBack,
  logUsageEvent = () => {},
  yearlyRegionalConsumption,
  setYearlyRegionalConsumption,
  yearlyRegionalFileName,
  setYearlyRegionalFileName,
  selectedRegionalConsumptionRegion,
  setSelectedRegionalConsumptionRegion,
  regionalParBufferPercent,
  setRegionalParBufferPercent,
  recipeRows = [],
}) {
  const [orderRows, setOrderRows] = useState([]);
  const [orderMeta, setOrderMeta] = useState({});
  const [orderFileName, setOrderFileName] = useState("");
  const [fmlRows, setFmlRows] = useState([]);
  const [fmlMissingRows, setFmlMissingRows] = useState([]);
  const [fmlRunningLowRows, setFmlRunningLowRows] = useState([]);

  const [message, setMessage] = useState("");
  const [loading, setLoading] = useState(false);
  const [view, setView] = useState("order");
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all");
  const [fmlSearch, setFmlSearch] = useState("");

  const [selectedInfoItem, setSelectedInfoItem] = useState(null);
  const [selectedRecipeUsageItem, setSelectedRecipeUsageItem] = useState(null);

  const activeShipCode =
    orderMeta.shipCode || normalizeShipCode(userShip) || userShip || "";

  const activeShipName =
    orderMeta.shipDisplayName || getShipDisplayName(activeShipCode);

  const regionalRegionOptions = useMemo(
    () => yearlyRegionalConsumption?.regionOptions || [],
    [yearlyRegionalConsumption]
  );

  const yearlyRegionalLookup = useMemo(
    () => buildYearlyRegionalLookup(yearlyRegionalConsumption),
    [yearlyRegionalConsumption]
  );

  const uploadOrderFile = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    setLoading(true);
    setMessage("Reading order file...");

    try {
      await new Promise((resolve) => window.setTimeout(resolve, 20));

      const arrayBuffer = await file.arrayBuffer();

      setMessage("Parsing order workbook...");

      await new Promise((resolve) => window.setTimeout(resolve, 20));

      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsed = parseNextOrderWorkbook(workbook, file.name);

      setOrderRows(parsed.rows);
      setOrderMeta(parsed.meta);
      setOrderFileName(file.name);
      setFmlRows(parsed.fmlRows);
      setFmlMissingRows(parsed.fmlMissingRows);
      setFmlRunningLowRows(parsed.fmlRunningLowRows);
      setView("order");
      setSearch("");
      setFmlSearch("");
      setFilter("all");

      setMessage(
        "Order file loaded. " +
          parsed.meta.totalItems +
          " item(s), " +
          parsed.meta.itemsNeedingOrder +
          " suggested order item(s), " +
          parsed.meta.coveredItems +
          " covered item(s), " +
          parsed.meta.preArrivalShortageItems +
          " item(s) short before arrival, " +
          parsed.meta.fmlMissingItems +
          " FML not ordered/not used item(s), " +
          parsed.meta.fmlRunningLowItems +
          " FML running low item(s)."
      );

      logUsageEvent("next_order_file_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        sheetName: parsed.meta.sheetName,
        shipName: parsed.meta.shipName,
        shipCode: parsed.meta.shipCode,
        isFreshProduceOrder: parsed.meta.isFreshProduceOrder,
        totalItems: parsed.meta.totalItems,
        itemsNeedingOrder: parsed.meta.itemsNeedingOrder,
        fmlMissingItems: parsed.meta.fmlMissingItems,
        fmlRunningLowItems: parsed.meta.fmlRunningLowItems,
      });
    } catch (error) {
      setOrderRows([]);
      setOrderMeta({});
      setOrderFileName(file.name);
      setFmlRows([]);
      setFmlMissingRows([]);
      setFmlRunningLowRows([]);

      const text = error?.message || "Could not read the order file.";

      setMessage(text);
      window.alert(text);
    } finally {
      setLoading(false);
      event.target.value = "";
    }
  };

  const uploadYearlyRegionalFile = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    try {
      setMessage("Loading yearly regional consumption file...");

      const arrayBuffer = await file.arrayBuffer();

      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsed = parseYearlyRegionalConsumptionWorkbook(workbook);

      setYearlyRegionalConsumption?.(parsed);
      setYearlyRegionalFileName?.(file.name);
      setSelectedRegionalConsumptionRegion?.("");

      setMessage(
        "Yearly regional file loaded. " +
          parsed.aggregates.length +
          " regional product records found across " +
          parsed.regionOptions.length +
          " region(s)."
      );

      logUsageEvent("yearly_regional_consumption_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        sourceSheet: parsed.sourceSheet,
        regions: parsed.regionOptions || [],
        aggregates: parsed.aggregates.length,
      });
    } catch (error) {
      const text =
        error?.message || "Could not load yearly regional consumption file.";

      setMessage(text);
      window.alert(text);
    } finally {
      event.target.value = "";
    }
  };

  const visibleOrderRows = useMemo(() => {
    const query = search.toLowerCase().trim();

    return orderRows.filter((item) => {
      const matchesFilter =
        filter === "all" ||
        (filter === "needsOrder" && Number(item.suggestedOrder || 0) > 0) ||
        (filter === "covered" &&
          Number(item.suggestedOrder || 0) <= 0 &&
          Number(item.averageConsumptionPerDay || 0) > 0) ||
        (filter === "review" &&
          Number(item.averageConsumptionPerDay || 0) <= 0) ||
        (filter === "shortBeforeArrival" &&
          Number(item.preArrivalShortage || 0) > 0) ||
        (filter === "fastSpoilage" && item.produceRuleType === "fast") ||
        (filter === "longHoldProduce" && item.produceRuleType === "long");

      if (!matchesFilter) return false;

      if (!query) return true;

      return [
        item.code,
        item.product,
        item.uom,
        item.statusLabel,
        item.produceRuleLabel,
        item.excelRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [orderRows, search, filter]);

  const visibleFmlMissingRows = useMemo(() => {
    const query = fmlSearch.toLowerCase().trim();

    if (!query) return fmlMissingRows;

    return fmlMissingRows.filter((item) =>
      [
        item.code,
        item.product,
        item.uom,
        item.department,
        item.category,
        item.subCategory,
        item.venueText,
        item.reason,
        item.excelRow,
        item.standardOrderRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(query)
    );
  }, [fmlMissingRows, fmlSearch]);

  const visibleFmlRunningLowRows = useMemo(() => {
    const query = fmlSearch.toLowerCase().trim();

    if (!query) return fmlRunningLowRows;

    return fmlRunningLowRows.filter((item) =>
      [
        item.code,
        item.product,
        item.uom,
        item.department,
        item.category,
        item.subCategory,
        item.venueText,
        item.reason,
        item.excelRow,
        item.standardOrderRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(query)
    );
  }, [fmlRunningLowRows, fmlSearch]);

  const filterCounts = useMemo(
    () => ({
      all: orderRows.length,
      needsOrder: orderRows.filter((item) => Number(item.suggestedOrder || 0) > 0)
        .length,
      covered: orderRows.filter(
        (item) =>
          Number(item.suggestedOrder || 0) <= 0 &&
          Number(item.averageConsumptionPerDay || 0) > 0
      ).length,
      review: orderRows.filter(
        (item) => Number(item.averageConsumptionPerDay || 0) <= 0
      ).length,
      shortBeforeArrival: orderRows.filter(
        (item) => Number(item.preArrivalShortage || 0) > 0
      ).length,
      fastSpoilage: orderRows.filter((item) => item.produceRuleType === "fast")
        .length,
      longHoldProduce: orderRows.filter((item) => item.produceRuleType === "long")
        .length,
    }),
    [orderRows]
  );

  const getOrderExportRows = (rows = visibleOrderRows) =>
    rows.map((item, index) => ({
      Line: index + 1,
      ExcelRow: item.excelRow,
      Code: item.code || "",
      Product: item.product,
      UOM: item.uom,
      StockOnHand: Number(item.stockOnHand || 0),
      OnboardOrderedQtyColumnY: Number(item.onboardOrderedQty ?? 0),
      FutureOrdersFtoNUsedForArrivalCoverage: Number(item.futureOrders || 0),
      PastConsumption: Number(item.pastConsumption || 0),
      AverageConsumptionPerDay: Number(item.averageConsumptionPerDay || 0),
      DaysUntilArrival: Number(item.daysUntilArrival || 0),
      ConsumptionUntilArrival: Number(item.consumptionUntilArrival || 0),
      EstimatedQtyAtArrival: Number(item.estimatedQtyAtArrival || 0),
      PreArrivalShortageHighlightedOnly: Number(item.preArrivalShortage || 0),
      UsableQtyAtArrivalForOrder: Number(item.usableQtyAtArrivalForOrder || 0),
      VoyageDays: Number(item.voyageDays || 0),
      VoyageNeed: Number(item.voyageNeed || 0),
      OrderBufferPercent: ORDER_BUFFER_PERCENT,
      OrderBufferQty: Number(item.orderBufferQty || 0),
      TargetQtyForVoyage: Number(item.targetQtyForVoyage || 0),
      SuggestedAdditionalOrder: Number(item.suggestedOrder || 0),
      CurrentFileParQ: Number(item.parLevel || 0),
      SuggestedParCurrentUsageVoyage: Number(item.currentUsageSuggestedPar || 0),
      SuggestedParCurrentUsage7Days: Number(
        item.currentUsageSuggestedPar7Days || 0
      ),
      CoverageDifferenceAfterFutureOrder: Number(
        item.futureCoverageDifference || 0
      ),
      FreshProduceRule: item.produceRuleLabel || "",
      Status: item.statusLabel,
    }));

  const getParComparisonExportRows = (rows = visibleOrderRows) =>
    rows.map((item, index) => {
      const allRegionsStats = getRegionalStatsForItem({
        item,
        yearlyRegionalConsumption,
        regionalLookup: yearlyRegionalLookup,
        regionFilter: YEARLY_REGION_ALL,
        shipFilter: "",
        parDays: PAR_REPORT_DAYS,
      });

      const selectedMarketStats =
        selectedRegionalConsumptionRegion &&
        selectedRegionalConsumptionRegion !== YEARLY_REGION_ALL
          ? getRegionalStatsForItem({
              item,
              yearlyRegionalConsumption,
              regionalLookup: yearlyRegionalLookup,
              regionFilter: selectedRegionalConsumptionRegion,
              shipFilter: "",
              parDays: PAR_REPORT_DAYS,
            })
          : allRegionsStats;

      // IMPORTANT:
      // This is column Y only.
      // Do not use futureOrders here.
      // Do not use || futureOrders as fallback, because blank/zero column Y must stay 0.
      const onboardOrderedQty = Number(item.onboardOrderedQty ?? 0);
      const suggestedOrderQty = Number(item.suggestedOrder ?? 0);

      return {
        Line: index + 1,
        ExcelRow: item.excelRow,
        Code: item.code || "",
        Product: item.product,
        UOM: item.uom,

        OnboardOrderedQtyColumnY: onboardOrderedQty,
        SuggestedOrderQty: suggestedOrderQty,
        DifferenceSuggestedVsOnboardOrdered:
          suggestedOrderQty - onboardOrderedQty,

        CurrentParLevelQ: Number(item.parLevel ?? 0),
        CurrentUsageAvgPerDay: Number(item.averageConsumptionPerDay ?? 0),
        SuggestedParCurrentUsage7Days: Number(
          item.currentUsageSuggestedPar7Days ?? 0
        ),

        YearlyAllRegionsAvgPerDay: Number(allRegionsStats.avgDailyQty ?? 0),
        YearlyAllRegionsSuggestedPar7Days: Number(
          allRegionsStats.suggestedPar ?? 0
        ),

        SelectedMarket:
          selectedRegionalConsumptionRegion &&
          selectedRegionalConsumptionRegion !== YEARLY_REGION_ALL
            ? selectedRegionalConsumptionRegion
            : "All regions",

        YearlySelectedMarketAvgPerDay: Number(
          selectedMarketStats.avgDailyQty ?? 0
        ),

        YearlySelectedMarketSuggestedPar7Days: Number(
          selectedMarketStats.suggestedPar ?? 0
        ),

        PreArrivalShortageHighlightedOnly: Number(item.preArrivalShortage ?? 0),
        EstimatedQtyAtArrival: Number(item.estimatedQtyAtArrival ?? 0),
        FreshProduceRule: item.produceRuleLabel || "",
        Status: item.statusLabel,
      };
    });

  const parComparisonRows = useMemo(() => {
    if (view !== "parReport") return [];

    return getParComparisonExportRows(visibleOrderRows);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    view,
    visibleOrderRows,
    yearlyRegionalLookup,
    selectedRegionalConsumptionRegion,
  ]);

  const displayedParComparisonRows = useMemo(
    () => parComparisonRows.slice(0, PAR_REPORT_DISPLAY_LIMIT),
    [parComparisonRows]
  );

  const exportOrderView = () => {
    exportRowsToExcel(
      getOrderExportRows(),
      "Suggested Order",
      `suggested-next-order-${activeShipCode || userShip || "ship"}.xlsx`
    );
  };

  const printOrderView = () => {
    printSimpleTable({
      title: "Suggested Next Order",
      subtitle: `${activeShipName} • ${orderFileName || "No file name"}`,
      rows: getOrderExportRows(),
    });
  };

  const exportParComparisonReport = () => {
    exportRowsToExcel(
      getParComparisonExportRows(visibleOrderRows),
      "Order vs Par",
      `order-vs-par-comparison-${activeShipCode || userShip || "ship"}.xlsx`
    );
  };

  const printParComparisonReport = () => {
    printSimpleTable({
      title: "Order vs Par Comparison",
      subtitle: `${activeShipName} • 7-day par comparison • ${orderFileName}`,
      rows: getParComparisonExportRows(visibleOrderRows),
    });
  };

  const exportFmlMissing = () => {
    const rows = visibleFmlMissingRows.map((item, index) => ({
      Line: index + 1,
      FMLRow: item.excelRow,
      StandardOrderRow: item.standardOrderRow || "Not found",
      Code: item.code || "",
      Product: item.product || "",
      UOM: item.uom || "",
      Department: item.department || "",
      Category: item.category || "",
      SubCategory: item.subCategory || "",
      VenuesFromFML: item.venueText || "",
      StockOnHand: Number(item.stockOnHand || 0),
      FutureOrders: Number(item.futureOrders || 0),
      PastConsumption: Number(item.pastConsumption || 0),
      AverageConsumptionPerDay: Number(item.averageConsumptionPerDay || 0),
      EstimatedQtyAtArrival: Number(item.estimatedQtyAtArrival || 0),
      SuggestedOrder: Number(item.suggestedOrder || 0),
      Reason: item.reason || "",
    }));

    exportRowsToExcel(
      rows,
      "FML Not Ordered Not Used",
      `fml-not-ordered-not-used-${activeShipCode || userShip || "ship"}.xlsx`
    );
  };

  const exportFmlRunningLow = () => {
    const rows = visibleFmlRunningLowRows.map((item, index) => ({
      Line: index + 1,
      FMLRow: item.excelRow,
      StandardOrderRow: item.standardOrderRow || "",
      Code: item.code || "",
      Product: item.product || "",
      UOM: item.uom || "",
      Department: item.department || "",
      Category: item.category || "",
      SubCategory: item.subCategory || "",
      VenuesFromFML: item.venueText || "",
      StockOnHand: Number(item.stockOnHand || 0),
      FutureOrders: Number(item.futureOrders || 0),
      PastConsumption: Number(item.pastConsumption || 0),
      AverageConsumptionPerDay: Number(item.averageConsumptionPerDay || 0),
      EstimatedQtyAtArrival: Number(item.estimatedQtyAtArrival || 0),
      TargetQtyForVoyage: Number(item.targetQtyForVoyage || 0),
      SuggestedOrder: Number(item.suggestedOrder || 0),
      Reason: item.reason || "",
    }));

    exportRowsToExcel(
      rows,
      "FML Running Low",
      `fml-running-low-${activeShipCode || userShip || "ship"}.xlsx`
    );
  };

  const selectedRecipeUsageRows = useMemo(
    () => getRecipeUsageForItem(selectedRecipeUsageItem, recipeRows),
    [selectedRecipeUsageItem, recipeRows]
  );

  const selectedInfoRegionalStats = useMemo(() => {
    if (!selectedInfoItem) return null;

    const voyageDays = Number(
      selectedInfoItem.voyageDays || orderMeta.voyageDays || 0
    );

    const allRegions = getRegionalStatsForItem({
      item: selectedInfoItem,
      yearlyRegionalConsumption,
      regionalLookup: yearlyRegionalLookup,
      regionFilter: YEARLY_REGION_ALL,
      shipFilter: "",
      parDays: voyageDays,
    });

    const selectedMarket =
      selectedRegionalConsumptionRegion &&
      selectedRegionalConsumptionRegion !== YEARLY_REGION_ALL
        ? getRegionalStatsForItem({
            item: selectedInfoItem,
            yearlyRegionalConsumption,
            regionalLookup: yearlyRegionalLookup,
            regionFilter: selectedRegionalConsumptionRegion,
            shipFilter: "",
            parDays: voyageDays,
          })
        : null;

    const sameShip = activeShipCode
      ? getRegionalStatsForItem({
          item: selectedInfoItem,
          yearlyRegionalConsumption,
          regionalLookup: yearlyRegionalLookup,
          regionFilter: selectedRegionalConsumptionRegion || YEARLY_REGION_ALL,
          shipFilter: activeShipCode,
          parDays: voyageDays,
        })
      : null;

    return {
      allRegions,
      selectedMarket,
      sameShip,
    };
  }, [
    selectedInfoItem,
    yearlyRegionalConsumption,
    yearlyRegionalLookup,
    selectedRegionalConsumptionRegion,
    activeShipCode,
    orderMeta.voyageDays,
  ]);

  const renderRegionalStatsBox = (title, stats) => (
    <div style={styles.infoBox}>
      <strong>{title}</strong>

      {!stats || !stats.matched ? (
        <div style={styles.emptyText}>No last-year regional match found.</div>
      ) : (
        <>
          <div>
            Total qty: <strong>{formatRegionalQty(stats.totalQty)}</strong>
          </div>
          <div>
            Total days: <strong>{formatQty(stats.totalDays)}</strong>
          </div>
          <div>
            Average/day: <strong>{formatRegionalQty(stats.avgDailyQty)}</strong>
          </div>
          <div>
            Suggested par for {formatQty(stats.parDays)} day(s) + 25%:{" "}
            <strong>{formatRegionalQty(stats.suggestedPar)}</strong>
          </div>
          <div>
            Evidence blocks: <strong>{stats.blocks}</strong>
          </div>

          {stats.matchedProducts.length > 0 && (
            <div style={styles.recipeMeta}>
              Match: {stats.matchedProducts.join("; ")}
            </div>
          )}
        </>
      )}
    </div>
  );

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={styles.headerLogo}
        />

        <div style={styles.headerActions}>
          <button type="button" style={styles.backButton} onClick={onBack}>
            ← Back
          </button>

          <div style={styles.shipBadge}>
            🛒 Generate Next Order {activeShipCode ? `• ${activeShipName}` : ""}
          </div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Latest Order Sheet</h2>

          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadOrderFile}
            style={styles.fileInput}
            disabled={loading}
          />

          <div style={styles.infoBox}>
            <div>
              📄 File: <strong>{orderFileName || "Not uploaded"}</strong>
            </div>
            <div>
              🚢 Ship detected:{" "}
              <strong>
                {orderMeta.shipName
                  ? `${orderMeta.shipName} → ${activeShipName}`
                  : "N/A"}
              </strong>
            </div>
            <div>
              📅 Order date B2: <strong>{orderMeta.orderDate || "N/A"}</strong>
            </div>
            <div>
              📅 Arrival date B3:{" "}
              <strong>{orderMeta.arrivalDate || "N/A"}</strong>
            </div>
            <div>
              ⏱️ Days until arrival:{" "}
              <strong>{formatQty(orderMeta.daysUntilArrival || 0)}</strong>
            </div>
            <div>
              🚢 Voyage days B6:{" "}
              <strong>{formatQty(orderMeta.voyageDays || 0)}</strong>
            </div>
            <div>
              🧮 Main order buffer: <strong>{ORDER_BUFFER_PERCENT}%</strong>
            </div>
            <div>
              📦 Column Y ordered qty:{" "}
              <strong>Used only in Order vs Par comparison</strong>
            </div>
            <div>
              🥬 Fresh produce spoilage rules:{" "}
              <strong>
                {orderMeta.isFreshProduceOrder ? "Applied" : "Not applied"}
              </strong>
            </div>
            <div>
              📊 Order vs Par report basis:{" "}
              <strong>{PAR_REPORT_DAYS} days + {ORDER_BUFFER_PERCENT}%</strong>
            </div>
            {loading && <div>Loading...</div>}
          </div>

          {message && <p style={styles.message}>{message}</p>}
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🌎 Last Year Regional Comparison</h2>

          <div style={styles.infoBox}>
            <div>
              📄 Yearly regional file:{" "}
              <strong>{yearlyRegionalFileName || "Not loaded"}</strong>
            </div>
            <div>
              🧭 Selected market:{" "}
              <strong>
                {!selectedRegionalConsumptionRegion
                  ? "Not selected"
                  : selectedRegionalConsumptionRegion === YEARLY_REGION_ALL
                    ? "All regions"
                    : selectedRegionalConsumptionRegion}
              </strong>
            </div>
            <div>
              🔎 Match rule:{" "}
              <strong>
                Code only. If no code, exact product description only.
              </strong>
            </div>
          </div>

          <label style={styles.label}>Upload / replace yearly regional file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadYearlyRegionalFile}
            style={styles.fileInput}
          />

          <label style={styles.label}>Market / region</label>
          <select
            value={selectedRegionalConsumptionRegion || ""}
            onChange={(event) =>
              setSelectedRegionalConsumptionRegion?.(event.target.value)
            }
            style={styles.searchInput}
          >
            <option value="">Select market / region</option>
            <option value={YEARLY_REGION_ALL}>All regions</option>

            {regionalRegionOptions.map((region) => (
              <option key={region} value={region}>
                {region}
              </option>
            ))}
          </select>

          <label style={styles.label}>Regional par buffer % reference only</label>
          <input
            type="number"
            min="0"
            step="1"
            value={regionalParBufferPercent || 0}
            onChange={(event) =>
              setRegionalParBufferPercent?.(Number(event.target.value || 0))
            }
            style={styles.searchInput}
          />

          <p style={styles.emptyText}>
            Main order calculation uses voyage need + 25%. The Order vs Par
            report uses 7 days + 25% only for par comparison.
          </p>
        </div>
      </section>

      <section style={styles.card}>
        <div
          style={{
            ...styles.header,
            boxShadow: "none",
            padding: 0,
            marginBottom: 16,
          }}
        >
          <div>
            <h2 style={styles.productTitle}>🧭 Reports</h2>
            <p style={{ ...styles.emptyText, margin: 0 }}>
              Suggested order stays simple. Use More Info, Recipe Usage, and
              Order vs Par for deeper checks.
            </p>
          </div>

          <div style={styles.headerActions}>
            <button
              type="button"
              style={{
                ...styles.viewModeButton,
                ...(view === "order" ? styles.viewModeButtonActive : {}),
              }}
              onClick={() => setView("order")}
            >
              Suggested Order
            </button>

            <button
              type="button"
              style={{
                ...styles.viewModeButton,
                ...(view === "consumption" ? styles.viewModeButtonActive : {}),
              }}
              onClick={() => setView("consumption")}
            >
              Consumption Report
            </button>

            <button
              type="button"
              style={{
                ...styles.viewModeButton,
                ...(view === "parReport" ? styles.viewModeButtonActive : {}),
              }}
              onClick={() => setView("parReport")}
            >
              Order vs Par
            </button>

            <button
              type="button"
              style={{
                ...styles.viewModeButton,
                ...(view === "fmlMissing" ? styles.viewModeButtonActive : {}),
              }}
              onClick={() => setView("fmlMissing")}
            >
              FML Not Ordered ({fmlMissingRows.length})
            </button>

            <button
              type="button"
              style={{
                ...styles.viewModeButton,
                ...(view === "fmlLow" ? styles.viewModeButtonActive : {}),
              }}
              onClick={() => setView("fmlLow")}
            >
              FML Running Low ({fmlRunningLowRows.length})
            </button>
          </div>
        </div>

        {orderRows.length === 0 && (
          <p style={styles.emptyText}>
            Upload the latest order sheet to generate suggested orders.
          </p>
        )}

        {(view === "order" ||
          view === "consumption" ||
          view === "parReport") &&
          orderRows.length > 0 && (
            <>
              <div style={styles.infoBox}>
                <div>
                  📦 Items loaded: <strong>{orderMeta.totalItems || 0}</strong>
                </div>
                <div>
                  🔵 Suggested order items:{" "}
                  <strong>{orderMeta.itemsNeedingOrder || 0}</strong>
                </div>
                <div>
                  ✅ Covered items: <strong>{orderMeta.coveredItems || 0}</strong>
                </div>
                <div>
                  ⚠️ Review items: <strong>{orderMeta.reviewItems || 0}</strong>
                </div>
                <div>
                  🚨 Short before arrival:{" "}
                  <strong>{orderMeta.preArrivalShortageItems || 0}</strong>
                </div>
                {orderMeta.isFreshProduceOrder && (
                  <>
                    <div>
                      🥬 Quick-spoil produce:{" "}
                      <strong>{orderMeta.fastSpoilageItems || 0}</strong>
                    </div>
                    <div>
                      🥔 Long-hold produce:{" "}
                      <strong>{orderMeta.longHoldProduceItems || 0}</strong>
                    </div>
                  </>
                )}
                <div>
                  📋 FML rows found: <strong>{fmlRows.length}</strong>
                </div>
              </div>

              <input
                placeholder="Search product, code, UOM or status..."
                value={search}
                onChange={(event) => setSearch(event.target.value)}
                style={{ ...styles.searchInput, marginTop: 14 }}
              />

              <div style={styles.viewModeBox}>
                <button
                  type="button"
                  style={{
                    ...styles.viewModeButton,
                    ...(filter === "all" ? styles.viewModeButtonActive : {}),
                  }}
                  onClick={() => setFilter("all")}
                >
                  All ({filterCounts.all})
                </button>

                <button
                  type="button"
                  style={{
                    ...styles.viewModeButton,
                    ...(filter === "needsOrder"
                      ? styles.viewModeButtonActive
                      : {}),
                  }}
                  onClick={() => setFilter("needsOrder")}
                >
                  Needs Order ({filterCounts.needsOrder})
                </button>

                <button
                  type="button"
                  style={{
                    ...styles.viewModeButton,
                    ...(filter === "covered" ? styles.viewModeButtonActive : {}),
                  }}
                  onClick={() => setFilter("covered")}
                >
                  Covered ({filterCounts.covered})
                </button>

                <button
                  type="button"
                  style={{
                    ...styles.viewModeButton,
                    ...(filter === "review" ? styles.viewModeButtonActive : {}),
                  }}
                  onClick={() => setFilter("review")}
                >
                  Review ({filterCounts.review})
                </button>

                <button
                  type="button"
                  style={{
                    ...styles.viewModeButton,
                    ...(filter === "shortBeforeArrival"
                      ? styles.viewModeButtonActive
                      : {}),
                  }}
                  onClick={() => setFilter("shortBeforeArrival")}
                >
                  Short Before Arrival ({filterCounts.shortBeforeArrival})
                </button>

                {orderMeta.isFreshProduceOrder && (
                  <>
                    <button
                      type="button"
                      style={{
                        ...styles.viewModeButton,
                        ...(filter === "fastSpoilage"
                          ? styles.viewModeButtonActive
                          : {}),
                      }}
                      onClick={() => setFilter("fastSpoilage")}
                    >
                      Quick Spoil ({filterCounts.fastSpoilage})
                    </button>

                    <button
                      type="button"
                      style={{
                        ...styles.viewModeButton,
                        ...(filter === "longHoldProduce"
                          ? styles.viewModeButtonActive
                          : {}),
                      }}
                      onClick={() => setFilter("longHoldProduce")}
                    >
                      Long Hold ({filterCounts.longHoldProduce})
                    </button>
                  </>
                )}

                {view === "parReport" ? (
                  <>
                    <button
                      type="button"
                      style={styles.backButton}
                      onClick={printParComparisonReport}
                    >
                      🖨️ Print Order vs Par
                    </button>

                    <button
                      type="button"
                      style={styles.primaryButton}
                      onClick={exportParComparisonReport}
                    >
                      📥 Export Order vs Par
                    </button>
                  </>
                ) : (
                  <>
                    <button
                      type="button"
                      style={styles.backButton}
                      onClick={printOrderView}
                    >
                      🖨️ Print
                    </button>

                    <button
                      type="button"
                      style={styles.primaryButton}
                      onClick={exportOrderView}
                    >
                      📥 Export Excel
                    </button>
                  </>
                )}
              </div>

              {visibleOrderRows.length === 0 && (
                <p style={styles.emptyText}>
                  No products match this search/filter.
                </p>
              )}
            </>
          )}

        {view === "order" && orderRows.length > 0 && (
          <div style={localStyles.orderGrid}>
            {visibleOrderRows.map((item) => (
              <div
                key={item.itemKey}
                style={{
                  ...localStyles.orderCard,
                  ...(item.statusType === "need" ? localStyles.orderCardNeed : {}),
                  ...(item.statusType === "covered"
                    ? localStyles.orderCardCovered
                    : {}),
                  ...(item.statusType === "review"
                    ? localStyles.orderCardReview
                    : {}),
                }}
              >
                <div style={localStyles.orderCardHeader}>
                  <div>
                    <div style={localStyles.productName}>{item.product}</div>
                    <div style={styles.recipeMeta}>
                      Code: {item.code || "N/A"} • UOM: {item.uom || "N/A"} •
                      Row {item.excelRow}
                    </div>
                  </div>

                  <div
                    style={{
                      ...localStyles.statusPill,
                      ...(item.statusType === "need"
                        ? localStyles.statusPillNeed
                        : {}),
                      ...(item.statusType === "covered"
                        ? localStyles.statusPillCovered
                        : {}),
                      ...(item.statusType === "review"
                        ? localStyles.statusPillReview
                        : {}),
                    }}
                  >
                    {item.statusLabel}
                  </div>
                </div>

                <div style={localStyles.metricGrid}>
                  <div style={localStyles.metricBox}>
                    <span>Current avg/day</span>
                    <strong>{formatQty(item.averageConsumptionPerDay)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Future F:N</span>
                    <strong>{formatQty(item.futureOrders)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Est. at arrival</span>
                    <strong
                      style={{
                        color:
                          Number(item.estimatedQtyAtArrival || 0) < 0
                            ? "#b00020"
                            : "#111",
                      }}
                    >
                      {formatQty(item.estimatedQtyAtArrival)}
                    </strong>
                  </div>

                  <div style={localStyles.metricBoxStrong}>
                    <span>Suggested order</span>
                    <strong>{formatQty(item.suggestedOrder)}</strong>
                  </div>
                </div>

                {Number(item.preArrivalShortage || 0) > 0 && (
                  <div style={localStyles.compactBad}>
  Short before arrival: {formatQty(item.preArrivalShortage)}.
  Highlight only, not added to next order.
</div>
                )}

                {orderMeta.isFreshProduceOrder && item.produceRuleType !== "standard" && (
                  <div
  style={
    item.produceRuleType === "fast"
      ? localStyles.compactWarning
      : localStyles.compactInfoBox
  }
>
                    {item.produceRuleLabel}
                    {item.produceOrderFullTarget
                      ? " — order full target quantity."
                      : ""}
                  </div>
                )}

                <div style={localStyles.compactInfoBox}>
                  <div>
                    Voyage need + 25% buffer:{" "}
                    <strong>{formatQty(item.targetQtyForVoyage)}</strong>
                  </div>
                  <div>
                    Coverage after future order:{" "}
                    <strong
                      style={{
                        color:
                          Number(item.futureCoverageDifference || 0) >= 0
                            ? "#2e7d32"
                            : "#b00020",
                      }}
                    >
                      {Number(item.futureCoverageDifference || 0) >= 0 ? "+" : ""}
                      {formatQty(item.futureCoverageDifference)}
                    </strong>
                  </div>
                </div>

                <div style={styles.headerActions}>
                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() => setSelectedInfoItem(item)}
                  >
                    ℹ️ More Info
                  </button>

                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() => setSelectedRecipeUsageItem(item)}
                  >
                    🍽️ Recipe Usage
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}

        {view === "consumption" && orderRows.length > 0 && (
          <div style={localStyles.reportList}>
            {visibleOrderRows.map((item) => (
              <div key={`consumption-${item.itemKey}`} style={localStyles.reportRow}>
                <div>
                  <strong>{item.product}</strong>
                  <div style={styles.recipeMeta}>
                    Code: {item.code || "N/A"} • UOM: {item.uom || "N/A"} • Row{" "}
                    {item.excelRow}
                  </div>
                </div>

                <div style={localStyles.reportMetrics}>
                  <span>Stock: {formatQty(item.stockOnHand)}</span>
                  <span>Ordered Y: {formatQty(item.onboardOrderedQty)}</span>
                  <span>Future F:N: {formatQty(item.futureOrders)}</span>
                  <span>Past: {formatQty(item.pastConsumption)}</span>
                  <span>Avg/day: {formatQty(item.averageConsumptionPerDay)}</span>
                  <span>Arrival: {formatQty(item.estimatedQtyAtArrival)}</span>
                  <span>Pre-arrival short: {formatQty(item.preArrivalShortage)}</span>
                  <span>Suggested: {formatQty(item.suggestedOrder)}</span>
                  <span>Current par Q: {formatQty(item.parLevel)}</span>
                  <span>
                    Suggested par voyage:{" "}
                    {formatQty(item.currentUsageSuggestedPar)}
                  </span>
                  <span>
                    Suggested par 7 days:{" "}
                    {formatQty(item.currentUsageSuggestedPar7Days)}
                  </span>
                </div>

                <div style={styles.headerActions}>
                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() => setSelectedInfoItem(item)}
                  >
                    ℹ️ More Info
                  </button>

                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() => setSelectedRecipeUsageItem(item)}
                  >
                    🍽️ Recipe Usage
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}

        {view === "parReport" && orderRows.length > 0 && (
          <div style={localStyles.reportList}>
            <div style={styles.infoBox}>
              <div>
                This report compares: current onboard ordered quantity from
                column Y, suggested order quantity, current par Q, current usage
                suggested 7-day par, and last-year 7-day suggested par.
              </div>

              <div>
                Report par basis: <strong>7 days + {ORDER_BUFFER_PERCENT}%</strong>
              </div>

              <div>
                Last-year match rule:{" "}
                <strong>
                  code only; if no code, exact normalized product description only.
                </strong>
              </div>

              <div>
                Main order calculation still uses F:N future orders for arrival
                coverage. This report compares suggested order only against
                column Y ordered quantity.
              </div>
            </div>

            {parComparisonRows.length > PAR_REPORT_DISPLAY_LIMIT && (
              <div style={styles.statusWarning}>
                Showing first {PAR_REPORT_DISPLAY_LIMIT} row(s) on screen to keep
                the app fast. Use search to narrow results or Export Order vs Par
                to download all {parComparisonRows.length} row(s).
              </div>
            )}

            {displayedParComparisonRows.map((row, index) => {
              const orderedColumnY = Number(row.OnboardOrderedQtyColumnY ?? 0);
              const suggestedOrder = Number(row.SuggestedOrderQty ?? 0);
              const difference = Number(
                row.DifferenceSuggestedVsOnboardOrdered ?? 0
              );

              return (
                <div
                  key={`par-${row.ExcelRow}-${row.Code}-${index}`}
                  style={localStyles.parRow}
                >
                  <div>
                    <strong>{row.Product}</strong>

                    <div style={styles.recipeMeta}>
                      Code: {row.Code || "N/A"} • UOM: {row.UOM || "N/A"} • Row{" "}
                      {row.ExcelRow}
                    </div>

                    {row.FreshProduceRule && (
                      <div style={styles.recipeMeta}>
                        Rule: {row.FreshProduceRule}
                      </div>
                    )}
                  </div>

                  <div style={localStyles.parMetricGrid}>
                    <div style={localStyles.metricBox}>
                      <span>Ordered col Y</span>
                      <strong>{formatQty(orderedColumnY)}</strong>
                    </div>

                    <div style={localStyles.metricBoxStrong}>
                      <span>Suggested order</span>
                      <strong>{formatQty(suggestedOrder)}</strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Suggested - ordered Y</span>
                      <strong
                        style={{
                          color: difference > 0 ? "#b00020" : "#2e7d32",
                        }}
                      >
                        {difference >= 0 ? "+" : ""}
                        {formatQty(difference)}
                      </strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Current par Q</span>
                      <strong>{formatQty(row.CurrentParLevelQ)}</strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Current usage par 7d</span>
                      <strong>
                        {formatQty(row.SuggestedParCurrentUsage7Days)}
                      </strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Year all regions par 7d</span>
                      <strong>
                        {formatQty(row.YearlyAllRegionsSuggestedPar7Days)}
                      </strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Year market par 7d</span>
                      <strong>
                        {formatQty(row.YearlySelectedMarketSuggestedPar7Days)}
                      </strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Year market avg/day</span>
                      <strong>{formatQty(row.YearlySelectedMarketAvgPerDay)}</strong>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {(view === "fmlMissing" || view === "fmlLow") && (
          <>
            <input
              placeholder="Search FML product, code, venue or reason..."
              value={fmlSearch}
              onChange={(event) => setFmlSearch(event.target.value)}
              style={styles.searchInput}
            />

            <div style={styles.headerActions}>
              {view === "fmlMissing" ? (
                <>
                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() =>
                      printSimpleTable({
                        title: "FML Not Ordered / Not Used",
                        subtitle: `${activeShipName} • ${orderFileName}`,
                        rows: visibleFmlMissingRows.map((item, index) => ({
                          Line: index + 1,
                          FMLRow: item.excelRow,
                          Code: item.code || "",
                          Product: item.product || "",
                          UOM: item.uom || "",
                          Venues: item.venueText || "",
                          FutureOrders: formatQty(item.futureOrders),
                          PastConsumption: formatQty(item.pastConsumption),
                          Reason: item.reason || "",
                        })),
                      })
                    }
                  >
                    🖨️ Print
                  </button>

                  <button
                    type="button"
                    style={styles.primaryButton}
                    onClick={exportFmlMissing}
                  >
                    📥 Export Excel
                  </button>
                </>
              ) : (
                <>
                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() =>
                      printSimpleTable({
                        title: "FML Running Low",
                        subtitle: `${activeShipName} • ${orderFileName}`,
                        rows: visibleFmlRunningLowRows.map((item, index) => ({
                          Line: index + 1,
                          FMLRow: item.excelRow,
                          Code: item.code || "",
                          Product: item.product || "",
                          UOM: item.uom || "",
                          Venues: item.venueText || "",
                          AvgPerDay: formatQty(item.averageConsumptionPerDay),
                          ArrivalQty: formatQty(item.estimatedQtyAtArrival),
                          SuggestedOrder: formatQty(item.suggestedOrder),
                          Reason: item.reason || "",
                        })),
                      })
                    }
                  >
                    🖨️ Print
                  </button>

                  <button
                    type="button"
                    style={styles.primaryButton}
                    onClick={exportFmlRunningLow}
                  >
                    📥 Export Excel
                  </button>
                </>
              )}
            </div>

            {view === "fmlMissing" && visibleFmlMissingRows.length === 0 && (
              <p style={styles.emptyText}>
                No FML not ordered / not used rows found.
              </p>
            )}

            {view === "fmlLow" && visibleFmlRunningLowRows.length === 0 && (
              <p style={styles.emptyText}>No FML running low rows found.</p>
            )}

            <div style={localStyles.fmlCompactGrid}>
  {(view === "fmlMissing"
    ? visibleFmlMissingRows
    : visibleFmlRunningLowRows
  ).map((item, index) => {
    const reasonText =
      view === "fmlMissing"
        ? item.standardOrderRow
          ? "No order / no use"
          : "Not in order sheet"
        : "Running low";

    return (
      <div
        key={`${view}-${item.code}-${item.product}-${index}`}
        style={{
          ...localStyles.fmlCompactCard,
          ...(view === "fmlMissing"
            ? localStyles.fmlCompactCardMissing
            : localStyles.fmlCompactCardLow),
        }}
      >
        <div style={localStyles.fmlTopLine}>FML row {item.excelRow}</div>

        <div style={localStyles.fmlName}>
          {item.product || "Unnamed item"}
        </div>

        <div style={localStyles.fmlMeta}>
          Code: {item.code || "N/A"} • UOM: {item.uom || "N/A"}
        </div>

        <div style={localStyles.fmlMiniGrid}>
          <div style={localStyles.fmlMiniBox}>
            <span>Future</span>
            <strong>{formatQty(item.futureOrders)}</strong>
          </div>

          <div style={localStyles.fmlMiniBox}>
            <span>Past</span>
            <strong>{formatQty(item.pastConsumption)}</strong>
          </div>

          {view === "fmlLow" && (
            <div style={localStyles.fmlMiniBoxStrong}>
              <span>Suggest</span>
              <strong>{formatQty(item.suggestedOrder)}</strong>
            </div>
          )}
        </div>

        <div style={localStyles.fmlReason}>{reasonText}</div>

        <button
          type="button"
          style={localStyles.fmlRecipeButton}
          onClick={() => setSelectedRecipeUsageItem(item)}
        >
          🍽️ Recipe / Details
        </button>
      </div>
    );
  })}
            </div>
          </>
        )}
      </section>

      {selectedInfoItem && (
        <div
          style={styles.modalBackdrop}
          onClick={() => setSelectedInfoItem(null)}
        >
          <div
            style={localStyles.infoModal}
            onClick={(event) => event.stopPropagation()}
          >
            <button
              type="button"
              style={styles.closeButton}
              onClick={() => setSelectedInfoItem(null)}
            >
              ✕
            </button>

            <h2 style={styles.productTitle}>ℹ️ {selectedInfoItem.product}</h2>

            <p style={styles.emptyText}>
              Code: {selectedInfoItem.code || "N/A"} • UOM:{" "}
              {selectedInfoItem.uom || "N/A"} • Excel row{" "}
              {selectedInfoItem.excelRow}
            </p>

            <section style={localStyles.detailGrid}>
              <div style={styles.infoBox}>
                <strong>Main calculation</strong>
                <div>
                  Stock on hand:{" "}
                  <strong>{formatQty(selectedInfoItem.stockOnHand)}</strong>
                </div>
                <div>
                  Column Y ordered qty:{" "}
                  <strong>{formatQty(selectedInfoItem.onboardOrderedQty)}</strong>
                </div>
                <div>
                  F:N future orders for arrival coverage:{" "}
                  <strong>{formatQty(selectedInfoItem.futureOrders)}</strong>
                </div>
                <div>
                  Days until arrival:{" "}
                  <strong>{formatQty(selectedInfoItem.daysUntilArrival)}</strong>
                </div>
                <div>
                  Consumption until arrival:{" "}
                  <strong>
                    {formatQty(selectedInfoItem.consumptionUntilArrival)}
                  </strong>
                </div>
                <div>
                  Estimated qty at arrival:{" "}
                  <strong
                    style={{
                      color:
                        Number(selectedInfoItem.estimatedQtyAtArrival || 0) < 0
                          ? "#b00020"
                          : "#111",
                    }}
                  >
                    {formatQty(selectedInfoItem.estimatedQtyAtArrival)}
                  </strong>
                </div>
                <div>
                  Pre-arrival shortage highlighted only:{" "}
                  <strong>{formatQty(selectedInfoItem.preArrivalShortage)}</strong>
                </div>
              </div>

              <div style={styles.infoBox}>
                <strong>Voyage need</strong>
                <div>
                  Current average/day:{" "}
                  <strong>
                    {formatQty(selectedInfoItem.averageConsumptionPerDay)}
                  </strong>
                </div>
                <div>
                  Voyage days:{" "}
                  <strong>{formatQty(selectedInfoItem.voyageDays)}</strong>
                </div>
                <div>
                  Voyage need:{" "}
                  <strong>{formatQty(selectedInfoItem.voyageNeed)}</strong>
                </div>
                <div>
                  25% order buffer:{" "}
                  <strong>{formatQty(selectedInfoItem.orderBufferQty)}</strong>
                </div>
                <div>
                  Target qty for voyage:{" "}
                  <strong>
                    {formatQty(selectedInfoItem.targetQtyForVoyage)}
                  </strong>
                </div>
              </div>

              <div style={styles.infoBox}>
                <strong>Order decision</strong>
                <div>
                  Suggested additional order:{" "}
                  <strong>{formatQty(selectedInfoItem.suggestedOrder)}</strong>
                </div>
                <div>
                  Coverage after future order:{" "}
                  <strong
                    style={{
                      color:
                        Number(selectedInfoItem.futureCoverageDifference || 0) >=
                        0
                          ? "#2e7d32"
                          : "#b00020",
                    }}
                  >
                    {Number(selectedInfoItem.futureCoverageDifference || 0) >= 0
                      ? "+"
                      : ""}
                    {formatQty(selectedInfoItem.futureCoverageDifference)}
                  </strong>
                </div>
                <div>
                  Status: <strong>{selectedInfoItem.statusLabel}</strong>
                </div>
                <div>{selectedInfoItem.statusDescription}</div>
                {selectedInfoItem.produceRuleLabel && (
                  <div>
                    Fresh produce rule:{" "}
                    <strong>{selectedInfoItem.produceRuleLabel}</strong>
                  </div>
                )}
              </div>

              <div style={styles.infoBox}>
                <strong>Par comparison</strong>
                <div>
                  Current file par Q:{" "}
                  <strong>{formatQty(selectedInfoItem.parLevel)}</strong>
                </div>
                <div>
                  Suggested par from current usage for voyage:{" "}
                  <strong>
                    {formatQty(selectedInfoItem.currentUsageSuggestedPar)}
                  </strong>
                </div>
                <div>
                  Suggested par from current usage for 7 days:{" "}
                  <strong>
                    {formatQty(selectedInfoItem.currentUsageSuggestedPar7Days)}
                  </strong>
                </div>
                <div>
                  Order vs Par compares suggested order against:{" "}
                  <strong>column Y ordered quantity</strong>
                </div>
              </div>
            </section>

            <h3 style={styles.sectionTitle}>🌎 Last year consumption comparison</h3>

            <div style={styles.infoBox}>
              Match rule: <strong>code only</strong>. If this item has no code,
              the system matches only the exact normalized product description.
            </div>

            <section style={localStyles.detailGrid}>
              {renderRegionalStatsBox(
                "All regions last year",
                selectedInfoRegionalStats?.allRegions
              )}

              {renderRegionalStatsBox(
                selectedRegionalConsumptionRegion
                  ? `${selectedRegionalConsumptionRegion} last year`
                  : "Selected market last year",
                selectedInfoRegionalStats?.selectedMarket
              )}

              {renderRegionalStatsBox(
                `${activeShipName} last year`,
                selectedInfoRegionalStats?.sameShip
              )}
            </section>
          </div>
        </div>
      )}

      {selectedRecipeUsageItem && (
        <div
          style={styles.modalBackdrop}
          onClick={() => setSelectedRecipeUsageItem(null)}
        >
          <div
            style={localStyles.infoModal}
            onClick={(event) => event.stopPropagation()}
          >
            <button
              type="button"
              style={styles.closeButton}
              onClick={() => setSelectedRecipeUsageItem(null)}
            >
              ✕
            </button>

            <h2 style={styles.productTitle}>
              🍽️ Recipe Usage:{" "}
              {selectedRecipeUsageItem.product || selectedRecipeUsageItem.name}
            </h2>

            <p style={styles.emptyText}>
              Uses the permanent Ingredient by Location file from page.js.
            </p>

            <div style={styles.infoBox}>
  <div>
    Recipe rows loaded:{" "}
    <strong>{Math.max((recipeRows || []).length - 1, 0)}</strong>
  </div>

  <div>
    Matches found: <strong>{selectedRecipeUsageRows.length}</strong>
  </div>

  {selectedRecipeUsageItem.excelRow && (
    <div>
      FML row: <strong>{selectedRecipeUsageItem.excelRow}</strong>
    </div>
  )}

  {selectedRecipeUsageItem.code && (
    <div>
      Code: <strong>{selectedRecipeUsageItem.code}</strong>
    </div>
  )}

  {selectedRecipeUsageItem.uom && (
    <div>
      UOM: <strong>{selectedRecipeUsageItem.uom}</strong>
    </div>
  )}

  {selectedRecipeUsageItem.venueText && (
    <div>
      FML venue(s): <strong>{selectedRecipeUsageItem.venueText}</strong>
    </div>
  )}

  {selectedRecipeUsageItem.reason && (
    <div>
      Reason: <strong>{selectedRecipeUsageItem.reason}</strong>
    </div>
  )}
</div>

            {selectedRecipeUsageRows.length === 0 ? (
              <p style={styles.emptyText}>
                No recipe usage found for this product/code in Ingredient by
                Location.
              </p>
            ) : (
              <div style={localStyles.recipeUsageList}>
                {selectedRecipeUsageRows.map((row, index) => (
                  <div
                    key={`${row.venue}-${row.recipeCode}-${row.recipeName}-${index}`}
                    style={localStyles.recipeUsageCard}
                  >
                    <div style={styles.recipeName}>{row.recipeName}</div>
                    <div style={styles.recipeMeta}>
                      Venue: <strong>{row.venue}</strong>
                    </div>
                    {row.menuName && (
                      <div style={styles.recipeMeta}>Menu: {row.menuName}</div>
                    )}
                    <div style={styles.recipeMeta}>
                      Recipe code: {row.recipeCode || "N/A"}
                    </div>
                    <div style={styles.recipeMeta}>
                      Ingredient code: {row.ingredientCode || "N/A"}
                    </div>
                    <div style={styles.recipeMeta}>
                      Ingredient: {row.ingredientName || "N/A"}
                    </div>
                    {row.assignedProduct && (
                      <div style={styles.recipeMeta}>
                        Assigned product: {row.assignedProduct}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}
    </main>
  );
}

  const localStyles = {
  orderGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(min(100%, 165px), 1fr))",
    gap: 8,
    marginTop: 12,
  },

  orderCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 9,
    background: "#fff",
    display: "grid",
    gap: 6,
    boxShadow: "0 4px 12px rgba(0,0,0,0.05)",
    fontSize: 12,
  },

  orderCardNeed: {
    border: "2px solid #0057b8",
    background: "#eef5ff",
  },

  orderCardCovered: {
    border: "2px solid #2e7d32",
    background: "#f0fff4",
  },

  orderCardReview: {
    border: "2px solid #8a5a00",
    background: "#fff8e1",
  },

  orderCardHeader: {
    display: "grid",
    gap: 5,
  },

  productName: {
    fontWeight: "bold",
    fontSize: 13,
    lineHeight: 1.12,
    overflowWrap: "anywhere",
  },

  statusPill: {
    justifySelf: "start",
    padding: "4px 7px",
    borderRadius: 999,
    background: "#f2f2f2",
    color: "#555",
    fontWeight: "bold",
    fontSize: 10,
    whiteSpace: "nowrap",
  },

  statusPillNeed: {
    background: "#0057b8",
    color: "#fff",
  },

  statusPillCovered: {
    background: "#2e7d32",
    color: "#fff",
  },

  statusPillReview: {
    background: "#8a5a00",
    color: "#fff",
  },

  metricGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    gap: 5,
  },

  metricBox: {
    padding: 6,
    borderRadius: 10,
    background: "#fff",
    border: "1px solid #ddd",
    display: "grid",
    gap: 2,
    textAlign: "center",
    fontSize: 10,
    lineHeight: 1.15,
  },

  metricBoxStrong: {
    padding: 6,
    borderRadius: 10,
    background: "#111",
    color: "#fff",
    border: "1px solid #111",
    display: "grid",
    gap: 2,
    textAlign: "center",
    fontSize: 10,
    lineHeight: 1.15,
  },

  compactInfoBox: {
    padding: 7,
    borderRadius: 10,
    background: "#f2f2f2",
    display: "grid",
    gap: 3,
    fontSize: 11,
    lineHeight: 1.2,
  },

  compactWarning: {
    padding: 7,
    borderRadius: 10,
    background: "#fff4d6",
    color: "#8a5a00",
    fontWeight: "bold",
    fontSize: 11,
    lineHeight: 1.2,
    textAlign: "center",
  },

  compactBad: {
    padding: 7,
    borderRadius: 10,
    background: "#b00020",
    color: "#fff",
    fontWeight: "bold",
    fontSize: 11,
    lineHeight: 1.2,
    textAlign: "center",
  },

  reportList: {
    display: "grid",
    gap: 8,
    marginTop: 14,
  },

  reportRow: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 12,
    background: "#fafafa",
    display: "grid",
    gap: 9,
  },

  reportMetrics: {
    display: "flex",
    flexWrap: "wrap",
    gap: 7,
    fontSize: 12,
    color: "#555",
  },

  parRow: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 12,
    background: "#fafafa",
    display: "grid",
    gap: 10,
  },

  parMetricGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(135px, 1fr))",
    gap: 8,
  },

  infoModal: {
    background: "#fff",
    borderRadius: 18,
    padding: 22,
    maxWidth: 1100,
    width: "96%",
    maxHeight: "90vh",
    overflowY: "auto",
    position: "relative",
    boxShadow: "0 20px 60px rgba(0,0,0,0.25)",
  },

  detailGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))",
    gap: 12,
    marginBottom: 12,
  },

  recipeUsageList: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(230px, 1fr))",
    gap: 10,
    marginTop: 14,
  },

  recipeUsageCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 12,
    background: "#fafafa",
    display: "grid",
    gap: 4,
  },
      fmlCompactGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(min(100%, 150px), 1fr))",
    gap: 8,
    marginTop: 12,
  },

  fmlCompactCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 9,
    background: "#fff",
    display: "grid",
    gap: 6,
    fontSize: 11,
    alignContent: "start",
    minWidth: 0,
  },

  fmlCompactCardMissing: {
    border: "2px solid #b00020",
    background: "#fff0f0",
  },

  fmlCompactCardLow: {
    border: "2px solid #0057b8",
    background: "#eef5ff",
  },

  fmlTopLine: {
    color: "#555",
    fontSize: 11,
    fontWeight: "bold",
  },

  fmlName: {
    fontWeight: "bold",
    fontSize: 12.5,
    lineHeight: 1.12,
    overflowWrap: "anywhere",
  },

  fmlMeta: {
    color: "#555",
    fontSize: 11,
    lineHeight: 1.2,
    overflowWrap: "anywhere",
  },

  fmlMiniGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    gap: 5,
  },

  fmlMiniBox: {
    padding: 5,
    borderRadius: 9,
    background: "#fff",
    border: "1px solid #ddd",
    display: "grid",
    gap: 2,
    textAlign: "center",
    fontSize: 10,
    lineHeight: 1.1,
  },

  fmlMiniBoxStrong: {
    gridColumn: "1 / -1",
    padding: 5,
    borderRadius: 9,
    background: "#111",
    color: "#fff",
    border: "1px solid #111",
    display: "grid",
    gap: 2,
    textAlign: "center",
    fontSize: 10,
    lineHeight: 1.1,
  },

  fmlReason: {
    padding: 6,
    borderRadius: 9,
    background: "#fff4d6",
    color: "#8a5a00",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 10.5,
    lineHeight: 1.15,
  },

  fmlRecipeButton: {
    border: "1px solid #ccc",
    borderRadius: 999,
    background: "#fff",
    color: "#111",
    padding: "7px 8px",
    cursor: "pointer",
    fontWeight: "bold",
    fontSize: 11,
    lineHeight: 1.1,
  },
};

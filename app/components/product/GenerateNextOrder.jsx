"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];
const REPORT_RENDER_BATCH = 120;
const HISTORICAL_CONSUMPTION_COLUMNS = [35, 36, 37, 38, 39, 40]; // AJ:AO, 0-based indexes
const CONSUMPTION_INCREASE_THRESHOLD_PERCENT = 25;

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const compactText = (value) =>
  cleanText(value)
    .replace(/[^A-Z0-9]/g, "")
    .replace(/S$/g, "");

const formatQty = (value) => Number(value || 0).toFixed(2);
const formatMoney = (value) => "$" + Number(value || 0).toFixed(2);

const toNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
};

const getHistoricalSailorDays = (cellA, cellB) => {
  const a = toNumber(cellA);
  const b = toNumber(cellB);

  if (!a && !b) return 0;
  if (a && !b) return a;
  if (!a && b) return b;

  const low = Math.min(Math.abs(a), Math.abs(b));
  const high = Math.max(Math.abs(a), Math.abs(b));

  // Older order files store total historical sailor-days directly in AI5:AN5
  // and the number of past days in AI6:AN6. If one value is much larger,
  // use that large value as the historical basis instead of multiplying again.
  if (low > 0 && high > low * 1000) return high;

  // Otherwise treat the two cells as sailors x days.
  return a * b;
};

const yieldToBrowser = () => new Promise((resolve) => setTimeout(resolve, 0));

const getCellValue = (worksheet, address) => worksheet?.[address]?.v ?? "";

const normalizeShipCode = (value) => {
  const text = cleanText(value).replace(/RESILIANT/g, "RESILIENT");

  if (!text) return "";
  if (text === "BRL" || text.includes("BRILLIANT")) return "BRL";
  if (text === "RL" || text.includes("RESILIENT")) return "RL";
  if (text === "SC" || text.includes("SCARLET") || text.includes("SCL")) return "SC";
  if (text === "VL" || text.includes("VALIANT") || text.includes("VAL")) return "VL";

  return "";
};

const getShipDisplayName = (ship) => {
  const map = {
    BRL: "Brilliant Lady",
    RL: "Resilient Lady",
    SC: "Scarlet Lady",
    VL: "Valiant Lady",
  };

  return map[ship] || ship || "N/A";
};

const getExcelDate = (value) => {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }

  const text = String(value || "").trim();
  if (!text) return null;

  const date = new Date(text);
  if (!isNaN(date.getTime())) return date;

  return null;
};

const formatDate = (value) => {
  const date = getExcelDate(value);
  if (!date) return String(value || "");

  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "2-digit",
  });
};

const daysBetween = (startValue, endValue) => {
  const start = getExcelDate(startValue);
  const end = getExcelDate(endValue);

  if (!start || !end) return 0;

  const diff = Math.ceil((end.getTime() - start.getTime()) / 86400000);
  return Math.max(diff, 0);
};

const sumRowRange = (row, startIndex, endIndex) => {
  let total = 0;

  for (let i = startIndex; i <= endIndex; i += 1) {
    total += toNumber(row?.[i]);
  }

  return total;
};

const averagePositiveValues = (values) => {
  const cleanValues = values
    .map((value) => Number(value || 0))
    .filter((value) => Number.isFinite(value) && value > 0);

  if (!cleanValues.length) return 0;

  return cleanValues.reduce((sum, value) => sum + value, 0) / cleanValues.length;
};

const getIncreasePercent = (currentValue, baselineValue) => {
  const current = Number(currentValue || 0);
  const baseline = Number(baselineValue || 0);

  if (baseline <= 0 || current <= 0) return 0;

  return ((current - baseline) / baseline) * 100;
};

const getHistoricalConsumptionPeriods = ({ workbookRows, row, targetSailors }) => {
  const daysRow = workbookRows[5] || []; // Excel row 6: # days
  const paxRow = workbookRows[6] || workbookRows[4] || []; // Excel row 7, fallback row 5
  const targetPax = Number(targetSailors || 0) > 0 ? Number(targetSailors || 0) : 2500;

  return HISTORICAL_CONSUMPTION_COLUMNS.map((columnIndex) => {
    const quantity = toNumber(row?.[columnIndex]);
    const days = toNumber(daysRow?.[columnIndex]);
    const pax = toNumber(paxRow?.[columnIndex]);

    const normalizedDaily =
      quantity > 0 && days > 0 && pax > 0
        ? (quantity / days / pax) * targetPax
        : 0;

    return {
      columnIndex,
      columnLetter: XLSX.utils.encode_col(columnIndex),
      quantity,
      days,
      pax,
      normalizedDaily,
    };
  });
};

const getConsumptionIncreaseMetrics = ({ workbookRows, row, targetSailors }) => {
  const periods = getHistoricalConsumptionPeriods({
    workbookRows,
    row,
    targetSailors,
  });

  // AJ:AO = 6 historical voyages.
  // Last 1 voyage = AO; baseline = average AJ:AN.
  const previousFive = periods.slice(0, 5);
  const latestOne = periods[5];

  const previousFiveVoyageDailyAverage = averagePositiveValues(
    previousFive.map((period) => period.normalizedDaily)
  );

  const latestOneVoyageDaily = Number(latestOne?.normalizedDaily || 0);
  const oneVoyageIncreasePercent = getIncreasePercent(
    latestOneVoyageDaily,
    previousFiveVoyageDailyAverage
  );

  // Last 2 voyages = AN:AO; baseline = average AJ:AM.
  const previousFour = periods.slice(0, 4);
  const latestTwo = periods.slice(4, 6);

  const previousFourVoyageDailyAverage = averagePositiveValues(
    previousFour.map((period) => period.normalizedDaily)
  );

  const latestTwoVoyageDailyAverage = averagePositiveValues(
    latestTwo.map((period) => period.normalizedDaily)
  );

  const twoVoyageIncreasePercent = getIncreasePercent(
    latestTwoVoyageDailyAverage,
    previousFourVoyageDailyAverage
  );

  return {
    historicalConsumptionPeriods: periods,

    latestOneVoyageColumn: latestOne?.columnLetter || "",
    latestOneVoyageQty: latestOne?.quantity || 0,
    latestOneVoyageDaily,
    previousFiveVoyageDailyAverage,
    oneVoyageIncreasePercent,
    oneVoyageIncreaseFlag:
      oneVoyageIncreasePercent >= CONSUMPTION_INCREASE_THRESHOLD_PERCENT,

    latestTwoVoyageColumns: latestTwo.map((period) => period.columnLetter).join(":"),
    latestTwoVoyageQty: latestTwo.reduce(
      (sum, period) => sum + Number(period.quantity || 0),
      0
    ),
    latestTwoVoyageDailyAverage,
    previousFourVoyageDailyAverage,
    twoVoyageIncreasePercent,
    twoVoyageIncreaseFlag:
      twoVoyageIncreasePercent >= CONSUMPTION_INCREASE_THRESHOLD_PERCENT,
  };
};


const productNamesMatch = (left, right) => {
  const a = compactText(left);
  const b = compactText(right);

  if (!a || !b) return false;
  if (a === b) return true;
  if (a.length > 8 && b.includes(a)) return true;
  if (b.length > 8 && a.includes(b)) return true;

  const aWords = cleanText(left)
    .split(" ")
    .map((word) => word.replace(/[^A-Z0-9]/g, ""))
    .filter((word) => word.length > 2);

  const bText = cleanText(right);
  const strongWords = aWords.filter(
    (word) => !["THE", "AND", "WITH", "FOR", "FRESH", "FROZEN", "CASE", "PACK"].includes(word)
  );

  if (strongWords.length >= 2 && strongWords.slice(0, 4).every((word) => bText.includes(word))) {
    return true;
  }

  return false;
};

const getSheetShipScope = (sheetName) => {
  const text = cleanText(sheetName);
  const ships = [];

  if (text.includes("SCL") || text.includes("SCARLET")) ships.push("SC");
  if (text.includes("VAL") || text.includes("VALIANT")) ships.push("VL");
  if (text.includes("RES") || text.includes("RESILIENT")) ships.push("RL");
  if (text.includes("BRL") || text.includes("BRILLIANT")) ships.push("BRL");

  return ships.length ? ships : [...SHIPS];
};

const getScopeText = (scope) => {
  if (!scope || scope.length === 0 || scope.length === SHIPS.length) return "Used by all ships";
  return "Used only on " + scope.join(", ");
};

const parseTemplateWorkbook = (workbook) => {
  const entries = [];

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    const shipScope = getSheetShipScope(sheetName);

    rows.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (cleanText(cell) !== "INGREDIENT NAME") return;

        const titleCells = [
          rows[rowIndex - 1]?.[colIndex],
          rows[rowIndex - 1]?.[colIndex - 1],
          rows[0]?.[colIndex],
          rows[0]?.[colIndex - 1],
          sheetName,
        ];

        const templateTitle =
          titleCells.map((value) => String(value || "").trim()).find(Boolean) || sheetName;

        for (let r = rowIndex + 1; r < rows.length; r += 1) {
          const product = String(rows[r]?.[colIndex] || "").replace(/\s+/g, " ").trim();
          const productKey = cleanText(product);

          if (!product) continue;
          if (productKey === "INGREDIENT NAME") continue;
          if (productKey === "CODE" || productKey === "UM") continue;
          if (productKey.includes("#REF")) continue;

          entries.push({
            product,
            normalizedProduct: compactText(product),
            sheetName,
            templateTitle,
            location: sheetName + " - " + templateTitle,
            shipScope,
            scopeText: getScopeText(shipScope),
          });
        }
      });
    });
  });

  return entries;
};

const findTemplateMatches = (templateEntries, productName, shipCode) => {
  if (!Array.isArray(templateEntries) || !templateEntries.length) return [];

  const matches = templateEntries.filter((entry) => {
    const shipAllowed = !shipCode || entry.shipScope.includes(shipCode);
    return shipAllowed && productNamesMatch(productName, entry.product);
  });

  const seen = new Set();

  return matches.filter((match) => {
    const key = cleanText(match.location + "|" + match.scopeText);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

const splitVenues = (value) =>
  String(value || "")
    .split(/[;,\n]/g)
    .map((venue) => venue.trim())
    .filter(Boolean);

const parseFmlRows = (workbook) => {
  const sheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "FML") ||
    workbook.SheetNames.find((name) => cleanText(name).includes("FML"));

  if (!sheetName) return [];

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });

  return rows
    .map((row, index) => {
      const excelRow = index + 1;
      const code = String(row[3] || "").trim();
      const product = String(row[4] || "").replace(/\s+/g, " ").trim();
      const venuesText = String(row[5] || "").trim();
      const category = String(row[2] || "").trim();

      return {
        fmlRow: excelRow,
        code,
        product,
        category,
        venuesText,
        venues: splitVenues(venuesText),
      };
    })
    .filter((item) => item.product && cleanText(item.product) !== "PRODUCT NAME")
    .filter((item) => item.code || item.product)
    .filter((item) => item.venues.length > 0);
};

const getFmlOrderCodeKey = (value) => {
  const cleaned = String(value ?? "")
    .trim()
    .replace(/\.0+$/g, "")
    .replace(/[^A-Z0-9]/gi, "")
    .toUpperCase()
    .replace(/^0+/, "");

  return cleaned || "";
};

const buildOrderedNotInFmlRows = (workbook, orderRows) => {
  const fmlSheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "FML") ||
    workbook.SheetNames.find((name) => cleanText(name).includes("FML"));

  const orderedRows = (orderRows || []).filter(
    (item) => Number(item.orderedByShip || item.orderedQty || 0) > 0
  );

  if (!orderedRows.length) return [];

  if (!fmlSheetName || !workbook.Sheets[fmlSheetName]) {
    return orderedRows.map((item) => ({
      id: "ordered-not-fml-" + item.excelRow + "-" + (item.code || item.product),
      excelRow: item.excelRow,
      code: item.code || "",
      product: item.product || "",
      unit: item.unit || "N/A",
      stock: Number(item.stock || 0),
      futureOrders: Number(item.futureOrders || 0),
      orderedByShip: Number(item.orderedByShip || item.orderedQty || 0),
      suggestedOrder: Number(item.suggestedOrder || 0),
      pastConsumption: Number(item.pastConsumption || 0),
      averagePerDay: Number(item.averagePerDay || 0),
      availableAtArrival: Number(item.availableAtArrival || 0),
      matchType: "No FML sheet",
      reason: "This item was ordered by the ship, but no FML sheet was found in the uploaded workbook.",
    }));
  }

  const fmlRows = XLSX.utils.sheet_to_json(workbook.Sheets[fmlSheetName], {
    header: 1,
    defval: "",
  });

  const fmlCodeSet = new Set();
  const fmlProductKeySet = new Set();
  const fmlProductNames = [];

  fmlRows.forEach((row) => {
    const code = String(row[3] || "").trim();
    const product = String(row[4] || "").replace(/\s+/g, " ").trim();

    if (!code && !product) return;
    if (cleanText(code) === "CODE") return;
    if (cleanText(product) === "PRODUCT NAME") return;
    if (cleanText(product).includes("PRODUCT NAME")) return;

    const codeKey = getFmlOrderCodeKey(code);
    const productKey = compactText(product);

    if (codeKey) fmlCodeSet.add(codeKey);
    if (productKey) fmlProductKeySet.add(productKey);
    if (product) fmlProductNames.push(product);
  });

  return orderedRows
    .filter((item) => {
      const codeKey = getFmlOrderCodeKey(item.code);
      const productKey = compactText(item.product);

      const foundByCode = Boolean(codeKey && fmlCodeSet.has(codeKey));
      const foundByExactProduct = Boolean(productKey && fmlProductKeySet.has(productKey));
      const foundByProductName = fmlProductNames.some((fmlProduct) =>
        productNamesMatch(item.product, fmlProduct)
      );

      return !foundByCode && !foundByExactProduct && !foundByProductName;
    })
    .map((item) => ({
      id: "ordered-not-fml-" + item.excelRow + "-" + (item.code || item.product),
      excelRow: item.excelRow,
      code: item.code || "",
      product: item.product || "",
      unit: item.unit || "N/A",
      stock: Number(item.stock || 0),
      futureOrders: Number(item.futureOrders || 0),
      orderedByShip: Number(item.orderedByShip || item.orderedQty || 0),
      suggestedOrder: Number(item.suggestedOrder || 0),
      pastConsumption: Number(item.pastConsumption || 0),
      averagePerDay: Number(item.averagePerDay || 0),
      availableAtArrival: Number(item.availableAtArrival || 0),
      matchType: "Not found by code or product name",
      reason:
        "This item was ordered by the ship in column Y, but it was not found in the FML sheet by code or product name.",
    }))
    .sort((a, b) => Number(a.excelRow || 0) - Number(b.excelRow || 0));
};

const findOrderWorksheetName = (workbook) => {
  const nonFml = workbook.SheetNames.filter((name) => !cleanText(name).includes("FML"));

  return (
    nonFml.find((name) => cleanText(name).includes("STANDARD")) ||
    nonFml.find((name) => cleanText(name).includes("ORDER")) ||
    nonFml[0] ||
    workbook.SheetNames[0]
  );
};

const parseOrderFile = async (file) => {
  await yieldToBrowser();

  const arrayBuffer = await file.arrayBuffer();
  await yieldToBrowser();

  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheetName = findOrderWorksheetName(workbook);
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  const shipNameRaw = getCellValue(worksheet, "B1");
  const shipCode = normalizeShipCode(shipNameRaw);
  const orderDate = getCellValue(worksheet, "B2");
  const arrivalDate = getCellValue(worksheet, "B3");
  const sailors = toNumber(getCellValue(worksheet, "B5"));
  const voyageDays = toNumber(getCellValue(worksheet, "B6"));
  const daysUntilArrival = daysBetween(orderDate, arrivalDate);

  const historicalDays = rows[4] || [];
  const historicalSailors = rows[5] || [];

  const orderRows = [];
  const orderByCode = {};

  const productRows = rows.slice(9);

  for (let rowOffset = 0; rowOffset < productRows.length; rowOffset += 1) {
    if (rowOffset > 0 && rowOffset % 80 === 0) {
      await yieldToBrowser();
    }

    const row = productRows[rowOffset];
    const excelRow = rowOffset + 10;
    const code = String(row[0] || "").trim();
    const product = String(row[1] || "").replace(/\s+/g, " ").trim();
    const unit = String(row[2] || "").trim() || "N/A";

    if (!product) continue;
    if (cleanText(product).includes("PRODUCT")) continue;

    const stock = toNumber(row[3]);
    const futureOrders = sumRowRange(row, 5, 13);
    const parLevel = toNumber(row[16]);
    const pastConsumption = sumRowRange(row, 34, 39);
    const orderedByShip = toNumber(row[24]); // Y - ordered by ship
    const consumptionIncreaseMetrics = getConsumptionIncreaseMetrics({
      workbookRows: rows,
      row,
      targetSailors: sailors,
    });

    let historicalSailorDays = 0;

    for (let c = 34; c <= 39; c += 1) {
      historicalSailorDays += getHistoricalSailorDays(historicalDays[c], historicalSailors[c]);
    }

    const averagePerSailorDay = historicalSailorDays > 0 ? pastConsumption / historicalSailorDays : 0;
    const averagePerDay = averagePerSailorDay * sailors;
    const usageUntilArrival = averagePerDay * daysUntilArrival;
    const availableAtArrival = stock + futureOrders - usageUntilArrival;
    const projectedVoyageNeed = averagePerDay * voyageDays;
    const rawSuggested = Math.max(projectedVoyageNeed - availableAtArrival, 0);

    let suggestedOrder = rawSuggested;
    let parCapApplied = false;
    let parCapLimit = 0;

    if (Number(voyageDays) === 14 && parLevel > 0) {
      parCapLimit = parLevel * 1.1;

      if (suggestedOrder > parCapLimit) {
        suggestedOrder = parCapLimit;
        parCapApplied = true;
      }
    }

    const suggestedQty = Number(suggestedOrder || 0);
    const orderedQty = Number(orderedByShip || 0);
    const orderDifference = orderedQty - suggestedQty;

    let orderDifferencePercent = 0;
    let orderComparisonStatus = "green";
    let orderComparisonLabel = "Within +/- 10%";

    if (suggestedQty > 0) {
      orderDifferencePercent = (orderDifference / suggestedQty) * 100;

      if (orderDifferencePercent > 10) {
        orderComparisonStatus = "red";
        orderComparisonLabel = "Ordered over suggested by more than 10%";
      } else if (orderDifferencePercent < -10) {
        orderComparisonStatus = "blue";
        orderComparisonLabel = "Ordered under suggested by more than 10%";
      }
    } else if (orderedQty > 0) {
      orderDifferencePercent = 100;
      orderComparisonStatus = "red";
      orderComparisonLabel = "Ordered but suggested quantity is 0";
    }

    let alertType = "normal";
    let alertLabel = "Review";

    if (stock === 0 && pastConsumption === 0) {
      alertType = "no-stock-no-consumption";
      alertLabel = "No stock / no consumption";
    } else if (stock > 0 && pastConsumption === 0) {
      alertType = "stock-no-consumption";
      alertLabel = "Stock but no consumption";
    } else if (suggestedOrder > 0) {
      alertType = "needs-order";
      alertLabel = "Needs order";
    }

    const item = {
      excelOrder: orderRows.length + 1,
      excelRow,
      code,
      product,
      unit,
      stock,
      futureOrders,
      parLevel,
      pastConsumption,
      historicalSailorDays,
      averagePerDay,
      usageUntilArrival,
      availableAtArrival,
      projectedVoyageNeed,
      rawSuggested,
      suggestedOrder,
      orderedByShip,
      orderedQty,
      suggestedQty,
      orderDifference,
      orderDifferencePercent,
      orderComparisonStatus,
      orderComparisonLabel,
      parCapApplied,
      parCapLimit,
      ...consumptionIncreaseMetrics,
      alertType,
      alertLabel,
    };

    orderRows.push(item);

    if (code) {
      orderByCode[cleanText(code)] = item;
      orderByCode[getFmlOrderCodeKey(code)] = item;
    }
  }

  await yieldToBrowser();

  const fmlRows = parseFmlRows(workbook);
  const fmlOrderedNotFmlRows = buildOrderedNotInFmlRows(workbook, orderRows);

  const counts = {
    totalItems: orderRows.length,
    itemsNeedingOrder: orderRows.filter((row) => row.suggestedOrder > 0).length,
    noConsumptionItems: orderRows.filter((row) => row.pastConsumption === 0).length,
    noStockItems: orderRows.filter((row) => row.stock === 0).length,
    parCapsApplied: orderRows.filter((row) => row.parCapApplied).length,
    negativeArrival: orderRows.filter((row) => row.availableAtArrival < 0).length,
    fmlNotUsed: 0,
    fmlRunningLow: 0,
    fmlOrderedNotFml: fmlOrderedNotFmlRows.length,
  };

  return {
    workbook,
    orderRows,
    orderByCode,
    fmlRows,
    fmlNotUsedRows: [],
    fmlRunningLowRows: [],
    fmlOrderedNotFmlRows,
    meta: {
      fileName: file.name,
      sheetName,
      shipNameRaw: String(shipNameRaw || ""),
      shipCode,
      orderDate,
      arrivalDate,
      sailors,
      voyageDays,
      daysUntilArrival,
      ...counts,
    },
  };
};

const buildFmlReportAsync = async ({
  mode,
  fmlRows,
  orderRows,
  orderByCode,
  templateEntries,
  shipCode,
  onProgress,
}) => {
  const result = [];
  const templateMatchCache = new Map();

  const findOrderMatch = (fmlItem) => {
    const fmlCodeKey = fmlItem.code ? cleanText(fmlItem.code) : "";
    const normalizedCodeKey = getFmlOrderCodeKey(fmlItem.code);

    if (fmlCodeKey && orderByCode[fmlCodeKey]) {
      return orderByCode[fmlCodeKey];
    }

    if (normalizedCodeKey && orderByCode[normalizedCodeKey]) {
      return orderByCode[normalizedCodeKey];
    }

    return orderRows.find((row) => productNamesMatch(fmlItem.product, row.product));
  };

  for (let i = 0; i < fmlRows.length; i += 1) {
    if (i > 0 && i % 20 === 0) {
      if (onProgress) {
        onProgress("Preparing FML report... " + i + " of " + fmlRows.length);
      }

      await yieldToBrowser();
    }

    const fmlItem = fmlRows[i];
    const orderMatch = findOrderMatch(fmlItem);

    if (!orderMatch) continue;

    const cacheKey = compactText(fmlItem.product) + "|" + (shipCode || "");
    let templateMatches = templateMatchCache.get(cacheKey);

    if (!templateMatches) {
      templateMatches = findTemplateMatches(templateEntries, fmlItem.product, shipCode);
      templateMatchCache.set(cacheKey, templateMatches);
    }

    if (!templateMatches.length) continue;

    const futureZero = Number(orderMatch.futureOrders || 0) === 0;
    const pastZero = Number(orderMatch.pastConsumption || 0) === 0;
    const pastPositive = Number(orderMatch.pastConsumption || 0) > 0;
    const runningLow =
      Number(orderMatch.availableAtArrival || 0) <= Number(orderMatch.averagePerDay || 0) &&
      Number(orderMatch.averagePerDay || 0) > 0;

    if (mode === "notUsed" && (!futureZero || !pastZero)) continue;
    if (mode === "runningLow" && (!futureZero || !pastPositive || !runningLow)) continue;

    result.push({
      id: mode + "-" + fmlItem.fmlRow + "-" + (orderMatch.code || orderMatch.product),
      fmlRow: fmlItem.fmlRow,
      code: orderMatch.code || fmlItem.code || "",
      product: orderMatch.product || fmlItem.product,
      unit: orderMatch.unit || "N/A",
      stock: orderMatch.stock,
      futureOrders: orderMatch.futureOrders,
      pastConsumption: orderMatch.pastConsumption,
      averagePerDay: orderMatch.averagePerDay,
      availableAtArrival: orderMatch.availableAtArrival,
      suggestedOrder: orderMatch.suggestedOrder,
      venues: fmlItem.venues,
      venuesText: fmlItem.venuesText,
      templateMatches,
      templateLocation: templateMatches[0]?.location || "",
      scopeText: templateMatches[0]?.scopeText || "",
      reason:
        mode === "notUsed"
          ? "FML item has venues, matches template, but has 0 future order and 0 past consumption."
          : "FML item matches template, has no future order, and will be low by arrival day.",
    });
  }

  if (onProgress) {
    onProgress("FML report ready. " + result.length + " records found.");
  }

  return result;
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

const printRows = (title, rows, columns) => {
  if (!rows.length) {
    window.alert("No rows to print.");
    return;
  }

  const headerHtml = columns.map((column) => "<th>" + column.label + "</th>").join("");
  const rowsHtml = rows
    .map(
      (row) =>
        "<tr>" +
        columns.map((column) => "<td>" + String(row[column.key] ?? "") + "</td>").join("") +
        "</tr>"
    )
    .join("");

  const html =
    "<html><head><title>" +
    title +
    "</title><style>body{font-family:Arial,sans-serif;padding:24px;}table{width:100%;border-collapse:collapse;margin-top:20px;}th,td{border:1px solid #ccc;padding:7px;text-align:left;font-size:12px;}th{background:#f2f2f2;}</style></head><body><h1>" +
    title +
    "</h1><table><thead><tr>" +
    headerHtml +
    "</tr></thead><tbody>" +
    rowsHtml +
    "</tbody></table></body></html>";

  const printWindow = window.open("", "_blank");
  if (!printWindow) return;

  printWindow.document.write(html);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
};

export default function GenerateNextOrder({
  styles,
  userShip,
  onBack,
  logUsageEvent = () => {},
}) {
  const [templateEntries, setTemplateEntries] = useState([]);
  const [templateStatus, setTemplateStatus] = useState("Loading attached ERP template...");
  const [nextOrderRows, setNextOrderRows] = useState([]);
  const [orderByCode, setOrderByCode] = useState({});
  const [fmlSourceRows, setFmlSourceRows] = useState([]);
  const [fmlNotUsedRows, setFmlNotUsedRows] = useState([]);
  const [fmlRunningLowRows, setFmlRunningLowRows] = useState([]);
  const [fmlOrderedNotFmlRows, setFmlOrderedNotFmlRows] = useState([]);
  const [nextOrderMeta, setNextOrderMeta] = useState({});
  const [nextOrderFileName, setNextOrderFileName] = useState("");
  const [nextOrderSearch, setNextOrderSearch] = useState("");
  const [fmlSearch, setFmlSearch] = useState("");
  const [fmlLowSearch, setFmlLowSearch] = useState("");
  const [fmlOrderedNotFmlSearch, setFmlOrderedNotFmlSearch] = useState("");
  const [nextOrderFilter, setNextOrderFilter] = useState("all");
  const [nextOrderView, setNextOrderView] = useState("order");
  const [nextOrderLoading, setNextOrderLoading] = useState(false);
  const [fmlReportLoading, setFmlReportLoading] = useState(false);
  const [fmlReportMessage, setFmlReportMessage] = useState("");
  const [fmlNotUsedPrepared, setFmlNotUsedPrepared] = useState(false);
  const [fmlRunningLowPrepared, setFmlRunningLowPrepared] = useState(false);
  const [nextOrderMessage, setNextOrderMessage] = useState("");
  const [reportDisplayLimit, setReportDisplayLimit] = useState(REPORT_RENDER_BATCH);

  useEffect(() => {
    const loadDefaultTemplate = async () => {
      try {
        const response = await fetch("/template.xlsx");

        if (!response.ok) {
          setTemplateStatus("Attached ERP template not found. Upload one if needed.");
          return;
        }

        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        const entries = parseTemplateWorkbook(workbook);

        setTemplateEntries(entries);
        setTemplateStatus("Attached ERP template loaded. " + entries.length + " item lines found.");
      } catch {
        setTemplateStatus("Could not load attached ERP template.");
      }
    };

    loadDefaultTemplate();
  }, []);

  useEffect(() => {
    setReportDisplayLimit(REPORT_RENDER_BATCH);
  }, [
    nextOrderView,
    nextOrderSearch,
    nextOrderFilter,
    fmlSearch,
    fmlLowSearch,
    fmlOrderedNotFmlSearch,
  ]);

  const uploadErpTemplate = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (loadEvent) => {
      try {
        const workbook = XLSX.read(loadEvent.target.result, { type: "binary" });
        const entries = parseTemplateWorkbook(workbook);

        setTemplateEntries(entries);
        setTemplateStatus("Custom ERP template loaded. " + entries.length + " item lines found.");

        logUsageEvent("next_order_erp_template_uploaded", {
          module: "generate_next_order",
          fileName: file.name,
          itemLines: entries.length,
        });
      } catch {
        setTemplateStatus("Could not load custom ERP template.");
      }
    };

    reader.readAsBinaryString(file);
  };

  const resetFmlReports = () => {
    setFmlNotUsedRows([]);
    setFmlRunningLowRows([]);
    setFmlOrderedNotFmlRows([]);
    setFmlNotUsedPrepared(false);
    setFmlRunningLowPrepared(false);
  };

  const uploadNextOrderFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setNextOrderLoading(true);
    setFmlReportLoading(false);
    setNextOrderMessage("Preparing next order report...");
    setFmlReportMessage("");

    try {
      await yieldToBrowser();

      const parsed = await parseOrderFile(file);

      setNextOrderRows(parsed.orderRows);
      setOrderByCode(parsed.orderByCode || {});
      setFmlSourceRows(parsed.fmlRows || []);
      setFmlNotUsedRows([]);
      setFmlRunningLowRows([]);
      setFmlOrderedNotFmlRows(parsed.fmlOrderedNotFmlRows || []);
      setFmlNotUsedPrepared(false);
      setFmlRunningLowPrepared(false);
      setNextOrderMeta(parsed.meta);
      setNextOrderFileName(file.name);
      setNextOrderSearch("");
      setFmlSearch("");
      setFmlLowSearch("");
      setFmlOrderedNotFmlSearch("");
      setNextOrderFilter("all");
      setNextOrderView("order");
      setNextOrderMessage(
        "Order file loaded. " +
          parsed.meta.totalItems +
          " product rows found. " +
          parsed.meta.itemsNeedingOrder +
          " need order. " +
          parsed.meta.fmlOrderedNotFml +
          " ordered item(s) are not in FML. Calculated using B2/B3 days to arrival, B5 sailors, B6 voyage days, F:N future orders, Y ordered by ship, AI:AN past consumption, and Q par where applicable. FML reports are preparing in the background."
      );

      logUsageEvent("next_order_file_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        ship: parsed.meta.shipCode,
        totalItems: parsed.meta.totalItems,
        fmlRows: (parsed.fmlRows || []).length,
        orderedNotInFml: parsed.meta.fmlOrderedNotFml,
      });

      window.setTimeout(() => {
        (async () => {
          try {
            setFmlReportLoading(true);
            setFmlReportMessage("Preparing FML reports in the background...");

            const notUsedRows = await buildFmlReportAsync({
              mode: "notUsed",
              fmlRows: parsed.fmlRows || [],
              orderRows: parsed.orderRows || [],
              orderByCode: parsed.orderByCode || {},
              templateEntries,
              shipCode: parsed.meta.shipCode,
              onProgress: setFmlReportMessage,
            });

            setFmlNotUsedRows(notUsedRows);
            setFmlNotUsedPrepared(true);
            setNextOrderMeta((prev) => ({ ...prev, fmlNotUsed: notUsedRows.length }));

            const runningLowRows = await buildFmlReportAsync({
              mode: "runningLow",
              fmlRows: parsed.fmlRows || [],
              orderRows: parsed.orderRows || [],
              orderByCode: parsed.orderByCode || {},
              templateEntries,
              shipCode: parsed.meta.shipCode,
              onProgress: setFmlReportMessage,
            });

            setFmlRunningLowRows(runningLowRows);
            setFmlRunningLowPrepared(true);
            setNextOrderMeta((prev) => ({ ...prev, fmlRunningLow: runningLowRows.length }));
            setFmlReportMessage(
              "FML reports ready. Not used: " +
                notUsedRows.length +
                ", running low: " +
                runningLowRows.length +
                ", ordered not in FML: " +
                (parsed.fmlOrderedNotFmlRows || []).length +
                "."
            );
          } catch (error) {
            setFmlReportMessage(error?.message || "Could not prepare FML reports.");
          } finally {
            setFmlReportLoading(false);
          }
        })();
      }, 50);
    } catch (error) {
      setNextOrderRows([]);
      setOrderByCode({});
      setFmlSourceRows([]);
      resetFmlReports();
      setNextOrderMessage(error?.message || "Could not prepare order file.");
    } finally {
      setNextOrderLoading(false);
      event.target.value = "";
    }
  };

  const prepareFmlReport = async (mode) => {
    if (!fmlSourceRows.length || !nextOrderRows.length) {
      setFmlReportMessage("Upload the latest order file first.");
      return;
    }

    if (mode === "notUsed" && fmlNotUsedPrepared) return;
    if (mode === "runningLow" && fmlRunningLowPrepared) return;

    setFmlReportLoading(true);
    setFmlReportMessage("Preparing FML report...");

    try {
      await yieldToBrowser();

      const rows = await buildFmlReportAsync({
        mode,
        fmlRows: fmlSourceRows,
        orderRows: nextOrderRows,
        orderByCode,
        templateEntries,
        shipCode: nextOrderMeta.shipCode,
        onProgress: setFmlReportMessage,
      });

      if (mode === "notUsed") {
        setFmlNotUsedRows(rows);
        setFmlNotUsedPrepared(true);
        setNextOrderMeta((prev) => ({ ...prev, fmlNotUsed: rows.length }));
      } else {
        setFmlRunningLowRows(rows);
        setFmlRunningLowPrepared(true);
        setNextOrderMeta((prev) => ({ ...prev, fmlRunningLow: rows.length }));
      }
    } catch (error) {
      setFmlReportMessage(error?.message || "Could not prepare FML report.");
    } finally {
      setFmlReportLoading(false);
    }
  };

  const filteredNextOrderRows = useMemo(() => {
    const query = nextOrderSearch.toLowerCase().trim();

    const rows = nextOrderRows.filter((row) => {
      const filterOk =
        nextOrderFilter === "all" ||
        (nextOrderFilter === "needs" && row.suggestedOrder > 0) ||
        (nextOrderFilter === "runningLow" && Number(row.availableAtArrival || 0) < 0) ||
        (nextOrderFilter === "noConsumption" && row.pastConsumption === 0) ||
        (nextOrderFilter === "noStock" && row.stock === 0);

      if (!filterOk) return false;
      if (!query) return true;

      return [
        row.product,
        row.code,
        row.unit,
        row.alertLabel,
        String(row.excelRow),
        String(row.excelOrder),
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    });

    if (nextOrderFilter === "runningLow") {
      return [...rows].sort((a, b) => {
        const aArrival = Number(a.availableAtArrival || 0);
        const bArrival = Number(b.availableAtArrival || 0);

        if (aArrival !== bArrival) return aArrival - bArrival;

        return Number(a.excelOrder || 0) - Number(b.excelOrder || 0);
      });
    }

    return rows;
  }, [nextOrderRows, nextOrderSearch, nextOrderFilter]);

  const orderedVsSuggestedRows = useMemo(() => {
    const query = nextOrderSearch.toLowerCase().trim();

    return nextOrderRows
      .map((row) => {
        const orderedQty = Number(row.orderedByShip || 0);
        const suggestedQty = Number(row.suggestedOrder || 0);
        const orderDifference = orderedQty - suggestedQty;

        let orderDifferencePercent = 0;
        let orderComparisonStatus = "green";
        let orderComparisonLabel = "Within +/- 10%";

        if (suggestedQty > 0) {
          orderDifferencePercent = (orderDifference / suggestedQty) * 100;

          if (orderDifferencePercent > 10) {
            orderComparisonStatus = "red";
            orderComparisonLabel = "Ordered over suggested by more than 10%";
          } else if (orderDifferencePercent < -10) {
            orderComparisonStatus = "blue";
            orderComparisonLabel = "Ordered under suggested by more than 10%";
          }
        } else if (orderedQty > 0) {
          orderDifferencePercent = 100;
          orderComparisonStatus = "red";
          orderComparisonLabel = "Ordered but suggested quantity is 0";
        }

        return {
          ...row,
          orderedQty,
          suggestedQty,
          orderDifference,
          orderDifferencePercent,
          orderComparisonStatus,
          orderComparisonLabel,
        };
      })
      .filter((row) => Number(row.orderedQty || 0) > 0 || Number(row.suggestedQty || 0) > 0)
      .filter((row) => {
        if (!query) return true;

        return [
          row.product,
          row.code,
          row.unit,
          row.orderComparisonLabel,
          String(row.excelRow),
          String(row.excelOrder),
          String(row.orderedQty),
          String(row.suggestedQty),
          String(row.orderDifference),
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort((a, b) => {
        const statusRank = {
          red: 0,
          blue: 1,
          green: 2,
        };

        const statusDiff =
          (statusRank[a.orderComparisonStatus] ?? 9) -
          (statusRank[b.orderComparisonStatus] ?? 9);

        if (statusDiff !== 0) return statusDiff;

        return Math.abs(Number(b.orderDifferencePercent || 0)) - Math.abs(Number(a.orderDifferencePercent || 0));
      });
  }, [nextOrderRows, nextOrderSearch]);

  const oneVoyageConsumptionIncreaseRows = useMemo(() => {
    const query = nextOrderSearch.toLowerCase().trim();

    return nextOrderRows
      .filter((row) => row.oneVoyageIncreaseFlag)
      .filter((row) => {
        if (!query) return true;

        return [
          row.product,
          row.code,
          row.unit,
          String(row.excelRow),
          String(row.latestOneVoyageDaily),
          String(row.previousFiveVoyageDailyAverage),
          String(row.oneVoyageIncreasePercent),
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort(
        (a, b) =>
          Number(b.oneVoyageIncreasePercent || 0) -
          Number(a.oneVoyageIncreasePercent || 0)
      );
  }, [nextOrderRows, nextOrderSearch]);

  const twoVoyageConsumptionIncreaseRows = useMemo(() => {
    const query = nextOrderSearch.toLowerCase().trim();

    return nextOrderRows
      .filter((row) => row.twoVoyageIncreaseFlag)
      .filter((row) => {
        if (!query) return true;

        return [
          row.product,
          row.code,
          row.unit,
          String(row.excelRow),
          String(row.latestTwoVoyageDailyAverage),
          String(row.previousFourVoyageDailyAverage),
          String(row.twoVoyageIncreasePercent),
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort(
        (a, b) =>
          Number(b.twoVoyageIncreasePercent || 0) -
          Number(a.twoVoyageIncreasePercent || 0)
      );
  }, [nextOrderRows, nextOrderSearch]);

  const filteredFmlNotUsedRows = useMemo(() => {
    const query = fmlSearch.toLowerCase().trim();
    if (!query) return fmlNotUsedRows;

    return fmlNotUsedRows.filter((row) =>
      [row.product, row.code, row.unit, row.venuesText, row.templateLocation, row.scopeText]
        .join(" ")
        .toLowerCase()
        .includes(query)
    );
  }, [fmlNotUsedRows, fmlSearch]);

  const filteredFmlRunningLowRows = useMemo(() => {
    const query = fmlLowSearch.toLowerCase().trim();
    if (!query) return fmlRunningLowRows;

    return fmlRunningLowRows.filter((row) =>
      [row.product, row.code, row.unit, row.venuesText, row.templateLocation, row.scopeText]
        .join(" ")
        .toLowerCase()
        .includes(query)
    );
  }, [fmlRunningLowRows, fmlLowSearch]);

  const filteredFmlOrderedNotFmlRows = useMemo(() => {
    const query = fmlOrderedNotFmlSearch.toLowerCase().trim();

    if (!query) return fmlOrderedNotFmlRows;

    return fmlOrderedNotFmlRows.filter((row) =>
      [
        row.product,
        row.code,
        row.unit,
        row.excelRow,
        row.matchType,
        row.reason,
        String(row.orderedByShip),
        String(row.suggestedOrder),
      ]
        .join(" ")
        .toLowerCase()
        .includes(query)
    );
  }, [fmlOrderedNotFmlRows, fmlOrderedNotFmlSearch]);

  const orderedVsSuggestedExportRows = orderedVsSuggestedRows.map((row, index) => ({
    Number: index + 1,
    ExcelRow: row.excelRow,
    Code: row.code,
    Product: row.product,
    UM: row.unit,
    CurrentQuantityOnHand: row.stock,
    UpcomingOrder: row.futureOrders,
    DailyConsumption: row.averagePerDay,
    OrderedByShipColumnY: row.orderedQty,
    SuggestedOrder: row.suggestedQty,
    Difference: row.orderDifference,
    DifferencePercent: row.orderDifferencePercent,
    Status: row.orderComparisonLabel,
  }));

  const getConsumptionIncreaseExportRows = (rows, modeLabel) =>
    rows.map((row, index) => ({
      Number: index + 1,
      Mode: modeLabel,
      ExcelRow: row.excelRow,
      Code: row.code,
      Product: row.product,
      UM: row.unit,
      StockOnHand: row.stock,
      FutureOrders: row.futureOrders,
      PastConsumptionTotal: row.pastConsumption,
      LatestOneVoyageColumn: row.latestOneVoyageColumn,
      LatestOneVoyageQty: row.latestOneVoyageQty,
      LatestOneVoyageDaily: row.latestOneVoyageDaily,
      PreviousFiveVoyageDailyAverage: row.previousFiveVoyageDailyAverage,
      OneVoyageIncreasePercent: row.oneVoyageIncreasePercent,
      LatestTwoVoyageColumns: row.latestTwoVoyageColumns,
      LatestTwoVoyageQty: row.latestTwoVoyageQty,
      LatestTwoVoyageDailyAverage: row.latestTwoVoyageDailyAverage,
      PreviousFourVoyageDailyAverage: row.previousFourVoyageDailyAverage,
      TwoVoyageIncreasePercent: row.twoVoyageIncreasePercent,
    }));

  const oneVoyageConsumptionIncreaseExportRows = getConsumptionIncreaseExportRows(
    oneVoyageConsumptionIncreaseRows,
    "1 Voyage Increase"
  );

  const twoVoyageConsumptionIncreaseExportRows = getConsumptionIncreaseExportRows(
    twoVoyageConsumptionIncreaseRows,
    "2 Voyage Increase"
  );

  const nextOrderExportRows = filteredNextOrderRows.map((row) => ({
    ExcelOrder: row.excelOrder,
    ExcelRow: row.excelRow,
    Code: row.code,
    Product: row.product,
    UM: row.unit,
    Stock: row.stock,
    FutureOrders: row.futureOrders,
    ParLevel: row.parLevel,
    PastConsumption: row.pastConsumption,
    HistoricalSailorDays: row.historicalSailorDays,
    AveragePerDay: row.averagePerDay,
    UsageUntilArrival: row.usageUntilArrival,
    AvailableAtArrival: row.availableAtArrival,
    ProjectedVoyageNeed: row.projectedVoyageNeed,
    RawSuggested: row.rawSuggested,
    SuggestedOrder: row.suggestedOrder,
    ParCapApplied: row.parCapApplied ? "Yes" : "No",
    Alert: row.alertLabel,
  }));

  const fmlExportRows = (rows) =>
    rows.map((row, index) => ({
      Number: index + 1,
      FmlRow: row.fmlRow,
      Code: row.code,
      Product: row.product,
      UM: row.unit,
      Stock: row.stock,
      FutureOrders: row.futureOrders,
      PastConsumption: row.pastConsumption,
      AveragePerDay: row.averagePerDay,
      AvailableAtArrival: row.availableAtArrival,
      SuggestedOrder: row.suggestedOrder,
      Venues: row.venuesText,
      TemplateLocation: row.templateLocation,
      TemplateScope: row.scopeText,
      Reason: row.reason,
    }));

  const fmlOrderedNotFmlExportRows = filteredFmlOrderedNotFmlRows.map((row, index) => ({
    Number: index + 1,
    ExcelRow: row.excelRow,
    Code: row.code,
    Product: row.product,
    UM: row.unit,
    Stock: row.stock,
    FutureOrders: row.futureOrders,
    OrderedByShipColumnY: row.orderedByShip,
    SuggestedOrder: row.suggestedOrder,
    PastConsumption: row.pastConsumption,
    AveragePerDay: row.averagePerDay,
    AvailableAtArrival: row.availableAtArrival,
    MatchType: row.matchType,
    Reason: row.reason,
  }));

  const countNeedsOrder = nextOrderRows.filter((row) => row.suggestedOrder > 0).length;
  const countRunningLowBeforeLoading = nextOrderRows.filter((row) => Number(row.availableAtArrival || 0) < 0).length;
  const countNoConsumption = nextOrderRows.filter((row) => row.pastConsumption === 0).length;
  const countNoStock = nextOrderRows.filter((row) => row.stock === 0).length;

  const visibleNextOrderRows = filteredNextOrderRows.slice(0, reportDisplayLimit);
  const visibleOrderedVsSuggestedRows = orderedVsSuggestedRows.slice(0, reportDisplayLimit);
  const visibleFmlNotUsedRows = filteredFmlNotUsedRows.slice(0, reportDisplayLimit);
  const visibleFmlRunningLowRows = filteredFmlRunningLowRows.slice(0, reportDisplayLimit);
  const visibleFmlOrderedNotFmlRows = filteredFmlOrderedNotFmlRows.slice(0, reportDisplayLimit);
  const visibleOneVoyageConsumptionIncreaseRows = oneVoyageConsumptionIncreaseRows.slice(0, reportDisplayLimit);
  const visibleTwoVoyageConsumptionIncreaseRows = twoVoyageConsumptionIncreaseRows.slice(0, reportDisplayLimit);

  const hasMoreNextOrderRows = filteredNextOrderRows.length > visibleNextOrderRows.length;
  const hasMoreOrderedVsSuggestedRows = orderedVsSuggestedRows.length > visibleOrderedVsSuggestedRows.length;
  const hasMoreFmlNotUsedRows = filteredFmlNotUsedRows.length > visibleFmlNotUsedRows.length;
  const hasMoreFmlRunningLowRows = filteredFmlRunningLowRows.length > visibleFmlRunningLowRows.length;
  const hasMoreFmlOrderedNotFmlRows = filteredFmlOrderedNotFmlRows.length > visibleFmlOrderedNotFmlRows.length;
  const hasMoreOneVoyageConsumptionIncreaseRows = oneVoyageConsumptionIncreaseRows.length > visibleOneVoyageConsumptionIncreaseRows.length;
  const hasMoreTwoVoyageConsumptionIncreaseRows = twoVoyageConsumptionIncreaseRows.length > visibleTwoVoyageConsumptionIncreaseRows.length;

  const orderSheetShip = nextOrderMeta.shipCode || userShip || "";

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Product Options
          </button>

          <div style={styles.shipBadge}>🚢 {orderSheetShip || userShip || "Ship"}</div>
        </div>
      </header>

      <div style={styles.viewModeBox}>
        <button
          style={{ ...styles.viewModeButton, ...(nextOrderView === "order" ? styles.viewModeButtonActive : {}) }}
          onClick={() => setNextOrderView("order")}
        >
          🛒 Next Order ({nextOrderRows.length})
        </button>

        <button
          style={{
            ...styles.viewModeButton,
            ...(nextOrderView === "orderedVsSuggested" ? styles.viewModeButtonActive : {}),
          }}
          onClick={() => setNextOrderView("orderedVsSuggested")}
        >
          📊 Ordered vs Suggested ({orderedVsSuggestedRows.length})
        </button>

        <button
          style={{
            ...styles.viewModeButton,
            ...(nextOrderView === "increase1" ? styles.viewModeButtonActive : {}),
          }}
          onClick={() => setNextOrderView("increase1")}
        >
          📈 Increase 1 Voyage ({oneVoyageConsumptionIncreaseRows.length})
        </button>

        <button
          style={{
            ...styles.viewModeButton,
            ...(nextOrderView === "increase2" ? styles.viewModeButtonActive : {}),
          }}
          onClick={() => setNextOrderView("increase2")}
        >
          📈 Increase 2 Voyages ({twoVoyageConsumptionIncreaseRows.length})
        </button>

        <button
          style={{ ...styles.viewModeButton, ...(nextOrderView === "fml" ? styles.viewModeButtonActive : {}) }}
          onClick={() => {
            setNextOrderView("fml");
            prepareFmlReport("notUsed");
          }}
        >
          📘 FML Not Ordered / Not Used ({fmlNotUsedRows.length})
        </button>

        <button
          style={{ ...styles.viewModeButton, ...(nextOrderView === "fmllow" ? styles.viewModeButtonActive : {}) }}
          onClick={() => {
            setNextOrderView("fmllow");
            prepareFmlReport("runningLow");
          }}
        >
          ⚠️ FML Running Low ({fmlRunningLowRows.length})
        </button>

        <button
          style={{ ...styles.viewModeButton, ...(nextOrderView === "fmlordered" ? styles.viewModeButtonActive : {}) }}
          onClick={() => setNextOrderView("fmlordered")}
        >
          🧾 Ordered Not In FML ({fmlOrderedNotFmlRows.length})
        </button>
      </div>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🛒 Generate Next Order</h2>

          <label style={styles.label}>Optional: replace attached ERP Food ordering template</label>
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadErpTemplate} style={styles.fileInput} />

          <label style={styles.label}>Step 1: Upload latest order file</label>
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadNextOrderFile} style={styles.fileInput} />

          <div style={styles.infoBox}>
            <div>📋 ERP template: <strong>{templateStatus}</strong></div>
            <div>📄 Order file: <strong>{nextOrderFileName || "Not uploaded"}</strong></div>
            <div>🚢 Ship from B1: <strong>{getShipDisplayName(nextOrderMeta.shipCode) || "Not loaded"}</strong></div>
            <div>📅 Order day B2: <strong>{formatDate(nextOrderMeta.orderDate)}</strong></div>
            <div>📦 Arrival day B3: <strong>{formatDate(nextOrderMeta.arrivalDate)}</strong></div>
            <div>👥 Sailors B5: <strong>{formatQty(nextOrderMeta.sailors)}</strong></div>
            <div>🗓️ Voyage days B6: <strong>{formatQty(nextOrderMeta.voyageDays)}</strong></div>
            <div>⏱️ Days until arrival: <strong>{formatQty(nextOrderMeta.daysUntilArrival)}</strong></div>
            {nextOrderMessage && (
              <div style={{ color: nextOrderMessage.includes("Could not") ? "#b00020" : "#555" }}>
                {nextOrderMessage}
              </div>
            )}
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>⚙️ Report Actions</h2>

          {nextOrderView === "order" && (
            <>
              <input
                placeholder="Search item, code, U/M, alert, or Excel row..."
                value={nextOrderSearch}
                onChange={(event) => setNextOrderSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.viewModeBox}>
                <button
                  style={{ ...styles.viewModeButton, ...(nextOrderFilter === "all" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setNextOrderFilter("all")}
                >
                  All ({nextOrderRows.length})
                </button>

                <button
                  style={{ ...styles.viewModeButton, ...(nextOrderFilter === "needs" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setNextOrderFilter("needs")}
                >
                  Needs to Order ({countNeedsOrder})
                </button>

                <button
                  style={{ ...styles.viewModeButton, ...(nextOrderFilter === "runningLow" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setNextOrderFilter("runningLow")}
                >
                  ⚠️ Running Low Before Loading ({countRunningLowBeforeLoading})
                </button>

                <button
                  style={{ ...styles.viewModeButton, ...(nextOrderFilter === "noConsumption" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setNextOrderFilter("noConsumption")}
                >
                  No Consumption ({countNoConsumption})
                </button>

                <button
                  style={{ ...styles.viewModeButton, ...(nextOrderFilter === "noStock" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setNextOrderFilter("noStock")}
                >
                  No Stock ({countNoStock})
                </button>
              </div>

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() =>
                    printRows("Generate Next Order", nextOrderExportRows, [
                      { key: "ExcelOrder", label: "#" },
                      { key: "ExcelRow", label: "Row" },
                      { key: "Code", label: "Code" },
                      { key: "Product", label: "Product" },
                      { key: "UM", label: "U/M" },
                      { key: "Stock", label: "Stock" },
                      { key: "FutureOrders", label: "Future" },
                      { key: "AvailableAtArrival", label: "At Arrival" },
                      { key: "SuggestedOrder", label: "Order" },
                      { key: "Alert", label: "Alert" },
                    ])
                  }
                >
                  🖨️ Print
                </button>

                <button
                  style={styles.primaryButton}
                  onClick={() => exportRowsToExcel(nextOrderExportRows, "Next Order", "next-order.xlsx")}
                >
                  📥 Export Excel
                </button>
              </div>
            </>
          )}

          {nextOrderView === "orderedVsSuggested" && (
            <>
              <input
                placeholder="Search ordered vs suggested item, code, U/M, or status..."
                value={nextOrderSearch}
                onChange={(event) => setNextOrderSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.infoBox}>
                <div>📘 Ordered by ship source: <strong>Column Y</strong></div>
                <div>📦 Current quantity on hand: <strong>Stock column D</strong></div>
                <div>🚚 Upcoming order: <strong>Future orders F:N</strong></div>
                <div>📊 Daily consumption: <strong>average daily usage from past consumption</strong></div>
                <div>🔴 Red: <strong>ordered more than suggested by over 10%</strong></div>
                <div>🔵 Blue: <strong>ordered less than suggested by over 10%</strong></div>
                <div>🟢 Green: <strong>within -10% to +10%</strong></div>
              </div>

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() =>
                    printRows("Ordered vs Suggested", orderedVsSuggestedExportRows, [
                      { key: "Number", label: "#" },
                      { key: "ExcelRow", label: "Row" },
                      { key: "Code", label: "Code" },
                      { key: "Product", label: "Product" },
                      { key: "UM", label: "U/M" },
                      { key: "CurrentQuantityOnHand", label: "On Hand" },
                      { key: "UpcomingOrder", label: "Upcoming Order" },
                      { key: "DailyConsumption", label: "Daily Consumption" },
                      { key: "OrderedByShipColumnY", label: "Ordered Y" },
                      { key: "SuggestedOrder", label: "Suggested" },
                      { key: "Difference", label: "Difference" },
                      { key: "DifferencePercent", label: "Diff %" },
                      { key: "Status", label: "Status" },
                    ])
                  }
                >
                  🖨️ Print
                </button>

                <button
                  style={styles.primaryButton}
                  onClick={() =>
                    exportRowsToExcel(
                      orderedVsSuggestedExportRows,
                      "Ordered vs Suggested",
                      "ordered-vs-suggested.xlsx"
                    )
                  }
                >
                  📥 Export Excel
                </button>
              </div>
            </>
          )}


          {nextOrderView === "increase1" && (
            <>
              <input
                placeholder="Search 1-voyage consumption increase item, code, U/M..."
                value={nextOrderSearch}
                onChange={(event) => setNextOrderSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.infoBox}>
                <div>📈 Report: <strong>Consumption Increase - Last 1 Voyage</strong></div>
                <div>📘 Source: <strong>AJ:AO historical consumption columns</strong></div>
                <div>📊 Logic: <strong>Latest voyage AO vs average of previous voyages AJ:AN</strong></div>
                <div>⚠️ Shows items with normalized daily consumption increase of <strong>25% or more</strong>.</div>
              </div>

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() =>
                    printRows("Consumption Increase - 1 Voyage", oneVoyageConsumptionIncreaseExportRows, [
                      { key: "Number", label: "#" },
                      { key: "ExcelRow", label: "Row" },
                      { key: "Code", label: "Code" },
                      { key: "Product", label: "Product" },
                      { key: "UM", label: "U/M" },
                      { key: "LatestOneVoyageDaily", label: "Latest Daily" },
                      { key: "PreviousFiveVoyageDailyAverage", label: "Previous Avg" },
                      { key: "OneVoyageIncreasePercent", label: "Increase %" },
                    ])
                  }
                >
                  🖨️ Print
                </button>

                <button
                  style={styles.primaryButton}
                  onClick={() =>
                    exportRowsToExcel(
                      oneVoyageConsumptionIncreaseExportRows,
                      "Increase 1 Voyage",
                      "consumption-increase-1-voyage.xlsx"
                    )
                  }
                >
                  📥 Export Excel
                </button>
              </div>
            </>
          )}

          {nextOrderView === "increase2" && (
            <>
              <input
                placeholder="Search 2-voyage consumption increase item, code, U/M..."
                value={nextOrderSearch}
                onChange={(event) => setNextOrderSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.infoBox}>
                <div>📈 Report: <strong>Consumption Increase - Last 2 Voyages</strong></div>
                <div>📘 Source: <strong>AJ:AO historical consumption columns</strong></div>
                <div>📊 Logic: <strong>Latest two voyages AN:AO vs average of previous voyages AJ:AM</strong></div>
                <div>⚠️ Shows items with normalized daily consumption increase of <strong>25% or more</strong>.</div>
              </div>

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() =>
                    printRows("Consumption Increase - 2 Voyages", twoVoyageConsumptionIncreaseExportRows, [
                      { key: "Number", label: "#" },
                      { key: "ExcelRow", label: "Row" },
                      { key: "Code", label: "Code" },
                      { key: "Product", label: "Product" },
                      { key: "UM", label: "U/M" },
                      { key: "LatestTwoVoyageDailyAverage", label: "Latest 2 Avg" },
                      { key: "PreviousFourVoyageDailyAverage", label: "Previous Avg" },
                      { key: "TwoVoyageIncreasePercent", label: "Increase %" },
                    ])
                  }
                >
                  🖨️ Print
                </button>

                <button
                  style={styles.primaryButton}
                  onClick={() =>
                    exportRowsToExcel(
                      twoVoyageConsumptionIncreaseExportRows,
                      "Increase 2 Voyages",
                      "consumption-increase-2-voyages.xlsx"
                    )
                  }
                >
                  📥 Export Excel
                </button>
              </div>
            </>
          )}


          {nextOrderView === "fml" && (
            <>
              <input
                placeholder="Search FML item, code, venue, or template location..."
                value={fmlSearch}
                onChange={(event) => setFmlSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() =>
                    printRows("FML Not Ordered / Not Used", fmlExportRows(filteredFmlNotUsedRows), [
                      { key: "Number", label: "#" },
                      { key: "FmlRow", label: "FML Row" },
                      { key: "Code", label: "Code" },
                      { key: "Product", label: "Product" },
                      { key: "Venues", label: "Venues" },
                      { key: "TemplateLocation", label: "Template" },
                      { key: "TemplateScope", label: "Scope" },
                      { key: "Reason", label: "Reason" },
                    ])
                  }
                >
                  🖨️ Print FML
                </button>

                <button
                  style={styles.primaryButton}
                  onClick={() => exportRowsToExcel(fmlExportRows(filteredFmlNotUsedRows), "FML Not Used", "fml-not-used.xlsx")}
                >
                  📥 Export FML
                </button>
              </div>
            </>
          )}

          {nextOrderView === "fmllow" && (
            <>
              <input
                placeholder="Search running-low FML item, code, venue, or template location..."
                value={fmlLowSearch}
                onChange={(event) => setFmlLowSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() =>
                    printRows("FML Running Low", fmlExportRows(filteredFmlRunningLowRows), [
                      { key: "Number", label: "#" },
                      { key: "FmlRow", label: "FML Row" },
                      { key: "Code", label: "Code" },
                      { key: "Product", label: "Product" },
                      { key: "Stock", label: "Stock" },
                      { key: "AveragePerDay", label: "Avg/Day" },
                      { key: "AvailableAtArrival", label: "At Arrival" },
                      { key: "Venues", label: "Venues" },
                      { key: "TemplateLocation", label: "Template" },
                      { key: "Reason", label: "Reason" },
                    ])
                  }
                >
                  🖨️ Print Low
                </button>

                <button
                  style={styles.primaryButton}
                  onClick={() => exportRowsToExcel(fmlExportRows(filteredFmlRunningLowRows), "FML Running Low", "fml-running-low.xlsx")}
                >
                  📥 Export Low
                </button>
              </div>
            </>
          )}

          {nextOrderView === "fmlordered" && (
            <>
              <input
                placeholder="Search ordered item missing from FML, code, row, or reason..."
                value={fmlOrderedNotFmlSearch}
                onChange={(event) => setFmlOrderedNotFmlSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.infoBox}>
                <div>📘 Report logic: <strong>Column Y ordered by ship is greater than 0</strong></div>
                <div>🧾 FML match: <strong>checks FML by code and product name</strong></div>
                <div>🔴 Shows: <strong>ordered items that are not found in FML</strong></div>
              </div>

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() =>
                    printRows("Ordered Not In FML", fmlOrderedNotFmlExportRows, [
                      { key: "Number", label: "#" },
                      { key: "ExcelRow", label: "Order Row" },
                      { key: "Code", label: "Code" },
                      { key: "Product", label: "Product" },
                      { key: "UM", label: "U/M" },
                      { key: "OrderedByShipColumnY", label: "Ordered Y" },
                      { key: "SuggestedOrder", label: "Suggested" },
                      { key: "Stock", label: "Stock" },
                      { key: "FutureOrders", label: "Future" },
                      { key: "Reason", label: "Reason" },
                    ])
                  }
                >
                  🖨️ Print
                </button>

                <button
                  style={styles.primaryButton}
                  onClick={() => exportRowsToExcel(fmlOrderedNotFmlExportRows, "Ordered Not In FML", "ordered-not-in-fml.xlsx")}
                >
                  📥 Export Excel
                </button>
              </div>
            </>
          )}
        </div>
      </section>

      {nextOrderLoading && (
        <section style={styles.card}>
          <h2 style={styles.cardTitle}>Preparing report...</h2>
          <p style={styles.emptyText}>Please wait while the workbook is processed.</p>
        </section>
      )}

      {fmlReportMessage && (nextOrderView === "fml" || nextOrderView === "fmllow") && (
        <section style={styles.card}>
          <p style={fmlReportLoading ? styles.warningText : styles.emptyText}>{fmlReportMessage}</p>
        </section>
      )}

      {nextOrderView === "order" && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>🛒 Next Order Report</h2>

          {filteredNextOrderRows.length === 0 && (
            <p style={styles.emptyText}>Upload the latest order file to generate the report.</p>
          )}

          <div style={localStyles.compactGrid}>
            {visibleNextOrderRows.map((row) => {
              const negativeArrival = Number(row.availableAtArrival || 0) < 0;
              const showPar = Number(nextOrderMeta.voyageDays || 0) === 14 && Number(row.parLevel || 0) > 0;

              return (
                <div key={row.excelRow + "-" + row.product} style={localStyles.nextOrderCard}>
                  <div style={localStyles.cardTopLine}>
                    <span>#{row.excelOrder}</span>
                    <span>Row {row.excelRow}</span>
                  </div>

                  <div style={localStyles.productName}>{row.product}</div>
                  <div style={styles.recipeMeta}>Code: {row.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>U/M: {row.unit || "N/A"}</div>

                  <div style={localStyles.calcStrip}>
                    <div>Past: <strong>{formatQty(row.pastConsumption)}</strong> · Avg/day: <strong>{formatQty(row.averagePerDay)}</strong></div>
                    <div>Days to arrival: <strong>{formatQty(nextOrderMeta.daysUntilArrival)}</strong> · Use until arrival: <strong>{formatQty(row.usageUntilArrival)}</strong></div>
                    <div>Voyage need: <strong>{formatQty(row.projectedVoyageNeed)}</strong> · Raw order: <strong>{formatQty(row.rawSuggested)}</strong></div>
                    {row.parCapApplied && <div style={localStyles.parCapNote}>Par cap applied: max {formatQty(row.parCapLimit)}</div>}
                  </div>

                  <div style={localStyles.metricGrid}>
                    <div style={localStyles.metricBox}>
                      <span>Stock</span>
                      <strong>{formatQty(row.stock)}</strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Future</span>
                      <strong>{formatQty(row.futureOrders)}</strong>
                    </div>

                    <div style={negativeArrival ? localStyles.metricBoxBad : localStyles.metricBox}>
                      <span>{negativeArrival ? "⚠️ At arrival" : "At arrival"}</span>
                      <strong>{formatQty(row.availableAtArrival)}</strong>
                    </div>

                    {showPar && (
                      <div style={row.parCapApplied ? localStyles.metricBoxWarning : localStyles.metricBox}>
                        <span>Par Q</span>
                        <strong>{formatQty(row.parLevel)}</strong>
                      </div>
                    )}
                  </div>

                  <div style={row.suggestedOrder > 0 ? localStyles.orderBadge : styles.statusGood}>
                    Order: {formatQty(row.suggestedOrder)}
                  </div>

                  <div
                    style={
                      row.alertType === "needs-order"
                        ? localStyles.blueBadge
                        : row.alertType === "stock-no-consumption"
                        ? styles.statusBad
                        : row.alertType === "no-stock-no-consumption"
                        ? localStyles.reviewBadge
                        : styles.statusNeutral
                    }
                  >
                    {row.alertLabel}
                  </div>
                </div>
              );
            })}
          </div>

          {hasMoreNextOrderRows && (
            <button
              style={styles.backButton}
              onClick={() => setReportDisplayLimit((value) => value + REPORT_RENDER_BATCH)}
            >
              Show more ({visibleNextOrderRows.length} / {filteredNextOrderRows.length})
            </button>
          )}
        </section>
      )}

      {nextOrderView === "orderedVsSuggested" && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📊 Ordered vs Suggested Report</h2>

          {orderedVsSuggestedRows.length === 0 && (
            <p style={styles.emptyText}>
              Upload the latest order file to compare ordered quantity from column Y against suggested order.
            </p>
          )}

          <div style={localStyles.compactGrid}>
            {visibleOrderedVsSuggestedRows.map((row, index) => {
              const cardStyle = {
                ...localStyles.orderedVsSuggestedCard,
                ...(row.orderComparisonStatus === "red"
                  ? localStyles.orderedVsSuggestedRed
                  : row.orderComparisonStatus === "blue"
                    ? localStyles.orderedVsSuggestedBlue
                    : localStyles.orderedVsSuggestedGreen),
              };

              const badgeStyle =
                row.orderComparisonStatus === "red"
                  ? localStyles.comparisonBadgeRed
                  : row.orderComparisonStatus === "blue"
                    ? localStyles.comparisonBadgeBlue
                    : localStyles.comparisonBadgeGreen;

              return (
                <div key={row.excelRow + "-" + row.product + "-ordered"} style={cardStyle}>
                  <div style={localStyles.cardTopLine}>
                    <span>#{index + 1}</span>
                    <span>Row {row.excelRow}</span>
                  </div>

                  <div style={localStyles.productName}>{row.product}</div>
                  <div style={styles.recipeMeta}>Code: {row.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>U/M: {row.unit || "N/A"}</div>

                  <div style={localStyles.metricGrid}>
                    <div style={localStyles.metricBox}>
                      <span>On Hand</span>
                      <strong>{formatQty(row.stock)}</strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Upcoming</span>
                      <strong>{formatQty(row.futureOrders)}</strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Daily Use</span>
                      <strong>{formatQty(row.averagePerDay)}</strong>
                    </div>
                  </div>

                  <div style={localStyles.metricGrid}>
                    <div style={localStyles.metricBox}>
                      <span>Ordered Y</span>
                      <strong>{formatQty(row.orderedQty)}</strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Suggested</span>
                      <strong>{formatQty(row.suggestedQty)}</strong>
                    </div>

                    <div style={localStyles.metricBox}>
                      <span>Diff</span>
                      <strong>{formatQty(row.orderDifference)}</strong>
                    </div>
                  </div>

                  <div style={localStyles.calcStrip}>
                    <div>
                      Difference %: <strong>{formatQty(row.orderDifferencePercent)}%</strong>
                    </div>
                  </div>

                  <div style={badgeStyle}>
                    {row.orderComparisonLabel}
                  </div>
                </div>
              );
            })}
          </div>

          {hasMoreOrderedVsSuggestedRows && (
            <button
              style={styles.backButton}
              onClick={() => setReportDisplayLimit((value) => value + REPORT_RENDER_BATCH)}
            >
              Show more ({visibleOrderedVsSuggestedRows.length} / {orderedVsSuggestedRows.length})
            </button>
          )}
        </section>
      )}


      {nextOrderView === "increase1" && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📈 Consumption Increase - Last 1 Voyage</h2>

          {oneVoyageConsumptionIncreaseRows.length === 0 && (
            <p style={styles.emptyText}>
              No items found with a 25% or higher increase in the latest voyage.
            </p>
          )}

          <div style={localStyles.compactGrid}>
            {visibleOneVoyageConsumptionIncreaseRows.map((row, index) => (
              <div
                key={row.excelRow + "-" + row.product + "-increase1"}
                style={{ ...localStyles.orderedVsSuggestedCard, ...localStyles.orderedVsSuggestedRed }}
              >
                <div style={localStyles.cardTopLine}>
                  <span>#{index + 1}</span>
                  <span>Row {row.excelRow}</span>
                </div>

                <div style={localStyles.productName}>{row.product}</div>
                <div style={styles.recipeMeta}>Code: {row.code || "N/A"}</div>
                <div style={styles.recipeMeta}>U/M: {row.unit || "N/A"}</div>

                <div style={localStyles.metricGrid}>
                  <div style={localStyles.metricBox}>
                    <span>Latest {row.latestOneVoyageColumn}</span>
                    <strong>{formatQty(row.latestOneVoyageDaily)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Prev Avg</span>
                    <strong>{formatQty(row.previousFiveVoyageDailyAverage)}</strong>
                  </div>

                  <div style={localStyles.metricBoxBad}>
                    <span>Increase</span>
                    <strong>{formatQty(row.oneVoyageIncreasePercent)}%</strong>
                  </div>
                </div>

                <div style={localStyles.comparisonBadgeRed}>
                  Latest voyage consumption increased by {formatQty(row.oneVoyageIncreasePercent)}%
                </div>
              </div>
            ))}
          </div>

          {hasMoreOneVoyageConsumptionIncreaseRows && (
            <button
              style={styles.backButton}
              onClick={() => setReportDisplayLimit((value) => value + REPORT_RENDER_BATCH)}
            >
              Show more ({visibleOneVoyageConsumptionIncreaseRows.length} / {oneVoyageConsumptionIncreaseRows.length})
            </button>
          )}
        </section>
      )}

      {nextOrderView === "increase2" && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📈 Consumption Increase - Last 2 Voyages</h2>

          {twoVoyageConsumptionIncreaseRows.length === 0 && (
            <p style={styles.emptyText}>
              No items found with a 25% or higher increase across the latest two voyages.
            </p>
          )}

          <div style={localStyles.compactGrid}>
            {visibleTwoVoyageConsumptionIncreaseRows.map((row, index) => (
              <div
                key={row.excelRow + "-" + row.product + "-increase2"}
                style={{ ...localStyles.orderedVsSuggestedCard, ...localStyles.orderedVsSuggestedRed }}
              >
                <div style={localStyles.cardTopLine}>
                  <span>#{index + 1}</span>
                  <span>Row {row.excelRow}</span>
                </div>

                <div style={localStyles.productName}>{row.product}</div>
                <div style={styles.recipeMeta}>Code: {row.code || "N/A"}</div>
                <div style={styles.recipeMeta}>U/M: {row.unit || "N/A"}</div>

                <div style={localStyles.metricGrid}>
                  <div style={localStyles.metricBox}>
                    <span>Latest {row.latestTwoVoyageColumns}</span>
                    <strong>{formatQty(row.latestTwoVoyageDailyAverage)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Prev Avg</span>
                    <strong>{formatQty(row.previousFourVoyageDailyAverage)}</strong>
                  </div>

                  <div style={localStyles.metricBoxBad}>
                    <span>Increase</span>
                    <strong>{formatQty(row.twoVoyageIncreasePercent)}%</strong>
                  </div>
                </div>

                <div style={localStyles.comparisonBadgeRed}>
                  Latest 2 voyages consumption increased by {formatQty(row.twoVoyageIncreasePercent)}%
                </div>
              </div>
            ))}
          </div>

          {hasMoreTwoVoyageConsumptionIncreaseRows && (
            <button
              style={styles.backButton}
              onClick={() => setReportDisplayLimit((value) => value + REPORT_RENDER_BATCH)}
            >
              Show more ({visibleTwoVoyageConsumptionIncreaseRows.length} / {twoVoyageConsumptionIncreaseRows.length})
            </button>
          )}
        </section>
      )}


      {nextOrderView === "fml" && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📘 FML Not Ordered / Not Used</h2>

          {filteredFmlNotUsedRows.length === 0 && (
            <p style={styles.emptyText}>
              {fmlReportLoading
                ? "Preparing FML report..."
                : fmlNotUsedPrepared
                ? "No FML not-ordered / not-used records found for this ship/template."
                : "Open this report to prepare FML not-ordered / not-used records."}
            </p>
          )}

          <div style={localStyles.compactGrid}>
            {visibleFmlNotUsedRows.map((row, index) => (
              <div key={row.id} style={localStyles.fmlCard}>
                <div style={localStyles.cardTopLine}>
                  <span>#{index + 1}</span>
                  <span>FML Row {row.fmlRow}</span>
                </div>

                <div style={localStyles.productName}>{row.product}</div>
                <div style={localStyles.miniMeta}>Code: {row.code || "N/A"} · U/M: {row.unit || "N/A"}</div>

                <div style={localStyles.metricGrid}>
                  <div style={localStyles.metricBox}>
                    <span>Stock</span>
                    <strong>{formatQty(row.stock)}</strong>
                  </div>
                  <div style={localStyles.metricBox}>
                    <span>Future</span>
                    <strong>{formatQty(row.futureOrders)}</strong>
                  </div>
                  <div style={localStyles.metricBox}>
                    <span>Past</span>
                    <strong>{formatQty(row.pastConsumption)}</strong>
                  </div>
                </div>

                <div style={localStyles.blueBadge}>FML venues</div>
                <div style={localStyles.compactTextBlock}>{row.venuesText}</div>
                <div style={localStyles.templateNote}>{row.scopeText}</div>
                <div style={localStyles.compactTextBlock}>Template: {row.templateLocation}</div>
                <div style={localStyles.reviewBadge}>{row.reason}</div>
              </div>
            ))}
          </div>

          {hasMoreFmlNotUsedRows && (
            <button
              style={styles.backButton}
              onClick={() => setReportDisplayLimit((value) => value + REPORT_RENDER_BATCH)}
            >
              Show more ({visibleFmlNotUsedRows.length} / {filteredFmlNotUsedRows.length})
            </button>
          )}
        </section>
      )}

      {nextOrderView === "fmllow" && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>⚠️ FML Running Low by Arrival</h2>

          {filteredFmlRunningLowRows.length === 0 && (
            <p style={styles.emptyText}>
              {fmlReportLoading
                ? "Preparing FML running-low report..."
                : fmlRunningLowPrepared
                ? "No FML running-low records found for this ship/template."
                : "Open this report to prepare FML running-low records."}
            </p>
          )}

          <div style={localStyles.compactGrid}>
            {visibleFmlRunningLowRows.map((row, index) => (
              <div key={row.id} style={localStyles.fmlCard}>
                <div style={localStyles.cardTopLine}>
                  <span>#{index + 1}</span>
                  <span>FML Row {row.fmlRow}</span>
                </div>

                <div style={localStyles.productName}>{row.product}</div>
                <div style={localStyles.miniMeta}>Code: {row.code || "N/A"} · U/M: {row.unit || "N/A"}</div>

                <div style={localStyles.metricGrid}>
                  <div style={localStyles.metricBox}>
                    <span>Stock</span>
                    <strong>{formatQty(row.stock)}</strong>
                  </div>
                  <div style={localStyles.metricBox}>
                    <span>Avg/day</span>
                    <strong>{formatQty(row.averagePerDay)}</strong>
                  </div>
                  <div style={Number(row.availableAtArrival || 0) < 0 ? localStyles.metricBoxBad : localStyles.metricBoxWarning}>
                    <span>At arrival</span>
                    <strong>{formatQty(row.availableAtArrival)}</strong>
                  </div>
                </div>

                <div style={localStyles.blueBadge}>FML venues</div>
                <div style={localStyles.compactTextBlock}>{row.venuesText}</div>
                <div style={localStyles.templateNote}>{row.scopeText}</div>
                <div style={localStyles.compactTextBlock}>Template: {row.templateLocation}</div>
                <div style={localStyles.reviewBadge}>{row.reason}</div>
              </div>
            ))}
          </div>

          {hasMoreFmlRunningLowRows && (
            <button
              style={styles.backButton}
              onClick={() => setReportDisplayLimit((value) => value + REPORT_RENDER_BATCH)}
            >
              Show more ({visibleFmlRunningLowRows.length} / {filteredFmlRunningLowRows.length})
            </button>
          )}
        </section>
      )}

      {nextOrderView === "fmlordered" && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>🧾 Ordered Items Not In FML</h2>

          {filteredFmlOrderedNotFmlRows.length === 0 && (
            <p style={styles.emptyText}>
              No ordered items missing from FML found. This report checks Column Y ordered quantities against the FML sheet.
            </p>
          )}

          <div style={localStyles.compactGrid}>
            {visibleFmlOrderedNotFmlRows.map((row, index) => (
              <div key={row.id || row.excelRow + "-" + row.product} style={localStyles.fmlOrderedCard}>
                <div style={localStyles.cardTopLine}>
                  <span>#{index + 1}</span>
                  <span>Order Row {row.excelRow}</span>
                </div>

                <div style={localStyles.productName}>{row.product}</div>
                <div style={localStyles.miniMeta}>Code: {row.code || "N/A"} · U/M: {row.unit || "N/A"}</div>

                <div style={localStyles.metricGrid}>
                  <div style={localStyles.metricBox}>
                    <span>Ordered Y</span>
                    <strong>{formatQty(row.orderedByShip)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Suggested</span>
                    <strong>{formatQty(row.suggestedOrder)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Stock</span>
                    <strong>{formatQty(row.stock)}</strong>
                  </div>
                </div>

                <div style={localStyles.metricGrid}>
                  <div style={localStyles.metricBox}>
                    <span>Future</span>
                    <strong>{formatQty(row.futureOrders)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Past</span>
                    <strong>{formatQty(row.pastConsumption)}</strong>
                  </div>

                  <div style={localStyles.metricBox}>
                    <span>Avg/day</span>
                    <strong>{formatQty(row.averagePerDay)}</strong>
                  </div>
                </div>

                <div style={localStyles.dangerBadge}>Not found in FML</div>
                <div style={localStyles.reviewBadge}>{row.reason}</div>
              </div>
            ))}
          </div>

          {hasMoreFmlOrderedNotFmlRows && (
            <button
              style={styles.backButton}
              onClick={() => setReportDisplayLimit((value) => value + REPORT_RENDER_BATCH)}
            >
              Show more ({visibleFmlOrderedNotFmlRows.length} / {filteredFmlOrderedNotFmlRows.length})
            </button>
          )}
        </section>
      )}
    </main>
  );
}

const localStyles = {
  compactGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(155px, 1fr))",
    gap: 7,
  },
  nextOrderCard: {
    border: "2px solid #0057b8",
    borderRadius: 14,
    padding: 8,
    background: "#eef5ff",
    display: "grid",
    gap: 6,
    fontSize: 11,
  },
  fmlCard: {
    border: "1.5px solid #0057b8",
    borderRadius: 10,
    padding: 6,
    background: "#eef5ff",
    display: "grid",
    gap: 4,
    fontSize: 10,
    minWidth: 0,
    alignContent: "start",
    overflow: "hidden",
  },
  fmlOrderedCard: {
    border: "1.5px solid #b00020",
    borderRadius: 10,
    padding: 6,
    background: "#fff0f0",
    display: "grid",
    gap: 4,
    fontSize: 10,
    minWidth: 0,
    alignContent: "start",
    overflow: "hidden",
  },
  cardTopLine: {
    display: "flex",
    justifyContent: "space-between",
    color: "#555",
    fontSize: 9,
    fontWeight: "bold",
    gap: 4,
  },
  productName: {
    fontWeight: "bold",
    fontSize: 11.5,
    lineHeight: 1.08,
    wordBreak: "break-word",
  },
  miniMeta: {
    color: "#555",
    fontSize: 9.5,
    lineHeight: 1.15,
  },
  compactTextBlock: {
    color: "#444",
    fontSize: 9.5,
    lineHeight: 1.12,
    wordBreak: "break-word",
    maxHeight: 46,
    overflowY: "auto",
  },
  calcStrip: {
    padding: "5px 6px",
    borderRadius: 8,
    background: "#fff",
    border: "1px solid #d7e7ff",
    color: "#333",
    fontSize: 9.5,
    lineHeight: 1.25,
    display: "grid",
    gap: 2,
  },
  parCapNote: {
    color: "#8a5a00",
    fontWeight: "bold",
  },
  metricGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
    gap: 4,
  },
  metricBox: {
    border: "1px solid #ddd",
    borderRadius: 7,
    padding: "4px 2px",
    background: "#fff",
    display: "grid",
    gap: 1,
    textAlign: "center",
    fontSize: 9.5,
    minHeight: 38,
  },
  metricBoxBad: {
    border: "1px solid #b00020",
    borderRadius: 7,
    padding: "4px 2px",
    background: "#fff0f0",
    color: "#b00020",
    display: "grid",
    gap: 1,
    textAlign: "center",
    fontWeight: "bold",
    fontSize: 9.5,
    minHeight: 38,
  },
  metricBoxWarning: {
    border: "1px solid #8a5a00",
    borderRadius: 7,
    padding: "4px 2px",
    background: "#fff4d6",
    color: "#8a5a00",
    display: "grid",
    gap: 1,
    textAlign: "center",
    fontWeight: "bold",
    fontSize: 9.5,
    minHeight: 38,
  },
  orderBadge: {
    padding: 6,
    borderRadius: 8,
    background: "#0057b8",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 11.5,
  },
  blueBadge: {
    padding: "4px 5px",
    borderRadius: 7,
    background: "#0057b8",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 9.5,
    lineHeight: 1.1,
  },
  dangerBadge: {
    padding: "4px 5px",
    borderRadius: 7,
    background: "#b00020",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 9.5,
    lineHeight: 1.1,
  },
  reviewBadge: {
    padding: "5px 5px",
    borderRadius: 7,
    background: "#fff4d6",
    color: "#8a5a00",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 9.5,
    lineHeight: 1.1,
  },
  templateNote: {
    color: "#0057b8",
    fontWeight: "bold",
    fontSize: 9.5,
    lineHeight: 1.1,
  },
  orderedVsSuggestedCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 8,
    background: "#fff",
    display: "grid",
    gap: 6,
    fontSize: 11,
  },
  orderedVsSuggestedRed: {
    border: "2px solid #b00020",
    background: "#fff0f0",
  },
  orderedVsSuggestedBlue: {
    border: "2px solid #0057b8",
    background: "#eef5ff",
  },
  orderedVsSuggestedGreen: {
    border: "2px solid #2e7d32",
    background: "#e8f5e9",
  },
  comparisonBadgeRed: {
    padding: 6,
    borderRadius: 8,
    background: "#b00020",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 10,
    lineHeight: 1.1,
  },
  comparisonBadgeBlue: {
    padding: 6,
    borderRadius: 8,
    background: "#0057b8",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 10,
    lineHeight: 1.1,
  },
  comparisonBadgeGreen: {
    padding: 6,
    borderRadius: 8,
    background: "#2e7d32",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
    fontSize: 10,
    lineHeight: 1.1,
  },
};

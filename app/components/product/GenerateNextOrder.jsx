"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

import {
  YEARLY_REGION_ALL,
  parseYearlyRegionalConsumptionWorkbook,
} from "../../lib/yearlyRegionalConsumption";

const ORDER_BUFFER_PERCENT = 25;

const SHIPS = ["SC", "VL", "BRL", "RL"];

const SHIP_DISPLAY_NAMES = {
  SC: "Scarlet",
  VL: "Valiant",
  BRL: "Brilliant",
  RL: "Resilient",
};

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const safeText = (value) =>
  String(value || "")
    .replace(/\s+/g, " ")
    .trim();

const formatQty = (value) => Number(value || 0).toFixed(2);

const formatMoney = (value) => "$" + Number(value || 0).toFixed(2);

const waitForBrowser = (ms = 0) =>
  new Promise((resolve) => window.setTimeout(resolve, ms));

const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

const toNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);

  return Number.isFinite(number) ? number : 0;
};

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

  if (!date) {
    return String(value || "").trim();
  }

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

const normalizeOrderCode = (value) => {
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

const normalizeShipCode = (value) => {
  const text = cleanText(value).replace(/RESILIANT/g, "RESILIENT");
  const compact = text.replace(/[^A-Z0-9]/g, "");

  if (!text) return "";

  // Important:
  // Some old files show Scarlet as V1 / V 1 / SCL.
  if (
    compact === "V1" ||
    compact === "SCL" ||
    text === "SC" ||
    text.includes("SCARLET")
  ) {
    return "SC";
  }

  if (compact === "VL" || text.includes("VALIANT")) return "VL";
  if (compact === "BRL" || text.includes("BRILLIANT")) return "BRL";
  if (compact === "RL" || text.includes("RESILIENT")) return "RL";

  return "";
};

const getShipDisplayName = (shipCode) =>
  SHIP_DISPLAY_NAMES[shipCode] || shipCode || "";

const getHistoricalSailorDays = (cellA, cellB) => {
  const a = toNumber(cellA);
  const b = toNumber(cellB);

  if (!a && !b) return 0;
  if (a && !b) return a;
  if (!a && b) return b;

  const low = Math.min(Math.abs(a), Math.abs(b));
  const high = Math.max(Math.abs(a), Math.abs(b));

  // Some files store total sailor-days in one row and days in the other.
  if (low > 0 && high > low * 1000) return high;

  // Otherwise treat the two cells as sailors x days.
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

const getProductNameMatchKey = (value) => {
  const tokens = [...new Set(getProductMatchTokens(value))].sort();

  return tokens.length ? tokens.join("|") : cleanText(value);
};

const sumColumns = (row, columnIndexes) =>
  columnIndexes.reduce((sum, columnIndex) => sum + toNumber(row[columnIndex]), 0);

const getOrderSheetName = (workbook) => {
  if (workbook.SheetNames.includes("Standard Order Template")) {
    return "Standard Order Template";
  }

  if (workbook.SheetNames.includes("Order Sheet")) {
    return "Order Sheet";
  }

  return workbook.SheetNames[0];
};

const buildRegionalIndex = (yearlyRegionalConsumption) => {
  const rows = Array.isArray(yearlyRegionalConsumption?.rows)
    ? yearlyRegionalConsumption.rows
    : [];

  const byCode = new Map();
  const byName = new Map();

  const addToMap = (map, key, row) => {
    if (!key) return;

    if (!map.has(key)) {
      map.set(key, []);
    }

    map.get(key).push(row);
  };

  rows.forEach((row) => {
    const codeKey = normalizeOrderCode(row.productCode);
    const nameKey = getProductNameMatchKey(row.productName);

    addToMap(byCode, codeKey, row);
    addToMap(byName, nameKey, row);
  });

  return {
    rows,
    byCode,
    byName,
  };
};

const getRegionalRowsForOrderItem = (item, regionalIndex) => {
  if (!regionalIndex) return [];

  const codeKey = normalizeOrderCode(item.code);

  if (codeKey && regionalIndex.byCode.has(codeKey)) {
    return regionalIndex.byCode.get(codeKey) || [];
  }

  const nameKey = getProductNameMatchKey(item.product);

  if (nameKey && regionalIndex.byName.has(nameKey)) {
    return regionalIndex.byName.get(nameKey) || [];
  }

  return [];
};

const summarizeRegionalRows = (rows) => {
  const summary = {
    totalQty: 0,
    totalValue: 0,
    totalDays: 0,
    blockCount: 0,
    avgDailyQty: 0,
    avgPrice: 0,
    regions: new Set(),
    ships: new Set(),
    months: new Set(),
  };

  rows.forEach((row) => {
    const qty = Number(row.qty || 0);
    const value = Number(row.value || 0);
    const days = Number(row.days || 0);

    summary.totalQty += Number.isFinite(qty) ? qty : 0;
    summary.totalValue += Number.isFinite(value) ? value : 0;
    summary.totalDays += Number.isFinite(days) ? days : 0;
    summary.blockCount += 1;

    if (row.region) summary.regions.add(row.region);
    if (row.ship) summary.ships.add(row.ship);
    if (row.monthName || row.monthKey) {
      summary.months.add(row.monthName || row.monthKey);
    }
  });

  summary.avgDailyQty =
    summary.totalDays > 0 ? summary.totalQty / summary.totalDays : 0;

  summary.avgPrice =
    summary.totalQty > 0 ? summary.totalValue / summary.totalQty : 0;

  return {
    ...summary,
    regions: [...summary.regions].sort(),
    ships: [...summary.ships].sort(),
    months: [...summary.months].sort(),
  };
};

const enrichOrderRowWithRegionalComparison = ({
  item,
  regionalIndex,
  selectedRegion,
}) => {
  const matchedRows = getRegionalRowsForOrderItem(item, regionalIndex);

  const orderShipRows = matchedRows.filter(
    (row) => normalizeShipCode(row.ship) === item.orderShipCode
  );

  const marketRows =
    selectedRegion && selectedRegion !== YEARLY_REGION_ALL
      ? matchedRows.filter((row) => String(row.region || "") === selectedRegion)
      : selectedRegion === YEARLY_REGION_ALL
      ? matchedRows
      : [];

  const yearSummary = summarizeRegionalRows(matchedRows);
  const marketSummary = summarizeRegionalRows(marketRows);
  const shipSummary = summarizeRegionalRows(orderShipRows);

  const suggestedParFromYearTotal =
    yearSummary.avgDailyQty > 0
      ? yearSummary.avgDailyQty * item.voyageDays * 1.25
      : 0;

  const suggestedParFromMarket =
    marketSummary.avgDailyQty > 0
      ? marketSummary.avgDailyQty * item.voyageDays * 1.25
      : 0;

  const suggestedParFromShip =
    shipSummary.avgDailyQty > 0
      ? shipSummary.avgDailyQty * item.voyageDays * 1.25
      : 0;

  const firstMatchedRow = matchedRows[0] || {};

  return {
    ...item,

    regionalMatchedProductCode: firstMatchedRow.productCode || "",
    regionalMatchedProductName: firstMatchedRow.productName || "",
    regionalMatchCount: matchedRows.length,

    yearAvgDailyQty: yearSummary.avgDailyQty,
    yearTotalQty: yearSummary.totalQty,
    yearTotalDays: yearSummary.totalDays,
    yearTotalValue: yearSummary.totalValue,
    yearBlockCount: yearSummary.blockCount,
    yearRegions: yearSummary.regions,
    yearShips: yearSummary.ships,

    marketAvgDailyQty: marketSummary.avgDailyQty,
    marketTotalQty: marketSummary.totalQty,
    marketTotalDays: marketSummary.totalDays,
    marketTotalValue: marketSummary.totalValue,
    marketBlockCount: marketSummary.blockCount,
    marketRegions: marketSummary.regions,
    marketShips: marketSummary.ships,

    shipYearAvgDailyQty: shipSummary.avgDailyQty,
    shipYearTotalQty: shipSummary.totalQty,
    shipYearTotalDays: shipSummary.totalDays,
    shipYearBlockCount: shipSummary.blockCount,

    suggestedParFromYearTotal,
    suggestedParFromMarket,
    suggestedParFromShip,
  };
};

const parseNextOrderWorkbook = async ({
  workbook,
  fileName,
  fallbackShip,
  onProgress,
}) => {
  const sheetName = getOrderSheetName(workbook);
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    throw new Error("Could not find an order worksheet.");
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  const orderShipName = safeText(rows[0]?.[1]);
  const orderShipCode =
    normalizeShipCode(orderShipName) || normalizeShipCode(fallbackShip);

  const rawOrderDate = rows[1]?.[1]; // B2
  const rawArrivalDate = rows[2]?.[1]; // B3
  const targetSailors = toNumber(rows[4]?.[1]); // B5
  const voyageDays = toNumber(rows[5]?.[1]); // B6

  const daysUntilArrival = getDaysBetweenCells(rawOrderDate, rawArrivalDate);
  const currentPeriodSailorDays = targetSailors * voyageDays;

  const futureOrderColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13]; // F:N
  const pastConsumptionColumns = [34, 35, 36, 37, 38, 39]; // AI:AN

  const historicalSailorDays = pastConsumptionColumns.reduce(
    (sum, colIndex) =>
      sum + getHistoricalSailorDays(rows[4]?.[colIndex], rows[5]?.[colIndex]),
    0
  );

  const parsedRows = [];

  for (let rowIndex = 9; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const excelRow = rowIndex + 1;

    const code = safeText(row[0]);
    const product = safeText(row[1]);
    const uom = safeText(row[2]);

    if (!product || !uom) continue;
    if (cleanText(product) === "PRODUCT") continue;
    if (cleanText(product).includes("PRODUCT NAME")) continue;

    const stockOnHand = toNumber(row[3]); // D
    const futureOrders = sumColumns(row, futureOrderColumns); // F:N
    const currentFileParLevel = toNumber(row[16]); // Q
    const pastConsumption = sumColumns(row, pastConsumptionColumns); // AI:AN

    const averageConsumptionPerSailorDay =
      historicalSailorDays > 0 ? pastConsumption / historicalSailorDays : 0;

    const averageConsumptionPerDay =
      averageConsumptionPerSailorDay > 0 && targetSailors > 0
        ? averageConsumptionPerSailorDay * targetSailors
        : voyageDays > 0
        ? pastConsumption / voyageDays
        : 0;

    const arrivalLeadDays = daysUntilArrival;
    const totalProjectionDays = arrivalLeadDays + voyageDays;

    // Correct logic:
    // Arrival days are only for estimating stock on arrival.
    const consumptionUntilArrival =
      averageConsumptionPerDay * arrivalLeadDays;

    const estimatedQtyAtArrival =
      stockOnHand + futureOrders - consumptionUntilArrival;

    // Order cannot fix stock that runs out before arrival.
    // For the new order calculation, negative arrival stock is treated as 0.
    const usableQtyAtArrival = Math.max(estimatedQtyAtArrival, 0);

    // New order is based only on voyage days + 25% buffer.
    const voyageNeedBeforeBuffer = averageConsumptionPerDay * voyageDays;

    const orderBufferQty =
      voyageNeedBeforeBuffer * (ORDER_BUFFER_PERCENT / 100);

    const targetQtyForVoyage = voyageNeedBeforeBuffer + orderBufferQty;

    const suggestedParFromCurrentUsage = targetQtyForVoyage;

    const suggestedOrder = Math.max(
      targetQtyForVoyage - usableQtyAtArrival,
      0
    );

    let alertType = suggestedOrder > 0 ? "order" : "normal";
    let alertLabel = suggestedOrder > 0 ? "Needs order" : "No order suggested";

    let alertDescription =
      "Stock on arrival is estimated from stock on hand + future orders - average daily consumption until arrival. Suggested order covers voyage days plus 25% buffer only.";

    if (estimatedQtyAtArrival < 0) {
      alertType = "red";
      alertLabel = "May run out before arrival";
      alertDescription =
        "Projected stock on arrival is below zero. The suggested order still covers only the voyage need plus 25% buffer.";
    }

    if (averageConsumptionPerDay <= 0 && stockOnHand <= 0 && futureOrders <= 0) {
      alertType = "blue";
      alertLabel = "No usage and no stock";
      alertDescription =
        "No average daily usage could be calculated and no stock/future order was found. Review manually.";
    }

    parsedRows.push({
      excelRow,
      orderShipCode,
      orderShipName,

      code,
      product,
      uom,

      stockOnHand,
      futureOrders,
      pastConsumption,

      currentFileParLevel,
      parLevel: currentFileParLevel,

      targetSailors,
      voyageDays,
      arrivalLeadDays,
      daysUntilArrival,
      totalProjectionDays,

      historicalSailorDays,
      currentPeriodSailorDays,
      averageConsumptionPerSailorDay,
      averageConsumptionPerDay,

      consumptionUntilArrival,
      estimatedQtyAtArrival,
      availableAtArrival: estimatedQtyAtArrival,
      usableQtyAtArrival,

      voyageNeedBeforeBuffer,
      orderBufferPercent: ORDER_BUFFER_PERCENT,
      orderBufferQty,
      targetQtyForVoyage,

      suggestedParFromCurrentUsage,
      parDifferenceVsCurrent:
        suggestedParFromCurrentUsage - currentFileParLevel,

      suggestedOrder,

      alertType,
      alertLabel,
      alertDescription,
      orderReason: alertDescription,
    });

    if (rowIndex % 250 === 0) {
      onProgress?.(
        `Reading order rows... ${rowIndex + 1} of ${rows.length}`
      );
      await waitForBrowser(0);
    }
  }

  return {
    rows: parsedRows,
    meta: {
      fileName,
      sheetName,
      orderShipName,
      orderShipCode,
      orderDate: formatDateCell(rawOrderDate),
      arrivalDate: formatDateCell(rawArrivalDate),
      targetSailors,
      voyageDays,
      daysUntilArrival,
      totalProjectionDays: daysUntilArrival + voyageDays,
      historicalSailorDays,
      currentPeriodSailorDays,
      totalItems: parsedRows.length,
      itemsNeedingOrder: parsedRows.filter(
        (item) => Number(item.suggestedOrder || 0) > 0
      ).length,
      runOutBeforeArrival: parsedRows.filter(
        (item) => Number(item.estimatedQtyAtArrival || 0) < 0
      ).length,
    },
  };
};

export default function GenerateNextOrder({
  styles,
  userShip,
  onBack,
  logUsageEvent = () => {},
  yearlyRegionalConsumption,
  setYearlyRegionalConsumption = () => {},
  yearlyRegionalFileName = "",
  setYearlyRegionalFileName = () => {},
  selectedRegionalConsumptionRegion = "",
  setSelectedRegionalConsumptionRegion = () => {},
}) {
  const [orderFileName, setOrderFileName] = useState("");
  const [orderMeta, setOrderMeta] = useState(null);
  const [baseRows, setBaseRows] = useState([]);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all");
  const [sortMode, setSortMode] = useState("excel");
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [selectedItem, setSelectedItem] = useState(null);

  const regionalIndex = useMemo(
    () => buildRegionalIndex(yearlyRegionalConsumption),
    [yearlyRegionalConsumption]
  );

  const rowsWithRegional = useMemo(
    () =>
      baseRows.map((item) =>
        enrichOrderRowWithRegionalComparison({
          item,
          regionalIndex,
          selectedRegion: selectedRegionalConsumptionRegion,
        })
      ),
    [baseRows, regionalIndex, selectedRegionalConsumptionRegion]
  );

  const filterCounts = useMemo(
    () => ({
      all: rowsWithRegional.length,
      needsOrder: rowsWithRegional.filter(
        (item) => Number(item.suggestedOrder || 0) > 0
      ).length,
      noOrder: rowsWithRegional.filter(
        (item) => Number(item.suggestedOrder || 0) <= 0
      ).length,
      runOut: rowsWithRegional.filter(
        (item) => Number(item.estimatedQtyAtArrival || 0) < 0
      ).length,
      regionalMatched: rowsWithRegional.filter(
        (item) => Number(item.regionalMatchCount || 0) > 0
      ).length,
    }),
    [rowsWithRegional]
  );

  const visibleRows = useMemo(() => {
    const query = search.toLowerCase().trim();

    let rows = rowsWithRegional.filter((item) => {
      const matchesSearch =
        !query ||
        [
          item.code,
          item.product,
          item.uom,
          item.alertLabel,
          item.excelRow,
          item.regionalMatchedProductName,
          item.regionalMatchedProductCode,
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);

      if (!matchesSearch) return false;

      if (filter === "needsOrder") {
        return Number(item.suggestedOrder || 0) > 0;
      }

      if (filter === "noOrder") {
        return Number(item.suggestedOrder || 0) <= 0;
      }

      if (filter === "runOut") {
        return Number(item.estimatedQtyAtArrival || 0) < 0;
      }

      if (filter === "regionalMatched") {
        return Number(item.regionalMatchCount || 0) > 0;
      }

      return true;
    });

    rows = [...rows];

    if (sortMode === "suggested") {
      rows.sort(
        (a, b) =>
          Number(b.suggestedOrder || 0) - Number(a.suggestedOrder || 0) ||
          String(a.product || "").localeCompare(String(b.product || ""))
      );
    } else if (sortMode === "arrivalStock") {
      rows.sort(
        (a, b) =>
          Number(a.estimatedQtyAtArrival || 0) -
            Number(b.estimatedQtyAtArrival || 0) ||
          String(a.product || "").localeCompare(String(b.product || ""))
      );
    } else if (sortMode === "product") {
      rows.sort((a, b) =>
        String(a.product || "").localeCompare(String(b.product || ""))
      );
    } else {
      rows.sort(
        (a, b) =>
          Number(a.excelRow || 0) - Number(b.excelRow || 0) ||
          String(a.product || "").localeCompare(String(b.product || ""))
      );
    }

    return rows;
  }, [rowsWithRegional, search, filter, sortMode]);

  const uploadOrderFile = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    setLoading(true);
    setMessage("Opening order workbook...");
    setSelectedItem(null);

    try {
      await waitForBrowser(30);

      const arrayBuffer = await file.arrayBuffer();

      setMessage("Reading Excel workbook...");
      await waitForBrowser(30);

      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsed = await parseNextOrderWorkbook({
        workbook,
        fileName: file.name,
        fallbackShip: userShip,
        onProgress: setMessage,
      });

      setOrderFileName(file.name);
      setOrderMeta(parsed.meta);
      setBaseRows(parsed.rows);
      setSearch("");
      setFilter("all");
      setSortMode("excel");

      setMessage(
        `Order file loaded. ${parsed.meta.totalItems} item(s). ${parsed.meta.itemsNeedingOrder} need order. ${parsed.meta.runOutBeforeArrival} may run out before arrival.`
      );

      logUsageEvent("next_order_file_uploaded", {
        module: "generate_next_order",
        ship: parsed.meta.orderShipCode || userShip,
        fileName: file.name,
        sheetName: parsed.meta.sheetName,
        totalItems: parsed.meta.totalItems,
        itemsNeedingOrder: parsed.meta.itemsNeedingOrder,
        runOutBeforeArrival: parsed.meta.runOutBeforeArrival,
        orderDate: parsed.meta.orderDate,
        arrivalDate: parsed.meta.arrivalDate,
        voyageDays: parsed.meta.voyageDays,
        daysUntilArrival: parsed.meta.daysUntilArrival,
      });
    } catch (error) {
      setOrderFileName(file.name);
      setOrderMeta(null);
      setBaseRows([]);

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

    setLoading(true);
    setMessage("Loading yearly regional consumption file...");

    try {
      await waitForBrowser(30);

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsed = parseYearlyRegionalConsumptionWorkbook(workbook);

      setYearlyRegionalConsumption(parsed);
      setYearlyRegionalFileName(file.name);
      setSelectedRegionalConsumptionRegion("");

      setMessage(
        `Yearly regional file loaded. ${parsed.aggregates?.length || 0} product record(s), ${
          parsed.regionOptions?.length || 0
        } region(s).`
      );

      logUsageEvent("yearly_regional_consumption_uploaded_in_next_order", {
        module: "generate_next_order",
        ship: userShip,
        fileName: file.name,
        regions: parsed.regionOptions || [],
        rows: parsed.rows?.length || 0,
      });
    } catch (error) {
      const text =
        error?.message || "Could not load yearly regional consumption file.";

      setMessage(text);
      window.alert(text);
    } finally {
      setLoading(false);
      event.target.value = "";
    }
  };

  const exportToExcel = () => {
    if (!rowsWithRegional.length) {
      window.alert("Upload the latest order workbook first.");
      return;
    }

    const exportRows = rowsWithRegional.map((item, index) => ({
      Line: index + 1,
      ExcelRow: item.excelRow,
      Ship: item.orderShipCode || "",
      Code: item.code || "",
      Product: item.product || "",
      UOM: item.uom || "",

      StockOnHand_D: Number(item.stockOnHand || 0),
      FutureOrders_F_to_N: Number(item.futureOrders || 0),
      PastConsumption_AI_to_AN: Number(item.pastConsumption || 0),
      CurrentPar_Q: Number(item.currentFileParLevel || 0),

      AvgPerDayFromOrder: Number(item.averageConsumptionPerDay || 0),
      DaysUntilArrival: Number(item.arrivalLeadDays || 0),
      VoyageDays: Number(item.voyageDays || 0),

      ConsumptionUntilArrival: Number(item.consumptionUntilArrival || 0),
      EstimatedQtyAtArrival: Number(item.estimatedQtyAtArrival || 0),
      UsableQtyAtArrivalForOrder: Number(item.usableQtyAtArrival || 0),

      VoyageNeedBeforeBuffer: Number(item.voyageNeedBeforeBuffer || 0),
      BufferPercent: Number(item.orderBufferPercent || 0),
      BufferQty: Number(item.orderBufferQty || 0),
      TargetQtyForVoyage: Number(item.targetQtyForVoyage || 0),

      SuggestedNextOrder: Number(item.suggestedOrder || 0),

      SuggestedParFromCurrentUsage: Number(
        item.suggestedParFromCurrentUsage || 0
      ),
      SuggestedParFromLastYearTotal: Number(
        item.suggestedParFromYearTotal || 0
      ),
      SuggestedParFromSelectedMarket: Number(
        item.suggestedParFromMarket || 0
      ),
      SuggestedParFromThisShipLastYear: Number(
        item.suggestedParFromShip || 0
      ),

      YearAvgDaily: Number(item.yearAvgDailyQty || 0),
      MarketAvgDaily: Number(item.marketAvgDailyQty || 0),
      ThisShipLastYearAvgDaily: Number(item.shipYearAvgDailyQty || 0),

      RegionalMatchedProductCode: item.regionalMatchedProductCode || "",
      RegionalMatchedProductName: item.regionalMatchedProductName || "",
      RegionalMatchRows: Number(item.regionalMatchCount || 0),

      Alert: item.alertLabel || "",
      Reason: item.alertDescription || "",
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Next Order");

    XLSX.writeFile(
      workbook,
      `next-order-${orderMeta?.orderShipCode || userShip || "ship"}.xlsx`
    );

    logUsageEvent("next_order_exported", {
      module: "generate_next_order",
      ship: orderMeta?.orderShipCode || userShip,
      rows: exportRows.length,
    });
  };

  const printReport = () => {
    if (!visibleRows.length) {
      window.alert("No rows to print.");
      return;
    }

    const html = `
      <html>
        <head>
          <title>Generated Next Order</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; color: #111; }
            h1 { margin-bottom: 4px; }
            .meta { margin: 3px 0; color: #555; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; margin-top: 18px; font-size: 11px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .qty { color: #0057b8; font-weight: bold; }
            .bad { color: #b00020; font-weight: bold; }
            .warning { margin-top: 14px; padding: 10px; background: #fff4d6; border-radius: 8px; color: #8a5a00; font-weight: bold; }
          </style>
        </head>

        <body>
          <h1>Generated Next Order</h1>

          <div class="meta">File: ${escapeHtml(orderFileName || "N/A")}</div>
          <div class="meta">Ship: ${escapeHtml(
            getShipDisplayName(orderMeta?.orderShipCode) ||
              orderMeta?.orderShipCode ||
              userShip ||
              "N/A"
          )}</div>
          <div class="meta">Order date: ${escapeHtml(
            orderMeta?.orderDate || "N/A"
          )}</div>
          <div class="meta">Arrival date: ${escapeHtml(
            orderMeta?.arrivalDate || "N/A"
          )}</div>
          <div class="meta">Days until arrival: ${formatQty(
            orderMeta?.daysUntilArrival
          )}</div>
          <div class="meta">Voyage days: ${formatQty(
            orderMeta?.voyageDays
          )}</div>
          <div class="meta">Buffer: ${ORDER_BUFFER_PERCENT}% of voyage need only</div>

          <div class="warning">
            Calculation: stock on hand + future orders - consumption until arrival = estimated stock on arrival.
            Suggested order = voyage need + 25% buffer - estimated usable stock on arrival.
          </div>

          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Code</th>
                <th>Product</th>
                <th>UOM</th>
                <th>Avg/day</th>
                <th>Arrival stock</th>
                <th>Voyage need</th>
                <th>25% buffer</th>
                <th>Target</th>
                <th>Suggested order</th>
                <th>Year avg/day</th>
                <th>Market avg/day</th>
                <th>Alert</th>
              </tr>
            </thead>

            <tbody>
              ${visibleRows
                .map(
                  (item, index) => `
                    <tr>
                      <td>${index + 1}</td>
                      <td>${escapeHtml(item.code || "")}</td>
                      <td>${escapeHtml(item.product || "")}</td>
                      <td>${escapeHtml(item.uom || "")}</td>
                      <td>${formatQty(item.averageConsumptionPerDay)}</td>
                      <td class="${
                        Number(item.estimatedQtyAtArrival || 0) < 0
                          ? "bad"
                          : ""
                      }">${formatQty(item.estimatedQtyAtArrival)}</td>
                      <td>${formatQty(item.voyageNeedBeforeBuffer)}</td>
                      <td>${formatQty(item.orderBufferQty)}</td>
                      <td>${formatQty(item.targetQtyForVoyage)}</td>
                      <td class="qty">${formatQty(item.suggestedOrder)}</td>
                      <td>${formatQty(item.yearAvgDailyQty)}</td>
                      <td>${formatQty(item.marketAvgDailyQty)}</td>
                      <td>${escapeHtml(item.alertLabel || "")}</td>
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

    logUsageEvent("next_order_printed", {
      module: "generate_next_order",
      ship: orderMeta?.orderShipCode || userShip,
      rows: visibleRows.length,
    });
  };

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
            🛒 Generate Next Order{" "}
            {orderMeta?.orderShipCode
              ? `• ${getShipDisplayName(orderMeta.orderShipCode)}`
              : userShip
              ? `• ${getShipDisplayName(userShip)}`
              : ""}
          </div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Latest Order Workbook</h2>

          <label style={styles.label}>Latest order file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadOrderFile}
            style={styles.fileInput}
            disabled={loading}
          />

          <div style={styles.infoBox}>
            <div>
              📄 Order file: <strong>{orderFileName || "Not uploaded"}</strong>
            </div>

            <div>
              📋 Sheet: <strong>{orderMeta?.sheetName || "N/A"}</strong>
            </div>

            <div>
              🚢 Ship detected:{" "}
              <strong>
                {orderMeta?.orderShipCode
                  ? `${getShipDisplayName(orderMeta.orderShipCode)} (${orderMeta.orderShipCode})`
                  : "N/A"}
              </strong>
            </div>

            <div>
              📅 Order date: <strong>{orderMeta?.orderDate || "N/A"}</strong>
            </div>

            <div>
              📅 Arrival date:{" "}
              <strong>{orderMeta?.arrivalDate || "N/A"}</strong>
            </div>

            <div>
              ⏳ Days until arrival:{" "}
              <strong>{formatQty(orderMeta?.daysUntilArrival)}</strong>
            </div>

            <div>
              🚢 Voyage days:{" "}
              <strong>{formatQty(orderMeta?.voyageDays)}</strong>
            </div>

            <div>
              🛡️ Buffer: <strong>{ORDER_BUFFER_PERCENT}% of voyage need only</strong>
            </div>
          </div>

          {message && <p style={styles.message}>{message}</p>}
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🌎 Last Year / Market Comparison</h2>

          <div style={styles.infoBox}>
            <div>
              📄 Yearly regional file:{" "}
              <strong>{yearlyRegionalFileName || "Not loaded"}</strong>
            </div>

            <div>
              📊 Regional rows:{" "}
              <strong>{yearlyRegionalConsumption?.rows?.length || 0}</strong>
            </div>

            <div>
              ✅ Matched with order:{" "}
              <strong>
                {filterCounts.regionalMatched} / {filterCounts.all}
              </strong>
            </div>
          </div>

          <label style={styles.label}>Upload / replace yearly regional file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadYearlyRegionalFile}
            style={styles.fileInput}
            disabled={loading}
          />

          <label style={styles.label}>Region / market</label>
          <select
            value={selectedRegionalConsumptionRegion}
            onChange={(event) =>
              setSelectedRegionalConsumptionRegion(event.target.value)
            }
            style={styles.select}
          >
            <option value="">Choose region / market</option>
            <option value={YEARLY_REGION_ALL}>All regions</option>

            {(yearlyRegionalConsumption?.regionOptions || []).map((region) => (
              <option key={region} value={region}>
                {region}
              </option>
            ))}
          </select>

          <p style={styles.emptyText}>
            The order calculation uses current order average consumption. Last
            year and market data are shown for par comparison only.
          </p>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔎 Search / Filter</h2>

          <input
            placeholder="Search code, product, UOM, alert..."
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            style={styles.searchInput}
          />

          <label style={styles.label}>Filter</label>
          <select
            value={filter}
            onChange={(event) => setFilter(event.target.value)}
            style={styles.select}
          >
            <option value="all">All ({filterCounts.all})</option>
            <option value="needsOrder">
              Needs order ({filterCounts.needsOrder})
            </option>
            <option value="noOrder">No order ({filterCounts.noOrder})</option>
            <option value="runOut">
              May run out before arrival ({filterCounts.runOut})
            </option>
            <option value="regionalMatched">
              Has last-year match ({filterCounts.regionalMatched})
            </option>
          </select>

          <label style={styles.label}>Sort</label>
          <select
            value={sortMode}
            onChange={(event) => setSortMode(event.target.value)}
            style={styles.select}
          >
            <option value="excel">Excel row order</option>
            <option value="suggested">Highest suggested order</option>
            <option value="arrivalStock">Lowest arrival stock</option>
            <option value="product">Product A-Z</option>
          </select>

          <div style={styles.headerActions}>
            <button
              type="button"
              style={styles.backButton}
              onClick={printReport}
              disabled={!visibleRows.length || loading}
            >
              🖨️ Print
            </button>

            <button
              type="button"
              style={styles.primaryButton}
              onClick={exportToExcel}
              disabled={!rowsWithRegional.length || loading}
            >
              📥 Export Excel
            </button>
          </div>
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
            <h2 style={styles.productTitle}>🛒 Suggested Next Order</h2>

            <p style={{ ...styles.emptyText, margin: 0 }}>
              Formula: estimated arrival stock first, then voyage need + 25%
              buffer.
            </p>
          </div>

          <div style={styles.shipBadge}>{visibleRows.length} item(s)</div>
        </div>

        {!rowsWithRegional.length && (
          <p style={styles.emptyText}>
            Upload the latest order workbook to generate suggested order lines.
          </p>
        )}

        {rowsWithRegional.length > 0 && visibleRows.length === 0 && (
          <p style={styles.emptyText}>No items match the current filter.</p>
        )}

        <div style={styles.nextOrderGrid || localStyles.nextOrderGrid}>
          {visibleRows.map((item, index) => {
            const cardStyle = {
              ...(styles.nextOrderCard || localStyles.nextOrderCard),
              ...(item.alertType === "red"
                ? styles.nextOrderCardRed || localStyles.nextOrderCardRed
                : {}),
              ...(item.alertType === "blue" || item.alertType === "order"
                ? styles.nextOrderCardBlue || localStyles.nextOrderCardBlue
                : {}),
            };

            return (
              <article
                key={`${item.excelRow}-${item.code}-${index}`}
                style={cardStyle}
              >
                <div style={styles.nextOrderTopLine || localStyles.topLine}>
                  <span>Row {item.excelRow}</span>
                  <span>{item.uom || "UOM"}</span>
                </div>

                <div style={styles.nextOrderName || localStyles.orderName}>
                  {item.product}
                </div>

                <div style={styles.nextOrderMeta || localStyles.meta}>
                  Code: {item.code || "N/A"}
                </div>

                <div style={localStyles.compactCalcBox}>
                  <div>
                    Stock on hand D: <strong>{formatQty(item.stockOnHand)}</strong>
                  </div>

                  <div>
                    Future orders F:N:{" "}
                    <strong>{formatQty(item.futureOrders)}</strong>
                  </div>

                  <div>
                    Past consumption AI:AN:{" "}
                    <strong>{formatQty(item.pastConsumption)}</strong>
                  </div>

                  <div>
                    Avg / day from order:{" "}
                    <strong>{formatQty(item.averageConsumptionPerDay)}</strong>
                  </div>

                  <div>
                    Days cover:{" "}
                    <strong>
                      {formatQty(item.totalProjectionDays)} ={" "}
                      {formatQty(item.arrivalLeadDays)} until arrival +{" "}
                      {formatQty(item.voyageDays)} voyage
                    </strong>
                  </div>

                  <div>
                    Use before arrival:{" "}
                    <strong>{formatQty(item.consumptionUntilArrival)}</strong>
                  </div>

                  <div>
                    Est. stock on arrival:{" "}
                    <strong
                      style={{
                        color:
                          Number(item.estimatedQtyAtArrival || 0) < 0
                            ? "#b00020"
                            : "#2e7d32",
                      }}
                    >
                      {formatQty(item.estimatedQtyAtArrival)}
                    </strong>
                  </div>

                  <div>
                    Voyage need:{" "}
                    <strong>{formatQty(item.voyageNeedBeforeBuffer)}</strong>
                  </div>

                  <div>
                    25% buffer:{" "}
                    <strong>{formatQty(item.orderBufferQty)}</strong>
                  </div>
                </div>

                <div style={localStyles.parCompareBox}>
                  <strong>Par comparison</strong>

                  <div>
                    Current par Q:{" "}
                    <strong>{formatQty(item.currentFileParLevel)}</strong>
                  </div>

                  <div>
                    Current usage par:{" "}
                    <strong>
                      {formatQty(item.suggestedParFromCurrentUsage)}
                    </strong>
                  </div>

                  <div>
                    Year avg/day:{" "}
                    <strong>{formatQty(item.yearAvgDailyQty)}</strong>
                  </div>

                  <div>
                    Market avg/day:{" "}
                    <strong>{formatQty(item.marketAvgDailyQty)}</strong>
                  </div>

                  <div>
                    Market suggested par:{" "}
                    <strong>{formatQty(item.suggestedParFromMarket)}</strong>
                  </div>
                </div>

                <div
                  style={
                    item.suggestedOrder > 0
                      ? styles.nextOrderSuggestedBlue ||
                        localStyles.suggestedBlue
                      : styles.nextOrderSuggestedNeutral ||
                        localStyles.suggestedNeutral
                  }
                >
                  Suggested next order: {formatQty(item.suggestedOrder)}
                </div>

                <div
                  style={
                    item.alertType === "red"
                      ? styles.nextOrderStatusRed || localStyles.statusRed
                      : item.alertType === "blue" || item.alertType === "order"
                      ? styles.nextOrderStatusBlue || localStyles.statusBlue
                      : styles.nextOrderStatusNeutral || localStyles.statusNeutral
                  }
                >
                  {item.alertLabel}
                </div>

                <button
                  type="button"
                  style={styles.backButton}
                  onClick={() => setSelectedItem(item)}
                >
                  🔎 Details / Yearly
                </button>
              </article>
            );
          })}
        </div>
      </section>

      {selectedItem && (
        <div
          style={styles.modalBackdrop}
          onClick={() => setSelectedItem(null)}
        >
          <div
            style={localStyles.detailsModalCard}
            onClick={(event) => event.stopPropagation()}
          >
            <button
              type="button"
              style={styles.closeButton}
              onClick={() => setSelectedItem(null)}
            >
              ✕
            </button>

            <h2 style={styles.productTitle}>{selectedItem.product}</h2>

            <p style={styles.recipeMeta}>
              Code: {selectedItem.code || "N/A"} • UOM:{" "}
              {selectedItem.uom || "N/A"} • Excel row:{" "}
              {selectedItem.excelRow}
            </p>

            <section style={localStyles.detailsGrid}>
              <div style={styles.infoBox}>
                <strong>1. Arrival stock calculation</strong>

                <div>
                  Stock on hand D:{" "}
                  <strong>{formatQty(selectedItem.stockOnHand)}</strong>
                </div>

                <div>
                  Future orders F:N:{" "}
                  <strong>{formatQty(selectedItem.futureOrders)}</strong>
                </div>

                <div>
                  Avg/day from order:{" "}
                  <strong>
                    {formatQty(selectedItem.averageConsumptionPerDay)}
                  </strong>
                </div>

                <div>
                  Days until arrival:{" "}
                  <strong>{formatQty(selectedItem.arrivalLeadDays)}</strong>
                </div>

                <div>
                  Use before arrival:{" "}
                  <strong>
                    {formatQty(selectedItem.consumptionUntilArrival)}
                  </strong>
                </div>

                <div>
                  Estimated stock on arrival:{" "}
                  <strong
                    style={{
                      color:
                        Number(selectedItem.estimatedQtyAtArrival || 0) < 0
                          ? "#b00020"
                          : "#2e7d32",
                    }}
                  >
                    {formatQty(selectedItem.estimatedQtyAtArrival)}
                  </strong>
                </div>

                <div>
                  Usable arrival stock for order:{" "}
                  <strong>{formatQty(selectedItem.usableQtyAtArrival)}</strong>
                </div>
              </div>

              <div style={styles.infoBox}>
                <strong>2. Voyage order need</strong>

                <div>
                  Voyage days:{" "}
                  <strong>{formatQty(selectedItem.voyageDays)}</strong>
                </div>

                <div>
                  Voyage need before buffer:{" "}
                  <strong>
                    {formatQty(selectedItem.voyageNeedBeforeBuffer)}
                  </strong>
                </div>

                <div>
                  25% buffer of voyage need:{" "}
                  <strong>{formatQty(selectedItem.orderBufferQty)}</strong>
                </div>

                <div>
                  Target stock for voyage:{" "}
                  <strong>{formatQty(selectedItem.targetQtyForVoyage)}</strong>
                </div>

                <div>
                  Suggested next order:{" "}
                  <strong>{formatQty(selectedItem.suggestedOrder)}</strong>
                </div>
              </div>

              <div style={styles.infoBox}>
                <strong>3. Current par comparison</strong>

                <div>
                  Current par in file Q:{" "}
                  <strong>{formatQty(selectedItem.currentFileParLevel)}</strong>
                </div>

                <div>
                  Suggested par from current usage:{" "}
                  <strong>
                    {formatQty(selectedItem.suggestedParFromCurrentUsage)}
                  </strong>
                </div>

                <div>
                  Difference vs current par:{" "}
                  <strong
                    style={{
                      color:
                        Number(selectedItem.parDifferenceVsCurrent || 0) > 0
                          ? "#b00020"
                          : "#2e7d32",
                    }}
                  >
                    {formatQty(selectedItem.parDifferenceVsCurrent)}
                  </strong>
                </div>
              </div>

              <div style={styles.infoBox}>
                <strong>4. Last year / market comparison</strong>

                <div>
                  Matched yearly product:{" "}
                  <strong>
                    {selectedItem.regionalMatchedProductName ||
                      selectedItem.regionalMatchedProductCode ||
                      "No match"}
                  </strong>
                </div>

                <div>
                  Match rows:{" "}
                  <strong>{selectedItem.regionalMatchCount || 0}</strong>
                </div>

                <div>
                  Previous year total qty:{" "}
                  <strong>{formatQty(selectedItem.yearTotalQty)}</strong>
                </div>

                <div>
                  Previous year total days:{" "}
                  <strong>{formatQty(selectedItem.yearTotalDays)}</strong>
                </div>

                <div>
                  Previous year avg/day:{" "}
                  <strong>{formatQty(selectedItem.yearAvgDailyQty)}</strong>
                </div>

                <div>
                  Suggested par from previous year total:{" "}
                  <strong>
                    {formatQty(selectedItem.suggestedParFromYearTotal)}
                  </strong>
                </div>

                <hr style={localStyles.hr} />

                <div>
                  Selected market:{" "}
                  <strong>
                    {!selectedRegionalConsumptionRegion
                      ? "Not selected"
                      : selectedRegionalConsumptionRegion === YEARLY_REGION_ALL
                      ? "All regions"
                      : selectedRegionalConsumptionRegion}
                  </strong>
                </div>

                <div>
                  Market total qty:{" "}
                  <strong>{formatQty(selectedItem.marketTotalQty)}</strong>
                </div>

                <div>
                  Market total days:{" "}
                  <strong>{formatQty(selectedItem.marketTotalDays)}</strong>
                </div>

                <div>
                  Market avg/day:{" "}
                  <strong>{formatQty(selectedItem.marketAvgDailyQty)}</strong>
                </div>

                <div>
                  Suggested par from selected market:{" "}
                  <strong>{formatQty(selectedItem.suggestedParFromMarket)}</strong>
                </div>

                <hr style={localStyles.hr} />

                <div>
                  This ship last-year avg/day:{" "}
                  <strong>{formatQty(selectedItem.shipYearAvgDailyQty)}</strong>
                </div>

                <div>
                  Suggested par from this ship last year:{" "}
                  <strong>{formatQty(selectedItem.suggestedParFromShip)}</strong>
                </div>
              </div>
            </section>

            <div style={styles.warningText}>
              Suggested order uses current order average/day only. Last year and
              market numbers are for par review and decision support.
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

const localStyles = {
  nextOrderGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
    gap: 12,
  },

  nextOrderCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 12,
    background: "#fafafa",
    display: "grid",
    gap: 8,
    textAlign: "left",
  },

  nextOrderCardBlue: {
    border: "2px solid #0057b8",
    background: "#eef5ff",
  },

  nextOrderCardRed: {
    border: "2px solid #b00020",
    background: "#fff0f0",
  },

  topLine: {
    display: "flex",
    justifyContent: "space-between",
    gap: 8,
    color: "#666",
    fontSize: 12,
    fontWeight: "bold",
  },

  orderName: {
    fontWeight: "bold",
    fontSize: 15,
    lineHeight: 1.15,
  },

  meta: {
    color: "#555",
    fontSize: 13,
  },

  compactCalcBox: {
    display: "grid",
    gap: 5,
    padding: 9,
    borderRadius: 12,
    background: "#fff",
    border: "1px solid #d8e7ff",
    fontSize: 13,
  },

  parCompareBox: {
    display: "grid",
    gap: 5,
    padding: 9,
    borderRadius: 12,
    background: "#f2f2f2",
    border: "1px solid #e1e1e1",
    fontSize: 13,
  },

  suggestedBlue: {
    padding: 9,
    borderRadius: 10,
    background: "#0057b8",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
  },

  suggestedNeutral: {
    padding: 9,
    borderRadius: 10,
    background: "#f2f2f2",
    color: "#555",
    fontWeight: "bold",
    textAlign: "center",
  },

  statusBlue: {
    padding: 8,
    borderRadius: 10,
    background: "#0057b8",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
  },

  statusRed: {
    padding: 8,
    borderRadius: 10,
    background: "#b00020",
    color: "#fff",
    fontWeight: "bold",
    textAlign: "center",
  },

  statusNeutral: {
    padding: 8,
    borderRadius: 10,
    background: "#f2f2f2",
    color: "#555",
    fontWeight: "bold",
    textAlign: "center",
  },

  detailsModalCard: {
    background: "#fff",
    borderRadius: 18,
    padding: 22,
    maxWidth: 1120,
    width: "96%",
    maxHeight: "90vh",
    overflowY: "auto",
    position: "relative",
    boxShadow: "0 20px 60px rgba(0,0,0,0.25)",
  },

  detailsGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
    gap: 14,
    marginBottom: 16,
  },

  hr: {
    border: 0,
    borderTop: "1px solid #ddd",
    margin: "8px 0",
  },
};

"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

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
    total += Number(row?.[i] || 0);
  }

  return total;
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
  const strongWords = aWords.filter((word) => !["THE", "AND", "WITH", "FOR", "FRESH", "FROZEN", "CASE", "PACK"].includes(word));

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
  const sailors = Number(getCellValue(worksheet, "B5") || 0);
  const voyageDays = Number(getCellValue(worksheet, "B6") || 0);
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

    const stock = Number(row[3] || 0);
    const futureOrders = sumRowRange(row, 5, 13);
    const parLevel = Number(row[16] || 0);
    const pastConsumption = sumRowRange(row, 34, 39);

    let historicalSailorDays = 0;
    for (let c = 34; c <= 39; c += 1) {
      historicalSailorDays += Number(historicalDays[c] || 0) * Number(historicalSailors[c] || 0);
    }

    const consumptionPerSailorDay = historicalSailorDays > 0 ? pastConsumption / historicalSailorDays : 0;
    const averagePerDay = consumptionPerSailorDay * sailors;
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
      parCapApplied,
      parCapLimit,
      alertType,
      alertLabel,
    };

    orderRows.push(item);
    if (code) orderByCode[cleanText(code)] = item;
  }

  await yieldToBrowser();

  const fmlRows = parseFmlRows(workbook);

  const counts = {
    totalItems: orderRows.length,
    itemsNeedingOrder: orderRows.filter((row) => row.suggestedOrder > 0).length,
    noConsumptionItems: orderRows.filter((row) => row.pastConsumption === 0).length,
    noStockItems: orderRows.filter((row) => row.stock === 0).length,
    parCapsApplied: orderRows.filter((row) => row.parCapApplied).length,
    negativeArrival: orderRows.filter((row) => row.availableAtArrival < 0).length,
    fmlNotUsed: 0,
    fmlRunningLow: 0,
  };

  return {
    workbook,
    orderRows,
    orderByCode,
    fmlRows,
    fmlNotUsedRows: [],
    fmlRunningLowRows: [],
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
    if (fmlItem.code && orderByCode[cleanText(fmlItem.code)]) {
      return orderByCode[cleanText(fmlItem.code)];
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

export default function GenerateNextOrder({ styles, userShip, onBack, logUsageEvent = () => {} }) {
  const [templateEntries, setTemplateEntries] = useState([]);
  const [templateStatus, setTemplateStatus] = useState("Loading attached ERP template...");
  const [nextOrderRows, setNextOrderRows] = useState([]);
  const [orderByCode, setOrderByCode] = useState({});
  const [fmlSourceRows, setFmlSourceRows] = useState([]);
  const [fmlNotUsedRows, setFmlNotUsedRows] = useState([]);
  const [fmlRunningLowRows, setFmlRunningLowRows] = useState([]);
  const [nextOrderMeta, setNextOrderMeta] = useState({});
  const [nextOrderFileName, setNextOrderFileName] = useState("");
  const [nextOrderSearch, setNextOrderSearch] = useState("");
  const [fmlSearch, setFmlSearch] = useState("");
  const [fmlLowSearch, setFmlLowSearch] = useState("");
  const [nextOrderFilter, setNextOrderFilter] = useState("all");
  const [nextOrderView, setNextOrderView] = useState("order");
  const [nextOrderLoading, setNextOrderLoading] = useState(false);
  const [fmlReportLoading, setFmlReportLoading] = useState(false);
  const [fmlReportMessage, setFmlReportMessage] = useState("");
  const [fmlNotUsedPrepared, setFmlNotUsedPrepared] = useState(false);
  const [fmlRunningLowPrepared, setFmlRunningLowPrepared] = useState(false);
  const [nextOrderMessage, setNextOrderMessage] = useState("");

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
      } catch (error) {
        setTemplateStatus("Could not load attached ERP template.");
      }
    };

    loadDefaultTemplate();
  }, []);

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
      } catch (error) {
        setTemplateStatus("Could not load custom ERP template.");
      }
    };

    reader.readAsBinaryString(file);
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
      const parsed = await parseOrderFile(file, templateEntries);

      setNextOrderRows(parsed.orderRows);
      setOrderByCode(parsed.orderByCode || {});
      setFmlSourceRows(parsed.fmlRows || []);
      setFmlNotUsedRows([]);
      setFmlRunningLowRows([]);
      setFmlNotUsedPrepared(false);
      setFmlRunningLowPrepared(false);
      setNextOrderMeta(parsed.meta);
      setNextOrderFileName(file.name);
      setNextOrderSearch("");
      setFmlSearch("");
      setFmlLowSearch("");
      setNextOrderFilter("all");
      setNextOrderView("order");
      setNextOrderMessage(
        "Order file loaded. " +
          parsed.meta.totalItems +
          " product rows found. " +
          parsed.meta.itemsNeedingOrder +
          " need order. Calculated using B2/B3 days to arrival, B5 sailors, B6 voyage days, F:N future orders, AI:AN past consumption, and Q par where applicable. FML reports are preparing in the background."
      );

      logUsageEvent("next_order_file_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        ship: parsed.meta.shipCode,
        totalItems: parsed.meta.totalItems,
        fmlRows: (parsed.fmlRows || []).length,
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
      setFmlNotUsedRows([]);
      setFmlRunningLowRows([]);
      setFmlNotUsedPrepared(false);
      setFmlRunningLowPrepared(false);
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

    return nextOrderRows.filter((row) => {
      const filterOk =
        nextOrderFilter === "all" ||
        (nextOrderFilter === "needs" && row.suggestedOrder > 0) ||
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
  }, [nextOrderRows, nextOrderSearch, nextOrderFilter]);

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

  const countNeedsOrder = nextOrderRows.filter((row) => row.suggestedOrder > 0).length;
  const countNoConsumption = nextOrderRows.filter((row) => row.pastConsumption === 0).length;
  const countNoStock = nextOrderRows.filter((row) => row.stock === 0).length;

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
            {nextOrderMessage && <div style={{ color: nextOrderMessage.includes("Could not") ? "#b00020" : "#555" }}>{nextOrderMessage}</div>}
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
            {filteredNextOrderRows.map((row) => {
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
            {filteredFmlNotUsedRows.map((row, index) => (
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
            {filteredFmlRunningLowRows.map((row, index) => (
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
};

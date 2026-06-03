"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

import {
  YEARLY_REGION_ALL,
  parseYearlyRegionalConsumptionWorkbook,
  formatRegionalQty,
} from "../../lib/yearlyRegionalConsumption";

const SHIPS = ["SC", "VL", "BRL", "RL"];

const SHIP_DISPLAY_NAMES = {
  SC: "Scarlet",
  VL: "Valiant",
  BRL: "Brilliant",
  RL: "Resilient",
};

const cleanText = (value) =>
  String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

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

const formatQty = (value) => Number(value || 0).toFixed(2);
const formatMoney = (value) => "$" + Number(value || 0).toFixed(2);

const getShipDisplayName = (shipCode) =>
  SHIP_DISPLAY_NAMES[shipCode] || shipCode || "";

const normalizeShipCode = (value) => {
  const text = cleanText(value)
    .replace(/RESILIANT/g, "RESILIENT")
    .replace(/\bV\s*[-]?\s*1\b/g, "V1")
    .replace(/\bS\s*C\s*L\b/g, "SCL");

  if (!text) return "";

  if (text === "BRL" || text.includes("BRILLIANT")) return "BRL";
  if (text === "RL" || text.includes("RESILIENT")) return "RL";
  if (text === "VL" || text.includes("VALIANT")) return "VL";

  if (
    text === "SC" ||
    text === "SCL" ||
    text === "V1" ||
    text.includes("SCARLET") ||
    text.includes("SCL") ||
    /\bV\s*[-]?\s*1\b/.test(text)
  ) {
    return "SC";
  }

  return "";
};

const normalizeVenue = (value) =>
  cleanText(value)
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\s*-\s*VV$/g, "")
    .replace(/\s*VV$/g, "")
    .replace(/\bTHE\s+/g, "")
    .replace(/\bSCL\b/g, "")
    .replace(/\bSC\b/g, "")
    .replace(/\bVAL\b/g, "")
    .replace(/\bVL\b/g, "")
    .replace(/\bRES\b/g, "")
    .replace(/\bRL\b/g, "")
    .replace(/\bBRL\b/g, "")
    .replace(/\bV1\b/g, "")
    .replace(/\bROJO\b/g, "")
    .replace(/\bARIYA\b/g, "")
    .replace(/\bONLY\b/g, "")
    .replace(/\bMANNOR\b/g, "MANOR")
    .replace(/\s+/g, " ")
    .trim();

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

const cleanTemplateTitle = (value) =>
  String(value || "")
    .replace(/\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/g, "")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*$/g, "")
    .trim();

const cleanTemplateSheetDisplay = (sheetName) =>
  String(sheetName || "")
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\bSCL\b/gi, "")
    .replace(/\bSC\b/gi, "")
    .replace(/\bV\s*[-]?\s*1\b/gi, "")
    .replace(/\bVAL\b/gi, "")
    .replace(/\bVL\b/gi, "")
    .replace(/\bRES\b/gi, "")
    .replace(/\bRL\b/gi, "")
    .replace(/\bBRL\b/gi, "")
    .replace(/\bROJO\b/gi, "")
    .replace(/\bARIYA\b/gi, "")
    .replace(/\bONLY\b/gi, "")
    .replace(/\s*-\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim();

const getTemplateSectionName = (templateName) => {
  const cleaned = cleanTemplateTitle(templateName);
  const parts = cleaned
    .split(/\s*-\s*/)
    .map((part) => part.trim())
    .filter(Boolean);

  return parts.length > 1 ? parts[parts.length - 1] : cleaned;
};

const getTemplateLocationDisplay = (sheetName, templateName) => {
  const sheetDisplay = cleanTemplateSheetDisplay(sheetName);
  const sectionName = getTemplateSectionName(templateName);

  if (!sheetDisplay && !sectionName) return "Template";
  if (!sheetDisplay) return sectionName;
  if (!sectionName) return sheetDisplay;

  const sheetKey = normalizeVenue(sheetDisplay);
  const sectionKey = normalizeVenue(sectionName);

  if (!sectionKey || sheetKey === sectionKey || sheetKey.includes(sectionKey)) {
    return sheetDisplay;
  }

  return `${sheetDisplay} - ${sectionName}`;
};

const getTemplateLocationKey = (sheetName, templateName) =>
  normalizeVenue(getTemplateLocationDisplay(sheetName, templateName));

const getTemplateSheetShipScope = (sheetName) => {
  const text = cleanText(sheetName)
    .replace(/RESILIANT/g, "RESILIENT")
    .replace(/\bV\s*[-]?\s*1\b/g, "V1")
    .replace(/\bS\s*C\s*L\b/g, "SCL");

  const scope = [];

  if (
    /\bSCL\b/.test(text) ||
    /\bSC\b/.test(text) ||
    /\bV1\b/.test(text) ||
    text.includes("SCARLET")
  ) {
    scope.push("SC");
  }

  if (
    /\bVAL\b/.test(text) ||
    /\bVL\b/.test(text) ||
    text.includes("VALIANT")
  ) {
    scope.push("VL");
  }

  if (
    /\bRES\b/.test(text) ||
    /\bRL\b/.test(text) ||
    text.includes("RESILIENT")
  ) {
    scope.push("RL");
  }

  if (/\bBRL\b/.test(text) || text.includes("BRILLIANT")) {
    scope.push("BRL");
  }

  return [...new Set(scope)];
};

const getTemplateShipScopeLabel = (shipScope) => {
  const scope = Array.isArray(shipScope) ? shipScope.filter(Boolean) : [];
  return scope.length ? "Used only on " + scope.join(", ") : "Used by all ships";
};

const templateShipScopeMatches = (shipScope, currentShipCode) => {
  const scope = Array.isArray(shipScope) ? shipScope.filter(Boolean) : [];

  if (!scope.length) return true;
  if (!currentShipCode) return false;

  return scope.includes(currentShipCode);
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

const getLooseVenueMatchKey = (value) =>
  normalizeVenue(value).replace(/[^A-Z0-9]/g, "");

const splitFmlVenues = (value) =>
  String(value || "")
    .split(",")
    .map((venue) => venue.replace(/\s+/g, " ").trim())
    .filter(Boolean);

const sanitizeFileName = (value) =>
  String(value || "ship")
    .replace(/[^a-z0-9]+/gi, "-")
    .replace(/^-+|-+$/g, "")
    .toLowerCase()
    .slice(0, 60);

const sx = (...parts) => Object.assign({}, ...parts.filter(Boolean));

const sortNextOrderRows = (rows) =>
  [...rows].sort((a, b) => {
    const rowDiff = Number(a.excelRow || 0) - Number(b.excelRow || 0);

    if (rowDiff !== 0) return rowDiff;

    return String(a.product || "").localeCompare(String(b.product || ""));
  });

const parseTemplateWorkbook = (workbook) => {
  const map = {};

  workbook.SheetNames.forEach((sheetName) => {
    const venueKey = normalizeVenue(sheetName);
    if (!venueKey) return;

    if (!map[venueKey]) map[venueKey] = {};

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
    });

    if (!rows.length) return;

    rows.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (cleanText(cell) !== "INGREDIENT NAME") return;

        const templateName = cleanTemplateTitle(
          rows[rowIndex - 1]?.[colIndex] ||
            rows[rowIndex - 1]?.[colIndex - 1] ||
            sheetName ||
            "Template"
        );

        const shipScope = getTemplateSheetShipScope(sheetName);

        const templateLocation = {
          locationKey: getTemplateLocationKey(sheetName, templateName),
          displayName: getTemplateLocationDisplay(sheetName, templateName),
          sheetName,
          templateName: templateName || sheetName || "Template",
          shipScope,
          shipScopeLabel: getTemplateShipScopeLabel(shipScope),
        };

        rows.slice(rowIndex + 1).forEach((dataRow) => {
          const product = String(dataRow[colIndex] || "").trim();
          if (!product) return;

          const productKey = cleanText(product);
          if (!productKey) return;

          if (
            productKey === "INGREDIENT NAME" ||
            productKey === "CODE" ||
            productKey === "UM" ||
            productKey.includes("#REF")
          ) {
            return;
          }

          if (!map[venueKey][productKey]) {
            map[venueKey][productKey] = {
              product,
              productCodes: new Set(),
              templates: new Set(),
              templateLocations: new Set(),
            };
          }

          const templateCode = String(dataRow[colIndex - 1] || "").trim();

          if (templateCode && cleanText(templateCode) !== "CODE") {
            map[venueKey][productKey].productCodes.add(templateCode);
          }

          map[venueKey][productKey].templates.add(
            templateName || sheetName || "Template"
          );

          map[venueKey][productKey].templateLocations.add(
            JSON.stringify(templateLocation)
          );
        });
      });
    });
  });

  Object.keys(map).forEach((venueKey) => {
    Object.keys(map[venueKey]).forEach((productKey) => {
      map[venueKey][productKey].productCodes = [
        ...(map[venueKey][productKey].productCodes || []),
      ];

      map[venueKey][productKey].templates = [
        ...(map[venueKey][productKey].templates || []),
      ];

      map[venueKey][productKey].templateLocations = [
        ...(map[venueKey][productKey].templateLocations || []),
      ]
        .map((locationText) => {
          try {
            return JSON.parse(locationText);
          } catch {
            return null;
          }
        })
        .filter(Boolean);
    });
  });

  return map;
};

const getTemplateMatchesForFmlProduct = (templateMap, fmlItem, currentShipCode) => {
  const matches = [];
  const seen = new Set();
  const fmlVenues = fmlItem.venues || [];
  const fmlVenueKeys = fmlVenues
    .map((venue) => getLooseVenueMatchKey(venue))
    .filter(Boolean);

  const fmlCodeKey = normalizeOrderCode(fmlItem.code);

  Object.entries(templateMap || {}).forEach(([venueKey, productsByKey]) => {
    Object.values(productsByKey || {}).forEach((templateItem) => {
      const templateCodes = Array.isArray(templateItem.productCodes)
        ? templateItem.productCodes
        : [];

      const codeMatches =
        fmlCodeKey &&
        templateCodes.some((code) => normalizeOrderCode(code) === fmlCodeKey);

      const nameMatches = productNamesMatch(
        fmlItem.product,
        templateItem.product
      );

      if (!codeMatches && !nameMatches) return;

      const locations =
        Array.isArray(templateItem.templateLocations) &&
        templateItem.templateLocations.length
          ? templateItem.templateLocations
          : [
              {
                locationKey: venueKey,
                displayName: venueKey,
                sheetName: "",
                templateName: "",
                shipScope: [],
                shipScopeLabel: "Used by all ships",
              },
            ];

      locations.forEach((location) => {
        const shipScope = Array.isArray(location.shipScope)
          ? location.shipScope
          : [];

        if (!templateShipScopeMatches(shipScope, currentShipCode)) return;

        const candidateKeys = [
          location.locationKey,
          location.displayName,
          location.sheetName,
          venueKey,
        ]
          .map((value) => getLooseVenueMatchKey(value))
          .filter(Boolean);

        const matchedFmlVenueIndexes = fmlVenueKeys
          .map((fmlKey, index) =>
            candidateKeys.some(
              (candidateKey) =>
                candidateKey === fmlKey ||
                candidateKey.includes(fmlKey) ||
                fmlKey.includes(candidateKey)
            )
              ? index
              : -1
          )
          .filter((index) => index >= 0);

        if (!matchedFmlVenueIndexes.length) return;

        const uniqueKey = [
          location.sheetName || venueKey,
          location.templateName || "",
          templateItem.product || "",
          shipScope.join("-") || "ALL",
        ].join("|");

        if (seen.has(uniqueKey)) return;
        seen.add(uniqueKey);

        matches.push({
          templateProduct: templateItem.product || fmlItem.product,
          templateName: location.templateName || "Template",
          sheetName: location.sheetName || "",
          displayName: location.displayName || location.locationKey || venueKey,
          shipScope,
          shipScopeLabel: getTemplateShipScopeLabel(shipScope),
          matchedVenues: [
            ...new Set(
              matchedFmlVenueIndexes
                .map((index) => fmlVenues[index])
                .filter(Boolean)
            ),
          ],
        });
      });
    });
  });

  return matches;
};

const parseFmlNotOrderedUnusedReport = (
  workbook,
  orderRows,
  currentShipCode,
  templateMap
) => {
  const fmlSheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "FML") ||
    workbook.SheetNames.find((name) => cleanText(name).includes("FML"));

  if (!fmlSheetName) return [];

  const ws = workbook.Sheets[fmlSheetName];
  if (!ws) return [];

  const decodedRange = XLSX.utils.decode_range(ws["!ref"] || "A1:I1");

  const fmlRange = {
    s: { r: decodedRange.s.r, c: 0 },
    e: { r: decodedRange.e.r, c: Math.max(decodedRange.e.c, 8) },
  };

  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
    range: fmlRange,
  });

  const orderByCode = {};
  const orderByProductKey = {};

  orderRows.forEach((item) => {
    const codeKey = normalizeOrderCode(item.code);
    const productKey = getProductReportKey(item.product);

    if (codeKey) orderByCode[codeKey] = item;
    if (productKey && !orderByProductKey[productKey]) {
      orderByProductKey[productKey] = item;
    }
  });

  const reportRows = [];
  const seen = new Set();

  rows.slice(3).forEach((row, index) => {
    const excelRow = index + 4;
    const department = String(row[0] || "").trim();
    const category = String(row[1] || "").trim();
    const subCategory = String(row[2] || "").trim();
    const code = String(row[3] || "").trim();
    const product = String(row[4] || "").replace(/\s+/g, " ").trim();
    const venueText = String(row[5] || "").replace(/\s+/g, " ").trim();
    const uom = String(row[8] || "").trim();

    if (!code || !product || !venueText) return;
    if (cleanText(code) === "PRODUCT") return;
    if (cleanText(product) === "PRODUCT NAME") return;

    const venues = splitFmlVenues(venueText);
    if (!venues.length) return;

    const codeKey = normalizeOrderCode(code);
    const productKey = getProductReportKey(product);
    const orderItem = orderByCode[codeKey] || orderByProductKey[productKey] || null;

    if (!orderItem) return;

    const futureOrders = Number(orderItem.futureOrders || 0);
    const pastConsumption = Number(orderItem.pastConsumption || 0);

    if (futureOrders > 0 || pastConsumption > 0) return;

    const templateMatches = getTemplateMatchesForFmlProduct(
      templateMap,
      { code, product, venues },
      currentShipCode
    );

    if (!templateMatches.length) return;

    const uniqueKey = codeKey || productKey || cleanText(product + "|" + excelRow);
    if (seen.has(uniqueKey)) return;
    seen.add(uniqueKey);

    const matchedVenues = [
      ...new Set(templateMatches.flatMap((match) => match.matchedVenues || [])),
    ];

    const templateShipScopeLabels = [
      ...new Set(
        templateMatches.map((match) => match.shipScopeLabel || "Used by all ships")
      ),
    ];

    const templateLocationNames = [
      ...new Set(
        templateMatches.map(
          (match) => match.displayName || match.templateName || "Template"
        )
      ),
    ];

    const templateSheetNames = [
      ...new Set(templateMatches.map((match) => match.sheetName).filter(Boolean)),
    ];

    reportRows.push({
      excelRow,
      standardOrderRow: orderItem?.excelRow || "",
      code,
      product,
      uom: orderItem?.uom || uom || "",
      department,
      category,
      subCategory,
      venues,
      venueText,
      matchedVenues,
      templateMatches,
      templateLocationNames,
      templateSheetNames,
      templateShipScopeLabels,
      templateShipScopeNote: templateShipScopeLabels.join("; "),
      stockOnHand: Number(orderItem?.stockOnHand || 0),
      futureOrders,
      pastConsumption,
      foundInOrderTemplate: Boolean(orderItem),
      foundInTemplate: true,
      currentShipCode,
      reason:
        "FML product matches the ERP template for this ship and has 0 future orders plus 0 past consumption in Standard Order Template.",
    });
  });

  return reportRows.sort(
    (a, b) => Number(a.excelRow || 0) - Number(b.excelRow || 0)
  );
};

const parseFmlRunningLowReport = (
  workbook,
  orderRows,
  currentShipCode,
  templateMap
) => {
  const fmlSheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "FML") ||
    workbook.SheetNames.find((name) => cleanText(name).includes("FML"));

  if (!fmlSheetName) return [];

  const ws = workbook.Sheets[fmlSheetName];
  if (!ws) return [];

  const decodedRange = XLSX.utils.decode_range(ws["!ref"] || "A1:I1");

  const fmlRange = {
    s: { r: decodedRange.s.r, c: 0 },
    e: { r: decodedRange.e.r, c: Math.max(decodedRange.e.c, 8) },
  };

  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
    range: fmlRange,
  });

  const orderByCode = {};
  const orderByProductKey = {};

  orderRows.forEach((item) => {
    const codeKey = normalizeOrderCode(item.code);
    const productKey = getProductReportKey(item.product);

    if (codeKey) orderByCode[codeKey] = item;
    if (productKey && !orderByProductKey[productKey]) {
      orderByProductKey[productKey] = item;
    }
  });

  const reportRows = [];
  const seen = new Set();

  rows.slice(3).forEach((row, index) => {
    const excelRow = index + 4;
    const department = String(row[0] || "").trim();
    const category = String(row[1] || "").trim();
    const subCategory = String(row[2] || "").trim();
    const code = String(row[3] || "").trim();
    const product = String(row[4] || "").replace(/\s+/g, " ").trim();
    const venueText = String(row[5] || "").replace(/\s+/g, " ").trim();
    const uom = String(row[8] || "").trim();

    if (!code || !product || !venueText) return;
    if (cleanText(code) === "PRODUCT") return;
    if (cleanText(product) === "PRODUCT NAME") return;

    const venues = splitFmlVenues(venueText);
    if (!venues.length) return;

    const codeKey = normalizeOrderCode(code);
    const productKey = getProductReportKey(product);
    const orderItem = orderByCode[codeKey] || orderByProductKey[productKey] || null;

    if (!orderItem) return;

    const futureOrders = Number(orderItem.futureOrders || 0);
    const pastConsumption = Number(orderItem.pastConsumption || 0);
    const averageConsumptionPerDay = Number(orderItem.averageConsumptionPerDay || 0);
    const availableAtArrival = Number(orderItem.availableAtArrival || 0);

    if (futureOrders > 0) return;
    if (pastConsumption <= 0 || averageConsumptionPerDay <= 0) return;

    const oneDayBuffer = averageConsumptionPerDay;
    const isRunningLowAtArrival = availableAtArrival <= oneDayBuffer;

    if (!isRunningLowAtArrival) return;

    const templateMatches = getTemplateMatchesForFmlProduct(
      templateMap,
      { code, product, venues },
      currentShipCode
    );

    if (!templateMatches.length) return;

    const uniqueKey = codeKey || productKey || cleanText(product + "|" + excelRow);
    if (seen.has(uniqueKey)) return;
    seen.add(uniqueKey);

    const matchedVenues = [
      ...new Set(templateMatches.flatMap((match) => match.matchedVenues || [])),
    ];

    const templateShipScopeLabels = [
      ...new Set(
        templateMatches.map((match) => match.shipScopeLabel || "Used by all ships")
      ),
    ];

    const templateLocationNames = [
      ...new Set(
        templateMatches.map(
          (match) => match.displayName || match.templateName || "Template"
        )
      ),
    ];

    const templateSheetNames = [
      ...new Set(templateMatches.map((match) => match.sheetName).filter(Boolean)),
    ];

    const daysOfCoverAtArrival =
      averageConsumptionPerDay > 0 ? availableAtArrival / averageConsumptionPerDay : 0;

    const reason =
      availableAtArrival <= 0
        ? "No future order. Based on average daily consumption, this product is expected to be out before or by arrival day."
        : "No future order. Based on average daily consumption, this product will have less than one day of stock at arrival.";

    reportRows.push({
      excelRow,
      standardOrderRow: orderItem?.excelRow || "",
      code,
      product,
      uom: orderItem?.uom || uom || "",
      department,
      category,
      subCategory,
      venues,
      venueText,
      matchedVenues,
      templateMatches,
      templateLocationNames,
      templateSheetNames,
      templateShipScopeLabels,
      templateShipScopeNote: templateShipScopeLabels.join("; "),
      stockOnHand: Number(orderItem?.stockOnHand || 0),
      futureOrders,
      pastConsumption,
      averageConsumptionPerDay,
      consumptionUntilArrival: Number(orderItem?.consumptionUntilArrival || 0),
      availableAtArrival,
      daysOfCoverAtArrival,
      suggestedOrder: Number(orderItem?.suggestedOrder || 0),
      foundInOrderTemplate: Boolean(orderItem),
      foundInTemplate: true,
      currentShipCode,
      reason,
    });
  });

  return reportRows.sort(
    (a, b) => Number(a.excelRow || 0) - Number(b.excelRow || 0)
  );
};

const parseNextOrderWorkbook = ({ workbook, templateMap, fallbackShipCode }) => {
  const sheetName = workbook.SheetNames.includes("Standard Order Template")
    ? "Standard Order Template"
    : workbook.SheetNames.includes("Order Sheet")
    ? "Order Sheet"
    : workbook.SheetNames[0];

  const ws = workbook.Sheets[sheetName];

  if (!ws) {
    throw new Error("Could not find the order worksheet.");
  }

  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
    cellDates: true,
  });

  const orderShipName = String(rows[0]?.[1] || "").trim();
  const orderShipCode =
    normalizeShipCode(orderShipName) || normalizeShipCode(fallbackShipCode);

  const rawOrderDate = rows[1]?.[1];
  const rawArrivalDate = rows[2]?.[1];
  const targetSailors = toNumber(rows[4]?.[1]);
  const targetDays = toNumber(rows[5]?.[1]);
  const daysUntilArrival = getDaysBetweenCells(rawOrderDate, rawArrivalDate);
  const currentPeriodSailorDays = targetSailors * targetDays;

  const futureOrderColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13];
  const pastConsumptionColumns = [34, 35, 36, 37, 38, 39];

  const historicalSailorDays = pastConsumptionColumns.reduce(
    (sum, colIndex) =>
      sum + getHistoricalSailorDays(rows[4]?.[colIndex], rows[5]?.[colIndex]),
    0
  );

  const parsedRows = [];

  rows.slice(9).forEach((row, rowOffset) => {
    const excelRow = rowOffset + 10;
    const code = String(row[0] || "").trim();
    const product = String(row[1] || "").replace(/\s+/g, " ").trim();
    const uom = String(row[2] || "").trim();

    if (!product || !uom) return;
    if (cleanText(code) === "CODE") return;
    if (cleanText(product) === "PRODUCT") return;
    if (cleanText(product) === "PRODUCT NAME") return;

    const stockOnHand = toNumber(row[3]);
    const parLevel = toNumber(row[16]);
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
      averageConsumptionPerSailorDay * targetSailors;

    const projectedNeed = averageConsumptionPerDay * targetDays;
    const consumptionUntilArrival = averageConsumptionPerDay * daysUntilArrival;
    const availableAtArrival = stockOnHand + futureOrders - consumptionUntilArrival;
    const rawSuggestedOrder = Math.max(projectedNeed - availableAtArrival, 0);

    const isFourteenDayLoad = Math.abs(Number(targetDays || 0) - 14) < 0.01;
    const parMaxAllowed = parLevel > 0 ? parLevel * 1.1 : 0;
    const parCapApplied = Boolean(
      isFourteenDayLoad &&
        parLevel > 0 &&
        rawSuggestedOrder > parMaxAllowed
    );

    const suggestedOrder = parCapApplied ? parMaxAllowed : rawSuggestedOrder;

    let parLevelNote = "Par level ignored because B6 is not exactly 14 days.";

    if (isFourteenDayLoad && parLevel > 0 && parCapApplied) {
      parLevelNote =
        "Par cap applied: 14-day load cannot exceed par level Q + 10%.";
    } else if (isFourteenDayLoad && parLevel > 0) {
      parLevelNote =
        "Par level considered: calculated order is within par level Q + 10%.";
    } else if (isFourteenDayLoad && parLevel <= 0) {
      parLevelNote = "14-day load, but no par level found in column Q.";
    }

    const hasNoPastConsumption = pastConsumption <= 0;
    const hasNoStockOnHand = stockOnHand <= 0;

    let alertType = suggestedOrder > 0 ? "order" : "normal";
    let alertLabel = suggestedOrder > 0 ? "Needs order" : "No order suggested";
    let alertDescription =
      "Average daily consumption x voyage days, adjusted for stock/future orders until order arrival. " +
      parLevelNote;

    if (hasNoPastConsumption && hasNoStockOnHand) {
      alertType = "blue";
      alertLabel = "No stock and no past consumption";
      alertDescription =
        "Blue review: stock on hand is 0 and past consumption is 0.";
    } else if (hasNoPastConsumption && stockOnHand > 0) {
      alertType = "red";
      alertLabel = "Stock on hand but no past consumption";
      alertDescription =
        "Red review: item has stock on hand but no past consumption.";
    }

    parsedRows.push({
      excelRow,
      code,
      product,
      uom,
      stockOnHand,
      parLevel,
      futureOrders,
      pastConsumption,
      rawSuggestedOrder,
      parMaxAllowed,
      parCapApplied,
      parLevelNote,
      historicalSailorDays,
      currentPeriodSailorDays,
      daysUntilArrival,
      averageConsumptionPerSailorDay,
      averageConsumptionPerDay,
      projectedNeed,
      consumptionUntilArrival,
      availableAtArrival,
      suggestedOrder,
      alertType,
      alertLabel,
      alertDescription,
      orderReason: alertDescription,
    });
  });

  const sortedRows = sortNextOrderRows(parsedRows);

  const fmlReportRows = parseFmlNotOrderedUnusedReport(
    workbook,
    sortedRows,
    orderShipCode,
    templateMap
  );

  const fmlRunningLowRows = parseFmlRunningLowReport(
    workbook,
    sortedRows,
    orderShipCode,
    templateMap
  );

  return {
    rows: sortedRows,
    fmlReportRows,
    fmlRunningLowRows,
    meta: {
      sheetName,
      shipName: orderShipName || fallbackShipCode || "",
      shipCode: orderShipCode,
      shipDisplayName: getShipDisplayName(orderShipCode),
      orderDate: formatDateCell(rawOrderDate),
      arrivalDate: formatDateCell(rawArrivalDate),
      targetSailors,
      targetDays,
      daysUntilArrival,
      currentPeriodSailorDays,
      historicalSailorDays,
      totalItems: sortedRows.length,
      itemsNeedingOrder: sortedRows.filter(
        (item) => Number(item.suggestedOrder || 0) > 0
      ).length,
      parCapItems: sortedRows.filter((item) => item.parCapApplied).length,
      blueReviewItems: sortedRows.filter((item) => item.alertType === "blue")
        .length,
      redReviewItems: sortedRows.filter((item) => item.alertType === "red")
        .length,
      fmlMissingItems: fmlReportRows.length,
      fmlRunningLowItems: fmlRunningLowRows.length,
      scarletAliasNote:
        orderShipCode === "SC" &&
        /\b(V\s*[-]?\s*1|SCL|SCARLET|SC)\b/i.test(orderShipName || "")
          ? "Scarlet detected from source ship value: " + orderShipName
          : "",
    },
  };
};

export default function GenerateNextOrder({
  styles = {},
  userShip = "",
  onBack = () => {},
  logUsageEvent = () => {},
  yearlyRegionalConsumption = null,
  setYearlyRegionalConsumption = () => {},
  yearlyRegionalFileName = "",
  setYearlyRegionalFileName = () => {},
  selectedRegionalConsumptionRegion = "",
  setSelectedRegionalConsumptionRegion = () => {},
  regionalParBufferPercent = 0,
  setRegionalParBufferPercent = () => {},
}) {
  const [templateMap, setTemplateMap] = useState({});
  const [templateStatus, setTemplateStatus] = useState("Loading default ERP template...");
  const [templateFileName, setTemplateFileName] = useState("Default ERP template");

  const [nextOrderFileName, setNextOrderFileName] = useState("");
  const [nextOrderSourceRows, setNextOrderSourceRows] = useState([]);
  const [nextOrderRows, setNextOrderRows] = useState([]);
  const [nextOrderMeta, setNextOrderMeta] = useState({});

  const [fmlMissingRows, setFmlMissingRows] = useState([]);
  const [fmlLowRows, setFmlLowRows] = useState([]);

  const [nextOrderSearch, setNextOrderSearch] = useState("");
  const [nextOrderFilter, setNextOrderFilter] = useState("all");
  const [nextOrderView, setNextOrderView] = useState("order");

  const [fmlMissingSearch, setFmlMissingSearch] = useState("");
  const [fmlLowSearch, setFmlLowSearch] = useState("");

  const [nextOrderLoading, setNextOrderLoading] = useState(false);
  const [nextOrderMessage, setNextOrderMessage] = useState("");
  const [yearlyRegionalMessage, setYearlyRegionalMessage] = useState("");

  const activeRows = nextOrderRows.length
    ? nextOrderRows
    : sortNextOrderRows(nextOrderSourceRows);

  const filterNextOrderRows = (rows) => {
    const term = nextOrderSearch.toLowerCase().trim();

    return rows.filter((item) => {
      const matchesFilter =
        nextOrderFilter === "all" ||
        (nextOrderFilter === "needsOrder" &&
          Number(item.suggestedOrder || 0) > 0) ||
        (nextOrderFilter === "parCap" && item.parCapApplied) ||
        (nextOrderFilter === "blue" && item.alertType === "blue") ||
        (nextOrderFilter === "red" && item.alertType === "red") ||
        (nextOrderFilter === "noConsumption" &&
          Number(item.pastConsumption || 0) <= 0) ||
        (nextOrderFilter === "noStock" && Number(item.stockOnHand || 0) <= 0);

      if (!matchesFilter) return false;
      if (!term) return true;

      return [
        item.code,
        item.product,
        item.uom,
        item.alertLabel,
        item.parLevelNote,
        item.orderReason,
        item.excelRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  };

  const visibleOrderRows = useMemo(
    () => filterNextOrderRows(activeRows),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [activeRows, nextOrderSearch, nextOrderFilter]
  );

  const visibleFmlMissingRows = useMemo(() => {
    const term = fmlMissingSearch.toLowerCase().trim();
    if (!term) return fmlMissingRows;

    return fmlMissingRows.filter((item) =>
      [
        item.code,
        item.product,
        item.uom,
        item.department,
        item.category,
        item.subCategory,
        item.venueText,
        (item.matchedVenues || []).join(" "),
        (item.templateLocationNames || []).join(" "),
        item.templateShipScopeNote,
        item.reason,
        item.excelRow,
        item.standardOrderRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term)
    );
  }, [fmlMissingRows, fmlMissingSearch]);

  const visibleFmlLowRows = useMemo(() => {
    const term = fmlLowSearch.toLowerCase().trim();
    if (!term) return fmlLowRows;

    return fmlLowRows.filter((item) =>
      [
        item.code,
        item.product,
        item.uom,
        item.department,
        item.category,
        item.subCategory,
        item.venueText,
        (item.matchedVenues || []).join(" "),
        (item.templateLocationNames || []).join(" "),
        item.templateShipScopeNote,
        item.reason,
        item.excelRow,
        item.standardOrderRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term)
    );
  }, [fmlLowRows, fmlLowSearch]);

  const filterCounts = useMemo(() => {
    const rows = activeRows;

    return {
      all: rows.length,
      needsOrder: rows.filter((item) => Number(item.suggestedOrder || 0) > 0)
        .length,
      parCap: rows.filter((item) => item.parCapApplied).length,
      blue: rows.filter((item) => item.alertType === "blue").length,
      red: rows.filter((item) => item.alertType === "red").length,
      noConsumption: rows.filter((item) => Number(item.pastConsumption || 0) <= 0)
        .length,
      noStock: rows.filter((item) => Number(item.stockOnHand || 0) <= 0).length,
    };
  }, [activeRows]);

  const loadDefaultTemplate = async () => {
    try {
      setTemplateStatus("Loading default ERP template...");

      const response = await fetch("/template.xlsx", {
        cache: "no-store",
      });

      if (!response.ok) {
        setTemplateMap({});
        setTemplateStatus("Default ERP template file was not found.");
        return;
      }

      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsedTemplate = parseTemplateWorkbook(workbook);

      setTemplateMap(parsedTemplate);
      setTemplateFileName("Default ERP Food ordering template");
      setTemplateStatus(
        "Default ERP template loaded. " +
          Object.keys(parsedTemplate).length +
          " venue sheet(s) prepared."
      );
    } catch (error) {
      setTemplateMap({});
      setTemplateStatus(
        error?.message || "Could not load default ERP template."
      );
    }
  };

  useEffect(() => {
    loadDefaultTemplate();
  }, []);

  const uploadTemplateFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setTemplateStatus("Loading custom ERP template...");

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsedTemplate = parseTemplateWorkbook(workbook);

      setTemplateMap(parsedTemplate);
      setTemplateFileName(file.name || "Custom ERP template");
      setTemplateStatus(
        "Custom ERP template loaded. " +
          Object.keys(parsedTemplate).length +
          " venue sheet(s) prepared."
      );

      logUsageEvent("product_template_file_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        sheetCount: workbook.SheetNames.length,
        venueCount: Object.keys(parsedTemplate).length,
      });
    } catch (error) {
      const text = error?.message || "Could not load custom template file.";
      setTemplateStatus(text);
      window.alert(text);
    } finally {
      event.target.value = "";
    }
  };

  const uploadNextOrderFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setNextOrderLoading(true);
    setNextOrderMessage("Reading order file...");

    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsed = parseNextOrderWorkbook({
        workbook,
        templateMap,
        fallbackShipCode: userShip,
      });

      setNextOrderFileName(file.name);
      setNextOrderSourceRows(parsed.rows);
      setNextOrderRows([]);
      setNextOrderMeta(parsed.meta);
      setFmlMissingRows(parsed.fmlReportRows || []);
      setFmlLowRows(parsed.fmlRunningLowRows || []);

      setNextOrderSearch("");
      setFmlMissingSearch("");
      setFmlLowSearch("");
      setNextOrderFilter("all");
      setNextOrderView("order");

      setNextOrderMessage(
        "Order file loaded. " +
          parsed.meta.totalItems +
          " product row(s) found. Ship detected: " +
          (parsed.meta.shipDisplayName || parsed.meta.shipCode || "N/A") +
          ". " +
          parsed.meta.itemsNeedingOrder +
          " need order, " +
          parsed.meta.parCapItems +
          " par cap, " +
          parsed.meta.blueReviewItems +
          " blue review, " +
          parsed.meta.redReviewItems +
          " red review, " +
          parsed.meta.fmlMissingItems +
          " FML not ordered/not used, " +
          parsed.meta.fmlRunningLowItems +
          " FML running low."
      );

      logUsageEvent("next_order_file_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        sheetName: parsed.meta.sheetName,
        shipName: parsed.meta.shipName,
        shipCode: parsed.meta.shipCode,
        totalItems: parsed.meta.totalItems,
        itemsNeedingOrder: parsed.meta.itemsNeedingOrder,
        parCapItems: parsed.meta.parCapItems,
        blueReviewItems: parsed.meta.blueReviewItems,
        redReviewItems: parsed.meta.redReviewItems,
        fmlMissingItems: parsed.meta.fmlMissingItems,
        fmlRunningLowItems: parsed.meta.fmlRunningLowItems,
      });
    } catch (error) {
      const text = error?.message || "Could not read the order file.";

      setNextOrderFileName(file.name);
      setNextOrderSourceRows([]);
      setNextOrderRows([]);
      setNextOrderMeta({});
      setFmlMissingRows([]);
      setFmlLowRows([]);
      setNextOrderMessage(text);
      window.alert(text);
    } finally {
      setNextOrderLoading(false);
      event.target.value = "";
    }
  };

  const generateNextOrderReport = () => {
    setNextOrderLoading(true);
    setNextOrderMessage("Generating next order...");

    window.setTimeout(() => {
      try {
        if (!nextOrderSourceRows.length) {
          setNextOrderRows([]);
          setNextOrderMessage("Upload the latest order file first.");
          return;
        }

        const rows = sortNextOrderRows(nextOrderSourceRows);

        setNextOrderRows(
          rows.map((item, index) => ({
            ...item,
            orderRank: index + 1,
          }))
        );

        setNextOrderMessage(
          "Generated " +
            rows.length +
            " product line(s). Ship: " +
            (nextOrderMeta.shipDisplayName ||
              nextOrderMeta.shipCode ||
              userShip ||
              "N/A") +
            ". Use filters and search to review products."
        );

        logUsageEvent("next_order_generated", {
          module: "generate_next_order",
          fileName: nextOrderFileName,
          ship: nextOrderMeta.shipCode || userShip,
          rowsGenerated: rows.length,
          itemsNeedingOrder: nextOrderMeta.itemsNeedingOrder || 0,
          parCapItems: nextOrderMeta.parCapItems || 0,
          blueReviewItems: nextOrderMeta.blueReviewItems || 0,
          redReviewItems: nextOrderMeta.redReviewItems || 0,
        });
      } catch (error) {
        setNextOrderRows([]);
        setNextOrderMessage(error?.message || "Could not generate next order.");
      } finally {
        setNextOrderLoading(false);
      }
    }, 25);
  };

  const uploadYearlyRegionalConsumptionFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setYearlyRegionalMessage("Loading yearly regional consumption file...");

      const arrayBuffer = await file.arrayBuffer();

      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const parsed = parseYearlyRegionalConsumptionWorkbook(workbook);

      setYearlyRegionalConsumption(parsed);
      setYearlyRegionalFileName(file.name);
      setSelectedRegionalConsumptionRegion("");

      setYearlyRegionalMessage(
        "Yearly regional file loaded. " +
          (parsed.aggregates?.length || 0) +
          " regional product record(s) found."
      );

      logUsageEvent("yearly_regional_consumption_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        sourceSheet: parsed.sourceSheet,
        regions: parsed.regionOptions || [],
        aggregates: parsed.aggregates?.length || 0,
      });
    } catch (error) {
      setYearlyRegionalConsumption(null);
      setYearlyRegionalFileName("");
      setSelectedRegionalConsumptionRegion("");

      const text =
        error?.message || "Could not load yearly regional consumption file.";

      setYearlyRegionalMessage(text);
      window.alert(text);
    } finally {
      event.target.value = "";
    }
  };

  const getRowsForOrderExport = () => visibleOrderRows;

  const exportNextOrderToExcel = () => {
    const rows = getRowsForOrderExport();

    if (!rows.length) {
      window.alert("No next-order lines found.");
      return;
    }

    const exportRows = rows.map((item, index) => ({
      Line: index + 1,
      ExcelRow: item.excelRow,
      ShipCode: nextOrderMeta.shipCode || "",
      Ship: nextOrderMeta.shipDisplayName || nextOrderMeta.shipName || "",
      Code: item.code || "",
      Product: item.product,
      UM: item.uom,
      StockOnHand: Number(item.stockOnHand || 0),
      ParLevel_Q_14Days: Number(item.parLevel || 0),
      FutureOrders_F_to_N: Number(item.futureOrders || 0),
      PastConsumption_AI_to_AN: Number(item.pastConsumption || 0),
      HistoricalSailorDays_AI5_AI6: Number(item.historicalSailorDays || 0),
      AverageConsumptionPerSailorDay: Number(
        item.averageConsumptionPerSailorDay || 0
      ),
      AverageConsumptionPerDay: Number(item.averageConsumptionPerDay || 0),
      DaysUntilArrival_B2_to_B3: Number(item.daysUntilArrival || 0),
      ConsumptionUntilArrival: Number(item.consumptionUntilArrival || 0),
      AvailableAtArrival: Number(item.availableAtArrival || 0),
      ProjectedVoyageNeed_B6: Number(item.projectedNeed || 0),
      RawSuggestedBeforePar: Number(item.rawSuggestedOrder || 0),
      ParMaxAllowed_Q_plus_10_percent: Number(item.parMaxAllowed || 0),
      ParCapApplied: item.parCapApplied ? "Yes" : "No",
      SuggestedNextOrder: Number(item.suggestedOrder || 0),
      Alert: item.alertLabel || "",
      AlertType: item.alertType || "",
      ParNote: item.parLevelNote || "",
      Reason:
        item.orderReason ||
        "Average daily consumption x voyage days, adjusted for stock/future orders until order arrival.",
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Next Order");

    XLSX.writeFile(
      wb,
      `next-order-${sanitizeFileName(
        nextOrderMeta.shipDisplayName ||
          nextOrderMeta.shipName ||
          nextOrderMeta.shipCode ||
          userShip ||
          "ship"
      )}.xlsx`
    );

    logUsageEvent("export_excel_clicked", {
      module: "generate_next_order",
      reportMode: "next_order",
      ship: nextOrderMeta.shipCode || userShip,
      rows: rows.length,
      search: nextOrderSearch,
      filter: nextOrderFilter,
    });
  };

  const exportFmlMissingReportToExcel = () => {
    const rows = visibleFmlMissingRows;

    if (!rows.length) {
      window.alert("No FML not ordered/not used rows found.");
      return;
    }

    const exportRows = rows.map((item, index) => ({
      Line: index + 1,
      FMLRow: item.excelRow,
      StandardOrderRow: item.standardOrderRow || "Not found",
      ShipCode: item.currentShipCode || nextOrderMeta.shipCode || "",
      Code: item.code || "",
      Product: item.product || "",
      UM: item.uom || "",
      Department: item.department || "",
      Category: item.category || "",
      SubCategory: item.subCategory || "",
      MatchedTemplateVenues: (item.matchedVenues || []).join(", "),
      TemplateLocations: (item.templateLocationNames || []).join(", "),
      TemplateSheets: (item.templateSheetNames || []).join(", "),
      TemplateShipScope: item.templateShipScopeNote || "Used by all ships",
      VenuesFromFMLColumnF: item.venueText || "",
      StockOnHand: Number(item.stockOnHand || 0),
      FutureOrders_F_to_N: Number(item.futureOrders || 0),
      PastConsumption_AI_to_AN: Number(item.pastConsumption || 0),
      FoundInStandardOrderTemplate: item.foundInOrderTemplate ? "Yes" : "No",
      FoundInTemplateForShip: item.foundInTemplate ? "Yes" : "No",
      Reason: item.reason || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "FML Not Ordered Not Used");

    XLSX.writeFile(
      wb,
      `fml-not-ordered-not-used-${sanitizeFileName(
        nextOrderMeta.shipDisplayName ||
          nextOrderMeta.shipCode ||
          userShip ||
          "ship"
      )}.xlsx`
    );
  };

  const exportFmlLowReportToExcel = () => {
    const rows = visibleFmlLowRows;

    if (!rows.length) {
      window.alert("No FML running-low rows found.");
      return;
    }

    const exportRows = rows.map((item, index) => ({
      Line: index + 1,
      FMLRow: item.excelRow,
      StandardOrderRow: item.standardOrderRow || "Not found",
      ShipCode: item.currentShipCode || nextOrderMeta.shipCode || "",
      Code: item.code || "",
      Product: item.product || "",
      UM: item.uom || "",
      Department: item.department || "",
      Category: item.category || "",
      SubCategory: item.subCategory || "",
      MatchedTemplateVenues: (item.matchedVenues || []).join(", "),
      TemplateLocations: (item.templateLocationNames || []).join(", "),
      TemplateSheets: (item.templateSheetNames || []).join(", "),
      TemplateShipScope: item.templateShipScopeNote || "Used by all ships",
      VenuesFromFMLColumnF: item.venueText || "",
      StockOnHand: Number(item.stockOnHand || 0),
      FutureOrders_F_to_N: Number(item.futureOrders || 0),
      PastConsumption_AI_to_AN: Number(item.pastConsumption || 0),
      AverageConsumptionPerDay: Number(item.averageConsumptionPerDay || 0),
      ConsumptionUntilArrival: Number(item.consumptionUntilArrival || 0),
      AvailableAtArrival: Number(item.availableAtArrival || 0),
      DaysCoverAtArrival: Number(item.daysOfCoverAtArrival || 0),
      SuggestedOrder: Number(item.suggestedOrder || 0),
      Reason: item.reason || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "FML Running Low");

    XLSX.writeFile(
      wb,
      `fml-running-low-${sanitizeFileName(
        nextOrderMeta.shipDisplayName ||
          nextOrderMeta.shipCode ||
          userShip ||
          "ship"
      )}.xlsx`
    );
  };

  const printRowsHtml = ({ title, rows, type }) => {
    const shipLabel =
      nextOrderMeta.shipDisplayName ||
      nextOrderMeta.shipName ||
      nextOrderMeta.shipCode ||
      userShip ||
      "N/A";

    const tableRows =
      type === "order"
        ? rows
            .map(
              (item, index) => `
                <tr>
                  <td>${index + 1}</td>
                  <td>${escapeHtml(item.excelRow || "")}</td>
                  <td>${escapeHtml(item.code || "")}</td>
                  <td>${escapeHtml(item.product || "")}</td>
                  <td>${escapeHtml(item.uom || "")}</td>
                  <td>${formatQty(item.stockOnHand)}</td>
                  <td>${formatQty(item.futureOrders)}</td>
                  <td>${formatQty(item.pastConsumption)}</td>
                  <td>${formatQty(item.averageConsumptionPerDay)}</td>
                  <td>${formatQty(item.availableAtArrival)}</td>
                  <td class="qty">${formatQty(item.suggestedOrder)}${
                item.parCapApplied ? " (Par cap)" : ""
              }</td>
                  <td class="${
                    item.alertType === "red"
                      ? "red"
                      : item.alertType === "blue" || item.alertType === "order"
                      ? "blue"
                      : ""
                  }">${escapeHtml(item.alertLabel || "")}</td>
                </tr>
              `
            )
            .join("")
        : rows
            .map(
              (item, index) => `
                <tr>
                  <td>${index + 1}</td>
                  <td>${escapeHtml(item.excelRow || "")}</td>
                  <td>${escapeHtml(item.code || "")}</td>
                  <td>${escapeHtml(item.product || "")}</td>
                  <td>${escapeHtml(item.uom || "")}</td>
                  <td class="blue">${escapeHtml(
                    (item.matchedVenues || []).join(", ") ||
                      item.venueText ||
                      ""
                  )}</td>
                  <td>${escapeHtml(
                    item.templateShipScopeNote || "Used by all ships"
                  )}</td>
                  <td>${formatQty(item.stockOnHand)}</td>
                  <td>${formatQty(item.futureOrders)}</td>
                  <td>${formatQty(item.pastConsumption)}</td>
                  <td class="warn">${escapeHtml(item.reason || "")}</td>
                </tr>
              `
            )
            .join("");

    const headers =
      type === "order"
        ? `
          <tr>
            <th>#</th>
            <th>Excel Row</th>
            <th>Code</th>
            <th>Product</th>
            <th>UM</th>
            <th>Stock</th>
            <th>Future Orders</th>
            <th>Past Consumption</th>
            <th>Avg / Day</th>
            <th>At Arrival</th>
            <th>Suggested</th>
            <th>Alert</th>
          </tr>
        `
        : `
          <tr>
            <th>#</th>
            <th>FML Row</th>
            <th>Code</th>
            <th>Product</th>
            <th>UM</th>
            <th>Matched Venue(s)</th>
            <th>Template Scope</th>
            <th>Stock</th>
            <th>Future Orders</th>
            <th>Past Consumption</th>
            <th>Reason</th>
          </tr>
        `;

    const html = `
      <html>
        <head>
          <title>${escapeHtml(title)}</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            .meta { margin: 2px 0; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .qty { color: #0057b8; font-weight: bold; }
            .blue { color: #0057b8; font-weight: bold; }
            .red { color: #b00020; font-weight: bold; }
            .warn { color: #8a5a00; font-weight: bold; }
          </style>
        </head>

        <body>
          <h1>${escapeHtml(title)}</h1>
          <div class="meta"><strong>Source file:</strong> ${escapeHtml(
            nextOrderFileName || "N/A"
          )}</div>
          <div class="meta"><strong>Ship:</strong> ${escapeHtml(shipLabel)}</div>
          <div class="meta"><strong>Ship code:</strong> ${escapeHtml(
            nextOrderMeta.shipCode || "N/A"
          )}</div>
          <div class="meta"><strong>Rows:</strong> ${rows.length}</div>
          <div class="meta"><strong>Generated:</strong> ${escapeHtml(
            new Date().toLocaleString()
          )}</div>

          <table>
            <thead>${headers}</thead>
            <tbody>${tableRows}</tbody>
          </table>
        </body>
      </html>
    `;

    const printWindow = window.open("", "_blank");

    if (!printWindow) {
      window.alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  };

  const printNextOrder = () => {
    if (!visibleOrderRows.length) {
      window.alert("No next-order lines found.");
      return;
    }

    printRowsHtml({
      title: "Generated Next Order",
      rows: visibleOrderRows,
      type: "order",
    });
  };

  const printFmlMissingReport = () => {
    if (!visibleFmlMissingRows.length) {
      window.alert("No FML not ordered/not used rows found.");
      return;
    }

    printRowsHtml({
      title: "FML Products Not Ordered / Not Used",
      rows: visibleFmlMissingRows,
      type: "fml",
    });
  };

  const printFmlLowReport = () => {
    if (!visibleFmlLowRows.length) {
      window.alert("No FML running-low rows found.");
      return;
    }

    printRowsHtml({
      title: "FML Products Running Low By Arrival",
      rows: visibleFmlLowRows,
      type: "fml",
    });
  };

  const orderShipLabel =
    nextOrderMeta.shipDisplayName ||
    nextOrderMeta.shipName ||
    nextOrderMeta.shipCode ||
    userShip ||
    "N/A";

  const activeRegionalRegionLabel = !selectedRegionalConsumptionRegion
    ? "Not selected"
    : selectedRegionalConsumptionRegion === YEARLY_REGION_ALL
    ? "All regions"
    : selectedRegionalConsumptionRegion;

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
            🛒 Generate Next Order {userShip ? `• ${getShipDisplayName(userShip)}` : ""}
          </div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Order File</h2>

          <label style={styles.label}>Latest order workbook</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadNextOrderFile}
            style={styles.fileInput}
            disabled={nextOrderLoading}
          />

          <button
            type="button"
            style={styles.primaryButton}
            onClick={generateNextOrderReport}
            disabled={nextOrderLoading || !nextOrderSourceRows.length}
          >
            {nextOrderLoading ? "Generating..." : "🛒 Generate Next Order"}
          </button>

          {nextOrderMessage && <p style={styles.message}>{nextOrderMessage}</p>}

          <div style={styles.infoBox}>
            <div>
              📄 Order file: <strong>{nextOrderFileName || "Not uploaded"}</strong>
            </div>

            <div>
              🚢 Detected ship: <strong>{orderShipLabel}</strong>
            </div>

            <div>
              🔑 Ship code used by app:{" "}
              <strong>{nextOrderMeta.shipCode || "N/A"}</strong>
            </div>

            {nextOrderMeta.scarletAliasNote && (
              <div style={{ color: "#0057b8", fontWeight: "bold" }}>
                ✅ {nextOrderMeta.scarletAliasNote}
              </div>
            )}

            <div>
              📅 Order date B2:{" "}
              <strong>{nextOrderMeta.orderDate || "N/A"}</strong>
            </div>

            <div>
              📅 Arrival date B3:{" "}
              <strong>{nextOrderMeta.arrivalDate || "N/A"}</strong>
            </div>

            <div>
              👥 Sailors B5:{" "}
              <strong>{formatQty(nextOrderMeta.targetSailors)}</strong>
            </div>

            <div>
              📆 Days B6:{" "}
              <strong>{formatQty(nextOrderMeta.targetDays)}</strong>
            </div>

            <div>
              ⏳ Days until arrival:{" "}
              <strong>{formatQty(nextOrderMeta.daysUntilArrival)}</strong>
            </div>

            <div>
              📋 Product rows:{" "}
              <strong>{nextOrderMeta.totalItems || activeRows.length}</strong>
            </div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📋 ERP Template</h2>

          <div style={styles.infoBox}>
            <div>
              📄 Template: <strong>{templateFileName || "N/A"}</strong>
            </div>

            <div>{templateStatus}</div>

            <div>
              ✅ Scarlet aliases supported:{" "}
              <strong>SC, SCL, V1, V 1, V-1, Scarlet</strong>
            </div>
          </div>

          <label style={styles.label}>Optional: replace template file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadTemplateFile}
            style={styles.fileInput}
          />

          <button
            type="button"
            style={styles.backButton}
            onClick={loadDefaultTemplate}
          >
            🔄 Reload Default Template
          </button>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🌎 Regional Par Optional</h2>

          <label style={styles.label}>
            Yearly regional consumption file May 2025 - April 2026
          </label>

          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadYearlyRegionalConsumptionFile}
            style={styles.fileInput}
          />

          <label style={styles.label}>Region / home port</label>
          <select
            value={selectedRegionalConsumptionRegion}
            onChange={(event) =>
              setSelectedRegionalConsumptionRegion(event.target.value)
            }
            style={styles.searchInput}
          >
            <option value="">Select region / origin</option>
            <option value={YEARLY_REGION_ALL}>All regions</option>

            {(yearlyRegionalConsumption?.regionOptions || []).map((region) => (
              <option key={region} value={region}>
                {region}
              </option>
            ))}
          </select>

          <label style={styles.label}>Regional par buffer %</label>
          <input
            type="number"
            min="0"
            step="1"
            value={regionalParBufferPercent}
            onChange={(event) =>
              setRegionalParBufferPercent(Number(event.target.value || 0))
            }
            style={styles.searchInput}
          />

          <div style={styles.infoBox}>
            <div>
              📄 File: <strong>{yearlyRegionalFileName || "Not loaded"}</strong>
            </div>

            <div>
              🧭 Region: <strong>{activeRegionalRegionLabel}</strong>
            </div>

            <div>
              🧮 Buffer:{" "}
              <strong>{formatRegionalQty(regionalParBufferPercent)}%</strong>
            </div>

            {yearlyRegionalMessage && <div>{yearlyRegionalMessage}</div>}
          </div>
        </div>
      </section>

      <section style={styles.card}>
        <div
          style={sx(styles.header, {
            boxShadow: "none",
            padding: 0,
            marginBottom: 16,
          })}
        >
          <div>
            <h2 style={styles.productTitle}>🧭 Report View</h2>
            <p style={sx(styles.emptyText, { margin: 0 })}>
              Review suggested order, FML not ordered/not used, and FML running
              low.
            </p>
          </div>

          <div style={styles.headerActions}>
            <button
              type="button"
              style={sx(
                styles.viewModeButton,
                nextOrderView === "order" ? styles.viewModeButtonActive : {}
              )}
              onClick={() => setNextOrderView("order")}
            >
              🛒 Next Order ({activeRows.length})
            </button>

            <button
              type="button"
              style={sx(
                styles.viewModeButton,
                nextOrderView === "fmlMissing" ? styles.viewModeButtonActive : {}
              )}
              onClick={() => setNextOrderView("fmlMissing")}
            >
              🔵 FML Not Ordered ({fmlMissingRows.length})
            </button>

            <button
              type="button"
              style={sx(
                styles.viewModeButton,
                nextOrderView === "fmlLow" ? styles.viewModeButtonActive : {}
              )}
              onClick={() => setNextOrderView("fmlLow")}
            >
              ⚠️ FML Running Low ({fmlLowRows.length})
            </button>
          </div>
        </div>

        {nextOrderView === "order" && (
          <>
            <div style={styles.infoBox}>
              <div>
                🚢 Ship:{" "}
                <strong>
                  {orderShipLabel} / {nextOrderMeta.shipCode || "N/A"}
                </strong>
              </div>

              <div>
                📦 Showing: <strong>{visibleOrderRows.length}</strong> /{" "}
                {activeRows.length}
              </div>

              <div>
                🛒 Need order: <strong>{filterCounts.needsOrder}</strong>
              </div>

              <div>
                📏 Par cap: <strong>{filterCounts.parCap}</strong>
              </div>

              <div>
                🔵 Blue review: <strong>{filterCounts.blue}</strong>
              </div>

              <div>
                🔴 Red review: <strong>{filterCounts.red}</strong>
              </div>
            </div>

            <div style={styles.viewModeBox}>
              {[
                ["all", "📋 All", filterCounts.all],
                ["needsOrder", "🛒 Needs Order", filterCounts.needsOrder],
                ["parCap", "📏 Par Cap", filterCounts.parCap],
                ["blue", "🔵 Blue Review", filterCounts.blue],
                ["red", "🔴 Red Review", filterCounts.red],
                ["noConsumption", "0️⃣ No Consumption", filterCounts.noConsumption],
                ["noStock", "📦 No Stock", filterCounts.noStock],
              ].map(([filterKey, label, count]) => (
                <button
                  key={filterKey}
                  type="button"
                  style={sx(
                    styles.viewModeButton,
                    nextOrderFilter === filterKey ? styles.viewModeButtonActive : {}
                  )}
                  onClick={() => setNextOrderFilter(filterKey)}
                >
                  {label} ({count})
                </button>
              ))}
            </div>

            <input
              placeholder="Search code, product, U/M, alert or row..."
              value={nextOrderSearch}
              onChange={(event) => setNextOrderSearch(event.target.value)}
              style={styles.searchInput}
            />

            <div style={styles.headerActions}>
              <button
                type="button"
                style={styles.backButton}
                onClick={printNextOrder}
                disabled={!visibleOrderRows.length}
              >
                🖨️ Print
              </button>

              <button
                type="button"
                style={styles.primaryButton}
                onClick={exportNextOrderToExcel}
                disabled={!visibleOrderRows.length}
              >
                📥 Export Excel
              </button>
            </div>

            {!activeRows.length && (
              <p style={styles.emptyText}>
                Upload the latest order workbook, then generate the next order.
              </p>
            )}

            {activeRows.length > 0 && !visibleOrderRows.length && (
              <p style={styles.emptyText}>
                No products match the current filter/search.
              </p>
            )}

            <div style={styles.equipmentGrid}>
              {visibleOrderRows.map((item, index) => (
                <div
                  key={`${item.excelRow}-${item.code}-${index}`}
                  style={sx(
                    styles.equipmentCard,
                    item.alertType === "red" ? styles.orderWarningCard : {},
                    item.alertType === "blue" || Number(item.suggestedOrder || 0) > 0
                      ? styles.orderNeededCard
                      : {}
                  )}
                >
                  <div style={styles.recipeMeta}>Excel row: {item.excelRow}</div>

                  <div style={styles.recipeName}>{item.product}</div>

                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>UM: {item.uom || "N/A"}</div>

                  <div style={styles.recipeMeta}>
                    Stock on hand: {formatQty(item.stockOnHand)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Future orders F:N: {formatQty(item.futureOrders)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Past consumption AI:AN: {formatQty(item.pastConsumption)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Avg / day: {formatQty(item.averageConsumptionPerDay)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Available at arrival: {formatQty(item.availableAtArrival)}
                  </div>

                  <div
                    style={
                      Number(item.suggestedOrder || 0) > 0
                        ? styles.suggestedOrderBad
                        : styles.suggestedOrderGood
                    }
                  >
                    Suggested next order: {formatQty(item.suggestedOrder)}
                    {item.parCapApplied ? " / Par cap" : ""}
                  </div>

                  {item.alertLabel && (
                    <div
                      style={
                        item.alertType === "red"
                          ? styles.statusBad
                          : item.alertType === "blue" || item.alertType === "order"
                          ? styles.statusWarning
                          : styles.statusNeutral
                      }
                    >
                      {item.alertLabel}
                    </div>
                  )}

                  <div style={styles.recipeMeta}>{item.parLevelNote}</div>
                </div>
              ))}
            </div>
          </>
        )}

        {nextOrderView === "fmlMissing" && (
          <>
            <div style={styles.infoBox}>
              <div>
                🔵 FML not ordered/not used rows:{" "}
                <strong>{visibleFmlMissingRows.length}</strong> /{" "}
                {fmlMissingRows.length}
              </div>

              <div>
                🚢 Ship matched as:{" "}
                <strong>{nextOrderMeta.shipCode || userShip || "N/A"}</strong>
              </div>

              <div>
                This report uses the ERP template and supports Scarlet aliases:
                SC, SCL, V1, V 1, V-1.
              </div>
            </div>

            <input
              placeholder="Search FML product, code, venue, template..."
              value={fmlMissingSearch}
              onChange={(event) => setFmlMissingSearch(event.target.value)}
              style={styles.searchInput}
            />

            <div style={styles.headerActions}>
              <button
                type="button"
                style={styles.backButton}
                onClick={printFmlMissingReport}
                disabled={!visibleFmlMissingRows.length}
              >
                🖨️ Print
              </button>

              <button
                type="button"
                style={styles.primaryButton}
                onClick={exportFmlMissingReportToExcel}
                disabled={!visibleFmlMissingRows.length}
              >
                📥 Export Excel
              </button>
            </div>

            {!fmlMissingRows.length && (
              <p style={styles.emptyText}>
                No FML not ordered/not used rows found. Upload the latest order
                file after the template is loaded.
              </p>
            )}

            <div style={styles.equipmentGrid}>
              {visibleFmlMissingRows.map((item, index) => (
                <div
                  key={`${item.excelRow}-${item.code}-${index}`}
                  style={sx(styles.equipmentCard, styles.orderNeededCard)}
                >
                  <div style={styles.recipeMeta}>FML row: {item.excelRow}</div>
                  <div style={styles.recipeName}>{item.product}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>UM: {item.uom || "N/A"}</div>
                  <div style={styles.recipeMeta}>
                    Venues: {item.venueText || "N/A"}
                  </div>

                  <div style={styles.templateFound}>
                    Matched: {(item.matchedVenues || []).join(", ") || "N/A"}
                  </div>

                  <div style={styles.recipeMeta}>
                    Template locations:{" "}
                    {(item.templateLocationNames || []).join(", ") || "N/A"}
                  </div>

                  <div style={styles.recipeMeta}>
                    Template scope: {item.templateShipScopeNote || "Used by all ships"}
                  </div>

                  <div style={styles.recipeMeta}>
                    Stock: {formatQty(item.stockOnHand)} / Future:{" "}
                    {formatQty(item.futureOrders)} / Past:{" "}
                    {formatQty(item.pastConsumption)}
                  </div>

                  <div style={styles.statusWarning}>{item.reason}</div>
                </div>
              ))}
            </div>
          </>
        )}

        {nextOrderView === "fmlLow" && (
          <>
            <div style={styles.infoBox}>
              <div>
                ⚠️ FML running-low rows:{" "}
                <strong>{visibleFmlLowRows.length}</strong> / {fmlLowRows.length}
              </div>

              <div>
                🚢 Ship matched as:{" "}
                <strong>{nextOrderMeta.shipCode || userShip || "N/A"}</strong>
              </div>

              <div>
                This report uses the ERP template and supports Scarlet aliases:
                SC, SCL, V1, V 1, V-1.
              </div>
            </div>

            <input
              placeholder="Search FML running-low product, code, venue, template..."
              value={fmlLowSearch}
              onChange={(event) => setFmlLowSearch(event.target.value)}
              style={styles.searchInput}
            />

            <div style={styles.headerActions}>
              <button
                type="button"
                style={styles.backButton}
                onClick={printFmlLowReport}
                disabled={!visibleFmlLowRows.length}
              >
                🖨️ Print
              </button>

              <button
                type="button"
                style={styles.primaryButton}
                onClick={exportFmlLowReportToExcel}
                disabled={!visibleFmlLowRows.length}
              >
                📥 Export Excel
              </button>
            </div>

            {!fmlLowRows.length && (
              <p style={styles.emptyText}>
                No FML running-low rows found. Upload the latest order file after
                the template is loaded.
              </p>
            )}

            <div style={styles.equipmentGrid}>
              {visibleFmlLowRows.map((item, index) => (
                <div
                  key={`${item.excelRow}-${item.code}-${index}`}
                  style={sx(styles.equipmentCard, styles.zeroCountCard)}
                >
                  <div style={styles.recipeMeta}>FML row: {item.excelRow}</div>
                  <div style={styles.recipeName}>{item.product}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>UM: {item.uom || "N/A"}</div>

                  <div style={styles.recipeMeta}>
                    Venues: {item.venueText || "N/A"}
                  </div>

                  <div style={styles.templateFound}>
                    Matched: {(item.matchedVenues || []).join(", ") || "N/A"}
                  </div>

                  <div style={styles.recipeMeta}>
                    Template scope: {item.templateShipScopeNote || "Used by all ships"}
                  </div>

                  <div style={styles.recipeMeta}>
                    Stock: {formatQty(item.stockOnHand)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Past consumption: {formatQty(item.pastConsumption)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Avg / day: {formatQty(item.averageConsumptionPerDay)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Available at arrival: {formatQty(item.availableAtArrival)}
                  </div>

                  <div style={styles.recipeMeta}>
                    Days cover: {formatQty(item.daysOfCoverAtArrival)}
                  </div>

                  <div style={styles.suggestedOrderBad}>
                    Suggested order: {formatQty(item.suggestedOrder)}
                  </div>

                  <div style={styles.statusWarning}>{item.reason}</div>
                </div>
              ))}
            </div>
          </>
        )}
      </section>
    </main>
  );
}

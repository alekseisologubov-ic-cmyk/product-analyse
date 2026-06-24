"use client";

import React, { useEffect, useMemo, useState } from "react";

const DEFAULT_ERP_LOCATION_TEMPLATE_PATH = "/erp-template-locations.xlsx";
const DEFAULT_FML_SHEET_NAME = "FML March 2026";
const MAX_PREVIEW_ROWS = 900;
const ALL_SECTIONS_SCOPE = "__ALL_SECTIONS__";
const FML_SOURCE_MODE = "fmlSource";
const FML_VENUE_NOTE_SOURCE_MODE = "fmlVenueNote";
const TEMPLATE_CODE_SOURCE_MODE = "templateCodes";

const loadXlsx = async () => {
  const module = await import("xlsx");
  return module.default || module;
};

const cleanText = (value) =>
  String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const cleanKey = (value) =>
  cleanText(value)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const normalizeSheetName = (value) => cleanKey(value).replace(/[^A-Z0-9]/g, "");

const normalizeProductCode = (value) => {
  let text = String(value ?? "").replace(/\u00a0/g, " ").trim();

  if (!text) return "";

  if (/^[+-]?\d+(\.\d+)?e[+-]?\d+$/i.test(text) || /^\d+\.0+$/.test(text)) {
    const numericValue = Number(text);
    if (Number.isFinite(numericValue)) {
      text = String(Math.round(numericValue));
    }
  }

  text = text.replace(/\.0+$/g, "").trim();

  const digits = text.replace(/[^0-9]/g, "");
  return digits || cleanKey(text);
};

const normalizeVenueProductCode = (value) => {
  const rawText = String(value ?? "").replace(/\u00a0/g, " ").trim();

  if (!rawText) return "";

  const digits = rawText.replace(/[^0-9]/g, "");
  if (!digits) return "";

  const letters = rawText.replace(/[^A-Za-z]/g, "");
  const mostlyText = letters.length > 2 && digits.length < rawText.length / 2;

  // Venue tabs must use the real product code as the lookup key.
  // This prevents old ingredient names from being treated as code values.
  if (mostlyText) return "";

  return normalizeProductCode(rawText);
};

const makeSafeFilePart = (value) =>
  cleanText(value || "report")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "report";

const getDateStamp = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const hour = String(now.getHours()).padStart(2, "0");
  const minute = String(now.getMinutes()).padStart(2, "0");

  return `${year}-${month}-${day}-${hour}${minute}`;
};

const escapeHtmlValue = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

const csvEscape = (value) => {
  const text = String(value ?? "");
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
};

const downloadBlob = ({ content, fileName, type }) => {
  const blob = content instanceof Blob ? content : new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

const cloneSheet = (worksheet) => JSON.parse(JSON.stringify(worksheet || {}));

const getCellDisplayValue = (cell) => {
  if (!cell) return "";
  if (cell.w !== undefined && cell.w !== null) return String(cell.w);
  if (cell.v !== undefined && cell.v !== null) return String(cell.v);
  return "";
};

const getCellValueByIndex = ({ worksheet, rowIndex, colIndex, XLSX }) => {
  if (!worksheet || !XLSX) return "";
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  return getCellDisplayValue(worksheet[address]);
};

const getWorksheetRows = ({ worksheet, XLSX }) => {
  if (!worksheet || !XLSX) return [];

  return XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
  });
};

const detectFmlSheetName = (sheetNames = []) => {
  const exact = sheetNames.find(
    (sheetName) => cleanKey(sheetName) === cleanKey(DEFAULT_FML_SHEET_NAME)
  );

  if (exact) return exact;

  const fmlWithDate = sheetNames.find((sheetName) => {
    const key = cleanKey(sheetName);
    return key.includes("FML") && (key.includes("MARCH") || key.includes("2026"));
  });

  if (fmlWithDate) return fmlWithDate;

  const anyFml = sheetNames.find((sheetName) => cleanKey(sheetName).includes("FML"));
  return anyFml || "";
};

const isHelperSheetName = (sheetName, fmlSheetName) => {
  const key = cleanKey(sheetName);

  if (!key) return true;
  if (sheetName === fmlSheetName) return true;
  if (key.includes("FML")) return true;
  if (key.includes("UNIT") && key.includes("MEASURE")) return true;
  if (key === "SHEET5" || key === "SHEET") return true;

  return ["README", "SUMMARY", "INDEX", "HELP", "SETTINGS", "LISTS"].some(
    (blocked) => key === blocked || key.includes(blocked)
  );
};

const getHeaderMap = (headerRow = []) => {
  const map = {};

  headerRow.forEach((value, index) => {
    const key = cleanKey(value).replace(/[^A-Z0-9]/g, "");
    if (key && map[key] === undefined) map[key] = index;
  });

  return map;
};

const getColumnIndex = (headerMap, aliases, fallbackIndex = -1) => {
  for (const alias of aliases) {
    const key = cleanKey(alias).replace(/[^A-Z0-9]/g, "");
    if (headerMap[key] !== undefined) return headerMap[key];
  }

  return fallbackIndex;
};

const escapeRegExp = (value) =>
  String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const aliasMatchesText = (textKey, alias) => {
  const aliasKey = cleanKey(alias);
  if (!textKey || !aliasKey) return false;

  if (aliasKey.length <= 3) {
    return new RegExp(`(^| )${escapeRegExp(aliasKey)}( |$)`).test(textKey);
  }

  return textKey === aliasKey || textKey.includes(aliasKey);
};

const VENUE_ALIAS_RULES = [
  { match: ["06 GARDE MANGER"], aliases: ["06", "GARDE MANGER", "GM"] },
  { match: ["07 KITCHEN TABLE"], aliases: ["07", "KITCHEN TABLE", "THE KITCHEN TABLE", "KT"] },
  { match: ["08 SPECIAL SERVICE FOOD"], aliases: ["08", "SPECIAL SERVICE FOOD", "SPECIAL SERVICE", "SSF"] },
  { match: ["13 TEST KITCHEN"], aliases: ["13", "TEST KITCHEN", "TK"] },
  { match: ["14 PINK AGAVE"], aliases: ["14", "PINK AGAVE", "PA"] },
  { match: ["15 RAZZLE"], aliases: ["15", "RAZZLE", "RAZZLE DAZZLE", "RD"] },
  { match: ["ROJO"], aliases: ["ROJO", "RAZZLE ROJO"] },
  { match: ["ARIYA"], aliases: ["ARIYA", "RAZZLE ARIYA"] },
  { match: ["16 EXTRA VIRGIN"], aliases: ["16", "EXTRA VIRGIN", "EV"] },
  { match: ["17 WAKE"], aliases: ["17", "WAKE", "THE WAKE"] },
  { match: ["18 GUNBAE"], aliases: ["18", "GUNBAE", "GB"] },
  { match: ["19 MANNOR", "19 MANOR"], aliases: ["19", "MANNOR", "MANOR", "ANOTHER ROSE"] },
  { match: ["LOLZ"], aliases: ["LOLZ"] },
  { match: ["20 THE DOCK"], aliases: ["20", "DOCK", "THE DOCK", "UP WITH A TWIST"] },
  { match: ["21 THE PIZZA PLACE"], aliases: ["21", "PIZZA", "PIZZA PLACE", "THE PIZZA PLACE"] },
  { match: ["22 LTS POPSTAR", "22 ITS POPSTAR"], aliases: ["22", "POPSTAR", "LTS POPSTAR", "ITS POPSTAR"] },
  { match: ["23 SOCIAL CLUB"], aliases: ["23", "SOCIAL CLUB"] },
  { match: ["24 SUN CLUB"], aliases: ["24", "SUN CLUB"] },
  { match: ["25 SIP LOUNGE"], aliases: ["25", "SIP", "SIP LOUNGE"] },
  { match: ["26 SHIP EATS"], aliases: ["26", "SHIP EATS"] },
  { match: ["27 QUICKEZE"], aliases: ["27", "QUICKEZE", "QUICKIEZE"] },
  { match: ["28 DINE AND DASH"], aliases: ["28", "DINE AND DASH", "EAT AND DRINK"] },
  { match: ["29 BURGER BAR"], aliases: ["29", "BURGER", "BURGER BAR"] },
  { match: ["30 HOT OF THE PRESS"], aliases: ["30", "HOT OF THE PRESS"] },
  { match: ["32 TACO"], aliases: ["32", "TACO"] },
  { match: ["33 BENTO BABY"], aliases: ["33", "BENTO BABY", "DIM SUM"] },
  { match: ["34 DAILY MIX"], aliases: ["34", "DAILY MIX"] },
  { match: ["35 NOODLE AROUND"], aliases: ["35", "NOODLE", "NOODLE AROUND"] },
  { match: ["36 BIMINI"], aliases: ["36", "BIMINI", "BIMNI", "BIMINI BEACH CLUB", "BBC"] },
  { match: ["37 THE GALLEY KITCHEN"], aliases: ["37", "GALLEY", "GALLEY KITCHEN", "THE GALLEY KITCHEN"] },
  { match: ["39 CHARGABLE", "39 CHARGEABLE"], aliases: ["39", "CHARGABLE", "CHARGEABLE", "CHARGEABLE ITEM"] },
  { match: ["44 PASTRY", "44 BAKERY"], aliases: ["44", "PASTRY", "BAKERY", "PASTRY BAKERY"] },
  { match: ["LIQUER", "LIQUOR"], aliases: ["LIQUER", "LIQUOR", "BAR"] },
];

const getVenueAliasesForSheetName = (sheetName) => {
  const sheetKey = cleanKey(sheetName);
  const aliases = new Set([sheetName, sheetKey.replace(/^\d+\s+/, "")]);

  VENUE_ALIAS_RULES.forEach((rule) => {
    if (rule.match.some((value) => sheetKey.includes(cleanKey(value)))) {
      rule.aliases.forEach((alias) => aliases.add(alias));
    }
  });

  return Array.from(aliases)
    .map(cleanText)
    .filter(Boolean)
    .filter((alias) => !/^\d+$/.test(alias));
};

const getFmlVenueText = (item = {}) =>
  [item.comments, item.adjustmentComments, item.notes]
    .map(cleanText)
    .filter(Boolean)
    .join(" / ");

const fmlItemMatchesVenueSheet = (fmlItem, sheetName) => {
  const textKey = cleanKey(getFmlVenueText(fmlItem));
  if (!textKey) return false;

  const sheetKey = cleanKey(sheetName);
  const aliases = getVenueAliasesForSheetName(sheetName);

  if (textKey.includes("ROJO") && !sheetKey.includes("ROJO")) return false;
  if (textKey.includes("ARIYA") && !sheetKey.includes("ARIYA")) return false;
  if (textKey.includes("BRL") && !sheetKey.includes("BRL") && sheetKey.includes("DOCK")) return false;

  return aliases.some((alias) => aliasMatchesText(textKey, alias));
};

const buildFmlIndex = ({ workbook, fmlSheetName, XLSX }) => {
  const worksheet = workbook?.Sheets?.[fmlSheetName];

  if (!worksheet || !XLSX) {
    return {
      rows: [],
      byCode: new Map(),
      headerRowIndex: -1,
      codeColumnIndex: -1,
      sourceSheetName: fmlSheetName || "",
    };
  }

  const rows = getWorksheetRows({ worksheet, XLSX });
  let headerRowIndex = -1;

  rows.slice(0, 30).some((row, index) => {
    const rowKey = cleanKey(row.join(" "));
    const hasCode = row.some((cell) => ["PRODUCT", "CODE", "ITEM CODE"].includes(cleanKey(cell)));
    const hasName = row.some((cell) => cleanKey(cell).includes("PRODUCT NAME") || cleanKey(cell).includes("DESCRIPTION"));

    if (hasCode && hasName && rowKey.includes("UM")) {
      headerRowIndex = index;
      return true;
    }

    return false;
  });

  if (headerRowIndex < 0) {
    headerRowIndex = 2;
  }

  const headerMap = getHeaderMap(rows[headerRowIndex] || []);
  const departmentIndex = getColumnIndex(headerMap, ["Department", "Departement"], 1);
  const categoryIndex = getColumnIndex(headerMap, ["Category"], 2);
  const subCategoryIndex = getColumnIndex(headerMap, ["SubCategory", "Sub Category"], 3);
  const codeIndex = getColumnIndex(headerMap, ["Product", "Code", "Item Code"], 4);
  const productNameIndex = getColumnIndex(headerMap, ["Product Name", "Final Description", "Description", "Ingredient Name"], 5);
  const typeIndex = getColumnIndex(headerMap, ["Type"], 6);
  const brandIndex = getColumnIndex(headerMap, ["Brand"], 7);
  const umIndex = getColumnIndex(headerMap, ["UM Ship", "UOM Ship", "UM", "UOM", "Unit"], 8);
  const allergensIndex = getColumnIndex(headerMap, ["Has Allergens", "Allergens"], 9);
  const picturesIndex = getColumnIndex(headerMap, ["Pictures", "Picture"], 10);
  const nutritionIndex = getColumnIndex(headerMap, ["Has Nutrition Facts", "Nutrition Facts"], 11);
  const priceUomIndex = getColumnIndex(headerMap, ["Price-U/M", "Price UM", "Price U/M"], 12);
  const priceIndex = getColumnIndex(headerMap, ["Price"], 13);
  const crewStaffIndex = getColumnIndex(headerMap, ["Crew/Staff", "Crew Staff"], 14);
  const sclIndex = getColumnIndex(headerMap, ["SCL", "SC", "Scarlet"], 15);
  const valIndex = getColumnIndex(headerMap, ["VAL", "VL", "Valiant"], 16);
  const resIndex = getColumnIndex(headerMap, ["RES", "RL", "Resilient"], 17);
  const brlIndex = getColumnIndex(headerMap, ["BRL", "Brilliant"], 18);
  const commentsIndex = getColumnIndex(headerMap, ["Comments"], 19);
  const adjustmentIndex = getColumnIndex(headerMap, ["FML Adjustment Comments", "Adjustment Comments"], 20);
  const notesIndex = getColumnIndex(headerMap, ["Notes"], 21);

  const byCode = new Map();
  const parsedRows = [];

  rows.slice(headerRowIndex + 1).forEach((row, offset) => {
    const excelRow = headerRowIndex + 2 + offset;
    const code = normalizeProductCode(row[codeIndex]);
    const productName = cleanText(row[productNameIndex]);

    if (!code || !productName) return;
    if (cleanKey(productName) === "PRODUCT NAME") return;

    const item = {
      code,
      productName,
      department: cleanText(row[departmentIndex]),
      category: cleanText(row[categoryIndex]),
      subCategory: cleanText(row[subCategoryIndex]),
      type: cleanText(row[typeIndex]),
      brand: cleanText(row[brandIndex]),
      um: cleanText(row[umIndex]),
      allergens: cleanText(row[allergensIndex]),
      pictures: cleanText(row[picturesIndex]),
      nutritionFacts: cleanText(row[nutritionIndex]),
      priceUom: cleanText(row[priceUomIndex]),
      price: cleanText(row[priceIndex]),
      crewStaff: cleanText(row[crewStaffIndex]),
      scl: cleanText(row[sclIndex]),
      val: cleanText(row[valIndex]),
      res: cleanText(row[resIndex]),
      brl: cleanText(row[brlIndex]),
      comments: cleanText(row[commentsIndex]),
      adjustmentComments: cleanText(row[adjustmentIndex]),
      notes: cleanText(row[notesIndex]),
      venueText: cleanText([row[commentsIndex], row[adjustmentIndex], row[notesIndex]].filter(Boolean).join(" / ")),
      fmlRow: excelRow,
      fmlAddress: `${fmlSheetName}!E${excelRow}:I${excelRow}`,
    };

    parsedRows.push(item);

    if (!byCode.has(code)) {
      byCode.set(code, item);
    }
  });

  return {
    rows: parsedRows,
    byCode,
    headerRowIndex,
    codeColumnIndex: codeIndex,
    sourceSheetName: fmlSheetName,
  };
};

const getSectionTypeFromTitle = (sectionTitle) => {
  const key = cleanKey(sectionTitle);

  if (!key) return "templateOnly";
  if (key.includes("LIQUOR") || key.includes("LIQUER") || key.includes("ALCOHOL") || key.includes("WINE")) return "liquor";
  if (key.includes("CHARG")) return "chargeable";
  if (key.includes("FISH") || key.includes("SEAFOOD")) return "fish";
  if (key.includes("MEAT")) return "meat";
  if (key.includes("VEGET") || key.includes("VEGETE") || key.includes("FRUIT")) return "produce";
  if (key.includes("PASTRY")) return "pastry";
  if (key.includes("BAKERY") || key.includes("BAKER")) return "bakery";
  if (key.includes("CANAPE") || key.includes("CIRCLE")) return "menuBasic";
  if (key.includes("CREW PARTY") || key.includes("SPECIAL REQUEST")) return "menuBasic";
  if (key.includes("MENU BASIC") || key.includes("MENU BADIC") || key.includes("PREPERATION") || key.includes("PREPARATION")) return "menuBasic";

  return "templateOnly";
};

const getFmlItemSearchKey = (item) =>
  cleanKey([
    item?.department,
    item?.category,
    item?.subCategory,
    item?.type,
    item?.productName,
    item?.comments,
    item?.adjustmentComments,
    item?.notes,
    item?.venueText,
  ].join(" "));

const isFmlMeatItem = (item) => {
  const key = getFmlItemSearchKey(item);
  return (
    key.includes("151") ||
    key.includes("BEEF") ||
    key.includes("VEAL") ||
    key.includes("PORK") ||
    key.includes("LAMB") ||
    key.includes("POULTRY") ||
    key.includes("KOSH") ||
    key.includes("GAME") ||
    key.includes("BACON") ||
    key.includes("SAUSAGE") ||
    key.includes("TURKEY") ||
    key.includes("CHICKEN")
  );
};

const isFmlFishItem = (item) => {
  const key = getFmlItemSearchKey(item);
  return key.includes("153") || key.includes("FISH") || key.includes("SEAFOOD") || key.includes("CAVIAR") || key.includes("SHELLFISH");
};

const isFmlProduceItem = (item) => {
  const key = getFmlItemSearchKey(item);
  const category = cleanKey(item?.category);
  const subCategory = cleanKey(item?.subCategory);

  return (
    category.includes("169") ||
    category.includes("171") ||
    category.includes("FRESH FRUIT") ||
    category.includes("FRESH VEGETABLE") ||
    subCategory.includes("FRESH FRUIT") ||
    subCategory.includes("FRESH BERR") ||
    subCategory.includes("FRESH VEGETABLE") ||
    subCategory.includes("FRESH HERB") ||
    (category.includes("156") && (subCategory.includes("FRUIT") || subCategory.includes("VEGETABLE") || subCategory.includes("BERR"))) ||
    key.includes("VEGETABLES FRUITS") ||
    key.includes("FRUITS BERRIES")
  );
};

const isFmlBakeryOrPastryItem = (item) => {
  const key = getFmlItemSearchKey(item);

  return (
    key.includes("157") ||
    key.includes("168") ||
    key.includes("BAKERY") ||
    key.includes("BAKING") ||
    key.includes("FLOUR") ||
    key.includes("SUGAR") ||
    key.includes("BREAD") ||
    key.includes("COOKIE") ||
    key.includes("MIXES") ||
    key.includes("PASTRY") ||
    key.includes("YEAST")
  );
};

const isFmlLiquorItem = (item) => {
  const key = getFmlItemSearchKey(item);

  return (
    key.includes("LIQUOR") ||
    key.includes("ALCOHOL") ||
    key.includes("WINE") ||
    key.includes("BEER") ||
    key.includes("VODKA") ||
    key.includes("GIN") ||
    key.includes("RUM") ||
    key.includes("TEQUILA") ||
    key.includes("BOURBON") ||
    key.includes("BRANDY") ||
    key.includes("WHISKEY") ||
    key.includes("WHISKY") ||
    key.includes("CHAMPAGNE") ||
    key.includes("COCKTAIL")
  );
};

const isFmlChargeableItem = (item) => {
  const key = getFmlItemSearchKey(item);
  return key.includes("CHARGEABLE") || key.includes("CHARGABLE") || key.includes("CHARGEABLE ITEM");
};

const sectionMatchesFmlItem = ({ sectionTitle, item }) => {
  const sectionType = getSectionTypeFromTitle(sectionTitle);

  if (!item?.code || !item?.productName) return false;

  if (sectionType === "meat") return isFmlMeatItem(item);
  if (sectionType === "fish") return isFmlFishItem(item);
  if (sectionType === "produce") return isFmlProduceItem(item);
  if (sectionType === "liquor") return isFmlLiquorItem(item);
  if (sectionType === "chargeable") return isFmlChargeableItem(item);
  if (sectionType === "pastry" || sectionType === "bakery") return isFmlBakeryOrPastryItem(item);

  if (sectionType === "menuBasic") {
    return (
      !isFmlMeatItem(item) &&
      !isFmlFishItem(item) &&
      !isFmlProduceItem(item) &&
      !isFmlLiquorItem(item) &&
      !isFmlChargeableItem(item)
    );
  }

  return false;
};

const getFmlItemsForSection = ({ sectionTitle, fmlIndex }) => {
  const sourceRows = Array.isArray(fmlIndex?.rows) ? fmlIndex.rows : [];
  const seenCodes = new Set();

  return sourceRows
    .filter((item) => sectionMatchesFmlItem({ sectionTitle, item }))
    .filter((item) => {
      const code = normalizeProductCode(item.code);
      if (!code || seenCodes.has(code)) return false;
      seenCodes.add(code);
      return true;
    })
    .sort((left, right) => {
      const leftRow = Number(left.fmlRow || 0);
      const rightRow = Number(right.fmlRow || 0);
      if (leftRow !== rightRow) return leftRow - rightRow;
      return String(left.productName || "").localeCompare(String(right.productName || ""));
    });
};

const guessBlockTitle = ({ rows, headerRowIndex, codeCol }) => {
  const candidates = [];

  for (let rowIndex = Math.max(0, headerRowIndex - 4); rowIndex < headerRowIndex; rowIndex += 1) {
    const row = rows[rowIndex] || [];

    for (let colIndex = Math.max(0, codeCol - 2); colIndex <= codeCol + 4; colIndex += 1) {
      const value = cleanText(row[colIndex]);

      if (!value) continue;
      if (cleanKey(value) === "CODE") continue;

      candidates.push(value);
    }
  }

  if (!candidates.length) return "Section";

  return candidates
    .sort((left, right) => right.length - left.length)[0]
    .replace(/\s+/g, " ")
    .trim();
};

const findNearbyColumn = ({ row, startCol, aliases, fallbackCol }) => {
  for (let colIndex = startCol; colIndex <= startCol + 5; colIndex += 1) {
    const key = cleanKey(row[colIndex]);
    if (!key) continue;

    if (aliases.some((alias) => key.includes(cleanKey(alias)))) {
      return colIndex;
    }
  }

  return fallbackCol;
};

const isRowOneProductLocator = (value) => {
  const text = cleanText(value);
  const key = cleanKey(text);

  if (!key) return false;
  if (key === "CODE") return false;
  if (key.includes("INGREDIENT NAME")) return false;
  if (key === "UM" || key === "UOM" || key.includes("UNIT OF MEASURE")) return false;

  // In the ERP location workbook, Excel row 1 contains the product-location
  // locator such as: "16 - Extra Virgin - Vegetables & Fruit 03/16/06".
  // Important: cleanKey removes hyphens, so check the original text, not the
  // normalized key. Otherwise blocks like Extra Virgin Vegetables & Fruit are
  // missed when the Code header row is blank.
  const hasDashLocator = /\d+\s*[-–—]\s*[^-–—]+/.test(text);
  const hasDateLocator = /\d{1,2}\s*\/\s*\d{1,2}\s*\/\s*\d{1,4}/.test(text);

  return /\d/.test(key) && (hasDashLocator || hasDateLocator);
};

const getCodeCountForColumn = ({ rows, codeCol, firstDataRowIndex }) => {
  let itemCount = 0;

  for (let dataRowIndex = firstDataRowIndex; dataRowIndex < rows.length; dataRowIndex += 1) {
    const code = normalizeVenueProductCode(rows[dataRowIndex]?.[codeCol]);
    if (code) itemCount += 1;
  }

  return itemCount;
};

const findCodeColumnForRowOneLocator = ({ rows, locatorCol, headerRowIndex }) => {
  const headerRow = rows[headerRowIndex] || [];
  const firstDataRowIndex = headerRowIndex + 1;
  const candidates = [];

  for (let colIndex = Math.max(0, locatorCol - 2); colIndex <= locatorCol + 1; colIndex += 1) {
    const headerKey = cleanKey(headerRow[colIndex]);
    const itemCount = getCodeCountForColumn({ rows, codeCol: colIndex, firstDataRowIndex });

    const headerScore =
      headerKey === "CODE"
        ? 100000
        : headerKey.includes("CODE")
          ? 50000
          : 0;

    if (!itemCount && !headerScore) continue;

    candidates.push({
      colIndex,
      itemCount,
      score: headerScore + itemCount * 100 - Math.abs(locatorCol - colIndex),
    });
  }

  candidates.sort((left, right) => right.score - left.score);

  return candidates[0]?.colIndex ?? locatorCol;
};

const buildVenueBlockFromRowOneLocator = ({ rows, locatorCol, title }) => {
  const headerRowIndex = 1;
  const headerRow = rows[headerRowIndex] || [];
  const firstDataRowIndex = headerRowIndex + 1;
  const codeCol = findCodeColumnForRowOneLocator({ rows, locatorCol, headerRowIndex });
  const itemCount = getCodeCountForColumn({ rows, codeCol, firstDataRowIndex });

  const foundNameCol = findNearbyColumn({
    row: headerRow,
    startCol: codeCol + 1,
    aliases: ["Ingredient", "Product", "Name", "Description"],
    fallbackCol: -1,
  });

  const nameCol =
    foundNameCol >= 0
      ? foundNameCol
      : locatorCol > codeCol && locatorCol <= codeCol + 3
        ? locatorCol
        : codeCol + 1;

  const umCol = findNearbyColumn({
    row: headerRow,
    startCol: nameCol + 1,
    aliases: ["UM", "UOM", "Unit"],
    fallbackCol: nameCol + 1,
  });

  return {
    title: cleanText(title) || "Section",
    headerRowIndex,
    locatorCol,
    codeCol,
    nameCol,
    umCol,
    itemCount,
    locatorSource: "row1",
  };
};

const findVenueBlocks = ({ workbook, sheetName, XLSX }) => {
  const worksheet = workbook?.Sheets?.[sheetName];
  if (!worksheet || !XLSX) return [];

  const rows = getWorksheetRows({ worksheet, XLSX });
  const blocks = [];
  const seenKeys = new Set();

  const addBlock = (block) => {
    if (!block) return;
    if (!block.itemCount && block.locatorSource !== "row1") return;

    const blockKey = `${block.headerRowIndex}|${block.codeCol}|${block.nameCol}|${block.umCol}`;
    if (seenKeys.has(blockKey)) return;
    seenKeys.add(blockKey);

    blocks.push({
      ...block,
      blockKey,
      blockIndex: blocks.length,
      title:
        cleanText(block.title) ||
        guessBlockTitle({ rows, headerRowIndex: block.headerRowIndex, codeCol: block.codeCol }),
    });
  };

  // Primary locator: Excel row 1. This row contains the real location/section
  // title for each code block. Some tabs place that title above Ingredient Name
  // instead of Code, and some tabs have a blank Code header, so this is more
  // reliable than looking for the word "Code" only.
  const rowOne = rows[0] || [];
  rowOne.forEach((cell, locatorCol) => {
    if (!isRowOneProductLocator(cell)) return;

    addBlock(
      buildVenueBlockFromRowOneLocator({
        rows,
        locatorCol,
        title: cell,
      })
    );
  });

  // Fallback for any older/custom sheet where row 1 is missing but Code headers
  // are present lower in the sheet.
  rows.slice(0, 35).forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      if (cleanKey(cell) !== "CODE") return;

      const nameCol = findNearbyColumn({
        row,
        startCol: colIndex + 1,
        aliases: ["Ingredient", "Product", "Name", "Description"],
        fallbackCol: colIndex + 1,
      });

      const umCol = findNearbyColumn({
        row,
        startCol: colIndex + 1,
        aliases: ["UM", "UOM", "Unit"],
        fallbackCol: -1,
      });

      const itemCount = getCodeCountForColumn({
        rows,
        codeCol: colIndex,
        firstDataRowIndex: rowIndex + 1,
      });

      addBlock({
        title: guessBlockTitle({ rows, headerRowIndex: rowIndex, codeCol: colIndex }),
        headerRowIndex: rowIndex,
        locatorCol: colIndex,
        codeCol: colIndex,
        nameCol,
        umCol,
        itemCount,
        locatorSource: "codeHeader",
      });
    });
  });

  return blocks.sort(
    (left, right) =>
      left.headerRowIndex - right.headerRowIndex || left.codeCol - right.codeCol
  );
};

const getVisibleVenueSheetNames = (workbook, fmlSheetName, XLSX) => {
  const sheetNames = workbook?.SheetNames || [];
  const sheetInfo = workbook?.Workbook?.Sheets || [];

  return sheetNames.filter((sheetName, index) => {
    if (isHelperSheetName(sheetName, fmlSheetName)) return false;

    const hidden = Number(sheetInfo[index]?.Hidden || 0);
    if (hidden === 1 || hidden === 2) return false;
    if (!workbook.Sheets?.[sheetName]) return false;

    if (XLSX) {
      return findVenueBlocks({ workbook, sheetName, XLSX }).length > 0;
    }

    return true;
  });
};

const makeVenueItemRecord = ({
  sheetName,
  rows,
  block,
  rowIndex,
  sectionRunningCounts,
  code,
  fmlItem,
  oldName,
  oldUm,
  XLSX,
  sourceMode,
}) => {
  const productName = fmlItem?.productName || oldName;
  const um = fmlItem?.um || oldUm;
  const sectionKey = `${block.blockKey}|${block.title}`;
  const currentCount = sectionRunningCounts.get(sectionKey) || 0;
  sectionRunningCounts.set(sectionKey, currentCount + 1);

  const codeAddress = XLSX.utils.encode_cell({ r: rowIndex, c: block.codeCol });
  const nameAddress = XLSX.utils.encode_cell({ r: rowIndex, c: block.nameCol });
  const umAddress = block.umCol >= 0
    ? XLSX.utils.encode_cell({ r: rowIndex, c: block.umCol })
    : "";

  return {
    key: `${sheetName}|${block.blockKey}|${rowIndex}|${block.codeCol}|${code}`,
    venueSheet: sheetName,
    sectionTitle: block.title,
    sectionIndex: block.blockIndex,
    sectionRowNumber: currentCount + 1,
    sourceRow: rowIndex + 1,
    targetRowIndex: rowIndex,
    code,
    displayCode: fmlItem?.code || code,
    productName,
    um,
    oldName,
    oldUm,
    sourceMode,
    sourceLabel:
      sourceMode === FML_SOURCE_MODE
        ? "FML March 2026 via venue template"
        : sourceMode === FML_VENUE_NOTE_SOURCE_MODE
          ? "FML venue note"
          : sourceMode === TEMPLATE_CODE_SOURCE_MODE
            ? "Venue template code"
            : "Venue template fallback",
    fmlMatched: Boolean(fmlItem),
    wasCorrected:
      Boolean(fmlItem) &&
      (cleanKey(oldName) !== cleanKey(fmlItem.productName) || cleanKey(oldUm) !== cleanKey(fmlItem.um) || normalizeProductCode(code) !== normalizeProductCode(fmlItem.code)),
    status: fmlItem ? "FML matched" : "Missing in FML",
    fmlItem,
    department: fmlItem?.department || "",
    category: fmlItem?.category || "",
    subCategory: fmlItem?.subCategory || "",
    type: fmlItem?.type || "",
    brand: fmlItem?.brand || "",
    allergens: fmlItem?.allergens || "",
    priceUom: fmlItem?.priceUom || "",
    price: fmlItem?.price || "",
    crewStaff: fmlItem?.crewStaff || "",
    scl: fmlItem?.scl || "",
    val: fmlItem?.val || "",
    res: fmlItem?.res || "",
    brl: fmlItem?.brl || "",
    comments: fmlItem?.comments || "",
    adjustmentComments: fmlItem?.adjustmentComments || "",
    notes: fmlItem?.notes || "",
    venueText: fmlItem?.venueText || "",
    fmlRow: fmlItem?.fmlRow || "",
    fmlAddress: fmlItem?.fmlAddress || "",
    codeAddress,
    nameAddress,
    umAddress,
    block,
  };
};

const buildVenueItemsFromTemplateCodes = ({ rows, sheetName, fmlIndex, XLSX, block, sectionRunningCounts, sourceMode = TEMPLATE_CODE_SOURCE_MODE }) => {
  const items = [];

  for (let rowIndex = block.headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const code = normalizeVenueProductCode(row[block.codeCol]);
    if (!code) continue;

    const fmlItem = fmlIndex?.byCode?.get(code) || null;
    const oldName = cleanText(row[block.nameCol]);
    const oldUm = block.umCol >= 0 ? cleanText(row[block.umCol]) : "";

    items.push(
      makeVenueItemRecord({
        sheetName,
        rows,
        block,
        rowIndex,
        sectionRunningCounts,
        code,
        fmlItem,
        oldName,
        oldUm,
        XLSX,
        sourceMode,
      })
    );
  }

  return items;
};

const buildVenueItemsFromFmlSection = ({ rows, sheetName, fmlIndex, XLSX, block, sectionRunningCounts }) => {
  // The venue/location tabs are the locator: they tell us which FML codes belong
  // under each row-1 section. FML supplies the final product details. This avoids
  // putting every FML item into every venue.
  const templateItems = buildVenueItemsFromTemplateCodes({
    rows,
    sheetName,
    fmlIndex,
    XLSX,
    block,
    sectionRunningCounts,
    sourceMode: FML_SOURCE_MODE,
  });

  const usedCodes = new Set(
    templateItems
      .map((item) => normalizeProductCode(item.code))
      .filter(Boolean)
  );

  const firstDataRowIndex = block.headerRowIndex + 1;
  const fmlVenueItems = getFmlItemsForSection({
    sectionTitle: block.title,
    fmlIndex,
  }).filter((fmlItem) => {
    const code = normalizeProductCode(fmlItem.code);
    if (!code || usedCodes.has(code)) return false;
    return fmlItemMatchesVenueSheet(fmlItem, sheetName);
  });

  const extraItems = fmlVenueItems.map((fmlItem, index) => {
    const rowIndex = firstDataRowIndex + templateItems.length + index;
    usedCodes.add(normalizeProductCode(fmlItem.code));

    return makeVenueItemRecord({
      sheetName,
      rows,
      block,
      rowIndex,
      sectionRunningCounts,
      code: fmlItem.code,
      fmlItem,
      oldName: "",
      oldUm: "",
      XLSX,
      sourceMode: FML_VENUE_NOTE_SOURCE_MODE,
    });
  });

  return [...templateItems, ...extraItems];
};


const buildVenueItemRows = ({ workbook, sheetName, fmlIndex, XLSX, sourceMode = FML_SOURCE_MODE }) => {
  const worksheet = workbook?.Sheets?.[sheetName];
  if (!worksheet || !XLSX) return [];

  const rows = getWorksheetRows({ worksheet, XLSX });
  const blocks = findVenueBlocks({ workbook, sheetName, XLSX });
  const items = [];
  const sectionRunningCounts = new Map();

  blocks.forEach((block) => {
    const blockItems =
      sourceMode === FML_SOURCE_MODE
        ? buildVenueItemsFromFmlSection({
            rows,
            sheetName,
            fmlIndex,
            XLSX,
            block,
            sectionRunningCounts,
          })
        : buildVenueItemsFromTemplateCodes({
            rows,
            sheetName,
            fmlIndex,
            XLSX,
            block,
            sectionRunningCounts,
            sourceMode: TEMPLATE_CODE_SOURCE_MODE,
          });

    items.push(...blockItems);
  });

  return items;
};

const buildAllVenuePlacementRows = ({ workbook, venueSheets, fmlIndex, XLSX, sourceMode = FML_SOURCE_MODE }) => {
  if (!workbook || !XLSX) return [];

  return (venueSheets || []).flatMap((sheetName) =>
    buildVenueItemRows({
      workbook,
      sheetName,
      fmlIndex,
      XLSX,
      sourceMode,
    })
  );
};

const buildFmlPlacementAuditRows = ({ fmlRows = [], placementRows = [] }) => {
  const placementsByCode = new Map();

  (placementRows || []).forEach((placement) => {
    const code = normalizeProductCode(placement.displayCode || placement.code);
    if (!code || !placement.fmlMatched) return;

    if (!placementsByCode.has(code)) placementsByCode.set(code, []);
    placementsByCode.get(code).push(placement);
  });

  return (fmlRows || []).flatMap((fmlItem, fmlIndex) => {
    const code = normalizeProductCode(fmlItem.code);
    const placements = placementsByCode.get(code) || [];

    if (!placements.length) {
      return [
        {
          key: `unplaced-${code || fmlIndex}-${fmlItem.fmlRow || fmlIndex}`,
          placementStatus: "Not placed in venue tabs",
          ...fmlItem,
          displayCode: fmlItem.code,
          venueSheet: "",
          sectionTitle: "",
          sourceRow: "",
          templateCell: "",
          sourceLabel: "FML only",
        },
      ];
    }

    return placements.map((placement, placementIndex) => ({
      key: `placed-${code}-${placement.key || placementIndex}`,
      placementStatus: placement.sourceMode === FML_VENUE_NOTE_SOURCE_MODE
        ? "Placed from FML venue note"
        : "Placed in venue tab",
      ...fmlItem,
      displayCode: placement.displayCode || fmlItem.code,
      productName: placement.productName || fmlItem.productName,
      um: placement.um || fmlItem.um,
      venueSheet: placement.venueSheet || "",
      sectionTitle: placement.sectionTitle || "",
      sourceRow: placement.sourceRow || "",
      templateCell: placement.codeAddress || "",
      sourceLabel: placement.sourceLabel || "",
    }));
  });
};

const getFmlPlacementAuditExportRows = (rows = []) =>
  (rows || []).map((row, index) => ({
    Line: index + 1,
    PlacementStatus: row.placementStatus || "",
    FMLRow: row.fmlRow || "",
    Code: row.displayCode || row.code || "",
    ProductName: row.productName || "",
    UM: row.um || "",
    Department: row.department || "",
    Category: row.category || "",
    SubCategory: row.subCategory || "",
    Type: row.type || "",
    Brand: row.brand || "",
    VenueTab: row.venueSheet || "",
    Row1LocatorSection: row.sectionTitle || "",
    TemplateRow: row.sourceRow || "",
    TemplateCell: row.templateCell || "",
    Source: row.sourceLabel || "",
    FMLVenueNotes: row.venueText || getFmlVenueText(row),
    Comments: row.comments || "",
    AdjustmentComments: row.adjustmentComments || "",
    Notes: row.notes || "",
  }));


const buildCorrectionOverrideMap = (items) => {
  const overrideMap = new Map();

  items.forEach((item) => {
    if (item.codeAddress) {
      overrideMap.set(item.codeAddress, {
        value: item.displayCode || item.code,
        fromFml: item.fmlMatched,
        field: "Code",
        fmlAddress: item.fmlAddress,
      });
    }

    if (item.nameAddress) {
      overrideMap.set(item.nameAddress, {
        value: item.productName,
        fromFml: item.fmlMatched,
        field: "Product Name",
        fmlAddress: item.fmlAddress,
      });
    }

    if (item.umAddress) {
      overrideMap.set(item.umAddress, {
        value: item.um,
        fromFml: item.fmlMatched,
        field: "UM Ship",
        fmlAddress: item.fmlAddress,
      });
    }
  });

  return overrideMap;
};

const buildCorrectedGridRows = ({ workbook, sheetName, items, XLSX }) => {
  const worksheet = workbook?.Sheets?.[sheetName];
  if (!worksheet || !XLSX) return [];

  const rows = getWorksheetRows({ worksheet, XLSX });
  const overrideMap = buildCorrectionOverrideMap(items);
  const decodedRange = worksheet["!ref"]
    ? XLSX.utils.decode_range(worksheet["!ref"])
    : { s: { r: 0, c: 0 }, e: { r: Math.max(rows.length - 1, 0), c: 0 } };

  const maxColFromRows = rows.reduce(
    (max, row) => Math.max(max, Array.isArray(row) ? row.length - 1 : 0),
    decodedRange.e.c
  );

  const maxItemRowIndex = (items || []).reduce(
    (max, item) => Math.max(max, Number(item.targetRowIndex ?? -1)),
    -1
  );
  const maxItemColIndex = (items || []).reduce(
    (max, item) => Math.max(max, Number(item.block?.umCol ?? item.block?.nameCol ?? item.block?.codeCol ?? -1)),
    -1
  );

  const startRow = decodedRange.s.r || 0;
  const endRow = Math.max(decodedRange.e.r || 0, rows.length - 1, maxItemRowIndex);
  const startCol = decodedRange.s.c || 0;
  const endCol = Math.max(decodedRange.e.c || 0, maxColFromRows, maxItemColIndex);
  const gridRows = [];

  for (let rowIndex = startRow; rowIndex <= endRow; rowIndex += 1) {
    const cells = [];
    let rowHasValue = false;
    let rowHasFml = false;

    for (let colIndex = startCol; colIndex <= endCol; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = worksheet[address];
      const override = overrideMap.get(address);
      const originalValue = cleanText(getCellDisplayValue(cell));
      const value = cleanText(override ? override.value : originalValue);
      const fromFml = Boolean(override?.fromFml);

      if (value) rowHasValue = true;
      if (fromFml) rowHasFml = true;

      cells.push({
        address,
        rowIndex,
        colIndex,
        value,
        originalValue,
        formula: String(cell?.f || ""),
        fromFml,
        overrideField: override?.field || "",
        fmlAddress: override?.fmlAddress || "",
      });
    }

    gridRows.push({
      rowIndex,
      excelRow: rowIndex + 1,
      cells,
      rowHasValue,
      rowHasFml,
      text: cells.map((cell) => cell.value).join(" "),
    });
  }

  return gridRows;
};

const buildResolvedCsv = ({ rows }) => {
  return rows
    .map((row) => row.cells.map((cell) => csvEscape(cell.value)).join(","))
    .join("\n");
};

const makeSafeSheetName = (value) => {
  const safe = cleanText(value || "Venue")
    .replace(/[\\/?*\[\]:]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  return (safe || "Venue").slice(0, 31);
};

const escapeSheetNameForFormula = (sheetName) =>
  `'${String(sheetName || "").replace(/'/g, "''")}'`;

const buildFmlLookupFormula = ({ codeAddress, fmlSheetName, returnIndex }) => {
  const safeCodeAddress = String(codeAddress || "").replace(/\$/g, "");
  const sheetRef = escapeSheetNameForFormula(fmlSheetName);

  return `IFERROR(VLOOKUP(${safeCodeAddress},${sheetRef}!$E:$I,${returnIndex},FALSE),IFERROR(VLOOKUP(VALUE(${safeCodeAddress}),${sheetRef}!$E:$I,${returnIndex},FALSE),""))`;
};

const ensureWorksheetRangeIncludesCell = ({ worksheet, rowIndex, colIndex, XLSX }) => {
  if (!worksheet || !XLSX) return;

  const currentRange = worksheet["!ref"]
    ? XLSX.utils.decode_range(worksheet["!ref"])
    : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };

  currentRange.s.r = Math.min(currentRange.s.r || 0, Number(rowIndex || 0));
  currentRange.s.c = Math.min(currentRange.s.c || 0, Number(colIndex || 0));
  currentRange.e.r = Math.max(currentRange.e.r || 0, Number(rowIndex || 0));
  currentRange.e.c = Math.max(currentRange.e.c || 0, Number(colIndex || 0));

  worksheet["!ref"] = XLSX.utils.encode_range(currentRange);
};

const clearCellValue = (worksheet, address) => {
  if (!worksheet || !address) return;

  if (!worksheet[address]) {
    return;
  }

  delete worksheet[address].v;
  delete worksheet[address].w;
  delete worksheet[address].f;
  delete worksheet[address].l;
  worksheet[address].t = "s";
  worksheet[address].v = "";
  worksheet[address].w = "";
};

const clearFmlGeneratedBlockCells = ({ worksheet, items, XLSX }) => {
  if (!worksheet || !XLSX) return;

  const fmlItems = (items || []).filter(
    (item) => item.sourceMode === FML_SOURCE_MODE || item.sourceMode === FML_VENUE_NOTE_SOURCE_MODE
  );
  if (!fmlItems.length) return;

  const blockMap = new Map();

  fmlItems.forEach((item) => {
    const blockKey = item.block?.blockKey;
    if (!blockKey) return;

    const columnIndexes = getFiniteColumnIndexes([
      item.block?.codeCol,
      item.block?.nameCol,
      item.block?.umCol,
    ]);

    if (!blockMap.has(blockKey)) {
      blockMap.set(blockKey, {
        block: item.block,
        maxRowIndex: Number(item.targetRowIndex || 0),
        columnIndexes,
      });
    }

    const record = blockMap.get(blockKey);
    record.maxRowIndex = Math.max(record.maxRowIndex, Number(item.targetRowIndex || 0));
    record.columnIndexes = [...new Set([...record.columnIndexes, ...columnIndexes])];
  });

  blockMap.forEach((record) => {
    const block = record.block;
    const startRowIndex = Number(block.headerRowIndex || 0) + 1;
    const existingEndRowIndex = startRowIndex + Math.max(Number(block.itemCount || 0), 0) + 8;
    const endRowIndex = Math.max(record.maxRowIndex, existingEndRowIndex);

    record.columnIndexes.forEach((colIndex) => {
      for (let rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex += 1) {
        const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
        clearCellValue(worksheet, address);
      }

      ensureWorksheetRangeIncludesCell({
        worksheet,
        rowIndex: endRowIndex,
        colIndex,
        XLSX,
      });
    });
  });
};

const writeStringCell = (worksheet, address, value) => {
  const text = cleanText(value);
  const current = worksheet[address] || {};

  worksheet[address] = {
    ...current,
    t: "s",
    v: text,
    w: text,
  };

  delete worksheet[address].f;
};

const writeFormulaCell = (worksheet, address, formula, cachedValue) => {
  const text = cleanText(cachedValue);
  const current = worksheet[address] || {};

  worksheet[address] = {
    ...current,
    t: "s",
    f: formula,
    v: text,
    w: text,
  };
};

const applyFmlCorrectionsToWorksheet = ({ worksheet, items, fmlSheetName, valuesOnly, XLSX }) => {
  if (!worksheet) return;

  clearFmlGeneratedBlockCells({ worksheet, items, XLSX });

  items.forEach((item) => {
    if (item.codeAddress) {
      writeStringCell(worksheet, item.codeAddress, item.displayCode || item.code);
      if (XLSX && Number.isFinite(Number(item.targetRowIndex)) && Number.isFinite(Number(item.block?.codeCol))) {
        ensureWorksheetRangeIncludesCell({
          worksheet,
          rowIndex: Number(item.targetRowIndex),
          colIndex: Number(item.block.codeCol),
          XLSX,
        });
      }
    }

    if (item.nameAddress) {
      if (XLSX && Number.isFinite(Number(item.targetRowIndex)) && Number.isFinite(Number(item.block?.nameCol))) {
        ensureWorksheetRangeIncludesCell({
          worksheet,
          rowIndex: Number(item.targetRowIndex),
          colIndex: Number(item.block.nameCol),
          XLSX,
        });
      }

      if (valuesOnly || !item.fmlMatched || !fmlSheetName) {
        writeStringCell(worksheet, item.nameAddress, item.productName);
      } else {
        writeFormulaCell(
          worksheet,
          item.nameAddress,
          buildFmlLookupFormula({
            codeAddress: item.codeAddress,
            fmlSheetName,
            returnIndex: 2,
          }),
          item.productName
        );
      }
    }

    if (item.umAddress) {
      if (XLSX && Number.isFinite(Number(item.targetRowIndex)) && Number.isFinite(Number(item.block?.umCol))) {
        ensureWorksheetRangeIncludesCell({
          worksheet,
          rowIndex: Number(item.targetRowIndex),
          colIndex: Number(item.block.umCol),
          XLSX,
        });
      }

      if (valuesOnly || !item.fmlMatched || !fmlSheetName) {
        writeStringCell(worksheet, item.umAddress, item.um);
      } else {
        writeFormulaCell(
          worksheet,
          item.umAddress,
          buildFmlLookupFormula({
            codeAddress: item.codeAddress,
            fmlSheetName,
            returnIndex: 5,
          }),
          item.um
        );
      }
    }
  });
};


const getSectionFilterKey = (item) => {
  const sectionIndex = Number(item?.sectionIndex || 0);
  const sectionTitle = cleanKey(item?.sectionTitle || "Section") || "SECTION";
  const blockKey = cleanKey(item?.block?.blockKey || "") || "BLOCK";

  return `${sectionIndex}|${blockKey}|${sectionTitle}`;
};

const getFiniteColumnIndexes = (values = []) =>
  values
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value) && value >= 0);

const buildSectionOptions = (items = []) => {
  const sectionMap = new Map();

  items.forEach((item) => {
    const key = getSectionFilterKey(item);
    const columnIndexes = getFiniteColumnIndexes([
      item?.block?.locatorCol,
      item?.block?.codeCol,
      item?.block?.nameCol,
      item?.block?.umCol,
    ]);

    if (!sectionMap.has(key)) {
      sectionMap.set(key, {
        key,
        title: item.sectionTitle || "Section",
        sectionIndex: Number(item.sectionIndex || 0),
        count: 0,
        minCol: columnIndexes.length ? Math.min(...columnIndexes) : 0,
        maxCol: columnIndexes.length ? Math.max(...columnIndexes) : 0,
        codeCol: Number(item?.block?.codeCol || 0),
        nameCol: Number(item?.block?.nameCol || 0),
        umCol: Number(item?.block?.umCol || 0),
        locatorCol: Number(item?.block?.locatorCol || item?.block?.codeCol || 0),
      });
    }

    const section = sectionMap.get(key);
    section.count += 1;

    if (columnIndexes.length) {
      section.minCol = Math.min(section.minCol, ...columnIndexes);
      section.maxCol = Math.max(section.maxCol, ...columnIndexes);
    }
  });

  return Array.from(sectionMap.values()).sort(
    (left, right) =>
      left.sectionIndex - right.sectionIndex ||
      left.minCol - right.minCol ||
      left.title.localeCompare(right.title)
  );
};

const filterItemsForSection = (items = [], sectionKey) => {
  if (!sectionKey || sectionKey === ALL_SECTIONS_SCOPE) return items;
  return items.filter((item) => getSectionFilterKey(item) === sectionKey);
};

const filterGridRowsForSection = (rows = [], sectionOption) => {
  if (!sectionOption) return rows;

  const minCol = Number(sectionOption.minCol || 0);
  const maxCol = Number(sectionOption.maxCol || minCol);

  return rows
    .map((row) => {
      const cells = (row.cells || []).filter(
        (cell) => cell.colIndex >= minCol && cell.colIndex <= maxCol
      );
      const rowHasValue = cells.some((cell) => cleanText(cell.value));
      const rowHasFml = cells.some((cell) => cell.fromFml);

      return {
        ...row,
        cells,
        rowHasValue,
        rowHasFml,
        text: cells.map((cell) => cell.value).join(" "),
      };
    })
    .filter((row) => row.rowHasValue || row.rowHasFml || row.excelRow <= 2);
};

const applySectionColumnFilterToWorksheet = ({ worksheet, sectionOption, XLSX }) => {
  if (!worksheet || !sectionOption || !XLSX) return;

  const decodedRange = worksheet["!ref"]
    ? XLSX.utils.decode_range(worksheet["!ref"])
    : { s: { r: 0, c: 0 }, e: { r: 0, c: Number(sectionOption.maxCol || 0) } };

  const minCol = Number(sectionOption.minCol || 0);
  const maxCol = Number(sectionOption.maxCol || minCol);
  const endCol = Math.max(decodedRange.e.c || 0, maxCol);
  const existingCols = worksheet["!cols"] || [];
  const nextCols = [];

  for (let colIndex = 0; colIndex <= endCol; colIndex += 1) {
    const current = { ...(existingCols[colIndex] || {}) };

    if (colIndex < minCol || colIndex > maxCol) {
      current.hidden = true;
    } else if (current.hidden) {
      delete current.hidden;
    }

    nextCols[colIndex] = current;
  }

  worksheet["!cols"] = nextCols;
};

const getSummary = (items) => {
  const sectionSet = new Set(items.map((item) => getSectionFilterKey(item)).filter(Boolean));

  return {
    totalItems: items.length,
    sections: sectionSet.size,
    fmlMatched: items.filter((item) => item.fmlMatched).length,
    missingFml: items.filter((item) => !item.fmlMatched).length,
    corrected: items.filter((item) => item.wasCorrected).length,
  };
};

const cardFallbackStyle = {
  background: "#fff",
  borderRadius: 16,
  padding: 20,
  boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
};

const primaryButtonFallbackStyle = {
  padding: "12px 14px",
  borderRadius: 10,
  border: 0,
  background: "#111",
  color: "#fff",
  fontWeight: "bold",
  cursor: "pointer",
};

const secondaryButtonFallbackStyle = {
  padding: "12px 14px",
  borderRadius: 10,
  border: "1px solid #ccc",
  background: "#fff",
  color: "#111",
  fontWeight: "bold",
  cursor: "pointer",
};

const tableHeaderStyle = {
  padding: 8,
  textAlign: "left",
  borderRight: "1px solid #333",
  whiteSpace: "nowrap",
  position: "sticky",
  top: 0,
  zIndex: 2,
  background: "#111",
  color: "#fff",
};

const tableCellStyle = {
  padding: "6px 8px",
  verticalAlign: "top",
  borderRight: "1px solid #eee",
  borderBottom: "1px solid #eee",
  minWidth: 90,
  maxWidth: 280,
  whiteSpace: "pre-wrap",
};

const statusBadgeStyle = (matched) => ({
  display: "inline-block",
  padding: "4px 8px",
  borderRadius: 999,
  fontWeight: "bold",
  background: matched ? "#e8f5e9" : "#fff0f0",
  color: matched ? "#2e7d32" : "#b00020",
  border: matched ? "1px solid #2e7d32" : "1px solid #b00020",
});

export default function ERPVenueIngredientsScreen({
  styles = {},
  setModule,
  logUsageEvent,
}) {
  const [workbook, setWorkbook] = useState(null);
  const [xlsxApi, setXlsxApi] = useState(null);
  const [sourceArrayBuffer, setSourceArrayBuffer] = useState(null);
  const [fileName, setFileName] = useState("");
  const [fmlSheetName, setFmlSheetName] = useState("");
  const [venueSheets, setVenueSheets] = useState([]);
  const [selectedVenueSheet, setSelectedVenueSheet] = useState("");
  const [message, setMessage] = useState("Loading ERP location template...");
  const [loading, setLoading] = useState(false);
  const [search, setSearch] = useState("");
  const [showFmlRowsOnly, setShowFmlRowsOnly] = useState(false);
  const [showUnplacedFmlOnly, setShowUnplacedFmlOnly] = useState(false);
  const [selectedSectionKey, setSelectedSectionKey] = useState(ALL_SECTIONS_SCOPE);
  const [itemSourceMode, setItemSourceMode] = useState(FML_SOURCE_MODE);

  const cardStyle = styles.card || cardFallbackStyle;
  const primaryButtonStyle = styles.primaryButton || primaryButtonFallbackStyle;
  const secondaryButtonStyle = styles.secondaryButton || secondaryButtonFallbackStyle;

  const fmlIndex = useMemo(
    () => buildFmlIndex({ workbook, fmlSheetName, XLSX: xlsxApi }),
    [workbook, fmlSheetName, xlsxApi]
  );

  const loadWorkbookFromArrayBuffer = async ({ arrayBuffer, nextFileName }) => {
    setLoading(true);
    setMessage("Reading ERP template workbook...");

    try {
      const XLSX = await loadXlsx();
      const workbookObject = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
        cellStyles: true,
        bookVBA: true,
      });

      const detectedFmlSheetName = detectFmlSheetName(workbookObject.SheetNames || []);
      const detectedVenueSheets = getVisibleVenueSheetNames(
        workbookObject,
        detectedFmlSheetName,
        XLSX
      );

      setWorkbook(workbookObject);
      setXlsxApi(XLSX);
      setSourceArrayBuffer(arrayBuffer);
      setFileName(nextFileName || "ERP location template");
      setFmlSheetName(detectedFmlSheetName);
      setVenueSheets(detectedVenueSheets);
      setSelectedVenueSheet((current) =>
        current && detectedVenueSheets.includes(current)
          ? current
          : detectedVenueSheets[0] || ""
      );
      setSelectedSectionKey(ALL_SECTIONS_SCOPE);

      if (!detectedFmlSheetName) {
        setMessage(
          "Workbook loaded, but FML March 2026 sheet was not found. Venue tabs can be shown, but FML-corrected item names and UMs cannot be rebuilt."
        );
      } else if (!detectedVenueSheets.length) {
        setMessage(
          `Workbook loaded. FML sheet found: ${detectedFmlSheetName}. No venue/location tabs with Code columns were detected.`
        );
      } else {
        setMessage(
          `Workbook loaded. FML source: ${detectedFmlSheetName}. ${detectedVenueSheets.length} venue/location tab(s) found. FML is now the source of truth; venue tabs are used as placement maps only.`
        );
      }

      logUsageEvent?.("erp_location_template_loaded", {
        module: "erp_location_ingredients",
        fileName: nextFileName || "ERP location template",
        fmlSheetName: detectedFmlSheetName,
        venueTabs: detectedVenueSheets.length,
      });
    } catch (error) {
      setWorkbook(null);
      setSourceArrayBuffer(null);
      setFileName(nextFileName || "");
      setFmlSheetName("");
      setVenueSheets([]);
      setSelectedVenueSheet("");
      setSelectedSectionKey(ALL_SECTIONS_SCOPE);
      setMessage(error?.message || "Could not read the ERP template workbook.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    let active = true;

    const loadDefaultWorkbook = async () => {
      setLoading(true);

      try {
        const response = await fetch(DEFAULT_ERP_LOCATION_TEMPLATE_PATH, {
          cache: "no-store",
        });

        if (!response.ok) {
          if (!active) return;
          setLoading(false);
          setMessage(
            "Default ERP location template was not found. Add public/erp-template-locations.xlsx or upload the file on this screen."
          );
          return;
        }

        const arrayBuffer = await response.arrayBuffer();
        if (!active) return;

        await loadWorkbookFromArrayBuffer({
          arrayBuffer,
          nextFileName: "erp-template-locations.xlsx",
        });
      } catch (error) {
        if (!active) return;
        setMessage(error?.message || "Could not load the default ERP location template.");
        setLoading(false);
      }
    };

    loadDefaultWorkbook();

    return () => {
      active = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleWorkbookUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      await loadWorkbookFromArrayBuffer({ arrayBuffer, nextFileName: file.name });
    } catch (error) {
      setMessage(error?.message || "Could not read the uploaded workbook.");
      window.alert(error?.message || "Could not read the uploaded workbook.");
    } finally {
      event.target.value = "";
    }
  };

  const selectedVenueItems = useMemo(
    () => buildVenueItemRows({
      workbook,
      sheetName: selectedVenueSheet,
      fmlIndex,
      XLSX: xlsxApi,
      sourceMode: itemSourceMode,
    }),
    [workbook, selectedVenueSheet, fmlIndex, xlsxApi, itemSourceMode]
  );

  const allVenuePlacementRows = useMemo(
    () => buildAllVenuePlacementRows({ workbook, venueSheets, fmlIndex, XLSX: xlsxApi, sourceMode: FML_SOURCE_MODE }),
    [workbook, venueSheets, fmlIndex, xlsxApi]
  );

  const fmlPlacedCodeSet = useMemo(
    () =>
      new Set(
        allVenuePlacementRows
          .filter((item) => item.fmlMatched)
          .map((item) => normalizeProductCode(item.displayCode || item.code))
          .filter(Boolean)
      ),
    [allVenuePlacementRows]
  );

  const fmlPlacementAuditRows = useMemo(
    () =>
      buildFmlPlacementAuditRows({
        fmlRows: fmlIndex.rows,
        placementRows: allVenuePlacementRows,
      }),
    [fmlIndex.rows, allVenuePlacementRows]
  );

  const fmlUnplacedRows = useMemo(
    () =>
      fmlIndex.rows.filter(
        (item) => !fmlPlacedCodeSet.has(normalizeProductCode(item.code))
      ),
    [fmlIndex.rows, fmlPlacedCodeSet]
  );

  const sectionOptions = useMemo(
    () => buildSectionOptions(selectedVenueItems),
    [selectedVenueItems]
  );

  useEffect(() => {
    if (selectedSectionKey === ALL_SECTIONS_SCOPE) return;

    const stillExists = sectionOptions.some(
      (section) => section.key === selectedSectionKey
    );

    if (!stillExists) {
      setSelectedSectionKey(ALL_SECTIONS_SCOPE);
    }
  }, [sectionOptions, selectedSectionKey]);

  const activeSectionOption = useMemo(
    () => sectionOptions.find((section) => section.key === selectedSectionKey) || null,
    [sectionOptions, selectedSectionKey]
  );

  const sectionFilteredItems = useMemo(
    () => filterItemsForSection(selectedVenueItems, selectedSectionKey),
    [selectedVenueItems, selectedSectionKey]
  );

  const selectedGridRows = useMemo(
    () => buildCorrectedGridRows({ workbook, sheetName: selectedVenueSheet, items: sectionFilteredItems, XLSX: xlsxApi }),
    [workbook, selectedVenueSheet, sectionFilteredItems, xlsxApi]
  );

  const sectionFilteredGridRows = useMemo(
    () => filterGridRowsForSection(selectedGridRows, activeSectionOption),
    [selectedGridRows, activeSectionOption]
  );

  const summary = useMemo(() => getSummary(sectionFilteredItems), [sectionFilteredItems]);
  const activeSectionName = activeSectionOption?.title || "All sections";
  const activeReportName = activeSectionOption
    ? `${selectedVenueSheet} - ${activeSectionName}`
    : selectedVenueSheet;

  const visibleGridRows = useMemo(() => {
    const term = search.toLowerCase().trim();
    let rows = sectionFilteredGridRows;

    if (showFmlRowsOnly) {
      rows = rows.filter((row) => row.rowHasFml);
    }

    if (term) {
      rows = rows.filter((row) => row.text.toLowerCase().includes(term));
    }

    return rows.slice(0, MAX_PREVIEW_ROWS);
  }, [sectionFilteredGridRows, search, showFmlRowsOnly]);

  const visibleItemRows = useMemo(() => {
    const term = search.toLowerCase().trim();

    return sectionFilteredItems.filter((item) => {
      if (showFmlRowsOnly && !item.fmlMatched) return false;
      if (!term) return true;

      return [
        item.venueSheet,
        item.sectionTitle,
        item.code,
        item.productName,
        item.um,
        item.department,
        item.category,
        item.subCategory,
        item.type,
        item.brand,
        item.allergens,
        item.comments,
        item.adjustmentComments,
        item.notes,
        item.venueText,
        item.sourceLabel,
        item.status,
        item.sourceRow,
        item.fmlRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  }, [sectionFilteredItems, search, showFmlRowsOnly]);

  const visibleFmlPlacementAuditRows = useMemo(() => {
    const term = search.toLowerCase().trim();

    return fmlPlacementAuditRows.filter((row) => {
      if (showUnplacedFmlOnly && row.placementStatus !== "Not placed in venue tabs") {
        return false;
      }

      if (!term) return true;

      return [
        row.placementStatus,
        row.fmlRow,
        row.code,
        row.displayCode,
        row.productName,
        row.um,
        row.department,
        row.category,
        row.subCategory,
        row.type,
        row.brand,
        row.venueSheet,
        row.sectionTitle,
        row.sourceRow,
        row.templateCell,
        row.comments,
        row.adjustmentComments,
        row.notes,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  }, [fmlPlacementAuditRows, search, showUnplacedFmlOnly]);

  const downloadFmlPlacementAuditExcel = async () => {
    if (!visibleFmlPlacementAuditRows.length || !xlsxApi) {
      window.alert("No FML placement audit rows to export.");
      return;
    }

    try {
      const XLSX = xlsxApi;
      const outputWorkbook = XLSX.utils.book_new();
      const rows = getFmlPlacementAuditExportRows(visibleFmlPlacementAuditRows);
      const worksheet = XLSX.utils.json_to_sheet(rows);

      worksheet["!cols"] = [
        { wch: 8 },
        { wch: 24 },
        { wch: 10 },
        { wch: 14 },
        { wch: 48 },
        { wch: 12 },
        { wch: 18 },
        { wch: 36 },
        { wch: 24 },
        { wch: 16 },
        { wch: 18 },
        { wch: 28 },
        { wch: 48 },
        { wch: 12 },
        { wch: 14 },
        { wch: 34 },
        { wch: 34 },
      ];

      XLSX.utils.book_append_sheet(outputWorkbook, worksheet, "FML Placement Audit");
      XLSX.writeFile(
        outputWorkbook,
        `erp-fml-placement-audit-${getDateStamp()}.xlsx`,
        { bookType: "xlsx" }
      );

      logUsageEvent?.("erp_fml_placement_audit_downloaded", {
        module: "erp_location_ingredients",
        rows: visibleFmlPlacementAuditRows.length,
        unplacedFmlRows: fmlUnplacedRows.length,
        fmlRows: fmlIndex.rows.length,
      });
    } catch (error) {
      window.alert(error?.message || "Could not download the FML placement audit.");
    }
  };

  const downloadSelectedVenueLinkedWorkbook = async () => {
    if (!workbook || !xlsxApi || !selectedVenueSheet) {
      window.alert("Choose a venue tab first.");
      return;
    }

    try {
      const XLSX = xlsxApi;
      const outputWorkbook = XLSX.utils.book_new();

      outputWorkbook.Sheets = {};
      outputWorkbook.SheetNames = [];

      const correctedVenueSheet = cloneSheet(workbook.Sheets[selectedVenueSheet]);
      applyFmlCorrectionsToWorksheet({
        worksheet: correctedVenueSheet,
        items: sectionFilteredItems,
        fmlSheetName,
        valuesOnly: false,
        XLSX,
      });

      applySectionColumnFilterToWorksheet({
        worksheet: correctedVenueSheet,
        sectionOption: activeSectionOption,
        XLSX,
      });

      outputWorkbook.Sheets[selectedVenueSheet] = correctedVenueSheet;
      outputWorkbook.SheetNames.push(selectedVenueSheet);

      if (fmlSheetName && workbook.Sheets[fmlSheetName]) {
        outputWorkbook.Sheets[fmlSheetName] = cloneSheet(workbook.Sheets[fmlSheetName]);
        outputWorkbook.SheetNames.push(fmlSheetName);
        outputWorkbook.Workbook = {
          Sheets: [
            { name: selectedVenueSheet, Hidden: 0 },
            { name: fmlSheetName, Hidden: 1 },
          ],
          CalcPr: { calcMode: "auto" },
        };
      }

      const sectionPart = activeSectionOption ? `-${makeSafeFilePart(activeSectionName)}` : "-all-sections";
      const sourcePart = itemSourceMode === FML_SOURCE_MODE ? "fml-venue-located" : "template-codes-fml-linked";
      const outputName = `erp-${makeSafeFilePart(selectedVenueSheet)}${sectionPart}-${sourcePart}-${getDateStamp()}.xlsx`;

      XLSX.writeFile(outputWorkbook, outputName, {
        bookType: "xlsx",
        cellStyles: true,
      });

      logUsageEvent?.("erp_selected_venue_fml_linked_workbook_downloaded", {
        module: "erp_location_ingredients",
        venueSheet: selectedVenueSheet,
        fmlSheetName,
        outputName,
        section: activeSectionName,
        rows: sectionFilteredItems.length,
        sourceMode: itemSourceMode,
      });
    } catch (error) {
      window.alert(error?.message || "Could not download the selected FML-linked venue workbook.");
    }
  };

  const downloadFullCorrectedWorkbook = async () => {
    if (!workbook || !xlsxApi || !sourceArrayBuffer) {
      window.alert("Upload or load the ERP location template first.");
      return;
    }

    try {
      const XLSX = xlsxApi;
      const outputWorkbook = {
        ...workbook,
        SheetNames: [...(workbook.SheetNames || [])],
        Sheets: {},
        Workbook: workbook.Workbook ? JSON.parse(JSON.stringify(workbook.Workbook)) : undefined,
      };

      const venueSheetSet = new Set(venueSheets);

      (workbook.SheetNames || []).forEach((sheetName) => {
        const clonedSheet = cloneSheet(workbook.Sheets[sheetName]);

        if (venueSheetSet.has(sheetName)) {
          const sheetItems = buildVenueItemRows({
            workbook,
            sheetName,
            fmlIndex,
            XLSX,
            sourceMode: itemSourceMode,
          });

          applyFmlCorrectionsToWorksheet({
            worksheet: clonedSheet,
            items: sheetItems,
            fmlSheetName,
            valuesOnly: false,
            XLSX,
          });
        }

        outputWorkbook.Sheets[sheetName] = clonedSheet;
      });

      const baseName = fileName.replace(/\.(xlsx|xlsm|xls)$/i, "") || "erp-template-locations";
      const sourcePart = itemSourceMode === FML_SOURCE_MODE ? "fml-venue-located" : "template-codes-fml-linked";
      const outputName = `${makeSafeFilePart(baseName)}-all-venues-${sourcePart}-${getDateStamp()}.xlsx`;

      XLSX.writeFile(outputWorkbook, outputName, {
        bookType: "xlsx",
        cellStyles: true,
      });

      logUsageEvent?.("erp_full_fml_linked_workbook_downloaded", {
        module: "erp_location_ingredients",
        outputName,
        venueTabs: venueSheets.length,
        fmlSheetName,
        sourceMode: itemSourceMode,
      });
    } catch (error) {
      window.alert(error?.message || "Could not download the full corrected workbook.");
    }
  };

  const downloadSelectedVenueResolvedExcel = async () => {
    if (!visibleGridRows.length || !xlsxApi) {
      window.alert("No visible rows to export.");
      return;
    }

    try {
      const XLSX = xlsxApi;
      const values = visibleGridRows.map((row) => row.cells.map((cell) => cell.value));
      const outputWorkbook = XLSX.utils.book_new();
      const worksheet = XLSX.utils.aoa_to_sheet(values);

      worksheet["!cols"] = values[0]?.map(() => ({ wch: 18 })) || [];

      XLSX.utils.book_append_sheet(
        outputWorkbook,
        worksheet,
        makeSafeSheetName(activeReportName || "Venue")
      );

      const outputName = `erp-${makeSafeFilePart(activeReportName || selectedVenueSheet)}-fml-values-${getDateStamp()}.xlsx`;
      XLSX.writeFile(outputWorkbook, outputName);

      logUsageEvent?.("erp_selected_venue_fml_values_downloaded", {
        module: "erp_location_ingredients",
        venueSheet: selectedVenueSheet,
        outputName,
        section: activeSectionName,
        rows: visibleGridRows.length,
      });
    } catch (error) {
      window.alert(error?.message || "Could not download FML values workbook.");
    }
  };

  const downloadVisibleCsv = () => {
    if (!visibleGridRows.length) {
      window.alert("No visible rows to export.");
      return;
    }

    downloadBlob({
      content: buildResolvedCsv({ rows: visibleGridRows }),
      fileName: `erp-${makeSafeFilePart(activeReportName || selectedVenueSheet)}-fml-corrected-${getDateStamp()}.csv`,
      type: "text/csv;charset=utf-8",
    });
  };

  const printVisibleReport = () => {
    if (!visibleGridRows.length) {
      window.alert("No visible rows to print.");
      return;
    }

    const html = `
      <html>
        <head>
          <title>ERP Venue Ingredients - ${escapeHtmlValue(activeReportName || selectedVenueSheet)}</title>
          <style>
            @page { size: landscape; margin: 10mm; }
            body { font-family: Arial, sans-serif; color: #111; padding: 18px; }
            h1 { margin: 0 0 4px; font-size: 22px; }
            .meta { color: #444; font-size: 12px; margin: 2px 0; }
            table { border-collapse: collapse; margin-top: 14px; font-size: 10px; }
            th, td { border: 1px solid #ccc; padding: 5px; text-align: left; vertical-align: top; }
            th { background: #111; color: #fff; }
            .linked { background: #eef5ff; font-weight: 700; }
            .missing { color: #b00020; font-weight: 700; }
            tr { break-inside: avoid; }
          </style>
        </head>
        <body>
          <h1>ERP Venue Ingredients - FML Corrected</h1>
          <div class="meta"><strong>Venue tab:</strong> ${escapeHtmlValue(selectedVenueSheet || "N/A")}</div>
          <div class="meta"><strong>Section:</strong> ${escapeHtmlValue(activeSectionName)}</div>
          <div class="meta"><strong>FML source:</strong> ${escapeHtmlValue(fmlSheetName || "Not found")}</div>
          <div class="meta"><strong>Source file:</strong> ${escapeHtmlValue(fileName || "N/A")}</div>
          <div class="meta"><strong>Visible rows:</strong> ${escapeHtmlValue(visibleGridRows.length)}</div>
          <div class="meta"><strong>Printed:</strong> ${escapeHtmlValue(new Date().toLocaleString())}</div>
          <table>
            <tbody>
              ${visibleGridRows
                .map(
                  (row) => `
                    <tr>
                      <th>${escapeHtmlValue(row.excelRow)}</th>
                      ${row.cells
                        .map(
                          (cell) => `
                            <td class="${cell.fromFml ? "linked" : ""}" title="${escapeHtmlValue(cell.fmlAddress || cell.address)}">
                              ${escapeHtmlValue(cell.value)}
                            </td>
                          `
                        )
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
      window.alert("Print window was blocked. Allow popups for this app and try again.");
      return;
    }

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();

    logUsageEvent?.("erp_venue_fml_corrected_report_printed", {
      module: "erp_location_ingredients",
      venueSheet: selectedVenueSheet,
      section: activeSectionName,
      rows: visibleGridRows.length,
    });
  };

  return (
    <div style={styles.page || { minHeight: "100vh", padding: 24, background: "#f5f5f5", fontFamily: "Arial, sans-serif" }}>
      <div style={styles.header || { display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", background: "#fff", borderRadius: 16, padding: 18, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ minWidth: 280 }}>
          <h1 style={{ margin: 0 }}>ERP Venue Ingredients</h1>
          <p style={styles.subtitle || { margin: "4px 0 0", color: "#666" }}>
            FML is the source of truth for product details. Venue tabs and row-1 locators decide the correct venue/location for each FML code, so FML items no longer appear in every venue unless that venue tab already contains the code or FML has a matching venue note.
          </p>
        </div>

        <button type="button" style={styles.backButton || secondaryButtonStyle} onClick={() => setModule?.("")}>
          Back
        </button>
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 14 }}>
        <div>
          <h2 style={{ margin: 0 }}>Files</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            Save the workbook as <strong>public/erp-template-locations.xlsx</strong>, or upload it here. The app uses Excel row 1 on each venue/location tab as the section locator, then reads the product codes underneath that locator. FML March 2026 is the source of truth for product details. Venue/location tabs are the placement map: tab name, row-1 locator, code column, and row position. FML venue notes can add new items only when they match the venue.
          </p>
        </div>

        <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
          Upload ERP template by locations
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={handleWorkbookUpload}
            disabled={loading}
            style={styles.fileInput}
          />
          <span style={{ color: "#666", fontSize: 13, fontWeight: "normal" }}>
            Current: {fileName || "No workbook loaded"}
          </span>
        </label>

        {message ? <div style={styles.infoBox || { padding: 12, borderRadius: 12, background: "#f2f2f2" }}>{message}</div> : null}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(145px, 1fr))", gap: 10, marginBottom: 16 }}>
        {[
          ["Venue tabs", venueSheets.length],
          ["FML rows", fmlIndex.rows.length],
          ["FML placed", fmlPlacedCodeSet.size],
          ["FML not placed", fmlUnplacedRows.length],
          ["All placements", allVenuePlacementRows.length],
          ["Sections", summary.sections],
          ["Selected items", summary.totalItems],
          ["Visible rows", visibleGridRows.length],
        ].map(([label, value]) => (
          <div key={label} style={{ ...cardStyle, padding: 14, textAlign: "center" }}>
            <div style={{ color: "#666", fontSize: 12, fontWeight: "bold", textTransform: "uppercase" }}>{label}</div>
            <div style={{ fontSize: 24, fontWeight: 900, marginTop: 4 }}>{value}</div>
          </div>
        ))}
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 14 }}>
        <h2 style={{ margin: 0 }}>View options</h2>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
          <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
            Venue / location tab
            <select
              value={selectedVenueSheet}
              onChange={(event) => {
                setSelectedVenueSheet(event.target.value);
                setSelectedSectionKey(ALL_SECTIONS_SCOPE);
              }}
              style={styles.selectInput || { width: "100%", padding: 11, borderRadius: 10, border: "1px solid #ccc", background: "#fff" }}
            >
              {venueSheets.length ? (
                venueSheets.map((sheetName) => (
                  <option key={sheetName} value={sheetName}>
                    {sheetName}
                  </option>
                ))
              ) : (
                <option value="">No venue tabs found</option>
              )}
            </select>
          </label>

          <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
            Item source
            <select
              value={itemSourceMode}
              onChange={(event) => {
                setItemSourceMode(event.target.value);
                setSelectedSectionKey(ALL_SECTIONS_SCOPE);
              }}
              style={styles.selectInput || { width: "100%", padding: 11, borderRadius: 10, border: "1px solid #ccc", background: "#fff" }}
            >
              <option value={FML_SOURCE_MODE}>FML via venue/location tabs</option>
              <option value={TEMPLATE_CODE_SOURCE_MODE}>Existing venue-tab codes</option>
            </select>
          </label>

          <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
            Section from row 1
            <select
              value={selectedSectionKey}
              onChange={(event) => setSelectedSectionKey(event.target.value)}
              style={styles.selectInput || { width: "100%", padding: 11, borderRadius: 10, border: "1px solid #ccc", background: "#fff" }}
            >
              <option value={ALL_SECTIONS_SCOPE}>All sections in this venue</option>
              {sectionOptions.map((section) => (
                <option key={section.key} value={section.key}>
                  {section.title} ({section.count})
                </option>
              ))}
            </select>
          </label>

          <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
            Search
            <input
              type="search"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search venue, section, code, FML item, UM, category, comments, notes..."
              style={styles.searchInput || { width: "100%", padding: 11, borderRadius: 10, border: "1px solid #ccc" }}
            />
          </label>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: "bold" }}>
          <input
            type="checkbox"
            checked={showFmlRowsOnly}
            onChange={(event) => setShowFmlRowsOnly(event.target.checked)}
          />
          Show only visible rows placed from FML
        </label>

        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button type="button" style={primaryButtonStyle} onClick={downloadSelectedVenueLinkedWorkbook} disabled={!selectedVenueSheet}>
            Download Selected Venue / Section Excel
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadFullCorrectedWorkbook} disabled={!sourceArrayBuffer}>
            Download Full FML-Built Workbook
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadFmlPlacementAuditExcel} disabled={!visibleFmlPlacementAuditRows.length}>
            Download FML Placement Audit
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadSelectedVenueResolvedExcel} disabled={!visibleGridRows.length}>
            Download Selected Values Excel
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={printVisibleReport} disabled={!visibleGridRows.length}>
            Print / Save PDF
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadVisibleCsv} disabled={!visibleGridRows.length}>
            Export CSV
          </button>
        </div>
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0 }}>FML-Corrected Item List</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            Showing {visibleItemRows.length} placement(s) from <strong>{selectedVenueSheet || "No venue"}</strong> / <strong>{activeSectionName}</strong>. Excel row 1 locates each section, codes stay under the Code column, and product name / UM come from FML March 2026.
          </p>
        </div>

        <div style={{ overflowX: "auto", border: "1px solid #ddd", borderRadius: 14 }}>
          <table style={{ width: "100%", minWidth: 1250, borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#111", color: "#fff" }}>
                <th style={tableHeaderStyle}>Section</th>
                <th style={tableHeaderStyle}>Row</th>
                <th style={tableHeaderStyle}>Code</th>
                <th style={tableHeaderStyle}>FML Ingredient Name</th>
                <th style={tableHeaderStyle}>UM</th>
                <th style={tableHeaderStyle}>SubCategory</th>
                <th style={tableHeaderStyle}>Type</th>
                <th style={tableHeaderStyle}>Brand</th>
                <th style={tableHeaderStyle}>Status</th>
                <th style={tableHeaderStyle}>Comments</th>
                <th style={tableHeaderStyle}>Notes</th>
              </tr>
            </thead>
            <tbody>
              {visibleItemRows.length ? (
                visibleItemRows.map((item) => (
                  <tr key={item.key}>
                    <td style={tableCellStyle}>{item.sectionTitle}</td>
                    <td style={tableCellStyle}>{item.sourceRow}</td>
                    <td style={tableCellStyle}>{item.displayCode}</td>
                    <td style={{ ...tableCellStyle, fontWeight: "bold" }}>{item.productName}</td>
                    <td style={tableCellStyle}>{item.um || "--"}</td>
                    <td style={tableCellStyle}>{item.subCategory || "--"}</td>
                    <td style={tableCellStyle}>{item.type || "--"}</td>
                    <td style={tableCellStyle}>{item.brand || "--"}</td>
                    <td style={tableCellStyle}>
                      <span style={statusBadgeStyle(item.fmlMatched)}>{item.status}</span>
                      {item.wasCorrected ? <div style={{ color: "#0057b8", fontWeight: "bold", marginTop: 5 }}>Updated from FML</div> : null}
                    </td>
                    <td style={tableCellStyle}>{item.comments || item.adjustmentComments || "--"}</td>
                    <td style={tableCellStyle}>{item.notes || "--"}</td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={11} style={{ ...tableCellStyle, textAlign: "center", color: "#777", padding: 22 }}>
                    Choose a venue tab or upload the ERP template workbook.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 12 }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
          <div>
            <h2 style={{ margin: 0 }}>FML Placement Audit</h2>
            <p style={styles.message || { color: "#555", fontSize: 14 }}>
              This checks every item from <strong>{fmlSheetName || DEFAULT_FML_SHEET_NAME}</strong> against all venue/location tabs. Placed items show the exact venue tab and row-1 locator. Items not found in any venue tab are listed as <strong>Not placed in venue tabs</strong>.
            </p>
          </div>
          <button type="button" style={secondaryButtonStyle} onClick={downloadFmlPlacementAuditExcel} disabled={!visibleFmlPlacementAuditRows.length}>
            Download FML Placement Audit
          </button>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: "bold" }}>
          <input
            type="checkbox"
            checked={showUnplacedFmlOnly}
            onChange={(event) => setShowUnplacedFmlOnly(event.target.checked)}
          />
          Show only FML items not placed in any venue tab
        </label>

        <div style={{ overflowX: "auto", border: "1px solid #ddd", borderRadius: 14, maxHeight: 420 }}>
          <table style={{ width: "100%", minWidth: 1450, borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#111", color: "#fff" }}>
                <th style={tableHeaderStyle}>Status</th>
                <th style={tableHeaderStyle}>FML Row</th>
                <th style={tableHeaderStyle}>Code</th>
                <th style={tableHeaderStyle}>Product Name</th>
                <th style={tableHeaderStyle}>UM</th>
                <th style={tableHeaderStyle}>Department</th>
                <th style={tableHeaderStyle}>Category</th>
                <th style={tableHeaderStyle}>SubCategory</th>
                <th style={tableHeaderStyle}>Venue Tab</th>
                <th style={tableHeaderStyle}>Row 1 Locator / Section</th>
                <th style={tableHeaderStyle}>Template Row</th>
              </tr>
            </thead>
            <tbody>
              {visibleFmlPlacementAuditRows.length ? (
                visibleFmlPlacementAuditRows.map((row) => (
                  <tr key={row.key}>
                    <td style={tableCellStyle}>
                      <span style={statusBadgeStyle(row.placementStatus !== "Not placed in venue tabs")}>{row.placementStatus}</span>
                    </td>
                    <td style={tableCellStyle}>{row.fmlRow || "--"}</td>
                    <td style={tableCellStyle}>{row.displayCode || row.code}</td>
                    <td style={{ ...tableCellStyle, fontWeight: "bold" }}>{row.productName}</td>
                    <td style={tableCellStyle}>{row.um || "--"}</td>
                    <td style={tableCellStyle}>{row.department || "--"}</td>
                    <td style={tableCellStyle}>{row.category || "--"}</td>
                    <td style={tableCellStyle}>{row.subCategory || "--"}</td>
                    <td style={tableCellStyle}>{row.venueSheet || "--"}</td>
                    <td style={tableCellStyle}>{row.sectionTitle || "--"}</td>
                    <td style={tableCellStyle}>{row.sourceRow || "--"}</td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={11} style={{ ...tableCellStyle, textAlign: "center", color: "#777", padding: 22 }}>
                    No FML placement audit rows found. Load the ERP workbook first.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div style={{ ...cardStyle, display: "grid", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0 }}>Corrected Template Preview</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            This keeps the same row and column layout as the venue tab. Use the section dropdown to create a report for one row-1 locator, such as “06 - Garde Manger - 15 - Menu Basic Preperation 01/06/15”, or choose all sections. Code stays in the Code column, product name stays in the Ingredient Name column, and UM comes only from the FML sheet.
            {visibleGridRows.length >= MAX_PREVIEW_ROWS ? ` Preview is limited to ${MAX_PREVIEW_ROWS} rows for browser speed.` : ""}
          </p>
        </div>

        <div style={{ overflow: "auto", border: "1px solid #ddd", borderRadius: 14, maxHeight: "70vh" }}>
          <table style={{ borderCollapse: "collapse", fontSize: 12, minWidth: 1100, background: "#fff" }}>
            <tbody>
              {visibleGridRows.length ? (
                visibleGridRows.map((row) => (
                  <tr key={row.excelRow}>
                    <th style={{ ...tableHeaderStyle, left: 0, zIndex: 3, minWidth: 52 }}>
                      {row.excelRow}
                    </th>
                    {row.cells.map((cell) => (
                      <td
                        key={cell.address}
                        title={cell.fromFml ? `FML: ${cell.fmlAddress}` : cell.formula ? `Formula: ${cell.formula}` : cell.address}
                        style={{
                          ...tableCellStyle,
                          background: cell.fromFml ? "#eef5ff" : row.rowHasValue ? "#fff" : "#fafafa",
                          color: cell.value ? "#111" : "#aaa",
                          fontWeight: cell.fromFml ? "700" : "normal",
                        }}
                      >
                        {cell.value || ""}
                      </td>
                    ))}
                  </tr>
                ))
              ) : (
                <tr>
                  <td style={{ ...tableCellStyle, textAlign: "center", color: "#777", padding: 24 }}>
                    Upload the ERP template by locations workbook to view venue ingredients.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

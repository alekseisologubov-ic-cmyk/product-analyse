"use client";

import React, { useMemo, useState } from "react";

const DEFAULT_INGREDIENT_BY_LOCATION_PATH = "/ingredient-by-location.xlsx";
const DEFAULT_ERP_TEMPLATE_PATH = "/erp-template-locations.xlsx";
const FML_SHEET_NAME = "FML March 2026";
const UOM_SHEET_NAME = "Unit of Measure ship March 2026";
const ALL_VENUES_SCOPE = "ALL";
const MAX_PREVIEW_ROWS = 400;

const loadXlsx = async () => {
  const module = await import("xlsx");
  return module.default || module;
};

const yieldToBrowser = () =>
  new Promise((resolve) => {
    if (typeof window === "undefined") {
      resolve();
      return;
    }

    window.setTimeout(resolve, 0);
  });

const cleanText = (value) =>
  String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const cleanKey = (value) =>
  cleanText(value)
    .toUpperCase()
    .replace(/&/g, " AND ")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const normalizeHeader = (value) => cleanKey(value).replace(/[^A-Z0-9]/g, "");

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

const isLikelyProductCode = (value) => {
  const code = normalizeProductCode(value);
  return /^\d{3,}$/.test(code);
};

const excelAddress = (XLSX, rowIndex, columnIndex) =>
  XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });

const getWorksheetRows = ({ worksheet, XLSX, raw = false }) => {
  if (!worksheet || !XLSX) return [];

  return XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw,
  });
};

const getHeaderMap = (row = []) => {
  const map = new Map();

  row.forEach((cell, index) => {
    const key = normalizeHeader(cell);
    if (key && !map.has(key)) {
      map.set(key, index);
    }
  });

  return map;
};

const getColumnIndex = (headerMap, aliases = [], fallbackIndex = -1) => {
  for (const alias of aliases) {
    const index = headerMap.get(normalizeHeader(alias));
    if (index !== undefined) return index;
  }

  return fallbackIndex;
};

const isActiveFlag = (value) => {
  const text = cleanKey(value);

  if (!text) return true;
  if (["N", "NO", "FALSE", "INACTIVE", "0"].includes(text)) return false;

  return true;
};

const cloneCell = (cell) => {
  if (!cell) return null;
  return JSON.parse(JSON.stringify(cell));
};

const cloneStyle = (style) => {
  if (!style) return undefined;
  return JSON.parse(JSON.stringify(style));
};

const applyRedFontStyle = (cell) => {
  const currentStyle = cloneStyle(cell.s) || {};

  cell.s = {
    ...currentStyle,
    font: {
      ...(currentStyle.font || {}),
      color: { rgb: "FF0000" },
    },
  };

  return cell;
};

const setCell = ({ worksheet, address, cell, red = false }) => {
  const nextCell = cloneCell(cell) || { t: "s", v: "" };

  if (red) {
    applyRedFontStyle(nextCell);
  }

  worksheet[address] = nextCell;
};

const ensureWorksheetRange = ({ worksheet, XLSX, rowIndex, columnIndex }) => {
  const decoded = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");

  decoded.s.r = Math.min(decoded.s.r, rowIndex);
  decoded.s.c = Math.min(decoded.s.c, columnIndex);
  decoded.e.r = Math.max(decoded.e.r, rowIndex);
  decoded.e.c = Math.max(decoded.e.c, columnIndex);

  worksheet["!ref"] = XLSX.utils.encode_range(decoded);
};

const makeSafeFilePart = (value) =>
  String(value || "report")
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

const escapeRegExp = (value) => String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const aliasMatchesText = (textKey, alias) => {
  const aliasKey = cleanKey(alias);

  if (!textKey || !aliasKey) return false;

  if (aliasKey.length <= 3) {
    return new RegExp(`(^| )${escapeRegExp(aliasKey)}( |$)`).test(textKey);
  }

  return textKey === aliasKey || textKey.includes(aliasKey);
};

const VENUE_ALIAS_RULES = [
  { match: ["06 GARDE MANGER"], aliases: ["GARDE MANGER", "GM"] },
  { match: ["07 KITCHEN TABLE"], aliases: ["KITCHEN TABLE", "THE KITCHEN TABLE", "KT"] },
  { match: ["08 SPECIAL SERVICE FOOD"], aliases: ["SPECIAL SERVICE FOOD", "SPECIAL SERVICE", "SSF"] },
  { match: ["13 TEST KITCHEN"], aliases: ["TEST KITCHEN", "TK"] },
  { match: ["14 PINK AGAVE"], aliases: ["PINK AGAVE", "PA"] },
  { match: ["15 RAZZLE"], aliases: ["RAZZLE", "RAZZLE DAZZLE", "RD"] },
  { match: ["ROJO"], aliases: ["ROJO", "RAZZLE ROJO"] },
  { match: ["ARIYA"], aliases: ["ARIYA", "RAZZLE ARIYA"] },
  { match: ["16 EXTRA VIRGIN"], aliases: ["EXTRA VIRGIN", "EV"] },
  { match: ["17 WAKE"], aliases: ["WAKE", "THE WAKE"] },
  { match: ["18 GUNBAE"], aliases: ["GUNBAE", "GB"] },
  { match: ["19 MANNOR", "19 MANOR"], aliases: ["MANNOR", "MANOR", "ANOTHER ROSE"] },
  { match: ["LOLZ"], aliases: ["LOLZ"] },
  { match: ["20 THE DOCK"], aliases: ["DOCK", "THE DOCK", "UP WITH A TWIST"] },
  { match: ["21 THE PIZZA PLACE"], aliases: ["PIZZA", "PIZZA PLACE", "THE PIZZA PLACE"] },
  { match: ["22 LTS POPSTAR", "22 ITS POPSTAR"], aliases: ["POPSTAR", "LTS POPSTAR", "ITS POPSTAR"] },
  { match: ["23 SOCIAL CLUB"], aliases: ["SOCIAL CLUB"] },
  { match: ["24 SUN CLUB"], aliases: ["SUN CLUB"] },
  { match: ["25 SIP LOUNGE"], aliases: ["SIP", "SIP LOUNGE"] },
  { match: ["26 SHIP EATS"], aliases: ["SHIP EATS"] },
  { match: ["27 QUICKEZE"], aliases: ["QUICKEZE", "QUICKIEZE"] },
  { match: ["28 DINE AND DASH"], aliases: ["DINE AND DASH", "EAT AND DRINK"] },
  { match: ["29 BURGER BAR"], aliases: ["BURGER", "BURGER BAR"] },
  { match: ["30 HOT OF THE PRESS"], aliases: ["HOT OF THE PRESS"] },
  { match: ["32 TACO"], aliases: ["TACO"] },
  { match: ["33 BENTO BABY"], aliases: ["BENTO BABY", "DIM SUM"] },
  { match: ["34 DAILY MIX"], aliases: ["DAILY MIX"] },
  { match: ["35 NOODLE AROUND"], aliases: ["NOODLE", "NOODLE AROUND"] },
  { match: ["36 BIMINI"], aliases: ["BIMINI", "BIMNI", "BIMINI BEACH CLUB", "BBC"] },
  { match: ["37 THE GALLEY KITCHEN"], aliases: ["GALLEY", "GALLEY KITCHEN", "THE GALLEY KITCHEN"] },
  { match: ["39 CHARGABLE", "39 CHARGEABLE"], aliases: ["CHARGABLE", "CHARGEABLE", "CHARGEABLE ITEM"] },
  { match: ["44 PASTRY", "44 BAKERY"], aliases: ["PASTRY", "BAKERY", "PASTRY BAKERY"] },
  { match: ["LIQUER", "LIQUOR"], aliases: ["LIQUER", "LIQUOR", "BAR"] },
];

const normalizeVenueForMatch = (value) =>
  cleanKey(value)
    .replace(/^\d+\s+/, "")
    .replace(/\bVV\b/g, "")
    .replace(/\bONLY\b/g, "")
    .replace(/\bSCL\b/g, "")
    .replace(/\bVAL\b/g, "")
    .replace(/\bRES\b/g, "")
    .replace(/\bBRL\b/g, "")
    .replace(/\bLADY\b/g, "")
    .replace(/\s+/g, " ")
    .trim();

const getVenueAliasesForSheetName = (sheetName) => {
  const sheetKey = cleanKey(sheetName);
  const aliases = new Set([
    sheetName,
    normalizeVenueForMatch(sheetName),
  ]);

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

const ingredientRowMatchesVenueSheet = (ingredientRow, sheetName) => {
  const textKey = normalizeVenueForMatch([
    ingredientRow?.restaurantName,
    ingredientRow?.menuName,
  ].join(" "));

  if (!textKey) return false;

  const sheetKey = normalizeVenueForMatch(sheetName);
  const aliases = getVenueAliasesForSheetName(sheetName);

  if (textKey.includes("ROJO") && !sheetKey.includes("ROJO")) return false;
  if (textKey.includes("ARIYA") && !sheetKey.includes("ARIYA")) return false;

  if (aliases.some((alias) => aliasMatchesText(textKey, alias))) return true;
  if (sheetKey && sheetKey.length > 5 && textKey.includes(sheetKey)) return true;

  return false;
};

const getTemplateSectionType = (sectionTitle) => {
  const key = cleanKey(sectionTitle);

  if (!key) return "general";
  if (key.includes("LIQUOR") || key.includes("LIQUER") || key.includes("WINE") || key.includes("BEER") || key.includes("ALCOHOL")) return "liquor";
  if (key.includes("CHARG")) return "chargeable";
  if (key.includes("VEGET") || key.includes("VEGETE") || key.includes("FRUIT") || key.includes("PRODUCE")) return "produce";
  if (key.includes("FISH") || key.includes("SEAFOOD")) return "fish";
  if (key.includes("MEAT") || key.includes("BUTCHER")) return "meat";
  if (key.includes("PASTRY")) return "pastry";
  if (key.includes("BAKERY") || key.includes("BAKER")) return "bakery";
  if (key.includes("MENU BASIC") || key.includes("MENU BADIC") || key.includes("PREPERATION") || key.includes("PREPARATION")) return "menuBasic";

  return "general";
};

const ingredientRowSectionKey = (row) =>
  cleanKey([
    row?.category,
    row?.subCategory,
    row?.name,
    row?.assigned,
    row?.productName,
    row?.recipeName,
    row?.specialInstructions,
    row?.specialInstructions2,
  ].join(" "));

const ingredientLooksLikeProduce = (row) => {
  const key = ingredientRowSectionKey(row);
  return key.includes("VEGET") || key.includes("VEGETE") || key.includes("FRUIT") || key.includes("PRODUCE") || key.includes("FRESH VEGETABLE");
};

const ingredientLooksLikeFish = (row) => {
  const key = ingredientRowSectionKey(row);
  return key.includes("FISH") || key.includes("SEAFOOD") || key.includes("SALMON") || key.includes("TUNA") || key.includes("SHRIMP") || key.includes("LOBSTER") || key.includes("CRAB") || key.includes("SCALLOP");
};

const ingredientLooksLikeMeat = (row) => {
  const key = ingredientRowSectionKey(row);
  return key.includes("MEAT") || key.includes("BEEF") || key.includes("VEAL") || key.includes("PORK") || key.includes("LAMB") || key.includes("POULTRY") || key.includes("CHICKEN") || key.includes("TURKEY") || key.includes("BACON") || key.includes("SAUSAGE");
};

const ingredientLooksLikeBakery = (row) => {
  const key = ingredientRowSectionKey(row);
  return key.includes("BAKERY") || key.includes("BREAD") || key.includes("FLOUR") || key.includes("DOUGH") || key.includes("CROISSANT") || key.includes("PASTRY");
};

const ingredientLooksLikeLiquor = (row) => {
  const key = ingredientRowSectionKey(row);
  return key.includes("LIQUOR") || key.includes("LIQUER") || key.includes("WINE") || key.includes("BEER") || key.includes("VODKA") || key.includes("RUM") || key.includes("GIN") || key.includes("TEQUILA") || key.includes("WHISKEY");
};

const sectionMatchesIngredientRow = (sectionTitle, ingredientRow) => {
  const type = getTemplateSectionType(sectionTitle);

  if (type === "produce") return ingredientLooksLikeProduce(ingredientRow);
  if (type === "fish") return ingredientLooksLikeFish(ingredientRow);
  if (type === "meat") return ingredientLooksLikeMeat(ingredientRow);
  if (type === "bakery" || type === "pastry") return ingredientLooksLikeBakery(ingredientRow);
  if (type === "liquor") return ingredientLooksLikeLiquor(ingredientRow);
  if (type === "chargeable") return cleanKey(ingredientRow?.category).includes("CHARG") || cleanKey(ingredientRow?.subCategory).includes("CHARG");

  if (type === "menuBasic") {
    return !ingredientLooksLikeProduce(ingredientRow) &&
      !ingredientLooksLikeFish(ingredientRow) &&
      !ingredientLooksLikeMeat(ingredientRow) &&
      !ingredientLooksLikeBakery(ingredientRow) &&
      !ingredientLooksLikeLiquor(ingredientRow);
  }

  return true;
};

const detectFmlSheetName = (sheetNames = []) =>
  sheetNames.find((sheetName) => cleanKey(sheetName) === cleanKey(FML_SHEET_NAME)) ||
  sheetNames.find((sheetName) => cleanKey(sheetName).includes("FML")) ||
  "";

const detectUomSheetName = (sheetNames = []) =>
  sheetNames.find((sheetName) => cleanKey(sheetName) === cleanKey(UOM_SHEET_NAME)) ||
  sheetNames.find((sheetName) => cleanKey(sheetName).includes("UNIT OF MEASURE")) ||
  "";

const isVenueSheet = (sheetName, workbook) => {
  const key = cleanKey(sheetName);
  const worksheet = workbook?.Sheets?.[sheetName];

  if (!worksheet || !worksheet["!ref"]) return false;
  if (key.includes("FML")) return false;
  if (key.includes("UNIT OF MEASURE")) return false;
  if (key === "SHEET5") return false;
  if (key.includes("SUMMARY")) return false;
  if (key.includes("INDEX")) return false;

  return true;
};

const parseIngredientByLocationWorkbook = ({ workbook, XLSX }) => {
  const sheetName = workbook?.SheetNames?.[0] || "";
  const worksheet = workbook?.Sheets?.[sheetName];

  if (!worksheet || !XLSX) {
    return { rows: [], sheetName, headerRowIndex: -1 };
  }

  const rows = getWorksheetRows({ worksheet, XLSX, raw: false });
  let headerRowIndex = -1;

  rows.slice(0, 20).some((row, index) => {
    const headerMap = getHeaderMap(row || []);
    const hasRestaurant = getColumnIndex(headerMap, ["RestaurantName", "Restaurant Name", "Venue", "Location"], -1) >= 0;
    const hasCode = getColumnIndex(headerMap, ["Code", "Product Code", "Ingredient Code"], -1) >= 0;
    const hasName = getColumnIndex(headerMap, ["Name", "Product Name", "Ingredient Name"], -1) >= 0;

    if (hasRestaurant && hasCode && hasName) {
      headerRowIndex = index;
      return true;
    }

    return false;
  });

  if (headerRowIndex < 0) headerRowIndex = 0;

  const headerMap = getHeaderMap(rows[headerRowIndex] || []);
  const restaurantCodeIndex = getColumnIndex(headerMap, ["RestaurantCode", "Restaurant Code"], 0);
  const restaurantNameIndex = getColumnIndex(headerMap, ["RestaurantName", "Restaurant Name", "Venue", "Location"], 1);
  const menuCodeIndex = getColumnIndex(headerMap, ["MenuCode", "Menu Code"], 2);
  const menuNameIndex = getColumnIndex(headerMap, ["MenuName", "Menu Name"], 3);
  const categoryIndex = getColumnIndex(headerMap, ["Category"], 4);
  const subCategoryIndex = getColumnIndex(headerMap, ["SubCategory", "Sub Category"], 5);
  const codeIndex = getColumnIndex(headerMap, ["Code", "Product Code", "Ingredient Code"], 6);
  const nameIndex = getColumnIndex(headerMap, ["Name", "Product Name", "Ingredient Name"], 7);
  const activeIndex = getColumnIndex(headerMap, ["Active"], 11);
  const assignedIndex = getColumnIndex(headerMap, ["Assigned"], 12);
  const assignedTypeIndex = getColumnIndex(headerMap, ["AssignedType", "Assigned Type"], 13);
  const assignedActiveIndex = getColumnIndex(headerMap, ["AssignedActive", "Assigned Active"], 14);
  const recipeCodeIndex = getColumnIndex(headerMap, ["RecipeCode", "Recipe Code"], 15);
  const recipeNameIndex = getColumnIndex(headerMap, ["RecipeName", "Recipe Name"], 16);
  const specialInstructionsIndex = getColumnIndex(headerMap, ["SpecialInstructions", "Special Instructions"], 17);
  const specialInstructions2Index = getColumnIndex(headerMap, ["SpecialInstructions2", "Special Instructions 2"], 18);

  const parsedRows = [];
  const seen = new Set();

  rows.slice(headerRowIndex + 1).forEach((row, offset) => {
    const sourceRow = headerRowIndex + 2 + offset;
    const code = normalizeProductCode(row[codeIndex]);
    const restaurantName = cleanText(row[restaurantNameIndex]);
    const productName = cleanText(row[assignedIndex]) || cleanText(row[nameIndex]);
    const name = cleanText(row[nameIndex]);
    const assignedType = cleanKey(row[assignedTypeIndex]);

    if (!code || !restaurantName || !productName) return;
    if (!isActiveFlag(row[activeIndex]) || !isActiveFlag(row[assignedActiveIndex])) return;
    if (assignedType && !["P", "PRODUCT", "I", "INGREDIENT"].includes(assignedType)) return;

    const item = {
      key: `${sourceRow}-${code}`,
      sourceRow,
      restaurantCode: cleanText(row[restaurantCodeIndex]),
      restaurantName,
      menuCode: cleanText(row[menuCodeIndex]),
      menuName: cleanText(row[menuNameIndex]),
      category: cleanText(row[categoryIndex]),
      subCategory: cleanText(row[subCategoryIndex]),
      code,
      name,
      productName,
      assigned: cleanText(row[assignedIndex]),
      assignedType: cleanText(row[assignedTypeIndex]),
      recipeCode: cleanText(row[recipeCodeIndex]),
      recipeName: cleanText(row[recipeNameIndex]),
      specialInstructions: cleanText(row[specialInstructionsIndex]),
      specialInstructions2: cleanText(row[specialInstructions2Index]),
    };

    const rowKey = [
      item.restaurantName,
      item.menuName,
      item.category,
      item.subCategory,
      item.code,
      item.recipeCode,
      item.recipeName,
    ].map(cleanKey).join("|");

    if (seen.has(rowKey)) return;
    seen.add(rowKey);
    parsedRows.push(item);
  });

  return {
    rows: parsedRows,
    sheetName,
    headerRowIndex,
  };
};

const getTitleColumnIndexes = (rows) => {
  const firstRow = rows[0] || [];
  const indexes = [];

  firstRow.forEach((cell, columnIndex) => {
    const text = cleanText(cell);
    const key = cleanKey(text);

    if (!text) return;

    const looksLikeLocator =
      text.includes("-") ||
      /\b\d{1,2}\/\d{1,2}\/\d{1,4}\b/.test(text) ||
      key.includes("MENU BASIC") ||
      key.includes("PREPERATION") ||
      key.includes("PREPARATION") ||
      key.includes("VEGET") ||
      key.includes("FRUIT") ||
      key.includes("MEAT") ||
      key.includes("FISH") ||
      key.includes("SEAFOOD") ||
      key.includes("PASTRY") ||
      key.includes("BAKERY") ||
      key.includes("LIQUOR") ||
      key.includes("LIQUER");

    if (looksLikeLocator) {
      indexes.push(columnIndex);
    }
  });

  return indexes;
};

const findCodeColumnInRange = (rows, startColumn, endColumn) => {
  for (let columnIndex = startColumn; columnIndex <= endColumn; columnIndex += 1) {
    const headerText = cleanKey(rows[1]?.[columnIndex]);
    if (headerText === "CODE" || headerText.includes("APOLLO") || headerText.includes("VV CODE")) {
      return { codeCol: columnIndex, nameCol: columnIndex + 1, umCol: columnIndex + 2, headerRowIndex: 1 };
    }
  }

  let bestColumn = -1;
  let bestCount = 0;

  for (let columnIndex = startColumn; columnIndex <= endColumn; columnIndex += 1) {
    let count = 0;

    for (let rowIndex = 2; rowIndex < Math.min(rows.length, 250); rowIndex += 1) {
      if (isLikelyProductCode(rows[rowIndex]?.[columnIndex])) {
        count += 1;
      }
    }

    if (count > bestCount) {
      bestCount = count;
      bestColumn = columnIndex;
    }
  }

  if (bestColumn >= 0 && bestCount > 0) {
    return { codeCol: bestColumn, nameCol: bestColumn + 1, umCol: bestColumn + 2, headerRowIndex: 1 };
  }

  return null;
};

const getLastDataRowForBlock = ({ rows, block }) => {
  let lastRowIndex = block.headerRowIndex;

  for (let rowIndex = block.headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const code = normalizeProductCode(rows[rowIndex]?.[block.codeCol]);
    const name = cleanText(rows[rowIndex]?.[block.nameCol]);

    if (code || name) {
      lastRowIndex = rowIndex;
    }
  }

  return lastRowIndex;
};

const findTemplateBlocks = ({ worksheet, XLSX }) => {
  if (!worksheet || !XLSX) return [];

  const rows = getWorksheetRows({ worksheet, XLSX, raw: false });
  const decodedRange = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");
  const titleColumns = getTitleColumnIndexes(rows);

  return titleColumns
    .map((titleColumn, index) => {
      const nextTitleColumn = titleColumns[index + 1];
      const endColumn = nextTitleColumn !== undefined
        ? nextTitleColumn - 1
        : Math.min(decodedRange.e.c, titleColumn + 4);

      const columnInfo = findCodeColumnInRange(rows, titleColumn, endColumn);

      if (!columnInfo) return null;

      const block = {
        key: `${titleColumn}-${cleanKey(rows[0]?.[titleColumn])}`,
        title: cleanText(rows[0]?.[titleColumn]),
        titleColumn,
        startColumn: titleColumn,
        endColumn,
        ...columnInfo,
      };

      block.lastDataRowIndex = getLastDataRowForBlock({ rows, block });
      return block;
    })
    .filter(Boolean);
};

const getExistingCodesForBlock = ({ rows, block }) => {
  const codes = new Set();

  for (let rowIndex = block.headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const code = normalizeProductCode(rows[rowIndex]?.[block.codeCol]);
    if (code) codes.add(code);
  }

  return codes;
};

const getRequiredRowsForVenueSection = ({ ingredientRows, sheetName, sectionTitle }) => {
  const seenCodes = new Set();

  return (ingredientRows || [])
    .filter((row) => ingredientRowMatchesVenueSheet(row, sheetName))
    .filter((row) => sectionMatchesIngredientRow(sectionTitle, row))
    .filter((row) => {
      const code = normalizeProductCode(row.code);
      if (!code || seenCodes.has(code)) return false;
      seenCodes.add(code);
      return true;
    })
    .sort((left, right) => {
      const category = String(left.category || "").localeCompare(String(right.category || ""));
      if (category !== 0) return category;

      const subCategory = String(left.subCategory || "").localeCompare(String(right.subCategory || ""));
      if (subCategory !== 0) return subCategory;

      return String(left.productName || left.name || "").localeCompare(String(right.productName || right.name || ""));
    });
};

const buildAnalysisRows = ({ workbook, XLSX, ingredientRows, scope = ALL_VENUES_SCOPE }) => {
  if (!workbook || !XLSX || !ingredientRows?.length) return [];

  const venueSheets = (workbook.SheetNames || [])
    .filter((sheetName) => isVenueSheet(sheetName, workbook))
    .filter((sheetName) => scope === ALL_VENUES_SCOPE || sheetName === scope);

  const rows = [];

  venueSheets.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const worksheetRows = getWorksheetRows({ worksheet, XLSX, raw: false });
    const blocks = findTemplateBlocks({ worksheet, XLSX });

    blocks.forEach((block) => {
      const existingCodes = getExistingCodesForBlock({ rows: worksheetRows, block });
      const requiredRows = getRequiredRowsForVenueSection({
        ingredientRows,
        sheetName,
        sectionTitle: block.title,
      });

      requiredRows.forEach((ingredientRow) => {
        const code = normalizeProductCode(ingredientRow.code);
        const exists = existingCodes.has(code);

        rows.push({
          key: `${sheetName}|${block.key}|${code}|${ingredientRow.sourceRow}`,
          status: exists ? "Already in template" : "Suggested addition",
          isNew: !exists,
          venueSheet: sheetName,
          sectionTitle: block.title,
          code,
          productName: ingredientRow.productName || ingredientRow.name,
          restaurantName: ingredientRow.restaurantName,
          menuName: ingredientRow.menuName,
          category: ingredientRow.category,
          subCategory: ingredientRow.subCategory,
          sourceRow: ingredientRow.sourceRow,
        });
      });
    });
  });

  return rows;
};

const makeCodeCell = (code, sourceStyle) => {
  const numericValue = Number(code);
  return {
    t: Number.isFinite(numericValue) ? "n" : "s",
    v: Number.isFinite(numericValue) ? numericValue : String(code || ""),
    s: cloneStyle(sourceStyle),
  };
};

const makeFormulaCell = ({ formula, fallbackValue, sourceStyle }) => ({
  t: "s",
  f: formula,
  v: fallbackValue || "",
  s: cloneStyle(sourceStyle),
});

const writeSuggestedItemToBlock = ({
  workbook,
  worksheet,
  XLSX,
  block,
  rowIndex,
  ingredientRow,
  fmlSheetName,
  uomSheetName,
}) => {
  const code = normalizeProductCode(ingredientRow.code);
  const codeAddress = excelAddress(XLSX, rowIndex, block.codeCol);
  const nameAddress = excelAddress(XLSX, rowIndex, block.nameCol);
  const umAddress = block.umCol >= 0 ? excelAddress(XLSX, rowIndex, block.umCol) : "";

  const sourceCodeAddress = excelAddress(XLSX, Math.max(block.headerRowIndex + 1, rowIndex - 1), block.codeCol);
  const sourceNameAddress = excelAddress(XLSX, Math.max(block.headerRowIndex + 1, rowIndex - 1), block.nameCol);
  const sourceUmAddress = block.umCol >= 0 ? excelAddress(XLSX, Math.max(block.headerRowIndex + 1, rowIndex - 1), block.umCol) : "";

  const codeCell = makeCodeCell(code, worksheet[sourceCodeAddress]?.s);
  const nameFormula = fmlSheetName
    ? `VLOOKUP(${codeAddress},'${fmlSheetName.replace(/'/g, "''")}'!E:F,2,FALSE)`
    : "";
  const umFormula = uomSheetName
    ? `VLOOKUP(${codeAddress},'${uomSheetName.replace(/'/g, "''")}'!A:C,3,FALSE)`
    : "";

  const nameCell = nameFormula
    ? makeFormulaCell({
        formula: nameFormula,
        fallbackValue: ingredientRow.productName || ingredientRow.name,
        sourceStyle: worksheet[sourceNameAddress]?.s,
      })
    : {
        t: "s",
        v: ingredientRow.productName || ingredientRow.name || "",
        s: cloneStyle(worksheet[sourceNameAddress]?.s),
      };

  const umCell = umFormula
    ? makeFormulaCell({
        formula: umFormula,
        fallbackValue: ingredientRow.um || "",
        sourceStyle: sourceUmAddress ? worksheet[sourceUmAddress]?.s : undefined,
      })
    : {
        t: "s",
        v: ingredientRow.um || "",
        s: cloneStyle(sourceUmAddress ? worksheet[sourceUmAddress]?.s : undefined),
      };

  setCell({ worksheet, address: codeAddress, cell: codeCell, red: true });
  setCell({ worksheet, address: nameAddress, cell: nameCell, red: true });

  if (umAddress) {
    setCell({ worksheet, address: umAddress, cell: umCell, red: true });
  }

  ensureWorksheetRange({ worksheet, XLSX, rowIndex, columnIndex: block.codeCol });
  ensureWorksheetRange({ worksheet, XLSX, rowIndex, columnIndex: block.nameCol });
  if (block.umCol >= 0) ensureWorksheetRange({ worksheet, XLSX, rowIndex, columnIndex: block.umCol });

  // Preserve row height when possible.
  if (worksheet["!rows"]?.[rowIndex - 1] && !worksheet["!rows"]?.[rowIndex]) {
    worksheet["!rows"][rowIndex] = cloneStyle(worksheet["!rows"][rowIndex - 1]);
  }
};

const applyIngredientLocationToTemplate = async ({ workbook, XLSX, ingredientRows, scope = ALL_VENUES_SCOPE, onProgress }) => {
  const fmlSheetName = detectFmlSheetName(workbook.SheetNames || []);
  const uomSheetName = detectUomSheetName(workbook.SheetNames || []);
  const venueSheets = (workbook.SheetNames || [])
    .filter((sheetName) => isVenueSheet(sheetName, workbook))
    .filter((sheetName) => scope === ALL_VENUES_SCOPE || sheetName === scope);

  const summaryRows = [];
  let existingCount = 0;
  let suggestedCount = 0;

  for (let sheetIndex = 0; sheetIndex < venueSheets.length; sheetIndex += 1) {
    const sheetName = venueSheets[sheetIndex];
    const worksheet = workbook.Sheets[sheetName];
    const rows = getWorksheetRows({ worksheet, XLSX, raw: false });
    const blocks = findTemplateBlocks({ worksheet, XLSX });

    onProgress?.(`Building ${sheetIndex + 1} of ${venueSheets.length}: ${sheetName}`);

    blocks.forEach((block) => {
      const existingCodes = getExistingCodesForBlock({ rows, block });
      const requiredRows = getRequiredRowsForVenueSection({
        ingredientRows,
        sheetName,
        sectionTitle: block.title,
      });
      const usedCodes = new Set(existingCodes);
      let nextRowIndex = Math.max(block.lastDataRowIndex + 1, block.headerRowIndex + 2);

      requiredRows.forEach((ingredientRow) => {
        const code = normalizeProductCode(ingredientRow.code);
        if (!code) return;

        if (usedCodes.has(code)) {
          existingCount += 1;
          summaryRows.push({
            Status: "Already in template",
            Venue: sheetName,
            Section: block.title,
            Code: code,
            Product: ingredientRow.productName || ingredientRow.name,
            Restaurant: ingredientRow.restaurantName,
            Menu: ingredientRow.menuName,
            Category: ingredientRow.category,
            SubCategory: ingredientRow.subCategory,
            IngredientSourceRow: ingredientRow.sourceRow,
          });
          return;
        }

        writeSuggestedItemToBlock({
          workbook,
          worksheet,
          XLSX,
          block,
          rowIndex: nextRowIndex,
          ingredientRow,
          fmlSheetName,
          uomSheetName,
        });

        suggestedCount += 1;
        usedCodes.add(code);

        summaryRows.push({
          Status: "Suggested addition - red in template",
          Venue: sheetName,
          Section: block.title,
          Code: code,
          Product: ingredientRow.productName || ingredientRow.name,
          Restaurant: ingredientRow.restaurantName,
          Menu: ingredientRow.menuName,
          Category: ingredientRow.category,
          SubCategory: ingredientRow.subCategory,
          IngredientSourceRow: ingredientRow.sourceRow,
        });

        nextRowIndex += 1;
      });
    });

    if (sheetIndex % 2 === 1) {
      await yieldToBrowser();
    }
  }

  const summaryWorksheet = XLSX.utils.json_to_sheet(summaryRows);
  summaryWorksheet["!cols"] = [
    { wch: 28 },
    { wch: 28 },
    { wch: 54 },
    { wch: 14 },
    { wch: 46 },
    { wch: 28 },
    { wch: 28 },
    { wch: 34 },
    { wch: 34 },
    { wch: 18 },
  ];

  const summarySheetName = "Suggested Additions";
  const existingSheetIndex = workbook.SheetNames.indexOf(summarySheetName);

  if (existingSheetIndex >= 0) {
    workbook.Sheets[summarySheetName] = summaryWorksheet;
  } else {
    workbook.SheetNames.push(summarySheetName);
    workbook.Sheets[summarySheetName] = summaryWorksheet;
  }

  return {
    existingCount,
    suggestedCount,
    summaryRows,
    venueSheetsProcessed: venueSheets.length,
  };
};

const summarizeAnalysisRows = (rows = []) => ({
  total: rows.length,
  existing: rows.filter((row) => !row.isNew).length,
  suggested: rows.filter((row) => row.isNew).length,
  venues: new Set(rows.map((row) => row.venueSheet).filter(Boolean)).size,
  sections: new Set(rows.map((row) => `${row.venueSheet}|${row.sectionTitle}`).filter(Boolean)).size,
});

const cardFallbackStyle = {
  background: "#fff",
  borderRadius: 16,
  padding: 20,
  boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
};

const primaryButtonFallbackStyle = {
  padding: "11px 14px",
  borderRadius: 10,
  border: 0,
  background: "#111",
  color: "#fff",
  fontWeight: "bold",
  cursor: "pointer",
};

const secondaryButtonFallbackStyle = {
  padding: "11px 14px",
  borderRadius: 10,
  border: "1px solid #bbb",
  background: "#fff",
  color: "#111",
  fontWeight: "bold",
  cursor: "pointer",
};

const tableHeaderStyle = {
  padding: 9,
  textAlign: "left",
  background: "#111",
  color: "#fff",
  borderRight: "1px solid #333",
  whiteSpace: "nowrap",
};

const tableCellStyle = {
  padding: 9,
  verticalAlign: "top",
  borderRight: "1px solid #eee",
  borderBottom: "1px solid #eee",
};

export default function ERPVenueIngredientsScreen({
  styles = {},
  setModule,
  logUsageEvent,
}) {
  const [ingredientRows, setIngredientRows] = useState([]);
  const [ingredientFileName, setIngredientFileName] = useState("");
  const [ingredientMessage, setIngredientMessage] = useState("Load or upload Ingredient by Location first.");
  const [templateArrayBuffer, setTemplateArrayBuffer] = useState(null);
  const [templateWorkbook, setTemplateWorkbook] = useState(null);
  const [templateFileName, setTemplateFileName] = useState("");
  const [xlsxApi, setXlsxApi] = useState(null);
  const [templateMessage, setTemplateMessage] = useState("Upload the ERP template after Ingredient by Location is loaded.");
  const [selectedScope, setSelectedScope] = useState(ALL_VENUES_SCOPE);
  const [analysisRows, setAnalysisRows] = useState([]);
  const [busy, setBusy] = useState(false);
  const [busyMessage, setBusyMessage] = useState("");

  const cardStyle = styles.card || cardFallbackStyle;
  const primaryButtonStyle = styles.primaryButton || primaryButtonFallbackStyle;
  const secondaryButtonStyle = styles.secondaryButton || secondaryButtonFallbackStyle;

  const venueSheets = useMemo(
    () =>
      templateWorkbook && xlsxApi
        ? (templateWorkbook.SheetNames || []).filter((sheetName) =>
            isVenueSheet(sheetName, templateWorkbook)
          )
        : [],
    [templateWorkbook, xlsxApi]
  );

  const analysisSummary = useMemo(() => summarizeAnalysisRows(analysisRows), [analysisRows]);
  const visibleAnalysisRows = useMemo(() => analysisRows.slice(0, MAX_PREVIEW_ROWS), [analysisRows]);

  const loadIngredientWorkbookFromArrayBuffer = async ({ arrayBuffer, fileName }) => {
    setBusy(true);
    setBusyMessage("Reading Ingredient by Location...");

    try {
      const XLSX = xlsxApi || await loadXlsx();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });
      const parsed = parseIngredientByLocationWorkbook({ workbook, XLSX });

      setXlsxApi(XLSX);
      setIngredientRows(parsed.rows || []);
      setIngredientFileName(fileName || "Ingredient by Location");
      setIngredientMessage(
        parsed.rows?.length
          ? `Ingredient by Location loaded. ${parsed.rows.length} active ingredient-location row(s) found from ${parsed.sheetName || "the first sheet"}.`
          : "Ingredient by Location loaded, but no usable rows were found. Check the headers."
      );
      setAnalysisRows([]);

      logUsageEvent?.("erp_simple_ingredient_location_loaded", {
        module: "erp_location_ingredients",
        fileName: fileName || "Ingredient by Location",
        rows: parsed.rows?.length || 0,
        sheetName: parsed.sheetName || "",
      });
    } catch (error) {
      setIngredientRows([]);
      setIngredientMessage(error?.message || "Could not read Ingredient by Location.");
      window.alert(error?.message || "Could not read Ingredient by Location.");
    } finally {
      setBusy(false);
      setBusyMessage("");
    }
  };

  const loadTemplateWorkbookFromArrayBuffer = async ({ arrayBuffer, fileName }) => {
    setBusy(true);
    setBusyMessage("Reading ERP template...");

    try {
      const XLSX = xlsxApi || await loadXlsx();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
        cellStyles: true,
        bookVBA: true,
      });
      const detectedVenueSheets = (workbook.SheetNames || []).filter((sheetName) =>
        isVenueSheet(sheetName, workbook)
      );

      setXlsxApi(XLSX);
      setTemplateArrayBuffer(arrayBuffer);
      setTemplateWorkbook(workbook);
      setTemplateFileName(fileName || "ERP template");
      setTemplateMessage(
        detectedVenueSheets.length
          ? `ERP template loaded. ${detectedVenueSheets.length} venue/location tab(s) found. Existing items will stay black; suggested new items will be added in red.`
          : "ERP template loaded, but no venue/location tabs were detected."
      );
      setSelectedScope((current) =>
        current === ALL_VENUES_SCOPE || detectedVenueSheets.includes(current)
          ? current
          : ALL_VENUES_SCOPE
      );
      setAnalysisRows([]);

      logUsageEvent?.("erp_simple_template_loaded", {
        module: "erp_location_ingredients",
        fileName: fileName || "ERP template",
        venueTabs: detectedVenueSheets.length,
      });
    } catch (error) {
      setTemplateArrayBuffer(null);
      setTemplateWorkbook(null);
      setTemplateFileName("");
      setTemplateMessage(error?.message || "Could not read ERP template.");
      window.alert(error?.message || "Could not read ERP template.");
    } finally {
      setBusy(false);
      setBusyMessage("");
    }
  };

  const loadDefaultIngredientFile = async () => {
    setBusy(true);
    setBusyMessage("Loading permanent Ingredient by Location...");

    try {
      const response = await fetch(DEFAULT_INGREDIENT_BY_LOCATION_PATH, { cache: "no-store" });

      if (!response.ok) {
        throw new Error("Permanent Ingredient by Location file was not found. Add public/ingredient-by-location.xlsx or upload the file here.");
      }

      const arrayBuffer = await response.arrayBuffer();
      await loadIngredientWorkbookFromArrayBuffer({
        arrayBuffer,
        fileName: "ingredient-by-location.xlsx",
      });
    } catch (error) {
      setIngredientMessage(error?.message || "Could not load permanent Ingredient by Location file.");
      window.alert(error?.message || "Could not load permanent Ingredient by Location file.");
    } finally {
      setBusy(false);
      setBusyMessage("");
    }
  };

  const loadDefaultTemplateFile = async () => {
    setBusy(true);
    setBusyMessage("Loading permanent ERP template...");

    try {
      const response = await fetch(DEFAULT_ERP_TEMPLATE_PATH, { cache: "no-store" });

      if (!response.ok) {
        throw new Error("Permanent ERP template file was not found. Add public/erp-template-locations.xlsx or upload the file here.");
      }

      const arrayBuffer = await response.arrayBuffer();
      await loadTemplateWorkbookFromArrayBuffer({
        arrayBuffer,
        fileName: "erp-template-locations.xlsx",
      });
    } catch (error) {
      setTemplateMessage(error?.message || "Could not load permanent ERP template file.");
      window.alert(error?.message || "Could not load permanent ERP template file.");
    } finally {
      setBusy(false);
      setBusyMessage("");
    }
  };

  const handleIngredientUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      await loadIngredientWorkbookFromArrayBuffer({ arrayBuffer, fileName: file.name });
    } finally {
      event.target.value = "";
    }
  };

  const handleTemplateUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      await loadTemplateWorkbookFromArrayBuffer({ arrayBuffer, fileName: file.name });
    } finally {
      event.target.value = "";
    }
  };

  const analyzeTemplate = async () => {
    if (!templateWorkbook || !xlsxApi) {
      window.alert("Upload or load the ERP template first.");
      return;
    }

    if (!ingredientRows.length) {
      window.alert("Load Ingredient by Location first.");
      return;
    }

    setBusy(true);
    setBusyMessage("Preparing simple comparison...");

    try {
      await yieldToBrowser();
      const rows = buildAnalysisRows({
        workbook: templateWorkbook,
        XLSX: xlsxApi,
        ingredientRows,
        scope: selectedScope,
      });

      setAnalysisRows(rows);
      setBusyMessage("");
      logUsageEvent?.("erp_simple_template_analyzed", {
        module: "erp_location_ingredients",
        scope: selectedScope,
        rows: rows.length,
        suggested: rows.filter((row) => row.isNew).length,
      });
    } catch (error) {
      setAnalysisRows([]);
      window.alert(error?.message || "Could not analyze the template.");
    } finally {
      setBusy(false);
      setBusyMessage("");
    }
  };

  const downloadCorrectedTemplate = async () => {
    if (!templateArrayBuffer) {
      window.alert("Upload or load the ERP template first.");
      return;
    }

    if (!ingredientRows.length) {
      window.alert("Load Ingredient by Location first.");
      return;
    }

    setBusy(true);
    setBusyMessage("Creating corrected ERP template...");

    try {
      const XLSX = xlsxApi || await loadXlsx();
      const workbook = XLSX.read(templateArrayBuffer.slice(0), {
        type: "array",
        cellDates: true,
        cellStyles: true,
        bookVBA: true,
      });

      const result = await applyIngredientLocationToTemplate({
        workbook,
        XLSX,
        ingredientRows,
        scope: selectedScope,
        onProgress: setBusyMessage,
      });

      const scopeName = selectedScope === ALL_VENUES_SCOPE ? "all-venues" : selectedScope;
      const outputName = `erp-template-${makeSafeFilePart(scopeName)}-ingredient-location-${getDateStamp()}.xlsx`;

      XLSX.writeFile(workbook, outputName, {
        bookType: "xlsx",
      });

      setBusyMessage("");
      setTemplateMessage(
        `Downloaded corrected ERP template. Existing placements stayed black. Suggested additions in red: ${result.suggestedCount}. Existing matched items: ${result.existingCount}.`
      );
      setAnalysisRows(result.summaryRows.map((row, index) => ({
        key: `download-${index}-${row.Code}`,
        status: row.Status,
        isNew: String(row.Status || "").toLowerCase().includes("suggested"),
        venueSheet: row.Venue,
        sectionTitle: row.Section,
        code: row.Code,
        productName: row.Product,
        restaurantName: row.Restaurant,
        menuName: row.Menu,
        category: row.Category,
        subCategory: row.SubCategory,
        sourceRow: row.IngredientSourceRow,
      })));

      logUsageEvent?.("erp_simple_corrected_template_downloaded", {
        module: "erp_location_ingredients",
        scope: selectedScope,
        outputName,
        existingCount: result.existingCount,
        suggestedCount: result.suggestedCount,
      });
    } catch (error) {
      window.alert(error?.message || "Could not create the corrected ERP template.");
    } finally {
      setBusy(false);
      setBusyMessage("");
    }
  };

  return (
    <div style={styles.page || { minHeight: "100vh", padding: 24, background: "#f5f5f5", fontFamily: "Arial, sans-serif" }}>
      <div style={styles.header || { display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", background: "#fff", borderRadius: 16, padding: 18, marginBottom: 16, flexWrap: "wrap" }}>
        <div>
          <h1 style={{ margin: 0 }}>ERP Template Builder</h1>
          <p style={styles.subtitle || { margin: "4px 0 0", color: "#666" }}>
            Simple flow: Ingredient by Location builds the required structure. ERP template receives missing items in red.
          </p>
        </div>

        <button type="button" style={styles.backButton || secondaryButtonStyle} onClick={() => setModule?.("")}>
          ← Back
        </button>
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 14 }}>
        <h2 style={{ margin: 0 }}>1. Load files</h2>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))", gap: 14 }}>
          <div style={{ display: "grid", gap: 10 }}>
            <strong>Ingredient by Location</strong>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button type="button" style={secondaryButtonStyle} onClick={loadDefaultIngredientFile} disabled={busy}>
                Load permanent file
              </button>
              <label style={{ ...secondaryButtonStyle, display: "inline-block" }}>
                Upload file
                <input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleIngredientUpload} disabled={busy} style={{ display: "none" }} />
              </label>
            </div>
            <span style={{ color: "#666", fontSize: 13 }}>
              Current: {ingredientFileName || "No file loaded"}
            </span>
            <div style={styles.infoBox || { padding: 12, borderRadius: 12, background: "#f7f7f7" }}>
              {ingredientMessage}
            </div>
          </div>

          <div style={{ display: "grid", gap: 10 }}>
            <strong>ERP template</strong>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button type="button" style={secondaryButtonStyle} onClick={loadDefaultTemplateFile} disabled={busy}>
                Load permanent template
              </button>
              <label style={{ ...secondaryButtonStyle, display: "inline-block" }}>
                Upload template
                <input type="file" accept=".xlsx,.xls,.xlsm" onChange={handleTemplateUpload} disabled={busy} style={{ display: "none" }} />
              </label>
            </div>
            <span style={{ color: "#666", fontSize: 13 }}>
              Current: {templateFileName || "No template loaded"}
            </span>
            <div style={styles.infoBox || { padding: 12, borderRadius: 12, background: "#f7f7f7" }}>
              {templateMessage}
            </div>
          </div>
        </div>

        {busy || busyMessage ? (
          <div style={styles.infoBox || { padding: 12, borderRadius: 12, background: "#fff4d6" }}>
            {busyMessage || "Working..."}
          </div>
        ) : null}
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 14 }}>
        <h2 style={{ margin: 0 }}>2. Choose report scope</h2>

        <div style={{ display: "grid", gridTemplateColumns: "minmax(240px, 420px) auto auto", gap: 10, alignItems: "end" }}>
          <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
            Venue / location
            <select
              value={selectedScope}
              onChange={(event) => {
                setSelectedScope(event.target.value);
                setAnalysisRows([]);
              }}
              style={styles.selectInput || { width: "100%", padding: 11, borderRadius: 10, border: "1px solid #ccc", background: "#fff" }}
              disabled={busy || !venueSheets.length}
            >
              <option value={ALL_VENUES_SCOPE}>All venue tabs</option>
              {venueSheets.map((sheetName) => (
                <option key={sheetName} value={sheetName}>{sheetName}</option>
              ))}
            </select>
          </label>

          <button type="button" style={secondaryButtonStyle} onClick={analyzeTemplate} disabled={busy || !ingredientRows.length || !templateWorkbook}>
            Preview comparison
          </button>

          <button type="button" style={primaryButtonStyle} onClick={downloadCorrectedTemplate} disabled={busy || !ingredientRows.length || !templateArrayBuffer}>
            Download corrected ERP template
          </button>
        </div>

        <p style={styles.message || { color: "#555", fontSize: 14, margin: 0 }}>
          Existing ingredient codes already in the ERP template stay untouched in black. Missing ingredient codes from Ingredient by Location are inserted into the matching venue/row-1 section and colored red.
        </p>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 10, marginBottom: 16 }}>
        {[
          ["Ingredient rows", ingredientRows.length],
          ["Venue tabs", venueSheets.length],
          ["Compared rows", analysisSummary.total],
          ["Already black", analysisSummary.existing],
          ["Suggested red", analysisSummary.suggested],
          ["Sections", analysisSummary.sections],
        ].map(([label, value]) => (
          <div key={label} style={{ ...cardStyle, padding: 14, textAlign: "center" }}>
            <div style={{ color: "#666", fontSize: 12, fontWeight: "bold", textTransform: "uppercase" }}>{label}</div>
            <div style={{ fontSize: 24, fontWeight: 900, marginTop: 4 }}>{value}</div>
          </div>
        ))}
      </div>

      <div style={{ ...cardStyle, display: "grid", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0 }}>Preview</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            Showing {visibleAnalysisRows.length} of {analysisRows.length} compared placement row(s). Red rows are suggested additions.
          </p>
        </div>

        <div style={{ overflowX: "auto", border: "1px solid #ddd", borderRadius: 14, maxHeight: 520 }}>
          <table style={{ width: "100%", minWidth: 1200, borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr>
                <th style={tableHeaderStyle}>Status</th>
                <th style={tableHeaderStyle}>Venue Tab</th>
                <th style={tableHeaderStyle}>Row 1 Section</th>
                <th style={tableHeaderStyle}>Code</th>
                <th style={tableHeaderStyle}>Product</th>
                <th style={tableHeaderStyle}>Ingredient Location Venue</th>
                <th style={tableHeaderStyle}>Category</th>
                <th style={tableHeaderStyle}>SubCategory</th>
                <th style={tableHeaderStyle}>Source Row</th>
              </tr>
            </thead>
            <tbody>
              {visibleAnalysisRows.length ? (
                visibleAnalysisRows.map((row) => (
                  <tr key={row.key} style={{ color: row.isNew ? "#b00020" : "#111", fontWeight: row.isNew ? "bold" : "normal" }}>
                    <td style={tableCellStyle}>{row.status}</td>
                    <td style={tableCellStyle}>{row.venueSheet}</td>
                    <td style={tableCellStyle}>{row.sectionTitle}</td>
                    <td style={tableCellStyle}>{row.code}</td>
                    <td style={tableCellStyle}>{row.productName}</td>
                    <td style={tableCellStyle}>{row.restaurantName}</td>
                    <td style={tableCellStyle}>{row.category}</td>
                    <td style={tableCellStyle}>{row.subCategory}</td>
                    <td style={tableCellStyle}>{row.sourceRow}</td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={9} style={{ ...tableCellStyle, textAlign: "center", color: "#777", padding: 22 }}>
                    Load Ingredient by Location and ERP template, then click Preview comparison or Download corrected ERP template.
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

import * as XLSXModule from "xlsx";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 120;

const XLSX = XLSXModule.default || XLSXModule;

const FML_SHEET_NAME = "FML March 2026";
const UOM_SHEET_NAME = "Unit of Measure ship March 2026";
const ALL_VENUES_SCOPE = "ALL";
const MAX_PREVIEW_ROWS = 400;

const cleanText = (value) =>
  String(value ?? "")
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

const normalizeTemplateProductCode = (value) => {
  const rawText = String(value ?? "").replace(/\u00a0/g, " ").trim();

  if (!rawText) return "";

  if (/^[+-]?\d+(\.\d+)?e[+-]?\d+$/i.test(rawText) || /^\d+\.0+$/.test(rawText)) {
    const normalized = normalizeProductCode(rawText);
    return /^\d{3,}$/.test(normalized) ? normalized : "";
  }

  const digits = rawText.replace(/[^0-9]/g, "");
  if (!digits || digits.length < 3) return "";

  const letters = rawText.replace(/[^A-Za-z]/g, "");
  const mostlyText = letters.length > 2 && digits.length < rawText.length / 2;

  // Code cells must be real product codes. This prevents product names or pack sizes
  // such as "12/2.4 LB" from being treated as template item codes.
  if (mostlyText) return "";

  return normalizeProductCode(rawText);
};

const isLikelyProductCode = (value) => /^\d{3,}$/.test(normalizeTemplateProductCode(value));

const excelAddress = (rowIndex, columnIndex) =>
  XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });

const getWorksheetRows = ({ worksheet, raw = false }) => {
  if (!worksheet) return [];

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

const ensureWorksheetRange = ({ worksheet, rowIndex, columnIndex }) => {
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
  const aliases = new Set([sheetName, normalizeVenueForMatch(sheetName)]);

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
    row?.menuName,
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
  const sectionKey = cleanKey(sectionTitle);
  const ingredientKey = ingredientRowSectionKey(ingredientRow);

  if (sectionKey && ingredientKey) {
    const simplifiedSection = sectionKey
      .replace(/^\d+\s+/, "")
      .replace(/\b\d{1,2}\s+\d{1,2}\s+\d{1,4}\b/g, "")
      .replace(/\bMENU BASIC PREPERATION\b/g, "MENU BASIC PREPARATION")
      .trim();

    if (
      simplifiedSection.length >= 8 &&
      (ingredientKey.includes(simplifiedSection) || simplifiedSection.includes(ingredientKey))
    ) {
      return true;
    }
  }

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

  return false;
};

const detectFmlSheetName = (sheetNames = []) =>
  sheetNames.find((sheetName) => cleanKey(sheetName) === cleanKey(FML_SHEET_NAME)) ||
  sheetNames.find((sheetName) => cleanKey(sheetName).includes("FML")) ||
  "";

const detectUomSheetName = (sheetNames = []) =>
  sheetNames.find((sheetName) => cleanKey(sheetName) === cleanKey(UOM_SHEET_NAME)) ||
  sheetNames.find((sheetName) => cleanKey(sheetName).includes("UNIT OF MEASURE")) ||
  "";

const isVenueSheetName = (sheetName) => {
  const key = cleanKey(sheetName);

  if (!key) return false;
  if (key.includes("FML")) return false;
  if (key.includes("UNIT OF MEASURE")) return false;
  if (key.includes("SUGGESTED ADDITIONS")) return false;
  if (key === "SHEET5") return false;
  if (key.includes("SUMMARY")) return false;
  if (key.includes("INDEX")) return false;

  return true;
};

const isVenueSheet = (sheetName, workbook) => {
  const worksheet = workbook?.Sheets?.[sheetName];
  if (!worksheet || !worksheet["!ref"]) return false;
  return isVenueSheetName(sheetName);
};

const parseIngredientByLocationWorkbook = (workbook) => {
  const sheetName = workbook?.SheetNames?.[0] || "";
  const worksheet = workbook?.Sheets?.[sheetName];

  if (!worksheet) {
    return { rows: [], sheetName, headerRowIndex: -1 };
  }

  const rows = getWorksheetRows({ worksheet, raw: false });
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
    const code = normalizeTemplateProductCode(rows[rowIndex]?.[block.codeCol]);
    const name = cleanText(rows[rowIndex]?.[block.nameCol]);

    if (code || name) {
      lastRowIndex = rowIndex;
    }
  }

  return lastRowIndex;
};

const findTemplateBlocks = ({ worksheet }) => {
  if (!worksheet) return [];

  const rows = getWorksheetRows({ worksheet, raw: false });
  const decodedRange = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");
  const titleColumns = getTitleColumnIndexes(rows);

  return titleColumns
    .map((titleColumn, index) => {
      const nextTitleColumn = titleColumns[index + 1];

      // Important fix:
      // In this ERP template, the row-1 locator is usually over the name column,
      // while the real Code column is one column to the left.
      // Example Gunbae:
      // row 1 title columns: C / G / K / O
      // real code columns:   B / F / J / N
      const searchStartColumn = Math.max(0, titleColumn - 1);
      const searchEndColumn = nextTitleColumn !== undefined
        ? Math.max(searchStartColumn, nextTitleColumn - 2)
        : Math.min(decodedRange.e.c, titleColumn + 4);

      const columnInfo = findCodeColumnInRange(rows, searchStartColumn, searchEndColumn);

      if (!columnInfo) return null;

      const block = {
        key: `${titleColumn}-${cleanKey(rows[0]?.[titleColumn])}`,
        title: cleanText(rows[0]?.[titleColumn]),
        titleColumn,
        startColumn: searchStartColumn,
        endColumn: Math.max(searchEndColumn, columnInfo.umCol || searchEndColumn),
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
    const code = normalizeTemplateProductCode(rows[rowIndex]?.[block.codeCol]);
    if (code) codes.add(code);
  }

  return codes;
};

const getExistingCodeInfoForSheet = ({ rows, blocks }) => {
  const codeInfo = new Map();

  (blocks || []).forEach((block) => {
    for (let rowIndex = block.headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
      const code = normalizeTemplateProductCode(rows[rowIndex]?.[block.codeCol]);
      if (!code || codeInfo.has(code)) continue;

      codeInfo.set(code, {
        section: block.title,
        rowIndex,
        codeCol: block.codeCol,
      });
    }
  });

  return codeInfo;
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

const buildAnalysisRows = ({ workbook, ingredientRows, scope = ALL_VENUES_SCOPE }) => {
  if (!workbook || !ingredientRows?.length) return [];

  const venueSheets = (workbook.SheetNames || [])
    .filter((sheetName) => isVenueSheet(sheetName, workbook))
    .filter((sheetName) => scope === ALL_VENUES_SCOPE || sheetName === scope);

  const rows = [];

  venueSheets.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const worksheetRows = getWorksheetRows({ worksheet, raw: false });
    const blocks = findTemplateBlocks({ worksheet });
    const existingCodeInfoBySheet = getExistingCodeInfoForSheet({ rows: worksheetRows, blocks });
    const previewedCodesForSheet = new Set();

    blocks.forEach((block) => {
      const existingCodes = getExistingCodesForBlock({ rows: worksheetRows, block });
      const requiredRows = getRequiredRowsForVenueSection({
        ingredientRows,
        sheetName,
        sectionTitle: block.title,
      });

      requiredRows.forEach((ingredientRow) => {
        const code = normalizeProductCode(ingredientRow.code);
        if (!code || previewedCodesForSheet.has(code)) return;

        const existsInSection = existingCodes.has(code);
        const existingInfo = existingCodeInfoBySheet.get(code);
        const existsInVenue = Boolean(existingInfo);

        rows.push({
          Status: existsInSection
            ? "Already in template"
            : existsInVenue
              ? `Already in template - different section: ${existingInfo.section || "unknown"}`
              : "Suggested addition",
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

        previewedCodesForSheet.add(code);
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
  worksheet,
  block,
  rowIndex,
  ingredientRow,
  fmlSheetName,
  uomSheetName,
}) => {
  const code = normalizeProductCode(ingredientRow.code);
  const codeAddress = excelAddress(rowIndex, block.codeCol);
  const nameAddress = excelAddress(rowIndex, block.nameCol);
  const umAddress = block.umCol >= 0 ? excelAddress(rowIndex, block.umCol) : "";

  const previousRowIndex = Math.max(block.headerRowIndex + 1, rowIndex - 1);
  const sourceCodeAddress = excelAddress(previousRowIndex, block.codeCol);
  const sourceNameAddress = excelAddress(previousRowIndex, block.nameCol);
  const sourceUmAddress = block.umCol >= 0 ? excelAddress(previousRowIndex, block.umCol) : "";

  const safeFmlName = String(fmlSheetName || "").replace(/'/g, "''");
  const safeUomName = String(uomSheetName || "").replace(/'/g, "''");

  const codeCell = makeCodeCell(code, worksheet[sourceCodeAddress]?.s);

  const nameFormula = fmlSheetName
    ? `VLOOKUP(${codeAddress},'${safeFmlName}'!E:F,2,FALSE)`
    : "";

  const umFormula = uomSheetName
    ? `VLOOKUP(${codeAddress},'${safeUomName}'!A:C,3,FALSE)`
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

  // Extra marker so the row is still identifiable even if the local xlsx package
  // does not preserve generated font color.
  worksheet[codeAddress].c = [{ a: "ERP Template Builder", t: "Suggested addition from Ingredient by Location" }];

  ensureWorksheetRange({ worksheet, rowIndex, columnIndex: block.codeCol });
  ensureWorksheetRange({ worksheet, rowIndex, columnIndex: block.nameCol });
  if (block.umCol >= 0) ensureWorksheetRange({ worksheet, rowIndex, columnIndex: block.umCol });

  if (worksheet["!rows"]?.[rowIndex - 1] && !worksheet["!rows"]?.[rowIndex]) {
    worksheet["!rows"][rowIndex] = cloneStyle(worksheet["!rows"][rowIndex - 1]);
  }
};

const applyIngredientLocationToTemplate = ({ workbook, ingredientRows, scope = ALL_VENUES_SCOPE }) => {
  const fmlSheetName = detectFmlSheetName(workbook.SheetNames || []);
  const uomSheetName = detectUomSheetName(workbook.SheetNames || []);

  const venueSheets = (workbook.SheetNames || [])
    .filter((sheetName) => isVenueSheet(sheetName, workbook))
    .filter((sheetName) => scope === ALL_VENUES_SCOPE || sheetName === scope);

  const summaryRows = [];
  let existingCount = 0;
  let suggestedCount = 0;

  for (const sheetName of venueSheets) {
    const worksheet = workbook.Sheets[sheetName];
    const rows = getWorksheetRows({ worksheet, raw: false });
    const blocks = findTemplateBlocks({ worksheet });
    const existingCodeInfoBySheet = getExistingCodeInfoForSheet({ rows, blocks });

    // Prevent duplicate red additions anywhere on the same venue tab.
    const usedSheetCodes = new Set(existingCodeInfoBySheet.keys());

    blocks.forEach((block) => {
      const existingCodes = getExistingCodesForBlock({ rows, block });
      const requiredRows = getRequiredRowsForVenueSection({
        ingredientRows,
        sheetName,
        sectionTitle: block.title,
      });

      let nextRowIndex = Math.max(block.lastDataRowIndex + 1, block.headerRowIndex + 2);

      requiredRows.forEach((ingredientRow) => {
        const code = normalizeProductCode(ingredientRow.code);
        if (!code) return;

        const existsInSection = existingCodes.has(code);
        const existingInfo = existingCodeInfoBySheet.get(code);
        const existsInVenue = Boolean(existingInfo);

        if (existsInSection || existsInVenue || usedSheetCodes.has(code)) {
          existingCount += 1;

          summaryRows.push({
            Status: existsInSection
              ? "Already in template"
              : `Already in template - different section: ${existingInfo?.section || "unknown"}`,
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
          worksheet,
          block,
          rowIndex: nextRowIndex,
          ingredientRow,
          fmlSheetName,
          uomSheetName,
        });

        suggestedCount += 1;
        usedSheetCodes.add(code);

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
  existing: rows.filter((row) => !String(row.Status || "").toLowerCase().includes("suggested")).length,
  suggested: rows.filter((row) => String(row.Status || "").toLowerCase().includes("suggested")).length,
  venues: new Set(rows.map((row) => row.Venue).filter(Boolean)).size,
  sections: new Set(rows.map((row) => `${row.Venue}|${row.Section}`).filter(Boolean)).size,
});

const jsonResponse = (payload, status = 200) =>
  Response.json(payload, {
    status,
    headers: {
      "Cache-Control": "no-store",
    },
  });

const getRequiredFile = (formData, name) => {
  const value = formData.get(name);
  if (!value || typeof value.arrayBuffer !== "function") {
    throw new Error(`Missing file: ${name}`);
  }
  return value;
};

const getFileBuffer = async (file) => Buffer.from(await file.arrayBuffer());

const getVenueSheetsFromTemplateBuffer = (templateBuffer) => {
  const workbook = XLSX.read(templateBuffer, {
    type: "buffer",
    bookSheets: true,
  });

  return (workbook.SheetNames || []).filter(isVenueSheetName);
};

const readIngredientRowsFromBuffer = (ingredientBuffer) => {
  const ingredientWorkbook = XLSX.read(ingredientBuffer, {
    type: "buffer",
    cellDates: true,
  });

  return parseIngredientByLocationWorkbook(ingredientWorkbook);
};

const readFullTemplateWorkbookFromBuffer = (templateBuffer) =>
  XLSX.read(templateBuffer, {
    type: "buffer",
    cellDates: true,
    cellStyles: true,
    bookVBA: true,
  });

export async function POST(request) {
  try {
    const formData = await request.formData();
    const action = String(formData.get("action") || "download");
    const scope = String(formData.get("scope") || ALL_VENUES_SCOPE);

    if (action === "inspect") {
      const templateFile = getRequiredFile(formData, "template");
      const templateBuffer = await getFileBuffer(templateFile);
      const venueSheets = getVenueSheetsFromTemplateBuffer(templateBuffer);

      return jsonResponse({
        ok: true,
        venueSheets,
        venueCount: venueSheets.length,
      });
    }

    const ingredientFile = getRequiredFile(formData, "ingredient");
    const templateFile = getRequiredFile(formData, "template");
    const ingredientBuffer = await getFileBuffer(ingredientFile);
    const templateBuffer = await getFileBuffer(templateFile);

    const ingredientParsed = readIngredientRowsFromBuffer(ingredientBuffer);
    const ingredientRows = ingredientParsed.rows || [];

    if (!ingredientRows.length) {
      throw new Error("No active ingredient-location rows were found in Ingredient by Location.");
    }

    const workbook = readFullTemplateWorkbookFromBuffer(templateBuffer);

    if (action === "analyze") {
      const rows = buildAnalysisRows({
        workbook,
        ingredientRows,
        scope,
      });
      const summary = summarizeAnalysisRows(rows);

      return jsonResponse({
        ok: true,
        summary,
        rows: rows.slice(0, MAX_PREVIEW_ROWS),
        previewLimit: MAX_PREVIEW_ROWS,
        ingredientRows: ingredientRows.length,
        ingredientSheetName: ingredientParsed.sheetName,
      });
    }

    const result = applyIngredientLocationToTemplate({
      workbook,
      ingredientRows,
      scope,
    });

    const scopeName = scope === ALL_VENUES_SCOPE ? "all-venues" : scope;
    const outputName = `erp-template-${makeSafeFilePart(scopeName)}-ingredient-location-${getDateStamp()}.xlsx`;

    const outputBuffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx",
      cellStyles: true,
    });

    return new Response(outputBuffer, {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${outputName}"`,
        "X-ERP-Existing-Count": String(result.existingCount),
        "X-ERP-Suggested-Count": String(result.suggestedCount),
        "X-ERP-Venues-Processed": String(result.venueSheetsProcessed),
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    return jsonResponse({
      ok: false,
      error: error?.message || "Could not build ERP template.",
    }, 500);
  }
}

import * as XLSXModule from "xlsx";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 120;

const XLSX = XLSXModule.default || XLSXModule;

const ALL_VENUES_SCOPE = "ALL";
const MAX_PREVIEW_ROWS = 600;
const FML_SHEET_NAME = "FML March 2026";
const UOM_SHEET_NAME = "Unit of Measure ship March 2026";

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

  if (mostlyText) return "";

  return normalizeProductCode(rawText);
};

const isLikelyTemplateCode = (value) =>
  /^\d{3,}$/.test(normalizeTemplateProductCode(value));

const excelAddress = (rowIndex, columnIndex) =>
  XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });

const isCellAddressKey = (key) => /^[A-Z]{1,3}[0-9]+$/.test(key);

const getCellValue = (cell, raw = false) => {
  if (!cell) return "";
  if (raw && cell.v !== undefined && cell.v !== null) return cell.v;
  if (cell.v !== undefined && cell.v !== null) return cell.v;
  if (cell.w !== undefined && cell.w !== null) return cell.w;
  if (cell.f) return `=${cell.f}`;
  return "";
};

const getWorksheetBounds = (worksheet) => {
  const bounds = {
    minRow: 0,
    minCol: 0,
    maxRow: 0,
    maxCol: 0,
    hasCells: false,
  };

  Object.keys(worksheet || {}).forEach((key) => {
    if (!isCellAddressKey(key)) return;

    const cell = worksheet[key];
    if (!cell) return;

    const address = XLSX.utils.decode_cell(key);

    if (!bounds.hasCells) {
      bounds.minRow = address.r;
      bounds.maxRow = address.r;
      bounds.minCol = address.c;
      bounds.maxCol = address.c;
      bounds.hasCells = true;
      return;
    }

    bounds.minRow = Math.min(bounds.minRow, address.r);
    bounds.maxRow = Math.max(bounds.maxRow, address.r);
    bounds.minCol = Math.min(bounds.minCol, address.c);
    bounds.maxCol = Math.max(bounds.maxCol, address.c);
  });

  if (worksheet?.["!ref"]) {
    try {
      const range = XLSX.utils.decode_range(worksheet["!ref"]);
      bounds.minRow = Math.min(bounds.minRow, range.s.r);
      bounds.minCol = Math.min(bounds.minCol, range.s.c);
      bounds.maxRow = Math.max(bounds.maxRow, range.e.r);
      bounds.maxCol = Math.max(bounds.maxCol, range.e.c);
      bounds.hasCells = true;
    } catch {}
  }

  return bounds;
};

const getWorksheetRows = (worksheet, raw = false) => {
  const bounds = getWorksheetBounds(worksheet);
  if (!bounds.hasCells) return [];

  const rows = [];

  for (let rowIndex = 0; rowIndex <= bounds.maxRow; rowIndex += 1) {
    const row = [];

    for (let columnIndex = 0; columnIndex <= bounds.maxCol; columnIndex += 1) {
      const address = excelAddress(rowIndex, columnIndex);
      row.push(getCellValue(worksheet[address], raw));
    }

    rows.push(row);
  }

  return rows;
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

const isInactiveFlag = (value) => {
  const text = cleanKey(value);
  return ["N", "NO", "FALSE", "INACTIVE", "0"].includes(text);
};

const cloneCell = (cell) => (cell ? JSON.parse(JSON.stringify(cell)) : null);
const cloneStyle = (style) => (style ? JSON.parse(JSON.stringify(style)) : undefined);

const applyFontColor = (cell, rgb) => {
  const style = cloneStyle(cell.s) || {};

  cell.s = {
    ...style,
    font: {
      ...(style.font || {}),
      color: { rgb },
    },
  };

  return cell;
};

const setCell = ({ worksheet, address, cell, color = "" }) => {
  const nextCell = cloneCell(cell) || { t: "s", v: "" };

  if (color === "red") {
    applyFontColor(nextCell, "FF0000");
  }

  if (color === "blue") {
    applyFontColor(nextCell, "0057B8");
  }

  worksheet[address] = nextCell;
};

const ensureWorksheetRange = ({ worksheet, rowIndex, columnIndex }) => {
  const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");

  range.s.r = Math.min(range.s.r, rowIndex);
  range.s.c = Math.min(range.s.c, columnIndex);
  range.e.r = Math.max(range.e.r, rowIndex);
  range.e.c = Math.max(range.e.c, columnIndex);

  worksheet["!ref"] = XLSX.utils.encode_range(range);
};

const makeSafeFilePart = (value) =>
  String(value || "report")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "report";

const getDateStamp = () => {
  const now = new Date();

  return (
    `${now.getFullYear()}-` +
    `${String(now.getMonth() + 1).padStart(2, "0")}-` +
    `${String(now.getDate()).padStart(2, "0")}-` +
    `${String(now.getHours()).padStart(2, "0")}` +
    `${String(now.getMinutes()).padStart(2, "0")}`
  );
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
  { match: ["18 GUNBAE"], aliases: ["GUNBAE", "GUNBAE KOREAN", "GB"] },
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

const ingredientRowSearchKey = (row) =>
  cleanKey([
    row?.restaurantName,
    row?.menuName,
    row?.category,
    row?.subCategory,
    row?.name,
    row?.productName,
    row?.assigned,
    row?.recipeName,
    row?.specialInstructions,
    row?.specialInstructions2,
  ].join(" "));

const ingredientLooksLikeProduce = (row) => {
  const key = ingredientRowSearchKey(row);
  return (
    key.includes("VEGET") ||
    key.includes("VEGETE") ||
    key.includes("FRUIT") ||
    key.includes("PRODUCE") ||
    key.includes("FRESH VEGETABLE")
  );
};

const ingredientLooksLikeFish = (row) => {
  const key = ingredientRowSearchKey(row);
  return (
    key.includes("FISH") ||
    key.includes("SEAFOOD") ||
    key.includes("SALMON") ||
    key.includes("TUNA") ||
    key.includes("SHRIMP") ||
    key.includes("LOBSTER") ||
    key.includes("CRAB") ||
    key.includes("SCALLOP")
  );
};

const ingredientLooksLikeMeat = (row) => {
  const key = ingredientRowSearchKey(row);
  return (
    key.includes("MEAT") ||
    key.includes("BEEF") ||
    key.includes("VEAL") ||
    key.includes("PORK") ||
    key.includes("LAMB") ||
    key.includes("POULTRY") ||
    key.includes("CHICKEN") ||
    key.includes("TURKEY") ||
    key.includes("BACON") ||
    key.includes("SAUSAGE")
  );
};

const ingredientLooksLikeLiquor = (row) => {
  const key = ingredientRowSearchKey(row);
  return (
    key.includes("LIQUOR") ||
    key.includes("LIQUER") ||
    key.includes("WINE") ||
    key.includes("BEER") ||
    key.includes("VODKA") ||
    key.includes("RUM") ||
    key.includes("GIN") ||
    key.includes("TEQUILA") ||
    key.includes("WHISKEY")
  );
};

const ingredientLooksLikeBakery = (row) => {
  const key = ingredientRowSearchKey(row);
  return (
    key.includes("BAKERY") ||
    key.includes("BREAD") ||
    key.includes("DOUGH") ||
    key.includes("PASTRY") ||
    key.includes("CROISSANT")
  );
};

const getTemplateSectionType = (sectionTitle) => {
  const key = cleanKey(sectionTitle);

  if (!key) return "general";
  if (key.includes("LIQUOR") || key.includes("LIQUER") || key.includes("WINE") || key.includes("BEER")) return "liquor";
  if (key.includes("CHARG")) return "chargeable";
  if (key.includes("VEGET") || key.includes("VEGETE") || key.includes("FRUIT") || key.includes("PRODUCE")) return "produce";
  if (key.includes("FISH") || key.includes("SEAFOOD")) return "fish";
  if (key.includes("MEAT") || key.includes("BUTCHER")) return "meat";
  if (key.includes("PASTRY")) return "pastry";
  if (key.includes("BAKERY") || key.includes("BAKER")) return "bakery";
  if (key.includes("MENU BASIC") || key.includes("MENU BADIC") || key.includes("PREPERATION") || key.includes("PREPARATION")) return "menuBasic";

  return "general";
};

const sectionMatchesIngredientRow = (sectionTitle, ingredientRow) => {
  const type = getTemplateSectionType(sectionTitle);
  const sectionKey = cleanKey(sectionTitle);
  const ingredientKey = ingredientRowSearchKey(ingredientRow);

  const simplifiedSection = sectionKey
    .replace(/^\d+\s+/, "")
    .replace(/\b\d{1,2}\s+\d{1,2}\s+\d{1,4}\b/g, "")
    .replace(/\bMENU BASIC PREPERATION\b/g, "MENU BASIC PREPARATION")
    .trim();

  if (
    simplifiedSection.length >= 10 &&
    ingredientKey &&
    (ingredientKey.includes(simplifiedSection) || simplifiedSection.includes(ingredientKey))
  ) {
    return true;
  }

  if (type === "produce") return ingredientLooksLikeProduce(ingredientRow);
  if (type === "fish") return ingredientLooksLikeFish(ingredientRow);
  if (type === "meat") return ingredientLooksLikeMeat(ingredientRow);
  if (type === "liquor") return ingredientLooksLikeLiquor(ingredientRow);
  if (type === "bakery" || type === "pastry") return ingredientLooksLikeBakery(ingredientRow);
  if (type === "chargeable") {
    return cleanKey(ingredientRow?.category).includes("CHARG") ||
      cleanKey(ingredientRow?.subCategory).includes("CHARG") ||
      cleanKey(ingredientRow?.menuName).includes("CHARG");
  }

  if (type === "menuBasic") {
    return (
      !ingredientLooksLikeProduce(ingredientRow) &&
      !ingredientLooksLikeFish(ingredientRow) &&
      !ingredientLooksLikeMeat(ingredientRow) &&
      !ingredientLooksLikeLiquor(ingredientRow)
    );
  }

  return false;
};

const findSheetName = (sheetNames = [], expectedName) =>
  sheetNames.find((sheetName) => cleanKey(sheetName) === cleanKey(expectedName)) ||
  sheetNames.find((sheetName) => cleanKey(sheetName).includes(cleanKey(expectedName))) ||
  "";

const isVenueSheetName = (sheetName) => {
  const key = cleanKey(sheetName);

  if (!key) return false;
  if (key.includes("FML")) return false;
  if (key.includes("UNIT OF MEASURE")) return false;
  if (key.includes("SUGGESTED ADDITIONS")) return false;
  if (key.includes("SUMMARY")) return false;
  if (key.includes("INDEX")) return false;
  if (key === "SHEET5") return false;

  return true;
};

const parseIngredientByLocationWorkbook = (workbook) => {
  const sheetName = workbook?.SheetNames?.[0] || "";
  const worksheet = workbook?.Sheets?.[sheetName];

  if (!worksheet) {
    return { rows: [], sheetName, headerRowIndex: -1 };
  }

  const rows = getWorksheetRows(worksheet, false);
  let headerRowIndex = -1;

  rows.slice(0, 30).some((row, index) => {
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
    const name = cleanText(row[nameIndex]);
    const productName = cleanText(row[assignedIndex]) || name;
    const assignedType = cleanKey(row[assignedTypeIndex]);

    if (!code || !restaurantName || !productName) return;
    if (isInactiveFlag(row[activeIndex]) || isInactiveFlag(row[assignedActiveIndex])) return;
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
      return {
        codeCol: columnIndex,
        nameCol: columnIndex + 1,
        umCol: columnIndex + 2,
        headerRowIndex: 1,
      };
    }
  }

  let bestColumn = -1;
  let bestCount = 0;

  for (let columnIndex = startColumn; columnIndex <= endColumn; columnIndex += 1) {
    let count = 0;

    for (let rowIndex = 2; rowIndex < rows.length; rowIndex += 1) {
      if (isLikelyTemplateCode(rows[rowIndex]?.[columnIndex])) {
        count += 1;
      }
    }

    if (count > bestCount) {
      bestCount = count;
      bestColumn = columnIndex;
    }
  }

  if (bestColumn >= 0 && bestCount > 0) {
    return {
      codeCol: bestColumn,
      nameCol: bestColumn + 1,
      umCol: bestColumn + 2,
      headerRowIndex: 1,
    };
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

const findTemplateBlocks = (worksheet) => {
  if (!worksheet) return [];

  const rows = getWorksheetRows(worksheet, false);
  const bounds = getWorksheetBounds(worksheet);
  const titleColumns = getTitleColumnIndexes(rows);

  return titleColumns
    .map((titleColumn, index) => {
      const nextTitleColumn = titleColumns[index + 1];

      const searchStartColumn = Math.max(0, titleColumn - 2);
      const searchEndColumn =
        nextTitleColumn !== undefined
          ? Math.max(searchStartColumn, nextTitleColumn - 1)
          : Math.min(bounds.maxCol, titleColumn + 4);

      const columnInfo = findCodeColumnInRange(
        rows,
        searchStartColumn,
        searchEndColumn
      );

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

const getExistingCodeRowsForBlock = ({ rows, block }) => {
  const items = [];
  const seen = new Set();

  for (let rowIndex = block.headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const code = normalizeTemplateProductCode(rows[rowIndex]?.[block.codeCol]);

    if (!code || seen.has(`${code}|${rowIndex}`)) continue;

    seen.add(`${code}|${rowIndex}`);

    items.push({
      code,
      rowIndex,
      product: cleanText(rows[rowIndex]?.[block.nameCol]),
    });
  }

  return items;
};

const getRequiredCodeMapForBlock = ({ ingredientRows, sheetName, sectionTitle }) => {
  const requiredMap = new Map();

  (ingredientRows || []).forEach((row) => {
    if (!ingredientRowMatchesVenueSheet(row, sheetName)) return;
    if (!sectionMatchesIngredientRow(sectionTitle, row)) return;

    const code = normalizeProductCode(row.code);
    if (!code || requiredMap.has(code)) return;

    requiredMap.set(code, row);
  });

  return requiredMap;
};

const getRequiredCodeInfoForVenue = ({ ingredientRows, sheetName, blocks }) => {
  const info = new Map();

  (blocks || []).forEach((block) => {
    const requiredMap = getRequiredCodeMapForBlock({
      ingredientRows,
      sheetName,
      sectionTitle: block.title,
    });

    requiredMap.forEach((ingredientRow, code) => {
      if (!info.has(code)) {
        info.set(code, {
          code,
          sections: [],
          ingredientRow,
        });
      }

      info.get(code).sections.push(block.title);
    });
  });

  return info;
};

const makeExistingNotUsedStatus = ({ code, requiredInfoForVenue }) => {
  const requiredInfo = requiredInfoForVenue.get(code);

  if (requiredInfo?.sections?.length) {
    return `Blue review - item used in this venue but different section: ${requiredInfo.sections.join(", ")}`;
  }

  return "Blue review - item exists in ERP template but not required by Ingredient by Location";
};

const buildAnalysisRows = ({ workbook, ingredientRows, scope = ALL_VENUES_SCOPE }) => {
  if (!workbook || !ingredientRows?.length) return [];

  const venueSheets = (workbook.SheetNames || [])
    .filter(isVenueSheetName)
    .filter((sheetName) => scope === ALL_VENUES_SCOPE || sheetName === scope);

  const reportRows = [];

  venueSheets.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = getWorksheetRows(worksheet, false);
    const blocks = findTemplateBlocks(worksheet);
    const requiredInfoForVenue = getRequiredCodeInfoForVenue({
      ingredientRows,
      sheetName,
      blocks,
    });

    blocks.forEach((block) => {
      const requiredMap = getRequiredCodeMapForBlock({
        ingredientRows,
        sheetName,
        sectionTitle: block.title,
      });

      const existingItems = getExistingCodeRowsForBlock({ rows, block });
      const existingCodes = new Set(existingItems.map((item) => item.code));

      existingItems.forEach((existingItem) => {
        if (requiredMap.has(existingItem.code)) return;

        reportRows.push({
          Status: makeExistingNotUsedStatus({
            code: existingItem.code,
            requiredInfoForVenue,
          }),
          Venue: sheetName,
          Section: block.title,
          Code: existingItem.code,
          Product: existingItem.product || "Existing ERP template item",
          Restaurant: "ERP Template",
          Menu: "Existing item not required here",
          Category: "Template extra / review",
          SubCategory: "Blue review",
          IngredientSourceRow: "ERP template",
        });
      });

      requiredMap.forEach((ingredientRow, code) => {
        reportRows.push({
          Status: existingCodes.has(code)
            ? "Already in template"
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
      });
    });
  });

  return reportRows;
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

  const nameFormula = fmlSheetName
    ? `VLOOKUP(${codeAddress},'${safeFmlName}'!E:F,2,FALSE)`
    : "";

  const umFormula = uomSheetName
    ? `VLOOKUP(${codeAddress},'${safeUomName}'!A:C,3,FALSE)`
    : "";

  setCell({
    worksheet,
    address: codeAddress,
    cell: makeCodeCell(code, worksheet[sourceCodeAddress]?.s),
    color: "red",
  });

  setCell({
    worksheet,
    address: nameAddress,
    cell: nameFormula
      ? makeFormulaCell({
          formula: nameFormula,
          fallbackValue: ingredientRow.productName || ingredientRow.name,
          sourceStyle: worksheet[sourceNameAddress]?.s,
        })
      : {
          t: "s",
          v: ingredientRow.productName || ingredientRow.name || "",
          s: cloneStyle(worksheet[sourceNameAddress]?.s),
        },
    color: "red",
  });

  if (umAddress) {
    setCell({
      worksheet,
      address: umAddress,
      cell: umFormula
        ? makeFormulaCell({
            formula: umFormula,
            fallbackValue: ingredientRow.um || "",
            sourceStyle: worksheet[sourceUmAddress]?.s,
          })
        : {
            t: "s",
            v: ingredientRow.um || "",
            s: cloneStyle(worksheet[sourceUmAddress]?.s),
          },
      color: "red",
    });
  }

  worksheet[codeAddress].c = [
    {
      a: "ERP Template Builder",
      t: "Suggested addition from Ingredient by Location",
    },
  ];

  ensureWorksheetRange({ worksheet, rowIndex, columnIndex: block.codeCol });
  ensureWorksheetRange({ worksheet, rowIndex, columnIndex: block.nameCol });
  if (block.umCol >= 0) {
    ensureWorksheetRange({ worksheet, rowIndex, columnIndex: block.umCol });
  }
};

const markExistingItemBlue = ({ worksheet, block, rowIndex }) => {
  [block.codeCol, block.nameCol, block.umCol]
    .filter((columnIndex) => columnIndex >= 0)
    .forEach((columnIndex) => {
      const address = excelAddress(rowIndex, columnIndex);
      const existingCell = worksheet[address];

      if (!existingCell) return;

      setCell({
        worksheet,
        address,
        cell: existingCell,
        color: "blue",
      });
    });
};

const applyIngredientLocationToTemplate = ({
  workbook,
  ingredientRows,
  scope = ALL_VENUES_SCOPE,
}) => {
  const fmlSheetName = findSheetName(workbook.SheetNames || [], FML_SHEET_NAME);
  const uomSheetName = findSheetName(workbook.SheetNames || [], UOM_SHEET_NAME);

  const venueSheets = (workbook.SheetNames || [])
    .filter(isVenueSheetName)
    .filter((sheetName) => scope === ALL_VENUES_SCOPE || sheetName === scope);

  const summaryRows = [];
  let existingCount = 0;
  let suggestedCount = 0;
  let blueCount = 0;

  venueSheets.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = getWorksheetRows(worksheet, false);
    const blocks = findTemplateBlocks(worksheet);
    const requiredInfoForVenue = getRequiredCodeInfoForVenue({
      ingredientRows,
      sheetName,
      blocks,
    });

    blocks.forEach((block) => {
      const requiredMap = getRequiredCodeMapForBlock({
        ingredientRows,
        sheetName,
        sectionTitle: block.title,
      });

      const existingItems = getExistingCodeRowsForBlock({ rows, block });
      const existingCodes = new Set(existingItems.map((item) => item.code));

      existingItems.forEach((existingItem) => {
        if (requiredMap.has(existingItem.code)) return;

        markExistingItemBlue({
          worksheet,
          block,
          rowIndex: existingItem.rowIndex,
        });

        blueCount += 1;

        summaryRows.push({
          Status: makeExistingNotUsedStatus({
            code: existingItem.code,
            requiredInfoForVenue,
          }),
          Venue: sheetName,
          Section: block.title,
          Code: existingItem.code,
          Product: existingItem.product || "Existing ERP template item",
          Restaurant: "ERP Template",
          Menu: "Existing item not required here",
          Category: "Template extra / review",
          SubCategory: "Blue review",
          IngredientSourceRow: "ERP template",
        });
      });

      let nextRowIndex = Math.max(
        block.lastDataRowIndex + 1,
        block.headerRowIndex + 2
      );

      requiredMap.forEach((ingredientRow, code) => {
        if (existingCodes.has(code)) {
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
          worksheet,
          block,
          rowIndex: nextRowIndex,
          ingredientRow,
          fmlSheetName,
          uomSheetName,
        });

        suggestedCount += 1;

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
  });

  const summaryWorksheet = XLSX.utils.json_to_sheet(summaryRows);

  summaryWorksheet["!cols"] = [
    { wch: 48 },
    { wch: 30 },
    { wch: 56 },
    { wch: 14 },
    { wch: 50 },
    { wch: 28 },
    { wch: 30 },
    { wch: 36 },
    { wch: 36 },
    { wch: 18 },
  ];

  const summarySheetName = "Suggested Additions";

  if (workbook.SheetNames.includes(summarySheetName)) {
    workbook.Sheets[summarySheetName] = summaryWorksheet;
  } else {
    workbook.SheetNames.push(summarySheetName);
    workbook.Sheets[summarySheetName] = summaryWorksheet;
  }

  return {
    existingCount,
    suggestedCount,
    blueCount,
    summaryRows,
    venueSheetsProcessed: venueSheets.length,
  };
};

const summarizeAnalysisRows = (rows = []) => ({
  total: rows.length,
  existing: rows.filter((row) =>
    String(row.Status || "").toLowerCase().includes("already in template")
  ).length,
  suggested: rows.filter((row) =>
    String(row.Status || "").toLowerCase().includes("suggested")
  ).length,
  blue: rows.filter((row) =>
    String(row.Status || "").toLowerCase().includes("blue")
  ).length,
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

const getFileBuffer = async (file) =>
  Buffer.from(await file.arrayBuffer());

const getVenueSheetsFromTemplateBuffer = (templateBuffer) => {
  const workbook = XLSX.read(templateBuffer, {
    type: "buffer",
    bookSheets: true,
    nodim: true,
  });

  return (workbook.SheetNames || []).filter(isVenueSheetName);
};

const readIngredientRowsFromBuffer = (ingredientBuffer) => {
  const workbook = XLSX.read(ingredientBuffer, {
    type: "buffer",
    cellDates: true,
    nodim: true,
  });

  return parseIngredientByLocationWorkbook(workbook);
};

const readFullTemplateWorkbookFromBuffer = (templateBuffer) =>
  XLSX.read(templateBuffer, {
    type: "buffer",
    cellDates: true,
    nodim: true,
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
        "X-ERP-Blue-Count": String(result.blueCount || 0),
        "X-ERP-Venues-Processed": String(result.venueSheetsProcessed),
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    return jsonResponse(
      {
        ok: false,
        error: error?.message || "Could not build ERP template.",
      },
      500
    );
  }
}

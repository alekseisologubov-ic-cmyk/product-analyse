"use client";

import React, { useEffect, useMemo, useState } from "react";

const DEFAULT_ERP_LOCATION_TEMPLATE_PATH = "/erp-template-locations.xlsx";
const DEFAULT_FML_SHEET_NAME = "FML March 2026";
const MAX_PREVIEW_ROWS = 900;

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

const findVenueBlocks = ({ workbook, sheetName, XLSX }) => {
  const worksheet = workbook?.Sheets?.[sheetName];
  if (!worksheet || !XLSX) return [];

  const rows = getWorksheetRows({ worksheet, XLSX });
  const blocks = [];
  const seenKeys = new Set();

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

      const blockKey = `${rowIndex}|${colIndex}|${nameCol}|${umCol}`;
      if (seenKeys.has(blockKey)) return;
      seenKeys.add(blockKey);

      let itemCount = 0;
      for (let dataRowIndex = rowIndex + 1; dataRowIndex < rows.length; dataRowIndex += 1) {
        const code = normalizeProductCode(rows[dataRowIndex]?.[colIndex]);
        if (code) itemCount += 1;
      }

      if (!itemCount) return;

      blocks.push({
        blockKey,
        blockIndex: blocks.length,
        title: guessBlockTitle({ rows, headerRowIndex: rowIndex, codeCol: colIndex }),
        headerRowIndex: rowIndex,
        codeCol: colIndex,
        nameCol,
        umCol,
        itemCount,
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

const buildVenueItemRows = ({ workbook, sheetName, fmlIndex, XLSX }) => {
  const worksheet = workbook?.Sheets?.[sheetName];
  if (!worksheet || !XLSX) return [];

  const rows = getWorksheetRows({ worksheet, XLSX });
  const blocks = findVenueBlocks({ workbook, sheetName, XLSX });
  const items = [];
  const sectionRunningCounts = new Map();

  blocks.forEach((block) => {
    for (let rowIndex = block.headerRowIndex + 1; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex] || [];
      const code = normalizeProductCode(row[block.codeCol]);
      if (!code) continue;

      const fmlItem = fmlIndex?.byCode?.get(code) || null;
      const oldName = cleanText(row[block.nameCol]);
      const oldUm = block.umCol >= 0 ? cleanText(row[block.umCol]) : "";
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

      items.push({
        key: `${sheetName}|${block.blockKey}|${rowIndex}|${block.codeCol}|${code}`,
        venueSheet: sheetName,
        sectionTitle: block.title,
        sectionIndex: block.blockIndex,
        sectionRowNumber: currentCount + 1,
        sourceRow: rowIndex + 1,
        code,
        displayCode: fmlItem?.code || code,
        productName,
        um,
        oldName,
        oldUm,
        fmlMatched: Boolean(fmlItem),
        wasCorrected:
          Boolean(fmlItem) &&
          (cleanKey(oldName) !== cleanKey(fmlItem.productName) || cleanKey(oldUm) !== cleanKey(fmlItem.um)),
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
        fmlRow: fmlItem?.fmlRow || "",
        fmlAddress: fmlItem?.fmlAddress || "",
        codeAddress,
        nameAddress,
        umAddress,
        block,
      });
    }
  });

  return items;
};

const buildCorrectionOverrideMap = (items) => {
  const overrideMap = new Map();

  items.forEach((item) => {
    if (item.codeAddress) {
      overrideMap.set(item.codeAddress, {
        value: item.displayCode,
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

  const startRow = decodedRange.s.r || 0;
  const endRow = Math.max(decodedRange.e.r || 0, rows.length - 1);
  const startCol = decodedRange.s.c || 0;
  const endCol = Math.max(decodedRange.e.c || 0, maxColFromRows);
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

const applyFmlCorrectionsToWorksheet = ({ worksheet, items, fmlSheetName, valuesOnly }) => {
  if (!worksheet) return;

  items.forEach((item) => {
    if (item.nameAddress) {
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

const getSummary = (items) => {
  const sectionSet = new Set(items.map((item) => item.sectionTitle).filter(Boolean));

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
          `Workbook loaded. FML source: ${detectedFmlSheetName}. ${detectedVenueSheets.length} venue/location tab(s) found. Item names and UM now come from FML, not from old venue-tab values.`
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
    () => buildVenueItemRows({ workbook, sheetName: selectedVenueSheet, fmlIndex, XLSX: xlsxApi }),
    [workbook, selectedVenueSheet, fmlIndex, xlsxApi]
  );

  const selectedGridRows = useMemo(
    () => buildCorrectedGridRows({ workbook, sheetName: selectedVenueSheet, items: selectedVenueItems, XLSX: xlsxApi }),
    [workbook, selectedVenueSheet, selectedVenueItems, xlsxApi]
  );

  const summary = useMemo(() => getSummary(selectedVenueItems), [selectedVenueItems]);

  const visibleGridRows = useMemo(() => {
    const term = search.toLowerCase().trim();
    let rows = selectedGridRows;

    if (showFmlRowsOnly) {
      rows = rows.filter((row) => row.rowHasFml);
    }

    if (term) {
      rows = rows.filter((row) => row.text.toLowerCase().includes(term));
    }

    return rows.slice(0, MAX_PREVIEW_ROWS);
  }, [selectedGridRows, search, showFmlRowsOnly]);

  const visibleItemRows = useMemo(() => {
    const term = search.toLowerCase().trim();

    return selectedVenueItems.filter((item) => {
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
        item.status,
        item.sourceRow,
        item.fmlRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  }, [selectedVenueItems, search, showFmlRowsOnly]);

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
        items: selectedVenueItems,
        fmlSheetName,
        valuesOnly: false,
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

      const outputName = `erp-${makeSafeFilePart(selectedVenueSheet)}-fml-linked-${getDateStamp()}.xlsx`;

      XLSX.writeFile(outputWorkbook, outputName, {
        bookType: "xlsx",
        cellStyles: true,
      });

      logUsageEvent?.("erp_selected_venue_fml_linked_workbook_downloaded", {
        module: "erp_location_ingredients",
        venueSheet: selectedVenueSheet,
        fmlSheetName,
        outputName,
        rows: selectedVenueItems.length,
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
          });

          applyFmlCorrectionsToWorksheet({
            worksheet: clonedSheet,
            items: sheetItems,
            fmlSheetName,
            valuesOnly: false,
          });
        }

        outputWorkbook.Sheets[sheetName] = clonedSheet;
      });

      const baseName = fileName.replace(/\.(xlsx|xlsm|xls)$/i, "") || "erp-template-locations";
      const outputName = `${makeSafeFilePart(baseName)}-all-venues-fml-linked-${getDateStamp()}.xlsx`;

      XLSX.writeFile(outputWorkbook, outputName, {
        bookType: "xlsx",
        cellStyles: true,
      });

      logUsageEvent?.("erp_full_fml_linked_workbook_downloaded", {
        module: "erp_location_ingredients",
        outputName,
        venueTabs: venueSheets.length,
        fmlSheetName,
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
        makeSafeSheetName(selectedVenueSheet || "Venue")
      );

      const outputName = `erp-${makeSafeFilePart(selectedVenueSheet)}-fml-values-${getDateStamp()}.xlsx`;
      XLSX.writeFile(outputWorkbook, outputName);

      logUsageEvent?.("erp_selected_venue_fml_values_downloaded", {
        module: "erp_location_ingredients",
        venueSheet: selectedVenueSheet,
        outputName,
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
      fileName: `erp-${makeSafeFilePart(selectedVenueSheet)}-fml-corrected-${getDateStamp()}.csv`,
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
          <title>ERP Venue Ingredients - ${escapeHtmlValue(selectedVenueSheet)}</title>
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
      rows: visibleGridRows.length,
    });
  };

  return (
    <div style={styles.page || { minHeight: "100vh", padding: 24, background: "#f5f5f5", fontFamily: "Arial, sans-serif" }}>
      <div style={styles.header || { display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", background: "#fff", borderRadius: 16, padding: 18, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ minWidth: 280 }}>
          <h1 style={{ margin: 0 }}>ERP Venue Ingredients</h1>
          <p style={styles.subtitle || { margin: "4px 0 0", color: "#666" }}>
            Venue tabs keep the template layout. Item name and UM come from <strong>{fmlSheetName || DEFAULT_FML_SHEET_NAME}</strong>.
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
            Save the workbook as <strong>public/erp-template-locations.xlsx</strong>, or upload it here. The app uses the venue/location tabs only for Code positions and section titles. The actual ingredient name and UM are rebuilt from the FML sheet.
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
          ["Sections", summary.sections],
          ["Venue items", summary.totalItems],
          ["FML matched", summary.fmlMatched],
          ["Missing FML", summary.missingFml],
          ["Corrected", summary.corrected],
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
              onChange={(event) => setSelectedVenueSheet(event.target.value)}
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
          Show only rows corrected from FML
        </label>

        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button type="button" style={primaryButtonStyle} onClick={downloadSelectedVenueLinkedWorkbook} disabled={!selectedVenueSheet}>
            Download Venue Excel Linked to FML
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadFullCorrectedWorkbook} disabled={!sourceArrayBuffer}>
            Download Full Corrected Workbook
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadSelectedVenueResolvedExcel} disabled={!visibleGridRows.length}>
            Download Values Excel
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
            Showing {visibleItemRows.length} item(s) from the selected venue tab. Codes and section placement come from the venue sheet; ingredient details come from FML.
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

      <div style={{ ...cardStyle, display: "grid", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0 }}>Corrected Template Preview</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            This keeps the same row and column layout as the venue tab. Blue-highlighted code, name, and UM cells are rebuilt from the FML sheet.
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

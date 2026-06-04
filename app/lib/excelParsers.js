import * as XLSX from "xlsx";
import JSZip from "jszip";

export const cleanParserText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

export const normalizeParserVenue = (value) =>
  cleanParserText(value)
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\s*-\s*VV$/g, "")
    .replace(/\s*VV$/g, "")
    .replace(/\bTHE\s+/g, "")
    .replace(/\bSCL\b/g, "")
    .replace(/\bVAL\b/g, "")
    .replace(/\bRES\b/g, "")
    .replace(/\bBRL\b/g, "")
    .replace(/\bROJO\b/g, "")
    .replace(/\bARIYA\b/g, "")
    .replace(/\bONLY\b/g, "")
    .replace(/\bMANNOR\b/g, "MANOR")
    .replace(/\s+/g, " ")
    .trim();

export const cleanParserTemplateTitle = (value) =>
  String(value || "")
    .replace(/\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/g, "")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*$/g, "")
    .trim();

export const cleanParserTemplateSheetDisplay = (sheetName) =>
  String(sheetName || "")
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\bSCL\b/gi, "")
    .replace(/\bVAL\b/gi, "")
    .replace(/\bRES\b/gi, "")
    .replace(/\bBRL\b/gi, "")
    .replace(/\bROJO\b/gi, "")
    .replace(/\bARIYA\b/gi, "")
    .replace(/\bONLY\b/gi, "")
    .replace(/\s*-\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim();

export const getParserTemplateSheetShipScope = (sheetName) => {
  const text = cleanParserText(sheetName)
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

export const getParserTemplateShipScopeLabel = (shipScope) => {
  const scope = Array.isArray(shipScope) ? shipScope.filter(Boolean) : [];

  return scope.length
    ? "Used only on " + scope.join(", ")
    : "Used by all ships";
};

export const getParserTemplateSectionName = (templateName) => {
  const cleaned = cleanParserTemplateTitle(templateName);

  const parts = cleaned
    .split(/\s*-\s*/)
    .map((part) => part.trim())
    .filter(Boolean);

  return parts.length > 1 ? parts[parts.length - 1] : cleaned;
};

export const getParserTemplateLocationDisplay = (sheetName, templateName) => {
  const sheetDisplay = cleanParserTemplateSheetDisplay(sheetName);
  const sectionName = getParserTemplateSectionName(templateName);

  if (!sheetDisplay && !sectionName) return "Template";
  if (!sheetDisplay) return sectionName;
  if (!sectionName) return sheetDisplay;

  const sheetKey = normalizeParserVenue(sheetDisplay);
  const sectionKey = normalizeParserVenue(sectionName);

  if (!sectionKey || sheetKey === sectionKey || sheetKey.includes(sectionKey)) {
    return sheetDisplay;
  }

  return `${sheetDisplay} - ${sectionName}`;
};

export const getParserTemplateLocationKey = (sheetName, templateName) =>
  normalizeParserVenue(
    getParserTemplateLocationDisplay(sheetName, templateName)
  );

export const readExcelFile = (file, callback, onError) => {
  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      const result = event.target?.result;

      if (!result) {
        throw new Error("The uploaded Excel file could not be read.");
      }

      const workbook = XLSX.read(result, {
        type: "array",
        cellDates: true,
      });

      callback(workbook);
    } catch (error) {
      const message = error?.message || "Could not read the Excel file.";

      if (onError) {
        onError(error);
      } else if (typeof window !== "undefined") {
        window.alert(message);
      }
    }
  };

  reader.onerror = () => {
    const error = new Error("Could not open the uploaded file.");

    if (onError) {
      onError(error);
    } else if (typeof window !== "undefined") {
      window.alert(error.message);
    }
  };

  reader.readAsArrayBuffer(file);
};

export const workbookToRows = (workbook) => {
  const worksheet = workbook?.Sheets?.[workbook?.SheetNames?.[0]];

  if (!worksheet) {
    return [];
  }

  return XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });
};

export const parseTemplateWorkbook = (workbook) => {
  const map = {};

  if (!workbook?.SheetNames?.length) {
    return map;
  }

  workbook.SheetNames.forEach((sheetName) => {
    const venueKey = normalizeParserVenue(sheetName);
    if (!venueKey) return;

    if (!map[venueKey]) map[venueKey] = {};

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
    });

    if (!rows.length) return;

    rows.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (cleanParserText(cell) !== "INGREDIENT NAME") return;

        const templateName = cleanParserTemplateTitle(
          rows[rowIndex - 1]?.[colIndex] ||
            rows[rowIndex - 1]?.[colIndex - 1] ||
            sheetName ||
            "Template"
        );

        const shipScope = getParserTemplateSheetShipScope(sheetName);

        const templateLocation = {
          locationKey: getParserTemplateLocationKey(sheetName, templateName),
          displayName: getParserTemplateLocationDisplay(sheetName, templateName),
          sheetName,
          templateName: templateName || sheetName || "Template",
          shipScope,
          shipScopeLabel: getParserTemplateShipScopeLabel(shipScope),
        };

        rows.slice(rowIndex + 1).forEach((dataRow) => {
          const product = String(dataRow[colIndex] || "").trim();
          if (!product) return;

          const productKey = cleanParserText(product);
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

          if (templateCode && cleanParserText(templateCode) !== "CODE") {
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

export const columnNumberToLetters = (columnNumberZeroBased) => {
  let number = Number(columnNumberZeroBased || 0) + 1;
  let letters = "";

  while (number > 0) {
    const remainder = (number - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    number = Math.floor((number - 1) / 26);
  }

  return letters;
};

const normalizeZipPath = (path) => {
  const parts = [];

  String(path || "")
    .replace(/^\/+/, "")
    .split("/")
    .forEach((part) => {
      if (!part || part === ".") return;

      if (part === "..") {
        parts.pop();
        return;
      }

      parts.push(part);
    });

  return parts.join("/");
};

const resolveZipPath = (basePath, target) => {
  const value = String(target || "");
  if (!value) return "";

  if (value.startsWith("/")) {
    return normalizeZipPath(value);
  }

  const baseDir = String(basePath || "").split("/").slice(0, -1).join("/");
  return normalizeZipPath(baseDir + "/" + value);
};

const getElementsByLocalName = (node, localName) => {
  if (!node) return [];

  return Array.from(node.getElementsByTagName("*")).filter(
    (element) => element.localName === localName
  );
};

const getXmlRelationships = (xmlText) => {
  if (typeof DOMParser === "undefined") {
    return {};
  }

  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  const rels = {};

  getElementsByLocalName(doc, "Relationship").forEach((node) => {
    rels[node.getAttribute("Id")] = node.getAttribute("Target");
  });

  return rels;
};

const getWorkbookSheetPath = async (zip, sheetNameToFind) => {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text");

  const workbookRelsXml = await zip
    .file("xl/_rels/workbook.xml.rels")
    ?.async("text");

  if (!workbookXml || !workbookRelsXml) return "";
  if (typeof DOMParser === "undefined") return "";

  const workbookDoc = new DOMParser().parseFromString(
    workbookXml,
    "application/xml"
  );

  const workbookRels = getXmlRelationships(workbookRelsXml);

  const sheets = getElementsByLocalName(workbookDoc, "sheet");
  const wanted = cleanParserText(sheetNameToFind);

  const sheetNode =
    sheets.find(
      (sheet) => cleanParserText(sheet.getAttribute("name")) === wanted
    ) ||
    sheets.find((sheet) =>
      cleanParserText(sheet.getAttribute("name")).includes(wanted)
    );

  if (!sheetNode) return "";

  const relationId =
    sheetNode.getAttribute("r:id") ||
    sheetNode.getAttribute("id") ||
    sheetNode.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id"
    );

  const target = workbookRels[relationId];
  if (!target) return "";

  return resolveZipPath("xl/workbook.xml", target);
};

export const isParserUsableImageValue = (value) => {
  const text = String(value || "").trim();

  if (!text) return false;
  if (text.startsWith("data:image/")) return true;
  if (/^https?:\/\//i.test(text)) return true;
  if (text.includes("drive.google.com")) return true;
  if (text.includes("sharepoint.com")) return true;
  if (text.includes("1drv.ms")) return true;

  return false;
};

export const extractEmbeddedImagesByCell = async (arrayBuffer, sheetName) => {
  const imageMap = {};

  if (!arrayBuffer || !sheetName) {
    return imageMap;
  }

  if (typeof DOMParser === "undefined") {
    return imageMap;
  }

  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const sheetPath = await getWorkbookSheetPath(zip, sheetName);

    if (!sheetPath) return imageMap;

    const sheetFileName = sheetPath.split("/").pop();

    const sheetRelsPath = sheetPath.replace(
      "/worksheets/" + sheetFileName,
      "/worksheets/_rels/" + sheetFileName + ".rels"
    );

    const sheetRelsXml = await zip.file(sheetRelsPath)?.async("text");
    if (!sheetRelsXml) return imageMap;

    const sheetRels = getXmlRelationships(sheetRelsXml);

    const drawingTarget = Object.values(sheetRels).find((target) =>
      String(target || "").includes("drawings/")
    );

    if (!drawingTarget) return imageMap;

    const drawingPath = resolveZipPath(sheetPath, drawingTarget);
    const drawingXml = await zip.file(drawingPath)?.async("text");

    if (!drawingXml) return imageMap;

    const drawingDoc = new DOMParser().parseFromString(
      drawingXml,
      "application/xml"
    );

    const drawingFileName = drawingPath.split("/").pop();

    const drawingRelsPath = drawingPath.replace(
      "/drawings/" + drawingFileName,
      "/drawings/_rels/" + drawingFileName + ".rels"
    );

    const drawingRelsXml = await zip.file(drawingRelsPath)?.async("text");
    const drawingRels = drawingRelsXml ? getXmlRelationships(drawingRelsXml) : {};

    const anchors = [
      ...getElementsByLocalName(drawingDoc, "oneCellAnchor"),
      ...getElementsByLocalName(drawingDoc, "twoCellAnchor"),
    ];

    for (const anchor of anchors) {
      const from = getElementsByLocalName(anchor, "from")[0];
      if (!from) continue;

      const colNode = getElementsByLocalName(from, "col")[0];
      const rowNode = getElementsByLocalName(from, "row")[0];

      const colNumber = Number(colNode?.textContent || 0);
      const rowNumber = Number(rowNode?.textContent || 0);

      const cellAddress =
        columnNumberToLetters(colNumber) + String(rowNumber + 1);

      let dataUrl = "";

      const cNvPr = getElementsByLocalName(anchor, "cNvPr")[0];
      const description = cNvPr?.getAttribute("descr") || "";

      if (description.startsWith("data:image/")) {
        dataUrl = description;
      }

      if (!dataUrl) {
        const blip = getElementsByLocalName(anchor, "blip")[0];

        const embedId =
          blip?.getAttribute("r:embed") ||
          blip?.getAttribute("embed") ||
          blip?.getAttributeNS(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "embed"
          );

        const imageTarget = drawingRels[embedId];

        if (imageTarget) {
          const imagePath = resolveZipPath(drawingPath, imageTarget);
          const imageFile = zip.file(imagePath);

          if (imageFile) {
            const base64 = await imageFile.async("base64");
            const extension = imagePath.split(".").pop()?.toLowerCase();

            const mime =
              extension === "jpg" || extension === "jpeg"
                ? "image/jpeg"
                : extension === "webp"
                ? "image/webp"
                : "image/png";

            dataUrl = "data:" + mime + ";base64," + base64;
          }
        }
      }

      if (dataUrl) {
        imageMap[cellAddress] = dataUrl;
      }
    }
  } catch {
    return imageMap;
  }

  return imageMap;
};

export const parseMusterWorkbook = (workbook, imageMapsBySheet = {}) => {
  const items = [];

  if (!workbook?.SheetNames?.length) {
    return items;
  }

  const findHeaderIndexes = (rows) => {
    let headerRowIndex = 0;
    let headerRow = rows[0] || [];

    rows.slice(0, 12).some((row, index) => {
      const cleanRow = row.map((cell) => cleanParserText(cell));
      const hasCode = cleanRow.some((cell) => cell.includes("CODE"));

      const hasName = cleanRow.some(
        (cell) =>
          cell.includes("FINAL DESCRIPTION") ||
          cell.includes("DESCRIPTION") ||
          cell.includes("ITEM NAME") ||
          cell === "NAME"
      );

      if (hasCode && hasName) {
        headerRowIndex = index;
        headerRow = row;
        return true;
      }

      return false;
    });

    const cleanHeaders = headerRow.map((cell) => cleanParserText(cell));

    const findIndex = (patterns, fallback) => {
      const found = cleanHeaders.findIndex((header) =>
        patterns.some((pattern) => header.includes(pattern))
      );

      return found >= 0 ? found : fallback;
    };

    return {
      headerRowIndex,
      categoryIndex: findIndex(["SUB CATEG", "SUB CATEGORY", "CATEGORY"], 2),
      codeIndex: findIndex(["CODE", "APOLLO", "VV"], 3),
      nameIndex: findIndex(
        ["FINAL DESCRIPTION", "DESCRIPTION", "ITEM NAME", "NAME"],
        4
      ),
      imageIndex: findIndex(["PHOTO", "PICTURE", "IMAGE", "LINK"], 7),
    };
  };

  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return;

    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
    });

    if (!rows.length) return;

    const indexes = findHeaderIndexes(rows);
    const imageMap = imageMapsBySheet[sheetName] || {};

    rows.slice(indexes.headerRowIndex + 1).forEach((row, dataIndex) => {
      const sourceRow = indexes.headerRowIndex + 2 + dataIndex;

      const category = String(row[indexes.categoryIndex] || "").trim();
      const code = String(row[indexes.codeIndex] || "").trim();

      const name = String(row[indexes.nameIndex] || "")
        .replace(/\s+/g, " ")
        .trim();

      if (!name) return;
      if (cleanParserText(name).includes("FINAL DESCRIPTION")) return;
      if (cleanParserText(code) === "CODE") return;
      if (cleanParserText(code).includes("APOLLO")) return;

      const imageFromColumnI = String(row[8] || "").trim();
      const embeddedImageFromColumnI = imageMap[`I${sourceRow}`] || "";

      const imageFromColumnH = String(row[7] || "").trim();
      const embeddedImageFromColumnH = imageMap[`H${sourceRow}`] || "";

      const imageFromDetectedPhotoColumn = String(
        row[indexes.imageIndex] || ""
      ).trim();

      const detectedImageColumnLetter =
        typeof indexes.imageIndex === "number"
          ? columnNumberToLetters(indexes.imageIndex)
          : "";

      const embeddedImageFromDetectedPhotoColumn = detectedImageColumnLetter
        ? imageMap[`${detectedImageColumnLetter}${sourceRow}`] || ""
        : "";

      const imageCandidates = [
        embeddedImageFromColumnI,
        imageFromColumnI,
        embeddedImageFromColumnH,
        imageFromColumnH,
        embeddedImageFromDetectedPhotoColumn,
        imageFromDetectedPhotoColumn,
      ]
        .map((value) => String(value || "").trim())
        .filter((value) => isParserUsableImageValue(value));

      const image = imageCandidates[0] || "";
      const imageFallback =
        imageCandidates.find((value) => value !== image) || "";

      items.push({
        sheetName,
        category,
        code,
        name,
        image,
        imageFallback,
        sourceRow,
      });
    });
  });

  return items;
};

const getFileFromParserInput = (input) => input?.file || input;

export const parseBarInventoryFile = async (input) => {
  const file = getFileFromParserInput(input);

  if (!file) {
    throw new Error("No Bar master list file selected.");
  }

  const arrayBuffer = await file.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
  });

  const oldSheetName =
    workbook.SheetNames.find((name) => cleanParserText(name) === "OLD") ||
    workbook.SheetNames.find((name) => cleanParserText(name).includes("OLD")) ||
    workbook.SheetNames[0];

  const worksheet = workbook.Sheets[oldSheetName];

  if (!worksheet) {
    return {
      workbook,
      items: [],
      sourceSheetName: oldSheetName || "",
    };
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  const imageMap = await extractEmbeddedImagesByCell(arrayBuffer, oldSheetName);

  const items = rows
    .slice(1)
    .map((row, index) => {
      const sourceRow = index + 2;

      const code = String(row[0] || "").trim();

      const name = String(row[1] || "")
        .replace(/\s+/g, " ")
        .trim();

      const unit = String(row[2] || "").trim();

      const imageFromCell = String(row[3] || "").trim();
      const imageFromEmbeddedPicture = imageMap["D" + sourceRow] || "";

      const imageFromColumnI = String(row[8] || "").trim();
      const imageFromEmbeddedColumnI = imageMap["I" + sourceRow] || "";

      const imageCandidates = [
        imageFromColumnI,
        imageFromEmbeddedColumnI,
        imageFromCell,
        imageFromEmbeddedPicture,
      ]
        .map((value) => String(value || "").trim())
        .filter((value) => isParserUsableImageValue(value));

      const image = imageCandidates[0] || "";
      const imageFallback =
        imageCandidates.find((value) => value !== image) || "";

      return {
        equipmentDepartment: "bar",
        sheetName: oldSheetName,
        stationName: oldSheetName,
        category: "Bar",
        code,
        name,
        unit,
        um: unit,
        image,
        imageFallback,
        sourceRow,
      };
    })
    .filter((item) => item.name)
    .filter((item) => cleanParserText(item.name) !== "PRODUCT NAME")
    .filter((item) => cleanParserText(item.code) !== "PRODUCT CODE")
    .filter((item) => item.code || item.name);

  return {
    workbook,
    items,
    sourceSheetName: oldSheetName,
  };
};

const RESTAURANT_PARSER_EXCLUDED_SHEETS = new Set([
  "NEW PAR LEVELS",
  "CORE ITEMS",
  "COPY OF TEMPLATE",
  "ERP TRANSFERS",
  "IMAGES",
  "WAREHOUSE",
  "SUMMARY",
  "INDEX",
  "COVER",
]);

const getRestaurantImageLookupKey = (value) =>
  cleanParserText(value).replace(/[^A-Z0-9]/g, "");

const formatRestaurantParserCode = (value) => {
  const raw = String(value ?? "").trim();

  if (!raw) return "";

  const numberValue = Number(raw);

  if (Number.isFinite(numberValue) && Math.abs(numberValue) >= 1) {
    return String(Math.trunc(numberValue));
  }

  return raw.replace(/\.0+$/, "");
};

const parseRestaurantNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
};

const isRestaurantVenueSheetName = (sheetName, firstSheetName) => {
  const text = cleanParserText(sheetName);

  if (!text) return false;
  if (sheetName === firstSheetName) return false;
  if (RESTAURANT_PARSER_EXCLUDED_SHEETS.has(text)) return false;
  if (text.includes("MASTER")) return false;
  if (text.includes("SUMMARY")) return false;
  if (text.includes("INDEX")) return false;
  if (text.includes("COVER")) return false;

  return true;
};

const shouldParseRestaurantSheet = (sheetName, firstSheetName) =>
  sheetName === firstSheetName ||
  isRestaurantVenueSheetName(sheetName, firstSheetName);

const extractFirstUrlFromText = (value) => {
  const text = String(value || "").trim();
  if (!text) return "";

  const urlMatch = text.match(/https?:\/\/[^"',\s)]+/i);
  return urlMatch?.[0] || text;
};

const getWorksheetCellDisplayValue = (worksheet, rowNumber, colIndex) => {
  if (!worksheet || !rowNumber || typeof colIndex !== "number" || colIndex < 0) {
    return "";
  }

  const cellAddress = columnNumberToLetters(colIndex) + String(rowNumber);
  const cell = worksheet[cellAddress];

  return extractFirstUrlFromText(
    cell?.l?.Target || cell?.v || cell?.w || cell?.f || ""
  );
};

const findRestaurantHeaderIndexes = (rows, isFirstSheet) => {
  let headerRowIndex = 0;
  let headerRow = rows[0] || [];
  let looseCandidate = null;

  rows.slice(0, 20).some((row, index) => {
    const cleanRow = row.map((cell) => cleanParserText(cell));

    const hasCode = cleanRow.some(
      (cell) =>
        cell === "CODE" ||
        cell === "ERP CODE" ||
        cell.includes("PRODUCT CODE") ||
        cell.includes("ITEM CODE") ||
        cell.includes("SKU") ||
        cell.includes("APOLLO") ||
        cell.includes("VV CODE")
    );

    const hasName = cleanRow.some(
      (cell) =>
        cell.includes("FINAL DESCRIPTION") ||
        cell.includes("DESCRIPTION") ||
        cell.includes("ITEM NAME") ||
        cell.includes("PRODUCT NAME") ||
        cell.includes("EQUIPMENT") ||
        cell === "NAME"
    );

    if (hasCode && hasName) {
      headerRowIndex = index;
      headerRow = row;
      return true;
    }

    if (!looseCandidate && (hasCode || hasName)) {
      looseCandidate = {
        index,
        row,
      };
    }

    return false;
  });

  if (headerRowIndex === 0 && looseCandidate) {
    headerRowIndex = looseCandidate.index;
    headerRow = looseCandidate.row;
  }

  const cleanHeaders = headerRow.map((cell) => cleanParserText(cell));

  const findIndex = (patterns, fallback) => {
    const found = cleanHeaders.findIndex((header) =>
      patterns.some(
        (pattern) => header === pattern || header.includes(pattern)
      )
    );

    return found >= 0 ? found : fallback;
  };

  return {
    headerRowIndex,

    // First sheet / MASTER usually: A code, B name, C picture.
    // Venue sheets usually: B code, C name, D picture.
    codeIndex: findIndex(
      ["CODE", "ERP CODE", "PRODUCT CODE", "ITEM CODE", "SKU", "APOLLO", "VV CODE"],
      isFirstSheet ? 0 : 1
    ),

    nameIndex: findIndex(
      ["FINAL DESCRIPTION", "DESCRIPTION", "ITEM NAME", "PRODUCT NAME", "EQUIPMENT", "NAME"],
      isFirstSheet ? 1 : 2
    ),

    imageIndex: findIndex(
      ["IMAGE", "PHOTO", "PICTURE", "LINK"],
      isFirstSheet ? 2 : 3
    ),

    locationIndex: findIndex(
      ["SPECIFIC LOCATION", "VENUE", "LOCATION", "RESTAURANT", "CATEGORY"],
      isFirstSheet ? -1 : 4
    ),

    parIndex: findIndex(["PAR LEVEL", "PAR"], -1),
    qtyIndex: findIndex(["QTY", "QUANTITY", "ON HAND", "TOTAL"], -1),
  };
};

const getRestaurantImageCandidatesForRow = ({
  worksheet,
  imageMap,
  row,
  sourceRow,
  detectedImageIndex,
  primaryImageIndex,
  masterImage,
  isFirstSheet,
}) => {
  // Restaurant picture rule:
  // MASTER / first sheet = column C = index 2
  // Venue sheets = column D = index 3
  const imageColumnIndexes = isFirstSheet
    ? [2] // force MASTER to use column C only
    : [
        primaryImageIndex,
        detectedImageIndex,
        3, // venue picture column D
        2, // backup fallback only for venue sheets
      ]
        .filter((index) => typeof index === "number" && index >= 0)
        .filter((index, position, array) => array.indexOf(index) === position);

  const candidates = [];

  imageColumnIndexes.forEach((columnIndex) => {
    const columnLetter = columnNumberToLetters(columnIndex);

    candidates.push(imageMap[`${columnLetter}${sourceRow}`] || "");
    candidates.push(getWorksheetCellDisplayValue(worksheet, sourceRow, columnIndex));
    candidates.push(String(row[columnIndex] || "").trim());
  });

  // Venue sheets can still use MASTER image as fallback by code/name.
  // MASTER itself should use column C only, so do not add masterImage for MASTER rows.
  if (!isFirstSheet) {
    candidates.push(masterImage || "");
  }

  return candidates
    .map((value) => String(value || "").trim())
    .filter((value) => isParserUsableImageValue(value))
    .filter((value, index, array) => array.indexOf(value) === index);
};

const addRestaurantMasterImageToLookup = (lookup, item) => {
  const image = String(item?.image || item?.imageFallback || "").trim();

  if (!image) return;

  const codeKey = getRestaurantImageLookupKey(item?.code);
  const nameKey = getRestaurantImageLookupKey(item?.name);

  if (codeKey && !lookup[codeKey]) lookup[codeKey] = image;
  if (nameKey && !lookup[nameKey]) lookup[nameKey] = image;
};

const parseRestaurantSheetItems = ({
  worksheet,
  rows,
  sheetName,
  isFirstSheet,
  isVenueSheet,
  imageMap,
  masterImageByKey,
}) => {
  const indexes = findRestaurantHeaderIndexes(rows, isFirstSheet);
  const parsedItems = [];

  const primaryImageIndex = isFirstSheet ? 2 : 3;
  const stationDisplayName = isVenueSheet ? sheetName : "MASTER";

  rows.slice(indexes.headerRowIndex + 1).forEach((row, dataIndex) => {
    const sourceRow = indexes.headerRowIndex + 2 + dataIndex;

    const code = formatRestaurantParserCode(row[indexes.codeIndex]);

    const name = String(row[indexes.nameIndex] || "")
      .replace(/\s+/g, " ")
      .trim();

    if (!code && !name) return;

    const cleanCode = cleanParserText(code);
    const cleanName = cleanParserText(name);

    if (cleanCode === "CODE") return;
    if (cleanCode === "ERP CODE") return;
    if (cleanName === "DESCRIPTION") return;
    if (cleanName === "FINAL DESCRIPTION") return;
    if (cleanName === "ITEM NAME") return;
    if (cleanName === "PRODUCT NAME") return;
    if (cleanName.includes("#REF")) return;
    if (cleanName.startsWith("TOTAL")) return;

    const location =
      indexes.locationIndex >= 0
        ? String(row[indexes.locationIndex] || "")
            .replace(/\s+/g, " ")
            .trim()
        : "";

    const codeLookupKey = getRestaurantImageLookupKey(code);
    const nameLookupKey = getRestaurantImageLookupKey(name);

    const masterImage =
      masterImageByKey[codeLookupKey] ||
      masterImageByKey[nameLookupKey] ||
      "";

    const imageCandidates = getRestaurantImageCandidatesForRow({
  worksheet,
  imageMap,
  row,
  sourceRow,
  detectedImageIndex: indexes.imageIndex,
  primaryImageIndex,
  masterImage,
  isFirstSheet,
});

    const image = imageCandidates[0] || "";
    const imageFallback =
      imageCandidates.find((value) => value !== image) || "";

    const parLevel =
      indexes.parIndex >= 0 ? parseRestaurantNumber(row[indexes.parIndex]) : 0;

    const qty =
      indexes.qtyIndex >= 0 ? parseRestaurantNumber(row[indexes.qtyIndex]) : 0;

    parsedItems.push({
      equipmentDepartment: "restaurant",

      // Important:
      // For the first sheet we expose it as MASTER so it becomes a selectable station/card.
      // For venue tabs we expose the tab name as the station.
      sheetName: stationDisplayName,
      stationName: stationDisplayName,
      sourceSheetName: sheetName,

      category: location || (isVenueSheet ? sheetName : "Restaurant Master"),
      code,
      name: name || code,
      unit: "EA",
      um: "EA",
      parLevel,
      qty,
      image,
      imageFallback,
      sourceRow,
      isVenueSheet,
      isMasterSheet: !isVenueSheet,
    });
  });

  return parsedItems;
};

export const parseRestaurantInventoryFile = async (input) => {
  const file = getFileFromParserInput(input);

  if (!file) {
    throw new Error("No Restaurant master list file selected.");
  }

  const arrayBuffer = await file.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
  });

  const firstSheetName = workbook.SheetNames[0] || "";
  const masterImageByKey = {};

  const masterItems = [];
  const venueItems = [];

  for (const sheetName of workbook.SheetNames) {
    if (!shouldParseRestaurantSheet(sheetName, firstSheetName)) continue;

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    const rows = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
    });

    if (!rows.length) continue;

    const isFirstSheet = sheetName === firstSheetName;
    const isVenueSheet = isRestaurantVenueSheetName(sheetName, firstSheetName);

    const imageMap = await extractEmbeddedImagesByCell(arrayBuffer, sheetName);

    const sheetItems = parseRestaurantSheetItems({
      worksheet,
      rows,
      sheetName,
      isFirstSheet,
      isVenueSheet,
      imageMap,
      masterImageByKey,
    });

    if (isFirstSheet) {
      const normalizedMasterItems = sheetItems.map((item) => ({
        ...item,
        equipmentDepartment: "restaurant",
        sheetName: sheetName || "MASTER",
        sourceSheetName: sheetName || "MASTER",
        stationName: "MASTER",
        category: item.category || "Restaurant Master",
        isVenueSheet: false,
      }));

      masterItems.push(...normalizedMasterItems);

      normalizedMasterItems.forEach((item) =>
        addRestaurantMasterImageToLookup(masterImageByKey, item)
      );

      continue;
    }

    venueItems.push(
      ...sheetItems.map((item) => ({
        ...item,
        equipmentDepartment: "restaurant",
        sourceSheetName: sheetName,
        stationName: item.stationName || sheetName,
        isVenueSheet: true,
      }))
    );
  }

  // Important:
  // Keep MASTER first, then all restaurant / venue sheets.
  // Before, MASTER was only used as fallback when no venues existed.
  const finalItems = [...masterItems, ...venueItems];

  const venueSheetNames = workbook.SheetNames.filter((sheetName) =>
    isRestaurantVenueSheetName(sheetName, firstSheetName)
  );

  const includedSheetNames = [
    firstSheetName,
    ...venueSheetNames,
  ].filter(Boolean);

  return {
    workbook,
    items: finalItems,
    sourceSheetName: includedSheetNames.length
      ? includedSheetNames.join(", ")
      : workbook.SheetNames.join(", "),
  };
};

export const parseEquipmentMasterFile = async (input) => {
  const file = getFileFromParserInput(input);

  const equipmentDepartment = String(
    input?.equipmentDepartment || "culinary"
  ).toLowerCase();

  if (!file) {
    throw new Error("No equipment master list file selected.");
  }

  if (equipmentDepartment === "bar") {
    return parseBarInventoryFile({ file });
  }

  if (equipmentDepartment === "restaurant") {
    return parseRestaurantInventoryFile({ file });
  }

  const arrayBuffer = await file.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
  });

  const imageMapsBySheet = {};

  for (const sheetName of workbook.SheetNames) {
    imageMapsBySheet[sheetName] = await extractEmbeddedImagesByCell(
      arrayBuffer,
      sheetName
    );
  }

  const items = parseMusterWorkbook(workbook, imageMapsBySheet).map((item) => ({
    ...item,
    equipmentDepartment: equipmentDepartment || "culinary",
  }));

  return {
    workbook,
    items,
    sourceSheetName: workbook.SheetNames.join(", "),
  };
};

const parserToNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
};

export const normalizeParserShipCode = (value) => {
  const text = cleanParserText(value);

  if (!text) return "";

  if (
    text === "SC" ||
    text.includes("SCARLET") ||
    text.includes("SCARLET LADY")
  ) {
    return "SC";
  }

  if (
    text === "VL" ||
    text.includes("VALIANT") ||
    text.includes("VALIANT LADY")
  ) {
    return "VL";
  }

  if (
    text === "BRL" ||
    text === "BR" ||
    text.includes("BRILLIANT") ||
    text.includes("BRILLIANT LADY")
  ) {
    return "BRL";
  }

  if (
    text === "RL" ||
    text === "RES" ||
    text.includes("RESILIENT") ||
    text.includes("RESILIENT LADY")
  ) {
    return "RL";
  }

  return "";
};

const normalizeParserOrderCode = (value) => {
  const raw = String(value ?? "").trim();
  if (!raw) return "";

  const numberValue = Number(raw);

  if (
    Number.isFinite(numberValue) &&
    raw.replace(/\.0+$/, "") === String(Math.trunc(numberValue))
  ) {
    return String(Math.trunc(numberValue));
  }

  return cleanParserText(raw).replace(/\.0+$/, "");
};

const splitParserFmlVenues = (value) =>
  String(value || "")
    .split(",")
    .map((venue) => venue.replace(/\s+/g, " ").trim())
    .filter(Boolean);

const getParserLooseVenueMatchKey = (value) =>
  normalizeParserVenue(value).replace(/[^A-Z0-9]/g, "");

const PARSER_PRODUCT_MATCH_STOP_WORDS = new Set([
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

const singularizeParserProductToken = (token) => {
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

const getParserProductMatchTokens = (value) =>
  cleanParserText(value)
    .replace(/[^A-Z0-9]+/g, " ")
    .split(" ")
    .map((token) => singularizeParserProductToken(token.trim()))
    .filter((token) => token && token.length > 2)
    .filter((token) => !/^\d+$/.test(token))
    .filter((token) => !PARSER_PRODUCT_MATCH_STOP_WORDS.has(token));

const parserProductNamesMatch = (left, right) => {
  const a = cleanParserText(left);
  const b = cleanParserText(right);

  if (!a || !b) return false;
  if (a === b) return true;

  if (a.length > 12 && (a.includes(b) || b.includes(a))) return true;
  if (b.length > 12 && (a.includes(b) || b.includes(a))) return true;

  const aTokens = getParserProductMatchTokens(a);
  const bTokens = getParserProductMatchTokens(b);

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

const getParserProductReportKey = (value) => {
  const displayValue = String(value || "").trim();
  if (!displayValue) return "";

  const tokens = [...new Set(getParserProductMatchTokens(displayValue))].sort();

  return tokens.length ? tokens.join("|") : cleanParserText(displayValue);
};

const parserTemplateShipScopeMatches = (shipScope, currentShipCode) => {
  const scope = Array.isArray(shipScope) ? shipScope.filter(Boolean) : [];

  if (!scope.length) return true;
  if (!currentShipCode) return false;

  return scope.includes(currentShipCode);
};

const getParserTemplateMatchesForFmlProduct = ({
  fmlItem,
  currentShipCode,
  templateMap = {},
}) => {
  const matches = [];
  const seen = new Set();

  const fmlVenues = fmlItem.venues || [];

  const fmlVenueKeys = fmlVenues
    .map((venue) => getParserLooseVenueMatchKey(venue))
    .filter(Boolean);

  const fmlCodeKey = normalizeParserOrderCode(fmlItem.code);

  Object.entries(templateMap || {}).forEach(([venueKey, productsByKey]) => {
    Object.values(productsByKey || {}).forEach((templateItem) => {
      const templateCodes = Array.isArray(templateItem.productCodes)
        ? templateItem.productCodes
        : [];

      const codeMatches =
        fmlCodeKey &&
        templateCodes.some(
          (code) => normalizeParserOrderCode(code) === fmlCodeKey
        );

      const nameMatches = parserProductNamesMatch(
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

        if (!parserTemplateShipScopeMatches(shipScope, currentShipCode)) {
          return;
        }

        const candidateKeys = [
          location.locationKey,
          location.displayName,
          location.sheetName,
          venueKey,
        ]
          .map((value) => getParserLooseVenueMatchKey(value))
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
          displayName:
            location.displayName || location.locationKey || venueKey,
          shipScope,
          shipScopeLabel: getParserTemplateShipScopeLabel(shipScope),
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

export const parseFmlNotOrderedUnusedReport = ({
  workbook,
  orderRows,
  currentShipCode,
  templateMap = {},
}) => {
  const fmlSheetName =
    workbook.SheetNames.find((name) => cleanParserText(name) === "FML") ||
    workbook.SheetNames.find((name) => cleanParserText(name).includes("FML"));

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

  const orderByCode = {};
  const orderByProductKey = {};

  orderRows.forEach((item) => {
    const codeKey = normalizeParserOrderCode(item.code);
    const productKey = getParserProductReportKey(item.product);

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

    const product = String(row[4] || "")
      .replace(/\s+/g, " ")
      .trim();

    const venueText = String(row[5] || "")
      .replace(/\s+/g, " ")
      .trim();

    const uom = String(row[8] || "").trim();

    if (!code || !product || !venueText) return;
    if (cleanParserText(code) === "PRODUCT") return;
    if (cleanParserText(product) === "PRODUCT NAME") return;

    const venues = splitParserFmlVenues(venueText);
    if (!venues.length) return;

    const codeKey = normalizeParserOrderCode(code);
    const productKey = getParserProductReportKey(product);

    const orderItem =
      orderByCode[codeKey] || orderByProductKey[productKey] || null;

    if (!orderItem) return;

    const futureOrders = Number(orderItem.futureOrders || 0);
    const pastConsumption = Number(orderItem.pastConsumption || 0);

    if (futureOrders > 0 || pastConsumption > 0) return;

    const templateMatches = getParserTemplateMatchesForFmlProduct({
      fmlItem: { code, product, venues },
      currentShipCode,
      templateMap,
    });

    if (!templateMatches.length) return;

    const uniqueKey =
      codeKey || productKey || cleanParserText(product + "|" + excelRow);

    if (seen.has(uniqueKey)) return;
    seen.add(uniqueKey);

    const matchedVenues = [
      ...new Set(
        templateMatches.flatMap((match) => match.matchedVenues || [])
      ),
    ];

    const templateShipScopeLabels = [
      ...new Set(
        templateMatches.map(
          (match) => match.shipScopeLabel || "Used by all ships"
        )
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

export const parseFmlRunningLowReport = ({
  workbook,
  orderRows,
  currentShipCode,
  templateMap = {},
}) => {
  const fmlSheetName =
    workbook.SheetNames.find((name) => cleanParserText(name) === "FML") ||
    workbook.SheetNames.find((name) => cleanParserText(name).includes("FML"));

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

  const orderByCode = {};
  const orderByProductKey = {};

  orderRows.forEach((item) => {
    const codeKey = normalizeParserOrderCode(item.code);
    const productKey = getParserProductReportKey(item.product);

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

    const product = String(row[4] || "")
      .replace(/\s+/g, " ")
      .trim();

    const venueText = String(row[5] || "")
      .replace(/\s+/g, " ")
      .trim();

    const uom = String(row[8] || "").trim();

    if (!code || !product || !venueText) return;
    if (cleanParserText(code) === "PRODUCT") return;
    if (cleanParserText(product) === "PRODUCT NAME") return;

    const venues = splitParserFmlVenues(venueText);
    if (!venues.length) return;

    const codeKey = normalizeParserOrderCode(code);
    const productKey = getParserProductReportKey(product);

    const orderItem =
      orderByCode[codeKey] || orderByProductKey[productKey] || null;

    if (!orderItem) return;

    const futureOrders = Number(orderItem.futureOrders || 0);
    const pastConsumption = Number(orderItem.pastConsumption || 0);
    const averageConsumptionPerDay = Number(
      orderItem.averageConsumptionPerDay || 0
    );
    const availableAtArrival = Number(orderItem.availableAtArrival || 0);

    if (futureOrders > 0) return;
    if (pastConsumption <= 0 || averageConsumptionPerDay <= 0) return;

    const oneDayBuffer = averageConsumptionPerDay;
    const isRunningLowAtArrival = availableAtArrival <= oneDayBuffer;

    if (!isRunningLowAtArrival) return;

    const templateMatches = getParserTemplateMatchesForFmlProduct({
      fmlItem: { code, product, venues },
      currentShipCode,
      templateMap,
    });

    if (!templateMatches.length) return;

    const uniqueKey =
      codeKey || productKey || cleanParserText(product + "|" + excelRow);

    if (seen.has(uniqueKey)) return;
    seen.add(uniqueKey);

    const matchedVenues = [
      ...new Set(
        templateMatches.flatMap((match) => match.matchedVenues || [])
      ),
    ];

    const templateShipScopeLabels = [
      ...new Set(
        templateMatches.map(
          (match) => match.shipScopeLabel || "Used by all ships"
        )
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
      averageConsumptionPerDay > 0
        ? availableAtArrival / averageConsumptionPerDay
        : 0;

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

const getParserHistoricalSailorDays = (cellA, cellB) => {
  const a = parserToNumber(cellA);
  const b = parserToNumber(cellB);

  if (!a && !b) return 0;
  if (a && !b) return a;
  if (!a && b) return b;

  const low = Math.min(Math.abs(a), Math.abs(b));
  const high = Math.max(Math.abs(a), Math.abs(b));

  if (low > 0 && high > low * 1000) return high;

  return a * b;
};

const parserExcelDateToDate = (value) => {
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

const formatParserDateCell = (value) => {
  const date = parserExcelDateToDate(value);

  if (!date) {
    return String(value || "").trim();
  }

  return date.toLocaleDateString();
};

const getParserDaysBetweenCells = (startValue, endValue) => {
  const startDate = parserExcelDateToDate(startValue);
  const endDate = parserExcelDateToDate(endValue);

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

export const parseNextOrderWorkbook = ({ workbook, templateMap = {} } = {}) => {
  if (!workbook?.SheetNames?.length) {
    throw new Error("Could not read the next order workbook.");
  }

  const sheetName = workbook.SheetNames.includes("Standard Order Template")
    ? "Standard Order Template"
    : workbook.SheetNames.includes("Order Sheet")
    ? "Order Sheet"
    : workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    throw new Error("Could not find the next order worksheet.");
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  const orderShipName = String(rows[0]?.[1] || "").trim();
  const orderShipCode = normalizeParserShipCode(orderShipName);
  const rawOrderDate = rows[1]?.[1];
  const rawArrivalDate = rows[2]?.[1];
  const targetSailors = parserToNumber(rows[4]?.[1]);
  const targetDays = parserToNumber(rows[5]?.[1]);

  const daysUntilArrival = getParserDaysBetweenCells(
    rawOrderDate,
    rawArrivalDate
  );

  const currentPeriodSailorDays = targetSailors * targetDays;

  const futureOrderColumns = [5, 6, 7, 8, 9, 10, 11, 12, 13];
  const pastConsumptionColumns = [34, 35, 36, 37, 38, 39];

  const historicalSailorDays = pastConsumptionColumns.reduce(
    (sum, colIndex) =>
      sum +
      getParserHistoricalSailorDays(
        rows[4]?.[colIndex],
        rows[5]?.[colIndex]
      ),
    0
  );

  const parsedRows = [];

  rows.slice(9).forEach((row, rowOffset) => {
    const excelRow = rowOffset + 10;
    const code = String(row[0] || "").trim();
    const product = String(row[1] || "").trim();
    const uom = String(row[2] || "").trim();

    if (!product || !uom) return;

    const stockOnHand = parserToNumber(row[3]);
    const parLevel = parserToNumber(row[16]);

    const futureOrders = futureOrderColumns.reduce(
      (sum, colIndex) => sum + parserToNumber(row[colIndex]),
      0
    );

    const pastConsumption = pastConsumptionColumns.reduce(
      (sum, colIndex) => sum + parserToNumber(row[colIndex]),
      0
    );

    const averageConsumptionPerSailorDay =
      historicalSailorDays > 0 ? pastConsumption / historicalSailorDays : 0;

    const averageConsumptionPerDay =
      averageConsumptionPerSailorDay * targetSailors;

    const projectedNeed = averageConsumptionPerDay * targetDays;
    const consumptionUntilArrival = averageConsumptionPerDay * daysUntilArrival;
    const availableAtArrival =
      stockOnHand + futureOrders - consumptionUntilArrival;

    const rawSuggestedOrder = Math.max(projectedNeed - availableAtArrival, 0);

    const isFourteenDayLoad = Math.abs(Number(targetDays || 0) - 14) < 0.01;
    const parMaxAllowed = parLevel > 0 ? parLevel * 1.1 : 0;

    const parCapApplied = Boolean(
      isFourteenDayLoad &&
        parLevel > 0 &&
        rawSuggestedOrder > parMaxAllowed
    );

    const suggestedOrder = parCapApplied
      ? parMaxAllowed
      : rawSuggestedOrder;

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
    let alertLabel =
      suggestedOrder > 0 ? "Needs order" : "No order suggested";

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

  const fmlReportRows = parseFmlNotOrderedUnusedReport({
    workbook,
    orderRows: parsedRows,
    currentShipCode: orderShipCode,
    templateMap,
  });

  const fmlRunningLowRows = parseFmlRunningLowReport({
    workbook,
    orderRows: parsedRows,
    currentShipCode: orderShipCode,
    templateMap,
  });

  return {
    rows: parsedRows,
    fmlReportRows,
    fmlRunningLowRows,
    meta: {
      sheetName,
      shipName: orderShipName,
      shipCode: orderShipCode,
      orderDate: formatParserDateCell(rawOrderDate),
      arrivalDate: formatParserDateCell(rawArrivalDate),
      targetSailors,
      targetDays,
      daysUntilArrival,
      currentPeriodSailorDays,
      historicalSailorDays,
      totalItems: parsedRows.length,
      itemsNeedingOrder: parsedRows.filter(
        (item) => Number(item.suggestedOrder || 0) > 0
      ).length,
      parCapItems: parsedRows.filter((item) => item.parCapApplied).length,
      blueReviewItems: parsedRows.filter((item) => item.alertType === "blue")
        .length,
      redReviewItems: parsedRows.filter((item) => item.alertType === "red")
        .length,
      fmlMissingItems: fmlReportRows.length,
      fmlRunningLowItems: fmlRunningLowRows.length,
    },
  };
};

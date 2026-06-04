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
  const ws = workbook?.Sheets?.[workbook?.SheetNames?.[0]];

  if (!ws) {
    return [];
  }

  return XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: "",
  });
};

export const parseTemplateWorkbook = (workbook) => {
  const map = {};

  workbook.SheetNames.forEach((sheetName) => {
    const venueKey = normalizeParserVenue(sheetName);
    if (!venueKey) return;

    if (!map[venueKey]) map[venueKey] = {};

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
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
        ...map[venueKey][productKey].templates,
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

const getElementsByLocalName = (node, localName) =>
  Array.from(node.getElementsByTagName("*")).filter(
    (element) => element.localName === localName
  );

const getXmlRelationships = (xmlText) => {
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
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
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

export const parseBarInventoryFile = async (file) => {
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

export const parseEquipmentMasterFile = async (input) => {
  const file = input?.file || input;
  const equipmentDepartment = String(
    input?.equipmentDepartment || "culinary"
  ).toLowerCase();

  if (!file) {
    throw new Error("No equipment master file selected.");
  }

  if (equipmentDepartment === "bar") {
    return parseBarInventoryFile(file);
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

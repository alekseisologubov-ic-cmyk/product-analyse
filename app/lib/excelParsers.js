import * as XLSX from "xlsx";

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
  return scope.length ? "Used only on " + scope.join(", ") : "Used by all ships";
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
  const ws = workbook.Sheets[workbook.SheetNames[0]];

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

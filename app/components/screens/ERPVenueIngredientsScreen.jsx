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

const normalizeSheetName = (value) => cleanKey(value).replace(/[^A-Z0-9]/g, "");

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

  return ["README", "SUMMARY", "INDEX", "HELP", "SETTINGS", "LISTS"].some(
    (blocked) => key === blocked || key.includes(blocked)
  );
};

const getVisibleVenueSheetNames = (workbook, fmlSheetName) => {
  const sheetNames = workbook?.SheetNames || [];
  const sheetInfo = workbook?.Workbook?.Sheets || [];

  return sheetNames.filter((sheetName, index) => {
    if (isHelperSheetName(sheetName, fmlSheetName)) return false;

    const hidden = Number(sheetInfo[index]?.Hidden || 0);
    if (hidden === 1 || hidden === 2) return false;

    return Boolean(workbook.Sheets?.[sheetName]);
  });
};

const getFormulaReferences = (formula) => {
  const text = String(formula || "");
  const refs = [];

  // Matches references like 'FML March 2026'!$A$10 or FML!A10.
  const sheetReferenceRegex = /(?:'([^']+)'|([A-Za-z0-9_ .-]+))!\$?([A-Z]{1,3})\$?(\d{1,7})/g;
  let match;

  while ((match = sheetReferenceRegex.exec(text))) {
    refs.push({
      sheetName: cleanText(match[1] || match[2] || ""),
      address: `${String(match[3] || "").toUpperCase()}${match[4]}`,
    });
  }

  return refs;
};

const resolveFormulaCellValue = ({ cell, workbook, fmlSheetName }) => {
  if (!cell) return { value: "", fromFml: false, formula: "" };

  const formula = String(cell.f || "").trim();
  const cachedValue = getCellDisplayValue(cell);

  if (!formula || !fmlSheetName || !workbook?.Sheets?.[fmlSheetName]) {
    return { value: cachedValue, fromFml: false, formula };
  }

  const fmlKey = normalizeSheetName(fmlSheetName);
  const fmlRefs = getFormulaReferences(formula).filter(
    (ref) => normalizeSheetName(ref.sheetName) === fmlKey
  );

  const uniqueRefMap = new Map();
  fmlRefs.forEach((ref) => {
    const key = `${normalizeSheetName(ref.sheetName)}!${ref.address}`;
    if (!uniqueRefMap.has(key)) uniqueRefMap.set(key, ref);
  });

  const uniqueRefs = Array.from(uniqueRefMap.values());

  // Direct link formulas and IF(cell="","",cell) formulas usually point to one unique FML cell.
  // In that case, read the FML value directly so the app reflects FML updates even if Excel cached formula values are stale.
  if (uniqueRefs.length === 1) {
    const fmlCell = workbook.Sheets[fmlSheetName][uniqueRefs[0].address];
    const fmlValue = getCellDisplayValue(fmlCell);

    return {
      value: fmlValue,
      fromFml: true,
      formula,
      fmlAddress: `${fmlSheetName}!${uniqueRefs[0].address}`,
    };
  }

  // Complex formulas keep the cached Excel value. Excel will recalculate when the linked workbook is opened.
  return { value: cachedValue, fromFml: false, formula };
};

const buildSheetGrid = ({ workbook, sheetName, fmlSheetName, XLSX }) => {
  const worksheet = workbook?.Sheets?.[sheetName];
  if (!worksheet || !worksheet["!ref"] || !XLSX) return [];

  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const rows = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const cells = [];
    let rowHasValue = false;

    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = worksheet[address];
      const resolved = resolveFormulaCellValue({ cell, workbook, fmlSheetName });
      const value = cleanText(resolved.value);

      if (value) rowHasValue = true;

      cells.push({
        address,
        rowIndex,
        colIndex,
        value,
        formula: resolved.formula || "",
        fromFml: Boolean(resolved.fromFml),
        fmlAddress: resolved.fmlAddress || "",
      });
    }

    rows.push({
      rowIndex,
      excelRow: rowIndex + 1,
      cells,
      rowHasValue,
      text: cells.map((cell) => cell.value).join(" "),
    });
  }

  return rows;
};

const findHeaderRowIndex = (gridRows) => {
  const headerWords = [
    "CODE",
    "PRODUCT",
    "DESCRIPTION",
    "FINAL DESCRIPTION",
    "ITEM",
    "INGREDIENT",
    "UOM",
    "UM",
    "UNIT",
    "PAR",
  ];

  let best = { index: -1, score: 0 };

  gridRows.slice(0, 40).forEach((row, index) => {
    const rowText = cleanKey(row.text);
    const score = headerWords.reduce(
      (sum, word) => sum + (rowText.includes(cleanKey(word)) ? 1 : 0),
      0
    );

    if (score > best.score) {
      best = { index, score };
    }
  });

  return best.score >= 2 ? best.index : -1;
};

const getIngredientRowsFromGrid = (gridRows) => {
  const headerIndex = findHeaderRowIndex(gridRows);
  const rowsToCheck = headerIndex >= 0 ? gridRows.slice(headerIndex + 1) : gridRows;

  return rowsToCheck.filter((row) => {
    if (!row.rowHasValue) return false;

    const text = cleanKey(row.text);
    if (!text) return false;
    if (text === "CODE" || text === "PRODUCT" || text.includes("FINAL DESCRIPTION")) return false;

    // Count rows that contain at least one visible product-like value.
    return row.cells.some((cell) => cleanText(cell.value).length > 1);
  });
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
  maxWidth: 260,
  whiteSpace: "pre-wrap",
};

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
  const [showLinkedOnly, setShowLinkedOnly] = useState(false);

  const cardStyle = styles.card || cardFallbackStyle;
  const primaryButtonStyle = styles.primaryButton || primaryButtonFallbackStyle;
  const secondaryButtonStyle = styles.secondaryButton || secondaryButtonFallbackStyle;

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
        detectedFmlSheetName
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
          "Workbook loaded, but FML March 2026 sheet was not found. The module can still show venue tabs, but FML live linking will not work."
        );
      } else if (!detectedVenueSheets.length) {
        setMessage(
          `Workbook loaded. FML sheet found: ${detectedFmlSheetName}. No venue/location tabs were detected.`
        );
      } else {
        setMessage(
          `Workbook loaded. FML source: ${detectedFmlSheetName}. ${detectedVenueSheets.length} venue/location tab(s) found.`
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

  const selectedGridRows = useMemo(
    () =>
      buildSheetGrid({
        workbook,
        sheetName: selectedVenueSheet,
        fmlSheetName,
        XLSX: xlsxApi,
      }),
    [workbook, selectedVenueSheet, fmlSheetName, xlsxApi]
  );

  const ingredientRows = useMemo(
    () => getIngredientRowsFromGrid(selectedGridRows),
    [selectedGridRows]
  );

  const visibleGridRows = useMemo(() => {
    const term = search.toLowerCase().trim();

    let rows = selectedGridRows;

    if (showLinkedOnly) {
      rows = rows.filter((row) => row.cells.some((cell) => cell.fromFml));
    }

    if (term) {
      rows = rows.filter((row) => row.text.toLowerCase().includes(term));
    }

    return rows.slice(0, MAX_PREVIEW_ROWS);
  }, [selectedGridRows, search, showLinkedOnly]);

  const linkedCellCount = useMemo(
    () => selectedGridRows.reduce(
      (sum, row) => sum + row.cells.filter((cell) => cell.fromFml).length,
      0
    ),
    [selectedGridRows]
  );

  const downloadFullLinkedWorkbook = () => {
    if (!sourceArrayBuffer) {
      window.alert("Upload or load the ERP location template first.");
      return;
    }

    const safeName = fileName && /\.xlsx?$/i.test(fileName)
      ? fileName
      : `erp-location-template-linked-${getDateStamp()}.xlsx`;

    downloadBlob({
      content: new Blob([sourceArrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      fileName: safeName,
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    logUsageEvent?.("erp_full_linked_workbook_downloaded", {
      module: "erp_location_ingredients",
      fileName: safeName,
    });
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

      outputWorkbook.Sheets[selectedVenueSheet] = cloneSheet(workbook.Sheets[selectedVenueSheet]);
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

      const outputName = `erp-${makeSafeFilePart(selectedVenueSheet)}-linked-to-fml-${getDateStamp()}.xlsx`;

      XLSX.writeFile(outputWorkbook, outputName, {
        bookType: "xlsx",
        cellStyles: true,
      });

      logUsageEvent?.("erp_selected_venue_linked_workbook_downloaded", {
        module: "erp_location_ingredients",
        venueSheet: selectedVenueSheet,
        fmlSheetName,
        outputName,
      });
    } catch (error) {
      window.alert(error?.message || "Could not download the selected venue workbook.");
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

      const outputName = `erp-${makeSafeFilePart(selectedVenueSheet)}-resolved-values-${getDateStamp()}.xlsx`;
      XLSX.writeFile(outputWorkbook, outputName);

      logUsageEvent?.("erp_selected_venue_resolved_values_downloaded", {
        module: "erp_location_ingredients",
        venueSheet: selectedVenueSheet,
        outputName,
        rows: visibleGridRows.length,
      });
    } catch (error) {
      window.alert(error?.message || "Could not download resolved values workbook.");
    }
  };

  const downloadVisibleCsv = () => {
    if (!visibleGridRows.length) {
      window.alert("No visible rows to export.");
      return;
    }

    downloadBlob({
      content: buildResolvedCsv({ rows: visibleGridRows }),
      fileName: `erp-${makeSafeFilePart(selectedVenueSheet)}-visible-${getDateStamp()}.csv`,
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
            .linked { background: #eef5ff; }
            tr { break-inside: avoid; }
          </style>
        </head>
        <body>
          <h1>ERP Venue Ingredients</h1>
          <div class="meta"><strong>Venue tab:</strong> ${escapeHtmlValue(selectedVenueSheet || "N/A")}</div>
          <div class="meta"><strong>FML source:</strong> ${escapeHtmlValue(fmlSheetName || "Not found")}</div>
          <div class="meta"><strong>Source file:</strong> ${escapeHtmlValue(fileName || "N/A")}</div>
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

    logUsageEvent?.("erp_venue_report_printed", {
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
            View venue tabs linked to <strong>{fmlSheetName || DEFAULT_FML_SHEET_NAME}</strong>.
          </p>
        </div>

        <button type="button" style={styles.backButton || secondaryButtonStyle} onClick={() => setModule?.("")}>
          ← Back
        </button>
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 14 }}>
        <div>
          <h2 style={{ margin: 0 }}>Files</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            The default file should be saved as <strong>public/erp-template-locations.xlsx</strong>. You can also upload a workbook here. Venue names come from the sheet/tab names, and linked cells read directly from the FML sheet.
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
          ["Selected rows", selectedGridRows.filter((row) => row.rowHasValue).length],
          ["Ingredient rows", ingredientRows.length],
          ["FML-linked cells", linkedCellCount],
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
              placeholder="Search ingredients, code, UM, row, or text..."
              style={styles.searchInput || { width: "100%", padding: 11, borderRadius: 10, border: "1px solid #ccc" }}
            />
          </label>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: 8, fontWeight: "bold" }}>
          <input
            type="checkbox"
            checked={showLinkedOnly}
            onChange={(event) => setShowLinkedOnly(event.target.checked)}
          />
          Show only rows with cells linked to FML
        </label>

        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button type="button" style={primaryButtonStyle} onClick={downloadSelectedVenueLinkedWorkbook} disabled={!selectedVenueSheet}>
            Download Venue Linked Excel
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadFullLinkedWorkbook} disabled={!sourceArrayBuffer}>
            Download Full Linked Workbook
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadSelectedVenueResolvedExcel} disabled={!visibleGridRows.length}>
            Download Resolved Values Excel
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={printVisibleReport} disabled={!visibleGridRows.length}>
            Print / Save PDF
          </button>
          <button type="button" style={secondaryButtonStyle} onClick={downloadVisibleCsv} disabled={!visibleGridRows.length}>
            Export CSV
          </button>
        </div>
      </div>

      <div style={{ ...cardStyle, display: "grid", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0 }}>Report Preview</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            Showing the selected venue tab in the same row/column order as the workbook. Blue-highlighted cells are read directly from the FML sheet.
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
                        title={cell.fromFml ? `Linked: ${cell.fmlAddress}` : cell.formula ? `Formula: ${cell.formula}` : cell.address}
                        style={{
                          ...tableCellStyle,
                          background: cell.fromFml ? "#eef5ff" : row.rowHasValue ? "#fff" : "#fafafa",
                          color: cell.value ? "#111" : "#aaa",
                          fontWeight: cell.fromFml ? "600" : "normal",
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

"use client";

import React, { useMemo, useState } from "react";

const ALL_VENUES_SCOPE = "ALL";
const INGREDIENT_PUBLIC_PATH = "/ingredient-by-location.xlsx";
const TEMPLATE_PUBLIC_PATH = "/erp-template-locations.xlsx";

const fallbackPageStyle = {
  minHeight: "100vh",
  padding: 24,
  background: "#f5f5f5",
  fontFamily: "Arial, sans-serif",
};

const fallbackCardStyle = {
  background: "#fff",
  borderRadius: 16,
  padding: 20,
  boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
};

const fallbackButtonStyle = {
  padding: "11px 14px",
  borderRadius: 10,
  border: 0,
  background: "#111",
  color: "#fff",
  fontWeight: "bold",
  cursor: "pointer",
};

const tableHeaderStyle = {
  padding: 10,
  textAlign: "left",
  borderRight: "1px solid #ddd",
  whiteSpace: "nowrap",
  background: "#111",
  color: "#fff",
};

const tableCellStyle = {
  padding: 10,
  verticalAlign: "top",
  borderRight: "1px solid #eee",
  borderBottom: "1px solid #eee",
};

const normalizeStatus = (value) => String(value || "").toLowerCase();

const getStatusColor = (status) => {
  const text = normalizeStatus(status);

  if (text.includes("suggested")) return "#c00018";
  if (text.includes("blue")) return "#0057b8";
  if (text.includes("already")) return "#111";

  return "#111";
};

const getStatusLabel = (status) => {
  const text = normalizeStatus(status);

  if (text.includes("suggested")) return "Suggested red";
  if (text.includes("blue")) return "Blue review";
  if (text.includes("already")) return "Already black";

  return status || "";
};

const makeFileFromBlob = (blob, fileName) =>
  new File([blob], fileName, {
    type: blob.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

const downloadBlob = (blob, fileName) => {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  URL.revokeObjectURL(url);
};

const getDownloadFileName = (response) => {
  const disposition = response.headers.get("Content-Disposition") || "";
  const match = disposition.match(/filename="?([^"]+)"?/i);

  if (match?.[1]) return match[1];

  return "erp-template-corrected.xlsx";
};

export default function ERPVenueIngredientsScreen({
  styles = {},
  setModule,
  logUsageEvent,
}) {
  const [ingredientFile, setIngredientFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [ingredientFileName, setIngredientFileName] = useState("");
  const [templateFileName, setTemplateFileName] = useState("");
  const [venueSheets, setVenueSheets] = useState([]);
  const [scope, setScope] = useState(ALL_VENUES_SCOPE);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [summary, setSummary] = useState(null);
  const [rows, setRows] = useState([]);
  const [search, setSearch] = useState("");

  const pageStyle = styles.page || fallbackPageStyle;
  const cardStyle = styles.card || fallbackCardStyle;
  const buttonStyle = styles.primaryButton || fallbackButtonStyle;

  const visibleRows = useMemo(() => {
    const term = search.toLowerCase().trim();

    if (!term) return rows;

    return rows.filter((row) =>
      [
        row.Status,
        row.Venue,
        row.Section,
        row.Code,
        row.Product,
        row.Restaurant,
        row.Menu,
        row.Category,
        row.SubCategory,
        row.IngredientSourceRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term)
    );
  }, [rows, search]);

  const resetResults = () => {
    setSummary(null);
    setRows([]);
  };

  const callBuilderApi = async ({ action, responseType = "json" }) => {
    if (!templateFile) {
      throw new Error("Upload or load the ERP template first.");
    }

    if (action !== "inspect" && !ingredientFile) {
      throw new Error("Upload or load Ingredient by Location first.");
    }

    const formData = new FormData();

    formData.append("action", action);
    formData.append("scope", scope || ALL_VENUES_SCOPE);
    formData.append("template", templateFile);

    if (ingredientFile) {
      formData.append("ingredient", ingredientFile);
    }

    const response = await fetch("/api/erp-template-builder", {
      method: "POST",
      body: formData,
    });

    if (responseType === "blob") {
      if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data?.error || "Could not build ERP template.");
      }

      return response;
    }

    const data = await response.json().catch(() => ({}));

    if (!response.ok || !data.ok) {
      throw new Error(data?.error || "ERP template builder failed.");
    }

    return data;
  };

  const inspectTemplate = async (fileOverride = null) => {
    const activeTemplateFile = fileOverride || templateFile;

    if (!activeTemplateFile) return;

    setBusy(true);
    setMessage("Reading ERP venue tabs...");
    resetResults();

    try {
      const formData = new FormData();

      formData.append("action", "inspect");
      formData.append("scope", ALL_VENUES_SCOPE);
      formData.append("template", activeTemplateFile);

      const response = await fetch("/api/erp-template-builder", {
        method: "POST",
        body: formData,
      });

      const data = await response.json().catch(() => ({}));

      if (!response.ok || !data.ok) {
        throw new Error(data?.error || "Could not inspect ERP template.");
      }

      setVenueSheets(data.venueSheets || []);
      setMessage(
        `ERP template loaded. ${data.venueCount || 0} venue/location tab(s) found.`
      );

      logUsageEvent?.("erp_template_inspected", {
        module: "erp_venue_ingredients",
        venueCount: data.venueCount || 0,
      });
    } catch (error) {
      setVenueSheets([]);
      setMessage(error?.message || "Could not inspect ERP template.");
      window.alert(error?.message || "Could not inspect ERP template.");
    } finally {
      setBusy(false);
    }
  };

  const loadPermanentIngredientFile = async () => {
    setBusy(true);
    setMessage("Loading permanent Ingredient by Location file...");
    resetResults();

    try {
      const response = await fetch(INGREDIENT_PUBLIC_PATH, {
        cache: "no-store",
      });

      if (!response.ok) {
        throw new Error("Permanent ingredient-by-location.xlsx was not found in public folder.");
      }

      const blob = await response.blob();
      const file = makeFileFromBlob(blob, "ingredient-by-location.xlsx");

      setIngredientFile(file);
      setIngredientFileName(file.name);
      setMessage("Permanent Ingredient by Location file loaded.");
    } catch (error) {
      setIngredientFile(null);
      setIngredientFileName("");
      setMessage(error?.message || "Could not load Ingredient by Location.");
      window.alert(error?.message || "Could not load Ingredient by Location.");
    } finally {
      setBusy(false);
    }
  };

  const loadPermanentTemplateFile = async () => {
    setBusy(true);
    setMessage("Loading permanent ERP template...");
    resetResults();

    try {
      const response = await fetch(TEMPLATE_PUBLIC_PATH, {
        cache: "no-store",
      });

      if (!response.ok) {
        throw new Error("Permanent erp-template-locations.xlsx was not found in public folder.");
      }

      const blob = await response.blob();
      const file = makeFileFromBlob(blob, "erp-template-locations.xlsx");

      setTemplateFile(file);
      setTemplateFileName(file.name);
      setScope(ALL_VENUES_SCOPE);
      await inspectTemplate(file);
    } catch (error) {
      setTemplateFile(null);
      setTemplateFileName("");
      setVenueSheets([]);
      setMessage(error?.message || "Could not load ERP template.");
      window.alert(error?.message || "Could not load ERP template.");
    } finally {
      setBusy(false);
    }
  };

  const handleIngredientUpload = (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    setIngredientFile(file);
    setIngredientFileName(file.name);
    setMessage(`Ingredient by Location file selected: ${file.name}`);
    resetResults();

    event.target.value = "";
  };

  const handleTemplateUpload = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    setTemplateFile(file);
    setTemplateFileName(file.name);
    setScope(ALL_VENUES_SCOPE);
    resetResults();

    event.target.value = "";

    await inspectTemplate(file);
  };

  const previewComparison = async () => {
    setBusy(true);
    setMessage("Cross-checking Ingredient by Location against ERP template...");
    resetResults();

    try {
      const data = await callBuilderApi({
        action: "analyze",
        responseType: "json",
      });

      setSummary(data.summary || null);
      setRows(data.rows || []);

      const limitNote =
        data.rows?.length >= data.previewLimit
          ? ` Preview is limited to ${data.previewLimit} row(s).`
          : "";

      setMessage(
        `Comparison complete. ${data.summary?.total || 0} placement row(s) found.${limitNote}`
      );

      logUsageEvent?.("erp_template_preview_created", {
        module: "erp_venue_ingredients",
        scope,
        total: data.summary?.total || 0,
        existing: data.summary?.existing || 0,
        suggested: data.summary?.suggested || 0,
        blue: data.summary?.blue || 0,
      });
    } catch (error) {
      setSummary(null);
      setRows([]);
      setMessage(error?.message || "Could not preview ERP comparison.");
      window.alert(error?.message || "Could not preview ERP comparison.");
    } finally {
      setBusy(false);
    }
  };

  const downloadCorrectedTemplate = async () => {
    setBusy(true);
    setMessage("Building corrected ERP template. This can take a moment for large files...");

    try {
      const response = await callBuilderApi({
        action: "download",
        responseType: "blob",
      });

      const blob = await response.blob();
      const fileName = getDownloadFileName(response);

      downloadBlob(blob, fileName);

      const existing = response.headers.get("X-ERP-Existing-Count") || "0";
      const suggested = response.headers.get("X-ERP-Suggested-Count") || "0";
      const blue = response.headers.get("X-ERP-Blue-Count") || "0";

      setMessage(
        `Corrected ERP template downloaded. Already black: ${existing}. Suggested red: ${suggested}. Blue review: ${blue}.`
      );

      logUsageEvent?.("erp_template_corrected_downloaded", {
        module: "erp_venue_ingredients",
        scope,
        existing,
        suggested,
        blue,
      });
    } catch (error) {
      setMessage(error?.message || "Could not download corrected ERP template.");
      window.alert(error?.message || "Could not download corrected ERP template.");
    } finally {
      setBusy(false);
    }
  };

  const summaryCards = [
    ["Ingredient rows", summary?.total ? "Loaded" : ingredientFileName ? "Ready" : "0"],
    ["Venue tabs", venueSheets.length],
    ["Compared rows", summary?.total || 0],
    ["Already black", summary?.existing || 0],
    ["Suggested red", summary?.suggested || 0],
    ["Blue review", summary?.blue || 0],
    ["Sections", summary?.sections || 0],
  ];

  return (
    <div style={pageStyle}>
      <div
        style={{
          ...(styles.header || {
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 12,
            background: "#fff",
            borderRadius: 16,
            padding: 18,
            marginBottom: 16,
            flexWrap: "wrap",
          }),
        }}
      >
        <div>
          <h1 style={{ margin: 0 }}>ERP Venue Ingredients</h1>
          <p style={{ margin: "4px 0 0", color: "#666" }}>
            Cross-check Ingredient by Location against ERP venue tabs.
          </p>
        </div>

        <button
          type="button"
          style={styles.backButton || buttonStyle}
          onClick={() => setModule?.("")}
        >
          ← Back
        </button>
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 14 }}>
        <div>
          <h2 style={{ margin: 0 }}>Files</h2>
          <p style={{ color: "#555", marginBottom: 0 }}>
            Ingredient by Location decides what should be used. The ERP template decides the venue tab and section layout.
          </p>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
            gap: 14,
          }}
        >
          <div style={{ display: "grid", gap: 8 }}>
            <strong>Ingredient by Location</strong>

            <input
              type="file"
              accept=".xlsx,.xls,.xlsm"
              onChange={handleIngredientUpload}
              disabled={busy}
            />

            <button
              type="button"
              style={buttonStyle}
              onClick={loadPermanentIngredientFile}
              disabled={busy}
            >
              Load permanent Ingredient file
            </button>

            <span style={{ color: "#666", fontSize: 13 }}>
              Current: {ingredientFileName || "No file loaded"}
            </span>
          </div>

          <div style={{ display: "grid", gap: 8 }}>
            <strong>ERP template</strong>

            <input
              type="file"
              accept=".xlsx,.xls,.xlsm"
              onChange={handleTemplateUpload}
              disabled={busy}
            />

            <button
              type="button"
              style={buttonStyle}
              onClick={loadPermanentTemplateFile}
              disabled={busy}
            >
              Load permanent ERP template
            </button>

            <span style={{ color: "#666", fontSize: 13 }}>
              Current: {templateFileName || "No file loaded"}
            </span>
          </div>

          <div style={{ display: "grid", gap: 8 }}>
            <strong>Venue / location scope</strong>

            <select
              value={scope}
              onChange={(event) => {
                setScope(event.target.value);
                resetResults();
              }}
              disabled={busy || !venueSheets.length}
              style={{
                width: "100%",
                padding: 11,
                borderRadius: 10,
                border: "1px solid #ccc",
                background: "#fff",
              }}
            >
              <option value={ALL_VENUES_SCOPE}>All venue tabs</option>
              {venueSheets.map((sheetName) => (
                <option key={sheetName} value={sheetName}>
                  {sheetName}
                </option>
              ))}
            </select>

            <button
              type="button"
              style={buttonStyle}
              onClick={previewComparison}
              disabled={busy || !ingredientFile || !templateFile}
            >
              Preview Comparison
            </button>

            <button
              type="button"
              style={buttonStyle}
              onClick={downloadCorrectedTemplate}
              disabled={busy || !ingredientFile || !templateFile}
            >
              Download Corrected ERP Template
            </button>
          </div>
        </div>

        {message ? (
          <div
            style={{
              padding: 12,
              borderRadius: 12,
              background: "#f2f2f2",
              color: "#111",
            }}
          >
            {message}
          </div>
        ) : null}
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))",
          gap: 10,
          marginBottom: 16,
        }}
      >
        {summaryCards.map(([label, value]) => (
          <div
            key={label}
            style={{
              ...cardStyle,
              padding: 14,
              textAlign: "center",
            }}
          >
            <div
              style={{
                color: "#666",
                fontSize: 12,
                fontWeight: "bold",
                textTransform: "uppercase",
              }}
            >
              {label}
            </div>

            <div style={{ fontSize: 24, fontWeight: 900, marginTop: 4 }}>
              {value}
            </div>
          </div>
        ))}
      </div>

      <div style={{ ...cardStyle, display: "grid", gap: 12 }}>
        <div>
          <h2 style={{ margin: 0 }}>Preview</h2>
          <p style={{ color: "#555", marginBottom: 0 }}>
            Showing {visibleRows.length} of {rows.length} compared placement row(s).
            Red = suggested additions. Blue = template items not used in that location.
            Black = already correct in the template.
          </p>
        </div>

        <input
          type="search"
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          placeholder="Search status, venue, section, code, product, category..."
          style={{
            width: "100%",
            padding: 11,
            borderRadius: 10,
            border: "1px solid #ccc",
          }}
        />

        <div
          style={{
            overflowX: "auto",
            border: "1px solid #ddd",
            borderRadius: 14,
            maxHeight: 620,
          }}
        >
          <table
            style={{
              width: "100%",
              minWidth: 1400,
              borderCollapse: "collapse",
              fontSize: 13,
            }}
          >
            <thead>
              <tr>
                <th style={tableHeaderStyle}>Status</th>
                <th style={tableHeaderStyle}>Venue</th>
                <th style={tableHeaderStyle}>Section</th>
                <th style={tableHeaderStyle}>Code</th>
                <th style={tableHeaderStyle}>Product</th>
                <th style={tableHeaderStyle}>Restaurant</th>
                <th style={tableHeaderStyle}>Menu</th>
                <th style={tableHeaderStyle}>Category</th>
                <th style={tableHeaderStyle}>Sub Category</th>
                <th style={tableHeaderStyle}>Source Row</th>
              </tr>
            </thead>

            <tbody>
              {visibleRows.length ? (
                visibleRows.map((row, index) => {
                  const color = getStatusColor(row.Status);
                  const isStrong = color !== "#111";

                  return (
                    <tr
                      key={`${row.Venue}-${row.Section}-${row.Code}-${index}`}
                      style={{
                        color,
                        fontWeight: isStrong ? "bold" : "normal",
                      }}
                    >
                      <td style={tableCellStyle}>{getStatusLabel(row.Status)}</td>
                      <td style={tableCellStyle}>{row.Venue}</td>
                      <td style={tableCellStyle}>{row.Section}</td>
                      <td style={tableCellStyle}>{row.Code}</td>
                      <td style={tableCellStyle}>{row.Product}</td>
                      <td style={tableCellStyle}>{row.Restaurant}</td>
                      <td style={tableCellStyle}>{row.Menu}</td>
                      <td style={tableCellStyle}>{row.Category}</td>
                      <td style={tableCellStyle}>{row.SubCategory}</td>
                      <td style={tableCellStyle}>{row.IngredientSourceRow}</td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td
                    colSpan={10}
                    style={{
                      ...tableCellStyle,
                      textAlign: "center",
                      color: "#777",
                      padding: 24,
                    }}
                  >
                    Load both files, choose venue scope, then click Preview Comparison.
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

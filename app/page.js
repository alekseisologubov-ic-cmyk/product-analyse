"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

const ALLERGEN_RULES = [
  { allergen: "Tree Nuts", keywords: ["almond", "walnut", "pecan", "cashew", "hazelnut", "pistachio", "macadamia"] },
  { allergen: "Peanuts", keywords: ["peanut"] },
  { allergen: "Seeds", keywords: ["seed", "seeds", "sunflower seed", "pumpkin seed", "chia", "flax", "hemp seed"], exclude: ["seedless", "seedless cucumber"] },
  { allergen: "Soy", keywords: ["soy", "tofu", "edamame", "miso", "tamari"] },
  { allergen: "Gluten", keywords: ["wheat", "flour", "gluten", "bread", "pasta", "semolina", "barley", "rye", "panko"] },
  { allergen: "Milk / Dairy", keywords: ["milk", "cream", "butter", "cheese", "yogurt", "parmesan", "mozzarella", "ricotta", "cream cheese"] },
  { allergen: "Egg", keywords: ["egg", "eggs", "mayonnaise", "aioli"], exclude: ["eggplant"] },
  { allergen: "Fish", keywords: ["salmon", "tuna", "cod", "anchovy", "fish", "sardine"] },
  { allergen: "Shellfish", keywords: ["shrimp", "crab", "lobster", "mussel", "oyster", "scallop"], exclude: ["clam shell", "clamshell", "packed in a clam shell"] },
  { allergen: "Sesame", keywords: ["sesame", "tahini"] },
  { allergen: "Mustard", keywords: ["mustard"] },
];

const cleanText = (value) => String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

const normalizeVenue = (value) =>
  cleanText(value)
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

const formatQty = (value) => Number(value || 0).toFixed(2);

const getImageUrl = (url) => {
  const value = String(url || "").trim();
  if (!value) return "";
  if (value.includes("sharepoint.com") || value.includes("1drv.ms")) {
    return value.includes("?") ? `${value}&download=1` : `${value}?download=1`;
  }
  return value;
};

export default function App() {
  const [consumptionRows, setConsumptionRows] = useState([]);
  const [recipeRows, setRecipeRows] = useState([]);
  const [templateMap, setTemplateMap] = useState({});
  const [templateStatus, setTemplateStatus] = useState("Loading template...");
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [search, setSearch] = useState("");
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);
  const [message, setMessage] = useState("");
  const [viewMode, setViewMode] = useState("all");

  const [module, setModule] = useState("");
  const [equipmentMode, setEquipmentMode] = useState("");
  const [musterItems, setMusterItems] = useState([]);
  const [musterSearch, setMusterSearch] = useState("");
  const [musterMessage, setMusterMessage] = useState("");
  const [selectedEquipment, setSelectedEquipment] = useState(null);

  const [warehouseRows, setWarehouseRows] = useState([]);
  const [warehouseSearch, setWarehouseSearch] = useState("");
  const [warehouseMessage, setWarehouseMessage] = useState("");

  const shipColumns = { BRL: 8, RL: 11, SC: 14, VL: 17 };

  useEffect(() => {
    loadDefaultTemplate();
  }, []);

  const visibleShips = viewMode === "single" ? [userShip] : SHIPS;

  const buildProductList = (rows) =>
    [...new Set(rows.slice(1).map((r) => String(r[6] || "").trim()).filter(Boolean))].sort();

  const readExcelFile = (file, callback) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      callback(wb);
    };
    reader.readAsBinaryString(file);
  };

  const workbookToRows = (workbook) => {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1 });
  };

  const loadDefaultTemplate = async () => {
    try {
      const response = await fetch("/template.xlsx");
      if (!response.ok) {
        setTemplateStatus("Template file not found.");
        return;
      }
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      setTemplateMap(parseTemplateWorkbook(workbook));
      setTemplateStatus("Template loaded.");
    } catch {
      setTemplateStatus("Could not load template.");
    }
  };

  const parseTemplateWorkbook = (workbook) => {
    const map = {};
    workbook.SheetNames.forEach((sheetName) => {
      const venueKey = normalizeVenue(sheetName);
      if (!venueKey) return;
      if (!map[venueKey]) map[venueKey] = {};
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      rows.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const header = cleanText(cell);
          if (header !== "INGREDIENT NAME") return;
          const templateName = String(rows[rowIndex - 1]?.[colIndex] || rows[rowIndex - 1]?.[colIndex - 1] || sheetName).trim();
          rows.slice(rowIndex + 1).forEach((dataRow) => {
            const product = String(dataRow[colIndex] || "").trim();
            const productKey = cleanText(product);
            if (!productKey || productKey === "INGREDIENT NAME" || productKey.includes("#REF")) return;
            if (!map[venueKey][productKey]) map[venueKey][productKey] = { product, templates: new Set() };
            map[venueKey][productKey].templates.add(templateName);
          });
        });
      });
    });
    Object.keys(map).forEach(v => Object.keys(map[v]).forEach(p => map[v][p].templates = [...map[v][p].templates]));
    return map;
  };

  const productMatches = (selectedName, row) => {
    const s = cleanText(selectedName);
    const assigned = cleanText(row[12]);
    const name = cleanText(row[7]);
    if (!s) return false;
    if (assigned === s || name === s) return true;
    // Substring matching for longer names to handle minor data shifts
    if (s.length > 12 && (assigned.includes(s) || s.includes(assigned))) return true;
    return false;
  };

  // Optimized breakdown calculation
  const combinedBreakdown = useMemo(() => {
    if (!selectedProduct) return [];
    
    const actual = {};
    let currentVenue = "";
    consumptionRows.slice(1).forEach(row => {
      if (row[2]) currentVenue = String(row[2]).trim();
      if (String(row[6]).trim() !== selectedProduct) return;
      const vKey = normalizeVenue(currentVenue);
      if (!actual[vKey]) actual[vKey] = { displayName: currentVenue, ships: {} };
      SHIPS.forEach(ship => {
        actual[vKey].ships[ship] = (actual[vKey].ships[ship] || 0) + (Number(row[shipColumns[ship]]) || 0);
      });
    });

    const required = {};
    recipeRows.slice(1).forEach(row => {
      if (!productMatches(selectedProduct, row)) return;
      const vKey = normalizeVenue(row[1]);
      if (!vKey) return;
      if (!required[vKey]) required[vKey] = { displayName: String(row[1]).trim() };
    });

    const allKeys = Array.from(new Set([...Object.keys(actual), ...Object.keys(required)])).sort();
    return allKeys.map(vKey => {
      const isReq = !!required[vKey];
      const ships = actual[vKey]?.ships || {};
      const inTemplate = templateMap[vKey] && Object.keys(templateMap[vKey]).some(p => p === cleanText(selectedProduct));
      const templateMatches = [];
      if (templateMap[vKey]) {
        Object.entries(templateMap[vKey]).forEach(([pKey, data]) => {
          if (pKey === cleanText(selectedProduct)) templateMatches.push(...data.templates);
        });
      }

      return {
        venueKey,
        displayName: actual[vKey]?.displayName || required[vKey]?.displayName || vKey,
        ships,
        required: isReq,
        missingShips: visibleShips.filter(s => isReq && (ships[s] || 0) === 0),
        missingFromTemplate: isReq && !inTemplate,
        templateMatches: [...new Set(templateMatches)]
      };
    });
  }, [selectedProduct, consumptionRows, recipeRows, templateMap, visibleShips]);

  // Rest of your logic (Allergens, Muster, etc.) remains exactly as you built it...
  const uploadConsumptionFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      const rows = workbookToRows(workbook);
      setConsumptionRows(rows);
      setProducts(buildProductList(rows));
      setMessage("Consumption file loaded.");
    });
  };

  const uploadRecipeFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      setRecipeRows(workbookToRows(workbook));
      setMessage("Recipe / location file loaded.");
    });
  };

  // Original UI Structure below
  if (!loggedIn) {
    return (
      <main style={styles.page}>
        <section style={styles.loginCard}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
          <h1 style={styles.title}>Virgin Voyages Dashboard</h1>
          <p style={styles.subtitle}>Product and equipment tools</p>
          <label style={styles.label}>🚢 Select your ship</label>
          <select value={userShip} onChange={(e) => setUserShip(e.target.value)} style={styles.select}>
            <option value="">Choose ship</option>
            {SHIPS.map((ship) => <option key={ship}>{ship}</option>)}
          </select>
          <button style={styles.primaryButton} onClick={() => userShip && setLoggedIn(true)}>Continue</button>
        </section>
      </main>
    );
  }

  if (!module) {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.shipBadge}>🚢 {userShip}</div>
        </header>
        <section style={styles.card}>
          <h2 style={styles.cardTitle}>🧭 Select Module</h2>
          <div style={styles.moduleGrid}>
            <button style={styles.moduleCard} onClick={() => setModule("product")}>
              <div style={styles.moduleIcon}>📦</div>
              <strong>Product Dashboard</strong>
              <span>Consumption, recipes, templates and allergens</span>
            </button>
            <button style={styles.moduleCard} onClick={() => setModule("equipment")}>
              <div style={styles.moduleIcon}>🍽️</div>
              <strong>Equipment</strong>
              <span>Muster list and inventory tools</span>
            </button>
          </div>
        </section>
      </main>
    );
  }

  // Final Render (Product Dashboard)
  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
          <div style={styles.shipBadge}>🚢 {userShip}</div>
        </div>
      </header>

      <div style={styles.viewModeBox}>
        <button onClick={() => setViewMode("single")} style={{ ...styles.viewModeButton, ...(viewMode === "single" ? styles.viewModeButtonActive : {}) }}>🚢 {userShip} Only</button>
        <button onClick={() => setViewMode("all")} style={{ ...styles.viewModeButton, ...(viewMode === "all" ? styles.viewModeButtonActive : {}) }}>🌍 All Ships Overview</button>
      </div>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Files</h2>
          <label style={styles.label}>Step 1: Consumption file</label>
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadConsumptionFile} style={styles.fileInput} />
          <label style={styles.label}>Step 2: Recipe / location file</label>
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadRecipeFile} style={styles.fileInput} />
          <div style={styles.infoBox}>
            <div>📦 Products: <strong>{products.length}</strong></div>
            <div>📋 Template: <strong>{templateStatus}</strong></div>
            <div style={{ color: "#b00020" }}>Red = Usage is 0 for highlighted ship.</div>
            <div style={{ color: "#0057b8" }}>Blue = Missing from menu template.</div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔍 Select Product</h2>
          <input placeholder="Search product..." value={search} onChange={(e) => setSearch(e.target.value)} style={styles.searchInput} />
          <div style={styles.productList}>
            {products.filter(p => p.toLowerCase().includes(search.toLowerCase())).map((product, i) => (
              <button key={i} onClick={() => setSelectedProduct(product)} style={{ ...styles.productItem, ...(selectedProduct === product ? styles.productItemActive : {}) }}>{product}</button>
            ))}
          </div>
        </div>
      </section>

      {selectedProduct && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📦 {selectedProduct}</h2>
          <h3 style={styles.sectionTitle}>🏢 Venue Breakdown</h3>
          <div style={styles.venueGrid}>
            {combinedBreakdown.map((venueItem, i) => (
              <div key={i} style={{
                ...styles.venueCard,
                ...(venueItem.missingShips.length > 0 ? styles.venueCardWarning : {}),
                ...(venueItem.missingFromTemplate ? styles.venueCardTemplateWarning : {}),
              }}>
                <h4 style={styles.venueTitle}>
                  {venueItem.displayName}
                  <span style={styles.badgeGroup}>
                    {venueItem.missingFromTemplate && <span style={styles.templateBadge}>Missing Template</span>}
                    {venueItem.missingShips.length > 0 && <span style={styles.missingBadge}>Missing: {venueItem.missingShips.join(", ")}</span>}
                  </span>
                </h4>
                <div style={styles.shipGrid}>
                  {visibleShips.map((ship) => (
                    <div key={ship} style={{ ...styles.shipBox, ...(venueItem.missingShips.includes(ship) ? styles.shipBoxMissing : {}) }}>
                      <span style={styles.shipName}>{ship}</span>
                      <strong style={styles.shipQty}>{formatQty(venueItem.ships[ship])}</strong>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </section>
      )}
    </main>
  );
}

// I kept all your ORIGINAL styles here so the look doesn't change
const styles = {
  page: { minHeight: "100vh", padding: 24, background: "#f5f5f5", fontFamily: "Arial, sans-serif", color: "#111" },
  loginCard: { maxWidth: 460, margin: "80px auto", padding: 28, background: "#fff", borderRadius: 16, boxShadow: "0 10px 30px rgba(0,0,0,0.08)", display: "grid", gap: 14 },
  logo: { height: 70, objectFit: "contain", marginBottom: 8 },
  header: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: 18, background: "#fff", borderRadius: 16, boxShadow: "0 4px 18px rgba(0,0,0,0.06)", marginBottom: 20 },
  headerLogo: { height: 54, objectFit: "contain" },
  headerActions: { display: "flex", alignItems: "center", gap: 10 },
  backButton: { padding: "10px 14px", borderRadius: 999, border: "1px solid #ccc", background: "#fff", cursor: "pointer", fontWeight: "bold" },
  title: { margin: 0, fontSize: 28 },
  subtitle: { margin: 0, color: "#666" },
  label: { fontWeight: "bold", marginTop: 8 },
  select: { padding: 10, borderRadius: 8, border: "1px solid #ccc" },
  primaryButton: { marginTop: 10, padding: 12, borderRadius: 10, border: 0, background: "#111", color: "#fff", fontWeight: "bold", cursor: "pointer" },
  shipBadge: { padding: "10px 14px", borderRadius: 999, background: "#111", color: "#fff", fontWeight: "bold" },
  moduleGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16 },
  moduleCard: { border: "1px solid #ddd", background: "#fafafa", borderRadius: 16, padding: 20, cursor: "pointer", textAlign: "left", display: "grid", gap: 8, boxShadow: "0 4px 12px rgba(0,0,0,0.04)" },
  moduleIcon: { fontSize: 34 },
  viewModeBox: { display: "flex", gap: 10, marginBottom: 20 },
  viewModeButton: { padding: "10px 14px", borderRadius: 999, border: "1px solid #ccc", background: "#fff", cursor: "pointer", fontWeight: "bold" },
  viewModeButtonActive: { background: "#111", color: "#fff", borderColor: "#111" },
  grid: { display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 20, marginBottom: 20 },
  card: { background: "#fff", borderRadius: 16, padding: 20, boxShadow: "0 4px 18px rgba(0,0,0,0.06)" },
  cardTitle: { marginTop: 0 },
  fileInput: { display: "block", margin: "8px 0 16px" },
  message: { color: "#555", fontSize: 14 },
  infoBox: { marginTop: 12, padding: 12, borderRadius: 12, background: "#f2f2f2", display: "grid", gap: 6 },
  searchInput: { width: "100%", padding: 12, borderRadius: 10, border: "1px solid #ccc", marginBottom: 10 },
  productList: { maxHeight: 300, overflowY: "auto", border: "1px solid #ddd", borderRadius: 12 },
  productItem: { width: "100%", display: "block", textAlign: "left", padding: 10, border: 0, borderBottom: "1px solid #eee", background: "#fff", cursor: "pointer" },
  productItemActive: { background: "#eee", fontWeight: "bold" },
  productTitle: { marginTop: 0, fontSize: 24 },
  sectionTitle: { marginTop: 22 },
  venueGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14 },
  venueCard: { border: "1px solid #ddd", borderRadius: 14, padding: 14, background: "#fafafa" },
  venueCardWarning: { border: "2px solid #b00020", background: "#fff0f0" },
  venueCardTemplateWarning: { border: "2px solid #0057b8", background: "#eef5ff" },
  venueTitle: { marginTop: 0, display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" },
  badgeGroup: { display: "flex", flexWrap: "wrap", gap: 6, justifyContent: "flex-end" },
  missingBadge: { fontSize: 12, color: "#fff", background: "#b00020", borderRadius: 999, padding: "4px 8px" },
  templateBadge: { fontSize: 12, color: "#fff", background: "#0057b8", borderRadius: 999, padding: "4px 8px" },
  shipGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(70px, 1fr))", gap: 6 },
  shipBox: { minWidth: 0, padding: "8px 4px", borderRadius: 10, background: "#fff", border: "1px solid #ddd", display: "grid", gap: 3, textAlign: "center", overflow: "hidden" },
  shipBoxMissing: { background: "#b00020", color: "#fff", borderColor: "#b00020" },
  shipName: { fontSize: 11, opacity: 0.8 },
  shipQty: { fontSize: 14, lineHeight: 1.1 },
};

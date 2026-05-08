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

const cleanText = (value) =>
  String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

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

  // Optimized Search Filter for Products (Prevents UI Lag)
  const filteredProducts = useMemo(() => {
    const query = search.toLowerCase().trim();
    if (!query) return products.slice(0, 100);
    return products.filter(p => p.toLowerCase().includes(query)).slice(0, 100);
  }, [products, search]);

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
          const templateName = String(rows[rowIndex - 1]?.[colIndex] || rows[rowIndex - 1]?.[colIndex - 1] || sheetName || "Template").trim();
          rows.slice(rowIndex + 1).forEach((dataRow) => {
            const product = String(dataRow[colIndex] || "").trim();
            if (!product) return;
            const productKey = cleanText(product);
            if (productKey === "INGREDIENT NAME" || productKey === "CODE" || productKey === "UM" || productKey.includes("#REF")) return;
            if (!map[venueKey][productKey]) map[venueKey][productKey] = { product, templates: new Set() };
            map[venueKey][productKey].templates.add(templateName);
          });
        });
      });
    });
    Object.keys(map).forEach((vK) => {
      Object.keys(map[vK]).forEach((pK) => {
        map[vK][pK].templates = [...map[vK][pK].templates];
      });
    });
    return map;
  };

  const uploadConsumptionFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      const rows = workbookToRows(workbook);
      setConsumptionRows(rows);
      setProducts([...new Set(rows.slice(1).map((r) => String(r[6] || "").trim()).filter(Boolean))].sort());
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

  const productMatches = (selectedProductName, row) => {
    const selected = cleanText(selectedProductName);
    const assignedProduct = cleanText(row[12]);
    const productName = cleanText(row[7]);
    if (!selected) return false;
    if (assignedProduct === selected || productName === selected) return true;
    if (assignedProduct.length > 12 && (selected.includes(assignedProduct) || assignedProduct.includes(selected))) return true;
    if (productName.length > 12 && (selected.includes(productName) || productName.includes(selected))) return true;
    return false;
  };

  // Heavy Breakdown Logic wrapped in useMemo for Data Handling
  const combinedBreakdown = useMemo(() => {
    if (!selectedProduct) return [];
    
    const actual = {};
    let currentVenue = "";
    consumptionRows.slice(1).forEach((row) => {
      if (row[2]) currentVenue = String(row[2]).trim();
      const venueKey = normalizeVenue(currentVenue);
      const productName = String(row[6] || "").trim();
      if (productName !== selectedProduct) return;
      if (!actual[venueKey]) actual[venueKey] = { displayName: currentVenue, ships: {} };
      SHIPS.forEach((ship) => {
        const qty = Number(row[shipColumns[ship]]) || 0;
        actual[venueKey].ships[ship] = (actual[venueKey].ships[ship] || 0) + qty;
      });
    });

    const required = {};
    recipeRows.slice(1).forEach((row) => {
      if (!productMatches(selectedProduct, row)) return;
      const venueRaw = String(row[1] || "").trim();
      const venueKey = normalizeVenue(venueRaw);
      if (!venueKey) return;
      if (!required[venueKey]) required[venueKey] = { displayName: venueRaw || venueKey };
    });

    const allVenueKeys = Array.from(new Set([...Object.keys(actual), ...Object.keys(required)])).sort();
    return allVenueKeys.map((venueKey) => {
      const actualVenue = actual[venueKey];
      const requiredVenue = required[venueKey];
      const ships = {};
      SHIPS.forEach((ship) => { ships[ship] = actualVenue?.ships?.[ship] || 0; });
      
      const venueTemplates = templateMap[venueKey] || {};
      const templateMatches = [];
      Object.entries(venueTemplates).forEach(([tKey, data]) => {
        if (tKey === cleanText(selectedProduct)) templateMatches.push(...data.templates);
      });

      const requiredByRecipe = !!requiredVenue;
      const inTemplate = templateMatches.length > 0;

      return {
        venueKey,
        displayName: actualVenue?.displayName || requiredVenue?.displayName || venueKey,
        ships,
        required: requiredByRecipe,
        missingShips: visibleShips.filter((ship) => requiredByRecipe && (ships[ship] || 0) === 0),
        missingFromTemplate: requiredByRecipe && !inTemplate,
        templateMatches: [...new Set(templateMatches)],
      };
    });
  }, [selectedProduct, consumptionRows, recipeRows, templateMap, visibleShips]);

  const totalConsumption = useMemo(() => {
    const totals = { BRL: 0, RL: 0, SC: 0, VL: 0 };
    combinedBreakdown.forEach((v) => {
      visibleShips.forEach((s) => { totals[s] += Number(v.ships[s] || 0); });
    });
    return { totals, allShips: visibleShips.reduce((sum, s) => sum + totals[s], 0) };
  }, [combinedBreakdown, visibleShips]);

  // UI remains identical to your original structure
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
        <button onClick={() => setViewMode("all")} style={{ ...styles.viewModeButton, ...(viewMode === "all" ? styles.viewModeButtonActive : {}) }}>🌍 All Ships</button>
      </div>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Data</h2>
          <input type="file" onChange={uploadConsumptionFile} style={styles.fileInput} />
          <input type="file" onChange={uploadRecipeFile} style={styles.fileInput} />
          <div style={styles.infoBox}>
            <div>📦 Loaded Items: {products.length}</div>
            <div style={{ color: "#b00020" }}>Red: 0 Usage Alert</div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔍 Product Search</h2>
          <input placeholder="Search items..." value={search} onChange={(e) => setSearch(e.target.value)} style={styles.searchInput} />
          <div style={styles.productList}>
            {filteredProducts.map((p, i) => (
              <button key={i} onClick={() => setSelectedProduct(p)} style={{ ...styles.productItem, ...(selectedProduct === p ? styles.productItemActive : {}) }}>{p}</button>
            ))}
          </div>
        </div>
      </section>

      {selectedProduct && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📦 {selectedProduct}</h2>
          <div style={styles.totalBox}>
            <div style={styles.totalMain}>Total across ships: {formatQty(totalConsumption.allShips)}</div>
          </div>
          <div style={styles.venueGrid}>
            {combinedBreakdown.map((v, i) => (
              <div key={i} style={{ ...styles.venueCard, ...(v.missingShips.length > 0 ? styles.venueCardWarning : {}), ...(v.missingFromTemplate ? styles.venueCardTemplateWarning : {}) }}>
                <h4 style={styles.venueTitle}>{v.displayName}</h4>
                <div style={styles.shipGrid}>
                  {visibleShips.map(s => (
                    <div key={s} style={{ ...styles.shipBox, ...(v.missingShips.includes(s) ? styles.shipBoxMissing : {}) }}>
                      <span style={styles.shipName}>{s}</span>
                      <strong style={styles.shipQty}>{formatQty(v.ships[s])}</strong>
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
  totalBox: { background: "#111", color: "#fff", borderRadius: 14, padding: 16, marginBottom: 18 },
  totalMain: { fontSize: 20, fontWeight: "bold" },
  venueGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14 },
  venueCard: { border: "1px solid #ddd", borderRadius: 14, padding: 14, background: "#fafafa" },
  venueCardWarning: { border: "2px solid #b00020", background: "#fff0f0" },
  venueCardTemplateWarning: { border: "2px solid #0057b8", background: "#eef5ff" },
  venueTitle: { marginTop: 0, display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" },
  shipGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(70px, 1fr))", gap: 6 },
  shipBox: { minWidth: 0, padding: "8px 4px", borderRadius: 10, background: "#fff", border: "1px solid #ddd", display: "grid", gap: 3, textAlign: "center" },
  shipBoxMissing: { background: "#b00020", color: "#fff", borderColor: "#b00020" },
  shipName: { fontSize: 11, opacity: 0.8 },
  shipQty: { fontSize: 14, fontWeight: "bold" },
};

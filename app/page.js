"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

/** * CONFIGURATION & CONSTANTS
 */
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

/**
 * UTILS & PARSING ENGINE
 */
const cleanText = (value) => String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

const normalizeVenue = (value) =>
  cleanText(value)
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\s*-\s*VV$/g, "")
    .replace(/\s*VV$/g, "")
    .replace(/\b(THE|SCL|VAL|RES|BRL|ROJO|ARIYA|ONLY)\b/g, "")
    .replace(/\bMANNOR\b/g, "MANOR")
    .replace(/\s+/g, " ")
    .trim();

const getImageUrl = (url) => {
  const value = String(url || "").trim();
  if (!value) return "";
  if (value.includes("sharepoint.com") || value.includes("1drv.ms")) {
    return value.includes("?") ? `${value}&download=1` : `${value}?download=1`;
  }
  return value;
};

export default function App() {
  // State Management
  const [consumptionRows, setConsumptionRows] = useState([]);
  const [recipeRows, setRecipeRows] = useState([]);
  const [templateMap, setTemplateMap] = useState({});
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [search, setSearch] = useState("");
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);
  const [viewMode, setViewMode] = useState("all");

  // Equipment/Warehouse State
  const [module, setModule] = useState("");
  const [equipmentMode, setEquipmentMode] = useState("");
  const [musterItems, setMusterItems] = useState([]);
  const [musterSearch, setMusterSearch] = useState("");
  const [selectedEquipment, setSelectedEquipment] = useState(null);
  const [warehouseRows, setWarehouseRows] = useState([]);
  const [warehouseSearch, setWarehouseSearch] = useState("");

  const shipColumns = { BRL: 8, RL: 11, SC: 14, VL: 17 };
  const visibleShips = viewMode === "single" ? [userShip] : SHIPS;

  useEffect(() => {
    loadDefaultTemplate();
  }, []);

  /**
   * FILE LOADERS
   */
  const loadDefaultTemplate = async () => {
    try {
      const response = await fetch("/template.xlsx");
      if (response.ok) {
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: "array" });
        setTemplateMap(parseTemplateWorkbook(workbook));
      }
    } catch (e) { console.error("Template load failed", e); }
  };

  const handleFileUpload = (e, callback) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      callback(wb);
    };
    reader.readAsBinaryString(file);
  };

  /**
   * PARSING LOGIC
   */
  const parseTemplateWorkbook = (workbook) => {
    const map = {};
    workbook.SheetNames.forEach((sheetName) => {
      const venueKey = normalizeVenue(sheetName);
      if (!venueKey) return;
      if (!map[venueKey]) map[venueKey] = {};

      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      rows.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (cleanText(cell) !== "INGREDIENT NAME") return;
          const templateName = String(rows[rowIndex - 1]?.[colIndex] || sheetName).trim();
          rows.slice(rowIndex + 1).forEach((dataRow) => {
            const product = String(dataRow[colIndex] || "").trim();
            const pKey = cleanText(product);
            if (!pKey || pKey.includes("#REF") || pKey === "CODE") return;
            if (!map[venueKey][pKey]) map[venueKey][pKey] = { product, templates: new Set() };
            map[venueKey][pKey].templates.add(templateName);
          });
        });
      });
    });
    // Convert Sets to Arrays
    Object.keys(map).forEach(v => Object.keys(map[v]).forEach(p => map[v][p].templates = [...map[v][p].templates]));
    return map;
  };

  /**
   * CORE COMPUTATIONS (useMemo for performance)
   */
  const filteredProducts = useMemo(() => {
    const query = search.toLowerCase();
    return products.filter(p => p.toLowerCase().includes(query)).slice(0, 100); // Caps UI render for speed
  }, [products, search]);

  const productMatches = (selectedName, rowValue) => {
    const s = cleanText(selectedName);
    const r = cleanText(rowValue);
    if (!s || !r) return false;
    return r === s || (r.length > 12 && (s.includes(r) || r.includes(s)));
  };

  const combinedBreakdown = useMemo(() => {
    if (!selectedProduct) return [];
    
    const actual = {};
    let currentVenue = "";
    
    // Calculate Consumption
    consumptionRows.slice(1).forEach(row => {
      if (row[2]) currentVenue = String(row[2]).trim();
      if (String(row[6]).trim() !== selectedProduct) return;
      const vKey = normalizeVenue(currentVenue);
      if (!actual[vKey]) actual[vKey] = { displayName: currentVenue, ships: {} };
      SHIPS.forEach(ship => {
        actual[vKey].ships[ship] = (actual[vKey].ships[ship] || 0) + (Number(row[shipColumns[ship]]) || 0);
      });
    });

    // Calculate Requirements
    const required = {};
    recipeRows.slice(1).forEach(row => {
      if (!productMatches(selectedProduct, row[12]) && !productMatches(selectedProduct, row[7])) return;
      const vRaw = String(row[1] || "").trim();
      const vKey = normalizeVenue(vRaw);
      if (!vKey) return;
      required[vKey] = true;
    });

    const allKeys = [...new Set([...Object.keys(actual), ...Object.keys(required)])].sort();
    return allKeys.map(k => {
      const isReq = !!required[k];
      const ships = actual[k]?.ships || {};
      const templateMatches = templateMap[k] ? Object.entries(templateMap[k])
        .filter(([pKey]) => productMatches(selectedProduct, pKey))
        .flatMap(([, d]) => d.templates) : [];

      return {
        key: k,
        name: actual[k]?.displayName || k,
        ships,
        required: isReq,
        missingShips: visibleShips.filter(s => isReq && (ships[s] || 0) === 0),
        inTemplate: templateMatches.length > 0,
        templateMatches: [...new Set(templateMatches)]
      };
    });
  }, [selectedProduct, consumptionRows, recipeRows, templateMap, visibleShips]);

  /**
   * UI SUB-COMPONENTS
   */
  const renderHeader = (title, backAction) => (
    <header style={styles.header}>
      <div style={styles.headerActions}>
        {backAction && <button style={styles.backButton} onClick={backAction}>← Back</button>}
        <img src="/virgin-logo.png" alt="Logo" style={styles.headerLogo} />
      </div>
      <div style={styles.headerActions}>
         <h2 style={{margin:0, fontSize: 18}}>{title}</h2>
         <div style={styles.shipBadge}>🚢 {userShip}</div>
      </div>
    </header>
  );

  /**
   * CONDITIONAL RENDERING (ROUTING)
   */
  if (!loggedIn) {
    return (
      <main style={styles.page}>
        <section style={styles.loginCard}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
          <h1 style={styles.title}>F&B Dashboard</h1>
          <label style={styles.label}>Select Ship</label>
          <select value={userShip} onChange={(e) => setUserShip(e.target.value)} style={styles.select}>
            <option value="">Choose...</option>
            {SHIPS.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <button style={styles.primaryButton} onClick={() => userShip && setLoggedIn(true)}>Enter System</button>
        </section>
      </main>
    );
  }

  // Main Module Selection
  if (!module) {
    return (
      <main style={styles.page}>
        {renderHeader("Command Center")}
        <div style={styles.moduleGrid}>
          <button style={styles.moduleCard} onClick={() => setModule("product")}>
            <span style={styles.moduleIcon}>📦</span>
            <strong>Product Intelligence</strong>
            <p>Consumption vs Recipes</p>
          </button>
          <button style={styles.moduleCard} onClick={() => setModule("equipment")}>
            <span style={styles.moduleIcon}>🍽️</span>
            <strong>Equipment & Muster</strong>
            <p>Inventory & Visual Guides</p>
          </button>
        </div>
      </main>
    );
  }

  // Equipment Module
  if (module === "equipment") {
    const warehouseItems = warehouseRows.slice(1).map(row => ({
      code: String(row[0] || ""),
      name: String(row[1] || ""),
      par: Number(row[6] || 0),
      onHand: Number(row[7] || 0),
      future: Number(row[12] || 0),
      suggested: Math.max(Number(row[6] || 0) - Number(row[7] || 0) - Number(row[12] || 0), 0)
    })).filter(i => (i.name + i.code).toLowerCase().includes(warehouseSearch.toLowerCase()));

    return (
        <main style={styles.page}>
            {renderHeader("Equipment", () => setModule(""))}
            <div style={styles.viewModeBox}>
                <button onClick={() => setEquipmentMode("muster")} style={{...styles.viewModeButton, ...(equipmentMode === "muster" ? styles.viewModeButtonActive : {})}}>📋 Muster List</button>
                <button onClick={() => setEquipmentMode("warehouse")} style={{...styles.viewModeButton, ...(equipmentMode === "warehouse" ? styles.viewModeButtonActive : {})}}>🏬 Warehouse</button>
            </div>

            {equipmentMode === "warehouse" ? (
                <section style={styles.card}>
                    <input type="file" onChange={(e) => handleFileUpload(e, wb => setWarehouseRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1})))} />
                    <input placeholder="Search items..." style={styles.searchInput} value={warehouseSearch} onChange={e => setWarehouseSearch(e.target.value)} />
                    <div style={styles.equipmentGrid}>
                        {warehouseItems.map((item, i) => (
                            <div key={i} style={{...styles.equipmentCard, borderLeft: item.suggested > 0 ? '5px solid #b00020' : '5px solid #2e7d32'}}>
                                <strong>{item.name}</strong>
                                <span>Code: {item.code}</span>
                                <div style={{display:'flex', justifyContent:'space-between', marginTop: 10}}>
                                    <small>On Hand: {item.onHand}</small>
                                    <small>Par: {item.par}</small>
                                </div>
                                {item.suggested > 0 && <div style={styles.suggestedOrderBad}>Order: {item.suggested}</div>}
                            </div>
                        ))}
                    </div>
                </section>
            ) : (
                <section style={styles.card}>
                    <input type="file" onChange={(e) => handleFileUpload(e, wb => {
                        const items = [];
                        wb.SheetNames.forEach(sn => {
                            const data = XLSX.utils.sheet_to_json(wb.Sheets[sn], {header: 1});
                            data.slice(1).forEach(r => items.push({sheet: sn, cat: r[2], code: r[3], name: r[4], img: r[7]}));
                        });
                        setMusterItems(items);
                    })} />
                    <div style={styles.equipmentGrid}>
                        {musterItems.filter(i => i.name?.toLowerCase().includes(musterSearch.toLowerCase())).map((item, i) => (
                            <div key={i} style={styles.equipmentCard} onClick={() => setSelectedEquipment(item)}>
                                {item.img && <img src={getImageUrl(item.img)} style={styles.equipmentImage} alt="item" />}
                                <strong>{item.name}</strong>
                                <small>{item.cat}</small>
                            </div>
                        ))}
                    </div>
                </section>
            )}
        </main>
    );
  }

  // Product Dashboard
  return (
    <main style={styles.page}>
      {renderHeader("Product Intelligence", () => setModule(""))}
      
      <div style={styles.grid}>
        <div style={styles.card}>
            <h3>Upload Data</h3>
            <div style={{display:'grid', gap: 10}}>
                <label>1. Consumption</label>
                <input type="file" onChange={(e) => handleFileUpload(e, wb => {
                    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
                    setConsumptionRows(rows);
                    setProducts([...new Set(rows.slice(1).map(r => String(r[6]).trim()).filter(Boolean))].sort());
                })} />
                <label>2. Recipes</label>
                <input type="file" onChange={(e) => handleFileUpload(e, wb => setRecipeRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1})))} />
            </div>
            
            <hr style={{margin: '20px 0', border: 0, borderTop: '1px solid #eee'}} />
            
            <input placeholder="Search products..." value={search} onChange={e => setSearch(e.target.value)} style={styles.searchInput} />
            <div style={styles.productList}>
                {filteredProducts.map(p => (
                    <button key={p} onClick={() => setSelectedProduct(p)} style={{...styles.productItem, background: selectedProduct === p ? '#f2f2f2' : 'transparent'}}>
                        {p}
                    </button>
                ))}
            </div>
        </div>

        <div style={styles.card}>
            {selectedProduct ? (
                <>
                    <h2>{selectedProduct}</h2>
                    <div style={styles.viewModeBox}>
                        <button onClick={() => setViewMode("all")} style={{...styles.viewModeButton, ...(viewMode === "all" ? styles.viewModeButtonActive : {})}}>All Ships</button>
                        <button onClick={() => setViewMode("single")} style={{...styles.viewModeButton, ...(viewMode === "single" ? styles.viewModeButtonActive : {})}}>{userShip} Only</button>
                    </div>

                    <div style={styles.venueGrid}>
                        {combinedBreakdown.map(v => (
                            <div key={v.key} style={{
                                ...styles.venueCard, 
                                borderTop: v.missingShips.length > 0 ? '4px solid #b00020' : (!v.inTemplate && v.required ? '4px solid #0057b8' : '1px solid #ddd')
                            }}>
                                <div style={{display:'flex', justifyContent:'space-between', alignItems:'flex-start'}}>
                                    <strong>{v.name}</strong>
                                    {v.required && <span style={{fontSize: 10, background: '#eee', padding: '2px 5px', borderRadius: 4}}>RECIPE REQ</span>}
                                </div>
                                
                                <div style={styles.shipGrid}>
                                    {visibleShips.map(s => (
                                        <div key={s} style={{...styles.shipBox, background: (v.required && (v.ships[s]||0) === 0) ? '#ffebee' : '#fff'}}>
                                            <small>{s}</small>
                                            <strong>{Number(v.ships[s]||0).toFixed(1)}</strong>
                                        </div>
                                    ))}
                                </div>
                                {!v.inTemplate && v.required && <div style={{color:'#0057b8', fontSize: 11, marginTop: 8}}>⚠️ Missing from Menu Template</div>}
                            </div>
                        ))}
                    </div>
                </>
            ) : (
                <div style={styles.emptyText}>Select a product to view discrepancy report.</div>
            )}
        </div>
      </div>
    </main>
  );
}

/**
 * UPDATED STYLES (Clean & Professional)
 */
const styles = {
  page: { minHeight: "100vh", padding: 20, background: "#f8f9fa", fontFamily: "Segoe UI, Roboto, Helvetica, Arial, sans-serif" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 20px", background: "#fff", borderRadius: 12, marginBottom: 20, boxShadow: "0 2px 10px rgba(0,0,0,0.05)" },
  headerLogo: { height: 40 },
  headerActions: { display: "flex", alignItems: "center", gap: 15 },
  loginCard: { maxWidth: 400, margin: "100px auto", padding: 30, background: "#fff", borderRadius: 16, boxShadow: "0 10px 40px rgba(0,0,0,0.1)", textAlign: "center" },
  logo: { height: 60, marginBottom: 20 },
  title: { fontSize: 24, fontWeight: "800", color: "#111", marginBottom: 20 },
  label: { display: "block", textAlign: "left", fontWeight: "bold", marginBottom: 5, fontSize: 14 },
  select: { width: "100%", padding: 12, borderRadius: 8, border: "1px solid #ddd", marginBottom: 20 },
  primaryButton: { width: "100%", padding: 14, background: "#111", color: "#fff", borderRadius: 8, fontWeight: "bold", border: "none", cursor: "pointer" },
  moduleGrid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, maxWidth: 800, margin: "40px auto" },
  moduleCard: { background: "#fff", padding: 30, borderRadius: 20, border: "1px solid #eee", cursor: "pointer", transition: "transform 0.2s", textAlign: "center" },
  moduleIcon: { fontSize: 40, display: "block", marginBottom: 10 },
  shipBadge: { background: "#111", color: "#fff", padding: "5px 12px", borderRadius: 20, fontSize: 12, fontWeight: "bold" },
  backButton: { background: "none", border: "none", cursor: "pointer", fontWeight: "bold", color: "#666" },
  grid: { display: "grid", gridTemplateColumns: "350px 1fr", gap: 20 },
  card: { background: "#fff", padding: 20, borderRadius: 16, boxShadow: "0 2px 12px rgba(0,0,0,0.04)" },
  searchInput: { width: "100%", padding: 12, borderRadius: 8, border: "1px solid #eee", marginBottom: 15, boxSizing: 'border-box' },
  productList: { maxHeight: "500px", overflowY: "auto", border: "1px solid #f0f0f0", borderRadius: 8 },
  productItem: { width: "100%", textAlign: "left", padding: "10px 15px", border: "none", borderBottom: "1px solid #f9f9f9", cursor: "pointer", fontSize: 13 },
  venueGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(280px, 1fr))", gap: 15, marginTop: 20 },
  venueCard: { background: "#fff", padding: 15, borderRadius: 12, border: "1px solid #eee" },
  shipGrid: { display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8, marginTop: 12 },
  shipBox: { textAlign: "center", padding: 5, borderRadius: 6, border: "1px solid #f0f0f0" },
  viewModeBox: { display: "flex", gap: 10, margin: "10px 0" },
  viewModeButton: { padding: "6px 15px", borderRadius: 20, border: "1px solid #ddd", background: "#fff", cursor: "pointer", fontSize: 12 },
  viewModeButtonActive: { background: "#111", color: "#fff", borderColor: "#111" },
  equipmentGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px, 1fr))", gap: 15, marginTop: 20 },
  equipmentCard: { background: "#fff", padding: 15, borderRadius: 12, border: "1px solid #eee", display: "flex", flexDirection: "column" },
  equipmentImage: { width: "100%", height: 120, objectFit: "cover", borderRadius: 8, marginBottom: 10 },
  suggestedOrderBad: { background: "#ffebee", color: "#b00020", padding: "5px", borderRadius: 4, textAlign: "center", marginTop: 10, fontWeight: "bold", fontSize: 12 },
  emptyText: { textAlign: "center", color: "#999", marginTop: 100 }
};

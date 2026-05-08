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
  if (value.startsWith("http")) {
      if (value.includes("sharepoint.com") || value.includes("1drv.ms")) {
        return value.includes("?") ? `${value}&download=1` : `${value}?download=1`;
      }
      return value;
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
      if (!rows.length) return;
      rows.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const header = cleanText(cell);
          if (header !== "INGREDIENT NAME") return;
          const templateName = String(rows[rowIndex - 1]?.[colIndex] || rows[rowIndex - 1]?.[colIndex - 1] || sheetName || "Template").trim();
          rows.slice(rowIndex + 1).forEach((dataRow) => {
            const product = String(dataRow[colIndex] || "").trim();
            if (!product) return;
            const productKey = cleanText(product);
            if (!productKey || productKey === "INGREDIENT NAME" || productKey === "CODE" || productKey === "UM" || productKey.includes("#REF")) return;
            if (!map[venueKey][productKey]) map[venueKey][productKey] = { product, templates: new Set() };
            map[venueKey][productKey].templates.add(templateName);
          });
        });
      });
    });
    Object.keys(map).forEach((venueKey) => {
      Object.keys(map[venueKey]).forEach((productKey) => {
        map[venueKey][productKey].templates = [...map[venueKey][productKey].templates];
      });
    });
    return map;
  };

  const parseMusterWorkbook = (workbook) => {
    const items = [];
    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      rows.slice(1).forEach((row) => {
        const category = String(row[2] || "").trim();
        const code = String(row[3] || "").trim();
        const name = String(row[4] || "").trim();
        const image = String(row[7] || "").trim();
        if (!category || !name) return;
        items.push({ sheetName, category, code, name, image });
      });
    });
    return items;
  };

  const parseWarehouseItems = () => {
    return warehouseRows
      .slice(1)
      .map((row) => {
        const code = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        const par = Number(row[6] || 0);
        const onHand = Number(row[7] || 0);
        const image = String(row[8] || "").trim(); 
        const future = Number(row[12] || 0);
        const suggested = Math.max(par - onHand - future, 0);
        return { code, name, par, onHand, future, suggested, image };
      })
      .filter((item) => item.name || item.code)
      .filter((item) => {
        const query = warehouseSearch.toLowerCase();
        return `${item.code} ${item.name}`.toLowerCase().includes(query);
      });
  };

  const uploadConsumptionFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      const rows = workbookToRows(workbook);
      setConsumptionRows(rows);
      setProducts(buildProductList(rows));
      setSelectedProduct("");
      setSelectedRecipe(null);
      setMessage("Consumption file loaded.");
    });
  };

  const uploadRecipeFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      setRecipeRows(workbookToRows(workbook));
      setSelectedRecipe(null);
      setMessage("Recipe / location file loaded.");
    });
  };

  const uploadTemplateFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      setTemplateMap(parseTemplateWorkbook(workbook));
      setTemplateStatus("Custom template loaded.");
    });
  };

  const uploadMusterFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      const items = parseMusterWorkbook(workbook);
      setMusterItems(items);
      setSelectedEquipment(null);
      setMusterMessage(`Equipment Muster List loaded from ${workbook.SheetNames.length} sheet(s).`);
    });
  };

  const uploadWarehouseFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcelFile(file, (workbook) => {
      const rows = workbookToRows(workbook);
      setWarehouseRows(rows);
      setWarehouseMessage("Warehouse inventory loaded.");
    });
  };

  const consumptionData = useMemo(() => consumptionRows.slice(1), [consumptionRows]);
  const recipeData = useMemo(() => recipeRows.slice(1), [recipeRows]);

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

  const getCombinedVenueBreakdown = (product) => {
    const actual = {};
    let currentVenue = "";
    consumptionData.forEach((row) => {
      if (row[2]) currentVenue = String(row[2]).trim();
      const venueKey = normalizeVenue(currentVenue);
      if (String(row[6]).trim() === product) {
        if (!actual[venueKey]) actual[venueKey] = { displayName: currentVenue, ships: {} };
        SHIPS.forEach((ship) => {
          actual[venueKey].ships[ship] = (actual[venueKey].ships[ship] || 0) + (Number(row[shipColumns[ship]]) || 0);
        });
      }
    });

    const required = {};
    recipeData.forEach((row) => {
      if (productMatches(product, row)) {
        const vKey = normalizeVenue(row[1]);
        if (vKey) required[vKey] = { displayName: String(row[1]).trim() };
      }
    });

    const allVenueKeys = Array.from(new Set([...Object.keys(actual), ...Object.keys(required)])).sort();
    return allVenueKeys.map((vK) => {
      const isReq = !!required[vK];
      const ships = actual[vK]?.ships || {};
      return {
        venueKey: vK,
        displayName: actual[vK]?.displayName || required[vK]?.displayName || vK,
        ships,
        required: isReq,
        missingShips: visibleShips.filter((s) => isReq && (ships[s] || 0) === 0),
      };
    });
  };

  const getRecipesUsingProduct = (product) => {
    const recipes = {};
    recipeData.forEach((row) => {
      const recipeCode = String(row[15] || "").trim();
      const recipeName = String(row[16] || "").trim();
      if ((!recipeCode && !recipeName) || (recipeName && !isNaN(Number(recipeName)))) return;
      if (!productMatches(product, row)) return;
      const key = `${recipeCode || "N/A"} - ${recipeName || "Unnamed Recipe"}`;
      if (!recipes[key]) recipes[key] = { key, recipeCode: recipeCode || "N/A", recipeName: recipeName || "Unnamed Recipe", venues: new Set() };
      if (row[1]) recipes[key].venues.add(row[1]);
    });
    return Object.values(recipes).map((r) => ({ ...r, venues: [...r.venues] }));
  };

  const getProductsInRecipe = (recipe) => {
    if (!recipe) return [];
    const items = {};
    recipeData.forEach((row) => {
      if (String(row[15]).trim() === recipe.recipeCode && String(row[16]).trim() === recipe.recipeName) {
        const product = String(row[12] || row[7] || "").trim();
        if (product) items[product] = true;
      }
    });
    return Object.keys(items).sort();
  };

  const parseMusterItems = () => {
    const grouped = {};
    musterItems.forEach((item) => {
      const searchText = `${item.sheetName} ${item.category} ${item.code} ${item.name}`.toLowerCase();
      if (musterSearch && !searchText.includes(musterSearch.toLowerCase())) return;
      const key = `${item.sheetName} / ${item.category}`;
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(item);
    });
    return grouped;
  };

  const combinedBreakdown = selectedProduct ? getCombinedVenueBreakdown(selectedProduct) : [];
  const recipesForProduct = selectedProduct ? getRecipesUsingProduct(selectedProduct) : [];
  const productsInRecipe = selectedRecipe ? getProductsInRecipe(selectedRecipe) : [];
  const filteredProducts = products.filter((p) => p.toLowerCase().includes(search.toLowerCase()));

  const totalConsumption = (() => {
    const totals = { BRL: 0, RL: 0, SC: 0, VL: 0 };
    combinedBreakdown.forEach((venue) => {
      visibleShips.forEach((ship) => { totals[ship] += Number(venue.ships[ship] || 0); });
    });
    return { totals, allShips: visibleShips.reduce((sum, ship) => sum + totals[ship], 0) };
  })();

  if (!loggedIn) {
    return (
      <main style={styles.page}>
        <section style={styles.loginCard}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
          <h1 style={styles.title}>Virgin Voyages Dashboard</h1>
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
            </button>
            <button style={styles.moduleCard} onClick={() => setModule("equipment")}>
              <div style={styles.moduleIcon}>🍽️</div>
              <strong>Equipment</strong>
            </button>
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && !equipmentMode) {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setModule("")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>
        <section style={styles.card}>
          <h2 style={styles.cardTitle}>🍽️ Equipment Options</h2>
          <div style={styles.moduleGrid}>
            <button style={styles.moduleCard} onClick={() => setEquipmentMode("muster")}>
              <div style={styles.moduleIcon}>📋</div>
              <strong>Equipment Muster List</strong>
            </button>
            <button style={styles.moduleCard} onClick={() => setEquipmentMode("inventory")}>
              <div style={styles.moduleIcon}>📊</div>
              <strong>Equipment Inventory</strong>
            </button>
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "inventory") {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>
        <section style={styles.card}>
          <h2 style={styles.cardTitle}>📊 Equipment Inventory</h2>
          <div style={styles.moduleGrid}>
            <button style={styles.moduleCard} onClick={() => setEquipmentMode("inuse")}>
              <div style={styles.moduleIcon}>✅</div>
              <strong>Inventory in Use</strong>
            </button>
            <button style={styles.moduleCard} onClick={() => setEquipmentMode("warehouse")}>
              <div style={styles.moduleIcon}>🏬</div>
              <strong>Inventory Warehouse</strong>
            </button>
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "inuse") {
    const warehouseItems = parseWarehouseItems();
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>
        <section style={styles.card}>
          <h2 style={styles.cardTitle}>✅ Inventory in Use</h2>
          <input placeholder="Search items..." value={warehouseSearch} onChange={(e) => setWarehouseSearch(e.target.value)} style={styles.searchInput} />
          <div style={styles.equipmentGrid}>
            {warehouseItems.map((item, i) => (
              <div key={i} style={styles.equipmentCard}>
                 {item.image && <img src={getImageUrl(item.image)} style={styles.equipmentImage} alt={item.name} />}
                 <strong>{item.name}</strong>
                 <span>On Hand: {item.onHand}</span>
              </div>
            ))}
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "warehouse") {
    const warehouseItems = parseWarehouseItems();
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>
        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📤 Upload Warehouse</h2>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadWarehouseFile} style={styles.fileInput} />
            <div style={styles.infoBox}>Column I = Picture Link</div>
          </div>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>🔍 Search</h2>
            <input placeholder="Search..." value={warehouseSearch} onChange={(e) => setWarehouseSearch(e.target.value)} style={styles.searchInput} />
          </div>
        </section>
        <section style={styles.card}>
          <div style={styles.equipmentGrid}>
            {warehouseItems.map((item, i) => (
              <div key={i} style={{ ...styles.equipmentCard, ...(item.suggested > 0 ? styles.orderWarningCard : {}) }}>
                {item.image ? (
                  <div style={{display:'flex', flexDirection:'column', gap:5}}>
                    <img src={getImageUrl(item.image)} alt={item.name} style={styles.equipmentImage} onError={(e) => { e.currentTarget.style.display = "none"; }} />
                    <a href={getImageUrl(item.image)} target="_blank" rel="noopener noreferrer" style={styles.imageLinkFull}>Open Picture</a>
                  </div>
                ) : <div style={styles.equipmentNoImage}>No image</div>}
                <div style={styles.recipeName}>{item.name}</div>
                <div style={styles.recipeMeta}>Code: {item.code}</div>
                <div style={item.suggested > 0 ? styles.suggestedOrderBad : styles.suggestedOrderGood}>Suggested: {formatQty(item.suggested)}</div>
              </div>
            ))}
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "muster") {
    const groupedMuster = parseMusterItems();
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>
        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📤 Upload Muster</h2>
            <input type="file" onChange={uploadMusterFile} style={styles.fileInput} />
          </div>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>🔍 Search</h2>
            <input placeholder="Search..." value={musterSearch} onChange={(e) => setMusterSearch(e.target.value)} style={styles.searchInput} />
          </div>
        </section>
        <section style={styles.card}>
          {Object.entries(groupedMuster).map(([cat, items]) => (
            <div key={cat} style={styles.equipmentCategory}>
              <h3 style={styles.sectionTitle}>🗂️ {cat}</h3>
              <div style={styles.equipmentGrid}>
                {items.map((item, index) => (
                  <button key={index} style={styles.equipmentCard} onClick={() => setSelectedEquipment(item)}>
                    {item.image ? <img src={getImageUrl(item.image)} alt={item.name} style={styles.equipmentImage} /> : <div style={styles.equipmentNoImage}>No image</div>}
                    <div style={styles.recipeName}>{item.name}</div>
                  </button>
                ))}
              </div>
            </div>
          ))}
          {selectedEquipment && (
            <div style={styles.modalBackdrop} onClick={() => setSelectedEquipment(null)}>
              <div style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
                <button style={styles.closeButton} onClick={() => setSelectedEquipment(null)}>✕</button>
                <h2>{selectedEquipment.name}</h2>
                {selectedEquipment.image && (
                    <div style={{textAlign:'center'}}>
                        <img src={getImageUrl(selectedEquipment.image)} alt={selectedEquipment.name} style={styles.modalImage} />
                        <a href={getImageUrl(selectedEquipment.image)} target="_blank" rel="noopener noreferrer" style={styles.imageLinkFull}>Open Original Link</a>
                    </div>
                )}
                <p>Code: {selectedEquipment.code}</p>
              </div>
            </div>
          )}
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
      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Data</h2>
          <input type="file" onChange={uploadConsumptionFile} style={styles.fileInput} />
          <input type="file" onChange={uploadRecipeFile} style={styles.fileInput} />
        </div>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔍 Product Search</h2>
          <input placeholder="Search..." value={search} onChange={(e) => setSearch(e.target.value)} style={styles.searchInput} />
          <div style={styles.productList}>
            {filteredProducts.map((p, i) => (
              <button key={i} onClick={() => setSelectedProduct(p)} style={{ ...styles.productItem, ...(selectedProduct === p ? styles.productItemActive : {}) }}>{p}</button>
            ))}
          </div>
        </div>
      </section>
      {selectedProduct && (
        <section style={styles.card}>
          <h2>📦 {selectedProduct}</h2>
          <div style={styles.venueGrid}>
            {combinedBreakdown.map((v, i) => (
              <div key={i} style={{ ...styles.venueCard, ...(v.missingShips.length > 0 ? styles.venueCardWarning : {}) }}>
                <strong>{v.displayName}</strong>
                <div style={styles.shipGrid}>
                  {visibleShips.map(s => (
                    <div key={s} style={{ ...styles.shipBox, ...(v.missingShips.includes(s) ? styles.shipBoxMissing : {}) }}>
                      <small>{s}</small>
                      <strong>{formatQty(v.ships[s])}</strong>
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
  select: { padding: 10, borderRadius: 8, border: "1px solid #ccc" },
  primaryButton: { marginTop: 10, padding: 12, borderRadius: 10, border: 0, background: "#111", color: "#fff", fontWeight: "bold", cursor: "pointer" },
  shipBadge: { padding: "10px 14px", borderRadius: 999, background: "#111", color: "#fff", fontWeight: "bold" },
  moduleGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16 },
  moduleCard: { border: "1px solid #ddd", background: "#fafafa", borderRadius: 16, padding: 20, cursor: "pointer", textAlign: "left", display: "grid", gap: 8, boxShadow: "0 4px 12px rgba(0,0,0,0.04)" },
  moduleIcon: { fontSize: 34 },
  grid: { display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 20, marginBottom: 20 },
  card: { background: "#fff", borderRadius: 16, padding: 20, boxShadow: "0 4px 18px rgba(0,0,0,0.06)" },
  cardTitle: { marginTop: 0 },
  fileInput: { display: "block", margin: "8px 0 16px" },
  infoBox: { marginTop: 12, padding: 12, borderRadius: 12, background: "#f2f2f2", display: "grid", gap: 6 },
  searchInput: { width: "100%", padding: 12, borderRadius: 10, border: "1px solid #ccc", marginBottom: 10 },
  productList: { maxHeight: 300, overflowY: "auto", border: "1px solid #ddd", borderRadius: 12 },
  productItem: { width: "100%", display: "block", textAlign: "left", padding: 10, border: 0, borderBottom: "1px solid #eee", background: "#fff", cursor: "pointer" },
  productItemActive: { background: "#eee", fontWeight: "bold" },
  venueGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14 },
  venueCard: { border: "1px solid #ddd", borderRadius: 14, padding: 14, background: "#fafafa" },
  venueCardWarning: { border: "2px solid #b00020", background: "#fff0f0" },
  shipGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(70px, 1fr))", gap: 6 },
  shipBox: { minWidth: 0, padding: "8px 4px", borderRadius: 10, background: "#fff", border: "1px solid #ddd", display: "grid", gap: 3, textAlign: "center" },
  shipBoxMissing: { background: "#b00020", color: "#fff", borderColor: "#b00020" },
  equipmentGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 },
  equipmentCard: { border: "1px solid #ddd", borderRadius: 14, padding: 14, background: "#fafafa", display: "grid", gap: 8, cursor: "pointer", textAlign: "left" },
  equipmentImage: { width: "100%", height: 150, objectFit: "cover", borderRadius: 10, background: "#eee" },
  equipmentNoImage: { height: 150, borderRadius: 10, background: "#eee", display: "flex", alignItems: "center", justifyContent: "center", color: "#777" },
  imageLinkFull: { display: "block", padding: "8px 12px", background: "#111", color: "#fff", borderRadius: 8, textAlign: "center", textDecoration: "none", fontSize: 13, fontWeight: "bold" },
  suggestedOrderBad: { marginTop: 8, padding: 8, borderRadius: 10, background: "#b00020", color: "#fff", fontWeight: "bold", textAlign: "center" },
  suggestedOrderGood: { marginTop: 8, padding: 8, borderRadius: 10, background: "#e8f5e9", color: "#2e7d32", fontWeight: "bold", textAlign: "center" },
  modalBackdrop: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999, padding: 20 },
  modalCard: { background: "#fff", borderRadius: 18, padding: 22, maxWidth: 760, width: "100%", maxHeight: "90vh", overflowY: "auto", position: "relative" },
  modalImage: { width: "100%", maxHeight: "65vh", objectFit: "contain", borderRadius: 14, background: "#f2f2f2", marginBottom: 15 },
  closeButton: { position: "absolute", top: 12, right: 12, border: 0, background: "#111", color: "#fff", borderRadius: 999, width: 34, height: 34, cursor: "pointer", fontWeight: "bold" },
};

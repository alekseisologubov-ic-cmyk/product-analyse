"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

const ALLERGEN_RULES = [
  { allergen: "Tree Nuts", keywords: ["almond", "walnut", "pecan", "cashew", "hazelnut", "pistachio", "macadamia"] },
  { allergen: "Peanuts", keywords: ["peanut"] },
  { allergen: "Soy", keywords: ["soy", "tofu", "edamame", "miso", "tamari"] },
  { allergen: "Gluten", keywords: ["wheat", "flour", "gluten", "bread", "pasta", "semolina", "barley", "rye", "panko"] },
  { allergen: "Milk / Dairy", keywords: ["milk", "cream", "butter", "cheese", "yogurt", "parmesan", "mozzarella", "ricotta", "cream cheese"] },
  { allergen: "Egg", keywords: ["egg", "mayonnaise", "aioli"] },
  { allergen: "Fish", keywords: ["salmon", "tuna", "cod", "anchovy", "fish", "sardine"] },
  { allergen: "Shellfish", keywords: ["shrimp", "crab", "lobster", "clam", "mussel", "oyster", "scallop"] },
  { allergen: "Sesame", keywords: ["sesame", "tahini"] },
  { allergen: "Mustard", keywords: ["mustard"] }
];

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const STORAGE_KEYS = {
  consumption: "vv_consumption_rows",
  recipe: "vv_recipe_rows",
};

export default function App() {
  const [consumptionRows, setConsumptionRows] = useState([]);
  const [recipeRows, setRecipeRows] = useState([]);
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [search, setSearch] = useState("");
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);
  const [message, setMessage] = useState("");

  const buildProductList = (rows) => {
    return [
      ...new Set(
        rows
          .slice(1)
          .map((r) => String(r[6] || "").trim())
          .filter(Boolean)
      ),
    ].sort();
  };

  const saveRows = (key, rows) => {
    try {
      localStorage.setItem(key, JSON.stringify(rows));
      setMessage("Files saved in this browser.");
    } catch {
      setMessage("Could not save files. File may be too large.");
    }
  };

  const loadSavedFiles = () => {
    try {
      const savedConsumption = localStorage.getItem(STORAGE_KEYS.consumption);
      const savedRecipe = localStorage.getItem(STORAGE_KEYS.recipe);

      if (!savedConsumption && !savedRecipe) {
        setMessage("No saved files found.");
        return;
      }

      if (savedConsumption) {
        const rows = JSON.parse(savedConsumption);
        setConsumptionRows(rows);
        setProducts(buildProductList(rows));
      }

      if (savedRecipe) {
        setRecipeRows(JSON.parse(savedRecipe));
      }

      setSelectedProduct("");
      setSelectedRecipe(null);
      setMessage("Saved files loaded.");
    } catch {
      setMessage("Could not load saved files.");
    }
  };

  const clearSavedFiles = () => {
    localStorage.removeItem(STORAGE_KEYS.consumption);
    localStorage.removeItem(STORAGE_KEYS.recipe);
    setConsumptionRows([]);
    setRecipeRows([]);
    setProducts([]);
    setSelectedProduct("");
    setSelectedRecipe(null);
    setMessage("Saved files cleared.");
  };

  const readExcel = (file, callback) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
      callback(rows);
    };
    reader.readAsBinaryString(file);
  };

  const uploadConsumptionFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcel(file, (rows) => {
      setConsumptionRows(rows);
      setProducts(buildProductList(rows));
      setSelectedProduct("");
      setSelectedRecipe(null);
      saveRows(STORAGE_KEYS.consumption, rows);
    });
  };

  const uploadRecipeFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcel(file, (rows) => {
      setRecipeRows(rows);
      saveRows(STORAGE_KEYS.recipe, rows);
    });
  };

  const consumptionData = useMemo(
    () => consumptionRows.slice(1),
    [consumptionRows]
  );

  const recipeData = useMemo(() => recipeRows.slice(1), [recipeRows]);

  const getConsumptionBreakdown = (product) => {
    let currentVenue = "";
    const result = {};

    const shipColumns = {
      BRL: 8,
      RL: 11,
      SC: 14,
      VL: 17,
    };

    consumptionData.forEach((row) => {
      if (row[2]) currentVenue = String(row[2]).trim();

      const venue = currentVenue || "Unknown";
      const productName = String(row[6] || "").trim();

      if (productName !== product) return;
      if (!result[venue]) result[venue] = {};

      SHIPS.forEach((ship) => {
        const qty = Number(row[shipColumns[ship]]) || 0;
        if (qty === 0) return;

        if (!result[venue][ship]) result[venue][ship] = 0;
        result[venue][ship] += qty;
      });
    });

    return result;
  };

  const getRecipesUsingProduct = (product) => {
    const recipes = {};
    const selectedCleanProduct = cleanText(product);

    recipeData.forEach((row) => {
      const assignedProduct = cleanText(row[12]); // M = Assigned full product name
      const recipeCode = String(row[15] || "").trim(); // P
      const recipeName = String(row[16] || "").trim(); // Q
      const venue = String(row[1] || "").trim(); // B

      if (!assignedProduct) return;
      if (!recipeCode && !recipeName) return;
      if (recipeName && !isNaN(Number(recipeName))) return;
      if (assignedProduct !== selectedCleanProduct) return;

      const key = `${recipeCode || "N/A"} - ${recipeName || "Unnamed Recipe"}`;

      if (!recipes[key]) {
        recipes[key] = {
          key,
          recipeCode: recipeCode || "N/A",
          recipeName: recipeName || "Unnamed Recipe",
          venues: new Set(),
        };
      }

      if (venue) recipes[key].venues.add(venue);
    });

    return Object.values(recipes).map((recipe) => ({
      ...recipe,
      venues: [...recipe.venues],
    }));
  };

  const getProductsInRecipe = (recipe) => {
    if (!recipe) return [];

    const items = {};

    recipeData.forEach((row) => {
      const recipeCode = String(row[15] || "").trim();
      const recipeName = String(row[16] || "").trim();

      if (recipeCode !== recipe.recipeCode || recipeName !== recipe.recipeName) return;

      const product = String(row[12] || row[7] || "").trim();
      if (!product) return;

      items[product] = true;
    });

    return Object.keys(items).sort();
  };

  const detectAllergens = (productsInRecipe) => {
    const found = {};

    productsInRecipe.forEach((product) => {
      const lowerProduct = product.toLowerCase();

      ALLERGEN_RULES.forEach((rule) => {
        const matchedKeyword = rule.keywords.find((keyword) =>
          lowerProduct.includes(keyword)
        );

        if (matchedKeyword) {
          if (!found[rule.allergen]) {
            found[rule.allergen] = new Set();
          }
          found[rule.allergen].add(product);
        }
      });
    });

    return Object.entries(found).map(([allergen, products]) => ({
      allergen,
      products: [...products].sort(),
    }));
  };

  if (!loggedIn) {
    return (
      <main style={styles.page}>
        <section style={styles.loginCard}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />

          <h1 style={styles.title}>Product Consumption Dashboard</h1>
          <p style={styles.subtitle}>Ship, venue & recipe usage analysis</p>

          <label style={styles.label}>🚢 Select your ship</label>
          <select
            value={userShip}
            onChange={(e) => setUserShip(e.target.value)}
            style={styles.select}
          >
            <option value="">Choose ship</option>
            {SHIPS.map((ship) => (
              <option key={ship}>{ship}</option>
            ))}
          </select>

          <button
            style={styles.primaryButton}
            onClick={() => userShip && setLoggedIn(true)}
          >
            Enter Dashboard
          </button>
        </section>
      </main>
    );
  }

  const breakdown = selectedProduct ? getConsumptionBreakdown(selectedProduct) : {};
  const recipesForProduct = selectedProduct ? getRecipesUsingProduct(selectedProduct) : [];
  const productsInRecipe = selectedRecipe ? getProductsInRecipe(selectedRecipe) : [];
  const allergenWarnings = selectedRecipe ? detectAllergens(productsInRecipe) : [];

  const filteredProducts = products.filter((p) =>
    p.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
        <div style={styles.shipBadge}>🚢 {userShip}</div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Files</h2>

          <label style={styles.label}>Step 1: Consumption file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadConsumptionFile}
            style={styles.fileInput}
          />

          <label style={styles.label}>Step 2: Recipe file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadRecipeFile}
            style={styles.fileInput}
          />

          <div style={styles.buttonRow}>
            <button style={styles.secondaryButton} onClick={loadSavedFiles}>
              Load Saved Files
            </button>
            <button style={styles.clearButton} onClick={clearSavedFiles}>
              Clear Saved
            </button>
          </div>

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.infoBox}>
            <div>📦 Products loaded: <strong>{products.length}</strong></div>
            <div>📘 Recipe rows loaded: <strong>{Math.max(recipeRows.length - 1, 0)}</strong></div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔍 Select Product</h2>

          <input
            placeholder="Search product..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={styles.searchInput}
          />

          <div style={styles.productList}>
            {filteredProducts.map((product, i) => (
              <button
                key={i}
                onClick={() => {
                  setSelectedProduct(product);
                  setSelectedRecipe(null);
                }}
                style={{
                  ...styles.productItem,
                  ...(selectedProduct === product ? styles.productItemActive : {}),
                }}
              >
                {product}
              </button>
            ))}
          </div>
        </div>
      </section>

      {selectedProduct && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📦 {selectedProduct}</h2>

          <h3 style={styles.sectionTitle}>🏢 Consumption by Venue and Ship</h3>

          <div style={styles.venueGrid}>
            {Object.entries(breakdown).map(([venue, ships], i) => (
              <div key={i} style={styles.venueCard}>
                <h4 style={styles.venueTitle}>{venue}</h4>

                <div style={styles.shipGrid}>
                  {SHIPS.map((ship) => (
                    <div
                      key={ship}
                      style={{
                        ...styles.shipBox,
                        ...(ship === userShip ? styles.shipBoxActive : {}),
                      }}
                    >
                      <span style={styles.shipName}>{ship}</span>
                      <strong>{ships[ship] || 0}</strong>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>

          <h3 style={styles.sectionTitle}>👨‍🍳 Recipes using this product</h3>

          {recipeRows.length === 0 && (
            <p style={styles.emptyText}>Upload or load the recipe file to see recipe details.</p>
          )}

          {recipeRows.length > 0 && recipesForProduct.length === 0 && (
            <p style={styles.emptyText}>No recipes found for this product.</p>
          )}

          <div style={styles.recipeList}>
            {recipesForProduct.map((recipe, i) => (
              <button
                key={i}
                onClick={() => setSelectedRecipe(recipe)}
                style={{
                  ...styles.recipeCard,
                  ...(selectedRecipe?.key === recipe.key ? styles.recipeCardActive : {}),
                }}
              >
                <div style={styles.recipeName}>{recipe.recipeName}</div>
                <div style={styles.recipeMeta}>Code: {recipe.recipeCode}</div>
                <div style={styles.recipeMeta}>
                  Venues: {recipe.venues.length ? recipe.venues.join(", ") : "N/A"}
                </div>
              </button>
            ))}
          </div>

          {selectedRecipe && (
            <div style={styles.ingredientsCard}>
              <h3 style={styles.sectionTitle}>🧾 Products used in recipe</h3>
              <h4 style={{ marginTop: 0 }}>
                {selectedRecipe.recipeName} ({selectedRecipe.recipeCode})
              </h4>

              {productsInRecipe.length === 0 ? (
                <p style={styles.emptyText}>No products found for this recipe.</p>
              ) : (
                <ul>
                  {productsInRecipe.map((product, i) => (
                    <li key={i}>{product}</li>
                  ))}
                </ul>
              )}

              <h3 style={styles.sectionTitle}>⚠️ Rule-Based Allergen Warning</h3>
              <p style={styles.warningText}>
                This is a keyword-based warning only. Verify against official allergen data before use.
              </p>

              {allergenWarnings.length === 0 ? (
                <p style={styles.emptyText}>No likely allergens detected by keyword rules.</p>
              ) : (
                <div style={styles.allergenList}>
                  {allergenWarnings.map((item, i) => (
                    <div key={i} style={styles.allergenCard}>
                      <strong>{item.allergen}</strong>
                      <ul>
                        {item.products.map((product, j) => (
                          <li key={j}>{product}</li>
                        ))}
                      </ul>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </section>
      )}
    </main>
  );
}

const styles = {
  page: {
    minHeight: "100vh",
    padding: 24,
    background: "#f5f5f5",
    fontFamily: "Arial, sans-serif",
    color: "#111",
  },
  loginCard: {
    maxWidth: 460,
    margin: "80px auto",
    padding: 28,
    background: "#fff",
    borderRadius: 16,
    boxShadow: "0 10px 30px rgba(0,0,0,0.08)",
    display: "grid",
    gap: 14,
  },
  logo: { height: 70, objectFit: "contain", marginBottom: 8 },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: 18,
    background: "#fff",
    borderRadius: 16,
    boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
    marginBottom: 20,
  },
  headerLogo: { height: 54, objectFit: "contain" },
  title: { margin: 0, fontSize: 28 },
  subtitle: { margin: 0, color: "#666" },
  label: { fontWeight: "bold", marginTop: 8 },
  select: { padding: 10, borderRadius: 8, border: "1px solid #ccc" },
  primaryButton: {
    marginTop: 10,
    padding: 12,
    borderRadius: 10,
    border: 0,
    background: "#111",
    color: "#fff",
    fontWeight: "bold",
    cursor: "pointer",
  },
  shipBadge: {
    padding: "10px 14px",
    borderRadius: 999,
    background: "#111",
    color: "#fff",
    fontWeight: "bold",
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "1fr 1.4fr",
    gap: 20,
    marginBottom: 20,
  },
  card: {
    background: "#fff",
    borderRadius: 16,
    padding: 20,
    boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
  },
  cardTitle: { marginTop: 0 },
  fileInput: { display: "block", margin: "8px 0 16px" },
  buttonRow: { display: "flex", gap: 10, marginTop: 10 },
  secondaryButton: {
    padding: 10,
    borderRadius: 8,
    border: "1px solid #111",
    background: "#fff",
    cursor: "pointer",
  },
  clearButton: {
    padding: 10,
    borderRadius: 8,
    border: "1px solid #ccc",
    background: "#eee",
    cursor: "pointer",
  },
  message: { color: "#555", fontSize: 14 },
  infoBox: {
    marginTop: 12,
    padding: 12,
    borderRadius: 12,
    background: "#f2f2f2",
    display: "grid",
    gap: 6,
  },
  searchInput: {
    width: "100%",
    padding: 12,
    borderRadius: 10,
    border: "1px solid #ccc",
    marginBottom: 10,
  },
  productList: {
    maxHeight: 300,
    overflowY: "auto",
    border: "1px solid #ddd",
    borderRadius: 12,
  },
  productItem: {
    width: "100%",
    display: "block",
    textAlign: "left",
    padding: 10,
    border: 0,
    borderBottom: "1px solid #eee",
    background: "#fff",
    cursor: "pointer",
  },
  productItemActive: { background: "#eee", fontWeight: "bold" },
  productTitle: { marginTop: 0, fontSize: 24 },
  sectionTitle: { marginTop: 22 },
  venueGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
    gap: 14,
  },
  venueCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 14,
    background: "#fafafa",
  },
  venueTitle: { marginTop: 0 },
  shipGrid: { display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 8 },
  shipBox: {
    padding: 10,
    borderRadius: 10,
    background: "#fff",
    border: "1px solid #ddd",
    display: "grid",
    gap: 4,
    textAlign: "center",
  },
  shipBoxActive: { background: "#111", color: "#fff" },
  shipName: { fontSize: 12, opacity: 0.8 },
  emptyText: { color: "#777" },
  recipeList: { display: "grid", gap: 10 },
  recipeCard: {
    width: "100%",
    textAlign: "left",
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 12,
    background: "#fafafa",
    cursor: "pointer",
  },
  recipeCardActive: { background: "#eee", borderColor: "#111" },
  recipeName: { fontWeight: "bold" },
  recipeMeta: { color: "#555", fontSize: 14, marginTop: 4 },
  ingredientsCard: {
    marginTop: 18,
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 14,
    background: "#fafafa",
  },
  warningText: {
    color: "#8a5a00",
    background: "#fff4d6",
    padding: 10,
    borderRadius: 8,
  },
  allergenList: {
    display: "grid",
    gap: 10,
  },
  allergenCard: {
    border: "1px solid #e1c16e",
    background: "#fff9e8",
    borderRadius: 10,
    padding: 10,
  },
};

"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "V1", "VL"];

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

export default function App() {
  const [consumptionRows, setConsumptionRows] = useState([]);
  const [recipeRows, setRecipeRows] = useState([]);
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [search, setSearch] = useState("");
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);

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

      const productList = [
        ...new Set(
          rows
            .slice(1)
            .map((r) => String(r[6] || "").trim()) // G = product in consumption file
            .filter(Boolean)
        ),
      ].sort();

      setProducts(productList);
      setSelectedProduct("");
    });
  };

  const uploadRecipeFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    readExcel(file, setRecipeRows);
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
      BRL: 8, // I
      RL: 11, // L
      V1: 14, // O
      VL: 17, // R
    };

    consumptionData.forEach((row) => {
      if (row[2]) currentVenue = String(row[2]).trim(); // C = venue

      const venue = currentVenue || "Unknown";
      const productName = String(row[6] || "").trim(); // G = product

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
      // IMPORTANT:
      // Recipe file:
      // M = Assigned full product name
      // P = RecipeCode
      // Q = RecipeName
      // B = Venue / Location
      const assignedProduct = cleanText(row[12]); // M = Assigned full product name
      const recipeCode = String(row[15] || "").trim(); // P
      const recipeName = String(row[16] || "").trim(); // Q
      const venue = String(row[1] || "").trim(); // B

      if (!assignedProduct) return;
      if (!recipeCode && !recipeName) return;
      if (recipeName && !isNaN(Number(recipeName))) return;

      // Exact clean match only.
      // This prevents beef matching wrong recipes.
      if (assignedProduct !== selectedCleanProduct) return;

      const key = `${recipeCode || "N/A"} - ${recipeName || "Unnamed Recipe"}`;

      if (!recipes[key]) {
        recipes[key] = {
          recipeCode: recipeCode || "N/A",
          recipeName: recipeName || "Unnamed Recipe",
          venues: new Set(),
        };
      }

      if (venue) recipes[key].venues.add(venue);
    });

    return Object.entries(recipes).map(([key, value]) => ({
      key,
      recipeCode: value.recipeCode,
      recipeName: value.recipeName,
      venues: [...value.venues],
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

  const breakdown = selectedProduct
    ? getConsumptionBreakdown(selectedProduct)
    : {};

  const recipesForProduct = selectedProduct
    ? getRecipesUsingProduct(selectedProduct)
    : [];

  const filteredProducts = products.filter((p) =>
    p.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
        <div>
          <h1 style={styles.headerTitle}>Product Consumption Dashboard</h1>
          <p style={styles.headerSubtitle}>Venue, ship & recipe analysis</p>
        </div>
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

          <div style={styles.infoBox}>
            <div>
              📦 Products loaded: <strong>{products.length}</strong>
            </div>
            <div>
              📘 Recipe rows loaded:{" "}
              <strong>{Math.max(recipeRows.length - 1, 0)}</strong>
            </div>
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
                onClick={() => setSelectedProduct(product)}
                style={{
                  ...styles.productItem,
                  ...(selectedProduct === product
                    ? styles.productItemActive
                    : {}),
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
            <p style={styles.emptyText}>
              Upload the recipe file to see recipe details.
            </p>
          )}

          {recipeRows.length > 0 && recipesForProduct.length === 0 && (
            <p style={styles.emptyText}>
              No recipes found for this product.
            </p>
          )}

          <div style={styles.recipeList}>
            {recipesForProduct.map((recipe, i) => (
              <div key={i} style={styles.recipeCard}>
                <div style={styles.recipeName}>{recipe.recipeName}</div>
                <div style={styles.recipeMeta}>Code: {recipe.recipeCode}</div>
                <div style={styles.recipeMeta}>
                  Venues:{" "}
                  {recipe.venues.length ? recipe.venues.join(", ") : "N/A"}
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
  logo: {
    height: 70,
    objectFit: "contain",
    marginBottom: 8,
  },
  headerLogo: {
    height: 54,
    objectFit: "contain",
  },
  title: {
    margin: 0,
    fontSize: 28,
  },
  subtitle: {
    margin: 0,
    color: "#666",
  },
  label: {
    fontWeight: "bold",
    marginTop: 8,
  },
  select: {
    padding: 10,
    borderRadius: 8,
    border: "1px solid #ccc",
  },
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
  header: {
    display: "flex",
    alignItems: "center",
    gap: 16,
    padding: 18,
    background: "#fff",
    borderRadius: 16,
    boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
    marginBottom: 20,
  },
  headerTitle: {
    margin: 0,
    fontSize: 26,
  },
  headerSubtitle: {
    margin: 0,
    color: "#666",
  },
  shipBadge: {
    marginLeft: "auto",
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
  cardTitle: {
    marginTop: 0,
  },
  fileInput: {
    display: "block",
    margin: "8px 0 16px",
  },
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
  productItemActive: {
    background: "#eee",
    fontWeight: "bold",
  },
  productTitle: {
    marginTop: 0,
    fontSize: 24,
  },
  sectionTitle: {
    marginTop: 22,
  },
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
  venueTitle: {
    marginTop: 0,
  },
  shipGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(4, 1fr)",
    gap: 8,
  },
  shipBox: {
    padding: 10,
    borderRadius: 10,
    background: "#fff",
    border: "1px solid #ddd",
    display: "grid",
    gap: 4,
    textAlign: "center",
  },
  shipBoxActive: {
    background: "#111",
    color: "#fff",
  },
  shipName: {
    fontSize: 12,
    opacity: 0.8,
  },
  emptyText: {
    color: "#777",
  },
  recipeList: {
    display: "grid",
    gap: 10,
  },
  recipeCard: {
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 12,
    background: "#fafafa",
  },
  recipeName: {
    fontWeight: "bold",
  },
  recipeMeta: {
    color: "#555",
    fontSize: 14,
    marginTop: 4,
  },
};

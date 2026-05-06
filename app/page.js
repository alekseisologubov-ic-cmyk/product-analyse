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
            .map((r) => String(r[6] || "").trim())
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

  const consumptionData = useMemo(() => consumptionRows.slice(1), [consumptionRows]);
  const recipeData = useMemo(() => recipeRows.slice(1), [recipeRows]);

  const getConsumptionBreakdown = (product) => {
    let currentVenue = "";
    const result = {};

    const shipColumns = {
      BRL: 8,
      RL: 11,
      V1: 14,
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
    const cleanProduct = cleanText(product);

    recipeData.forEach((row) => {
      const assignedProduct = cleanText(row[12]); // Column M
      const recipeCode = String(row[15] || "").trim();
      const recipeName = String(row[16] || "").trim();
      const venue = String(row[1] || "").trim();

      if (!assignedProduct) return;
      if (!recipeCode && !recipeName) return;
      if (recipeName && !isNaN(Number(recipeName))) return;

      if (assignedProduct !== cleanProduct) return;

      const key = `${recipeCode} - ${recipeName}`;

      if (!recipes[key]) {
        recipes[key] = {
          recipeCode,
          recipeName,
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
      <div style={styles.page}>
        <div style={styles.loginCard}>
          <img src="/virgin-logo.png" style={{ height: 70 }} />

          <h2>Select Your Ship</h2>

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

          <button style={styles.button} onClick={() => userShip && setLoggedIn(true)}>
            Enter
          </button>
        </div>
      </div>
    );
  }

  const breakdown = selectedProduct ? getConsumptionBreakdown(selectedProduct) : {};
  const recipesForProduct = selectedProduct ? getRecipesUsingProduct(selectedProduct) : [];

  return (
    <div style={styles.page}>
      <div style={styles.header}>
        <img src="/virgin-logo.png" style={{ height: 60 }} />
      </div>

      <div style={styles.grid}>
        <div style={styles.card}>
          <h3>Upload Files</h3>

          <input type="file" onChange={uploadConsumptionFile} />
          <input type="file" onChange={uploadRecipeFile} />

          <p>Products: {products.length}</p>
          <p>Recipe rows: {Math.max(recipeRows.length - 1, 0)}</p>
        </div>

        <div style={styles.card}>
          <h3>Search Product</h3>

          <input
            placeholder="Search..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={styles.input}
          />

          <div style={styles.list}>
            {products
              .filter((p) => p.toLowerCase().includes(search.toLowerCase()))
              .map((p, i) => (
                <div
                  key={i}
                  onClick={() => setSelectedProduct(p)}
                  style={{
                    ...styles.item,
                    background: selectedProduct === p ? "#eee" : "#fff",
                  }}
                >
                  {p}
                </div>
              ))}
          </div>
        </div>
      </div>

      {selectedProduct && (
        <div style={styles.card}>
          <h2>{selectedProduct}</h2>

          <h3>Consumption</h3>

          {Object.entries(breakdown).map(([venue, ships], i) => (
            <div key={i}>
              <strong>{venue}</strong>
              <div>
                {SHIPS.map((s) => (
                  <span key={s} style={{ marginRight: 10 }}>
                    {s}: {ships[s] || 0}
                  </span>
                ))}
              </div>
            </div>
          ))}

          <h3>Recipes</h3>

          {recipesForProduct.length === 0 && <p>No recipes found</p>}

          {recipesForProduct.map((r, i) => (
            <div key={i}>
              {r.recipeName} ({r.recipeCode})
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const styles = {
  page: { padding: 20, background: "#f5f5f5" },
  header: { textAlign: "center", marginBottom: 20 },
  loginCard: {
    maxWidth: 400,
    margin: "80px auto",
    padding: 20,
    background: "#fff",
    borderRadius: 12,
    display: "grid",
    gap: 10,
  },
  select: { padding: 10 },
  button: { padding: 10, background: "#111", color: "#fff" },
  grid: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 },
  card: { background: "#fff", padding: 20, borderRadius: 12 },
  input: { padding: 10, width: "100%" },
  list: { maxHeight: 200, overflowY: "auto", border: "1px solid #ccc" },
  item: { padding: 6, cursor: "pointer" },
};

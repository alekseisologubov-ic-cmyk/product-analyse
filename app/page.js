"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "V1", "VL"];

// 🔥 NEW: cleaning function to match products between files
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
  const [selectedRecipe, setSelectedRecipe] = useState("");
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

  // 🔹 FILE 1 (consumption)
  const uploadConsumptionFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcel(file, (rows) => {
      setConsumptionRows(rows);

      const productList = [
        ...new Set(
          rows
            .slice(1)
            .map((r) => String(r[6] || "").trim()) // Column G
            .filter(Boolean)
        ),
      ].sort();

      setProducts(productList);
    });
  };

  // 🔹 FILE 2 (recipes)
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

  // 🔹 Consumption breakdown
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

  // 🔥 FIXED: recipes matching (this is what was missing)
  const getRecipesUsingProduct = (product) => {
    const recipes = {};
    const cleanProduct = cleanText(product);

    recipeData.forEach((row) => {
      const rawProduct = row[7]; // Column H
      const productName = cleanText(rawProduct);

      // smart match (handles different naming)
      if (
        productName !== cleanProduct &&
        !productName.includes(cleanProduct) &&
        !cleanProduct.includes(productName)
      ) {
        return;
      }

      const venue = String(row[1] || "").trim(); // Column B
      const recipeCode = String(row[15] || "").trim(); // P
      const recipeName = String(row[16] || "").trim(); // Q

      if (!recipeCode && !recipeName) return;

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
      <div style={{ padding: 20 }}>
        <h2>Select Your Ship</h2>

        <select value={userShip} onChange={(e) => setUserShip(e.target.value)}>
          <option value="">Choose ship</option>
          {SHIPS.map((ship) => (
            <option key={ship}>{ship}</option>
          ))}
        </select>

        <br /><br />

        <button onClick={() => userShip && setLoggedIn(true)}>
          Enter
        </button>
      </div>
    );
  }

  const recipesForProduct = selectedProduct
    ? getRecipesUsingProduct(selectedProduct)
    : [];

  return (
    <div style={{ padding: 20 }}>
      <h2>Your Ship: {userShip}</h2>

      <div>
        <strong>Step 1: Upload consumption file</strong>
        <br />
        <input type="file" onChange={uploadConsumptionFile} />
      </div>

      <br />

      <div>
        <strong>Step 2: Upload recipe file</strong>
        <br />
        <input type="file" onChange={uploadRecipeFile} />
      </div>

      <br />

      <input
        placeholder="Search product..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />

      <div style={{ maxHeight: 200, overflowY: "scroll", border: "1px solid #ccc" }}>
        {products
          .filter((p) => p.toLowerCase().includes(search.toLowerCase()))
          .map((p, i) => (
            <div
              key={i}
              onClick={() => {
                setSelectedProduct(p);
                setSelectedRecipe("");
              }}
              style={{ cursor: "pointer", padding: 4 }}
            >
              {p}
            </div>
          ))}
      </div>

      {selectedProduct && (
        <div style={{ marginTop: 20 }}>
          <h3>{selectedProduct}</h3>

          <h4>Consumption by Venue and Ship</h4>

          {Object.entries(getConsumptionBreakdown(selectedProduct)).map(
            ([venue, ships], i) => (
              <div key={i}>
                <strong>{venue}</strong>
                <div>
                  {SHIPS.map((ship) => (
                    <span key={ship} style={{ marginRight: 10 }}>
                      {ship}: {ships[ship] || 0}
                    </span>
                  ))}
                </div>
              </div>
            )
          )}

          <h4>Recipes using this product</h4>

          {recipesForProduct.length === 0 && (
            <p>No recipes found for this product.</p>
          )}

          {recipesForProduct.map((recipe, i) => (
            <div key={i}>
              {recipe.recipeName} ({recipe.recipeCode})
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

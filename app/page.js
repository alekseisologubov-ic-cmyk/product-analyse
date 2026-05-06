"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "V1", "VL"];

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

  const uploadConsumptionFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcel(file, (rows) => {
      setConsumptionRows(rows);

      const productList = [
        ...new Set(
          rows
            .slice(1)
            .map((r) => String(r[6] || "").trim()) // G = product
            .filter(Boolean)
        ),
      ].sort();

      setProducts(productList);
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
      BRL: 8,  // I
      RL: 11,  // L
      V1: 14,  // O
      VL: 17,  // R
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

    recipeData.forEach((row) => {
      const productName = String(row[7] || "").trim(); // H = product
      if (productName !== product) return;

      const venue = String(row[1] || "").trim(); // B = venue
      const recipeCode = String(row[15] || "").trim(); // P = recipe code
      const recipeName = String(row[16] || "").trim(); // Q = recipe name

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

  const getProductsInRecipe = (recipeKey) => {
    const productsInRecipe = {};

    recipeData.forEach((row) => {
      const productName = String(row[7] || "").trim(); // H = product
      const recipeCode = String(row[15] || "").trim(); // P
      const recipeName = String(row[16] || "").trim(); // Q
      const key = `${recipeCode} - ${recipeName}`;

      if (key !== recipeKey) return;
      if (!productName) return;

      if (!productsInRecipe[productName]) {
        productsInRecipe[productName] = 0;
      }

      productsInRecipe[productName] += 1;
    });

    return Object.keys(productsInRecipe).sort();
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

        <br />
        <br />

        <button onClick={() => userShip && setLoggedIn(true)}>Enter</button>
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
        style={{ width: 350, padding: 8 }}
      />

      <div
        style={{
          maxHeight: 220,
          overflowY: "scroll",
          border: "1px solid #ccc",
          marginTop: 10,
        }}
      >
        {products
          .filter((p) => p.toLowerCase().includes(search.toLowerCase()))
          .map((product, i) => (
            <div
              key={i}
              onClick={() => {
                setSelectedProduct(product);
                setSelectedRecipe("");
              }}
              style={{
                cursor: "pointer",
                padding: 6,
                background: selectedProduct === product ? "#ddd" : "white",
              }}
            >
              {product}
            </div>
          ))}
      </div>

      {selectedProduct && (
        <div style={{ marginTop: 25 }}>
          <h2>{selectedProduct}</h2>

          <h3>Consumption by Venue and Ship</h3>

          {Object.entries(getConsumptionBreakdown(selectedProduct)).map(
            ([venue, ships], i) => (
              <div key={i} style={{ borderBottom: "1px solid #ddd", padding: 10 }}>
                <strong>{venue}</strong>

                <div style={{ display: "flex", gap: 15 }}>
                  {SHIPS.map((ship) => (
                    <div
                      key={ship}
                      style={{
                        fontWeight: ship === userShip ? "bold" : "normal",
                      }}
                    >
                      {ship}: {ships[ship] || 0}
                    </div>
                  ))}
                </div>
              </div>
            )
          )}

          <h3>Recipes using this product</h3>

          {recipeRows.length === 0 && (
            <p>Please upload the recipe file to see recipes.</p>
          )}

          {recipesForProduct.length === 0 && recipeRows.length > 0 && (
            <p>No recipes found for this product.</p>
          )}

          {recipesForProduct.map((recipe, i) => (
            <div
              key={i}
              onClick={() => setSelectedRecipe(recipe.key)}
              style={{
                cursor: "pointer",
                padding: 8,
                borderBottom: "1px solid #ddd",
                background: selectedRecipe === recipe.key ? "#eee" : "white",
              }}
            >
              <strong>{recipe.recipeName}</strong>
              <br />
              Code: {recipe.recipeCode}
              <br />
              Venues: {recipe.venues.join(", ")}
            </div>
          ))}
        </div>
      )}

      {selectedRecipe && (
        <div style={{ marginTop: 25 }}>
          <h3>Products used in recipe</h3>
          <h4>{selectedRecipe}</h4>

          <ul>
            {getProductsInRecipe(selectedRecipe).map((product, i) => (
              <li key={i}>{product}</li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}

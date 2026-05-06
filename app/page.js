"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "V1", "VL"];

export default function App() {
  const [rawRows, setRawRows] = useState([]);
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [search, setSearch] = useState("");
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);

  const handleLogin = () => {
    if (!userShip) return;
    setLoggedIn(true);
  };

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      setRawRows(rows);

      const uniqueProducts = [
        ...new Set(
          rows
            .slice(1)
            .map((r) => String(r[7] || "").trim()) // H = Product
            .filter(Boolean)
        ),
      ].sort();

      setProducts(uniqueProducts);
    };

    reader.readAsBinaryString(file);
  };

  const dataRows = useMemo(() => rawRows.slice(1), [rawRows]);

  const getBreakdown = (product) => {
    let currentVenue = "";
    const result = {};

    const shipColumns = {
      BRL: 8,  // I
      RL: 11,  // L
      V1: 14,  // O
      VL: 17,  // R
    };

    dataRows.forEach((row) => {
      if (row[1]) currentVenue = String(row[1]).trim(); // B = Venue

      const venue = currentVenue || "Unknown";
      const productName = String(row[7] || "").trim(); // H = Product
      if (productName !== product) return;

      const recipeCode = String(row[15] || "").trim(); // P
      const recipeName = String(row[16] || "").trim(); // Q

      if (!result[venue]) {
        result[venue] = {
          ships: {},
          recipes: {},
        };
      }

      SHIPS.forEach((ship) => {
        const qty = Number(row[shipColumns[ship]]) || 0;
        if (qty === 0) return;

        if (!result[venue].ships[ship]) result[venue].ships[ship] = 0;
        result[venue].ships[ship] += qty;

        const recipeKey = `${recipeCode} - ${recipeName}`;
        if (recipeName || recipeCode) {
          if (!result[venue].recipes[recipeKey]) {
            result[venue].recipes[recipeKey] = 0;
          }
          result[venue].recipes[recipeKey] += qty;
        }
      });
    });

    return result;
  };

  if (!loggedIn) {
    return (
      <div style={{ padding: 20 }}>
        <h2>Select Your Ship</h2>
        <select value={userShip} onChange={(e) => setUserShip(e.target.value)}>
          <option value="">Choose ship</option>
          {SHIPS.map((s) => (
            <option key={s}>{s}</option>
          ))}
        </select>

        <br /><br />

        <button onClick={handleLogin}>Enter</button>
      </div>
    );
  }

  return (
    <div style={{ padding: 20 }}>
      <h2>Your Ship: {userShip}</h2>

      <input type="file" onChange={handleFileUpload} />

      <br /><br />

      <input
        placeholder="Search product..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        style={{ width: 300, padding: 8 }}
      />

      <div style={{ maxHeight: 250, overflowY: "scroll", border: "1px solid #ccc", marginTop: 10 }}>
        {products
          .filter((p) => p.toLowerCase().includes(search.toLowerCase()))
          .map((p, i) => (
            <div
              key={i}
              onClick={() => setSelectedProduct(p)}
              style={{
                cursor: "pointer",
                padding: 6,
                background: selectedProduct === p ? "#ddd" : "white",
              }}
            >
              {p}
            </div>
          ))}
      </div>

      {selectedProduct && (
        <div style={{ marginTop: 20 }}>
          <h3>{selectedProduct}</h3>

          {Object.entries(getBreakdown(selectedProduct)).map(([venue, data], i) => (
            <div key={i} style={{ borderBottom: "1px solid #ddd", padding: 12 }}>
              <h4>{venue}</h4>

              <div style={{ display: "flex", gap: 16, marginBottom: 10 }}>
                {SHIPS.map((ship) => (
                  <div
                    key={ship}
                    style={{
                      fontWeight: ship === userShip ? "bold" : "normal",
                    }}
                  >
                    {ship}: {data.ships[ship] || 0}
                  </div>
                ))}
              </div>

              <strong>Recipes using this product:</strong>
              <ul>
                {Object.entries(data.recipes).map(([recipe, qty], j) => (
                  <li key={j}>
                    {recipe} — Qty: {qty}
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

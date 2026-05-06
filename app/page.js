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
          rows.slice(1).map((r) => String(r[6] || "").trim()).filter(Boolean)
        ),
      ];

      setProducts(uniqueProducts);
    };

    reader.readAsBinaryString(file);
  };

  const dataRows = useMemo(() => rawRows.slice(1), [rawRows]);

  const getBreakdown = (product) => {
    let currentVenue = "";
    const result = {};

    const shipColumns = {
      BRL: 8,
      RL: 11,
      V1: 14,
      VL: 17,
    };

    dataRows.forEach((row) => {
      if (row[2]) currentVenue = row[2];

      const venue = currentVenue;
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
      />

      <div style={{ maxHeight: 200, overflowY: "scroll", border: "1px solid #ccc" }}>
        {products
          .filter((p) => p.toLowerCase().includes(search.toLowerCase()))
          .map((p, i) => (
            <div key={i} onClick={() => setSelectedProduct(p)} style={{ cursor: "pointer" }}>
              {p}
            </div>
          ))}
      </div>

      {selectedProduct && (
        <div>
          <h3>{selectedProduct}</h3>
          {Object.entries(getBreakdown(selectedProduct)).map(([venue, ships], i) => (
            <div key={i}>
              <strong>{venue}</strong>
              <div style={{ display: "flex", gap: 10 }}>
                {SHIPS.map((ship) => (
                  <div key={ship} style={{ fontWeight: ship === userShip ? "bold" : "normal" }}>
                    {ship}: {ships[ship] || 0}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

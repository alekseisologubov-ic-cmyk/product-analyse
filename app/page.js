"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

const cleanText = (value) =>
  String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

const formatQty = (value) => Number(value || 0).toFixed(2);

export default function App() {

  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);

  const [module, setModule] = useState("");
  const [equipmentMode, setEquipmentMode] = useState("");

  const [musterItems, setMusterItems] = useState([]);
  const [musterMessage, setMusterMessage] = useState("");

  const [inUseRows, setInUseRows] = useState([]);
  const [inUseSearch, setInUseSearch] = useState("");
  const [inUseMessage, setInUseMessage] = useState("");

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

  const parseMusterWorkbook = (workbook) => {
    const items = [];

    workbook.SheetNames.forEach((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

      rows.slice(1).forEach((row) => {
        const category = row[2];
        const code = row[3];
        const name = row[4];

        if (!code || !name) return;

        items.push({
          sheetName,
          category,
          code: String(code).trim(),
          name: String(name).trim(),
        });
      });
    });

    return items;
  };

  const uploadMusterFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      setMusterItems(parseMusterWorkbook(workbook));
      setMusterMessage("Muster list loaded.");
    });
  };

  const uploadInUseFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      setInUseRows(workbookToRows(workbook));
      setInUseMessage("Inventory in Use loaded.");
    });
  };

  const parseInUseItems = () => {
    const actualMap = {};

    inUseRows.slice(1).forEach((row) => {
      const code = String(row[0] || "").trim();
      const onHand = Number(row[7] || 0);

      if (!code) return;

      actualMap[cleanText(code)] = onHand;
    });

    return musterItems
      .map((item) => {
        const onHand = actualMap[cleanText(item.code)] || 0;

        let status = "Missing";
        if (onHand > 0) status = "In Use";
        if (onHand === 0 && actualMap[cleanText(item.code)] !== undefined) status = "Zero Count";

        return { ...item, onHand, status };
      })
      .filter((item) =>
        `${item.code} ${item.name} ${item.status}`
          .toLowerCase()
          .includes(inUseSearch.toLowerCase())
      )
      .sort((a, b) => {
        const order = { Missing: 0, "Zero Count": 1, "In Use": 2 };
        return order[a.status] - order[b.status];
      });
  };
    if (!loggedIn) {
    return (
      <main style={styles.page}>
        <div style={styles.card}>
          <h2>Select Ship</h2>
          <select value={userShip} onChange={(e) => setUserShip(e.target.value)}>
            <option value="">Select</option>
            {SHIPS.map((s) => <option key={s}>{s}</option>)}
          </select>
          <button style={styles.primaryButton} onClick={() => setLoggedIn(true)}>Continue</button>
        </div>
      </main>
    );
  }

  if (!module) {
    return (
      <main style={styles.page}>
        <div style={styles.card}>
          <button style={styles.primaryButton} onClick={() => setModule("equipment")}>
            Equipment
          </button>
        </div>
      </main>
    );
  }

  if (!equipmentMode) {
    return (
      <main style={styles.page}>
        <div style={styles.card}>
          <button style={styles.primaryButton} onClick={() => setEquipmentMode("inuse")}>
            Inventory in Use
          </button>
        </div>
      </main>
    );
  }

  if (equipmentMode === "inuse") {
    const items = parseInUseItems();

    const missing = items.filter(i => i.status === "Missing");
    const zero = items.filter(i => i.status === "Zero Count");
    const ok = items.filter(i => i.status === "In Use");

    return (
      <main style={styles.page}>
        <div style={styles.grid}>
          <div style={styles.card}>
            <h3>Upload Files</h3>

            <label>Muster List</label>
            <input type="file" onChange={uploadMusterFile} />

            <label>Inventory In Use</label>
            <input type="file" onChange={uploadInUseFile} />

            <p>{musterMessage}</p>
            <p>{inUseMessage}</p>
          </div>

          <div style={styles.card}>
            <input
              placeholder="Search..."
              value={inUseSearch}
              onChange={(e) => setInUseSearch(e.target.value)}
              style={styles.searchInput}
            />
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={{ color: "#b00020" }}>❌ Missing from Inventory</h2>

          <div style={styles.grid}>
            {missing.map((item, i) => (
              <div key={i} style={{ ...styles.card, ...styles.red }}>
                <b>{item.name}</b>
                <div>{item.code}</div>
                <div>Status: Missing</div>
              </div>
            ))}
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={{ color: "#8a5a00" }}>⚠️ Zero Count</h2>

          <div style={styles.grid}>
            {zero.map((item, i) => (
              <div key={i} style={{ ...styles.card, ...styles.orange }}>
                <b>{item.name}</b>
                <div>{item.code}</div>
                <div>On Hand: 0</div>
              </div>
            ))}
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={{ color: "#2e7d32" }}>✅ In Use</h2>

          <div style={styles.grid}>
            {ok.map((item, i) => (
              <div key={i} style={styles.card}>
                <b>{item.name}</b>
                <div>{item.code}</div>
                <div>On Hand: {formatQty(item.onHand)}</div>
              </div>
            ))}
          </div>
        </div>
      </main>
    );
  }

  return null;
}

const styles = {
  page: { padding: 20, background: "#f5f5f5", fontFamily: "Arial" },
  grid: { display: "grid", gap: 12 },
  card: {
    padding: 14,
    borderRadius: 12,
    background: "#fff",
    boxShadow: "0 4px 12px rgba(0,0,0,0.06)"
  },
  primaryButton: {
    padding: 10,
    borderRadius: 10,
    background: "#111",
    color: "#fff",
    border: 0,
    cursor: "pointer"
  },
  searchInput: {
    width: "100%",
    padding: 10,
    borderRadius: 10,
    border: "1px solid #ccc"
  },
  red: { border: "2px solid #b00020", background: "#fff0f0" },
  orange: { border: "2px solid #8a5a00", background: "#fff4d6" }
};

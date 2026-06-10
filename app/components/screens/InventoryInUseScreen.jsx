"use client";

import React from "react";

export default function InventoryInUseScreen(props) {
  const {
    styles,
    formatQty,
    getShipDisplayName,
    inUseMessage,
    inUseRows,
    inUseSearch,
    musterItems,
    musterMessage,
    parseInUseItems,
    setEquipmentMode,
    setInUseSearch,
    uploadInUseFile,
    uploadMusterFile,
    userShip,
  } = props;

  const inUseItems = parseInUseItems();
      const missingItems = inUseItems.filter((item) => item.status === "Missing");
      const zeroItems = inUseItems.filter((item) => item.status === "Zero Count");
      const activeItems = inUseItems.filter((item) => item.status === "In Use");

      return (
        <main style={styles.page}>
          <header style={styles.header}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
            <div style={styles.headerActions}>
              <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
              <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
            </div>
          </header>

          <section style={styles.grid}>
            <div style={styles.card}>
              <h2 style={styles.cardTitle}>📤 Upload Inventory in Use</h2>

              <label style={styles.label}>Step 1: Equipment Master List file</label>
              <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadMusterFile} style={styles.fileInput} />

              <label style={styles.label}>Step 2: Inventory in Use file</label>
              <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadInUseFile} style={styles.fileInput} />

              {musterMessage && <p style={styles.message}>{musterMessage}</p>}
              {inUseMessage && <p style={styles.message}>{inUseMessage}</p>}

              <div style={styles.infoBox}>
                <div>❌ Missing from Inventory: <strong>{missingItems.length}</strong></div>
                <div>⚠️ Zero Count: <strong>{zeroItems.length}</strong></div>
                <div>✅ In Use: <strong>{activeItems.length}</strong></div>
                <div>Master List: C = Category, D = Code, E = Name</div>
                <div>In Use: A = Code, B = Name, H = On Hand</div>
              </div>
            </div>

            <div style={styles.card}>
              <h2 style={styles.cardTitle}>🔍 Search Inventory in Use</h2>

              <input
                placeholder="Search code, name, category, sheet or status..."
                value={inUseSearch}
                onChange={(e) => setInUseSearch(e.target.value)}
                style={styles.searchInput}
              />

              <p style={styles.emptyText}>
                Missing items are shown first in red under Missing from Inventory.
              </p>
            </div>
          </section>

          <section style={styles.card}>
            <h2 style={{ ...styles.productTitle, color: "#b00020" }}>❌ Missing from Inventory</h2>

            {musterItems.length === 0 && <p style={styles.emptyText}>Upload the Equipment Master List first.</p>}
            {inUseRows.length === 0 && <p style={styles.emptyText}>Upload the Inventory in Use file to compare.</p>}
            {musterItems.length > 0 && inUseRows.length > 0 && missingItems.length === 0 && (
              <p style={styles.emptyText}>No missing items found.</p>
            )}

            <div style={styles.equipmentGrid}>
              {missingItems.map((item, i) => (
                <div key={`${item.code}-missing-${i}`} style={{ ...styles.equipmentCard, ...styles.orderWarningCard }}>
                  <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                  <div style={styles.recipeMeta}>Category: {item.category}</div>
                  <div style={styles.statusBad}>Missing</div>
                </div>
              ))}
            </div>
          </section>

          <section style={styles.card}>
            <h2 style={{ ...styles.productTitle, color: "#8a5a00" }}>⚠️ Zero Count</h2>

            <div style={styles.equipmentGrid}>
              {zeroItems.map((item, i) => (
                <div key={`${item.code}-zero-${i}`} style={{ ...styles.equipmentCard, ...styles.zeroCountCard }}>
                  <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                  <div style={styles.recipeMeta}>Category: {item.category}</div>
                  <div style={styles.recipeMeta}>On Hand: {formatQty(item.onHand)}</div>
                  <div style={styles.statusWarning}>Zero Count</div>
                </div>
              ))}
            </div>
          </section>

          <section style={styles.card}>
            <h2 style={{ ...styles.productTitle, color: "#2e7d32" }}>✅ In Use</h2>

            <div style={styles.equipmentGrid}>
              {activeItems.map((item, i) => (
                <div key={`${item.code}-active-${i}`} style={styles.equipmentCard}>
                  <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                  <div style={styles.recipeMeta}>Category: {item.category}</div>
                  <div style={styles.recipeMeta}>On Hand: {formatQty(item.onHand)}</div>
                  <div style={styles.statusGood}>In Use</div>
                </div>
              ))}
            </div>
          </section>
        </main>
      );
}

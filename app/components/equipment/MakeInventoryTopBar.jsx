"use client";

import React from "react";

export default function MakeInventoryTopBar({
  styles,
  isAdmin,

  inventoryReportMode,
  setInventoryReportMode,

  makeInventoryShip,
  userShip,
  refreshMakeInventoryData,

  inventoryLoading,
  reportBusy,

  resetInventoryRun,
  printInventorySummary,
  clearMyInventory,
  clearShipInventory,
  generateFinalInventoryReport,
  exportInventorySummaryToExcel,

  allStationsSubmitted,

  inventoryStation,
  currentStationProgress,
  myReportRows,
  makeInventoryItems,

  startedStations,
  submittedStations,
  stationProgressRows,
}) {
  return (
    <section style={styles.card}>
      <div
        style={{
          ...styles.header,
          boxShadow: "none",
          padding: 0,
          marginBottom: 16,
        }}
      >
        <h2 style={styles.productTitle}>📡 Live Inventory Station Status</h2>

        <div style={styles.headerActions}>
          <button
            style={{
              ...styles.viewModeButton,
              ...(inventoryReportMode === "my" ? styles.viewModeButtonActive : {}),
            }}
            onClick={() => setInventoryReportMode("my")}
          >
            👤 My Report
          </button>

          <button
            style={{
              ...styles.viewModeButton,
              ...(inventoryReportMode === "summary" ? styles.viewModeButtonActive : {}),
            }}
            onClick={() => setInventoryReportMode("summary")}
          >
            🌍 Summary Report
          </button>

          <button
            style={styles.backButton}
            onClick={() => refreshMakeInventoryData(makeInventoryShip || userShip)}
            disabled={inventoryLoading || reportBusy}
          >
            🔄 Refresh
          </button>

          {isAdmin && (
  <button
    style={styles.deleteButton}
    onClick={resetInventoryRun}
    disabled={inventoryLoading || reportBusy}
  >
    ♻️ Reset Inventory
  </button>
)}

          <button
            style={styles.backButton}
            onClick={printInventorySummary}
            disabled={reportBusy}
          >
            {reportBusy ? "Preparing..." : "🖨️ Print / PDF"}
          </button>

          {inventoryReportMode === "my" && (
            <button
              style={styles.deleteButton}
              onClick={clearMyInventory}
              disabled={inventoryLoading || reportBusy}
            >
              🧹 Clear My Report
            </button>
          )}

          {isAdmin && inventoryReportMode === "summary" && (
  <button
    style={styles.deleteButton}
    onClick={clearShipInventory}
    disabled={inventoryLoading || reportBusy}
  >
    🧹 Clear Ship Records
  </button>
)}

          {inventoryReportMode === "summary" ? (
  isAdmin ? (
    <button
      style={styles.primaryButton}
      onClick={generateFinalInventoryReport}
      disabled={reportBusy || !allStationsSubmitted}
    >
      📥 Generate Final Report
    </button>
  ) : null
) : (
  <button
    style={styles.primaryButton}
    onClick={exportInventorySummaryToExcel}
    disabled={reportBusy}
  >
    📥 Export Excel
  </button>
)}
        </div>
      </div>

      <div style={styles.infoBox}>
        <div>
          Current station:{" "}
          <strong>
            {inventoryStation || "Not selected"} -{" "}
            {currentStationProgress?.statusLabel || "Not Started"}
          </strong>
        </div>

        <div>
          My live count:{" "}
          <strong>
            {myReportRows.length} / {makeInventoryItems.length || 0}
          </strong>
        </div>

        <div>
          Stations started: <strong>{startedStations}</strong>
        </div>

        <div>
          Stations submitted:{" "}
          <strong>
            {submittedStations} / {stationProgressRows.length}
          </strong>
        </div>

        {allStationsSubmitted ? (
          <div style={{ color: "#2e7d32", fontWeight: "bold" }}>
            ✅ All stations submitted. Final report is ready.
          </div>
        ) : (
          <div style={{ color: "#8a5a00", fontWeight: "bold" }}>
            ⏳ Waiting for all stations to submit before final report.
          </div>
        )}
      </div>

      <div style={styles.stationStatusGrid}>
        {stationProgressRows.map((item) => (
          <div
            key={item.station}
            style={{
              ...styles.stationStatusCard,
              ...(item.status === "started" ? styles.stationStatusStarted : {}),
              ...(item.status === "submitted" ? styles.stationStatusSubmitted : {}),
            }}
          >
            <strong>{item.station}</strong>
            <span>{item.statusLabel}</span>
            <small>
              Counted items: {item.countedItems} / {makeInventoryItems.length || 0}
            </small>
            {item.userName && <small>User: {item.userName}</small>}
          </div>
        ))}
      </div>
    </section>
  );
}

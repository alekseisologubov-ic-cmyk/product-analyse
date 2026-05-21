"use client";

import React from "react";

export default function EquipmentInventoryOptions({
  styles,
  userShip,
  activeEquipmentDepartmentLabel,
  getShipDisplayName,
  onBack,
  onOpenInUse,
  onOpenWarehouse,
  onOpenMakeInventory,
}) {
  const shipDisplayName =
    typeof getShipDisplayName === "function"
      ? getShipDisplayName(userShip)
      : userShip;

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={styles.headerLogo}
        />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Back
          </button>

          <div style={styles.shipBadge}>🚢 {shipDisplayName}</div>
        </div>
      </header>

      <section style={styles.card}>
        <h2 style={styles.cardTitle}>
          📊 {activeEquipmentDepartmentLabel} Inventory
        </h2>

        <div style={styles.moduleGrid}>
          <button style={styles.moduleCard} onClick={onOpenInUse}>
            <div style={styles.moduleIcon}>✅</div>
            <strong>Inventory in Use</strong>
            <span>Compare muster list against in-use inventory</span>
          </button>

          <button style={styles.moduleCard} onClick={onOpenWarehouse}>
            <div style={styles.moduleIcon}>🏬</div>
            <strong>Inventory Warehouse</strong>
            <span>Par, on hand, future order and suggested order</span>
          </button>

          <button style={styles.moduleCard} onClick={onOpenMakeInventory}>
            <div style={styles.moduleIcon}>📝</div>
            <strong>Make Inventory</strong>
            <span>Multi-user counts, my report and ship summary</span>
          </button>
        </div>
      </section>
    </main>
  );
}

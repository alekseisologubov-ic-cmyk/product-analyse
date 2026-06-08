"use client";

import React from "react";

export default function EquipmentDepartmentOptions({
  styles,
  userShip,
  equipmentDepartment,
  activeEquipmentDepartmentLabel,
  activeEquipmentDepartmentIcon,
  getShipDisplayName,
  onBackToModules,
  onOpenProductDashboard,
  onOpenNextOrder,
  onOpenAllergen,
  onOpenTraining,
  onOpenMuster,
  onOpenBreakageReport,
  onOpenInventory,
  onOpenTemperature,
}) {
  const title =
    equipmentDepartment === "restaurant"
      ? "Restaurant Options"
      : `${activeEquipmentDepartmentLabel} Options`;

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
          <button style={styles.backButton} onClick={onBackToModules}>
            ← Modules
          </button>

          <div style={styles.shipBadge}>🚢 {shipDisplayName}</div>
        </div>
      </header>

      <section style={styles.card}>
        <h2 style={styles.cardTitle}>
          {activeEquipmentDepartmentIcon} {title}
        </h2>

        <div style={styles.moduleGrid}>
          {equipmentDepartment === "culinary" && (
            <button style={styles.moduleCard} onClick={onOpenProductDashboard}>
              <div style={styles.moduleIcon}>📊</div>
              <strong>Product Dashboard</strong>
              <span>Consumption, recipes, templates and product reports</span>
            </button>
          )}

          {(equipmentDepartment === "culinary" ||
            equipmentDepartment === "bar") && (
            <button style={styles.moduleCard} onClick={onOpenNextOrder}>
              <div style={styles.moduleIcon}>🛒</div>
              <strong>Generate Next Order</strong>
              <span>
                Upload the latest order file and calculate suggested next-order
                quantities
              </span>
            </button>
          )}

          {(equipmentDepartment === "culinary" ||
            equipmentDepartment === "restaurant") && (
            <button style={styles.moduleCard} onClick={onOpenAllergen}>
              <div style={styles.moduleIcon}>🧬</div>
              <strong>Allergen Matrix</strong>
              <span>
                Recipe, ingredient, sub-recipe, and possible hidden allergen
                review by venue
              </span>
            </button>
          )}

          {equipmentDepartment === "culinary" && (
            <button style={styles.moduleCard} onClick={onOpenTraining}>
              <div style={styles.moduleIcon}>🎓</div>
              <strong>Training</strong>
              <span>
                Monthly station training tracker by crew assignment and
                training links
              </span>
            </button>
          )}

          <button style={styles.moduleCard} onClick={onOpenMuster}>
            <div style={styles.moduleIcon}>📋</div>
            <strong>Equipment Master List</strong>
            <span>Grouped by sheet and subcategory</span>
          </button>

          <button style={styles.moduleCard} onClick={onOpenBreakageReport}>
            <div style={styles.moduleIcon}>🧾</div>
            <strong>Breakage Report</strong>
            <span>Report broken equipment from this department master list</span>
          </button>

          <button style={styles.moduleCard} onClick={onOpenInventory}>
            <div style={styles.moduleIcon}>📦</div>
            <strong>Equipment Inventory</strong>
            <span>Inventory in use, warehouse stock and make inventory</span>
          </button>

          {equipmentDepartment === "culinary" && (
            <button style={styles.moduleCard} onClick={onOpenTemperature}>
              <div style={styles.moduleIcon}>🌡️</div>
              <strong>Take Temperature</strong>
              <span>Take food temperature pictures and save by date</span>
            </button>
          )}
        </div>
      </section>
    </main>
  );
}

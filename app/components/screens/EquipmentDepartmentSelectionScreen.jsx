"use client";

import React from "react";

export default function EquipmentDepartmentSelectionScreen(props) {
  const {
    styles,
    getShipDisplayName,
    logUsageEvent,
    setEquipmentDepartment,
    setEquipmentMode,
    setModule,
    userShip,
  } = props;

  return (
        <main style={styles.page}>
          <header style={styles.header}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
            <div style={styles.headerActions}>
              <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
              <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
            </div>
          </header>

          <section style={styles.card}>
            <h2 style={styles.cardTitle}>🍽️ Equipment Department</h2>
            <p style={styles.emptyText}>Choose which operation area you want to work with.</p>

            <div style={styles.moduleGrid}>
              <button
                style={styles.moduleCard}
                onClick={() => {
                  setEquipmentDepartment("culinary");
                  setEquipmentMode("");
                  logUsageEvent("equipment_department_opened", { module: "equipment_culinary", ship: userShip });
                }}
              >
                <div style={styles.moduleIcon}>👨‍🍳</div>
                <strong>Culinary</strong>
                <span>Current equipment tools: master list, inventory in use, warehouse and make inventory.</span>
              </button>

              <button
                style={styles.moduleCard}
                onClick={() => {
                  setEquipmentDepartment("bar");
                  setEquipmentMode("");
                  logUsageEvent("equipment_department_opened", { module: "equipment_bar", ship: userShip });
                }}
              >
                <div style={styles.moduleIcon}>🍸</div>
                <strong>Bar</strong>
                <span>Master list and inventory tools for Bar equipment.</span>
              </button>

              <button
                style={styles.moduleCard}
                onClick={() => {
                  setEquipmentDepartment("restaurant");
                  setEquipmentMode("");
                  logUsageEvent("equipment_department_opened", { module: "equipment_restaurant", ship: userShip });
                }}
              >
                <div style={styles.moduleIcon}>🍽️</div>
                <strong>Restaurant</strong>
                <span>Master list and inventory tools for Restaurant equipment.</span>
              </button>
            </div>
          </section>
        </main>
      );
}

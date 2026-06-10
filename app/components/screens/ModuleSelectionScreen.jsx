"use client";

import React from "react";

export default function ModuleSelectionScreen(props) {
  const {
    styles,
    getShipDisplayName,
    isAdmin,
    logUsageEvent,
    setEquipmentDepartment,
    setEquipmentMode,
    setModule,
    setProductMode,
    userShip,
  } = props;

  return (
        <main style={styles.page}>
          <header style={styles.header}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
          </header>

          <section style={styles.card}>
            <h2 style={styles.cardTitle}>🧭 Select Module</h2>

            <div style={styles.moduleGrid}>
    {isAdmin && (
      <button
        style={styles.moduleCard}
        onClick={() => {
          setModule("admin");
          logUsageEvent("module_opened", {
            module: "admin_dashboard",
            ship: userShip,
          });
        }}
      >
        <div style={styles.moduleIcon}>🛡️</div>
        <strong>Admin Dashboard</strong>
        <span>Usage, inventory status, logs, and admin tools</span>
      </button>
    )}

    <button
      style={styles.moduleCard}
      onClick={() => {
        setModule("equipment");
        setEquipmentDepartment("culinary");
        setEquipmentMode("");
        setProductMode("");
        logUsageEvent("department_opened", {
          module: "department_culinary",
          equipmentDepartment: "culinary",
          ship: userShip,
        });
      }}
    >
      <div style={styles.moduleIcon}>👨‍🍳</div>
      <strong>Culinary</strong>
      <span>Product dashboard, next order, inventory, master list and temperature checks</span>
    </button>

    <button
      style={styles.moduleCard}
      onClick={() => {
        setModule("equipment");
        setEquipmentDepartment("bar");
        setEquipmentMode("");
        setProductMode("");
        logUsageEvent("department_opened", {
          module: "department_bar",
          equipmentDepartment: "bar",
          ship: userShip,
        });
      }}
    >
      <div style={styles.moduleIcon}>🍸</div>
      <strong>Bar</strong>
      <span>Generate next order, inventory and master list for Bar equipment</span>
    </button>

    <button
      style={styles.moduleCard}
      onClick={() => {
        setModule("equipment");
        setEquipmentDepartment("restaurant");
        setEquipmentMode("");
        setProductMode("");
        logUsageEvent("department_opened", {
          module: "department_restaurant",
          equipmentDepartment: "restaurant",
          ship: userShip,
        });
      }}
    >
      <div style={styles.moduleIcon}>🍽️</div>
      <strong>Restaurant</strong>
     <span>Inventory and master list for Restaurant equipment</span>
    </button>
  </div>
          </section>
        </main>
      );
}

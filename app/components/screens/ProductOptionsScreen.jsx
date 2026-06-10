"use client";

import React from "react";

export default function ProductOptionsScreen(props) {
  const {
    styles,
    getShipDisplayName,
    logUsageEvent,
    setFmlLowRows,
    setFmlLowSearch,
    setFmlMissingRows,
    setFmlMissingSearch,
    setModule,
    setNextOrderFilter,
    setNextOrderMessage,
    setNextOrderRows,
    setNextOrderSearch,
    setNextOrderView,
    setProductMode,
    setProductReportView,
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
            <h2 style={styles.cardTitle}>📦 Product Options</h2>
            <p style={styles.emptyText}>Choose whether you want to review products or generate the next order.</p>

            <div style={styles.moduleGrid}>
              <button
                style={styles.moduleCard}
                onClick={() => {
                  setProductMode("dashboard");
                  setProductReportView("main");
                  logUsageEvent("product_option_opened", { module: "product_dashboard", ship: userShip });
                }}
              >
                <div style={styles.moduleIcon}>📊</div>
                <strong>Product Dashboard</strong>
                <span>Use the existing product dashboard with consumption, recipes, templates, allergens and reports.</span>
              </button>

              <button
                style={styles.moduleCard}
                onClick={() => {
                  setProductMode("nextorder");
                  setNextOrderRows([]);
                  setFmlMissingRows([]);
          setFmlLowRows([]);
                  setNextOrderSearch("");
                  setFmlMissingSearch("");
          setFmlLowSearch("");
                  setNextOrderFilter("all");
                  setNextOrderView("order");
                  setNextOrderMessage("");
                  logUsageEvent("product_option_opened", { module: "generate_next_order", ship: userShip });
                }}
              >
                <div style={styles.moduleIcon}>🛒</div>
                <strong>Generate Next Order</strong>
                <span>Use the attached ERP template, upload the latest order workbook, and calculate suggested next-order quantities.</span>
              </button>
            </div>
          </section>
        </main>
      );
}

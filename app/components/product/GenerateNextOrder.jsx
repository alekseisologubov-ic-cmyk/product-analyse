"use client";

import React from "react";

export default function GenerateNextOrder({
  styles,
  userShip,
  onBack,
}) {
  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Product Options
          </button>

          <div style={styles.shipBadge}>🚢 {userShip}</div>
        </div>
      </header>

      <section style={styles.card}>
        <h2 style={styles.cardTitle}>🛒 Generate Next Order</h2>
        <p style={styles.emptyText}>
          Generate Next Order module is loading as a separate component.
        </p>
      </section>
    </main>
  );
}

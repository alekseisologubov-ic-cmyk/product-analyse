"use client";

import React from "react";

export default function AdminScreen(props) {
  const {
    styles,
    AdminDashboard,
    getShipDisplayName,
    isAdmin,
    normalizedUserEmail,
    setModule,
    supabase,
    userShip,
  } = props;

  if (!isAdmin) {
        return (
          <main style={styles.page}>
            <header style={styles.header}>
              <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
              <div style={styles.headerActions}>
                <button style={styles.backButton} onClick={() => setModule("")}>
                  ← Modules
                </button>
                <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
              </div>
            </header>

            <section style={styles.card}>
              <h2 style={styles.cardTitle}>Access denied</h2>
              <p style={styles.emptyText}>
                This dashboard is available only for admin users.
              </p>
            </section>
          </main>
        );
      }

      return (
        <AdminDashboard
          styles={styles}
          supabase={supabase}
          userEmail={normalizedUserEmail}
          userShip={userShip}
          onBack={() => setModule("")}
        />
      );
}

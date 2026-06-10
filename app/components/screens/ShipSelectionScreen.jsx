"use client";

import React from "react";

export default function ShipSelectionScreen(props) {
  const {
    styles,
    getShipDisplayName,
    logUsageEvent,
    normalizeAppEmail,
    resetUserEmail,
    setLoggedIn,
    setUserShip,
    setWelcomeStarted,
    SHIPS,
    userEmail,
    userShip,
  } = props;

  return (
        <main style={styles.page}>
          <section style={styles.loginCard}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
            <h1 style={styles.title}>Choose Your Ship</h1>
            <p style={styles.subtitle}>Select your vessel to start the dashboard.</p>

            <div style={styles.infoBox}>
              <div>👤 Signed in as: <strong>{normalizeAppEmail(userEmail)}</strong></div>
              <button type="button" style={styles.inlineLinkButton} onClick={resetUserEmail}>Use different email</button>
            </div>

            <label style={styles.label}>🚢 Select your ship</label>
            <select value={userShip} onChange={(e) => setUserShip(e.target.value)} style={styles.select}>
              <option value="">Choose ship</option>
              {SHIPS.map((ship) => (
    <option key={ship} value={ship}>
      {getShipDisplayName(ship)}
    </option>
  ))}
            </select>

            <button
              style={styles.primaryButton}
              onClick={() => {
                if (!userShip) return;
                logUsageEvent("ship_selected", { ship: userShip, module: "welcome", userEmail: normalizeAppEmail(userEmail) });
                setLoggedIn(true);
              }}
            >
              Continue
            </button>

            <button style={styles.backButton} onClick={() => setWelcomeStarted(false)}>
              ← Back to Start
            </button>
          </section>
        </main>
      );
}

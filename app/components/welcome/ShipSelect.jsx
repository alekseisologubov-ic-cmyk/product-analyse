"use client";

import React from "react";

export default function ShipSelect({
  styles,
  ships,
  userEmail,
  userShip,
  setUserShip,
  onContinue,
  onBack,
  onUseDifferentEmail,
}) {
  return (
    <main style={styles.page}>
      <section style={styles.loginCard}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />

        <h1 style={styles.title}>Choose Your Ship</h1>
        <p style={styles.subtitle}>Select your vessel to start the dashboard.</p>

        <div style={styles.infoBox}>
          <div>
            👤 Signed in as: <strong>{userEmail}</strong>
          </div>

          <button style={styles.inlineLinkButton} onClick={onUseDifferentEmail}>
            Use different email
          </button>
        </div>

        <label style={styles.label}>🚢 Select your ship</label>
        <select
          value={userShip}
          onChange={(e) => setUserShip(e.target.value)}
          style={styles.select}
        >
          <option value="">Choose ship</option>
          {ships.map((ship) => (
            <option key={ship} value={ship}>
              {ship}
            </option>
          ))}
        </select>

        <button style={styles.primaryButton} onClick={onContinue}>
          Continue
        </button>

        <button style={styles.backButton} onClick={onBack}>
          ← Back to Start
        </button>
      </section>
    </main>
  );
}

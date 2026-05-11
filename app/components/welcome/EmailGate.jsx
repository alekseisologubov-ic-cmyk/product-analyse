"use client";

import React from "react";

export default function EmailGate({
  styles,
  userEmail,
  setUserEmail,
  emailError,
  setEmailError,
  rememberEmail,
  setRememberEmail,
  onContinue,
  onBack,
}) {
  return (
    <main style={styles.page}>
      <section style={styles.loginCard}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />

        <h1 style={styles.title}>Enter Your Email</h1>
        <p style={styles.subtitle}>Use your Virgin Voyages email to continue.</p>

        <label style={styles.label}>✉️ Virgin Voyages email</label>
        <input
          type="email"
          value={userEmail}
          onChange={(e) => {
            setUserEmail(e.target.value);
            setEmailError("");
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") onContinue();
          }}
          placeholder="name@virginvoyages.com"
          style={styles.searchInput}
          autoComplete="email"
        />

        {emailError && <div style={styles.emailError}>{emailError}</div>}

        <label style={styles.checkboxRow}>
          <input
            type="checkbox"
            checked={rememberEmail}
            onChange={(e) => setRememberEmail(e.target.checked)}
          />
          <span>Remember me on this device</span>
        </label>

        <div style={styles.infoBox}>
          <div>
            🔒 Only emails ending with <strong>@virginvoyages.com</strong> are allowed.
          </div>
          <div>📊 Usage tracking will be connected to this email.</div>
        </div>

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

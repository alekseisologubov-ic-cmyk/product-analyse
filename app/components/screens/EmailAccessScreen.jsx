"use client";

import React from "react";

export default function EmailAccessScreen(props) {
  const {
    styles,
    dashboardRegionalMatchedCount,
    emailCodeSent,
    emailError,
    emailMessage,
    emailOtpCode,
    filteredProductCostReportRows,
    otpLoading,
    productCostReportRows,
    productCostReportRowsWithRegionalPar,
    rememberEmail,
    selectedRegionalConsumptionRegion,
    sendAccessCode,
    setEmailCodeSent,
    setEmailError,
    setEmailMessage,
    setEmailOtpCode,
    setRememberEmail,
    setUserEmail,
    setWelcomeStarted,
    userEmail,
    userShip,
    verifyAccessCode,
    viewMode,
    YEARLY_REGION_ALL,
  } = props;

  return (
        <main style={styles.page}>
          <section style={styles.loginCard}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
            <h1 style={styles.title}>Email Access Code</h1>
            <p style={styles.subtitle}>Enter your Virgin Voyages email and access code.</p>

            <label style={styles.label}>✉️ Virgin Voyages email</label>
            <input
              type="email"
              value={userEmail}
              disabled={otpLoading || emailCodeSent}
              onChange={(e) => {
                setUserEmail(e.target.value);
                setEmailError("");
                setEmailMessage("");
              }}
              onKeyDown={(e) => {
                if (e.key === "Enter" && !emailCodeSent) sendAccessCode();
              }}
              placeholder="name@virginvoyages.com"
              style={styles.searchInput}
              autoComplete="email"
            />

            {emailCodeSent && (
              <>
                <label style={styles.label}>🔐 Access code</label>
                <input
                  type="password"
                  inputMode="numeric"
                  value={emailOtpCode}
                  disabled={otpLoading}
                  onChange={(e) => {
                    setEmailOtpCode(e.target.value);
                    setEmailError("");
                    setEmailMessage("");
                  }}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") verifyAccessCode();
                  }}
                  placeholder="Enter access code..."
                  style={styles.searchInput}
                  autoComplete="one-time-code"
                />
              </>
            )}

            {emailError && <div style={styles.emailError}>{emailError}</div>}
            {emailMessage && <div style={{ ...styles.infoBox, color: "#2e7d32", fontWeight: "bold" }}>{emailMessage}</div>}

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
      📦 Products in report:{" "}
      <strong>{filteredProductCostReportRows.length}</strong> / {productCostReportRows.length}
    </div>

    <div>
      🚢 View: <strong>{viewMode === "single" ? userShip : "All Ships"}</strong>
    </div>

    <div>
      🌎 Regional par region:{" "}
      <strong>
        {selectedRegionalConsumptionRegion === YEARLY_REGION_ALL
          ? "All regions"
          : selectedRegionalConsumptionRegion}
      </strong>
    </div>

    <div>
      📈 Regional matches:{" "}
      <strong>
        {dashboardRegionalMatchedCount} / {productCostReportRowsWithRegionalPar.length}
      </strong>
    </div>

    <div>
      📘 Source columns: C Venue, I/L/O/R Quantity, J/M/P/S Total Cost, H/K/N/Q Unit Price.
    </div>
  </div>

            {!emailCodeSent ? (
              <button style={styles.primaryButton} onClick={sendAccessCode} disabled={otpLoading}>
                Continue
              </button>
            ) : (
              <>
                <button style={styles.primaryButton} onClick={verifyAccessCode} disabled={otpLoading}>
                  Verify Code
                </button>
                <button
                  style={styles.backButton}
                  onClick={() => {
                    setEmailCodeSent(false);
                    setEmailOtpCode("");
                    setEmailError("");
                    setEmailMessage("");
                  }}
                  disabled={otpLoading}
                >
                  Change Email
                </button>
              </>
            )}

            <button
              style={styles.backButton}
              onClick={() => {
                setWelcomeStarted(false);
                setEmailError("");
                setEmailMessage("");
                setEmailOtpCode("");
                setEmailCodeSent(false);
              }}
              disabled={otpLoading}
            >
              ← Back to AHOY
            </button>
          </section>
        </main>
      );
}

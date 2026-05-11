"use client";

import React from "react";

export default function EmailGate({
  styles,
  userEmail = "",
  setUserEmail = () => {},
  emailError = "",
  setEmailError = () => {},
  emailMessage = "",
  setEmailMessage = () => {},
  emailOtpCode = "",
  setEmailOtpCode = () => {},
  emailCodeSent = false,
  setEmailCodeSent = () => {},
  rememberEmail = false,
  setRememberEmail = () => {},
  otpLoading = false,
  onSendCode = () => {},
  onVerifyCode = () => {},
  onBack = () => {},
}) {
  return (
    <main style={styles.page}>
      <section style={styles.loginCard}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={styles.logo}
        />

        <h1 style={styles.title}>Email Access Code</h1>

        <p style={styles.subtitle}>
          Enter your Virgin Voyages email and verify the access code.
        </p>

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
            if (e.key === "Enter" && !emailCodeSent) {
              onSendCode();
            }
          }}
          placeholder="name@virginvoyages.com"
          style={styles.searchInput}
          autoComplete="email"
        />

        {emailCodeSent && (
          <>
            <label style={styles.label}>🔐 Access code</label>

            <input
              type="text"
              inputMode="numeric"
              value={emailOtpCode}
              disabled={otpLoading}
              onChange={(e) => {
                setEmailOtpCode(e.target.value);
                setEmailError("");
                setEmailMessage("");
              }}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  onVerifyCode();
                }
              }}
              placeholder="Enter code from email..."
              style={styles.searchInput}
              autoComplete="one-time-code"
            />
          </>
        )}

        {emailError && (
          <div style={styles.emailError}>
            {emailError}
          </div>
        )}

        {emailMessage && (
          <div
            style={{
              ...styles.infoBox,
              color: "#2e7d32",
              fontWeight: "bold",
            }}
          >
            {emailMessage}
          </div>
        )}

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
            🔒 Only emails ending with{" "}
            <strong>@virginvoyages.com</strong> are allowed.
          </div>
          <div>
            📩 A one-time access code will be sent to your email.
          </div>
          <div>
            📊 Usage tracking will be connected to your verified email.
          </div>
        </div>

        {!emailCodeSent ? (
          <button
            style={styles.primaryButton}
            onClick={onSendCode}
            disabled={otpLoading}
          >
            {otpLoading ? "Sending..." : "Send Access Code"}
          </button>
        ) : (
          <>
            <button
              style={styles.primaryButton}
              onClick={onVerifyCode}
              disabled={otpLoading}
            >
              {otpLoading ? "Verifying..." : "Verify Code"}
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

            <button
              style={styles.backButton}
              onClick={onSendCode}
              disabled={otpLoading}
            >
              Resend Code
            </button>
          </>
        )}

        <button
          style={styles.backButton}
          onClick={onBack}
          disabled={otpLoading}
        >
          ← Back to AHOY
        </button>
      </section>
    </main>
  );
}

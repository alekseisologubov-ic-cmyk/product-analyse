"use client";

import React from "react";

export default function WelcomePage({ styles, onStart }) {
  return (
    <main style={styles.welcomePage}>
      <style>{`
        @keyframes vvMarquee {
          0% { transform: translateX(0); }
          100% { transform: translateX(-50%); }
        }

        @keyframes vvGlow {
          0%, 100% { box-shadow: 0 18px 50px rgba(0,0,0,0.18); }
          50% { box-shadow: 0 22px 70px rgba(176,0,32,0.28); }
        }
      `}</style>

      <section style={styles.welcomeHero}>
        <div style={styles.welcomeGlowCard}>
          <img
            src="/virgin-logo.png"
            alt="Virgin Voyages"
            style={styles.welcomeLogo}
          />

          <div style={styles.runningLineWrapper}>
            <div style={styles.runningLineTrack}>
              <span style={styles.runningLineText}>
                Use it • Save Time • Be the Reason Someone Smiles Today •
              </span>
              <span style={styles.runningLineText}>
                Use it • Save Time • Be the Reason Someone Smiles Today •
              </span>
              <span style={styles.runningLineText}>
                Use it • Save Time • Be the Reason Someone Smiles Today •
              </span>
              <span style={styles.runningLineText}>
                Use it • Save Time • Be the Reason Someone Smiles Today •
              </span>
            </div>
          </div>

          <button
            style={styles.ahoyStartButton}
            onClick={onStart}
            aria-label="Start"
          >
            AHOY
          </button>

          <p style={styles.welcomeSubtitle}>
            Press AHOY to start.
          </p>
        </div>
      </section>
    </main>
  );
}

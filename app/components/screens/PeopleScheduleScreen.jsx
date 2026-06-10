"use client";

import React from "react";

export default function PeopleScheduleScreen(props) {
  const {
    styles,
    Suspense,
    logUsageEvent,
    PeopleScheduleModule,
    setModule,
    userShip,
  } = props;

  return (
        <Suspense
          fallback={
            <main style={styles.page}>
              <header style={styles.header}>
                <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
                <div style={styles.headerActions}>
                  <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
                  <div style={styles.shipBadge}>🚢 Loading</div>
                </div>
              </header>

              <section style={styles.card}>
                <h2 style={styles.cardTitle}>👥 Loading People & Schedule...</h2>
                <p style={styles.emptyText}>Preparing the rotation planner only when needed.</p>
              </section>
            </main>
          }
        >
          <PeopleScheduleModule
            userShip={userShip}
            onBack={() => setModule("")}
            styles={styles}
            logUsageEvent={logUsageEvent}
          />
        </Suspense>
      );
}

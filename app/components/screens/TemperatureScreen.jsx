"use client";

import React from "react";

export default function TemperatureScreen(props) {
  const {
    styles,
    Suspense,
    equipmentDepartment,
    isAdmin,
    logUsageEvent,
    normalizeAppEmail,
    setEquipmentMode,
    setModule,
    supabase,
    TemperatureCheckModule,
    userEmail,
    userShip,
  } = props;

  return (
        <Suspense
          fallback={
            <main style={styles.page}>
              <section style={styles.card}>
                <h2 style={styles.cardTitle}>🌡️ Loading Take Temperature...</h2>
                <p style={styles.emptyText}>Preparing temperature photo tools.</p>
              </section>
            </main>
          }
        >
          <TemperatureCheckModule
            styles={styles}
            supabase={supabase}
            userShip={userShip}
            userEmail={normalizeAppEmail(userEmail)}
  isAdmin={isAdmin}
            onBack={() => {
    if (equipmentDepartment) {
      setModule("equipment");
      setEquipmentMode("");
      return;
    }

    setModule("");
  }}
            logUsageEvent={logUsageEvent}
          />
        </Suspense>
      );
}

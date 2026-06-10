"use client";

import React from "react";

export default function TrainingScreen(props) {
  const {
    styles,
    Suspense,
    AppProvider,
    equipmentDepartment,
    isAdmin,
    logUsageEvent,
    normalizeAppEmail,
    setEquipmentMode,
    setModule,
    supabase,
    TrainingModule,
    userEmail,
    userShip,
  } = props;

  return (
        <Suspense
          fallback={
            <main style={styles.page}>
              <section style={styles.card}>
                <h2 style={styles.cardTitle}>🎓 Loading Training Module...</h2>
                <p style={styles.emptyText}>Preparing station training tracker.</p>
              </section>
            </main>
          }
        >
          <AppProvider
    value={{
      supabase,
      userShip,
      userEmail: normalizeAppEmail(userEmail),
      isAdmin,
      logUsageEvent,
    }}
  >
    <TrainingModule
      styles={styles}
      onBack={() => {
        if (equipmentDepartment) {
          setModule("equipment");
          setEquipmentMode("");
          return;
        }

        setModule("");
      }}
    />
  </AppProvider>
        </Suspense>
      );
}

"use client";

import React from "react";

export default function BreakageScreen(props) {
  const {
    styles,
    Suspense,
    activeEquipmentDepartmentLabel,
    BreakageReportModule,
    equipmentDepartment,
    getShipDisplayName,
    logUsageEvent,
    normalizeAppEmail,
    setEquipmentMode,
    setModule,
    supabase,
    userEmail,
    userShip,
  } = props;

  return (
      <Suspense
        fallback={
          <main style={styles.page}>
            <section style={styles.card}>
              <h2 style={styles.cardTitle}>🧾 Loading Breakage Report...</h2>
              <p style={styles.emptyText}>
                Preparing {activeEquipmentDepartmentLabel} equipment master list.
              </p>
            </section>
          </main>
        }
      >
        <BreakageReportModule
          styles={styles}
          supabase={supabase}
          userShip={userShip}
          userEmail={normalizeAppEmail(userEmail)}
          equipmentDepartment={equipmentDepartment}
          activeEquipmentDepartmentLabel={activeEquipmentDepartmentLabel}
          getShipDisplayName={getShipDisplayName}
          logUsageEvent={logUsageEvent}
          onBack={() => {
            setModule("equipment");
            setEquipmentMode("");
          }}
        />
      </Suspense>
    );
}

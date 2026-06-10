"use client";

import React from "react";

export default function GenerateNextOrderScreen(props) {
  const {
    styles,
    Suspense,
    equipmentDepartment,
    GenerateNextOrder,
    logUsageEvent,
    recipeRows,
    regionalParBufferPercent,
    selectedRegionalConsumptionRegion,
    setEquipmentMode,
    setModule,
    setProductMode,
    setRegionalParBufferPercent,
    setSelectedRegionalConsumptionRegion,
    setYearlyRegionalConsumption,
    setYearlyRegionalFileName,
    userShip,
    yearlyRegionalConsumption,
    yearlyRegionalFileName,
  } = props;

  return (
        <Suspense
          fallback={
            <main style={styles.page}>
              <section style={styles.card}>
                <h2 style={styles.cardTitle}>🛒 Loading Generate Next Order...</h2>
                <p style={styles.emptyText}>Preparing order tools.</p>
              </section>
            </main>
          }
        >
          <GenerateNextOrder
    styles={styles}
    userShip={userShip}
    onBack={() => {
      if (equipmentDepartment) {
        setModule("equipment");
        setEquipmentMode("");
        setProductMode("");
        return;
      }

      setProductMode("");
    }}
    logUsageEvent={logUsageEvent}
    yearlyRegionalConsumption={yearlyRegionalConsumption}
    setYearlyRegionalConsumption={setYearlyRegionalConsumption}
    yearlyRegionalFileName={yearlyRegionalFileName}
    setYearlyRegionalFileName={setYearlyRegionalFileName}
    selectedRegionalConsumptionRegion={selectedRegionalConsumptionRegion}
    setSelectedRegionalConsumptionRegion={setSelectedRegionalConsumptionRegion}
    regionalParBufferPercent={regionalParBufferPercent}
    setRegionalParBufferPercent={setRegionalParBufferPercent}
    recipeRows={recipeRows}
  />
        </Suspense>
      );
}

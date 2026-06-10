"use client";

import React from "react";

export default function AllergenScreen(props) {
  const {
    styles,
    Suspense,
    AppProvider,
    AllergenModule,
    equipmentDepartment,
    isAdmin,
    logUsageEvent,
    normalizeAppEmail,
    recipeRows,
    setEquipmentMode,
    setModule,
    setRecipeRows,
    supabase,
    userEmail,
    userShip,
  } = props;

  return (
      <Suspense
        fallback={
          <main style={styles.page}>
            <section style={styles.card}>
              <h2 style={styles.cardTitle}>🧬 Loading Allergen Module...</h2>
              <p style={styles.emptyText}>Preparing recipe allergen matrix.</p>
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
          <AllergenModule
    key={`${userShip}-${normalizeAppEmail(userEmail)}`}
    styles={styles}
    supabase={supabase}
    userShip={userShip}
    userEmail={normalizeAppEmail(userEmail)}
    isAdmin={isAdmin}
    recipeRows={recipeRows}
    setRecipeRows={setRecipeRows}
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
        </AppProvider>
      </Suspense>
    );
}

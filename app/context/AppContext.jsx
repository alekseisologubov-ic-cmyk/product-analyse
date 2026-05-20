"use client";

import React, { createContext, useContext, useMemo } from "react";

const AppContext = createContext(null);

export const normalizeContextEmail = (value) =>
  String(value || "").trim().toLowerCase();

export const getCulinaryAdminShipFromEmail = (value) => {
  const email = normalizeContextEmail(value);
  const localPart = email.split("@")[0] || "";

  const adminShipMap = {
    "val.cul.admin": "VL",
    "vl.cul.admin": "VL",
    "valiant.cul.admin": "VL",

    "sc.cul.admin": "SC",
    "scl.cul.admin": "SC",
    "scarlet.cul.admin": "SC",

    "rl.cul.admin": "RL",
    "res.cul.admin": "RL",
    "resilient.cul.admin": "RL",

    "brl.cul.admin": "BRL",
    "brilliant.cul.admin": "BRL",
  };

  return adminShipMap[localPart] || "";
};

export function AppProvider({ value = {}, children }) {
  const {
    supabase = null,
    userShip = "",
    userEmail = "",
    isAdmin = false,
    logUsageEvent = () => {},
  } = value;

  const contextValue = useMemo(() => {
    const normalizedUserEmail = normalizeContextEmail(userEmail);
    const culinaryAdminShip = getCulinaryAdminShipFromEmail(normalizedUserEmail);
    const isShipCulinaryAdmin =
      Boolean(culinaryAdminShip) && culinaryAdminShip === userShip;

    return {
      supabase,
      userShip,
      userEmail: normalizedUserEmail,
      normalizedUserEmail,
      isAdmin,
      culinaryAdminShip,
      isShipCulinaryAdmin,
      canManageTraining: Boolean(isAdmin || isShipCulinaryAdmin),
      logUsageEvent,
    };
  }, [supabase, userShip, userEmail, isAdmin, logUsageEvent]);

  return (
    <AppContext.Provider value={contextValue}>
      {children}
    </AppContext.Provider>
  );
}

export function useAppContext() {
  const context = useContext(AppContext);

  if (!context) {
    return {
      supabase: null,
      userShip: "",
      userEmail: "",
      normalizedUserEmail: "",
      isAdmin: false,
      culinaryAdminShip: "",
      isShipCulinaryAdmin: false,
      canManageTraining: false,
      logUsageEvent: () => {},
    };
  }

  return context;
}

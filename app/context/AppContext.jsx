"use client";

import React, { createContext, useContext, useMemo } from "react";

const AppContext = createContext(null);

export const normalizeContextEmail = (value) =>
  String(value || "").trim().toLowerCase();

export const SHIP_LABELS = {
  SC: "Scarlet",
  VL: "Valiant",
  BRL: "Brilliant",
  RL: "Resilient",
};

export const getContextShipDisplayName = (shipCode) =>
  SHIP_LABELS[shipCode] || shipCode || "";

export const getCulinaryAdminShipFromEmail = (value) => {
  const email = normalizeContextEmail(value);
  const localPart = email.split("@")[0] || "";

  const adminShipMap = {
    // Scarlet
    "scarlet.cul.admin": "SC",
    "sc.cul.admin": "SC",
    "scl.cul.admin": "SC",

    // Valiant
    "valiant.cul.admin": "VL",
    "val.cul.admin": "VL",
    "vl.cul.admin": "VL",

    // Brilliant
    "brilliant.cul.admin": "BRL",
    "brl.cul.admin": "BRL",

    // Resilient
    "resilient.cul.admin": "RL",
    "res.cul.admin": "RL",
    "rl.cul.admin": "RL",
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

    const canManageTraining = Boolean(isAdmin || isShipCulinaryAdmin);

    return {
      supabase,

      userShip,
      shipDisplayName: getContextShipDisplayName(userShip),

      userEmail: normalizedUserEmail,
      normalizedUserEmail,

      isAdmin,

      culinaryAdminShip,
      culinaryAdminShipDisplayName: getContextShipDisplayName(culinaryAdminShip),

      isShipCulinaryAdmin,

      // General training access.
      // Global admins can manage all training.
      // Ship culinary admins can manage training only for their own ship.
      canManageTraining,

      // Station assignment files are different per ship.
      // Only the matching ship's culinary admin can upload the crew/station file.
      canUploadStationAssignments: Boolean(isAdmin || isShipCulinaryAdmin),

      // Training links are global for all ships.
      // Keep this for global admins only.
      canReplaceTrainingLinks: Boolean(isAdmin),

      // Month/reset affects the selected ship/month training run.
      canManageTrainingMonth: Boolean(isAdmin || isShipCulinaryAdmin),

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
      shipDisplayName: "",

      userEmail: "",
      normalizedUserEmail: "",

      isAdmin: false,

      culinaryAdminShip: "",
      culinaryAdminShipDisplayName: "",

      isShipCulinaryAdmin: false,

      canManageTraining: false,
      canUploadStationAssignments: false,
      canReplaceTrainingLinks: false,
      canManageTrainingMonth: false,

      logUsageEvent: () => {},
    };
  }

  return context;
}

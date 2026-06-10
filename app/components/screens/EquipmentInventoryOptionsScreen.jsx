"use client";

import React from "react";

export default function EquipmentInventoryOptionsScreen(props) {
  const {
    styles,
    activeEquipmentDepartmentLabel,
    equipmentDepartment,
    EquipmentInventoryOptions,
    getShipDisplayName,
    logUsageEvent,
    setEquipmentMode,
    userShip,
  } = props;

  return (
      <EquipmentInventoryOptions
        styles={styles}
        userShip={userShip}
        activeEquipmentDepartmentLabel={activeEquipmentDepartmentLabel}
        getShipDisplayName={getShipDisplayName}
        onBack={() => setEquipmentMode("")}
        onOpenInUse={() => {
          setEquipmentMode("inuse");

          logUsageEvent("equipment_inventory_option_opened", {
            module: "inventory_in_use",
            equipmentDepartment,
            ship: userShip,
          });
        }}
        onOpenWarehouse={() => {
          setEquipmentMode("warehouse");

          logUsageEvent("equipment_inventory_option_opened", {
            module: "inventory_warehouse",
            equipmentDepartment,
            ship: userShip,
          });
        }}
        onOpenMakeInventory={() => {
          setEquipmentMode("makeinventory");

          logUsageEvent("equipment_inventory_option_opened", {
            module: "make_inventory",
            equipmentDepartment,
            ship: userShip,
          });
        }}
      />
    );
}

"use client";

import React from "react";

export default function EquipmentDepartmentOptionsScreen(props) {
  const {
    styles,
    activeEquipmentDepartmentIcon,
    activeEquipmentDepartmentLabel,
    equipmentDepartment,
    EquipmentDepartmentOptions,
    getShipDisplayName,
    logUsageEvent,
    setEquipmentDepartment,
    setEquipmentMode,
    setFmlLowRows,
    setFmlLowSearch,
    setFmlMissingRows,
    setFmlMissingSearch,
    setModule,
    setNextOrderFilter,
    setNextOrderMessage,
    setNextOrderRows,
    setNextOrderSearch,
    setNextOrderView,
    setProductCostReportSearch,
    setProductMode,
    setProductReportView,
    setSearch,
    setSelectedProduct,
    setSelectedRecipe,
    userShip,
  } = props;

  return (
      <EquipmentDepartmentOptions
        styles={styles}
        userShip={userShip}
        equipmentDepartment={equipmentDepartment}
        activeEquipmentDepartmentLabel={activeEquipmentDepartmentLabel}
        activeEquipmentDepartmentIcon={activeEquipmentDepartmentIcon}
        getShipDisplayName={getShipDisplayName}
        onBackToModules={() => {
          setModule("");
          setEquipmentDepartment("");
          setEquipmentMode("");
          setProductMode("");
        }}
        onOpenProductDashboard={() => {
          setModule("product");
          setProductMode("dashboard");

          logUsageEvent("culinary_option_opened", {
            module: "product_dashboard",
            equipmentDepartment,
            ship: userShip,
          });
        }}
        onOpenNextOrder={() => {
          setModule("product");
          setProductMode("nextorder");
          setNextOrderRows([]);
          setFmlMissingRows([]);
          setFmlLowRows([]);
          setNextOrderSearch("");
          setFmlMissingSearch("");
          setFmlLowSearch("");
          setNextOrderFilter("all");
          setNextOrderView("order");
          setNextOrderMessage("");

          logUsageEvent("department_option_opened", {
            module: "generate_next_order",
            equipmentDepartment,
            ship: userShip,
          });
        }}
              onOpenAllergen={() => {
          setSelectedProduct("");
          setSelectedRecipe(null);
          setSearch("");
          setProductCostReportSearch("");
          setProductReportView("main");

          setModule("allergen");

          logUsageEvent("module_opened", {
            module: "allergen",
            equipmentDepartment,
            ship: userShip,
          });
        }}
        onOpenTraining={() => {
          setModule("training");

          logUsageEvent("module_opened", {
            module: "training",
            equipmentDepartment,
            ship: userShip,
          });
        }}
        onOpenMuster={() => {
    setEquipmentMode("muster");

    logUsageEvent("equipment_option_opened", {
      module: `equipment_${equipmentDepartment}_muster`,
      equipmentDepartment,
      ship: userShip,
    });
  }}

  onOpenBreakageReport={() => {
    setModule("breakage");
    setEquipmentMode("");
    setProductMode("");

    logUsageEvent("equipment_option_opened", {
      module: `equipment_${equipmentDepartment}_breakage_report`,
      equipmentDepartment,
      ship: userShip,
    });
  }}

  onOpenInventory={() => {
          setEquipmentMode("inventory");

          logUsageEvent("equipment_option_opened", {
            module: `equipment_${equipmentDepartment}_inventory`,
            equipmentDepartment,
            ship: userShip,
          });
        }}
        onOpenTemperature={() => {
          setModule("temperature");

          logUsageEvent("module_opened", {
            module: "temperature_check",
            equipmentDepartment,
            ship: userShip,
          });
        }}
      />
    );
}

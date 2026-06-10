"use client";

import React from "react";

export default function ProductDashboardScreen(props) {
  const {
    styles,
    activeEquipmentDepartmentLabel,
    allergenWarnings,
    combinedBreakdown,
    consumptionRows,
    dashboardRegionalMatchedCount,
    equipmentDepartment,
    exportMainConsumptionCostReportToExcel,
    exportSelectedProductConsumptionReportToExcel,
    exportTopNotInUseByLocationReportToExcel,
    exportYearlyRegionalConsumptionReportToExcel,
    filteredProductCostReportRows,
    filteredProducts,
    formatMoney,
    formatQty,
    formatRegionalQty,
    getShipDisplayName,
    getSubRecipeIngredients,
    isAdmin,
    loadPermanentIngredientByLocationForProductDashboard,
    message,
    printMainConsumptionCostReport,
    printSelectedProductConsumptionReport,
    printTopNotInUseByLocationReport,
    productCostReportRows,
    productCostReportRowsWithRegionalPar,
    productCostReportSearch,
    productMissingReportLoading,
    productMissingReportMessage,
    productReportView,
    products,
    productsInRecipe,
    recipeRows,
    recipesForProduct,
    refreshProductMissingReport,
    regionalParBufferPercent,
    search,
    selectedProduct,
    selectedRecipe,
    selectedRegionalConsumptionRegion,
    setEquipmentMode,
    setModule,
    setProductCostReportSearch,
    setProductMode,
    setProductReportView,
    setRegionalParBufferPercent,
    setSearch,
    setSelectedProduct,
    setSelectedRecipe,
    setSelectedRegionalConsumptionRegion,
    setViewMode,
    setYearlyRegionalReportMonths,
    setYearlyRegionalReportRegion,
    setYearlyRegionalReportSearch,
    setYearlyRegionalReportShip,
    setYearlyRegionalReportSort,
    showProductMissingReport,
    templateStatus,
    toggleProductMissingReport,
    toggleYearlyRegionalReportMonth,
    topNotInUseReport,
    totalConsumption,
    uploadConsumptionFile,
    uploadRecipeFile,
    uploadTemplateFile,
    uploadYearlyRegionalConsumptionFile,
    userShip,
    viewMode,
    visibleShips,
    YEARLY_REGION_ALL,
    YEARLY_REPORT_ALL_SHIPS,
    yearlyRegionalActiveMonthKeys,
    yearlyRegionalConsumption,
    yearlyRegionalConsumptionReportRows,
    yearlyRegionalFileName,
    yearlyRegionalMessage,
    yearlyRegionalMonthOptions,
    yearlyRegionalReportMonths,
    yearlyRegionalReportRegion,
    yearlyRegionalReportSearch,
    yearlyRegionalReportShip,
    yearlyRegionalReportSort,
  } = props;

  return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button
    style={styles.backButton}
    onClick={() => {
      if (equipmentDepartment) {
        setModule("equipment");
        setEquipmentMode("");
        setProductMode("");
        return;
      }

      setProductMode("");
    }}
  >
    {equipmentDepartment ? `← ${activeEquipmentDepartmentLabel} Options` : "← Product Options"}
  </button>
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
          </div>
        </header>

        <div style={styles.viewModeBox}>
          <button onClick={() => setViewMode("single")} style={{ ...styles.viewModeButton, ...(viewMode === "single" ? styles.viewModeButtonActive : {}) }}>
            🚢 {getShipDisplayName(userShip)} Only
          </button>

          <button onClick={() => setViewMode("all")} style={{ ...styles.viewModeButton, ...(viewMode === "all" ? styles.viewModeButtonActive : {}) }}>
            🌍 All Ships Overview
          </button>
        </div>

        <section style={styles.grid}>
    <div style={styles.card}>
      <h2 style={styles.cardTitle}>📤 Upload Files</h2>

      <label style={styles.label}>Step 1: Consumption file</label>
      <input
        type="file"
        accept=".xlsx,.xls,.xlsm"
        onChange={uploadConsumptionFile}
        style={styles.fileInput}
      />

      <label style={styles.label}>Step 2: Permanent Ingredient by Location file</label>

      <div style={styles.infoBox}>
        <div>
          📄 Ingredient by Location file loads automatically for all users.
        </div>
        <div>
          🔒 Only admins can replace the permanent file.
        </div>
      </div>

      <button
        type="button"
        style={styles.backButton}
        onClick={() => loadPermanentIngredientByLocationForProductDashboard()}
      >
        🔄 Reload Permanent Ingredient by Location
      </button>

      {isAdmin && (
        <>
          <label style={styles.label}>
            Admin only: replace permanent Ingredient by Location file
          </label>

          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadRecipeFile}
            style={styles.fileInput}
          />
        </>
      )}

      <label style={styles.label}>Optional: Replace template file</label>

      <input
        type="file"
        accept=".xlsx,.xls,.xlsm"
        onChange={uploadTemplateFile}
        style={styles.fileInput}
      />

      <label style={styles.label}>
        Optional: Yearly regional consumption file May 2025 - April 2026
      </label>

      <input
        type="file"
        accept=".xlsx,.xls,.xlsm"
        onChange={uploadYearlyRegionalConsumptionFile}
        style={styles.fileInput}
      />

      <label style={styles.label}>Region / home port</label>

      <select
    value={selectedRegionalConsumptionRegion}
    onChange={(e) => setSelectedRegionalConsumptionRegion(e.target.value)}
    style={styles.searchInput}
  >
    <option value="">Select region / origin</option>
    <option value={YEARLY_REGION_ALL}>All regions</option>

    {(yearlyRegionalConsumption?.regionOptions || []).map((region) => (
      <option key={region} value={region}>
        {region}
      </option>
    ))}
  </select>

      <label style={styles.label}>Regional par buffer %</label>

      <input
        type="number"
        min="0"
        step="1"
        value={regionalParBufferPercent}
        onChange={(e) => setRegionalParBufferPercent(Number(e.target.value || 0))}
        style={styles.searchInput}
      />

      {message && <p style={styles.message}>{message}</p>}

      <div style={styles.infoBox}>
        <div>
          📦 Products loaded: <strong>{products.length}</strong>
        </div>

        <div>
          📘 Recipe rows loaded:{" "}
          <strong>{Math.max(recipeRows.length - 1, 0)}</strong>
        </div>

        <div>
          📋 Template: <strong>{templateStatus}</strong>
        </div>

        <div>
    🌎 Yearly regional file:{" "}
    <strong>{yearlyRegionalFileName || "Not uploaded"}</strong>
  </div>

  <div>
    🧭 Selected region:{" "}
    <strong>
      {!selectedRegionalConsumptionRegion
        ? "Not selected"
        : selectedRegionalConsumptionRegion === YEARLY_REGION_ALL
        ? "All regions"
        : selectedRegionalConsumptionRegion}
    </strong>
  </div>

  <div>
    📈 Regional matches:{" "}
    <strong>
      {dashboardRegionalMatchedCount} / {productCostReportRowsWithRegionalPar.length}
    </strong>
  </div>

        <div>
          🧮 Regional buffer:{" "}
          <strong>{formatQty(regionalParBufferPercent)}%</strong>
        </div>

        <div>{yearlyRegionalMessage}</div>

        <div style={{ color: "#b00020" }}>
          Red = recipe/location or template charge location expects usage, but consumption is 0 for visible ship(s).
        </div>

        <div style={{ color: "#0057b8" }}>
          Blue = product is in recipe/location, but missing from the matching venue template.
        </div>
      </div>
    </div>

    <div style={styles.card}>
      <h2 style={styles.cardTitle}>🧭 Product Report View</h2>
      <p style={styles.emptyText}>Choose the report you want to work with.</p>

      <div style={styles.reportModeGrid}>
        <button
          style={{
            ...styles.reportModeButton,
            ...(productReportView === "main" ? styles.reportModeButtonActive : {}),
          }}
          onClick={() => setProductReportView("main")}
        >
          <strong>💰 Main Report</strong>
          <span>Consumption and cost by product, venue and ship</span>
        </button>
          <button
    style={{
      ...styles.reportModeButton,
      ...(productReportView === "yearlyRegional"
        ? styles.reportModeButtonActive
        : {}),
    }}
    onClick={() => setProductReportView("yearlyRegional")}
  >
    <strong>🌎 Yearly Regional Consumption</strong>
    <span>Consumption by month, region / home port, ship, and product</span>
  </button>

        <button
          style={{
            ...styles.reportModeButton,
            ...(productReportView === "consumption"
              ? styles.reportModeButtonActive
              : {}),
          }}
          onClick={() => setProductReportView("consumption")}
        >
          <strong>📊 Consumption Report</strong>
          <span>Consumption vs locations and template</span>
        </button>

        <button
          style={{
            ...styles.reportModeButton,
            ...(productReportView === "reports" ? styles.reportModeButtonActive : {}),
          }}
          onClick={() => setProductReportView("reports")}
        >
          <strong>📋 Generate Report</strong>
          <span>Top 50 items not in use by location</span>
        </button>
      </div>
    </div>

    {productReportView === "consumption" && (
      <div style={styles.card}>
        <h2 style={styles.cardTitle}>🔍 Select Product</h2>

        <input
          placeholder="Search product..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={styles.searchInput}
        />

        <div style={styles.productList}>
          {filteredProducts.map((product, i) => (
            <button
              key={i}
              onClick={() => {
                setSelectedProduct(product);
                setSelectedRecipe(null);
              }}
              style={{
                ...styles.productItem,
                ...(selectedProduct === product ? styles.productItemActive : {}),
              }}
            >
              {product}
            </button>
          ))}
        </div>
      </div>
    )}
  </section>

  {productReportView === "yearlyRegional" && (
    <section style={styles.card}>
      <div
        style={{
          ...styles.header,
          boxShadow: "none",
          padding: 0,
          marginBottom: 18,
        }}
      >
        <div>
          <h2 style={styles.productTitle}>🌎 Yearly Regional Consumption</h2>
          <p style={{ ...styles.emptyText, margin: 0 }}>
            Uses the yearly May 2025 - April 2026 file. Filter by month, region /
            home port, ship, and product.
          </p>
        </div>

        <div style={styles.headerActions}>
          <button
            style={styles.primaryButton}
            onClick={exportYearlyRegionalConsumptionReportToExcel}
          >
            📥 Export Excel
          </button>
        </div>
      </div>

      {!yearlyRegionalConsumption && (
        <p style={styles.emptyText}>
          Upload or load the yearly regional consumption file to use this report.
        </p>
      )}

      <div style={styles.grid}>
        <div style={styles.card}>
          <h3 style={styles.cardTitle}>🔎 Filters</h3>

          <label style={styles.label}>Search product, code, category, region, or ship</label>
          <input
            placeholder="Search yearly regional data..."
            value={yearlyRegionalReportSearch}
            onChange={(event) => setYearlyRegionalReportSearch(event.target.value)}
            style={styles.searchInput}
          />

          <label style={styles.label}>Region / home port</label>
          <select
            value={yearlyRegionalReportRegion}
            onChange={(event) => setYearlyRegionalReportRegion(event.target.value)}
            style={styles.searchInput}
          >
            <option value={YEARLY_REGION_ALL}>All regions</option>

            {(yearlyRegionalConsumption?.regionOptions || []).map((region) => (
              <option key={region} value={region}>
                {region}
              </option>
            ))}
          </select>

          <label style={styles.label}>Ship</label>
          <select
            value={yearlyRegionalReportShip}
            onChange={(event) => setYearlyRegionalReportShip(event.target.value)}
            style={styles.searchInput}
          >
            <option value={YEARLY_REPORT_ALL_SHIPS}>All ships</option>
            <option value="BRL">BRL</option>
            <option value="RL">RL</option>
            <option value="SC">SC</option>
            <option value="VL">VL</option>
          </select>

          <label style={styles.label}>Sort by</label>
          <select
            value={yearlyRegionalReportSort}
            onChange={(event) => setYearlyRegionalReportSort(event.target.value)}
            style={styles.searchInput}
          >
            <option value="qty">Highest quantity</option>
            <option value="value">Highest value</option>
            <option value="daily">Highest daily consumption</option>
            <option value="product">Product name A-Z</option>
          </select>
        </div>

        <div style={styles.card}>
          <h3 style={styles.cardTitle}>📅 Month comparison</h3>

          <div style={styles.viewModeBox}>
            <button
              style={{
                ...styles.viewModeButton,
                ...(yearlyRegionalReportMonths.length === 0
                  ? styles.viewModeButtonActive
                  : {}),
              }}
              onClick={() => setYearlyRegionalReportMonths([])}
            >
              All months
            </button>

            {yearlyRegionalMonthOptions.map((month) => {
              const active =
                yearlyRegionalReportMonths.length === 0 ||
                yearlyRegionalReportMonths.includes(month.monthKey);

              return (
                <button
                  key={month.monthKey}
                  style={{
                    ...styles.viewModeButton,
                    ...(active ? styles.viewModeButtonActive : {}),
                  }}
                  onClick={() => toggleYearlyRegionalReportMonth(month.monthKey)}
                >
                  {month.monthName}
                </button>
              );
            })}
          </div>

          <div style={styles.infoBox}>
            <div>
              📄 File:{" "}
              <strong>{yearlyRegionalFileName || "Not loaded"}</strong>
            </div>

            <div>
              📦 Product rows shown:{" "}
              <strong>{yearlyRegionalConsumptionReportRows.length}</strong>
            </div>

            <div>
              🧭 Region filter:{" "}
              <strong>
                {yearlyRegionalReportRegion === YEARLY_REGION_ALL
                  ? "All regions"
                  : yearlyRegionalReportRegion}
              </strong>
            </div>

            <div>
              🚢 Ship filter:{" "}
              <strong>
                {yearlyRegionalReportShip === YEARLY_REPORT_ALL_SHIPS
                  ? "All ships"
                  : yearlyRegionalReportShip}
              </strong>
            </div>

            <div>
              📅 Months selected:{" "}
              <strong>
                {yearlyRegionalReportMonths.length === 0
                  ? "All months"
                  : yearlyRegionalMonthOptions
                      .filter((month) =>
                        yearlyRegionalReportMonths.includes(month.monthKey)
                      )
                      .map((month) => month.monthName)
                      .join(", ")}
              </strong>
            </div>
          </div>
        </div>
      </div>

      {yearlyRegionalConsumption &&
        yearlyRegionalConsumptionReportRows.length === 0 && (
          <p style={styles.emptyText}>
            No yearly regional rows match the current filters.
          </p>
        )}

      <div style={styles.costReportLineList}>
        {yearlyRegionalConsumptionReportRows.map((item, index) => {
          const activeMonthSet = new Set(yearlyRegionalActiveMonthKeys);
          const visibleMonths = yearlyRegionalMonthOptions.filter((month) =>
            activeMonthSet.has(month.monthKey)
          );

          return (
            <div
              key={`${item.productCode || item.productKey || item.productName}-${index}`}
              style={styles.costReportLine}
            >
              <div style={styles.costLineMain}>
                <div style={styles.costLineProduct}>
                  {item.productName || "Unnamed product"}
                </div>

                <div style={styles.costLineMeta}>
                  {item.productCode ? "Code: " + item.productCode + " • " : ""}
                  U/M: {item.unitMeasure || "N/A"}
                </div>

                <div style={styles.costLineMeta}>
                  {item.categoryName || "No category"}
                  {item.subCategoryName ? " • " + item.subCategoryName : ""}
                </div>

                <div style={styles.statusGood}>
                  Daily: {formatRegionalQty(item.avgDailyQty)} • Days:{" "}
                  {formatQty(item.totalDays)} • Blocks: {item.blocks}
                </div>
              </div>

              <div style={styles.costLineTotals}>
                <span>Total Qty</span>
                <strong>{formatRegionalQty(item.totalQty)}</strong>
                <span>{formatMoney(item.totalValue)}</span>
                {item.avgPrice > 0 && (
                  <small>{formatMoney(item.avgPrice)} / unit</small>
                )}
              </div>

              <div style={styles.costLineVenues}>
                {visibleMonths.map((month) => {
                  const monthData = item.months?.[month.monthKey] || {};
                  const hasMonthData =
                    Number(monthData.qty || 0) !== 0 ||
                    Number(monthData.value || 0) !== 0;

                  return (
                    <div key={month.monthKey} style={styles.costLineVenue}>
                      <div style={styles.costLineVenueTitle}>{month.monthName}</div>

                      {hasMonthData ? (
                        <>
                          <div style={styles.costLineShipChips}>
                            <div style={styles.costLineShipChip}>
                              <span style={styles.costLineShipName}>Qty</span>
                              <strong>{formatRegionalQty(monthData.qty)}</strong>
                            </div>

                            <div style={styles.costLineShipChip}>
                              <span style={styles.costLineShipName}>Value</span>
                              <strong>{formatMoney(monthData.value)}</strong>
                            </div>

                            <div style={styles.costLineShipChip}>
                              <span style={styles.costLineShipName}>Daily</span>
                              <strong>
                                {formatRegionalQty(monthData.avgDailyQty)}
                              </strong>
                            </div>

                            <div style={styles.costLineShipChip}>
                              <span style={styles.costLineShipName}>Days</span>
                              <strong>{formatQty(monthData.days)}</strong>
                            </div>
                          </div>

                          <div style={styles.costLinePriceNote}>
                            Ships: {(monthData.ships || []).join(", ") || "N/A"}
                          </div>

                          <div style={styles.recipeMeta}>
                            Regions: {(monthData.regions || []).join(", ") || "N/A"}
                          </div>
                        </>
                      ) : (
                        <div style={styles.statusNeutral}>No consumption</div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    </section>
  )}
        {productReportView === "main" && (
        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 18 }}>
            <div>
              <h2 style={styles.productTitle}>💰 Main Consumption & Cost Report</h2>
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Automatically generated from the consumption file. Lowest unit price is highlighted when ship prices differ.
              </p>
            </div>

            <div style={styles.headerActions}>
              <button style={styles.backButton} onClick={printMainConsumptionCostReport}>
                🖨️ Print
              </button>
              <button style={styles.primaryButton} onClick={exportMainConsumptionCostReportToExcel}>
                📥 Export Excel
              </button>
            </div>
          </div>

          <div style={styles.infoBox}>
            <div>📦 Products in report: <strong>{filteredProductCostReportRows.length}</strong> / {productCostReportRows.length}</div>
            <div>🚢 View: <strong>{viewMode === "single" ? userShip : "All Ships"}</strong></div>
            <div>📘 Source columns: C Venue, I/L/O/R Quantity, J/M/P/S Total Cost, H/K/N/Q Unit Price.</div>
          </div>

          <input
            placeholder="Search product, code, or venue..."
            value={productCostReportSearch}
            onChange={(e) => setProductCostReportSearch(e.target.value)}
            style={{ ...styles.searchInput, marginTop: 14 }}
          />

          {consumptionRows.length === 0 && (
            <p style={styles.emptyText}>Upload the consumption file to generate this report.</p>
          )}

          {consumptionRows.length > 0 && filteredProductCostReportRows.length === 0 && (
            <p style={styles.emptyText}>No products found for this search.</p>
          )}

          <div style={styles.costReportLineList}>
            {filteredProductCostReportRows.map((item) => {
    const itemVenues = Array.isArray(item.venues) ? item.venues : [];

    return (
              <div key={item.productKey} style={styles.costReportLine}>
                <div style={styles.costLineMain}>
    <div style={styles.costLineProduct}>{item.product}</div>

    <div style={styles.costLineMeta}>
      {item.code ? "Code: " + item.code + " • " : ""}
      {itemVenues.length} venue{itemVenues.length === 1 ? "" : "s"}
    </div>

    {item.regionalHasData ? (
      <div style={styles.statusGood}>
        Region:{" "}
        {item.regionalRegion === YEARLY_REGION_ALL
          ? "All regions"
          : item.regionalRegion}{" "}
        • Daily: {formatRegionalQty(item.regionalAvgDailyQty)} • Suggested Par:{" "}
        {formatRegionalQty(item.regionalSuggestedParLevel)}
      </div>
    ) : yearlyRegionalConsumption ? (
      <div style={styles.statusNeutral}>
        No regional yearly match
      </div>
    ) : null}
  </div>

                <div style={styles.costLineTotals}>
                  <span>Total</span>
                  <strong>{formatQty(item.visibleTotalQty)}</strong>
                  <span>{formatMoney(item.visibleTotalCost)}</span>
                </div>

                <div style={styles.costLineVenues}>
                  {itemVenues.map((venue) => (
                    <div key={venue.venueKey} style={styles.costLineVenue}>
                      <div style={styles.costLineVenueTitle}>{venue.location}</div>

                      <div style={styles.costLineShipChips}>
                        {visibleShips.map((ship) => {
                          const shipData = venue.ships[ship] || { qty: 0, cost: 0, unitPrice: 0, isLowestUnitPrice: false };
                          const hasData = Number(shipData.qty || 0) !== 0 || Number(shipData.cost || 0) !== 0;

                          if (!hasData) return null;

                          return (
                            <div
                              key={ship}
                              style={{
                                ...styles.costLineShipChip,
                                ...(shipData.isLowestUnitPrice ? styles.costLineShipLowest : {}),
                              }}
                            >
                              <span style={styles.costLineShipName}>{ship}</span>
                              <strong>{formatQty(shipData.qty)}</strong>
                              <span>{formatMoney(shipData.cost)}</span>
                              {shipData.unitPrice > 0 && (
                                <small>{formatMoney(shipData.unitPrice)} / unit</small>
                              )}
                            </div>
                          );
                        })}
                      </div>

                      {venue.hasPriceDifference && (
                        <div style={styles.costLinePriceNote}>Lowest unit price highlighted.</div>
                      )}
                    </div>
                  ))}
                </div>
                          </div>
            );
          })}
          </div>

        </section>
        )}

        {productReportView === "reports" && (
        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: showProductMissingReport ? 18 : 0 }}>
            <div>
              <h2 style={styles.productTitle}>📋 Generate Report</h2>
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Generate the Top 50 items not in use by the location where they should be used.
              </p>
            </div>

            <div style={styles.headerActions}>
              <button
                style={styles.backButton}
                onClick={toggleProductMissingReport}
                disabled={productMissingReportLoading}
              >
                {productMissingReportLoading ? "Preparing..." : showProductMissingReport ? "Hide Report" : "Open Report"}
              </button>

              {showProductMissingReport && (
                <button
                  style={styles.backButton}
                  onClick={refreshProductMissingReport}
                  disabled={productMissingReportLoading}
                >
                  🔄 Refresh Report
                </button>
              )}

              <button
                style={styles.backButton}
                onClick={printTopNotInUseByLocationReport}
                disabled={productMissingReportLoading}
              >
                🖨️ Print
              </button>

              <button
                style={styles.primaryButton}
                onClick={exportTopNotInUseByLocationReportToExcel}
                disabled={productMissingReportLoading}
              >
                📥 Export Excel
              </button>
            </div>
          </div>

          {showProductMissingReport && (
            <>
              <div style={styles.infoBox}>
                <div>📋 Report: <strong>Top 50 Not In Use By Location</strong></div>
                <div>🚢 View: <strong>{viewMode === "single" ? userShip : "All Ships"}</strong></div>
                <div>🔎 Rows found: <strong>{topNotInUseReport.length}</strong></div>
                {productMissingReportLoading && <div>Preparing report, please wait...</div>}
                {productMissingReportMessage && <div style={{ color: topNotInUseReport.length ? "#555" : "#8a5a00" }}>{productMissingReportMessage}</div>}
                <div style={{ color: "#b00020" }}>
                  Shows products that are expected by recipe/location or template charge location, but usage is 0 for one or more visible ship(s).
                </div>
              </div>

              {!productMissingReportLoading && topNotInUseReport.length === 0 && (
                <p style={styles.emptyText}>
                  No not-in-use records found. Upload consumption, recipe/location, and template files, then open the report again.
                </p>
              )}

              <div style={styles.equipmentGrid}>
                {topNotInUseReport.map((item, index) => (
                  <div key={`${item.product}-${item.venueKey}-${index}`} style={{ ...styles.equipmentCard, ...styles.orderWarningCard }}>
                    <div style={styles.recipeMeta}>#{index + 1}</div>
                    <div style={styles.recipeName}>{item.product}</div>
                    <div style={styles.recipeMeta}>Location: {item.location}</div>
                    <div style={styles.recipeMeta}>Source: {item.source}</div>

                    {item.templateMatches.length > 0 && (
                      <div style={styles.templateFound}>Template/Menu: {item.templateMatches.join(", ")}</div>
                    )}

                    {item.missingFromTemplate && (
                      <div style={styles.templateWarningText}>Also missing from matching template.</div>
                    )}

                    <div style={styles.shipGrid}>
                      {visibleShips.map((ship) => {
                        const isMissing = item.missingShips.includes(ship);

                        return (
                          <div
                            key={ship}
                            style={{ ...styles.shipBox, ...(ship === userShip ? styles.shipBoxActive : {}), ...(isMissing ? styles.shipBoxMissing : {}) }}
                          >
                            <span style={styles.shipName}>{ship}</span>
                            <strong style={styles.shipQty}>{formatQty(item.ships?.[ship])}</strong>
                          </div>
                        );
                      })}
                    </div>

                    <div style={styles.statusBad}>Missing: {item.missingShips.join(", ")}</div>

                    <button
                      style={styles.backButton}
                      onClick={() => {
                        setProductReportView("consumption");
                        setSelectedProduct(item.product);
                        setSelectedRecipe(null);
                      }}
                    >
                      Open Product
                    </button>
                  </div>
                ))}
              </div>
            </>
          )}
        </section>
        )}

        {productReportView === "consumption" && selectedProduct && (
          <section style={styles.card}>
            <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 18 }}>
              <div>
                <h2 style={styles.productTitle}>📦 {selectedProduct}</h2>
                <p style={{ ...styles.emptyText, margin: 0 }}>
                  Consumption vs locations and template for the selected product.
                </p>
              </div>

              <div style={styles.headerActions}>
                <button style={styles.backButton} onClick={printSelectedProductConsumptionReport}>
                  🖨️ Print
                </button>
                <button style={styles.primaryButton} onClick={exportSelectedProductConsumptionReportToExcel}>
                  📥 Export Excel
                </button>
              </div>
            </div>

            <h3 style={styles.sectionTitle}>📊 Total Consumption</h3>

            <div style={styles.totalBox}>
              <div style={styles.totalMain}>
                {viewMode === "single" ? `${userShip} Total: ` : "Total All Ships: "}
                {formatQty(totalConsumption.allShips)}
              </div>

              <div style={styles.totalShipGrid}>
                {visibleShips.map((ship) => (
                  <div key={ship} style={styles.totalShipBox}>
                    <span>{ship}</span>
                    <strong>{formatQty(totalConsumption.totals[ship])}</strong>
                  </div>
                ))}
              </div>
            </div>

            <h3 style={styles.sectionTitle}>🏢 Consumption by Venue and Ship</h3>

            <div style={styles.venueGrid}>
              {combinedBreakdown.map((venueItem, i) => (
                <div key={i} style={{ ...styles.venueCard, ...(venueItem.missingShips.length > 0 ? styles.venueCardWarning : {}), ...(venueItem.missingFromTemplate ? styles.venueCardTemplateWarning : {}) }}>
                  <h4 style={styles.venueTitle}>
                    {venueItem.displayName}
                    <span style={styles.badgeGroup}>
                      {venueItem.requiredFromTemplate && <span style={styles.chargeBadge}>Template Charge</span>}
                      {venueItem.missingFromTemplate && <span style={styles.templateBadge}>Missing Template</span>}
                      {venueItem.missingShips.length > 0 && <span style={styles.missingBadge}>Missing: {venueItem.missingShips.join(", ")}</span>}
                    </span>
                  </h4>

                  {venueItem.requiredFromTemplate && (
                    <div style={styles.templateFound}>
                      Template charge location: {venueItem.templateMatches.join(", ") || "Template"}
                    </div>
                  )}

                  {venueItem.missingFromTemplate && (
                    <div style={styles.templateWarningText}>
                      Product is used in the recipe/location file for this venue, but it is not found in this venue template.
                    </div>
                  )}

                  <div style={styles.shipGrid}>
                    {visibleShips.map((ship) => {
                      const isMissing = venueItem.required && (venueItem.ships[ship] || 0) === 0;

                      return (
                        <div key={ship} style={{ ...styles.shipBox, ...(ship === userShip ? styles.shipBoxActive : {}), ...(isMissing ? styles.shipBoxMissing : {}) }}>
                          <span style={styles.shipName}>{ship}</span>
                          <strong style={styles.shipQty}>{formatQty(venueItem.ships[ship])}</strong>
                        </div>
                      );
                    })}
                  </div>

                  {venueItem.missingShips.length > 0 && (
                    <div style={styles.warningSmall}>
                      {venueItem.requiredFromTemplate
                        ? "Product appears in the template charge location, but usage is 0 for highlighted ship(s)."
                        : "Product appears in recipe/location file for this venue, but usage is 0 for highlighted ship(s)."}
                    </div>
                  )}
                </div>
              ))}
            </div>

            <h3 style={styles.sectionTitle}>👨‍🍳 Recipes using this product</h3>

            {recipeRows.length === 0 && <p style={styles.emptyText}>Upload the recipe/location file to see recipe details.</p>}
            {recipeRows.length > 0 && recipesForProduct.length === 0 && <p style={styles.emptyText}>No recipes found for this product.</p>}

            <div style={styles.recipeList}>
              {recipesForProduct.map((recipe, i) => (
                <button key={i} onClick={() => setSelectedRecipe(recipe)} style={{ ...styles.recipeCard, ...(selectedRecipe?.key === recipe.key ? styles.recipeCardActive : {}) }}>
                  <div style={styles.recipeName}>{recipe.recipeName}</div>
                  <div style={styles.recipeMeta}>Code: {recipe.recipeCode}</div>
                  <div style={styles.recipeMeta}>Venues: {recipe.venues.length ? recipe.venues.join(", ") : "N/A"}</div>
                </button>
              ))}
            </div>

            {selectedRecipe && (
              <div style={styles.ingredientsCard}>
                <h3 style={styles.sectionTitle}>🧾 Products used in recipe</h3>
                <h4 style={{ marginTop: 0 }}>{selectedRecipe.recipeName} ({selectedRecipe.recipeCode})</h4>

                {productsInRecipe.length === 0 ? (
                  <p style={styles.emptyText}>No products found for this recipe.</p>
                ) : (
                  <ul>
                    {productsInRecipe.map((product, i) => {
                      const subIngredients = getSubRecipeIngredients(product);

                      return (
                        <li key={i} style={{ marginBottom: 10 }}>
                          <strong>{product}</strong>
                          {subIngredients.length > 0 && (
                            <ul style={styles.subRecipeList}>
                              {subIngredients.map((subItem, j) => <li key={j}>{subItem}</li>)}
                            </ul>
                          )}
                        </li>
                      );
                    })}
                  </ul>
                )}

                <h3 style={styles.sectionTitle}>⚠️ Rule-Based Allergen Warning</h3>
                <p style={styles.warningText}>This is a keyword-based warning only. Verify against official allergen data before use.</p>

                {allergenWarnings.length === 0 ? (
                  <p style={styles.emptyText}>No likely allergens detected by keyword rules.</p>
                ) : (
                  <div style={styles.allergenList}>
                    {allergenWarnings.map((item, i) => (
                      <div key={i} style={styles.allergenCard}>
                        <strong>{item.allergen}</strong>
                        <ul>
                          {item.products.map((product, j) => <li key={j}>{product}</li>)}
                        </ul>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </section>
        )}
      </main>
    );
}

"use client";

import React from "react";

export default function MakeInventoryScreen(props) {
  const {
    styles,
    activeEquipmentDepartmentLabel,
    cleanText,
    clearMyInventory,
    clearShipInventory,
    confirmInventoryQty,
    currentInventoryItem,
    deleteInventoryItem,
    drivePictureLibraryByCode,
    editingInventoryId,
    editInventoryItem,
    equipmentDepartment,
    exportInventoryStatusToExcel,
    exportInventorySummaryToExcel,
    extraInventoryCode,
    extraInventoryMessage,
    extraInventoryName,
    extraInventoryPhotoFile,
    extraInventoryPhotoInputKey,
    extraInventoryQty,
    extraInventorySaving,
    formatQty,
    generateFinalInventoryReport,
    getActiveInventoryStationList,
    getAllInventoryStationsSubmitted,
    getCurrentStationProgress,
    getEffectiveInventoryUserName,
    getEquipmentDisplayImage,
    getEquipmentFallbackImage,
    getFilteredMakeInventoryItems,
    getImageUrl,
    getInventoryItemKey,
    getInventoryStationLabel,
    getInventoryStationProgressRows,
    getMyInventoryRows,
    getMyInventoryStatusRows,
    getShipDisplayName,
    getShipSummaryRows,
    getSummaryStationOptions,
    getVisibleInventoryReportRows,
    handleInventoryCountSheetTemplateFile,
    inventoryCountSheetTemplateName,
    inventoryError,
    inventoryLoading,
    inventoryQty,
    inventoryReportMode,
    inventoryStation,
    inventoryUserName,
    inventoryUserPosition,
    isAdmin,
    loadMasterInventoryItems,
    makeInventoryItems,
    makeInventoryMessage,
    makeInventorySearch,
    makeInventoryShip,
    MakeInventoryTopBar,
    masterInventoryLoading,
    masterInventorySource,
    pictureLibraryBusy,
    printInventoryStatus,
    printInventorySummary,
    refreshMakeInventoryData,
    reportBusy,
    resetInventoryRun,
    saveExtraInventoryItem,
    setCurrentInventoryItem,
    setEditingInventoryId,
    setEquipmentMode,
    setExtraInventoryCode,
    setExtraInventoryName,
    setExtraInventoryPhotoFile,
    setExtraInventoryQty,
    setInventoryQty,
    setInventoryReportMode,
    setInventoryStation,
    setInventoryUserName,
    setInventoryUserPosition,
    setMakeInventoryMessage,
    setMakeInventorySearch,
    setMakeInventoryShip,
    setShowVariance,
    setSummaryStationFilter,
    SHIPS,
    showVariance,
    submitInventoryStationCount,
    summaryStationFilter,
    supabase,
    syncMasterInventoryPicturesFromDrive,
    userShip,
  } = props;

  const filteredMakeInventoryItems = getFilteredMakeInventoryItems();
      const myReportRows = getMyInventoryRows();
      const summaryReportRows = getShipSummaryRows();
          const stationProgressRows = getInventoryStationProgressRows();
      const currentStationProgress = getCurrentStationProgress();
      const allStationsSubmitted = getAllInventoryStationsSubmitted();
      const startedStations = stationProgressRows.filter((item) => item.status === "started").length;
      const submittedStations = stationProgressRows.filter((item) => item.status === "submitted").length;
      const currentStationSubmitted = currentStationProgress?.status === "submitted";
      const visibleReportRows = getVisibleInventoryReportRows();
      const summaryStationOptions = getSummaryStationOptions();
      const activeInventoryStations = getActiveInventoryStationList();
      const inventoryStationLabel = getInventoryStationLabel();
      const selectedSummaryStationLabel =
        summaryStationFilter === "ALL"
          ? equipmentDepartment === "bar"
            ? "All Bars"
            : "All Stations"
          : summaryStationFilter;
      const inventoryStatusRows = getMyInventoryStatusRows();
      const statusCountedItems = inventoryStatusRows.filter((item) => item.status === "Counted");
      const statusPendingItems = inventoryStatusRows.filter((item) => item.status !== "Counted");
      const userName = getEffectiveInventoryUserName();
      const inventoryReady = Boolean(makeInventoryShip && inventoryStation && userName && supabase);
      const countedKeysForMe = new Set(myReportRows.map((item) => item.itemKey || cleanText(item.code || item.name)));
          const countedRecordByKey = new Map(
        myReportRows.map((item) => [
          item.itemKey || cleanText(item.code || item.name),
          item,
        ])
      );
      const currentInventoryDisplayImage = currentInventoryItem
    ? getEquipmentDisplayImage(currentInventoryItem)
    : "";

  const currentInventoryFallbackImage = currentInventoryItem
    ? getEquipmentFallbackImage(currentInventoryItem)
    : "";

  const currentInventoryOpenImage =
    currentInventoryDisplayImage || currentInventoryFallbackImage || "";

          const selectInventoryItemForCounting = (item) => {
        const itemKey = getInventoryItemKey(item);
        const countedRecord = countedRecordByKey.get(itemKey);

        if (!inventoryReady) {
          setMakeInventoryMessage("Choose ship, station and user before counting.");
          return;
        }

        if (currentStationSubmitted) {
          setMakeInventoryMessage(
            `${inventoryStation} has already submitted count. Waiting for all stations before final report.`
          );
          return;
        }

        const displayImage = getEquipmentDisplayImage(item);
  const fallbackImage = getEquipmentFallbackImage(item);

  setCurrentInventoryItem({
    ...item,
    image: displayImage,
    imageFallback: fallbackImage,
    itemKey,
  });

        setInventoryQty(countedRecord ? String(countedRecord.qty ?? "") : "");
        setEditingInventoryId(countedRecord?.id || null);
        setMakeInventoryMessage(`Selected: ${item.name}`);
      };
      const sortedMakeInventoryItems = filteredMakeInventoryItems
        .map((item, index) => ({
          item,
          index,
          itemKey: getInventoryItemKey(item),
        }))
        .sort((a, b) => {
          const aCounted = countedKeysForMe.has(a.itemKey);
          const bCounted = countedKeysForMe.has(b.itemKey);

          if (aCounted !== bCounted) {
            return aCounted ? 1 : -1;
          }

          return a.index - b.index;
        })
        .map((entry) => entry.item);

      return (
        <main style={styles.page}>
          <header style={styles.header}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
            <div style={styles.headerActions}>
              <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
              <div style={styles.shipBadge}>🚢 {getShipDisplayName(makeInventoryShip || userShip)}</div>
            </div>
          </header>
        <MakeInventoryTopBar
    styles={styles}
      isAdmin={isAdmin}
    inventoryReportMode={inventoryReportMode}
    setInventoryReportMode={setInventoryReportMode}
    makeInventoryShip={makeInventoryShip}
    userShip={userShip}
    refreshMakeInventoryData={refreshMakeInventoryData}
    inventoryLoading={inventoryLoading}
    reportBusy={reportBusy}
    resetInventoryRun={resetInventoryRun}
    printInventorySummary={printInventorySummary}
    clearMyInventory={clearMyInventory}
    clearShipInventory={clearShipInventory}
    generateFinalInventoryReport={generateFinalInventoryReport}
    exportInventorySummaryToExcel={exportInventorySummaryToExcel}
    allStationsSubmitted={allStationsSubmitted}
    inventoryStation={inventoryStation}
    currentStationProgress={currentStationProgress}
    myReportRows={myReportRows}
    makeInventoryItems={makeInventoryItems}
    startedStations={startedStations}
    submittedStations={submittedStations}
    stationProgressRows={stationProgressRows}
  />

          <section style={styles.grid}>
            <div style={styles.card}>
              <h2 style={styles.cardTitle}>📝 {activeEquipmentDepartmentLabel} Make Inventory</h2>

              <label style={styles.label}>Choose ship for this inventory</label>
              <select
                value={makeInventoryShip}
                onChange={(e) => setMakeInventoryShip(e.target.value)}
                style={styles.select}
              >
                <option value="">Choose ship</option>
                {SHIPS.map((ship) => (
    <option key={ship} value={ship}>
      {getShipDisplayName(ship)}
    </option>
  ))}
              </select>

              <label style={styles.label}>Choose {inventoryStationLabel}</label>
              <select
                value={inventoryStation}
                onChange={(e) => setInventoryStation(e.target.value)}
                style={styles.select}
              >
                <option value="">Choose {inventoryStationLabel}</option>
                {activeInventoryStations.map((station) => (
                  <option key={station} value={station}>{station}</option>
                ))}
              </select>

              <label style={styles.label}>Choose user</label>
              <input
                placeholder="Type your name..."
                value={inventoryUserName}
                onChange={(e) => setInventoryUserName(e.target.value)}
                style={styles.searchInput}
              />

              <label style={styles.label}>Position</label>
              <input
                placeholder="Type your position..."
                value={inventoryUserPosition}
                onChange={(e) => setInventoryUserPosition(e.target.value)}
                style={styles.searchInput}
              />

  <button
    type="button"
    style={styles.backButton}
    onClick={() => loadMasterInventoryItems(makeInventoryShip || userShip)}
    disabled={masterInventoryLoading}
  >
    {masterInventoryLoading ? "Refreshing..." : "🔄 Refresh Shared Master List"}
  </button>
  {isAdmin && equipmentDepartment === "culinary" && (
    <button
      type="button"
      style={styles.backButton}
      onClick={syncMasterInventoryPicturesFromDrive}
      disabled={pictureLibraryBusy || masterInventoryLoading}
    >
      {pictureLibraryBusy
        ? "Syncing pictures..."
        : "🖼️ Sync Pictures With Master List"}
    </button>
  )}

              {makeInventoryMessage && <p style={styles.message}>{makeInventoryMessage}</p>}

              <div style={styles.infoBox}>
                <div>🚢 Inventory ship: <strong>{makeInventoryShip || "Not selected"}</strong></div>
                <div>📍 {equipmentDepartment === "bar" ? "Bar" : "Station"}: <strong>{inventoryStation || "Not selected"}</strong></div>
                <div>👤 User: <strong>{userName || "Not selected"}</strong></div>
                <div>📋 Shared master items: <strong>{makeInventoryItems.length}</strong></div>
                <div>🖼️ Drive pictures loaded: <strong>{Object.keys(drivePictureLibraryByCode).length}</strong></div>
                <div>📂 Master source: <strong>{masterInventorySource || "Not loaded"}</strong></div>
                {masterInventoryLoading && <div>Loading shared master inventory...</div>}
                <div>✅ My counted items: <strong>{myReportRows.length}</strong></div>
                <div>🌍 Ship summary records: <strong>{summaryReportRows.length}</strong></div>
                <div>Duplicate entries are automatically updated instead of added twice.</div>
                {inventoryLoading && <div>Loading shared inventory records...</div>}
                {inventoryError && <div style={{ color: "#b00020" }}>{inventoryError}</div>}
              </div>
            </div>

            <div style={styles.card}>
              <h2 style={styles.cardTitle}>🔍 Search Product</h2>

              <input
                placeholder="Search code, name, category or sheet..."
                value={makeInventorySearch}
                onChange={(e) => setMakeInventorySearch(e.target.value)}
                style={styles.searchInput}
              />

              {!inventoryReady && (
                <p style={styles.warningText}>
                  Choose ship, station and user before counting. Supabase must be connected.
                </p>
              )}

              <p style={styles.emptyText}>
                Select the correct product from the master list. Green means already counted by this user for this station.
              </p>
            </div>
                <div style={styles.card}>
    <h2 style={styles.cardTitle}>➕ Item Not In Master List</h2>

    <p style={styles.emptyText}>
      Use this only when the item is found during inventory but does not exist in the shared master list.
    </p>

    <label style={styles.label}>Code / SKU optional</label>
    <input
      placeholder="Enter code or SKU..."
      value={extraInventoryCode}
      onChange={(event) => setExtraInventoryCode(event.target.value)}
      style={styles.searchInput}
    />

    <label style={styles.label}>Item name</label>
    <input
      placeholder="Enter item name..."
      value={extraInventoryName}
      onChange={(event) => setExtraInventoryName(event.target.value)}
      style={styles.searchInput}
    />

    <label style={styles.label}>Quantity</label>
    <input
      type="number"
      min="0"
      placeholder="Enter quantity..."
      value={extraInventoryQty}
      onChange={(event) => setExtraInventoryQty(event.target.value)}
      style={styles.searchInput}
    />

    <label style={styles.label}>Take / upload photo optional</label>
    <input
      key={extraInventoryPhotoInputKey}
      type="file"
      accept="image/*"
      capture="environment"
      onChange={(event) =>
        setExtraInventoryPhotoFile(event.target.files?.[0] || null)
      }
      style={styles.fileInput}
    />

    {extraInventoryPhotoFile && (
      <div style={styles.statusNeutral}>
        Photo selected: {extraInventoryPhotoFile.name}
      </div>
    )}

    <button
      type="button"
      style={styles.primaryButton}
      onClick={saveExtraInventoryItem}
      disabled={extraInventorySaving || !inventoryReady || currentStationSubmitted}
    >
      {extraInventorySaving ? "Saving..." : "Save Extra Item"}
    </button>

    {extraInventoryMessage && (
      <p style={styles.message}>{extraInventoryMessage}</p>
    )}

    {!inventoryReady && (
      <p style={styles.warningText}>
        Choose ship, station and user before saving an extra item.
      </p>
    )}

    {currentStationSubmitted && (
      <p style={styles.warningText}>
        This station already submitted count. Extra items cannot be added now.
      </p>
    )}
  </div>
          </section>

                  <section style={styles.card}>
            <h2 style={styles.productTitle}>📦 Select Product for Inventory</h2>

            {makeInventoryItems.length === 0 && (
              <p style={styles.emptyText}>
                Upload the shared master inventory file once for this ship, or click Refresh if another user already uploaded it.
              </p>
            )}

            <div style={styles.equipmentGrid}>
              {sortedMakeInventoryItems.map((item, index) => {
                const itemKey = getInventoryItemKey(item);
                const alreadyCounted = countedKeysForMe.has(itemKey);
                const countedRecord = countedRecordByKey.get(itemKey);
                const displayImage = getEquipmentDisplayImage(item);
                const fallbackImage = getEquipmentFallbackImage(item);

                return (
                  <button
                    key={`${item.sheetName}-${item.code}-${index}`}
                    style={{
    ...styles.inventoryItemCard,
    ...(alreadyCounted ? styles.countedCard : {}),
  }}
                    onClick={() =>
    selectInventoryItemForCounting({
      ...item,
      image: displayImage,
      imageFallback: fallbackImage,
    })
  }
                  >
                    {displayImage ? (
    <div style={styles.inventoryImageFrame}>
      <img
    src={getImageUrl(displayImage, "w360")}
    alt={item.name}
    loading="lazy"
    decoding="async"
    style={styles.inventoryCardImage}
    data-fallback-src={
      fallbackImage && fallbackImage !== displayImage
        ? getImageUrl(fallbackImage, "w360")
        : ""
    }
    onError={(e) => {
    const fallbackSrc = e.currentTarget.dataset.fallbackSrc;

    if (fallbackSrc && e.currentTarget.dataset.usedFallback !== "true") {
      e.currentTarget.dataset.usedFallback = "true";
      e.currentTarget.src = fallbackSrc;
      return;
    }

    e.currentTarget.style.display = "none";
    const fallback = e.currentTarget.nextElementSibling;
    if (fallback) fallback.style.display = "flex";
  }}
  />

      <div style={{ ...styles.inventoryNoImage, display: "none" }}>
        No image
      </div>
    </div>
  ) : (
    <div style={styles.inventoryNoImage}>No image</div>
  )}
                    <div style={styles.recipeName}>{item.name}</div>
                    <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                    <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                    <div style={styles.recipeMeta}>Category: {item.category}</div>

                    {alreadyCounted && (
                      <div style={styles.statusGood}>
                        Already Counted: {formatQty(countedRecord?.qty || 0)}
                      </div>
                    )}
                  </button>
                );
              })}
            </div>
                  </section>

          {currentInventoryItem ? (
            <div
              style={styles.modalBackdrop}
              onClick={() => {
                setCurrentInventoryItem(null);
                setInventoryQty("");
                setEditingInventoryId(null);
              }}
            >
              <div style={styles.modalCard} onClick={(event) => event.stopPropagation()}>
                <button
                  type="button"
                  style={styles.closeButton}
                  onClick={() => {
                    setCurrentInventoryItem(null);
                    setInventoryQty("");
                    setEditingInventoryId(null);
                  }}
                >
                  ✕
                </button>

                <h2 style={styles.productTitle}>
                  {editingInventoryId ? "✏️ Update Quantity" : "✅ Insert Quantity"}
                </h2>

                <div style={styles.grid}>
                  <div>
    {currentInventoryDisplayImage ? (
      <div>
        <div style={styles.modalPictureFrame}>
          <img
            src={getImageUrl(currentInventoryDisplayImage, "w720")}
            alt={currentInventoryItem.name}
            style={styles.modalPreviewImage}
            data-fallback-src={
              currentInventoryFallbackImage &&
              currentInventoryFallbackImage !== currentInventoryDisplayImage
                ? getImageUrl(currentInventoryFallbackImage, "w720")
                : ""
            }
            onError={(e) => {
              const fallbackSrc = e.currentTarget.dataset.fallbackSrc;

              if (
                fallbackSrc &&
                e.currentTarget.dataset.usedFallback !== "true"
              ) {
                e.currentTarget.dataset.usedFallback = "true";
                e.currentTarget.src = fallbackSrc;
                return;
              }

              e.currentTarget.style.display = "none";
              const fallback = e.currentTarget.nextElementSibling;
              if (fallback) fallback.style.display = "flex";
            }}
          />

          <div style={{ ...styles.modalNoImage, display: "none" }}>
            Picture could not be loaded
          </div>
        </div>

        <a
          href={currentInventoryOpenImage}
          target="_blank"
          rel="noreferrer"
          style={{
            ...styles.imageButton,
            display: "block",
            textDecoration: "none",
          }}
        >
          Open Picture
        </a>
      </div>
    ) : (
      <div style={styles.modalNoImage}>No image</div>
    )}
  </div>

                  <div>
                    <h3 style={{ marginTop: 0 }}>{currentInventoryItem.name}</h3>

                    <p><strong>Ship:</strong> {makeInventoryShip || userShip}</p>
                    <p><strong>Station:</strong> {inventoryStation || "N/A"}</p>
                    <p><strong>User:</strong> {userName || "N/A"}</p>
                    <p><strong>Code:</strong> {currentInventoryItem.code || "N/A"}</p>
                    <p><strong>Sheet:</strong> {currentInventoryItem.sheetName}</p>
                    <p><strong>Category:</strong> {currentInventoryItem.category}</p>

                    {!inventoryReady && (
                      <div style={styles.warningText}>
                        Choose ship, station, and user before confirming quantity.
                      </div>
                    )}

                    {currentStationSubmitted && (
                      <div style={styles.statusWarning}>
                        This station has already submitted count. You can view this item, but cannot update quantity.
                      </div>
                    )}

                    <label style={styles.label}>Insert quantity</label>
                    <input
                      autoFocus
                      type="number"
                      min="0"
                      value={inventoryQty}
                      onChange={(event) => setInventoryQty(event.target.value)}
                      onKeyDown={(event) => {
                        if (
                          event.key === "Enter" &&
                          inventoryReady &&
                          !inventoryLoading &&
                          !currentStationSubmitted
                        ) {
                          confirmInventoryQty();
                        }
                      }}
                      style={styles.searchInput}
                      placeholder="Enter quantity..."
                    />

                    <div style={styles.headerActions}>
                      <button
                        type="button"
                        style={styles.backButton}
                        onClick={() => {
                          setCurrentInventoryItem(null);
                          setInventoryQty("");
                          setEditingInventoryId(null);
                        }}
                      >
                        Cancel
                      </button>

                      <button
                        type="button"
                        style={styles.primaryButton}
                        onClick={confirmInventoryQty}
                        disabled={!inventoryReady || inventoryLoading || currentStationSubmitted}
                      >
                        {inventoryLoading
                          ? "Saving..."
                          : editingInventoryId
                            ? "Update Quantity"
                            : "Confirm Quantity"}
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          ) : null}

          <section style={styles.card}>
            <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 20 }}>
    <h2 style={styles.productTitle}>📄 Inventory Report</h2>
  </div>
            {inventoryReportMode === "summary" && (
              <div style={styles.reportFilterBox}>
                <label style={styles.label}>📍 {equipmentDepartment === "bar" ? "Bar Filter" : "Station Filter"}</label>
                <select
                  value={summaryStationFilter}
                  onChange={(e) => setSummaryStationFilter(e.target.value)}
                  style={styles.select}
                >
                  <option value="ALL">{equipmentDepartment === "bar" ? "All Bars" : "All Stations"}</option>
                  {summaryStationOptions.map((station) => (
                    <option key={station} value={station}>{station}</option>
                  ))}
                </select>

                <div style={styles.recipeMeta}>
                  Managers can filter the ship summary by galley/station to see where equipment is being consumed.
                </div>
              </div>
            )}
  <div style={styles.reportFilterBox}>
    <label style={styles.label}>
      📤 Upload Inventory Sheet Sample For Excel Report
    </label>

    <input
      type="file"
      accept=".xlsx,.xlsm"
      onChange={handleInventoryCountSheetTemplateFile}
      style={styles.fileInput}
    />

    <div style={styles.recipeMeta}>
      Export Excel will use this uploaded file as the sample/template. It will keep
      the same item positions and only write counts into column S.
    </div>

    {inventoryCountSheetTemplateName && (
      <div style={styles.statusGood}>
        Uploaded sample: {inventoryCountSheetTemplateName}
      </div>
    )}
  </div>

            <div style={styles.infoBox}>
              <div>🚢 Ship: <strong>{makeInventoryShip || userShip}</strong></div>
              {inventoryReportMode === "my" ? (
                <>
                  <div>📍 {equipmentDepartment === "bar" ? "Bar" : "Station"}: <strong>{inventoryStation || "Not selected"}</strong></div>
                  <div>👤 User: <strong>{userName || "Not selected"}</strong></div>
                  <div>✅ My records: <strong>{myReportRows.length}</strong></div>
                </>
              ) : (
                <>
                  <div>🌍 Report: <strong>All users for selected ship</strong></div>
                  <div>📍 Station Filter: <strong>{selectedSummaryStationLabel}</strong></div>
                  <div>📦 Summary items shown: <strong>{summaryReportRows.length}</strong></div>
                </>
              )}
            </div>

            {visibleReportRows.length === 0 && (
              <p style={styles.emptyText}>
                {inventoryReportMode === "summary"
                  ? `No ship summary records yet for ${selectedSummaryStationLabel}.`
                  : "Your confirmed quantities will appear here."}
              </p>
            )}

            <div style={styles.equipmentGrid}>
              {inventoryReportMode === "summary"
                ? visibleReportRows.map((item, index) => (
                    <div key={`${item.itemKey}-summary-${index}`} style={styles.equipmentCard}>
                      <div style={styles.recipeName}>{item.name}</div>
                      <div style={styles.recipeMeta}>Ship: {item.ship}</div>
                      <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                      <div style={styles.recipeMeta}>Category: {item.category}</div>
                      <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                      <div style={styles.statusGood}>Total Quantity: {formatQty(item.totalQty)}</div>
                      <div style={styles.recipeMeta}>Stations: {item.stations.join(", ") || "N/A"}</div>
                      {summaryStationFilter !== "ALL" && (
                        <div style={styles.statusWarning}>Filtered Station: {summaryStationFilter}</div>
                      )}
                      <div style={styles.recipeMeta}>Users: {item.users.join(", ") || "N/A"}</div>
                      <div style={styles.recipeMeta}>Records: {item.recordCount}</div>
                      <div style={styles.recipeMeta}>Last Updated: {item.confirmedAt}</div>
                    </div>
                  ))
                : visibleReportRows.map((item) => (
                    <div key={item.id} style={styles.equipmentCard}>
                  {item.image && (
    <img
      src={getImageUrl(item.image, "w360")}
      alt={item.name}
      style={styles.equipmentImage}
      onError={(event) => {
        event.currentTarget.style.display = "none";
      }}
    />
  )}
                      <div style={styles.recipeName}>{item.name}</div>
                      <div style={styles.recipeMeta}>Ship: {item.ship}</div>
                      <div style={styles.recipeMeta}>Station: {item.station}</div>
                      <div style={styles.recipeMeta}>User: {item.userName}</div>
                      <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                      <div style={styles.recipeMeta}>Category: {item.category}</div>
                      <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                      <div style={styles.statusGood}>Quantity: {formatQty(item.qty)}</div>
                      <div style={styles.recipeMeta}>Confirmed: {item.confirmedAt}</div>

                      <div style={styles.headerActions}>
                        <button style={styles.backButton} onClick={() => editInventoryItem(item)}>
                          ✏️ Edit
                        </button>

                        <button style={styles.deleteButton} onClick={() => deleteInventoryItem(item)}>
                          🗑️ Delete
                        </button>
                      </div>
                    </div>
                  ))}
                      </div>

            <div style={styles.finishBar}>
              {inventoryReportMode === "my" ? (
                <>
                  <div>
                    <strong>{inventoryStation || "Station"}</strong>
                    <div style={styles.recipeMeta}>
                      Live count: {myReportRows.length} / {makeInventoryItems.length || 0}
                    </div>
                    <div style={styles.recipeMeta}>
                      Status: {currentStationProgress?.statusLabel || "Not Started"}
                    </div>
                  </div>

                  {currentStationSubmitted ? (
                    <div style={styles.statusWarning}>
                      ✅ {inventoryStation} - Count Submitted. Waiting for all stations.
                    </div>
                  ) : (
                    <button
                      style={styles.primaryButton}
                      onClick={submitInventoryStationCount}
                      disabled={!inventoryReady || reportBusy}
                    >
                      ✅ Finish / Submit Station Count
                    </button>
                  )}
                </>
              ) : (
                <>
                  <div>
                    <strong>Final Inventory Report</strong>
                    <div style={styles.recipeMeta}>
                      Submitted stations: {submittedStations} / {stationProgressRows.length}
                    </div>
                    <div style={styles.recipeMeta}>
                      Uploaded sample: {inventoryCountSheetTemplateName || "Not uploaded"}
                    </div>
                  </div>

                  {allStationsSubmitted ? (
                    <button
                      style={styles.primaryButton}
                      onClick={generateFinalInventoryReport}
                      disabled={reportBusy}
                    >
                      📥 Generate Final Report
                    </button>
                  ) : (
                    <div style={styles.statusWarning}>
                      ⏳ Final report will unlock after all stations submit.
                    </div>
                  )}
                </>
              )}
            </div>
          </section>

          <section style={styles.card}>
            <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: showVariance ? 20 : 0 }}>
              <h2 style={styles.productTitle}>📊 Count Status</h2>

              <div style={styles.headerActions}>
                <button
                  style={styles.backButton}
                  onClick={() => setShowVariance((value) => !value)}
                >
                  {showVariance ? "Hide Status" : "Open Status"}
                </button>

                {showVariance && (
                  <>
                    <button style={styles.backButton} onClick={printInventoryStatus}>
                      🖨️ Print
                    </button>

                    <button style={styles.primaryButton} onClick={exportInventoryStatusToExcel}>
                      📥 Export Excel
                    </button>
                  </>
                )}
              </div>
            </div>

            {showVariance && (
              <>
                <div style={styles.infoBox}>
                  <div>📋 Master items: <strong>{makeInventoryItems.length}</strong></div>
                  <div>✅ Counted by me: <strong>{statusCountedItems.length}</strong></div>
                  <div>📋 Remaining for me: <strong>{statusPendingItems.length}</strong></div>
                </div>

                {makeInventoryItems.length === 0 && (
                  <p style={styles.emptyText}>Upload the master inventory file to see count status.</p>
                )}

                <h3 style={styles.sectionTitle}>📋 My Master Inventory Status</h3>

                <div style={styles.equipmentGrid}>
                  {inventoryStatusRows.map((item, index) => (
                    <div
                      key={`${item.code}-status-${index}`}
                      style={{ ...styles.equipmentCard, ...(item.status === "Counted" ? styles.countedCard : {}) }}
                    >
                      <div style={styles.recipeName}>{item.name}</div>
                      <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                      <div style={styles.recipeMeta}>Category: {item.category}</div>
                      <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>

                      {item.status === "Counted" ? (
                        <>
                          <div style={styles.statusGood}>Counted: {formatQty(item.countedQty)}</div>
                          <div style={styles.recipeMeta}>Counted: {item.countedAt}</div>
                        </>
                      ) : (
                        <div style={styles.statusNeutral}>Pending Count</div>
                      )}
                    </div>
                  ))}
                </div>
              </>
            )}
          </section>
        </main>
      );
}

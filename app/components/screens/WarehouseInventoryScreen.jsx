"use client";

import React from "react";

export default function WarehouseInventoryScreen(props) {
  const {
    styles,
    activeEquipmentDepartmentLabel,
    formatQty,
    getEquipmentDisplayImage,
    getImageUrl,
    getShipDisplayName,
    loadMasterInventoryItems,
    makeInventoryShip,
    parseWarehouseItems,
    selectedEquipment,
    setEquipmentMode,
    setSelectedEquipment,
    setWarehouseFilter,
    setWarehouseSearch,
    uploadWarehouseFile,
    userShip,
    warehouseFilter,
    warehouseMessage,
    warehouseRows,
    warehouseSearch,
  } = props;

  const allWarehouseItems = parseWarehouseItems();
      const isWarehouseOverstock = (item) => Number(item.onHand || 0) - Number(item.par || 0) > 10;
      const needsWarehouseOrder = (item) => Number(item.suggested || 0) > 0;
      const warehouseItems = allWarehouseItems.filter((item) => {
        if (warehouseFilter === "overstock") return isWarehouseOverstock(item);
        if (warehouseFilter === "needsOrder") return needsWarehouseOrder(item);
        return true;
      });
      const totalSuggested = allWarehouseItems.reduce((sum, item) => sum + item.suggested, 0);
      const criticalItems = allWarehouseItems.filter(needsWarehouseOrder).length;
      const warehouseItemsWithPictures = allWarehouseItems.filter((item) => item.image).length;
      const warehouseOverstockItems = allWarehouseItems.filter(isWarehouseOverstock).length;

      return (
        <main style={styles.page}>
          <header style={styles.header}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
            <div style={styles.headerActions}>
              <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
              <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
            </div>
          </header>

          <section style={styles.grid}>
            <div style={styles.card}>
              <h2 style={styles.cardTitle}>📤 Upload Warehouse Inventory</h2>
              <label style={styles.label}>Warehouse inventory file</label>
              <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadWarehouseFile} style={styles.fileInput} />

              {warehouseMessage && <p style={styles.message}>{warehouseMessage}</p>}

              <div style={styles.infoBox}>
                <div>📦 Items loaded: <strong>{allWarehouseItems.length}</strong></div>
                <div>👀 Showing now: <strong>{warehouseItems.length}</strong></div>
                <div>🖼️ Pictures matched: <strong>{warehouseItemsWithPictures}</strong></div>
                <div>🔵 Items needing order: <strong>{criticalItems}</strong></div>
                <div>🛒 Total suggested order: <strong>{formatQty(totalSuggested)}</strong></div>
                <div>🔴 Over par by more than 10: <strong>{warehouseOverstockItems}</strong></div>
                <div>A = Code, B = Name, G = Par, H = On Hand, M = Future Order</div>
                <div>Pictures are matched from the shared MEL master list by code/name.</div>
              </div>
            </div>

            <div style={styles.card}>
              <h2 style={styles.cardTitle}>🔍 Search Warehouse</h2>
              <input
                placeholder="Search code or equipment name..."
                value={warehouseSearch}
                onChange={(e) => setWarehouseSearch(e.target.value)}
                style={styles.searchInput}
              />
              <p style={styles.emptyText}>Blue = suggested order needed. Red = on hand is more than 10 above par level.</p>

              <div style={styles.viewModeBox}>
                <button
                  style={{ ...styles.viewModeButton, ...(warehouseFilter === "all" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setWarehouseFilter("all")}
                >
                  📋 All ({allWarehouseItems.length})
                </button>

                <button
                  style={{ ...styles.viewModeButton, ...(warehouseFilter === "overstock" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setWarehouseFilter("overstock")}
                >
                  🔴 Overstock ({warehouseOverstockItems})
                </button>

                <button
                  style={{ ...styles.viewModeButton, ...(warehouseFilter === "needsOrder" ? styles.viewModeButtonActive : {}) }}
                  onClick={() => setWarehouseFilter("needsOrder")}
                >
                  🔵 Needs Order ({criticalItems})
                </button>
              </div>

              <button style={styles.backButton} onClick={() => loadMasterInventoryItems(makeInventoryShip || userShip)}>
                🔄 Refresh Pictures
              </button>
            </div>
          </section>

          <section style={styles.card}>
            <h2 style={styles.productTitle}>🏬 {activeEquipmentDepartmentLabel} Inventory Warehouse</h2>

            {warehouseRows.length === 0 && <p style={styles.emptyText}>Upload the warehouse inventory file to begin.</p>}
            {warehouseRows.length > 0 && warehouseItems.length === 0 && (
              <p style={styles.emptyText}>
                No warehouse items match the current search/filter.
              </p>
            )}

            <div style={styles.equipmentGrid}>
              {warehouseItems.map((item, i) => {
    const displayImage = getEquipmentDisplayImage(item);
    const overstockAmount = Number(item.onHand || 0) - Number(item.par || 0);
    const isOverstock = overstockAmount > 10;

                return (
                <div
                  key={`${item.code}-${i}`}
                  style={{
                    ...styles.equipmentCard,
                    ...(item.suggested > 0 ? styles.orderNeededCard : {}),
                    ...(isOverstock ? styles.overstockCard : {}),
                  }}
                >
                  {displayImage ? (
                    <div>
                      <img
                        src={getImageUrl(displayImage)}
                        alt={item.name}
                        style={styles.equipmentImage}
                        onClick={() => setSelectedEquipment({
                          name: item.name,
                          code: item.code,
                          category: item.masterCategory || "Warehouse Item",
                          sheetName: item.masterSheetName || "Warehouse",
                          image: displayImage,
                        })}
                        onError={(e) => {
                          e.currentTarget.style.display = "none";
                          const link = e.currentTarget.nextElementSibling;
                          if (link) link.style.display = "block";
                        }}
                      />
                      <button
                        style={styles.imageButton}
                        onClick={() => setSelectedEquipment({
                          name: item.name,
                          code: item.code,
                          category: item.masterCategory || "Warehouse Item",
                          sheetName: item.masterSheetName || "Warehouse",
                          image: displayImage,
                        })}
                      >
                        View Picture
                      </button>
                      <a href={displayImage} target="_blank" rel="noreferrer" style={styles.imageLink}>Open Picture</a>
                    </div>
                  ) : (
                    <div style={styles.equipmentNoImage}>No image matched</div>
                  )}

                  <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Par Level: {formatQty(item.par)}</div>
                  <div style={styles.recipeMeta}>On Hand: {formatQty(item.onHand)}</div>
                  <div style={styles.recipeMeta}>Future Order: {formatQty(item.future)}</div>
                  {item.imageSource && <div style={styles.recipeMeta}>Picture match: {item.imageSource}</div>}
                  {isOverstock && (
                    <div style={styles.overstockWarning}>
                      Overstock Alert: {formatQty(overstockAmount)} above par
                    </div>
                  )}
                  <div style={item.suggested > 0 ? styles.suggestedOrderBad : styles.suggestedOrderGood}>
                    Suggested Next Order: {formatQty(item.suggested)}
                  </div>
                </div>
                );
              })}
            </div>

            {selectedEquipment && (
    <div style={styles.modalBackdrop} onClick={() => setSelectedEquipment(null)}>
      <div style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
        <button style={styles.closeButton} onClick={() => setSelectedEquipment(null)}>
          ✕
        </button>

        <h2>{selectedEquipment.name}</h2>
        <p><strong>Code:</strong> {selectedEquipment.code || "N/A"}</p>
        <p><strong>Sheet:</strong> {selectedEquipment.sheetName || "N/A"}</p>
        <p><strong>Category:</strong> {selectedEquipment.category || "N/A"}</p>

        {selectedEquipment.image || selectedEquipment.imageFallback ? (
          <div>
            <img
              src={getImageUrl(selectedEquipment.image || selectedEquipment.imageFallback)}
              alt={selectedEquipment.name}
              style={styles.modalImage}
              onError={(e) => {
                const fallbackSrc = selectedEquipment.imageFallback
                  ? getImageUrl(selectedEquipment.imageFallback)
                  : "";

                if (fallbackSrc && e.currentTarget.dataset.usedFallback !== "true") {
                  e.currentTarget.dataset.usedFallback = "true";
                  e.currentTarget.src = fallbackSrc;
                  return;
                }

                e.currentTarget.style.display = "none";
                const link = e.currentTarget.nextElementSibling;
                if (link) link.style.display = "block";
              }}
            />

            <a
              href={selectedEquipment.imageFallback || selectedEquipment.image}
              target="_blank"
              rel="noreferrer"
              style={{ ...styles.imageLink, display: "block" }}
            >
              Open Picture
            </a>
          </div>
        ) : (
          <div style={styles.equipmentNoImage}>No image</div>
        )}
      </div>
    </div>
  )}
          </section>
        </main>
      );
}

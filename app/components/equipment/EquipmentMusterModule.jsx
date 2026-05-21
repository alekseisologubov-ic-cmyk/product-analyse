"use client";

import React from "react";

export default function EquipmentMusterModule({
  styles,
  userShip,
  makeInventoryShip,
  isAdmin,
  equipmentDepartment,
  activeEquipmentDepartmentLabel,
  getShipDisplayName,
  groupedMuster,
  totalItems,
  musterItems,
  musterSearch,
  setMusterSearch,
  musterMessage,
  pictureLibraryMessage,
  pictureLibraryBusy,
  masterInventoryLoading,
  uploadMusterFile,
  loadMasterInventoryItems,
  syncMasterInventoryPicturesFromDrive,
  uploadEquipmentPictureZipFile = null,
  getEquipmentDisplayImage,
  getEquipmentFallbackImage,
  getImageUrl,
  selectedEquipment,
  setSelectedEquipment,
  onBack,
}) {
  const shipDisplayName =
    typeof getShipDisplayName === "function"
      ? getShipDisplayName(userShip)
      : userShip;

  const sheetCount = [...new Set((musterItems || []).map((item) => item.sheetName))].length;
  const groupCount = Object.keys(groupedMuster || {}).length;
  const refreshShip = makeInventoryShip || userShip;

  const getDisplayImage = (item) =>
    typeof getEquipmentDisplayImage === "function"
      ? getEquipmentDisplayImage(item)
      : item?.image || "";

  const getFallbackImage = (item) =>
    typeof getEquipmentFallbackImage === "function"
      ? getEquipmentFallbackImage(item)
      : item?.imageFallback || "";

  const getImageSrc = (value) =>
    typeof getImageUrl === "function" ? getImageUrl(value) : value;

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={styles.headerLogo}
        />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Back
          </button>

          <div style={styles.shipBadge}>🚢 {shipDisplayName}</div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Muster File</h2>

          <label style={styles.label}>Muster list file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadMusterFile}
            style={styles.fileInput}
          />

          <button
            style={styles.backButton}
            onClick={() => loadMasterInventoryItems(refreshShip)}
          >
            🔄 Refresh Shared Master List
          </button>

          {isAdmin && equipmentDepartment === "culinary" && (
            <>
              <button
                style={styles.backButton}
                onClick={syncMasterInventoryPicturesFromDrive}
                disabled={pictureLibraryBusy || masterInventoryLoading}
              >
                {pictureLibraryBusy ? "Syncing pictures..." : "🖼️ Sync Picture Library"}
              </button>

              {typeof uploadEquipmentPictureZipFile === "function" && (
                <>
                  <label style={styles.label}>Upload Culinary Picture ZIP</label>
                  <input
                    type="file"
                    accept=".zip"
                    onChange={uploadEquipmentPictureZipFile}
                    style={styles.fileInput}
                    disabled={pictureLibraryBusy || masterInventoryLoading}
                  />
                </>
              )}

              {pictureLibraryMessage && (
                <p style={styles.message}>{pictureLibraryMessage}</p>
              )}
            </>
          )}

          {musterMessage && <p style={styles.message}>{musterMessage}</p>}

          <div style={styles.infoBox}>
            <div>📋 Items loaded: <strong>{totalItems}</strong></div>
            <div>📄 Sheets included: <strong>{sheetCount}</strong></div>
            <div>🗂️ Groups: <strong>{groupCount}</strong></div>
            <div>C = Sub Category, D = Code, E = Name, H/I = Picture Link</div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔍 Search Equipment</h2>

          <input
            placeholder="Search equipment, code, sheet or sub category..."
            value={musterSearch}
            onChange={(event) => setMusterSearch(event.target.value)}
            style={styles.searchInput}
          />

          <p style={styles.emptyText}>
            Click any equipment card to open the picture and full details.
          </p>
        </div>
      </section>

      <section style={styles.card}>
        <h2 style={styles.productTitle}>
          📋 {activeEquipmentDepartmentLabel} Muster List
        </h2>

        {(musterItems || []).length === 0 && (
          <p style={styles.emptyText}>Upload the muster list file to begin.</p>
        )}

        {Object.entries(groupedMuster || {}).map(([category, items]) => (
          <div key={category} style={styles.equipmentCategory}>
            <h3 style={styles.sectionTitle}>🗂️ {category}</h3>

            <div style={styles.equipmentGrid}>
              {items.map((item, index) => {
                const displayImage = getDisplayImage(item);
                const fallbackImage = getFallbackImage(item);
                const imageSrc = displayImage ? getImageSrc(displayImage) : "";
                const fallbackSrc =
                  fallbackImage && fallbackImage !== displayImage
                    ? getImageSrc(fallbackImage)
                    : "";

                return (
                  <button
                    key={`${item.sheetName}-${item.code}-${index}`}
                    style={styles.equipmentCard}
                    onClick={() =>
                      setSelectedEquipment({
                        ...item,
                        image: displayImage,
                        imageFallback: fallbackImage,
                      })
                    }
                  >
                    {displayImage ? (
                      <div>
                        <img
                          src={imageSrc}
                          alt={item.name}
                          style={styles.equipmentImage}
                          data-fallback-src={fallbackSrc}
                          onError={(event) => {
                            const nextSrc = event.currentTarget.dataset.fallbackSrc;

                            if (
                              nextSrc &&
                              event.currentTarget.dataset.usedFallback !== "true"
                            ) {
                              event.currentTarget.dataset.usedFallback = "true";
                              event.currentTarget.src = nextSrc;
                              return;
                            }

                            event.currentTarget.style.display = "none";
                            const link = event.currentTarget.nextElementSibling;
                            if (link) link.style.display = "block";
                          }}
                        />

                        <a
                          href={fallbackImage || displayImage}
                          target="_blank"
                          rel="noreferrer"
                          style={styles.imageLink}
                        >
                          Open Picture
                        </a>
                      </div>
                    ) : (
                      <div style={styles.equipmentNoImage}>No image</div>
                    )}

                    <div style={styles.recipeName}>{item.name}</div>
                    <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                    <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                    <div style={styles.recipeMeta}>Category: {item.category}</div>
                  </button>
                );
              })}
            </div>
          </div>
        ))}

        {selectedEquipment && (
          <div
            style={styles.modalBackdrop}
            onClick={() => setSelectedEquipment(null)}
          >
            <div style={styles.modalCard} onClick={(event) => event.stopPropagation()}>
              <button
                style={styles.closeButton}
                onClick={() => setSelectedEquipment(null)}
              >
                ✕
              </button>

              <h2>{selectedEquipment.name}</h2>

              <p><strong>Code:</strong> {selectedEquipment.code || "N/A"}</p>
              <p><strong>Sheet:</strong> {selectedEquipment.sheetName || "N/A"}</p>
              <p><strong>Category:</strong> {selectedEquipment.category || "N/A"}</p>

              {selectedEquipment.image || selectedEquipment.imageFallback ? (
                <div>
                  <img
                    src={getImageSrc(selectedEquipment.image || selectedEquipment.imageFallback)}
                    alt={selectedEquipment.name}
                    style={styles.modalImage}
                    data-fallback-src={
                      selectedEquipment.imageFallback
                        ? getImageSrc(selectedEquipment.imageFallback)
                        : ""
                    }
                    onError={(event) => {
                      const fallbackSrc = event.currentTarget.dataset.fallbackSrc;

                      if (
                        fallbackSrc &&
                        event.currentTarget.dataset.usedFallback !== "true"
                      ) {
                        event.currentTarget.dataset.usedFallback = "true";
                        event.currentTarget.src = fallbackSrc;
                        return;
                      }

                      event.currentTarget.style.display = "none";
                      const link = event.currentTarget.nextElementSibling;
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

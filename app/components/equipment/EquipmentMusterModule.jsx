"use client";

import React, { useEffect, useMemo, useState } from "react";

const cleanMusterSearchText = (value) =>
  String(value || "").toLowerCase().replace(/\s+/g, " ").trim();
const getMusterDuplicateKey = (item) => {
  const code = String(item?.code || "")
    .trim()
    .replace(/\.0+$/g, "")
    .replace(/^0+(?=\d)/g, "")
    .replace(/\s+/g, "")
    .toUpperCase();

  if (code) return "CODE:" + code;

  const name = String(item?.name || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();

  return name ? "NAME:" + name : "";
};

const getMusterLocationText = (item) =>
  [item?.sheetName, item?.category].filter(Boolean).join(" / ");

const getMusterItemSearchText = (item) =>
  `${item?.sheetName || ""} ${item?.category || ""} ${item?.code || ""} ${
    item?.name || ""
  }`;

const createDedupedMusterItem = (item) => ({
  ...item,
  duplicateCount: 1,
  duplicateSearchText: getMusterItemSearchText(item),
  duplicateLocations: [getMusterLocationText(item)].filter(Boolean),
});

const mergeDedupedMusterItem = (existing, duplicate) => {
  const duplicateLocation = getMusterLocationText(duplicate);

  return {
    ...existing,

    // Keep the first visible row, but rescue image data from duplicate rows.
    image: existing.image || duplicate.image || "",
    imageFallback:
      existing.imageFallback ||
      duplicate.imageFallback ||
      duplicate.image ||
      "",

    duplicateCount: Number(existing.duplicateCount || 1) + 1,
    duplicateSearchText:
      String(existing.duplicateSearchText || "") +
      " " +
      getMusterItemSearchText(duplicate),

    duplicateLocations: [
      ...new Set([
        ...(existing.duplicateLocations || []),
        duplicateLocation,
      ].filter(Boolean)),
    ],
  };
};

const getGoogleDriveId = (value) => {
  const text = String(value || "").trim();

  const fileMatch = text.match(/drive\.google\.com\/file\/d\/([^/]+)/);
  const idMatch = text.match(/[?&]id=([^&]+)/);

  return fileMatch?.[1] || idMatch?.[1] || "";
};

const addUniqueImageCandidate = (list, value) => {
  const text = String(value || "").trim();

  if (!text) return;
  if (list.includes(text)) return;

  list.push(text);
};

const buildImageCandidates = ({
  displayImage,
  fallbackImage,
  getImageSrc,
  size = "w360",
}) => {
  const candidates = [];

  [displayImage, fallbackImage].forEach((value) => {
    const rawValue = String(value || "").trim();

    if (!rawValue) return;

    const convertedValue =
      typeof getImageSrc === "function" ? getImageSrc(rawValue, size) : rawValue;

    addUniqueImageCandidate(candidates, convertedValue);

    const googleDriveId = getGoogleDriveId(rawValue);

    if (googleDriveId) {
      addUniqueImageCandidate(
        candidates,
        `https://drive.google.com/thumbnail?id=${googleDriveId}&sz=${size}`
      );

      addUniqueImageCandidate(
        candidates,
        `https://drive.google.com/uc?export=view&id=${googleDriveId}`
      );
    }

    addUniqueImageCandidate(candidates, rawValue);
  });

  return candidates;
};

function MusterImagePreview({
  styles,
  item,
  displayImage,
  fallbackImage,
  getImageSrc,
  height = 150,
  size = "w360",
  showOpenButton = false,
}) {
  const [candidateIndex, setCandidateIndex] = useState(0);

  const candidates = useMemo(
    () =>
      buildImageCandidates({
        displayImage,
        fallbackImage,
        getImageSrc,
        size,
      }),
    [displayImage, fallbackImage, getImageSrc, size]
  );

  const openCandidates = useMemo(
    () =>
      buildImageCandidates({
        displayImage,
        fallbackImage,
        getImageSrc,
        size: "w1200",
      }),
    [displayImage, fallbackImage, getImageSrc]
  );

  useEffect(() => {
    setCandidateIndex(0);
  }, [displayImage, fallbackImage]);

  const previewSrc = candidates[candidateIndex] || "";
  const openSrc = openCandidates[0] || candidates[0] || displayImage || fallbackImage || "";
  const hasAnyImage = Boolean(displayImage || fallbackImage);
  const previewUnavailable = hasAnyImage && !previewSrc;

  const openPicture = (event) => {
    event.stopPropagation();

    if (!openSrc) return;

    window.open(openSrc, "_blank", "noopener,noreferrer");
  };

  if (!hasAnyImage) {
    return <div style={styles.equipmentNoImage}>No image</div>;
  }

  return (
    <div>
      {!previewUnavailable ? (
        <div
          style={{
            width: "100%",
            height,
            borderRadius: 10,
            overflow: "hidden",
            background: "#eee",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <img
            src={previewSrc}
            alt={item?.name || "Equipment picture"}
            loading="lazy"
            decoding="async"
            style={{
              width: "100%",
              height: "100%",
              objectFit: "cover",
              display: "block",
              cursor: "pointer",
              background: "#eee",
            }}
            onClick={(event) => {
              event.stopPropagation();
              openPicture(event);
            }}
            onError={() => {
              setCandidateIndex((currentIndex) => {
                const nextIndex = currentIndex + 1;

                if (nextIndex < candidates.length) {
                  return nextIndex;
                }

                return candidates.length;
              });
            }}
          />
        </div>
      ) : (
        <div
          style={{
            width: "100%",
            height,
            borderRadius: 10,
            background: "#eee",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            color: "#777",
            fontWeight: "bold",
            fontSize: 13,
            textAlign: "center",
            padding: 8,
            boxSizing: "border-box",
          }}
        >
          Picture preview unavailable
        </div>
      )}

      {(showOpenButton || previewUnavailable) && openSrc && (
        <button
          type="button"
          style={styles.imageButton}
          onClick={openPicture}
        >
          Open Picture
        </button>
      )}
    </div>
  );
}

export default function EquipmentMusterModule({
  styles,
  userShip,
  makeInventoryShip,
  isAdmin,
  equipmentDepartment,
  activeEquipmentDepartmentLabel,
  getShipDisplayName,
  musterItems,
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
  onBack,
}) {
  const [musterSearch, setMusterSearch] = useState("");
  const [selectedEquipment, setSelectedEquipment] = useState(null);

  const departmentLabel = activeEquipmentDepartmentLabel || "Equipment";

  const shipDisplayName =
    typeof getShipDisplayName === "function"
      ? getShipDisplayName(userShip)
      : userShip;

  const refreshShip = makeInventoryShip || userShip;

  const groupedMuster = useMemo(() => {
  const grouped = {};
  const query = cleanMusterSearchText(musterSearch);
  const uniqueByCode = new Map();

  (musterItems || []).forEach((item) => {
    const duplicateKey = getMusterDuplicateKey(item);

    if (!duplicateKey) return;

    if (!uniqueByCode.has(duplicateKey)) {
      uniqueByCode.set(duplicateKey, createDedupedMusterItem(item));
      return;
    }

    uniqueByCode.set(
      duplicateKey,
      mergeDedupedMusterItem(uniqueByCode.get(duplicateKey), item)
    );
  });

  Array.from(uniqueByCode.values()).forEach((item) => {
    const searchText = cleanMusterSearchText(
      item.duplicateSearchText || getMusterItemSearchText(item)
    );

    if (query && !searchText.includes(query)) return;

    const groupKey = `${item.sheetName || "Unknown Sheet"} / ${
      item.category || "Uncategorized"
    }`;

    if (!grouped[groupKey]) {
      grouped[groupKey] = [];
    }

    grouped[groupKey].push(item);
  });

  return grouped;
}, [musterItems, musterSearch]);

  const totalItems = useMemo(
    () =>
      Object.values(groupedMuster).reduce(
        (sum, items) => sum + items.length,
        0
      ),
    [groupedMuster]
  );
  const hiddenDuplicateCount = useMemo(
  () =>
    Object.values(groupedMuster).reduce(
      (sum, items) =>
        sum +
        items.reduce(
          (itemSum, item) =>
            itemSum + Math.max(Number(item.duplicateCount || 1) - 1, 0),
          0
        ),
      0
    ),
  [groupedMuster]
);

  const sheetCount = useMemo(
    () =>
      [...new Set((musterItems || []).map((item) => item.sheetName))].filter(
        Boolean
      ).length,
    [musterItems]
  );

  const groupCount = Object.keys(groupedMuster || {}).length;

  const getDisplayImage = (item) =>
    typeof getEquipmentDisplayImage === "function"
      ? getEquipmentDisplayImage(item)
      : item?.image || item?.imageFallback || "";

  const getFallbackImage = (item) =>
    typeof getEquipmentFallbackImage === "function"
      ? getEquipmentFallbackImage(item)
      : item?.imageFallback || "";

  const getImageSrc = (value, size = "w800") =>
    typeof getImageUrl === "function" ? getImageUrl(value, size) : String(value || "").trim();

  const openSelectedEquipment = (item) => {
    const displayImage = getDisplayImage(item);
    const fallbackImage = getFallbackImage(item);

    setSelectedEquipment({
      ...item,
      image: displayImage,
      imageFallback: fallbackImage,
    });
  };

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={styles.headerLogo}
        />

        <div style={styles.headerActions}>
          <button type="button" style={styles.backButton} onClick={onBack}>
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
            type="button"
            style={styles.backButton}
            onClick={() => loadMasterInventoryItems(refreshShip)}
          >
            🔄 Refresh Shared Master List
          </button>

          {isAdmin && equipmentDepartment === "culinary" && (
            <>
              <button
                type="button"
                style={styles.backButton}
                onClick={syncMasterInventoryPicturesFromDrive}
                disabled={pictureLibraryBusy || masterInventoryLoading}
              >
                {pictureLibraryBusy
                  ? "Syncing pictures..."
                  : "🖼️ Sync Picture Library"}
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
            <div>
              📋 Items shown: <strong>{totalItems}</strong>
            </div>
            <div>
              📋 Master items loaded: <strong>{(musterItems || []).length}</strong>
            </div>
            <div>
              📄 Sheets included: <strong>{sheetCount}</strong>
            </div>
            <div>
              🗂️ Groups shown: <strong>{groupCount}</strong>
            </div>
            <div>
              C = Sub Category, D = Code, E = Name, I = Product Picture, H =
              backup only
            </div>
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
        <h2 style={styles.productTitle}>📋 {departmentLabel} Muster List</h2>

        {(musterItems || []).length === 0 && (
          <p style={styles.emptyText}>Upload the muster list file to begin.</p>
        )}

        {(musterItems || []).length > 0 && totalItems === 0 && (
          <p style={styles.emptyText}>No equipment matched your search.</p>
        )}

        {Object.entries(groupedMuster || {}).map(([category, items]) => (
          <div key={category} style={styles.equipmentCategory}>
            <h3 style={styles.sectionTitle}>🗂️ {category}</h3>

            <div style={styles.equipmentGrid}>
              {items.map((item, index) => {
                const displayImage = getDisplayImage(item);
                const fallbackImage = getFallbackImage(item);

                return (
                  <div
                    key={`${item.sheetName || "sheet"}-${
                      item.code || "code"
                    }-${index}`}
                    role="button"
                    tabIndex={0}
                    style={styles.equipmentCard}
                    onClick={() => openSelectedEquipment(item)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter" || event.key === " ") {
                        event.preventDefault();
                        openSelectedEquipment(item);
                      }
                    }}
                  >
                    <MusterImagePreview
                      styles={styles}
                      item={item}
                      displayImage={displayImage}
                      fallbackImage={fallbackImage}
                      getImageSrc={getImageSrc}
                      height={150}
                      size="w360"
                      showOpenButton={false}
                    />

                    <div style={styles.recipeName}>{item.name}</div>
                    <div style={styles.recipeMeta}>
                      Code: {item.code || "N/A"}
                    </div>
                    <div style={styles.recipeMeta}>
                      Sheet: {item.sheetName || "N/A"}
                    </div>
                    <div style={styles.recipeMeta}>
                      Category: {item.category || "N/A"}
                    </div>
                  </div>
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
            <div
              style={styles.modalCard}
              onClick={(event) => event.stopPropagation()}
            >
              <button
                type="button"
                style={styles.closeButton}
                onClick={() => setSelectedEquipment(null)}
              >
                ✕
              </button>

              <h2>{selectedEquipment.name}</h2>

              <p>
                <strong>Code:</strong> {selectedEquipment.code || "N/A"}
              </p>
              <p>
                <strong>Sheet:</strong> {selectedEquipment.sheetName || "N/A"}
              </p>
              <p>
                <strong>Category:</strong>{" "}
                {selectedEquipment.category || "N/A"}
              </p>

              <MusterImagePreview
                styles={styles}
                item={selectedEquipment}
                displayImage={selectedEquipment.image}
                fallbackImage={selectedEquipment.imageFallback}
                getImageSrc={getImageSrc}
                height={420}
                size="w1200"
                showOpenButton={true}
              />
            </div>
          </div>
        )}
      </section>
    </main>
  );
}

"use client";

import React, { useEffect, useMemo, useState } from "react";

const cleanMusterSearchText = (value) =>
  String(value || "").toLowerCase().replace(/\s+/g, " ").trim();

const cleanMusterKeyText = (value) =>
  String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

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

const getMusterStationName = (item, equipmentDepartment) => {
  const department = String(equipmentDepartment || "").toLowerCase();

  if (department === "restaurant") {
    const stationName =
      item?.stationName ||
      item?.sheetName ||
      item?.sourceSheetName ||
      "MASTER";

    return (
      String(stationName || "MASTER").replace(/\s+/g, " ").trim() || "MASTER"
    );
  }

  return (
    String(item?.sheetName || "Unknown Sheet")
      .replace(/\s+/g, " ")
      .trim() || "Unknown Sheet"
  );
};

const getMusterStationKey = (value) =>
  cleanMusterKeyText(value || "MASTER") || "MASTER";

const isMasterStationName = (value) => getMusterStationKey(value) === "MASTER";

const isMasterRestaurantItem = (item, equipmentDepartment) => {
  const department = String(equipmentDepartment || "").toLowerCase();

  if (department !== "restaurant") return false;

  if (item?.isMasterSheet) return true;

  const stationName = getMusterStationName(item, equipmentDepartment);
  const sheetName = item?.sheetName;
  const category = item?.category;

  return (
    isMasterStationName(stationName) ||
    isMasterStationName(sheetName) ||
    isMasterStationName(category)
  );
};

const getMusterRestaurantSectionName = (item, equipmentDepartment) => {
  if (isMasterRestaurantItem(item, equipmentDepartment)) {
    return "MASTER";
  }

  const stationName = getMusterStationName(item, equipmentDepartment);
  const category = String(item?.category || item?.sourceSheetName || "")
    .replace(/\s+/g, " ")
    .trim();

  const stationKey = getMusterStationKey(stationName);
  const categoryKey = getMusterStationKey(category);

  if (!category || categoryKey === stationKey) {
    return stationName || "Restaurant";
  }

  return `${stationName} - ${category}`;
};

const getMusterRestaurantSectionKey = (item, equipmentDepartment) => {
  if (isMasterRestaurantItem(item, equipmentDepartment)) {
    return "MASTER";
  }

  const stationName = getMusterStationName(item, equipmentDepartment);
  const category = String(item?.category || item?.sourceSheetName || stationName)
    .replace(/\s+/g, " ")
    .trim();

  const stationKey = getMusterStationKey(stationName);
  const categoryKey = getMusterStationKey(category);

  return `${stationKey}|||${categoryKey}`;
};

const getMusterLocationText = (item) =>
  [
    item?.stationName || item?.sheetName,
    item?.sourceSheetName,
    item?.category,
  ]
    .filter(Boolean)
    .join(" / ");

const getMusterItemSearchText = (item) =>
  `${item?.stationName || ""} ${item?.sheetName || ""} ${
    item?.sourceSheetName || ""
  } ${item?.category || ""} ${item?.code || ""} ${item?.name || ""}`;

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
  const openSrc =
    openCandidates[0] || candidates[0] || displayImage || fallbackImage || "";
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
        <button type="button" style={styles.imageButton} onClick={openPicture}>
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
  const [selectedRestaurantSection, setSelectedRestaurantSection] = useState("");

  const departmentLabel = activeEquipmentDepartmentLabel || "Equipment";
  const isRestaurantDepartment =
    String(equipmentDepartment || "").toLowerCase() === "restaurant";

  const shipDisplayName =
    typeof getShipDisplayName === "function"
      ? getShipDisplayName(userShip)
      : userShip;

  const refreshShip = makeInventoryShip || userShip;

  const getDisplayImage = (item) =>
    typeof getEquipmentDisplayImage === "function"
      ? getEquipmentDisplayImage(item)
      : item?.image || item?.imageFallback || "";

  const getFallbackImage = (item) =>
    typeof getEquipmentFallbackImage === "function"
      ? getEquipmentFallbackImage(item)
      : item?.imageFallback || "";

  const getImageSrc = (value, size = "w800") =>
    typeof getImageUrl === "function"
      ? getImageUrl(value, size)
      : String(value || "").trim();

  const allRestaurantSectionKeys = useMemo(() => {
    if (!isRestaurantDepartment) return new Set();

    return new Set(
      (musterItems || [])
        .map((item) =>
          getMusterRestaurantSectionKey(item, equipmentDepartment)
        )
        .filter(Boolean)
    );
  }, [musterItems, equipmentDepartment, isRestaurantDepartment]);

  useEffect(() => {
    if (!isRestaurantDepartment) {
      setSelectedRestaurantSection("");
      return;
    }

    if (
      selectedRestaurantSection &&
      !allRestaurantSectionKeys.has(selectedRestaurantSection)
    ) {
      setSelectedRestaurantSection("");
    }
  }, [
    isRestaurantDepartment,
    selectedRestaurantSection,
    allRestaurantSectionKeys,
  ]);

  useEffect(() => {
    setSelectedEquipment(null);
  }, [selectedRestaurantSection, equipmentDepartment]);

  const selectedRestaurantSectionName = useMemo(() => {
    if (!selectedRestaurantSection) return "";

    const matchingItem = (musterItems || []).find(
      (item) =>
        getMusterRestaurantSectionKey(item, equipmentDepartment) ===
        selectedRestaurantSection
    );

    return (
      getMusterRestaurantSectionName(matchingItem, equipmentDepartment) ||
      selectedRestaurantSection
    );
  }, [musterItems, equipmentDepartment, selectedRestaurantSection]);

  const restaurantSectionCards = useMemo(() => {
    if (!isRestaurantDepartment) return [];

    const query = cleanMusterSearchText(musterSearch);
    const sectionMap = new Map();

    (musterItems || []).forEach((item) => {
      const sectionKey = getMusterRestaurantSectionKey(
        item,
        equipmentDepartment
      );

      const sectionName = getMusterRestaurantSectionName(
        item,
        equipmentDepartment
      );

      const stationName = getMusterStationName(item, equipmentDepartment);
      const category = String(item?.category || sectionName || "")
        .replace(/\s+/g, " ")
        .trim();

      if (!sectionMap.has(sectionKey)) {
        sectionMap.set(sectionKey, {
          sectionKey,
          sectionName,
          stationName,
          category,
          rawItems: [],
          uniqueKeys: new Set(),
          pictureCount: 0,
          searchText: `${sectionName} ${stationName} ${category}`,
          sampleImage: "",
          sampleFallbackImage: "",
          isMaster: isMasterRestaurantItem(item, equipmentDepartment),
        });
      }

      const section = sectionMap.get(sectionKey);
      const duplicateKey = getMusterDuplicateKey(item);

      section.rawItems.push(item);

      if (duplicateKey) {
        section.uniqueKeys.add(duplicateKey);
      }

      const image = String(item?.image || "").trim();
      const fallbackImage = String(item?.imageFallback || "").trim();

      if (image || fallbackImage) {
        section.pictureCount += 1;
      }

      if (!section.sampleImage && image) {
        section.sampleImage = image;
      }

      if (!section.sampleFallbackImage && fallbackImage) {
        section.sampleFallbackImage = fallbackImage;
      }

      section.searchText += " " + getMusterItemSearchText(item);
    });

    return Array.from(sectionMap.values())
      .map((section) => ({
        ...section,
        rawCount: section.rawItems.length,
        uniqueCount: section.uniqueKeys.size || section.rawItems.length,
      }))
      .filter((section) => {
        if (!query) return true;

        return cleanMusterSearchText(section.searchText).includes(query);
      })
      .sort((a, b) => {
        if (a.isMaster !== b.isMaster) return a.isMaster ? -1 : 1;

        return String(a.sectionName || "").localeCompare(
          String(b.sectionName || "")
        );
      });
  }, [
    musterItems,
    musterSearch,
    equipmentDepartment,
    isRestaurantDepartment,
  ]);

  const visibleMusterItems = useMemo(() => {
    if (!isRestaurantDepartment) {
      return musterItems || [];
    }

    if (!selectedRestaurantSection) {
      return [];
    }

    return (musterItems || []).filter(
      (item) =>
        getMusterRestaurantSectionKey(item, equipmentDepartment) ===
        selectedRestaurantSection
    );
  }, [
    musterItems,
    equipmentDepartment,
    isRestaurantDepartment,
    selectedRestaurantSection,
  ]);

  const groupedMuster = useMemo(() => {
    const grouped = {};
    const query = cleanMusterSearchText(musterSearch);
    const uniqueByCode = new Map();

    (visibleMusterItems || []).forEach((item) => {
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

      const groupKey = isRestaurantDepartment
        ? item.category || selectedRestaurantSectionName || "Restaurant"
        : `${item.sheetName || "Unknown Sheet"} / ${
            item.category || "Uncategorized"
          }`;

      if (!grouped[groupKey]) {
        grouped[groupKey] = [];
      }

      grouped[groupKey].push(item);
    });

    return grouped;
  }, [
    visibleMusterItems,
    musterSearch,
    isRestaurantDepartment,
    selectedRestaurantSectionName,
  ]);

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

  const sheetCount = useMemo(() => {
    if (isRestaurantDepartment) {
      const values = (musterItems || []).map((item) =>
        getMusterRestaurantSectionName(item, equipmentDepartment)
      );

      return [...new Set(values)].filter(Boolean).length;
    }

    return [...new Set((musterItems || []).map((item) => item.sheetName))]
      .filter(Boolean)
      .length;
  }, [musterItems, equipmentDepartment, isRestaurantDepartment]);

  const groupCount = Object.keys(groupedMuster || {}).length;

  const restaurantUniqueItemsShown = restaurantSectionCards.reduce(
    (sum, section) => sum + Number(section.uniqueCount || 0),
    0
  );

  const restaurantPicturesShown = restaurantSectionCards.reduce(
    (sum, section) => sum + Number(section.pictureCount || 0),
    0
  );

  const openSelectedEquipment = (item) => {
    const displayImage = getDisplayImage(item);
    const fallbackImage = getFallbackImage(item);

    setSelectedEquipment({
      ...item,
      image: displayImage,
      imageFallback: fallbackImage,
    });
  };

  const openRestaurantSection = (sectionKey) => {
    setSelectedRestaurantSection(sectionKey);
    setMusterSearch("");
  };

  const searchPlaceholder =
    isRestaurantDepartment && !selectedRestaurantSection
      ? "Search MASTER, restaurant, category, code or equipment..."
      : "Search equipment, code, sheet or subcategory...";

  const helpText = isRestaurantDepartment
    ? "Restaurant file: A = picture, B = code, C = description. MASTER and each restaurant/category tab open as cards first."
    : equipmentDepartment === "bar"
    ? "Bar file: picture is read from column D or column I when available."
    : "C = Subcategory, D = Code, E = Name, I = Product Picture, H = backup only.";

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
          <h2 style={styles.cardTitle}>📤 Upload Master List File</h2>

          <label style={styles.label}>Master list file</label>
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
            disabled={masterInventoryLoading}
          >
            {masterInventoryLoading
              ? "Refreshing..."
              : "🔄 Refresh Shared Master List"}
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
            {isRestaurantDepartment && !selectedRestaurantSection ? (
              <>
                <div>
                  🍽️ MASTER / section cards shown:{" "}
                  <strong>{restaurantSectionCards.length}</strong>
                </div>

                <div>
                  📦 Unique equipment records shown:{" "}
                  <strong>{restaurantUniqueItemsShown}</strong>
                </div>

                <div>
                  🖼️ Rows with pictures:{" "}
                  <strong>{restaurantPicturesShown}</strong>
                </div>

                <div>
                  📋 Total rows loaded:{" "}
                  <strong>{(musterItems || []).length}</strong>
                </div>

                <div>
                  📄 MASTER / sections included: <strong>{sheetCount}</strong>
                </div>

                <div>{helpText}</div>
              </>
            ) : (
              <>
                {isRestaurantDepartment && selectedRestaurantSection && (
                  <div>
                    🍽️ Selected:{" "}
                    <strong>{selectedRestaurantSectionName}</strong>
                  </div>
                )}

                <div>
                  📋 Items shown: <strong>{totalItems}</strong>
                </div>

                <div>
                  🔁 Duplicate codes hidden:{" "}
                  <strong>{hiddenDuplicateCount}</strong>
                </div>

                <div>
                  📋 Master rows loaded:{" "}
                  <strong>{(musterItems || []).length}</strong>
                </div>

                <div>
                  📄 Sheets / sections included: <strong>{sheetCount}</strong>
                </div>

                <div>
                  🗂️ Groups shown: <strong>{groupCount}</strong>
                </div>

                <div>{helpText}</div>
              </>
            )}
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔍 Search Equipment</h2>

          <input
            placeholder={searchPlaceholder}
            value={musterSearch}
            onChange={(event) => setMusterSearch(event.target.value)}
            style={styles.searchInput}
          />

          {isRestaurantDepartment && selectedRestaurantSection && (
            <button
              type="button"
              style={styles.backButton}
              onClick={() => {
                setSelectedRestaurantSection("");
                setMusterSearch("");
              }}
            >
              ← Back to MASTER / Sections
            </button>
          )}

          <p style={styles.emptyText}>
            {isRestaurantDepartment && !selectedRestaurantSection
              ? "Select MASTER or a restaurant/category card first. Then the equipment for that section will be shown."
              : "Click any equipment card to open the picture and full details."}
          </p>
        </div>
      </section>

      <section style={styles.card}>
        <div
          style={{
            ...styles.header,
            boxShadow: "none",
            padding: 0,
            marginBottom: 16,
          }}
        >
          <div>
            <h2 style={styles.productTitle}>
              {isRestaurantDepartment && selectedRestaurantSection
                ? `🍽️ ${selectedRestaurantSectionName}`
                : `📋 ${departmentLabel} Master List`}
            </h2>

            {isRestaurantDepartment && selectedRestaurantSection && (
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Showing equipment assigned to{" "}
                <strong>{selectedRestaurantSectionName}</strong>.
              </p>
            )}

            {isRestaurantDepartment && !selectedRestaurantSection && (
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Choose MASTER or a restaurant/category section to view its
                equipment.
              </p>
            )}
          </div>

          {isRestaurantDepartment && selectedRestaurantSection && (
            <button
              type="button"
              style={styles.backButton}
              onClick={() => {
                setSelectedRestaurantSection("");
                setMusterSearch("");
              }}
            >
              ← MASTER / Sections
            </button>
          )}
        </div>

        {(musterItems || []).length === 0 && (
          <p style={styles.emptyText}>Upload the master list file to begin.</p>
        )}

        {isRestaurantDepartment &&
          (musterItems || []).length > 0 &&
          !selectedRestaurantSection && (
            <>
              {restaurantSectionCards.length === 0 && (
                <p style={styles.emptyText}>
                  No MASTER / restaurant section matched your search.
                </p>
              )}

              <div style={styles.equipmentGrid}>
                {restaurantSectionCards.map((section) => {
                  const isMaster = section.isMaster;

                  return (
                    <div
                      key={section.sectionKey}
                      role="button"
                      tabIndex={0}
                      style={{
                        ...styles.equipmentCard,
                        ...(isMaster ? styles.countedCard : {}),
                      }}
                      onClick={() => openRestaurantSection(section.sectionKey)}
                      onKeyDown={(event) => {
                        if (event.key === "Enter" || event.key === " ") {
                          event.preventDefault();
                          openRestaurantSection(section.sectionKey);
                        }
                      }}
                    >
                      {section.sampleImage || section.sampleFallbackImage ? (
                        <MusterImagePreview
                          styles={styles}
                          item={{ name: section.sectionName }}
                          displayImage={section.sampleImage}
                          fallbackImage={section.sampleFallbackImage}
                          getImageSrc={getImageSrc}
                          height={130}
                          size="w360"
                          showOpenButton={false}
                        />
                      ) : (
                        <div
                          style={{
                            fontSize: 42,
                            lineHeight: 1,
                            textAlign: "center",
                            marginBottom: 4,
                            minHeight: 130,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            background: "#f2f2f2",
                            borderRadius: 10,
                          }}
                        >
                          {isMaster ? "📘" : "🍽️"}
                        </div>
                      )}

                      <div style={styles.recipeName}>
                        {section.sectionName}
                      </div>

                      <div style={styles.recipeMeta}>
                        Type: {isMaster ? "Master full list" : "Restaurant section"}
                      </div>

                      {!isMaster && section.stationName && (
                        <div style={styles.recipeMeta}>
                          Restaurant: {section.stationName}
                        </div>
                      )}

                      {!isMaster && section.category && (
                        <div style={styles.recipeMeta}>
                          Category: {section.category}
                        </div>
                      )}

                      <div style={styles.statusGood}>
                        Equipment: {section.uniqueCount}
                      </div>

                      <div style={styles.recipeMeta}>
                        Total rows: {section.rawCount}
                      </div>

                      <div style={styles.recipeMeta}>
                        Pictures: {section.pictureCount}
                      </div>

                      <button
                        type="button"
                        style={styles.imageButton}
                        onClick={(event) => {
                          event.stopPropagation();
                          openRestaurantSection(section.sectionKey);
                        }}
                      >
                        Open Equipment
                      </button>
                    </div>
                  );
                })}
              </div>
            </>
          )}

        {(!isRestaurantDepartment || selectedRestaurantSection) && (
          <>
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
                        key={`${item.stationName || item.sheetName || "sheet"}-${
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

                        {isRestaurantDepartment && (
                          <div style={styles.recipeMeta}>
                            Section:{" "}
                            {getMusterRestaurantSectionName(
                              item,
                              equipmentDepartment
                            )}
                          </div>
                        )}

                        {isRestaurantDepartment && (
                          <div style={styles.recipeMeta}>
                            Restaurant:{" "}
                            {item.stationName || item.sheetName || "N/A"}
                          </div>
                        )}

                        <div style={styles.recipeMeta}>
                          Sheet:{" "}
                          {item.sourceSheetName || item.sheetName || "N/A"}
                        </div>

                        <div style={styles.recipeMeta}>
                          Category: {item.category || "N/A"}
                        </div>

                        {Number(item.duplicateCount || 1) > 1 && (
                          <div style={styles.statusNeutral}>
                            Duplicate code hidden:{" "}
                            {Number(item.duplicateCount || 1) - 1}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            ))}
          </>
        )}

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

              {isRestaurantDepartment && (
                <p>
                  <strong>Section:</strong>{" "}
                  {getMusterRestaurantSectionName(
                    selectedEquipment,
                    equipmentDepartment
                  )}
                </p>
              )}

              {isRestaurantDepartment && (
                <p>
                  <strong>Restaurant:</strong>{" "}
                  {selectedEquipment.stationName ||
                    selectedEquipment.sheetName ||
                    "N/A"}
                </p>
              )}

              <p>
                <strong>Sheet:</strong>{" "}
                {selectedEquipment.sourceSheetName ||
                  selectedEquipment.sheetName ||
                  "N/A"}
              </p>

              <p>
                <strong>Category:</strong>{" "}
                {selectedEquipment.category || "N/A"}
              </p>

              {Number(selectedEquipment.duplicateCount || 1) > 1 && (
                <>
                  <p>
                    <strong>Duplicate rows hidden:</strong>{" "}
                    {Number(selectedEquipment.duplicateCount || 1) - 1}
                  </p>

                  <p>
                    <strong>Found in:</strong>{" "}
                    {(selectedEquipment.duplicateLocations || []).join(", ") ||
                      "N/A"}
                  </p>
                </>
              )}

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

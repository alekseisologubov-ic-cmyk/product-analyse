"use client";

import React, { useMemo, useState } from "react";

const normalizeText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const getTokens = (value) =>
  normalizeText(value)
    .split(" ")
    .map((item) => item.trim())
    .filter((item) => item.length >= 3)
    .filter(
      (item) =>
        ![
          "THE",
          "AND",
          "FOR",
          "WITH",
          "HIGH",
          "LOW",
          "SMALL",
          "LARGE",
          "SILVER",
          "STEEL",
          "S/S",
          "EQUIPMENT",
        ].includes(item)
    );

const resizeImageFileToDataUrl = (
  file,
  { maxWidth = 1280, maxHeight = 1280, quality = 0.72 } = {}
) =>
  new Promise((resolve, reject) => {
    const image = new Image();
    const objectUrl = URL.createObjectURL(file);

    image.onload = () => {
      try {
        const originalWidth = image.naturalWidth || image.width;
        const originalHeight = image.naturalHeight || image.height;

        const scale = Math.min(
          maxWidth / originalWidth,
          maxHeight / originalHeight,
          1
        );

        const targetWidth = Math.max(1, Math.round(originalWidth * scale));
        const targetHeight = Math.max(1, Math.round(originalHeight * scale));

        const canvas = document.createElement("canvas");
        canvas.width = targetWidth;
        canvas.height = targetHeight;

        const context = canvas.getContext("2d");
        context.drawImage(image, 0, 0, targetWidth, targetHeight);

        const dataUrl = canvas.toDataURL("image/jpeg", quality);

        URL.revokeObjectURL(objectUrl);
        resolve(dataUrl);
      } catch (error) {
        URL.revokeObjectURL(objectUrl);
        reject(error);
      }
    };

    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error("Could not resize image."));
    };

    image.src = objectUrl;
  });

const waitForBrowser = () => new Promise((resolve) => setTimeout(resolve, 0));

const getComparableImageUrl = (url) => {
  const value = String(url || "").trim();
  if (!value) return "";

  if (value.startsWith("data:image/")) return value;

  const googleDriveFileMatch = value.match(/drive\.google\.com\/file\/d\/([^/]+)/);
  const googleDriveIdMatch = value.match(/[?&]id=([^&]+)/);
  const googleDriveId = googleDriveFileMatch?.[1] || googleDriveIdMatch?.[1];

  if (googleDriveId) {
    return `https://drive.google.com/thumbnail?id=${googleDriveId}&sz=w360`;
  }

  if (value.includes("sharepoint.com") || value.includes("1drv.ms")) {
    return value.includes("?") ? `${value}&download=1` : `${value}?download=1`;
  }

  return value;
};

const loadImageForCompare = (src) =>
  new Promise((resolve, reject) => {
    const image = new Image();

    image.crossOrigin = "anonymous";

    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Could not load comparison image."));

    image.src = src;
  });

const getImageHash = async (src) => {
  const image = await loadImageForCompare(src);

  const size = 16;
  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;

  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.drawImage(image, 0, 0, size, size);

  const imageData = context.getImageData(0, 0, size, size).data;
  const grayValues = [];

  for (let index = 0; index < imageData.length; index += 4) {
    const red = imageData[index];
    const green = imageData[index + 1];
    const blue = imageData[index + 2];

    grayValues.push(red * 0.299 + green * 0.587 + blue * 0.114);
  }

  const average =
    grayValues.reduce((sum, value) => sum + value, 0) / grayValues.length;

  return grayValues.map((value) => (value >= average ? 1 : 0));
};

const getHashSimilarity = (leftHash, rightHash) => {
  if (!leftHash?.length || !rightHash?.length) return 0;

  const length = Math.min(leftHash.length, rightHash.length);
  let same = 0;

  for (let index = 0; index < length; index += 1) {
    if (leftHash[index] === rightHash[index]) same += 1;
  }

  return same / length;
};
const getLocalPossibleMatches = (items, searchText) => {
  const searchTokens = getTokens(searchText);
  if (!searchTokens.length) return [];

  return items
    .map((item) => {
      const itemText = `${item.code || ""} ${item.name || ""} ${item.category || ""} ${item.sheetName || ""}`;
      const itemTokens = new Set(getTokens(itemText));

      let score = 0;

      searchTokens.forEach((token) => {
        if (itemTokens.has(token)) score += 10;
        if (normalizeText(itemText).includes(token)) score += 4;
      });

      const nameText = normalizeText(item.name);
      const searchClean = normalizeText(searchText);

      if (nameText && searchClean && nameText.includes(searchClean)) score += 30;
      if (nameText && searchClean && searchClean.includes(nameText)) score += 25;

      return { item, score };
    })
    .filter((entry) => entry.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 6)
    .map((entry) => entry.item);
};

export default function InventoryAiHelper({
  styles,
  items = [],
  inventoryReady,
  currentStationSubmitted,
  inventoryStation,
  onUseItem,
}) {
  const [previewUrl, setPreviewUrl] = useState("");
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  const [result, setResult] = useState(null);
  const [visualMatches, setVisualMatches] = useState([]);
  const [manualSearch, setManualSearch] = useState("");

  const candidates = useMemo(
    () =>
      items.map((item, index) => ({
        index,
        code: item.code || "",
        name: item.name || "",
        category: item.category || "",
        sheetName: item.sheetName || "",
      })),
    [items]
  );

      const localMatchEntries = useMemo(() => {
    const searchText = result
      ? [
          result.visualName,
          result.visualDescription,
          result.equipmentCategory,
          ...(result.likelySearchTerms || []),
        ]
          .filter(Boolean)
          .join(" ")
      : manualSearch;

    const searchTokens = getTokens(searchText);

    if (!searchTokens.length) return [];

    return items
      .map((item) => {
        const itemText = `${item.code || ""} ${item.name || ""} ${item.category || ""} ${item.sheetName || ""}`;
        const itemClean = normalizeText(itemText);
        const itemNameClean = normalizeText(item.name);
        const searchClean = normalizeText(searchText);
        const itemTokens = new Set(getTokens(itemText));

        let score = 0;

        searchTokens.forEach((token) => {
          if (itemTokens.has(token)) score += 12;
          if (itemClean.includes(token)) score += 5;
        });

        if (searchClean && itemNameClean === searchClean) score += 80;
        if (searchClean && itemNameClean.includes(searchClean)) score += 45;
        if (
          searchClean &&
          searchClean.includes(itemNameClean) &&
          itemNameClean.length > 8
        ) {
          score += 35;
        }

        return {
          item,
          score,
        };
      })
      .filter((entry) => entry.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, 8);
  }, [items, result, manualSearch]);

  const matchedItem = useMemo(() => {
  const visualBest = visualMatches[0];

  if (visualBest?.item && Number(visualBest.similarity || 0) >= 0.82) {
    return visualBest.item;
  }

  const best = localMatchEntries[0];

  if (!best) return null;

  return best.score >= 25 ? best.item : null;
}, [localMatchEntries, visualMatches]);

  const possibleMatches = useMemo(() => {
  const seen = new Set();

  return [
    ...(matchedItem ? [matchedItem] : []),
    ...visualMatches.map((entry) => entry.item).filter(Boolean),
    ...localMatchEntries.map((entry) => entry.item),
  ]
    .filter((item) => {
      const key = `${item.code || ""}|${item.name || ""}|${item.sheetName || ""}`;

      if (seen.has(key)) return false;

      seen.add(key);
      return true;
    })
    .slice(0, 6);
}, [localMatchEntries, matchedItem, visualMatches]);
  const runVisualMasterListSearch = async (imageDataUrl) => {
  if (!imageDataUrl) return [];

  const itemsWithImages = items
    .map((item, index) => ({
      item,
      index,
      imageUrl: getComparableImageUrl(item.image),
    }))
    .filter((entry) => entry.imageUrl);

  if (!itemsWithImages.length) {
    setVisualMatches([]);
    return [];
  }

  try {
    const sourceHash = await getImageHash(imageDataUrl);
    const matches = [];

    for (let index = 0; index < itemsWithImages.length; index += 1) {
      if (index > 0 && index % 20 === 0) {
        await waitForBrowser();
      }

      const entry = itemsWithImages[index];

      try {
        const candidateHash = await getImageHash(entry.imageUrl);
        const similarity = getHashSimilarity(sourceHash, candidateHash);

        if (similarity >= 0.62) {
          matches.push({
            item: entry.item,
            similarity,
          });
        }
      } catch {
        // Some external image links cannot be compared because the browser blocks pixel access.
      }
    }

    const sortedMatches = matches
      .sort((a, b) => b.similarity - a.similarity)
      .slice(0, 6);

    setVisualMatches(sortedMatches);

    return sortedMatches;
  } catch {
    setVisualMatches([]);
    return [];
  }
};
  const handlePhotoSelected = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    if (!items.length) {
      setMessage("Upload or refresh the master inventory list first.");
      event.target.value = "";
      return;
    }

    setBusy(true);
setResult(null);
setVisualMatches([]);
setMessage("Reading equipment picture...");
setPreviewUrl(URL.createObjectURL(file));

    let imageDataUrl = "";

try {
  imageDataUrl = await resizeImageFileToDataUrl(file, {
        maxWidth: 1280,
        maxHeight: 1280,
        quality: 0.72,
      });

      const response = await fetch("/api/identify-equipment", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
  imageDataUrl,
}),
      });

      const responseText = await response.text();

      let data = {};

      try {
        setMessage("AI identified the equipment. Master list search completed.");
      } catch {
        throw new Error(
          responseText
            ? responseText.slice(0, 180)
            : "Server returned a non-JSON response."
        );
      }

      if (!response.ok) {
        throw new Error(data?.error || responseText || "Could not identify equipment.");
      }

      setResult(data.result || null);
      setMessage("AI helper finished. Confirm the match before counting.");
    } catch (error) {
  const errorText = error?.message || "Could not identify equipment.";

  const isAiUnavailable =
    errorText.toLowerCase().includes("quota") ||
    errorText.toLowerCase().includes("billing") ||
    errorText.toLowerCase().includes("exceeded") ||
    errorText.toLowerCase().includes("rate limit") ||
    errorText.toLowerCase().includes("temporarily unavailable");

  setResult(null);
  setVisualMatches([]);
  setManualSearch("");

  setMessage(
    isAiUnavailable
      ? "AI is unavailable. Type equipment name below to search the uploaded master list."
      : `${errorText} Type equipment name below to search the uploaded master list.`
  );
}
} finally {
      setBusy(false);
      event.target.value = "";
    }
  };

  const handleUseItem = (item) => {
    if (!inventoryReady) {
      setMessage("Choose ship, station and user before counting.");
      return;
    }

    if (currentStationSubmitted) {
      setMessage(
        `${inventoryStation || "This station"} has already submitted count. Reset inventory before editing.`
      );
      return;
    }

    onUseItem(item);
  };

  return (
    <section style={styles.card}>
      <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 14 }}>
        <div>
          <h2 style={styles.productTitle}>🤖 AI Equipment Helper</h2>
          <p style={{ ...styles.emptyText, margin: 0 }}>
            Take a picture. AI will identify the equipment and search the uploaded master list.
          </p>
        </div>
      </div>

      <label style={styles.label}>Take / upload equipment picture</label>
      <input
        type="file"
        accept="image/*"
        capture="environment"
        onChange={handlePhotoSelected}
        style={styles.fileInput}
      />

      {previewUrl && (
        <div style={styles.infoBox}>
          <img
            src={previewUrl}
            alt="Equipment preview"
            style={{
              width: "100%",
              maxHeight: 280,
              objectFit: "contain",
              borderRadius: 12,
              background: "#f2f2f2",
            }}
          />
        </div>
      )}

      {message && <div style={busy ? styles.warningText : styles.infoBox}>{message}</div>}

      {result && (
        <div style={styles.infoBox}>
          <div>
            AI identified: <strong>{result.visualName || "Unknown equipment"}</strong>
          </div>

          {result.visualDescription && (
            <div>Description: {result.visualDescription}</div>
          )}

          <div>
            Match status: <strong>{result.matchStatus || "not_found"}</strong>
          </div>

          <div>
            Confidence: <strong>{Math.round(Number(result.confidence || 0) * 100)}%</strong>
          </div>

          {result.notes && <div>Notes: {result.notes}</div>}
        </div>
      )}

      {matchedItem && (
        <div style={{ ...styles.equipmentCard, ...styles.countedCard }}>
          <div style={styles.recipeName}>Best Match</div>
          <div style={styles.recipeName}>{matchedItem.name}</div>
          <div style={styles.recipeMeta}>Code: {matchedItem.code || "N/A"}</div>
          <div style={styles.recipeMeta}>Sheet: {matchedItem.sheetName || "N/A"}</div>
          <div style={styles.recipeMeta}>Category: {matchedItem.category || "N/A"}</div>

          <button
            style={styles.primaryButton}
            onClick={() => handleUseItem(matchedItem)}
          >
            ✅ Use This Equipment
          </button>
        </div>
      )}

      {result && !matchedItem && (
        <div style={styles.warningText}>
          Equipment was not found as an exact match in the master list. Review possible uploaded-list matches below, or use the AI name to search manually.
        </div>
      )}

      {possibleMatches.length > 0 && (
        <>
          <h3 style={styles.sectionTitle}>Possible Matches</h3>

          <div style={styles.equipmentGrid}>
            {possibleMatches.map((item, index) => (
              <div key={`${item.code}-${item.name}-${index}`} style={styles.equipmentCard}>
                <div style={styles.recipeName}>{item.name}</div>
                <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                <div style={styles.recipeMeta}>Sheet: {item.sheetName || "N/A"}</div>
                <div style={styles.recipeMeta}>Category: {item.category || "N/A"}</div>
                {visualMatches.find((entry) => entry.item === item) && (
  <div style={styles.statusGood}>
    Picture similarity:{" "}
    {Math.round(
      Number(
        visualMatches.find((entry) => entry.item === item)?.similarity || 0
      ) * 100
    )}
    %
  </div>
)}

                <button
                  style={styles.backButton}
                  onClick={() => handleUseItem(item)}
                >
                  Use This Match
                </button>
              </div>
            ))}
          </div>
        </>
      )}
    </section>
  );
}

"use client";

import React, { useEffect, useMemo, useState } from "react";

const BUCKET_NAME = "temperature-photos";

const todayDateKey = () => new Date().toISOString().slice(0, 10);

const safeFileName = (value) =>
  String(value || "photo")
    .replace(/[^a-z0-9._-]+/gi, "-")
    .replace(/-+/g, "-")
    .slice(0, 80);

const fileToDataUrl = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Could not read file."));
    reader.readAsDataURL(file);
  });
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

const toNumberOrNull = (value) => {
  const number = Number(String(value ?? "").replace(/[^0-9.\-]/g, ""));
  return Number.isFinite(number) ? number : null;
};

export default function TemperatureCheckModule({
  styles,
  supabase,
  userShip,
  userEmail,
  onBack,
  logUsageEvent = () => {},
}) {
  const [selectedFile, setSelectedFile] = useState(null);
  const [previewUrl, setPreviewUrl] = useState("");
  const [photoDataUrl, setPhotoDataUrl] = useState("");
  const [analyzing, setAnalyzing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [message, setMessage] = useState("");

  const [station, setStation] = useState("");
  const [userName, setUserName] = useState("");
  const [productName, setProductName] = useState("");
  const [foodCategory, setFoodCategory] = useState("");
  const [temperatureValue, setTemperatureValue] = useState("");
  const [temperatureUnit, setTemperatureUnit] = useState("F");
  const [temperatureText, setTemperatureText] = useState("");
  const [readerVisible, setReaderVisible] = useState(false);
  const [isChickenOrPoultry, setIsChickenOrPoultry] = useState(false);
  const [confidence, setConfidence] = useState(0);
  const [notes, setNotes] = useState("");

  const [dateFilter, setDateFilter] = useState(todayDateKey());
  const [savedChecks, setSavedChecks] = useState([]);
  const [loadingSaved, setLoadingSaved] = useState(false);

  const canSave = useMemo(
    () =>
      Boolean(
        supabase &&
          selectedFile &&
          userShip &&
          productName.trim() &&
          String(temperatureValue).trim()
      ),
    [supabase, selectedFile, userShip, productName, temperatureValue]
  );

  const resetCurrentPhoto = () => {
    setSelectedFile(null);
    setPreviewUrl("");
    setPhotoDataUrl("");
    setProductName("");
    setFoodCategory("");
    setTemperatureValue("");
    setTemperatureText("");
    setReaderVisible(false);
    setIsChickenOrPoultry(false);
    setConfidence(0);
    setNotes("");
  };

  const handlePhotoSelected = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    setSelectedFile(file);
    setPreviewUrl(URL.createObjectURL(file));
    setMessage("Photo selected. Analyzing...");
    setAnalyzing(true);

    try {
      const dataUrl = await resizeImageFileToDataUrl(file, {
  maxWidth: 1280,
  maxHeight: 1280,
  quality: 0.72,
});
      setPhotoDataUrl(dataUrl);

      const response = await fetch("/api/analyze-temperature", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          imageDataUrl: dataUrl,
        }),
      });

      const responseText = await response.text();

let data = {};

try {
  data = responseText ? JSON.parse(responseText) : {};
} catch {
  throw new Error(
    responseText
      ? responseText.slice(0, 180)
      : "Server returned a non-JSON response."
  );
}

if (!response.ok) {
  throw new Error(data?.error || responseText || "Could not analyze photo.");
}

      const result = data.result || {};

      setProductName(result.productName || "");
      setFoodCategory(result.foodCategory || "");
      setTemperatureText(result.temperatureText || "");
      setTemperatureValue(
        result.temperatureValue === null || result.temperatureValue === undefined
          ? ""
          : String(result.temperatureValue)
      );
      setTemperatureUnit(result.temperatureUnit || "F");
      setReaderVisible(Boolean(result.readerVisible));
      setIsChickenOrPoultry(Boolean(result.isChickenOrPoultry));
      setConfidence(Number(result.confidence || 0));
      setNotes(result.notes || "");

      setMessage("AI analysis completed. Please verify before saving.");
    } catch (error) {
      setMessage(error?.message || "Could not analyze photo. Enter product and temperature manually.");
    } finally {
      setAnalyzing(false);
      event.target.value = "";
    }
  };

  const loadSavedChecks = async () => {
    if (!supabase) {
      setMessage("Supabase is not connected.");
      return;
    }

    setLoadingSaved(true);

    try {
      const start = new Date(`${dateFilter}T00:00:00`);
      const end = new Date(start.getTime() + 24 * 60 * 60 * 1000);

      let query = supabase
        .from("temperature_checks")
        .select("*")
        .gte("taken_at", start.toISOString())
        .lt("taken_at", end.toISOString())
        .order("taken_at", { ascending: false })
        .limit(100);

      if (userShip) {
        query = query.eq("ship", userShip);
      }

      const { data, error } = await query;

      if (error) {
        throw error;
      }

      setSavedChecks(data || []);
    } catch (error) {
      setMessage(error?.message || "Could not load saved temperature checks.");
    } finally {
      setLoadingSaved(false);
    }
  };

  useEffect(() => {
    loadSavedChecks();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dateFilter, userShip]);

  const saveTemperatureCheck = async () => {
    if (!supabase) {
      window.alert("Supabase is not connected.");
      return;
    }

    if (!selectedFile) {
      window.alert("Take or upload a photo first.");
      return;
    }

    if (!userShip) {
      window.alert("Choose ship first.");
      return;
    }

    if (!productName.trim()) {
      window.alert("Confirm product name before saving.");
      return;
    }

    const tempNumber = toNumberOrNull(temperatureValue);

    if (tempNumber === null) {
      window.alert("Confirm temperature value before saving.");
      return;
    }

    const confirmed = window.confirm(
      `Save temperature check?\n\nProduct: ${productName}\nTemperature: ${tempNumber} ${temperatureUnit}\n\nPlease confirm the AI reading is correct.`
    );

    if (!confirmed) return;

    setSaving(true);
    setMessage("Saving temperature check...");

    try {
      const dateFolder = todayDateKey();
      const extension =
        selectedFile.name.split(".").pop()?.toLowerCase() ||
        (selectedFile.type.includes("png") ? "png" : "jpg");

      const fileName = `${Date.now()}-${safeFileName(productName)}.${extension}`;
      const imagePath = `${dateFolder}/${fileName}`;

      const { error: uploadError } = await supabase.storage
        .from(BUCKET_NAME)
        .upload(imagePath, selectedFile, {
          cacheControl: "3600",
          upsert: false,
          contentType: selectedFile.type || "image/jpeg",
        });

      if (uploadError) {
        throw uploadError;
      }

      const publicUrlResult = supabase.storage
        .from(BUCKET_NAME)
        .getPublicUrl(imagePath);

      const imageUrl = publicUrlResult?.data?.publicUrl || "";

      const payload = {
        ship: userShip || "",
        department: "culinary",
        station: station || "",
        user_email: userEmail || "",
        user_name: userName || "",
        product_name: productName.trim(),
        food_category: foodCategory || "",
        temperature_value: tempNumber,
        temperature_unit: temperatureUnit || "F",
        temperature_text: temperatureText || `${tempNumber} ${temperatureUnit || "F"}`,
        reader_visible: readerVisible,
        is_chicken_or_poultry: isChickenOrPoultry,
        confidence: Number(confidence || 0),
        notes: notes || "",
        image_bucket: BUCKET_NAME,
        image_path: imagePath,
        image_url: imageUrl,
        taken_at: new Date().toISOString(),
      };

      const { error: insertError } = await supabase
        .from("temperature_checks")
        .insert(payload);

      if (insertError) {
        throw insertError;
      }

      setMessage(`Saved ${productName} / ${tempNumber} ${temperatureUnit}.`);

      logUsageEvent("temperature_check_saved", {
        module: "temperature_check",
        ship: userShip,
        productName,
        temperatureValue: tempNumber,
        temperatureUnit,
        isChickenOrPoultry,
      });

      resetCurrentPhoto();
      await loadSavedChecks();
    } catch (error) {
      setMessage(error?.message || "Could not save temperature check.");
      window.alert(error?.message || "Could not save temperature check.");
    } finally {
      setSaving(false);
    }
  };

    return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Modules
          </button>
          <div style={styles.shipBadge}>🌡️ {userShip || "Ship"}</div>
        </div>
      </header>

      <section style={styles.card}>
        <h2 style={styles.productTitle}>🌡️ Take Temperature Picture</h2>
        <p style={styles.emptyText}>
          Take a picture, confirm product name and temperature, then save.
        </p>

        <label style={styles.label}>Take picture / upload picture</label>
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
              alt="Temperature check preview"
              style={{
                width: "100%",
                maxHeight: 420,
                objectFit: "contain",
                borderRadius: 14,
                background: "#f2f2f2",
              }}
            />
          </div>
        )}

        {analyzing && (
          <div style={styles.warningText}>
            Reading picture...
          </div>
        )}

        {message && (
          <div style={styles.infoBox}>
            {message}
          </div>
        )}

        <div style={styles.formRow}>
          <div>
            <label style={styles.label}>Product name</label>
            <input
              value={productName}
              onChange={(event) => setProductName(event.target.value)}
              placeholder="Example: Chicken Breast"
              style={styles.searchInput}
            />
          </div>

          <div>
            <label style={styles.label}>Temperature</label>
            <input
              type="number"
              value={temperatureValue}
              onChange={(event) => setTemperatureValue(event.target.value)}
              placeholder="Example: 38.5"
              style={styles.searchInput}
            />
          </div>
        </div>

        <label style={styles.label}>Unit</label>
        <select
          value={temperatureUnit}
          onChange={(event) => setTemperatureUnit(event.target.value)}
          style={styles.select}
        >
          <option value="F">F</option>
          <option value="C">C</option>
          <option value="unknown">Unknown</option>
        </select>

        <button
          style={styles.primaryButton}
          onClick={saveTemperatureCheck}
          disabled={!canSave || analyzing || saving}
        >
          {saving ? "Saving..." : "💾 Save Picture"}
        </button>

        {!canSave && (
          <p style={styles.emptyText}>
            Picture, product name, and temperature are required before saving.
          </p>
        )}
      </section>
    </main>
  );
}

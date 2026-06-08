"use client";

import React, { useEffect, useMemo, useState } from "react";
import { cleanText, formatQty, getImageUrl } from "../../lib/appHelpers";
import {
  EQUIPMENT_PICTURE_BUCKET,
  getMasterInventoryScope,
} from "../../constants/inventoryConfig";
import { makeStorageSafePart } from "../../lib/inventoryImageHelpers";

const SHIPS = ["SC", "VL", "BRL", "RL"];

const getMonthKey = (date = new Date()) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");

  return `${year}-${month}`;
};

const getMonthLabel = (monthKey = getMonthKey()) => {
  const [year, month] = String(monthKey || "").split("-");
  const date = new Date(Number(year), Number(month) - 1, 1);

  if (Number.isNaN(date.getTime())) return monthKey;

  return date.toLocaleDateString(undefined, {
    month: "long",
    year: "numeric",
  });
};

const getFileExtension = (file) => {
  const fileName = String(file?.name || "").toLowerCase();

  if (fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")) return "jpg";
  if (fileName.endsWith(".webp")) return "webp";
  if (fileName.endsWith(".gif")) return "gif";

  return "png";
};

const makeBreakageItemKey = (item) =>
  cleanText(
    [
      item?.equipmentDepartment || "equipment",
      item?.sheetName || "",
      item?.category || "",
      item?.code || "",
      item?.name || "",
      item?.sourceRow || "",
    ].join("|")
  );

const normalizeBreakageRow = (row) => ({
  id: row.id || "",
  ship: row.ship || "",
  department: row.department || "",
  monthKey: row.month_key || "",
  itemKey: row.item_key || "",
  code: row.code || "",
  name: row.item_name || "",
  category: row.category || "",
  sheetName: row.sheet_name || "",
  image: row.image || "",
  qty: Number(row.qty || 0),
  notes: row.notes || "",
  reportPhoto: row.report_photo || "",
  userName: row.user_name || "",
  userPosition: row.user_position || "",
  reportedAt: row.reported_at || row.created_at || "",
  confirmedAt: row.reported_at
    ? new Date(row.reported_at).toLocaleString()
    : "",
});

export default function BreakageReportModule({
  styles,
  supabase,
  userShip,
  userEmail,
  equipmentDepartment,
  activeEquipmentDepartmentLabel,
  getShipDisplayName,
  logUsageEvent,
  onBack,
}) {
  const monthKey = getMonthKey();
  const monthLabel = getMonthLabel(monthKey);
  const departmentKey = equipmentDepartment || "culinary";
  const departmentLabel = activeEquipmentDepartmentLabel || "Equipment";

  const [entryShip, setEntryShip] = useState(userShip || "");
  const [reportedBy, setReportedBy] = useState("");
  const [reportedPosition, setReportedPosition] = useState("");

  const [masterItems, setMasterItems] = useState([]);
  const [masterLoading, setMasterLoading] = useState(false);
  const [message, setMessage] = useState("");

  const [reportRows, setReportRows] = useState([]);
  const [reportLoading, setReportLoading] = useState(false);

  const [search, setSearch] = useState("");

  const [currentItem, setCurrentItem] = useState(null);
  const [breakageQty, setBreakageQty] = useState("");
  const [breakageNotes, setBreakageNotes] = useState("");
  const [breakagePhotoFile, setBreakagePhotoFile] = useState(null);
  const [breakagePhotoInputKey, setBreakagePhotoInputKey] = useState(0);
  const [savingBreakage, setSavingBreakage] = useState(false);

  const [extraCode, setExtraCode] = useState("");
  const [extraName, setExtraName] = useState("");
  const [extraQty, setExtraQty] = useState("");
  const [extraNotes, setExtraNotes] = useState("");
  const [extraPhotoFile, setExtraPhotoFile] = useState(null);
  const [extraPhotoInputKey, setExtraPhotoInputKey] = useState(0);
  const [savingExtra, setSavingExtra] = useState(false);

  const getShipName =
    typeof getShipDisplayName === "function"
      ? getShipDisplayName
      : (ship) => ship;

  const effectiveReportedBy = [reportedBy.trim(), reportedPosition.trim()]
    .filter(Boolean)
    .join(" - ");

  const loadMasterItems = async () => {
    if (!supabase) {
      setMasterItems([]);
      setMessage("Supabase is not connected. Master list cannot load.");
      return;
    }

    setMasterLoading(true);
    setMessage(`Loading ${departmentLabel} equipment master list...`);

    try {
      const scope = getMasterInventoryScope(departmentKey);

      const { data, error } = await supabase
        .from("inventory_master_items")
        .select(
          "id, ship, item_key, code, item_name, category, sheet_name, image, source_row, sort_order"
        )
        .eq("ship", scope)
        .order("sort_order", { ascending: true })
        .limit(5000);

      if (error) throw error;

      const rows = (data || [])
        .map((row, index) => {
          const item = {
            id: row.id || `${scope}-${index}`,
            equipmentDepartment: departmentKey,
            departmentLabel,
            itemKey: row.item_key || "",
            code: row.code || "",
            name: row.item_name || "",
            category: row.category || "",
            sheetName: row.sheet_name || "",
            image: row.image || "",
            sourceRow: row.source_row || "",
            sortOrder: Number(row.sort_order ?? index),
          };

          return {
            ...item,
            itemKey: item.itemKey || makeBreakageItemKey(item),
          };
        })
        .filter((item) => item.name || item.code);

      setMasterItems(rows);
      setMessage(
        rows.length
          ? `${departmentLabel} equipment master list loaded. ${rows.length} item(s).`
          : `No ${departmentLabel} master list found. Upload it first from Equipment Master List.`
      );
    } catch (error) {
      setMasterItems([]);
      setMessage(
        error?.message ||
          `Could not load ${departmentLabel} equipment master list.`
      );
    } finally {
      setMasterLoading(false);
    }
  };

  const loadReportRows = async () => {
    if (!supabase) {
      setReportRows([]);
      return;
    }

    setReportLoading(true);

    try {
      const { data, error } = await supabase
        .from("equipment_breakage_reports")
        .select("*")
        .eq("department", departmentKey)
        .eq("month_key", monthKey)
        .order("reported_at", { ascending: false })
        .limit(5000);

      if (error) throw error;

      setReportRows((data || []).map(normalizeBreakageRow));
    } catch {
      setReportRows([]);
    } finally {
      setReportLoading(false);
    }
  };

  useEffect(() => {
    setEntryShip((current) => current || userShip || "");
  }, [userShip]);

  useEffect(() => {
    loadMasterItems();
    loadReportRows();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [departmentKey]);

  const visibleItems = useMemo(() => {
    const term = search.toLowerCase().trim();

    if (!term) return masterItems;

    return masterItems.filter((item) =>
      [
        item.departmentLabel,
        item.code,
        item.name,
        item.category,
        item.sheetName,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term)
    );
  }, [masterItems, search]);

  const getItemMonthlyCount = (item, ship = entryShip) => {
    const itemKey = item?.itemKey || makeBreakageItemKey(item);

    if (!itemKey || !ship) return 0;

    return reportRows
      .filter((row) => row.ship === ship)
      .filter((row) => row.department === departmentKey)
      .filter((row) => row.monthKey === monthKey)
      .filter((row) => row.itemKey === itemKey)
      .reduce((sum, row) => sum + Number(row.qty || 0), 0);
  };

  const uploadBreakagePhoto = async ({ file, item, ship }) => {
    if (!supabase || !file) return "";

    const extension = getFileExtension(file);
    const contentType = file.type || `image/${extension}`;
    const itemPart = makeStorageSafePart(item?.code || item?.name || "breakage");

    const path =
      "breakage/" +
      makeStorageSafePart(ship || "ship") +
      "/" +
      makeStorageSafePart(departmentKey) +
      "/" +
      Date.now() +
      "-" +
      itemPart +
      "." +
      extension;

    const { error } = await supabase.storage
      .from(EQUIPMENT_PICTURE_BUCKET)
      .upload(path, file, {
        contentType,
        upsert: true,
        cacheControl: "31536000",
      });

    if (error) throw error;

    const { data } = supabase.storage
      .from(EQUIPMENT_PICTURE_BUCKET)
      .getPublicUrl(path);

    return data?.publicUrl || "";
  };

  const validateBaseFields = () => {
    if (!supabase) {
      window.alert("Supabase is not connected. Breakage cannot be saved.");
      return false;
    }

    if (!entryShip) {
      window.alert("Choose ship before saving breakage.");
      return false;
    }

    if (!reportedBy.trim()) {
      window.alert("Enter reported by name before saving breakage.");
      return false;
    }

    return true;
  };

  const saveBreakageRecord = async ({ item, qty, notes, photoFile }) => {
    const safeQty = Number(qty || 0);

    if (String(qty).trim() === "" || Number.isNaN(safeQty) || safeQty <= 0) {
      window.alert("Enter broken quantity greater than 0.");
      return false;
    }

    const photoUrl = photoFile
      ? await uploadBreakagePhoto({
          file: photoFile,
          item,
          ship: entryShip,
        })
      : "";

    const now = new Date().toISOString();

    const payload = {
      ship: entryShip,
      department: departmentKey,
      month_key: monthKey,

      user_name: effectiveReportedBy,
      user_position: reportedPosition || "",

      item_key: item.itemKey || makeBreakageItemKey(item),
      code: item.code || "",
      item_name: item.name || "",
      category: item.category || "",
      sheet_name: item.sheetName || "",
      image: item.image || photoUrl || "",

      qty: safeQty,
      notes: notes || "",

      reported_at: now,
      updated_at: now,
    };

    if (photoUrl) {
      payload.report_photo = photoUrl;
    }

    const { data, error } = await supabase
      .from("equipment_breakage_reports")
      .insert(payload)
      .select("*")
      .single();

    if (error) throw error;

    const savedRow = normalizeBreakageRow(data);
    setReportRows((current) => [savedRow, ...current]);

    if (typeof logUsageEvent === "function") {
      logUsageEvent("equipment_breakage_saved", {
        module: `equipment_${departmentKey}_breakage_report`,
        ship: entryShip,
        equipmentDepartment: departmentKey,
        userEmail,
        userName: effectiveReportedBy,
        itemName: item.name || "",
        code: item.code || "",
        qty: safeQty,
        hasPhoto: Boolean(photoUrl),
        monthKey,
      });
    }

    return true;
  };

  const confirmMasterItemBreakage = async () => {
    if (savingBreakage) return;
    if (!currentItem) return;
    if (!validateBaseFields()) return;

    setSavingBreakage(true);
    setMessage("Saving breakage...");

    try {
      const saved = await saveBreakageRecord({
        item: currentItem,
        qty: breakageQty,
        notes: breakageNotes,
        photoFile: breakagePhotoFile,
      });

      if (!saved) return;

      setMessage(`Saved breakage: ${currentItem.name}.`);
      setCurrentItem(null);
      setBreakageQty("");
      setBreakageNotes("");
      setBreakagePhotoFile(null);
      setBreakagePhotoInputKey((value) => value + 1);
    } catch (error) {
      const text = error?.message || "Could not save breakage.";
      setMessage(text);
      window.alert(text);
    } finally {
      setSavingBreakage(false);
    }
  };

  const saveExtraBrokenItem = async () => {
    if (savingExtra) return;
    if (!validateBaseFields()) return;

    const name = String(extraName || "").replace(/\s+/g, " ").trim();

    if (!name) {
      window.alert("Enter item name.");
      return;
    }

    setSavingExtra(true);
    setMessage("Saving broken item not in master list...");

    try {
      const item = {
        equipmentDepartment: departmentKey,
        departmentLabel,
        itemKey: cleanText(
          `${departmentKey}|EXTRA|${extraCode || "NO-CODE"}|${name}|${Date.now()}`
        ),
        code: extraCode || "EXTRA",
        name,
        category: "Extra item - not in master list",
        sheetName: "Extra Item",
        image: "",
      };

      const saved = await saveBreakageRecord({
        item,
        qty: extraQty,
        notes: extraNotes,
        photoFile: extraPhotoFile,
      });

      if (!saved) return;

      setMessage(`Saved broken item not in master list: ${name}.`);
      setExtraCode("");
      setExtraName("");
      setExtraQty("");
      setExtraNotes("");
      setExtraPhotoFile(null);
      setExtraPhotoInputKey((value) => value + 1);
    } catch (error) {
      const text = error?.message || "Could not save broken extra item.";
      setMessage(text);
      window.alert(text);
    } finally {
      setSavingExtra(false);
    }
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
          <button style={styles.backButton} onClick={onBack}>
            ← {departmentLabel} Options
          </button>

          <div style={styles.shipBadge}>
            🧾 {departmentLabel} Breakage Report
          </div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🧾 Report Breakage</h2>

          <label style={styles.label}>Ship</label>
          <select
            value={entryShip}
            onChange={(event) => setEntryShip(event.target.value)}
            style={styles.select}
          >
            <option value="">Choose ship</option>
            {SHIPS.map((ship) => (
              <option key={ship} value={ship}>
                {getShipName(ship)}
              </option>
            ))}
          </select>

          <label style={styles.label}>Reported by</label>
          <input
            value={reportedBy}
            onChange={(event) => setReportedBy(event.target.value)}
            placeholder="Type your name..."
            style={styles.searchInput}
          />

          <label style={styles.label}>Position optional</label>
          <input
            value={reportedPosition}
            onChange={(event) => setReportedPosition(event.target.value)}
            placeholder="Type your position..."
            style={styles.searchInput}
          />

          <button
            type="button"
            style={styles.backButton}
            onClick={() => {
              loadMasterItems();
              loadReportRows();
            }}
            disabled={masterLoading || reportLoading}
          >
            {masterLoading || reportLoading
              ? "Refreshing..."
              : "🔄 Refresh Master List"}
          </button>

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.infoBox}>
            <div>
              🧭 Department: <strong>{departmentLabel}</strong>
            </div>
            <div>
              🚢 Ship: <strong>{entryShip || "Not selected"}</strong>
            </div>
            <div>
              👤 User: <strong>{effectiveReportedBy || "Not selected"}</strong>
            </div>
            <div>
              📋 Master list items: <strong>{masterItems.length}</strong>
            </div>
            <div>
              📅 Report month: <strong>{monthLabel}</strong>
            </div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>➕ Item Not In Master List</h2>

          <p style={styles.emptyText}>
            Use this only when broken equipment is not found in the{" "}
            {departmentLabel} master list.
          </p>

          <label style={styles.label}>Code optional</label>
          <input
            value={extraCode}
            onChange={(event) => setExtraCode(event.target.value)}
            placeholder="Enter code..."
            style={styles.searchInput}
          />

          <label style={styles.label}>Item name</label>
          <input
            value={extraName}
            onChange={(event) => setExtraName(event.target.value)}
            placeholder="Enter item name..."
            style={styles.searchInput}
          />

          <label style={styles.label}>Broken quantity</label>
          <input
            type="number"
            min="1"
            step="1"
            value={extraQty}
            onChange={(event) => setExtraQty(event.target.value)}
            placeholder="Enter broken quantity..."
            style={styles.searchInput}
          />

          <label style={styles.label}>Notes optional</label>
          <input
            value={extraNotes}
            onChange={(event) => setExtraNotes(event.target.value)}
            placeholder="Example: cracked, chipped, broken..."
            style={styles.searchInput}
          />

          <label style={styles.label}>Take / upload picture</label>
          <input
            key={extraPhotoInputKey}
            type="file"
            accept="image/*"
            capture="environment"
            onChange={(event) =>
              setExtraPhotoFile(event.target.files?.[0] || null)
            }
            style={styles.fileInput}
          />

          {extraPhotoFile && (
            <div style={styles.statusNeutral}>
              Picture selected: {extraPhotoFile.name}
            </div>
          )}

          <button
            type="button"
            style={styles.primaryButton}
            onClick={saveExtraBrokenItem}
            disabled={savingExtra}
          >
            {savingExtra ? "Saving..." : "Save Broken Item"}
          </button>
        </div>
      </section>

      <section style={styles.card}>
        <h2 style={styles.productTitle}>📋 {departmentLabel} Equipment Master List</h2>

        <label style={styles.label}>Search equipment</label>
        <input
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          placeholder="Search code, name, category, sheet..."
          style={styles.searchInput}
        />

        {visibleItems.length === 0 && (
          <p style={styles.emptyText}>
            No equipment found. Upload the {departmentLabel} master list first
            from Equipment Master List.
          </p>
        )}

        <div style={styles.equipmentGrid}>
          {visibleItems.map((item) => {
            const itemCount = getItemMonthlyCount(item, entryShip);

            return (
              <button
                key={`${item.equipmentDepartment}-${item.itemKey}-${item.id}`}
                type="button"
                style={styles.inventoryItemCard}
                onClick={() => {
                  setCurrentItem(item);
                  setBreakageQty("");
                  setBreakageNotes("");
                  setBreakagePhotoFile(null);
                  setBreakagePhotoInputKey((value) => value + 1);
                }}
              >
                {item.image ? (
                  <div style={styles.inventoryImageFrame}>
                    <img
                      src={getImageUrl(item.image, "w360")}
                      alt={item.name}
                      loading="lazy"
                      decoding="async"
                      style={styles.inventoryCardImage}
                      onError={(event) => {
                        event.currentTarget.style.display = "none";
                        const fallback =
                          event.currentTarget.nextElementSibling;
                        if (fallback) fallback.style.display = "flex";
                      }}
                    />

                    <div
                      style={{
                        ...styles.inventoryNoImage,
                        display: "none",
                      }}
                    >
                      No image
                    </div>
                  </div>
                ) : (
                  <div style={styles.inventoryNoImage}>No image</div>
                )}

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

                {entryShip && itemCount > 0 && (
                  <div style={styles.statusWarning}>
                    This month broken: {formatQty(itemCount)}
                  </div>
                )}
              </button>
            );
          })}
        </div>
      </section>

      {currentItem && (
        <div
          style={styles.modalBackdrop}
          onClick={() => {
            setCurrentItem(null);
            setBreakageQty("");
            setBreakageNotes("");
            setBreakagePhotoFile(null);
          }}
        >
          <div
            style={styles.modalCard}
            onClick={(event) => event.stopPropagation()}
          >
            <button
              type="button"
              style={styles.closeButton}
              onClick={() => {
                setCurrentItem(null);
                setBreakageQty("");
                setBreakageNotes("");
                setBreakagePhotoFile(null);
              }}
            >
              ✕
            </button>

            <h2 style={styles.productTitle}>🧾 Report Broken Equipment</h2>

            <div style={styles.grid}>
              <div>
                {currentItem.image ? (
                  <div>
                    <div style={styles.modalPictureFrame}>
                      <img
                        src={getImageUrl(currentItem.image, "w720")}
                        alt={currentItem.name}
                        style={styles.modalPreviewImage}
                        onError={(event) => {
                          event.currentTarget.style.display = "none";
                          const fallback =
                            event.currentTarget.nextElementSibling;
                          if (fallback) fallback.style.display = "flex";
                        }}
                      />

                      <div
                        style={{
                          ...styles.modalNoImage,
                          display: "none",
                        }}
                      >
                        Picture could not be loaded
                      </div>
                    </div>

                    <a
                      href={currentItem.image}
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
                <h3 style={{ marginTop: 0 }}>{currentItem.name}</h3>

                <p>
                  <strong>Ship:</strong> {entryShip || "Not selected"}
                </p>
                <p>
                  <strong>Department:</strong> {departmentLabel}
                </p>
                <p>
                  <strong>Code:</strong> {currentItem.code || "N/A"}
                </p>
                <p>
                  <strong>Sheet:</strong> {currentItem.sheetName || "N/A"}
                </p>
                <p>
                  <strong>Category:</strong> {currentItem.category || "N/A"}
                </p>

                {entryShip && (
                  <div style={styles.infoBox}>
                    <div>
                      This month broken count:{" "}
                      <strong>
                        {formatQty(getItemMonthlyCount(currentItem, entryShip))}
                      </strong>
                    </div>
                  </div>
                )}

                <label style={styles.label}>Broken quantity</label>
                <input
                  autoFocus
                  type="number"
                  min="1"
                  step="1"
                  value={breakageQty}
                  onChange={(event) => setBreakageQty(event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" && !savingBreakage) {
                      confirmMasterItemBreakage();
                    }
                  }}
                  placeholder="Enter broken quantity..."
                  style={styles.searchInput}
                />

                <label style={styles.label}>Notes optional</label>
                <input
                  value={breakageNotes}
                  onChange={(event) => setBreakageNotes(event.target.value)}
                  placeholder="Example: cracked, chipped, dropped..."
                  style={styles.searchInput}
                />

                <label style={styles.label}>Take / upload picture</label>
                <input
                  key={breakagePhotoInputKey}
                  type="file"
                  accept="image/*"
                  capture="environment"
                  onChange={(event) =>
                    setBreakagePhotoFile(event.target.files?.[0] || null)
                  }
                  style={styles.fileInput}
                />

                {breakagePhotoFile && (
                  <div style={styles.statusNeutral}>
                    Picture selected: {breakagePhotoFile.name}
                  </div>
                )}

                <div style={styles.headerActions}>
                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() => {
                      setCurrentItem(null);
                      setBreakageQty("");
                      setBreakageNotes("");
                      setBreakagePhotoFile(null);
                    }}
                  >
                    Cancel
                  </button>

                  <button
                    type="button"
                    style={styles.primaryButton}
                    onClick={confirmMasterItemBreakage}
                    disabled={savingBreakage}
                  >
                    {savingBreakage ? "Saving..." : "Confirm Breakage"}
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

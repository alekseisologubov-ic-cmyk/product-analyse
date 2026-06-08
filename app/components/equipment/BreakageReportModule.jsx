"use client";

import React, { useEffect, useMemo, useState } from "react";
import { cleanText, formatQty, getImageUrl } from "../../lib/appHelpers";
import {
  EQUIPMENT_PICTURE_BUCKET,
  getMasterInventoryScope,
} from "../../constants/inventoryConfig";
import { makeStorageSafePart } from "../../lib/inventoryImageHelpers";

const SHIPS = ["SC", "VL", "BRL", "RL"];

const DEPARTMENTS = [
  { key: "culinary", label: "Culinary", icon: "👨‍🍳" },
  { key: "bar", label: "Bar", icon: "🍸" },
  { key: "restaurant", label: "Restaurant", icon: "🍽️" },
];

const ALL_DEPARTMENTS = "ALL";

const getMonthKey = (date = new Date()) => {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");

  return `${year}-${month}`;
};

const makeItemKey = (item) =>
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

const getFileExtension = (file) => {
  const fileName = String(file?.name || "").toLowerCase();

  if (fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")) return "jpg";
  if (fileName.endsWith(".webp")) return "webp";
  if (fileName.endsWith(".gif")) return "gif";

  return "png";
};

export default function BreakageReportModule({
  styles,
  supabase,
  userShip,
  userEmail,
  getShipDisplayName,
  logUsageEvent,
  onBack,
}) {
  const [entryShip, setEntryShip] = useState(userShip || "");
  const [reportedBy, setReportedBy] = useState("");
  const [reportedPosition, setReportedPosition] = useState("");

  const [masterItems, setMasterItems] = useState([]);
  const [masterLoading, setMasterLoading] = useState(false);
  const [message, setMessage] = useState("");

  const [search, setSearch] = useState("");
  const [departmentFilter, setDepartmentFilter] = useState(ALL_DEPARTMENTS);

  const [currentItem, setCurrentItem] = useState(null);
  const [breakageQty, setBreakageQty] = useState("");
  const [breakageNotes, setBreakageNotes] = useState("");
  const [breakagePhotoFile, setBreakagePhotoFile] = useState(null);
  const [breakagePhotoInputKey, setBreakagePhotoInputKey] = useState(0);
  const [savingBreakage, setSavingBreakage] = useState(false);

  const [extraDepartment, setExtraDepartment] = useState("culinary");
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
    setMessage("Loading equipment master list...");

    try {
      const scopes = DEPARTMENTS.map((department) =>
        getMasterInventoryScope(department.key)
      );

      const { data, error } = await supabase
        .from("inventory_master_items")
        .select(
          "id, ship, item_key, code, item_name, category, sheet_name, image, source_row, sort_order"
        )
        .in("ship", scopes)
        .order("ship", { ascending: true })
        .order("sort_order", { ascending: true })
        .limit(5000);

      if (error) throw error;

      const departmentByScope = {};
      DEPARTMENTS.forEach((department) => {
        departmentByScope[getMasterInventoryScope(department.key)] =
          department.key;
      });

      const rows = (data || [])
        .map((row, index) => {
          const equipmentDepartment =
            departmentByScope[row.ship] || "culinary";

          const departmentConfig = DEPARTMENTS.find(
            (department) => department.key === equipmentDepartment
          );

          const item = {
            id: row.id || `${row.ship}-${index}`,
            equipmentDepartment,
            departmentLabel: departmentConfig?.label || equipmentDepartment,
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
            itemKey: item.itemKey || makeItemKey(item),
          };
        })
        .filter((item) => item.name || item.code);

      setMasterItems(rows);
      setMessage(
        rows.length
          ? `Equipment master list loaded. ${rows.length} item(s).`
          : "No equipment master list found. Upload master lists from the existing Equipment Master List page first."
      );
    } catch (error) {
      setMasterItems([]);
      setMessage(error?.message || "Could not load equipment master list.");
    } finally {
      setMasterLoading(false);
    }
  };

  useEffect(() => {
    setEntryShip((current) => current || userShip || "");
  }, [userShip]);

  useEffect(() => {
    loadMasterItems();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const visibleItems = useMemo(() => {
    const term = search.toLowerCase().trim();

    return masterItems.filter((item) => {
      if (
        departmentFilter !== ALL_DEPARTMENTS &&
        item.equipmentDepartment !== departmentFilter
      ) {
        return false;
      }

      if (!term) return true;

      return [
        item.departmentLabel,
        item.code,
        item.name,
        item.category,
        item.sheetName,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  }, [masterItems, search, departmentFilter]);

  const uploadBreakagePhoto = async ({ file, item, ship }) => {
    if (!supabase || !file) return "";

    const extension = getFileExtension(file);
    const contentType = file.type || `image/${extension}`;
    const department = item?.equipmentDepartment || extraDepartment || "equipment";
    const itemPart = makeStorageSafePart(item?.code || item?.name || "breakage");

    const path =
      "breakage/" +
      makeStorageSafePart(ship || "ship") +
      "/" +
      makeStorageSafePart(department) +
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
      department: item.equipmentDepartment || "culinary",
      month_key: getMonthKey(),

      user_name: effectiveReportedBy,
      user_position: reportedPosition || "",

      item_key: item.itemKey || makeItemKey(item),
      code: item.code || "",
      item_name: item.name || "",
      category: item.category || "",
      sheet_name: item.sheetName || "",
      image: item.image || "",

      qty: safeQty,
      notes: notes || "",
      report_photo: photoUrl,

      reported_at: now,
      updated_at: now,
    };

    const { error } = await supabase
      .from("equipment_breakage_reports")
      .insert(payload);

    if (error) throw error;

    if (typeof logUsageEvent === "function") {
      logUsageEvent("equipment_breakage_saved", {
        module: "breakage_report",
        ship: entryShip,
        department: item.equipmentDepartment || "",
        userEmail,
        userName: effectiveReportedBy,
        itemName: item.name || "",
        code: item.code || "",
        qty: safeQty,
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

      setCurrentItem(null);
      setBreakageQty("");
      setBreakageNotes("");
      setBreakagePhotoFile(null);
      setBreakagePhotoInputKey((value) => value + 1);

      setMessage(`Saved breakage: ${currentItem.name}.`);
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
        equipmentDepartment: extraDepartment,
        itemKey: cleanText(
          `${extraDepartment}|EXTRA|${extraCode || "NO-CODE"}|${name}|${Date.now()}`
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

      setExtraCode("");
      setExtraName("");
      setExtraQty("");
      setExtraNotes("");
      setExtraPhotoFile(null);
      setExtraPhotoInputKey((value) => value + 1);

      setMessage(`Saved broken item not in master list: ${name}.`);
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
            ← Modules
          </button>

          <div style={styles.shipBadge}>🧾 Breakage Report</div>
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
            onClick={loadMasterItems}
            disabled={masterLoading}
          >
            {masterLoading ? "Loading..." : "🔄 Refresh Equipment Master List"}
          </button>

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.infoBox}>
            <div>
              🚢 Ship: <strong>{entryShip || "Not selected"}</strong>
            </div>
            <div>
              👤 User: <strong>{effectiveReportedBy || "Not selected"}</strong>
            </div>
            <div>
              📋 Master list items: <strong>{masterItems.length}</strong>
            </div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>➕ Item Not In Master List</h2>

          <label style={styles.label}>Department</label>
          <select
            value={extraDepartment}
            onChange={(event) => setExtraDepartment(event.target.value)}
            style={styles.select}
          >
            {DEPARTMENTS.map((department) => (
              <option key={department.key} value={department.key}>
                {department.label}
              </option>
            ))}
          </select>

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
        <h2 style={styles.productTitle}>📋 Equipment Master List</h2>

        <div style={styles.grid}>
          <div>
            <label style={styles.label}>Department filter</label>
            <select
              value={departmentFilter}
              onChange={(event) => setDepartmentFilter(event.target.value)}
              style={styles.select}
            >
              <option value={ALL_DEPARTMENTS}>All Departments</option>
              {DEPARTMENTS.map((department) => (
                <option key={department.key} value={department.key}>
                  {department.label}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label style={styles.label}>Search equipment</label>
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Search code, name, category, sheet..."
              style={styles.searchInput}
            />
          </div>
        </div>

        {visibleItems.length === 0 && (
          <p style={styles.emptyText}>
            No equipment found. Upload the department master list first from
            Equipment Master List.
          </p>
        )}

        <div style={styles.equipmentGrid}>
          {visibleItems.map((item) => (
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
                      const fallback = event.currentTarget.nextElementSibling;
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
                Department: {item.departmentLabel}
              </div>
              <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
              <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
              <div style={styles.recipeMeta}>Category: {item.category}</div>
            </button>
          ))}
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
                  <strong>Department:</strong> {currentItem.departmentLabel}
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

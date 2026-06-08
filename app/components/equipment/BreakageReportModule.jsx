"use client";

import React, { useEffect, useMemo, useState } from "react";
import { cleanText, formatQty, getImageUrl } from "../../lib/appHelpers";
import { getMasterInventoryScope } from "../../constants/inventoryConfig";

const ALL = "ALL";

const SHIPS = ["SC", "VL", "BRL", "RL"];

const DEPARTMENTS = [
  { key: "culinary", label: "Culinary", icon: "👨‍🍳" },
  { key: "bar", label: "Bar", icon: "🍸" },
  { key: "restaurant", label: "Restaurant", icon: "🍽️" },
];

const DEPARTMENT_LABELS = {
  culinary: "Culinary",
  bar: "Bar",
  restaurant: "Restaurant",
};

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

const makeBreakageItemKey = (item) =>
  cleanText(
    [
      item?.equipmentDepartment || item?.department || "equipment",
      item?.sheetName || "",
      item?.category || "",
      item?.code || "",
      item?.name || "",
      item?.sourceRow || "",
    ].join("|")
  );

const normalizeReportRow = (row) => ({
  id: row.id,
  ship: row.ship || "",
  department: row.department || "",
  itemKey: row.item_key || "",
  code: row.code || "",
  name: row.item_name || "",
  category: row.category || "",
  sheetName: row.sheet_name || "",
  image: row.image || "",
  qty: Number(row.qty || 0),
  notes: row.notes || "",
  userName: row.user_name || "",
  userPosition: row.user_position || "",
  monthKey: row.month_key || "",
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
  getShipDisplayName,
  logUsageEvent,
  onBack,
}) {
  const monthKey = getMonthKey();
  const monthLabel = getMonthLabel(monthKey);

  const [masterItems, setMasterItems] = useState([]);
  const [masterLoading, setMasterLoading] = useState(false);
  const [masterMessage, setMasterMessage] = useState("");

  const [reportRows, setReportRows] = useState([]);
  const [reportLoading, setReportLoading] = useState(false);
  const [reportMessage, setReportMessage] = useState("");

  const [entryShip, setEntryShip] = useState(userShip || "");
  const [reportedBy, setReportedBy] = useState("");
  const [reportedPosition, setReportedPosition] = useState("");

  const [equipmentSearch, setEquipmentSearch] = useState("");
  const [reportSearch, setReportSearch] = useState("");

  const [equipmentDepartmentFilter, setEquipmentDepartmentFilter] =
    useState(ALL);
  const [reportShipFilter, setReportShipFilter] = useState(ALL);
  const [reportDepartmentFilter, setReportDepartmentFilter] = useState(ALL);

  const [currentItem, setCurrentItem] = useState(null);
  const [breakageQty, setBreakageQty] = useState("");
  const [breakageNotes, setBreakageNotes] = useState("");
  const [saving, setSaving] = useState(false);

  const getShipName =
    typeof getShipDisplayName === "function"
      ? getShipDisplayName
      : (ship) => ship;

  const departmentScopes = useMemo(() => {
    const scopeMap = {};

    DEPARTMENTS.forEach((department) => {
      scopeMap[department.key] = getMasterInventoryScope(department.key);
    });

    return scopeMap;
  }, []);

  const departmentByScope = useMemo(() => {
    const result = {};

    Object.entries(departmentScopes).forEach(([department, scope]) => {
      result[scope] = department;
    });

    return result;
  }, [departmentScopes]);

  const effectiveReportedBy = [reportedBy.trim(), reportedPosition.trim()]
    .filter(Boolean)
    .join(" - ");

  const loadMasterItems = async () => {
    if (!supabase) {
      setMasterItems([]);
      setMasterMessage("Supabase is not connected. Master list cannot load.");
      return;
    }

    setMasterLoading(true);
    setMasterMessage("Loading permanent equipment master list...");

    try {
      const scopes = Object.values(departmentScopes).filter(Boolean);

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

      const rows = (data || [])
        .map((row, index) => {
          const department =
            departmentByScope[row.ship] || row.department || "culinary";

          const item = {
            id: row.id || `${row.ship}-${index}`,
            equipmentDepartment: department,
            departmentLabel: DEPARTMENT_LABELS[department] || department,
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
        .filter((item) => item.name || item.code)
        .sort((a, b) => {
          const departmentOrder =
            DEPARTMENTS.findIndex((item) => item.key === a.equipmentDepartment) -
            DEPARTMENTS.findIndex((item) => item.key === b.equipmentDepartment);

          if (departmentOrder !== 0) return departmentOrder;

          return (
            Number(a.sortOrder || 0) - Number(b.sortOrder || 0) ||
            String(a.name || "").localeCompare(String(b.name || ""))
          );
        });

      setMasterItems(rows);
      setMasterMessage(
        rows.length
          ? `Permanent equipment master list loaded. ${rows.length} item(s).`
          : "No permanent master list found. Upload each department master list once from the existing Equipment Master List page."
      );
    } catch (error) {
      setMasterItems([]);
      setMasterMessage(
        error?.message || "Could not load permanent equipment master list."
      );
    } finally {
      setMasterLoading(false);
    }
  };

  const loadReportRows = async () => {
    if (!supabase) {
      setReportRows([]);
      setReportMessage("Supabase is not connected. Report cannot load.");
      return;
    }

    setReportLoading(true);
    setReportMessage("Loading monthly breakage report...");

    try {
      const { data, error } = await supabase
        .from("equipment_breakage_reports")
        .select("*")
        .eq("month_key", monthKey)
        .order("reported_at", { ascending: false })
        .limit(5000);

      if (error) throw error;

      const rows = (data || []).map(normalizeReportRow);

      setReportRows(rows);
      setReportMessage(
        rows.length
          ? `Monthly breakage report loaded. ${rows.length} record(s) for ${monthLabel}.`
          : `No breakage records yet for ${monthLabel}.`
      );
    } catch (error) {
      setReportRows([]);
      setReportMessage(
        error?.message || "Could not load monthly breakage report."
      );
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
  }, []);

  const visibleMasterItems = useMemo(() => {
    const term = equipmentSearch.toLowerCase().trim();

    return masterItems.filter((item) => {
      if (
        equipmentDepartmentFilter !== ALL &&
        item.equipmentDepartment !== equipmentDepartmentFilter
      ) {
        return false;
      }

      if (!term) return true;

      return [
        item.departmentLabel,
        item.sheetName,
        item.category,
        item.code,
        item.name,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  }, [masterItems, equipmentSearch, equipmentDepartmentFilter]);

  const visibleReportRows = useMemo(() => {
    const term = reportSearch.toLowerCase().trim();

    return reportRows.filter((row) => {
      if (reportShipFilter !== ALL && row.ship !== reportShipFilter) {
        return false;
      }

      if (
        reportDepartmentFilter !== ALL &&
        row.department !== reportDepartmentFilter
      ) {
        return false;
      }

      if (!term) return true;

      return [
        row.ship,
        row.department,
        row.code,
        row.name,
        row.category,
        row.sheetName,
        row.userName,
        row.notes,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  }, [reportRows, reportSearch, reportShipFilter, reportDepartmentFilter]);

  const summaryRows = useMemo(() => {
    const grouped = new Map();

    visibleReportRows.forEach((row) => {
      const key = `${row.ship}|${row.department}|${row.itemKey}`;

      if (!grouped.has(key)) {
        grouped.set(key, {
          ship: row.ship,
          department: row.department,
          itemKey: row.itemKey,
          code: row.code || "",
          name: row.name || "",
          category: row.category || "",
          sheetName: row.sheetName || "",
          image: row.image || "",
          totalQty: 0,
          records: 0,
          users: new Set(),
          lastReportedAt: "",
        });
      }

      const item = grouped.get(key);

      item.totalQty += Number(row.qty || 0);
      item.records += 1;

      if (row.userName) item.users.add(row.userName);

      if (
        !item.lastReportedAt ||
        new Date(row.reportedAt || 0).getTime() >
          new Date(item.lastReportedAt || 0).getTime()
      ) {
        item.lastReportedAt = row.reportedAt;
      }
    });

    return Array.from(grouped.values())
      .map((item) => ({
        ...item,
        users: Array.from(item.users).sort(),
        confirmedAt: item.lastReportedAt
          ? new Date(item.lastReportedAt).toLocaleString()
          : "",
      }))
      .sort(
        (a, b) =>
          Number(b.totalQty || 0) - Number(a.totalQty || 0) ||
          String(a.name || "").localeCompare(String(b.name || ""))
      );
  }, [visibleReportRows]);

  const getItemMonthlyCount = (item, shipFilter = entryShip) => {
    const itemKey = item?.itemKey || makeBreakageItemKey(item);

    return reportRows
      .filter((row) => row.monthKey === monthKey)
      .filter((row) => row.itemKey === itemKey)
      .filter((row) => row.department === item.equipmentDepartment)
      .filter((row) => shipFilter === ALL || row.ship === shipFilter)
      .reduce((sum, row) => sum + Number(row.qty || 0), 0);
  };

  const openBreakagePopup = (item) => {
    if (!entryShip) {
      setReportMessage("Choose ship before reporting broken equipment.");
      return;
    }

    setCurrentItem(item);
    setBreakageQty("");
    setBreakageNotes("");
  };

  const confirmBreakage = async () => {
    if (saving) return;

    if (!supabase) {
      window.alert("Supabase is not connected. Breakage cannot be saved.");
      return;
    }

    if (!currentItem) {
      setReportMessage("Select equipment before confirming breakage.");
      return;
    }

    if (!entryShip) {
      window.alert("Choose ship before confirming breakage.");
      return;
    }

    if (!reportedBy.trim()) {
      window.alert("Enter reported by name before confirming breakage.");
      return;
    }

    const qty = Number(breakageQty || 0);

    if (String(breakageQty).trim() === "" || Number.isNaN(qty) || qty <= 0) {
      window.alert("Enter a breakage number greater than 0.");
      return;
    }

    const previousCount = getItemMonthlyCount(currentItem, entryShip);
    const now = new Date().toISOString();

    const payload = {
      ship: entryShip,
      department: currentItem.equipmentDepartment,
      month_key: monthKey,

      user_name: effectiveReportedBy,
      user_position: reportedPosition || "",

      item_key: currentItem.itemKey || makeBreakageItemKey(currentItem),
      code: currentItem.code || "",
      item_name: currentItem.name || "",
      category: currentItem.category || "",
      sheet_name: currentItem.sheetName || "",
      image: currentItem.image || "",

      qty,
      notes: breakageNotes || "",

      reported_at: now,
      updated_at: now,
    };

    setSaving(true);
    setReportMessage("Saving breakage report...");

    try {
      const { data, error } = await supabase
        .from("equipment_breakage_reports")
        .insert(payload)
        .select("*")
        .single();

      if (error) throw error;

      const savedRow = normalizeReportRow(data);

      setReportRows((current) => [savedRow, ...current]);
      setCurrentItem(null);
      setBreakageQty("");
      setBreakageNotes("");

      setReportMessage(
        `Saved: ${currentItem.name} / Added ${formatQty(
          qty
        )}. ${entryShip} monthly total is now ${formatQty(previousCount + qty)}.`
      );

      if (typeof logUsageEvent === "function") {
        logUsageEvent("equipment_breakage_saved", {
          module: "breakage_report",
          ship: entryShip,
          department: currentItem.equipmentDepartment,
          userEmail,
          userName: effectiveReportedBy,
          itemName: currentItem.name,
          code: currentItem.code,
          qty,
          previousCount,
          newMonthlyCount: previousCount + qty,
          monthKey,
        });
      }
    } catch (error) {
      const text = error?.message || "Could not save breakage report.";
      setReportMessage(text);
      window.alert(text);
    } finally {
      setSaving(false);
    }
  };

  const totalBreakageQty = visibleReportRows.reduce(
    (sum, row) => sum + Number(row.qty || 0),
    0
  );

  const currentItemShipCount = currentItem
    ? getItemMonthlyCount(currentItem, entryShip)
    : 0;

  const currentItemAllShipsCount = currentItem
    ? getItemMonthlyCount(currentItem, ALL)
    : 0;

  const insertedQty = Number(breakageQty || 0);
  const nextShipCount =
    currentItemShipCount + (Number.isFinite(insertedQty) ? insertedQty : 0);

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
          <h2 style={styles.cardTitle}>🧾 Report Broken Equipment</h2>

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

          <div style={styles.headerActions}>
            <button
              type="button"
              style={styles.backButton}
              onClick={loadMasterItems}
              disabled={masterLoading}
            >
              {masterLoading ? "Loading..." : "🔄 Refresh Master List"}
            </button>

            <button
              type="button"
              style={styles.backButton}
              onClick={loadReportRows}
              disabled={reportLoading}
            >
              {reportLoading ? "Loading..." : "🔄 Refresh Report"}
            </button>
          </div>

          <div style={styles.infoBox}>
            <div>
              📅 Active report month: <strong>{monthLabel}</strong>
            </div>
            <div>
              📋 Permanent master items:{" "}
              <strong>{masterItems.length}</strong>
            </div>
            <div>
              🧾 Report records shown:{" "}
              <strong>{visibleReportRows.length}</strong>
            </div>
            <div>
              🔢 Total broken quantity shown:{" "}
              <strong>{formatQty(totalBreakageQty)}</strong>
            </div>
            <div>
              This block loads independently from Inventory and only uses the
              permanent master list.
            </div>
          </div>

          {masterMessage && <p style={styles.message}>{masterMessage}</p>}
          {reportMessage && <p style={styles.message}>{reportMessage}</p>}
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔎 Filters</h2>

          <label style={styles.label}>Master list department</label>
          <select
            value={equipmentDepartmentFilter}
            onChange={(event) =>
              setEquipmentDepartmentFilter(event.target.value)
            }
            style={styles.select}
          >
            <option value={ALL}>All Departments</option>
            {DEPARTMENTS.map((department) => (
              <option key={department.key} value={department.key}>
                {department.label}
              </option>
            ))}
          </select>

          <label style={styles.label}>Search equipment master list</label>
          <input
            value={equipmentSearch}
            onChange={(event) => setEquipmentSearch(event.target.value)}
            placeholder="Search equipment, code, category, sheet..."
            style={styles.searchInput}
          />

          <label style={styles.label}>Report ship filter</label>
          <select
            value={reportShipFilter}
            onChange={(event) => setReportShipFilter(event.target.value)}
            style={styles.select}
          >
            <option value={ALL}>All Ships</option>
            {SHIPS.map((ship) => (
              <option key={ship} value={ship}>
                {getShipName(ship)}
              </option>
            ))}
          </select>

          <label style={styles.label}>Report department filter</label>
          <select
            value={reportDepartmentFilter}
            onChange={(event) => setReportDepartmentFilter(event.target.value)}
            style={styles.select}
          >
            <option value={ALL}>All Departments</option>
            {DEPARTMENTS.map((department) => (
              <option key={department.key} value={department.key}>
                {department.label}
              </option>
            ))}
          </select>

          <label style={styles.label}>Search report</label>
          <input
            value={reportSearch}
            onChange={(event) => setReportSearch(event.target.value)}
            placeholder="Search report records..."
            style={styles.searchInput}
          />
        </div>
      </section>

      <section style={styles.card}>
        <h2 style={styles.productTitle}>📋 Equipment Master List</h2>

        {masterItems.length === 0 && (
          <p style={styles.emptyText}>
            No permanent master list loaded yet. Upload each department master
            list once from the existing Equipment Master List page.
          </p>
        )}

        {masterItems.length > 0 && visibleMasterItems.length === 0 && (
          <p style={styles.emptyText}>
            No equipment matches the selected filter.
          </p>
        )}

        <div style={styles.equipmentGrid}>
          {visibleMasterItems.map((item) => {
            const shipCount = entryShip
              ? getItemMonthlyCount(item, entryShip)
              : 0;

            const allShipsCount = getItemMonthlyCount(item, ALL);

            return (
              <button
                key={`${item.equipmentDepartment}-${item.itemKey}-${item.id}`}
                type="button"
                style={{
                  ...styles.inventoryItemCard,
                  ...(allShipsCount > 0 ? styles.zeroCountCard : {}),
                }}
                onClick={() => openBreakagePopup(item)}
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

                <div style={styles.statusNeutral}>
                  {entryShip || "Ship"} this month: {formatQty(shipCount)}
                </div>

                <div
                  style={allShipsCount > 0 ? styles.statusWarning : styles.statusNeutral}
                >
                  All ships this month: {formatQty(allShipsCount)}
                </div>
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
                  <strong>Department:</strong>{" "}
                  {DEPARTMENT_LABELS[currentItem.equipmentDepartment] ||
                    currentItem.equipmentDepartment}
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

                <div style={styles.infoBox}>
                  <div>
                    📅 Month: <strong>{monthLabel}</strong>
                  </div>
                  <div>
                    Current count for {entryShip}:{" "}
                    <strong>{formatQty(currentItemShipCount)}</strong>
                  </div>
                  <div>
                    Current count all ships:{" "}
                    <strong>{formatQty(currentItemAllShipsCount)}</strong>
                  </div>
                  <div>
                    New {entryShip} total after confirm:{" "}
                    <strong>{formatQty(nextShipCount)}</strong>
                  </div>
                </div>

                <label style={styles.label}>Broken quantity</label>
                <input
                  autoFocus
                  type="number"
                  min="1"
                  step="1"
                  value={breakageQty}
                  onChange={(event) => setBreakageQty(event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" && !saving) {
                      confirmBreakage();
                    }
                  }}
                  placeholder="Enter broken quantity..."
                  style={styles.searchInput}
                />

                <label style={styles.label}>Notes optional</label>
                <input
                  value={breakageNotes}
                  onChange={(event) => setBreakageNotes(event.target.value)}
                  placeholder="Example: chipped, cracked, dropped, missing..."
                  style={styles.searchInput}
                />

                <div style={styles.headerActions}>
                  <button
                    type="button"
                    style={styles.backButton}
                    onClick={() => {
                      setCurrentItem(null);
                      setBreakageQty("");
                      setBreakageNotes("");
                    }}
                  >
                    Cancel
                  </button>

                  <button
                    type="button"
                    style={styles.primaryButton}
                    onClick={confirmBreakage}
                    disabled={saving}
                  >
                    {saving ? "Saving..." : "Confirm Breakage"}
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

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
            <h2 style={styles.productTitle}>📄 Monthly Breakage Report</h2>
            <p style={{ ...styles.emptyText, margin: 0 }}>
              Current month report with filters for all ships, specific ship,
              all departments, or one department.
            </p>
          </div>
        </div>

        <div style={styles.infoBox}>
          <div>
            📅 Month: <strong>{monthLabel}</strong>
          </div>
          <div>
            🚢 Ship filter:{" "}
            <strong>
              {reportShipFilter === ALL
                ? "All Ships"
                : getShipName(reportShipFilter)}
            </strong>
          </div>
          <div>
            🧭 Department filter:{" "}
            <strong>
              {reportDepartmentFilter === ALL
                ? "All Departments"
                : DEPARTMENT_LABELS[reportDepartmentFilter]}
            </strong>
          </div>
          <div>
            🧾 Records: <strong>{visibleReportRows.length}</strong>
          </div>
          <div>
            📦 Summary equipment items: <strong>{summaryRows.length}</strong>
          </div>
          <div>
            🔢 Total broken quantity:{" "}
            <strong>{formatQty(totalBreakageQty)}</strong>
          </div>
        </div>

        {summaryRows.length === 0 && (
          <p style={styles.emptyText}>
            No breakage records match the selected filters.
          </p>
        )}

        {summaryRows.length > 0 && (
          <>
            <h3 style={styles.sectionTitle}>📊 Summary</h3>

            <div style={styles.equipmentGrid}>
              {summaryRows.map((item) => (
                <div
                  key={`${item.ship}-${item.department}-${item.itemKey}`}
                  style={styles.equipmentCard}
                >
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
                  <div style={styles.recipeMeta}>
                    Ship: {getShipName(item.ship)}
                  </div>
                  <div style={styles.recipeMeta}>
                    Department:{" "}
                    {DEPARTMENT_LABELS[item.department] || item.department}
                  </div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Category: {item.category}</div>
                  <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                  <div style={styles.statusBad}>
                    Total Broken: {formatQty(item.totalQty)}
                  </div>
                  <div style={styles.recipeMeta}>Records: {item.records}</div>
                  <div style={styles.recipeMeta}>
                    Users: {item.users.join(", ") || "N/A"}
                  </div>
                  <div style={styles.recipeMeta}>
                    Last Reported: {item.confirmedAt || "N/A"}
                  </div>
                </div>
              ))}
            </div>
          </>
        )}

        {visibleReportRows.length > 0 && (
          <>
            <h3 style={styles.sectionTitle}>🧾 Entries</h3>

            <div style={styles.equipmentGrid}>
              {visibleReportRows.map((row) => (
                <div key={row.id} style={styles.equipmentCard}>
                  <div style={styles.recipeName}>{row.name}</div>
                  <div style={styles.recipeMeta}>
                    Ship: {getShipName(row.ship)}
                  </div>
                  <div style={styles.recipeMeta}>
                    Department:{" "}
                    {DEPARTMENT_LABELS[row.department] || row.department}
                  </div>
                  <div style={styles.recipeMeta}>Code: {row.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Category: {row.category}</div>
                  <div style={styles.recipeMeta}>Sheet: {row.sheetName}</div>
                  <div style={styles.statusBad}>
                    Added Broken Qty: {formatQty(row.qty)}
                  </div>
                  <div style={styles.recipeMeta}>
                    Reported by: {row.userName || "N/A"}
                  </div>
                  <div style={styles.recipeMeta}>
                    Confirmed: {row.confirmedAt || "N/A"}
                  </div>
                  {row.notes && (
                    <div style={styles.statusNeutral}>Notes: {row.notes}</div>
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

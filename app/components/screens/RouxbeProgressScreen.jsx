"use client";

import React, { useEffect, useMemo, useState } from "react";

const SHIP_SHEET_ALIASES = {
  SC: ["SCARLET", "SCARLET LADY", "SC"],
  VL: ["VALIANT", "VALIANT LADY", "VL"],
  RL: ["RESILIENT", "RESILIANT", "RESILIENT LADY", "RL", "RES"],
  BRL: ["BRILLIANT", "BRILIANT", "BRILLIANT LADY", "BRL", "BR"],
};

const SHIP_DISPLAY_NAMES = {
  SC: "Scarlet",
  VL: "Valiant",
  RL: "Resilient",
  BRL: "Brilliant",
};

const DEFAULT_ROUXBE_ROSTER_PATH = "/rouxbe-groups.xlsx";

const loadXlsx = async () => {
  const module = await import("xlsx");
  return module.default || module;
};

const cleanText = (value) =>
  String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

const normalizeHeader = (value) => cleanText(value).replace(/[^A-Z0-9]/g, "");

const normalizeName = (value) =>
  cleanText(value)
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const normalizeCrewId = (value) => {
  let text = String(value || "").trim();

  if (!text) return "";

  if (text.includes("@")) {
    text = text.split("@")[0];
  }

  if (/^[+-]?\d+(\.\d+)?e[+-]?\d+$/i.test(text) || /^\d+\.0+$/.test(text)) {
    const numericValue = Number(text);
    if (Number.isFinite(numericValue)) {
      text = String(Math.round(numericValue));
    }
  }

  const digits = text.replace(/\D+/g, "");
  return digits || cleanText(text);
};

const toNumber = (value) => {
  const text = String(value ?? "")
    .replace(/%/g, "")
    .replace(/,/g, "")
    .trim();

  if (!text || text === "--") return 0;

  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
};

const getRosterNameKeys = (name) => {
  const text = String(name || "").replace(/\s+/g, " ").trim();
  const keys = new Set();

  const add = (value) => {
    const key = normalizeName(value);
    if (key) keys.add(key);
  };

  add(text);

  if (text.includes(",")) {
    const [last, ...rest] = text.split(",");
    const first = rest.join(" ").trim();

    add(`${first} ${last}`);
    add(`${last} ${first}`);
  }

  return Array.from(keys);
};

const getProgressNameKeys = (firstName, lastName) => {
  const keys = new Set();
  const first = String(firstName || "").trim();
  const last = String(lastName || "").trim();

  [
    `${first} ${last}`,
    `${last} ${first}`,
  ].forEach((value) => {
    const key = normalizeName(value);
    if (key) keys.add(key);
  });

  return Array.from(keys);
};

const getShipSheetName = (workbook, userShip) => {
  const sheetNames = workbook?.SheetNames || [];
  const aliases = SHIP_SHEET_ALIASES[userShip] || [];

  const exactMatch = sheetNames.find((sheetName) =>
    aliases.some((alias) => cleanText(sheetName) === cleanText(alias))
  );

  if (exactMatch) return exactMatch;

  const containsMatch = sheetNames.find((sheetName) =>
    aliases.some((alias) => cleanText(sheetName).includes(cleanText(alias)))
  );

  if (containsMatch) return containsMatch;

  return sheetNames[0] || "";
};

const findGroupLabel = (rows, headerRowIndex, nameColumnIndex) => {
  const startRow = Math.max(0, headerRowIndex - 8);
  const columnsToCheck = [nameColumnIndex - 1, nameColumnIndex, nameColumnIndex + 1];

  for (let rowIndex = startRow; rowIndex <= headerRowIndex; rowIndex += 1) {
    const row = rows[rowIndex] || [];

    for (const columnIndex of columnsToCheck) {
      if (columnIndex < 0) continue;

      const value = String(row[columnIndex] || "").replace(/\s+/g, " ").trim();

      if (cleanText(value).includes("GROUP")) {
        return value;
      }
    }
  }

  return "Group";
};

const parseRouxbeRosterWorkbook = (workbook, userShip, XLSX) => {
  const sheetName = getShipSheetName(workbook, userShip);

  if (!sheetName) {
    return {
      sheetName: "",
      rows: [],
      sheetNames: workbook?.SheetNames || [],
    };
  }

  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
  });

  const headerBlocks = [];

  rows.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      const current = cleanText(cell);
      const next = cleanText(row[columnIndex + 1]);
      const afterNext = cleanText(row[columnIndex + 2]);

      if (current === "NAME" && next.includes("ID")) {
        headerBlocks.push({
          headerRowIndex: rowIndex,
          nameColumnIndex: columnIndex,
          idColumnIndex: columnIndex + 1,
          positionColumnIndex: columnIndex + 2,
          group: findGroupLabel(rows, rowIndex, columnIndex),
          hasPosition: afterNext.includes("POSITION"),
        });
      }
    });
  });

  const rosterRows = [];
  const seenRowKeys = new Set();

  headerBlocks.forEach((block) => {
    rows.slice(block.headerRowIndex + 1).forEach((row, offset) => {
      const sourceRowNumber = block.headerRowIndex + 2 + offset;
      const number = String(row[block.nameColumnIndex - 1] || "").trim();
      const name = String(row[block.nameColumnIndex] || "").replace(/\s+/g, " ").trim();
      const id = normalizeCrewId(row[block.idColumnIndex]);
      const position = String(row[block.positionColumnIndex] || "").replace(/\s+/g, " ").trim();
      const normalizedName = normalizeName(name);

      if (!id && !normalizedName) return;
      if (!normalizedName && !id) return;
      if (cleanText(name) === "NAME") return;

      const rowKey = `${block.group}|${sourceRowNumber}|${block.nameColumnIndex}|${id || normalizedName}`;

      if (seenRowKeys.has(rowKey)) return;
      seenRowKeys.add(rowKey);

      rosterRows.push({
        rosterKey: rowKey,
        ship: userShip,
        sheetName,
        group: block.group,
        number,
        name,
        id,
        position,
        sourceRowNumber,
        nameKeys: getRosterNameKeys(name),
      });
    });
  });

  return {
    sheetName,
    sheetNames: workbook.SheetNames || [],
    rows: rosterRows,
  };
};


const parseCsvText = (text) => {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        value += '"';
        index += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        value += char;
      }

      continue;
    }

    if (char === '"') {
      inQuotes = true;
    } else if (char === ",") {
      row.push(value);
      value = "";
    } else if (char === "\n") {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else if (char === "\r") {
      if (next !== "\n") {
        row.push(value);
        rows.push(row);
        row = [];
        value = "";
      }
    } else {
      value += char;
    }
  }

  if (value || row.length) {
    row.push(value);
    rows.push(row);
  }

  return rows.filter((item) => item.some((cell) => String(cell || "").trim()));
};

const getCsvCellGetter = (headers) => {
  const headerMap = {};

  headers.forEach((header, index) => {
    const key = normalizeHeader(header);
    if (key && headerMap[key] === undefined) {
      headerMap[key] = index;
    }
  });

  return (row, aliases) => {
    for (const alias of aliases) {
      const index = headerMap[normalizeHeader(alias)];

      if (index !== undefined) {
        return String(row[index] ?? "").trim();
      }
    }

    return "";
  };
};

const getProgressCategory = ({ matched, percentComplete, status, finishAt }) => {
  const statusText = cleanText(status);
  const finishText = String(finishAt || "").trim();

  if (!matched) return "No Record";
  if (statusText.includes("COMPLETE") || percentComplete >= 100 || (finishText && finishText !== "--")) return "Completed";
  if (percentComplete > 0 || statusText.includes("IN PROGRESS")) return "In Progress";
  return "Not Started";
};

const getProgressPriority = (row) => {
  const category = getProgressCategory({
    matched: true,
    percentComplete: row.percentComplete,
    status: row.status,
    finishAt: row.finishAt,
  });

  const categoryScore =
    category === "Completed" ? 1000000 :
    category === "In Progress" ? 500000 :
    0;

  const lastEngagedTime = Date.parse(row.lastEngaged || "") || 0;
  return categoryScore + Number(row.percentComplete || 0) * 1000 + lastEngagedTime / 100000000000;
};

const parseRouxbeProgressCsv = (text) => {
  const rows = parseCsvText(text);

  if (!rows.length) return [];

  const headers = rows[0];
  const getCell = getCsvCellGetter(headers);

  return rows.slice(1).map((row, index) => {
    const firstName = getCell(row, ["First Name", "First"]);
    const lastName = getCell(row, ["Last Name", "Last"]);
    const email = getCell(row, ["Email", "User Email"]);
    const customId = getCell(row, ["Custom ID", "CustomID"]);
    const percentComplete = toNumber(getCell(row, ["Percent Complete", "Percent"]));
    const grade = toNumber(getCell(row, ["Grade"]));
    const status = getCell(row, ["Status"]);
    const progress = getCell(row, ["Progress"]);
    const course = getCell(row, ["Course"]);
    const enrolledAt = getCell(row, ["Enrolled At"]);
    const startAt = getCell(row, ["Start At", "Started At"]);
    const finishAt = getCell(row, ["Finish At", "Finished At"]);
    const lastEngaged = getCell(row, ["Last Engaged"]);
    const adminUrl = getCell(row, ["Admin URL", "Admin"]);
    const id = normalizeCrewId(email);

    return {
      progressKey: `${id || normalizeName(`${firstName} ${lastName}`)}-${index}`,
      id,
      firstName,
      lastName,
      fullName: `${firstName} ${lastName}`.replace(/\s+/g, " ").trim(),
      nameKeys: getProgressNameKeys(firstName, lastName),
      customId,
      course,
      grade,
      progress,
      percentComplete,
      status,
      enrolledAt,
      startAt,
      finishAt,
      lastEngaged,
      adminUrl,
      sourceRowNumber: index + 2,
    };
  }).filter((row) => row.id || row.fullName);
};

const buildProgressIndexes = (progressRows) => {
  const byId = new Map();
  const byName = new Map();

  const putBest = (map, key, row) => {
    if (!key) return;

    const current = map.get(key);

    if (!current || getProgressPriority(row) >= getProgressPriority(current)) {
      map.set(key, row);
    }
  };

  progressRows.forEach((row) => {
    putBest(byId, row.id, row);
    row.nameKeys.forEach((nameKey) => putBest(byName, nameKey, row));
  });

  return { byId, byName };
};

const mergeRosterWithProgress = (rosterRows, progressRows) => {
  const indexes = buildProgressIndexes(progressRows);

  return rosterRows.map((person) => {
    let progress = person.id ? indexes.byId.get(person.id) : null;

    if (!progress) {
      for (const nameKey of person.nameKeys || []) {
        progress = indexes.byName.get(nameKey);
        if (progress) break;
      }
    }

    const matched = Boolean(progress);
    const percentComplete = matched ? Number(progress.percentComplete || 0) : 0;
    const category = getProgressCategory({
      matched,
      percentComplete,
      status: progress?.status || "",
      finishAt: progress?.finishAt || "",
    });

    return {
      ...person,
      matched,
      progressCategory: category,
      progressStatus: progress?.status || (matched ? "" : "No progress record"),
      progressText: progress?.progress || "",
      percentComplete,
      grade: progress?.grade || 0,
      course: progress?.course || "",
      enrolledAt: progress?.enrolledAt || "",
      startAt: progress?.startAt || "",
      finishAt: progress?.finishAt || "",
      lastEngaged: progress?.lastEngaged || "",
      adminUrl: progress?.adminUrl || "",
      progressName: progress?.fullName || "",
    };
  });
};

const getSummary = (rows) => {
  const total = rows.length;
  const matched = rows.filter((row) => row.matched).length;
  const completed = rows.filter((row) => row.progressCategory === "Completed").length;
  const inProgress = rows.filter((row) => row.progressCategory === "In Progress").length;
  const notStarted = rows.filter((row) => row.progressCategory === "Not Started").length;
  const noRecord = rows.filter((row) => row.progressCategory === "No Record").length;
  const averagePercent = matched
    ? rows.filter((row) => row.matched).reduce((sum, row) => sum + Number(row.percentComplete || 0), 0) / matched
    : 0;

  return {
    total,
    matched,
    completed,
    inProgress,
    notStarted,
    noRecord,
    averagePercent,
  };
};

const makeAdminUrl = (value) => {
  const url = String(value || "").trim();
  if (!url) return "";
  return /^https?:\/\//i.test(url) ? url : `https://${url}`;
};

const csvEscape = (value) => {
  const text = String(value ?? "");
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
};

const buildExportCsv = (rows) => {
  const headers = [
    "Ship",
    "Sheet",
    "Group",
    "Name",
    "Crew ID",
    "Position",
    "Progress Category",
    "Rouxbe Status",
    "Progress",
    "Percent Complete",
    "Grade",
    "Course",
    "Enrolled At",
    "Started At",
    "Finished At",
    "Last Engaged",
  ];

  const body = rows.map((row) => [
    row.ship,
    row.sheetName,
    row.group,
    row.name,
    row.id,
    row.position,
    row.progressCategory,
    row.progressStatus,
    row.progressText,
    row.percentComplete,
    row.grade,
    row.course,
    row.enrolledAt,
    row.startAt,
    row.finishAt,
    row.lastEngaged,
  ]);

  return [headers, ...body]
    .map((row) => row.map(csvEscape).join(","))
    .join("\n");
};

const statusStyle = (category) => {
  if (category === "Completed") {
    return { background: "#e8f5e9", color: "#2e7d32", border: "1px solid #2e7d32" };
  }

  if (category === "In Progress") {
    return { background: "#eef5ff", color: "#0057b8", border: "1px solid #0057b8" };
  }

  if (category === "No Record") {
    return { background: "#fff0f0", color: "#b00020", border: "1px solid #b00020" };
  }

  return { background: "#fff4d6", color: "#8a5a00", border: "1px solid #8a5a00" };
};

export default function RouxbeProgressScreen({
  styles = {},
  userShip = "",
  getShipDisplayName,
  setModule,
  logUsageEvent,
}) {
  const [rosterRows, setRosterRows] = useState([]);
  const [rosterMeta, setRosterMeta] = useState({ sheetName: "", fileName: "" });
  const [rosterLoading, setRosterLoading] = useState(false);
  const [progressRows, setProgressRows] = useState([]);
  const [progressFileName, setProgressFileName] = useState("");
  const [message, setMessage] = useState("");
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");

  const shipName =
    typeof getShipDisplayName === "function"
      ? getShipDisplayName(userShip)
      : SHIP_DISPLAY_NAMES[userShip] || userShip || "Ship";

  const loadRosterFromArrayBuffer = async ({ arrayBuffer, fileName }) => {
    const XLSX = await loadXlsx();
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const parsed = parseRouxbeRosterWorkbook(workbook, userShip, XLSX);

    setRosterRows(parsed.rows);
    setRosterMeta({
      sheetName: parsed.sheetName,
      fileName,
      sheetNames: parsed.sheetNames || [],
    });

    setMessage(
      parsed.rows.length
        ? `Roster loaded for ${shipName}. Tab used: ${parsed.sheetName}. ${parsed.rows.length} enrolled CM(s) found.`
        : `Roster loaded, but no enrolled CM rows were found for ${shipName}. Check the tab layout.`
    );

    logUsageEvent?.("rouxbe_roster_loaded", {
      module: "rouxbe_progress",
      ship: userShip,
      fileName,
      sheetName: parsed.sheetName,
      enrolledRows: parsed.rows.length,
    });
  };

  useEffect(() => {
    if (!userShip) return;

    let active = true;

    const loadDefaultRoster = async () => {
      setRosterLoading(true);
      setMessage("Loading Rouxbe enrolled CM roster...");

      try {
        const response = await fetch(DEFAULT_ROUXBE_ROSTER_PATH, { cache: "no-store" });

        if (!response.ok) {
          if (!active) return;

          setRosterRows([]);
          setRosterMeta({ sheetName: "", fileName: "" });
          setMessage(
            "Permanent Rouxbe Groups file was not found. Add public/rouxbe-groups.xlsx or upload the roster file on this screen."
          );
          return;
        }

        const arrayBuffer = await response.arrayBuffer();

        if (!active) return;

        await loadRosterFromArrayBuffer({
          arrayBuffer,
          fileName: "rouxbe-groups.xlsx",
        });
      } catch (error) {
        if (!active) return;

        setRosterRows([]);
        setRosterMeta({ sheetName: "", fileName: "" });
        setMessage(error?.message || "Could not load the Rouxbe roster file.");
      } finally {
        if (active) setRosterLoading(false);
      }
    };

    loadDefaultRoster();

    return () => {
      active = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userShip]);

  const handleRosterUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setRosterLoading(true);
      setMessage("Loading uploaded Rouxbe Groups file...");
      const arrayBuffer = await file.arrayBuffer();
      await loadRosterFromArrayBuffer({ arrayBuffer, fileName: file.name });
    } catch (error) {
      setRosterRows([]);
      setRosterMeta({ sheetName: "", fileName: file.name });
      setMessage(error?.message || "Could not read the uploaded Rouxbe Groups file.");
      window.alert(error?.message || "Could not read the uploaded Rouxbe Groups file.");
    } finally {
      setRosterLoading(false);
      event.target.value = "";
    }
  };

  const handleProgressUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setMessage("Loading Rouxbe progress CSV...");
      const text = await file.text();
      const parsed = parseRouxbeProgressCsv(text);

      setProgressRows(parsed);
      setProgressFileName(file.name);
      setMessage(
        `Progress file loaded. ${parsed.length} Rouxbe progress row(s) found. Progress is now matched to the ${shipName} enrolled CM list.`
      );

      logUsageEvent?.("rouxbe_progress_file_uploaded", {
        module: "rouxbe_progress",
        ship: userShip,
        fileName: file.name,
        progressRows: parsed.length,
        rosterRows: rosterRows.length,
      });
    } catch (error) {
      setProgressRows([]);
      setProgressFileName(file.name);
      setMessage(error?.message || "Could not read the Rouxbe progress CSV.");
      window.alert(error?.message || "Could not read the Rouxbe progress CSV.");
    } finally {
      event.target.value = "";
    }
  };

  const mergedRows = useMemo(
    () => mergeRosterWithProgress(rosterRows, progressRows),
    [rosterRows, progressRows]
  );

  const summary = useMemo(() => getSummary(mergedRows), [mergedRows]);

  const visibleRows = useMemo(() => {
    const term = search.toLowerCase().trim();

    return mergedRows.filter((row) => {
      if (statusFilter !== "all" && row.progressCategory !== statusFilter) {
        return false;
      }

      if (!term) return true;

      return [
        row.group,
        row.name,
        row.id,
        row.position,
        row.progressCategory,
        row.progressStatus,
        row.progressText,
        row.course,
        row.lastEngaged,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term);
    });
  }, [mergedRows, search, statusFilter]);

  const downloadVisibleCsv = () => {
    if (!visibleRows.length) {
      window.alert("No Rouxbe progress rows to export.");
      return;
    }

    const csv = buildExportCsv(visibleRows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `rouxbe-progress-${userShip || "ship"}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const cardStyle = styles.card || {
    background: "#fff",
    borderRadius: 16,
    padding: 20,
    boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
  };

  const buttonStyle = styles.primaryButton || {
    marginTop: 10,
    padding: 12,
    borderRadius: 10,
    border: 0,
    background: "#111",
    color: "#fff",
    fontWeight: "bold",
    cursor: "pointer",
  };

  const filterButton = (value, label, count) => (
    <button
      key={value}
      type="button"
      onClick={() => setStatusFilter(value)}
      style={{
        ...(styles.viewModeButton || {}),
        ...(statusFilter === value ? styles.viewModeButtonActive || { background: "#111", color: "#fff" } : {}),
      }}
    >
      {label} {typeof count === "number" ? `(${count})` : ""}
    </button>
  );

  return (
    <div style={styles.page || { minHeight: "100vh", padding: 24, background: "#f5f5f5", fontFamily: "Arial, sans-serif" }}>
      <div style={styles.header || { display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center", background: "#fff", borderRadius: 16, padding: 18, marginBottom: 16 }}>
        <div>
          <h1 style={{ margin: 0 }}>Rouxbe Progress</h1>
          <p style={styles.subtitle || { margin: "4px 0 0", color: "#666" }}>
            Ship: <strong>{shipName}</strong>{rosterMeta.sheetName ? ` / roster tab: ${rosterMeta.sheetName}` : ""}
          </p>
        </div>
        <button type="button" style={styles.backButton || buttonStyle} onClick={() => setModule?.("")}>
          ← Back
        </button>
      </div>

      <div style={{ ...cardStyle, marginBottom: 16, display: "grid", gap: 14 }}>
        <div>
          <h2 style={{ margin: 0 }}>Files</h2>
          <p style={styles.message || { color: "#555", fontSize: 14 }}>
            The enrolled CM roster comes from the Rouxbe Groups workbook. The app uses the selected ship to choose the correct tab automatically.
          </p>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 14 }}>
          <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
            Optional roster replacement
            <input
              type="file"
              accept=".xlsx,.xls,.xlsm"
              onChange={handleRosterUpload}
              disabled={rosterLoading}
              style={styles.fileInput}
            />
            <span style={{ color: "#666", fontSize: 13, fontWeight: "normal" }}>
              Current: {rosterMeta.fileName || "No roster loaded"}
            </span>
          </label>

          <label style={{ display: "grid", gap: 6, fontWeight: "bold" }}>
            Upload Rouxbe progress CSV
            <input
              type="file"
              accept=".csv,text/csv"
              onChange={handleProgressUpload}
              style={styles.fileInput}
            />
            <span style={{ color: "#666", fontSize: 13, fontWeight: "normal" }}>
              Current: {progressFileName || "No progress file uploaded"}
            </span>
          </label>
        </div>

        {message ? <div style={styles.infoBox || { padding: 12, borderRadius: 12, background: "#f2f2f2" }}>{message}</div> : null}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(145px, 1fr))", gap: 10, marginBottom: 16 }}>
        {[
          ["Enrolled", summary.total],
          ["Matched", summary.matched],
          ["Completed", summary.completed],
          ["In Progress", summary.inProgress],
          ["Not Started", summary.notStarted],
          ["No Record", summary.noRecord],
          ["Avg %", `${Math.round(summary.averagePercent)}%`],
        ].map(([label, value]) => (
          <div key={label} style={{ ...cardStyle, padding: 14, textAlign: "center" }}>
            <div style={{ color: "#666", fontSize: 12, fontWeight: "bold", textTransform: "uppercase" }}>{label}</div>
            <div style={{ fontSize: 24, fontWeight: 900, marginTop: 4 }}>{value}</div>
          </div>
        ))}
      </div>

      <div style={{ ...cardStyle, display: "grid", gap: 12 }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
          <div>
            <h2 style={{ margin: 0 }}>Enrolled CM Progress</h2>
            <p style={styles.message || { color: "#555", fontSize: 14 }}>
              Showing {visibleRows.length} of {mergedRows.length} enrolled CM(s). Emails are used only for matching and are not shown here.
            </p>
          </div>
          <button type="button" style={buttonStyle} onClick={downloadVisibleCsv}>
            Export visible CSV
          </button>
        </div>

        <input
          type="search"
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          placeholder="Search by name, ID, group, position, course, or status..."
          style={styles.searchInput || { width: "100%", padding: 11, borderRadius: 10, border: "1px solid #ccc" }}
        />

        <div style={styles.viewModeBox || { display: "flex", gap: 8, flexWrap: "wrap" }}>
          {filterButton("all", "All", mergedRows.length)}
          {filterButton("Completed", "Completed", summary.completed)}
          {filterButton("In Progress", "In Progress", summary.inProgress)}
          {filterButton("Not Started", "Not Started", summary.notStarted)}
          {filterButton("No Record", "No Record", summary.noRecord)}
        </div>

        <div style={{ overflowX: "auto", border: "1px solid #ddd", borderRadius: 14 }}>
          <table style={{ width: "100%", minWidth: 1120, borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#111", color: "#fff" }}>
                <th style={tableHeaderStyle}>Group</th>
                <th style={tableHeaderStyle}>Name</th>
                <th style={tableHeaderStyle}>ID</th>
                <th style={tableHeaderStyle}>Position</th>
                <th style={tableHeaderStyle}>Status</th>
                <th style={tableHeaderStyle}>Progress</th>
                <th style={tableHeaderStyle}>%</th>
                <th style={tableHeaderStyle}>Grade</th>
                <th style={tableHeaderStyle}>Last Engaged</th>
                <th style={tableHeaderStyle}>Course</th>
                <th style={tableHeaderStyle}>Admin</th>
              </tr>
            </thead>
            <tbody>
              {visibleRows.length ? (
                visibleRows.map((row) => (
                  <tr key={row.rosterKey} style={{ borderBottom: "1px solid #eee" }}>
                    <td style={tableCellStyle}>{row.group}</td>
                    <td style={{ ...tableCellStyle, fontWeight: "bold" }}>{row.name}</td>
                    <td style={tableCellStyle}>{row.id}</td>
                    <td style={tableCellStyle}>{row.position}</td>
                    <td style={tableCellStyle}>
                      <span style={{ display: "inline-block", padding: "5px 8px", borderRadius: 999, fontWeight: "bold", ...statusStyle(row.progressCategory) }}>
                        {row.progressCategory}
                      </span>
                      {row.progressStatus ? <div style={{ color: "#666", fontSize: 12, marginTop: 4 }}>{row.progressStatus}</div> : null}
                    </td>
                    <td style={tableCellStyle}>{row.progressText || "--"}</td>
                    <td style={tableCellStyle}>
                      <div style={{ fontWeight: 900 }}>{Math.round(Number(row.percentComplete || 0))}%</div>
                      <div style={{ height: 7, background: "#eee", borderRadius: 999, overflow: "hidden", marginTop: 5 }}>
                        <div style={{ width: `${Math.max(0, Math.min(100, Number(row.percentComplete || 0)))}%`, height: "100%", background: "#111" }} />
                      </div>
                    </td>
                    <td style={tableCellStyle}>{row.grade || "--"}</td>
                    <td style={tableCellStyle}>{row.lastEngaged || "--"}</td>
                    <td style={tableCellStyle}>{row.course || "--"}</td>
                    <td style={tableCellStyle}>
                      {row.adminUrl ? (
                        <a href={makeAdminUrl(row.adminUrl)} target="_blank" rel="noreferrer" style={{ color: "#0057b8", fontWeight: "bold" }}>
                          Open
                        </a>
                      ) : "--"}
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={11} style={{ ...tableCellStyle, textAlign: "center", color: "#777", padding: 22 }}>
                    Upload the Rouxbe progress CSV to see matched progress for the selected ship.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

const tableHeaderStyle = {
  padding: 10,
  textAlign: "left",
  borderRight: "1px solid #333",
  whiteSpace: "nowrap",
};

const tableCellStyle = {
  padding: 10,
  verticalAlign: "top",
  borderRight: "1px solid #eee",
};

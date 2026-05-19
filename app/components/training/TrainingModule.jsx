"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const todayMonthKey = () => new Date().toISOString().slice(0, 7);

const safeText = (value) => String(value || "").replace(/\s+/g, " ").trim();

const stationNameFromHeader = (value) =>
  safeText(value)
    .replace(/\s*-\s*\d+\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim();

const getCell = (row, index) => safeText(row?.[index]);

const looksLikeStationHeader = (row) => {
  const stationText = getCell(row, 1); // B
  const nameHeader = cleanText(row?.[3]); // D
  const actualPositionHeader = cleanText(row?.[5]); // F

  if (!stationText) return false;

  // Your file marks station rows yellow, but browser XLSX parsing does not reliably expose cell colors.
  // This format also has B = "Station - count", D = "Name", F = "Actual Position".
  if (nameHeader === "NAME" && actualPositionHeader.includes("ACTUAL POSITION")) return true;

  return /\s-\s*\d+\s*$/.test(stationText) && nameHeader === "NAME";
};

const parseStationAssignmentWorkbook = (workbook) => {
  const preferredSheet =
    workbook.SheetNames.find((name) => cleanText(name) === "MASTER PAGE") ||
    workbook.SheetNames.find((name) => cleanText(name).includes("MASTER")) ||
    workbook.SheetNames[0];

  const worksheet = workbook.Sheets[preferredSheet];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  const assignments = [];
  let currentStation = "";

  rows.forEach((row, index) => {
    const excelRow = index + 1;

    // Start reading from row 37 only.
    if (excelRow < 37) return;

    const columnB = getCell(row, 1); // B
    const columnD = getCell(row, 3); // D
    const columnF = getCell(row, 5); // F
    const columnI = getCell(row, 8); // I

    const columnDClean = cleanText(columnD);
    const columnFClean = cleanText(columnF);
    const columnIClean = cleanText(columnI);

    const isStationHeader =
      columnB &&
      (
        columnDClean === "NAME" ||
        columnFClean.includes("ACTUAL POSITION") ||
        columnIClean.includes("UNIQUE ID")
      );

    if (isStationHeader) {
      currentStation = stationNameFromHeader(columnB);
      return;
    }

    if (!currentStation) return;

    const position = columnB;
    const crewName = columnD;
    const actualPosition = columnF;
    const nationality = getCell(row, 6); // G
    const cabinNo = getCell(row, 7); // H
    const employeeNumber = columnI; // I

    if (!position || !crewName) return;
    if (cleanText(position).includes("TOTAL")) return;
    if (cleanText(crewName) === "NAME") return;
    if (cleanText(crewName).includes("#N/A")) return;
    if (cleanText(employeeNumber).includes("#N/A")) return;

    assignments.push({
      station: currentStation,
      position,
      crewName,
      actualPosition,
      nationality,
      cabinNo,
      employeeNumber,
      sourceRow: excelRow,
      sourceSheet: preferredSheet,
    });
  });

  return assignments;
};
const parseTrainingLinksWorkbook = (workbook) => {
  const sheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "LINKS") ||
    workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  return rows
    .map((row, index) => {
      const trainingName = getCell(row, 1); // B
      const trainingUrl = getCell(row, 2); // C
      const note = getCell(row, 3); // D

      return {
        trainingName,
        trainingUrl,
        note,
        sourceRow: index + 1,
        sortOrder: index,
      };
    })
    .filter((item) => item.trainingName && item.trainingUrl)
    .filter((item) => cleanText(item.trainingName) !== "NAME")
    .filter((item) => /^https?:\/\//i.test(item.trainingUrl));
};

const getCompletionKey = ({ ship, monthKey, station, crewName, trainingName }) =>
  [
    cleanText(ship),
    cleanText(monthKey),
    cleanText(station),
    cleanText(crewName),
    cleanText(trainingName),
  ].join("|");

export default function TrainingModule({
  styles,
  supabase,
  userShip,
  userEmail,
  isAdmin = false,
  onBack,
  logUsageEvent = () => {},
}) {
  const [monthKey, setMonthKey] = useState(todayMonthKey());
  const [assignments, setAssignments] = useState([]);
  const [trainingLinks, setTrainingLinks] = useState([]);
  const [completions, setCompletions] = useState([]);
  const [selectedStation, setSelectedStation] = useState("");
  const [selectedTraining, setSelectedTraining] = useState(null);
  const [stationSearch, setStationSearch] = useState("");
  const [trainingSearch, setTrainingSearch] = useState("");
  const [crewSearch, setCrewSearch] = useState("");
  const [message, setMessage] = useState("");
  const [loading, setLoading] = useState(false);

  const completionMap = useMemo(() => {
    const map = new Map();

    completions.forEach((item) => {
      map.set(
        getCompletionKey({
          ship: item.ship,
          monthKey: item.month_key,
          station: item.station,
          crewName: item.crew_name,
          trainingName: item.training_name,
        }),
        item
      );
    });

    return map;
  }, [completions]);

  const stationGroups = useMemo(() => {
    const grouped = new Map();

    assignments.forEach((item) => {
      if (!grouped.has(item.station)) grouped.set(item.station, []);
      grouped.get(item.station).push(item);
    });

    return Array.from(grouped.entries())
      .map(([station, crew]) => ({
        station,
        crew: crew.sort((a, b) => a.position.localeCompare(b.position) || a.crewName.localeCompare(b.crewName)),
      }))
      .filter((group) => {
        const query = stationSearch.toLowerCase().trim();
        if (!query) return true;
        return group.station.toLowerCase().includes(query);
      })
      .sort((a, b) => a.station.localeCompare(b.station));
  }, [assignments, stationSearch]);

  const selectedCrew = useMemo(() => {
    const query = crewSearch.toLowerCase().trim();

    return assignments
      .filter((item) => item.station === selectedStation)
      .filter((item) => {
        if (!query) return true;

        return [
          item.crewName,
          item.position,
          item.actualPosition,
          item.nationality,
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort((a, b) => a.position.localeCompare(b.position) || a.crewName.localeCompare(b.crewName));
  }, [assignments, selectedStation, crewSearch]);

  const filteredTrainingLinks = useMemo(() => {
    const query = trainingSearch.toLowerCase().trim();

    if (!query) return trainingLinks;

    return trainingLinks.filter((item) =>
      [item.trainingName, item.note, item.trainingUrl]
        .join(" ")
        .toLowerCase()
        .includes(query)
    );
  }, [trainingLinks, trainingSearch]);

  const getTrainingProgressForStation = (station, trainingName) => {
    const crew = assignments.filter((item) => item.station === station);

    if (!crew.length) return { completed: 0, total: 0, percent: 0 };

    const completed = crew.filter((person) =>
      completionMap.has(
        getCompletionKey({
          ship: userShip,
          monthKey,
          station,
          crewName: person.crewName,
          trainingName,
        })
      )
    ).length;

    return {
      completed,
      total: crew.length,
      percent: Math.round((completed / crew.length) * 100),
    };
  };

  const getStationOverallProgress = (station) => {
    const crew = assignments.filter((item) => item.station === station);
    const total = crew.length * trainingLinks.length;

    if (!total) return { completed: 0, total: 0, percent: 0 };

    let completed = 0;

    crew.forEach((person) => {
      trainingLinks.forEach((training) => {
        if (
          completionMap.has(
            getCompletionKey({
              ship: userShip,
              monthKey,
              station,
              crewName: person.crewName,
              trainingName: training.trainingName,
            })
          )
        ) {
          completed += 1;
        }
      });
    });

    return {
      completed,
      total,
      percent: Math.round((completed / total) * 100),
    };
  };

  const isCrewTrainingComplete = (person, training = selectedTraining) => {
    if (!training) return false;

    return completionMap.has(
      getCompletionKey({
        ship: userShip,
        monthKey,
        station: person.station,
        crewName: person.crewName,
        trainingName: training.trainingName,
      })
    );
  };

  const loadTrainingData = async () => {
    if (!supabase || !userShip) return;

    setLoading(true);

    try {
      const [assignmentResult, linksResult, completionResult] = await Promise.all([
        supabase
          .from("training_station_assignments")
          .select("*")
          .eq("ship", userShip)
          .eq("month_key", monthKey)
          .order("station", { ascending: true })
          .order("sort_order", { ascending: true }),

        supabase
          .from("training_links")
          .select("*")
          .eq("ship", userShip)
          .order("sort_order", { ascending: true }),

        supabase
          .from("training_completion_records")
          .select("*")
          .eq("ship", userShip)
          .eq("month_key", monthKey),
      ]);

      if (assignmentResult.error) throw assignmentResult.error;
      if (linksResult.error) throw linksResult.error;
      if (completionResult.error) throw completionResult.error;

      const nextAssignments = (assignmentResult.data || []).map((item) => ({
        station: item.station || "",
        position: item.position || "",
        crewName: item.crew_name || "",
        actualPosition: item.actual_position || "",
        nationality: item.nationality || "",
        sourceRow: Number(item.source_row || 0),
        sourceSheet: item.source_sheet || "",
      }));

      const nextLinks = (linksResult.data || []).map((item) => ({
        trainingName: item.training_name || "",
        trainingUrl: item.training_url || "",
        note: item.note || "",
        sourceRow: Number(item.source_row || 0),
        sortOrder: Number(item.sort_order || 0),
      }));

      setAssignments(nextAssignments);
      setTrainingLinks(nextLinks);
      setCompletions(completionResult.data || []);

      if (!selectedStation && nextAssignments.length) {
        setSelectedStation(nextAssignments[0].station);
      }

      setMessage(
        `Loaded ${nextAssignments.length} station assignment(s), ${nextLinks.length} training link(s), and ${(completionResult.data || []).length} completion record(s).`
      );
    } catch (error) {
      setMessage(error?.message || "Could not load training data.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadTrainingData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userShip, monthKey]);

  const replaceStationAssignments = async (rows, fileName) => {
    if (!supabase) {
      setAssignments(rows);
      setMessage("Supabase is not connected. Station assignments loaded locally only.");
      return;
    }

    setLoading(true);

    try {
      const { error: deleteError } = await supabase
        .from("training_station_assignments")
        .delete()
        .eq("ship", userShip)
        .eq("month_key", monthKey);

      if (deleteError) throw deleteError;

      const payload = rows.map((item, index) => ({
        ship: userShip,
        month_key: monthKey,
        station: item.station,
        position: item.position,
        crew_name: item.crewName,
        actual_position: item.actualPosition,
        nationality: item.nationality,
        source_sheet: item.sourceSheet,
        source_row: item.sourceRow,
        source_file: fileName || "",
        sort_order: index,
        updated_at: new Date().toISOString(),
      }));

      const batchSize = 200;

      for (let index = 0; index < payload.length; index += batchSize) {
        const { error } = await supabase
          .from("training_station_assignments")
          .insert(payload.slice(index, index + batchSize));

        if (error) throw error;
      }

      setAssignments(rows);
      setSelectedStation(rows[0]?.station || "");
      setMessage(`Station assignments saved. ${rows.length} crew assignment(s) found.`);

      logUsageEvent("training_station_assignments_uploaded", {
        module: "training",
        ship: userShip,
        monthKey,
        rows: rows.length,
        fileName,
      });
    } catch (error) {
      setMessage(error?.message || "Could not save station assignments.");
      window.alert(error?.message || "Could not save station assignments.");
    } finally {
      setLoading(false);
    }
  };

  const replaceTrainingLinks = async (rows, fileName) => {
    if (!supabase) {
      setTrainingLinks(rows);
      setMessage("Supabase is not connected. Training links loaded locally only.");
      return;
    }

    setLoading(true);

    try {
      const { error: deleteError } = await supabase
        .from("training_links")
        .delete()
        .eq("ship", userShip);

      if (deleteError) throw deleteError;

      const payload = rows.map((item, index) => ({
        ship: userShip,
        training_name: item.trainingName,
        training_url: item.trainingUrl,
        note: item.note || "",
        source_row: item.sourceRow,
        source_file: fileName || "",
        sort_order: index,
        updated_at: new Date().toISOString(),
      }));

      const batchSize = 200;

      for (let index = 0; index < payload.length; index += batchSize) {
        const { error } = await supabase
          .from("training_links")
          .insert(payload.slice(index, index + batchSize));

        if (error) throw error;
      }

      setTrainingLinks(rows);
      setSelectedTraining(rows[0] || null);
      setMessage(`Training links saved. ${rows.length} training link(s) found.`);

      logUsageEvent("training_links_uploaded", {
        module: "training",
        ship: userShip,
        rows: rows.length,
        fileName,
      });
    } catch (error) {
      setMessage(error?.message || "Could not save training links.");
      window.alert(error?.message || "Could not save training links.");
    } finally {
      setLoading(false);
    }
  };

  const uploadStationAssignmentFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setMessage("Reading station assignment file...");

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
      const rows = parseStationAssignmentWorkbook(workbook);

      if (!rows.length) {
        throw new Error("No station assignments found. Expected station headers in column B and crew names in column D.");
      }

      await replaceStationAssignments(rows, file.name);
    } catch (error) {
      setMessage(error?.message || "Could not read station assignment file.");
      window.alert(error?.message || "Could not read station assignment file.");
    } finally {
      event.target.value = "";
    }
  };

  const uploadTrainingLinksFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setMessage("Reading training links file...");

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
      const rows = parseTrainingLinksWorkbook(workbook);

      if (!rows.length) {
        throw new Error("No training links found. Expected training name in column B and link in column C.");
      }

      await replaceTrainingLinks(rows, file.name);
    } catch (error) {
      setMessage(error?.message || "Could not read training links file.");
      window.alert(error?.message || "Could not read training links file.");
    } finally {
      event.target.value = "";
    }
  };

  const markTrainingComplete = async (person, training = selectedTraining) => {
    if (!training) {
      window.alert("Choose a training first.");
      return;
    }

    if (!supabase) {
      window.alert("Supabase is not connected. Completion cannot be saved.");
      return;
    }

    const now = new Date().toISOString();

    const payload = {
      ship: userShip,
      month_key: monthKey,
      station: person.station,
      crew_name: person.crewName,
      position: person.position,
      actual_position: person.actualPosition,
      training_name: training.trainingName,
      training_url: training.trainingUrl,
      completed: true,
      completed_at: now,
      completed_by: userEmail || "",
      updated_at: now,
    };

    try {
      const { data, error } = await supabase
        .from("training_completion_records")
        .upsert(payload, {
          onConflict: "ship,month_key,station,crew_name,training_name",
        })
        .select("*")
        .single();

      if (error) throw error;

      setCompletions((prev) => {
        const key = getCompletionKey({
          ship: payload.ship,
          monthKey: payload.month_key,
          station: payload.station,
          crewName: payload.crew_name,
          trainingName: payload.training_name,
        });

        return [
          ...(data ? [data] : [payload]),
          ...prev.filter(
            (item) =>
              getCompletionKey({
                ship: item.ship,
                monthKey: item.month_key,
                station: item.station,
                crewName: item.crew_name,
                trainingName: item.training_name,
              }) !== key
          ),
        ];
      });

      setMessage(`${person.crewName} completed ${training.trainingName}.`);

      logUsageEvent("training_completed", {
        module: "training",
        ship: userShip,
        monthKey,
        station: person.station,
        crewName: person.crewName,
        trainingName: training.trainingName,
      });
    } catch (error) {
      setMessage(error?.message || "Could not save completion.");
      window.alert(error?.message || "Could not save completion.");
    }
  };

  const resetMonthlyCompletions = async () => {
    if (!isAdmin) return;

    if (!supabase) {
      window.alert("Supabase is not connected.");
      return;
    }

    const confirmed = window.confirm(
      `Reset all training completions for ${userShip} / ${monthKey}? This cannot be undone.`
    );

    if (!confirmed) return;

    setLoading(true);

    try {
      const { error } = await supabase
        .from("training_completion_records")
        .delete()
        .eq("ship", userShip)
        .eq("month_key", monthKey);

      if (error) throw error;

      setCompletions([]);
      setMessage(`Training completions reset for ${userShip} / ${monthKey}.`);
    } catch (error) {
      setMessage(error?.message || "Could not reset completions.");
      window.alert(error?.message || "Could not reset completions.");
    } finally {
      setLoading(false);
    }
  };

  const selectedStationProgress = selectedStation
    ? getStationOverallProgress(selectedStation)
    : { completed: 0, total: 0, percent: 0 };

  const selectedTrainingProgress =
    selectedStation && selectedTraining
      ? getTrainingProgressForStation(selectedStation, selectedTraining.trainingName)
      : { completed: 0, total: 0, percent: 0 };

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Modules
          </button>

          <div style={styles.shipBadge}>🎓 {userShip || "Ship"}</div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🎓 Training Module</h2>

          <label style={styles.label}>Training month</label>
          <input
            type="month"
            value={monthKey}
            onChange={(event) => {
              setMonthKey(event.target.value);
              setSelectedTraining(null);
            }}
            style={styles.searchInput}
          />

          <label style={styles.label}>Upload station assignment file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadStationAssignmentFile}
            style={styles.fileInput}
          />

          <label style={styles.label}>Upload training links file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadTrainingLinksFile}
            style={styles.fileInput}
          />

          <div style={styles.infoBox}>
            <div>🚢 Ship: <strong>{userShip || "Not selected"}</strong></div>
            <div>📅 Month: <strong>{monthKey}</strong></div>
            <div>📍 Stations: <strong>{stationGroups.length}</strong></div>
            <div>👥 Crew assignments: <strong>{assignments.length}</strong></div>
            <div>📚 Training links: <strong>{trainingLinks.length}</strong></div>
            <div>✅ Completion records: <strong>{completions.length}</strong></div>
            {loading && <div>Loading / saving...</div>}
          </div>

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={loadTrainingData} disabled={loading}>
              🔄 Refresh
            </button>

            {isAdmin && (
              <button style={styles.deleteButton} onClick={resetMonthlyCompletions} disabled={loading}>
                🧹 Reset Month
              </button>
            )}
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📍 Stations</h2>

          <input
            placeholder="Search station..."
            value={stationSearch}
            onChange={(event) => setStationSearch(event.target.value)}
            style={styles.searchInput}
          />

          {!stationGroups.length && (
            <p style={styles.emptyText}>
              Upload the station assignment file. Station headers are read from column B.
            </p>
          )}

          <div style={localStyles.stationGrid}>
            {stationGroups.map((group) => {
              const progress = getStationOverallProgress(group.station);

              return (
                <button
                  key={group.station}
                  style={{
                    ...localStyles.stationCard,
                    ...(selectedStation === group.station ? localStyles.stationCardActive : {}),
                  }}
                  onClick={() => {
                    setSelectedStation(group.station);
                    setSelectedTraining(null);
                  }}
                >
                  <strong>{group.station}</strong>
                  <span>{group.crew.length} crew member(s)</span>
                  <span>{progress.completed} / {progress.total} monthly completions</span>
                  <div style={localStyles.progressOuter}>
                    <div style={{ ...localStyles.progressInner, width: `${progress.percent}%` }} />
                  </div>
                  <span>{progress.percent}%</span>
                </button>
              );
            })}
          </div>
        </div>
      </section>

      {selectedStation && (
        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 16 }}>
            <div>
              <h2 style={styles.productTitle}>📚 {selectedStation} Trainings</h2>
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Select one training. Crew assigned to this station will appear below.
              </p>
            </div>

            <div style={styles.shipBadge}>
              {selectedStationProgress.completed} / {selectedStationProgress.total}
            </div>
          </div>

          <input
            placeholder="Search training..."
            value={trainingSearch}
            onChange={(event) => setTrainingSearch(event.target.value)}
            style={styles.searchInput}
          />

          {!trainingLinks.length && (
            <p style={styles.emptyText}>
              Upload the training links file. Training name is read from column B and link from column C.
            </p>
          )}

          <div style={localStyles.trainingGrid}>
            {filteredTrainingLinks.map((training) => {
              const progress = getTrainingProgressForStation(selectedStation, training.trainingName);

              return (
                <button
                  key={training.trainingName}
                  style={{
                    ...localStyles.trainingCard,
                    ...(selectedTraining?.trainingName === training.trainingName
                      ? localStyles.trainingCardActive
                      : {}),
                    ...(progress.total > 0 && progress.completed === progress.total
                      ? localStyles.trainingCardComplete
                      : {}),
                  }}
                  onClick={() => setSelectedTraining(training)}
                >
                  <strong>{training.trainingName}</strong>
                  {training.note && <span>{training.note}</span>}
                  <span>{progress.completed} / {progress.total} complete</span>
                  <div style={localStyles.progressOuter}>
                    <div style={{ ...localStyles.progressInner, width: `${progress.percent}%` }} />
                  </div>
                </button>
              );
            })}
          </div>
        </section>
      )}

      {selectedStation && selectedTraining && (
        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 16 }}>
            <div>
              <h2 style={styles.productTitle}>✅ {selectedTraining.trainingName}</h2>
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Crew assigned to {selectedStation}. Open the training link, then tap Done.
              </p>
            </div>

            <div style={styles.shipBadge}>
              {selectedTrainingProgress.completed} / {selectedTrainingProgress.total}
            </div>
          </div>

          <input
            placeholder="Search crew name, position, nationality..."
            value={crewSearch}
            onChange={(event) => setCrewSearch(event.target.value)}
            style={styles.searchInput}
          />

          <div style={styles.infoBox}>
            <div>📍 Station: <strong>{selectedStation}</strong></div>
            <div>📚 Training: <strong>{selectedTraining.trainingName}</strong></div>
            <div>🔗 Link: <strong>{selectedTraining.trainingUrl}</strong></div>
          </div>

          <div style={styles.equipmentGrid}>
            {selectedCrew.map((person) => {
              const complete = isCrewTrainingComplete(person, selectedTraining);

              return (
                <div
                  key={`${person.station}-${person.crewName}-${person.sourceRow}`}
                  style={{
                    ...styles.equipmentCard,
                    ...(complete ? styles.countedCard : {}),
                  }}
                >
                  <div style={styles.recipeName}>{person.crewName}</div>
                  <div style={styles.recipeMeta}>Position: {person.position || "N/A"}</div>
                  <div style={styles.recipeMeta}>Actual Position: {person.actualPosition || "N/A"}</div>
                  <div style={styles.recipeMeta}>Nationality: {person.nationality || "N/A"}</div>

                  {complete ? (
                    <div style={styles.statusGood}>Completed for {monthKey}</div>
                  ) : (
                    <div style={styles.statusWarning}>Pending for {monthKey}</div>
                  )}

                  <div style={styles.headerActions}>
                    <a
                      href={selectedTraining.trainingUrl}
                      target="_blank"
                      rel="noreferrer"
                      style={{
                        ...styles.backButton,
                        textDecoration: "none",
                        display: "inline-block",
                      }}
                    >
                      Open Training
                    </a>

                    <button
                      style={complete ? styles.backButton : styles.primaryButton}
                      onClick={() => markTrainingComplete(person, selectedTraining)}
                      disabled={complete}
                    >
                      {complete ? "Done" : "Mark Done"}
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      )}
    </main>
  );
}

const localStyles = {
  stationGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(160px, 1fr))",
    gap: 10,
  },
  stationCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 12,
    background: "#fafafa",
    display: "grid",
    gap: 6,
    textAlign: "left",
    cursor: "pointer",
  },
  stationCardActive: {
    border: "2px solid #111",
    background: "#f2f2f2",
  },
  trainingGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))",
    gap: 10,
  },
  trainingCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 12,
    background: "#fff",
    display: "grid",
    gap: 6,
    textAlign: "left",
    cursor: "pointer",
  },
  trainingCardActive: {
    border: "2px solid #111",
    background: "#f2f2f2",
  },
  trainingCardComplete: {
    border: "2px solid #2e7d32",
    background: "#e8f5e9",
  },
  progressOuter: {
    height: 8,
    borderRadius: 999,
    background: "#ddd",
    overflow: "hidden",
  },
  progressInner: {
    height: "100%",
    borderRadius: 999,
    background: "#2e7d32",
  },
};

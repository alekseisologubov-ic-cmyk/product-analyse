"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { useAppContext } from "../../context/AppContext";

const TRAINING_LINK_SCOPE = "GLOBAL";

const SHIP_LABELS = {
  SC: "Scarlet",
  VL: "Valiant",
  BRL: "Brilliant",
  RL: "Resilient",
};

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const safeText = (value) => String(value || "").replace(/\s+/g, " ").trim();

const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

const todayMonthKey = () => new Date().toISOString().slice(0, 7);

const getTrainingShipLabel = (shipCode) =>
  SHIP_LABELS[shipCode] || shipCode || "Not assigned";

const stationNameFromHeader = (value) =>
  safeText(value)
    .replace(/\s*-\s*\d+\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim();

const getCrewIdentity = ({ crewKey, employeeNumber, crewName }) =>
  cleanText(crewKey || employeeNumber || crewName || "");

const getCompletionKey = ({
  ship,
  monthKey,
  crewKey,
  employeeNumber,
  crewName,
  trainingName,
}) =>
  [
    cleanText(ship),
    cleanText(monthKey),
    getCrewIdentity({ crewKey, employeeNumber, crewName }),
    cleanText(trainingName),
  ].join("|");

const formatDateTime = (value) => {
  if (!value) return "N/A";

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);

  return date.toLocaleString();
};

const parseStationAssignmentWorkbook = (workbook) => {
  const candidateSheetNames = workbook.SheetNames.filter((name) => {
    const text = cleanText(name);

    return (
      text.includes("MASTER") ||
      text.includes("MANNING") ||
      text.includes("MANING") ||
      text.includes("CREW") ||
      text.includes("ROSTER")
    );
  });

  const sheetNamesToCheck = candidateSheetNames.length
    ? candidateSheetNames
    : workbook.SheetNames;

  const readCell = (worksheet, rowNumber, columnIndexZeroBased) => {
    const address = XLSX.utils.encode_cell({
      r: rowNumber - 1,
      c: columnIndexZeroBased,
    });

    return safeText(worksheet[address]?.v);
  };

  const looksLikeStationText = (value) => {
    const text = cleanText(value);
    const raw = String(value || "");

    if (!text) return false;

    return (
      /\s-\s*\d+\s*$/.test(raw) ||
      text.includes("CULINARY") ||
      text.includes("GALLEY") ||
      text.includes("KITCHEN") ||
      text.includes("BAKERY") ||
      text.includes("PASTRY") ||
      text.includes("BUTCHER") ||
      text.includes("FISH PREP") ||
      text.includes("VEG PREP") ||
      text.includes("POT WASH") ||
      text.includes("RESTAURANT") ||
      text.includes("BAR") ||
      text.includes("EXTRA VIRGIN") ||
      text.includes("PINK AGAVE") ||
      text.includes("RAZZLE") ||
      text.includes("WAKE") ||
      text.includes("GUNBAE") ||
      text.includes("DOCK") ||
      text.includes("SOCIAL") ||
      text.includes("PIZZA") ||
      text.includes("MANOR") ||
      text.includes("TEST KITCHEN") ||
      text.includes("SUN CLUB")
    );
  };

  const isStationHeaderRow = (worksheet, rowNumber) => {
    const columnB = readCell(worksheet, rowNumber, 1); // B
    const columnD = cleanText(readCell(worksheet, rowNumber, 3)); // D
    const columnF = cleanText(readCell(worksheet, rowNumber, 5)); // F
    const columnI = cleanText(readCell(worksheet, rowNumber, 8)); // I

    if (!columnB) return false;

    const hasCrewHeader =
      columnD === "NAME" ||
      columnD.includes("CREW NAME") ||
      columnF.includes("ACTUAL POSITION") ||
      columnI.includes("UNIQUE ID") ||
      columnI.includes("EMPLOYEE");

    return hasCrewHeader && looksLikeStationText(columnB);
  };

  const parseSheet = (sheetName) => {
    const worksheet = workbook.Sheets[sheetName];

    if (!worksheet || !worksheet["!ref"]) {
      return [];
    }

    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    const assignments = [];
    let currentStation = "";

    for (
      let rowNumber = range.s.r + 1;
      rowNumber <= range.e.r + 1;
      rowNumber += 1
    ) {
      const columnB = readCell(worksheet, rowNumber, 1); // B - station header or position
      const columnD = readCell(worksheet, rowNumber, 3); // D - crew name
      const columnF = readCell(worksheet, rowNumber, 5); // F - actual position
      const columnG = readCell(worksheet, rowNumber, 6); // G - nationality
      const columnH = readCell(worksheet, rowNumber, 7); // H - cabin no
      const columnI = readCell(worksheet, rowNumber, 8); // I - employee number / unique ID

      if (isStationHeaderRow(worksheet, rowNumber)) {
        currentStation = stationNameFromHeader(columnB);
        continue;
      }

      if (!currentStation) continue;

      const position = columnB;
      const crewName = columnD;
      const actualPosition = columnF;
      const nationality = columnG;
      const cabinNo = columnH;
      const employeeNumber = columnI;

      const positionClean = cleanText(position);
      const crewNameClean = cleanText(crewName);
      const employeeNumberClean = cleanText(employeeNumber);

      if (!position || !crewName) continue;
      if (positionClean.includes("TOTAL")) continue;
      if (positionClean === "POSITION") continue;
      if (crewNameClean === "NAME") continue;
      if (crewNameClean.includes("#N/A")) continue;
      if (employeeNumberClean.includes("#N/A")) continue;

      assignments.push({
        station: currentStation,
        position,
        crewName,
        actualPosition,
        nationality,
        cabinNo,
        employeeNumber,
        sourceRow: rowNumber,
        sourceSheet: sheetName,
      });
    }

    return assignments;
  };

  const allAssignments = [];
  const seen = new Set();

  sheetNamesToCheck.forEach((sheetName) => {
    parseSheet(sheetName).forEach((item) => {
      const personKey = cleanText(item.employeeNumber || item.crewName);
      const uniqueKey = [
        cleanText(item.station),
        cleanText(item.position),
        personKey,
      ].join("|");

      if (!personKey || seen.has(uniqueKey)) return;

      seen.add(uniqueKey);
      allAssignments.push(item);
    });
  });

  return allAssignments;
};

const parseTrainingLinksWorkbook = (workbook) => {
  const sheetName =
    workbook.SheetNames.find((name) => cleanText(name) === "LINKS") ||
    workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet || !worksheet["!ref"]) {
    return [];
  }

  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  const readCell = (rowNumber, columnIndexZeroBased) => {
    const address = XLSX.utils.encode_cell({
      r: rowNumber - 1,
      c: columnIndexZeroBased,
    });

    return safeText(worksheet[address]?.v);
  };

  const readCellLink = (rowNumber, columnIndexZeroBased) => {
    const address = XLSX.utils.encode_cell({
      r: rowNumber - 1,
      c: columnIndexZeroBased,
    });

    const cell = worksheet[address];

    return safeText(cell?.l?.Target) || safeText(cell?.l?.target) || safeText(cell?.v);
  };

  const links = [];

  for (let rowNumber = 1; rowNumber <= range.e.r + 1; rowNumber += 1) {
    const trainingName = readCell(rowNumber, 1); // B
    const linkFromColumnB = readCellLink(rowNumber, 1); // B hyperlink
    const linkFromColumnC = readCellLink(rowNumber, 2); // C link or hyperlink
    const note = readCell(rowNumber, 3); // D

    const trainingUrl =
      /^https?:\/\//i.test(linkFromColumnC)
        ? linkFromColumnC
        : /^https?:\/\//i.test(linkFromColumnB)
          ? linkFromColumnB
          : "";

    if (!trainingName || !trainingUrl) continue;
    if (cleanText(trainingName) === "NAME") continue;
    if (cleanText(trainingName).includes("TRAINING NAME")) continue;

    links.push({
      trainingName,
      trainingUrl,
      note,
      sourceRow: rowNumber,
      sortOrder: rowNumber,
    });
  }

  return links;
};

const exportRowsToExcel = (rows, sheetName, fileName) => {
  if (!rows.length) {
    window.alert("No rows to export.");
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  XLSX.writeFile(workbook, fileName);
};

const printRows = ({ title, rows, columns }) => {
  if (!rows.length) {
    window.alert("No rows to print.");
    return;
  }

  const headerHtml = columns
    .map((column) => `<th>${escapeHtml(column.label)}</th>`)
    .join("");

  const rowsHtml = rows
    .map(
      (row) =>
        "<tr>" +
        columns
          .map((column) => `<td>${escapeHtml(row[column.key] ?? "")}</td>`)
          .join("") +
        "</tr>"
    )
    .join("");

  const html = `
    <html>
      <head>
        <title>${escapeHtml(title)}</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 24px; }
          h1 { margin-bottom: 4px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
          th, td { border: 1px solid #ccc; padding: 7px; text-align: left; vertical-align: top; }
          th { background: #f2f2f2; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(title)}</h1>
        <table>
          <thead><tr>${headerHtml}</tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </body>
    </html>
  `;

  const printWindow = window.open("", "_blank");
  if (!printWindow) return;

  printWindow.document.write(html);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
};

export default function TrainingModule({ styles, onBack }) {
  const {
    supabase,
    userShip,
    shipDisplayName,
    userEmail,
    isAdmin,
    culinaryAdminShip,
    culinaryAdminShipDisplayName,
    canUploadStationAssignments,
    canReplaceTrainingLinks,
    canManageTrainingMonth,
    logUsageEvent,
  } = useAppContext();

  const currentShipLabel = shipDisplayName || getTrainingShipLabel(userShip);
  const adminShipLabel =
    culinaryAdminShipDisplayName || getTrainingShipLabel(culinaryAdminShip);

  const [monthKey, setMonthKey] = useState(todayMonthKey());
  const [assignments, setAssignments] = useState([]);
  const [trainingLinks, setTrainingLinks] = useState([]);
  const [completions, setCompletions] = useState([]);
  const [selectedStation, setSelectedStation] = useState("");
  const [selectedTraining, setSelectedTraining] = useState(null);
  const [trainingLaunchModal, setTrainingLaunchModal] = useState(null);
  const [trainingLaunchCrewSearch, setTrainingLaunchCrewSearch] = useState("");
  const [trainingLaunchCrew, setTrainingLaunchCrew] = useState(null);
  const [stationSearch, setStationSearch] = useState("");
  const [trainingSearch, setTrainingSearch] = useState("");
  const [crewSearch, setCrewSearch] = useState("");
  const [reportSearch, setReportSearch] = useState("");
  const [reportStationFilter, setReportStationFilter] = useState("ALL");
  const [reportTrainingFilter, setReportTrainingFilter] = useState("ALL");
  const [reportMode, setReportMode] = useState("station");
  const [message, setMessage] = useState("");
  const [loading, setLoading] = useState(false);

  const completionMap = useMemo(() => {
    const map = new Map();

    completions.forEach((item) => {
      map.set(
        getCompletionKey({
          ship: item.ship,
          monthKey: item.month_key,
          crewKey: item.crew_key,
          employeeNumber: item.employee_number,
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
        crew: crew.sort(
          (a, b) =>
            a.position.localeCompare(b.position) ||
            a.crewName.localeCompare(b.crewName)
        ),
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
          item.employeeNumber,
          item.position,
          item.actualPosition,
          item.nationality,
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort(
        (a, b) =>
          a.position.localeCompare(b.position) ||
          a.crewName.localeCompare(b.crewName)
      );
  }, [assignments, selectedStation, crewSearch]);

  const assignmentByCrewKey = useMemo(() => {
    const map = new Map();

    assignments.forEach((person) => {
      const key = getCrewIdentity({
        employeeNumber: person.employeeNumber,
        crewName: person.crewName,
      });

      if (key && !map.has(key)) {
        map.set(key, person);
      }
    });

    return map;
  }, [assignments]);

  const trainingLaunchCrewOptions = useMemo(() => {
    if (!selectedStation) return [];

    const query = trainingLaunchCrewSearch.toLowerCase().trim();

    return assignments
      .filter((item) => item.station === selectedStation)
      .filter((item) => {
        if (!query) return true;

        return [
          item.crewName,
          item.employeeNumber,
          item.position,
          item.actualPosition,
          item.nationality,
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort(
        (a, b) =>
          a.position.localeCompare(b.position) ||
          a.crewName.localeCompare(b.crewName)
      );
  }, [assignments, selectedStation, trainingLaunchCrewSearch]);

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
          employeeNumber: person.employeeNumber,
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
              employeeNumber: person.employeeNumber,
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
        employeeNumber: person.employeeNumber,
        crewName: person.crewName,
        trainingName: training.trainingName,
      })
    );
  };

  const allStationOptions = useMemo(
    () => [...new Set(assignments.map((item) => item.station).filter(Boolean))].sort(),
    [assignments]
  );

  const allTrainingOptions = useMemo(
    () => [...new Set(trainingLinks.map((item) => item.trainingName).filter(Boolean))].sort(),
    [trainingLinks]
  );

  const completionSummaryRows = useMemo(() => {
    const query = reportSearch.toLowerCase().trim();

    return completions
      .map((item, index) => {
        const crewIdentity = getCrewIdentity({
          crewKey: item.crew_key,
          employeeNumber: item.employee_number,
          crewName: item.crew_name,
        });

        const currentAssignment = assignmentByCrewKey.get(crewIdentity);

        return {
          Number: index + 1,
          Ship: item.ship || userShip || "",
          Month: item.month_key || monthKey,
          Station: currentAssignment?.station || item.station || "",
          Name: currentAssignment?.crewName || item.crew_name || "",
          EmployeeNumber:
            currentAssignment?.employeeNumber || item.employee_number || "",
          Position: currentAssignment?.position || item.position || "",
          ActualPosition:
            currentAssignment?.actualPosition || item.actual_position || "",
          Training: item.training_name || "",
          CompletedAt: formatDateTime(item.completed_at),
          CompletedBy: item.completed_by || "",
        };
      })
      .filter((row) => reportStationFilter === "ALL" || row.Station === reportStationFilter)
      .filter((row) => reportTrainingFilter === "ALL" || row.Training === reportTrainingFilter)
      .filter((row) => {
        if (!query) return true;

        return [
          row.Station,
          row.Name,
          row.EmployeeNumber,
          row.Position,
          row.ActualPosition,
          row.Training,
          row.CompletedAt,
          row.CompletedBy,
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort(
        (a, b) =>
          a.Station.localeCompare(b.Station) ||
          a.Name.localeCompare(b.Name) ||
          a.Training.localeCompare(b.Training)
      )
      .map((row, index) => ({ ...row, Number: index + 1 }));
  }, [
    completions,
    assignmentByCrewKey,
    userShip,
    monthKey,
    reportStationFilter,
    reportTrainingFilter,
    reportSearch,
  ]);

  const notYetCompletedRows = useMemo(() => {
    const query = reportSearch.toLowerCase().trim();
    const rows = [];

    assignments.forEach((person) => {
      trainingLinks.forEach((training) => {
        const completed = completionMap.has(
          getCompletionKey({
            ship: userShip,
            monthKey,
            employeeNumber: person.employeeNumber,
            crewName: person.crewName,
            trainingName: training.trainingName,
          })
        );

        if (completed) return;

        rows.push({
          Ship: userShip || "",
          Month: monthKey,
          Station: person.station || "",
          Name: person.crewName || "",
          EmployeeNumber: person.employeeNumber || "",
          Position: person.position || "",
          ActualPosition: person.actualPosition || "",
          Training: training.trainingName || "",
          TrainingUrl: training.trainingUrl || "",
          Status: "Not Yet Completed",
        });
      });
    });

    return rows
      .filter((row) => reportStationFilter === "ALL" || row.Station === reportStationFilter)
      .filter((row) => reportTrainingFilter === "ALL" || row.Training === reportTrainingFilter)
      .filter((row) => {
        if (!query) return true;

        return [
          row.Station,
          row.Name,
          row.EmployeeNumber,
          row.Position,
          row.ActualPosition,
          row.Training,
          row.Status,
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort(
        (a, b) =>
          a.Station.localeCompare(b.Station) ||
          a.Name.localeCompare(b.Name) ||
          a.Training.localeCompare(b.Training)
      )
      .map((row, index) => ({
        Number: index + 1,
        ...row,
      }));
  }, [
    assignments,
    trainingLinks,
    completionMap,
    userShip,
    monthKey,
    reportStationFilter,
    reportTrainingFilter,
    reportSearch,
  ]);

  const reportColumns =
    reportMode === "summary"
      ? [
          { key: "Number", label: "#" },
          { key: "Station", label: "Station" },
          { key: "Name", label: "Crew Name" },
          { key: "EmployeeNumber", label: "Employee #" },
          { key: "Position", label: "Position" },
          { key: "Training", label: "Training" },
          { key: "CompletedAt", label: "Completed Date / Time" },
          { key: "CompletedBy", label: "Completed By" },
        ]
      : [
          { key: "Number", label: "#" },
          { key: "Station", label: "Station" },
          { key: "Name", label: "Crew Name" },
          { key: "EmployeeNumber", label: "Employee #" },
          { key: "Position", label: "Position" },
          { key: "Training", label: "Training" },
          { key: "Status", label: "Status" },
        ];

  const visibleReportRows =
    reportMode === "summary" ? completionSummaryRows : notYetCompletedRows;

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
          .eq("ship", TRAINING_LINK_SCOPE)
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
        cabinNo: item.cabin_no || "",
        employeeNumber: item.employee_number || "",
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
    if (!canUploadStationAssignments) {
      throw new Error(
        "Only this ship's Culinary admin can replace station assignments."
      );
    }

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
        cabin_no: item.cabinNo || "",
        employee_number: item.employeeNumber || "",
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
    if (!canReplaceTrainingLinks) {
      throw new Error("Only global admins can replace the training links list.");
    }

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
        .eq("ship", TRAINING_LINK_SCOPE);

      if (deleteError) throw deleteError;

      const payload = rows.map((item, index) => ({
        ship: TRAINING_LINK_SCOPE,
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
    if (!canUploadStationAssignments) {
      const text =
        culinaryAdminShip && culinaryAdminShip !== userShip
          ? `This email is Culinary admin for ${adminShipLabel} only. Current selected ship is ${currentShipLabel}.`
          : "Only this ship's Culinary admin can upload the station assignment file.";

      setMessage(text);
      window.alert(text);
      event.target.value = "";
      return;
    }

    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setMessage("Reading station assignment file...");

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const rows = parseStationAssignmentWorkbook(workbook);

      if (!rows.length) {
        throw new Error(
          "No station assignments found. Checked sheets: " +
            workbook.SheetNames.join(", ") +
            ". Expected station header in column B with Name in column D, Actual Position in column F, or Unique ID / Employee number in column I."
        );
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
    if (!canReplaceTrainingLinks) {
      const text = "Only global admins can replace the permanent training links list.";
      setMessage(text);
      window.alert(text);
      event.target.value = "";
      return;
    }

    const file = event.target.files?.[0];
    if (!file) return;

    try {
      setMessage("Reading training links file...");

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, {
        type: "array",
        cellDates: true,
      });

      const rows = parseTrainingLinksWorkbook(workbook);

      if (!rows.length) {
        throw new Error(
          "No training links found. Expected training name in column B and link in column C, or hyperlink in column B."
        );
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

    const crewKey = getCrewIdentity({
      employeeNumber: person.employeeNumber,
      crewName: person.crewName,
    });

    const payload = {
      ship: userShip,
      month_key: monthKey,
      station: person.station,
      crew_key: crewKey,
      crew_name: person.crewName,
      employee_number: person.employeeNumber || "",
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
          onConflict: "ship,month_key,crew_key,training_name",
        })
        .select("*")
        .single();

      if (error) throw error;

      setCompletions((prev) => {
        const key = getCompletionKey({
          ship: payload.ship,
          monthKey: payload.month_key,
          crewKey: payload.crew_key,
          employeeNumber: payload.employee_number,
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
                crewKey: item.crew_key,
                employeeNumber: item.employee_number,
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
        employeeNumber: person.employeeNumber || "",
        trainingName: training.trainingName,
      });
    } catch (error) {
      setMessage(error?.message || "Could not save completion.");
      window.alert(error?.message || "Could not save completion.");
    }
  };

  const resetMonthlyCompletions = async () => {
    if (!canManageTrainingMonth) {
      window.alert("Only Culinary admins can reset monthly training completions.");
      return;
    }

    if (!supabase) {
      window.alert("Supabase is not connected.");
      return;
    }

    const confirmed = window.confirm(
      `Reset all training completions for ${currentShipLabel} / ${monthKey}? This cannot be undone.`
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
      setMessage(`Training completions reset for ${currentShipLabel} / ${monthKey}.`);
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

          <div style={styles.shipBadge}>🎓 {currentShipLabel || "Ship"}</div>
        </div>
      </header>

      <div style={styles.viewModeBox}>
        <button
          style={{
            ...styles.viewModeButton,
            ...(reportMode === "station" ? styles.viewModeButtonActive : {}),
          }}
          onClick={() => setReportMode("station")}
        >
          📍 Station Training
        </button>

        <button
          style={{
            ...styles.viewModeButton,
            ...(reportMode === "summary" ? styles.viewModeButtonActive : {}),
          }}
          onClick={() => setReportMode("summary")}
        >
          ✅ Summary Report ({completionSummaryRows.length})
        </button>

        <button
          style={{
            ...styles.viewModeButton,
            ...(reportMode === "notCompleted" ? styles.viewModeButtonActive : {}),
          }}
          onClick={() => setReportMode("notCompleted")}
        >
          ⏳ Not Yet Completed ({notYetCompletedRows.length})
        </button>
      </div>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🎓 Training Module</h2>

          <label style={styles.label}>Training month</label>
          <input
            type="month"
            value={monthKey}
            disabled={!canManageTrainingMonth}
            onChange={(event) => {
              setMonthKey(event.target.value);
              setSelectedTraining(null);
              setTrainingLaunchModal(null);
              setTrainingLaunchCrew(null);
              setTrainingLaunchCrewSearch("");
            }}
            style={styles.searchInput}
          />

          {!canManageTrainingMonth && (
            <div style={styles.recipeMeta}>
              Month changes are available only for Culinary admins.
            </div>
          )}

          {canUploadStationAssignments && (
            <>
              <label style={styles.label}>
                Upload station assignment file for {currentShipLabel}
              </label>
              <input
                type="file"
                accept=".xlsx,.xls,.xlsm"
                onChange={uploadStationAssignmentFile}
                style={styles.fileInput}
              />
            </>
          )}

          {canReplaceTrainingLinks && (
            <>
              <label style={styles.label}>
                Admin only: replace global training links file
              </label>
              <input
                type="file"
                accept=".xlsx,.xls,.xlsm"
                onChange={uploadTrainingLinksFile}
                style={styles.fileInput}
              />
            </>
          )}

          {!canUploadStationAssignments && (
            <div style={styles.infoBox}>
              <div>🔒 Station assignment upload is ship-specific.</div>
              <div>
                Your Culinary admin ship: <strong>{adminShipLabel}</strong>
              </div>
              <div>
                Current selected ship: <strong>{currentShipLabel}</strong>
              </div>
            </div>
          )}

          <div style={styles.infoBox}>
            <div>🚢 Ship: <strong>{currentShipLabel || "Not selected"}</strong></div>
            <div>📅 Month: <strong>{monthKey}</strong></div>
            <div>📍 Stations: <strong>{stationGroups.length}</strong></div>
            <div>👥 Crew assignments: <strong>{assignments.length}</strong></div>
            <div>📚 Training links: <strong>{trainingLinks.length}</strong></div>
            <div>✅ Completion records: <strong>{completions.length}</strong></div>
            <div>
              🔐 Station assignment upload:
              <strong>
                {canUploadStationAssignments
                  ? " Allowed for this ship"
                  : " Not allowed for this ship"}
              </strong>
            </div>
            <div>
              🔗 Global training links upload:
              <strong>{canReplaceTrainingLinks ? " Allowed" : " Not allowed"}</strong>
            </div>
            <div>
              👤 Culinary admin ship:
              <strong> {adminShipLabel}</strong>
            </div>
            {isAdmin && (
              <div>
                🛡️ Global admin: <strong>Yes</strong>
              </div>
            )}
            {loading && <div>Loading / saving...</div>}
          </div>

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={loadTrainingData} disabled={loading}>
              🔄 Refresh
            </button>

            {canManageTrainingMonth && (
              <button
                style={styles.deleteButton}
                onClick={resetMonthlyCompletions}
                disabled={loading}
              >
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
                    ...(selectedStation === group.station
                      ? localStyles.stationCardActive
                      : {}),
                  }}
                  onClick={() => {
                    setReportMode("station");
                    setSelectedStation(group.station);
                    setSelectedTraining(null);
                  }}
                >
                  <strong>{group.station}</strong>
                  <span>{group.crew.length} crew member(s)</span>
                  <span>
                    {progress.completed} / {progress.total} monthly completions
                  </span>
                  <div style={localStyles.progressOuter}>
                    <div
                      style={{
                        ...localStyles.progressInner,
                        width: `${progress.percent}%`,
                      }}
                    />
                  </div>
                  <span>{progress.percent}%</span>
                </button>
              );
            })}
          </div>
        </div>
      </section>

      {reportMode !== "station" && (
        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 16 }}>
            <div>
              <h2 style={styles.productTitle}>
                {reportMode === "summary"
                  ? "✅ Training Summary Report"
                  : "⏳ Not Yet Completed"}
              </h2>
              <p style={{ ...styles.emptyText, margin: 0 }}>
                {reportMode === "summary"
                  ? "Shows completed trainings with crew name, employee number, and completion date/time stamp."
                  : "Shows all crew members and trainings that have not been submitted yet."}
              </p>
            </div>

            <div style={styles.shipBadge}>{visibleReportRows.length} row(s)</div>
          </div>

          <section style={styles.grid}>
            <div>
              <label style={styles.label}>Station filter</label>
              <select
                value={reportStationFilter}
                onChange={(event) => setReportStationFilter(event.target.value)}
                style={styles.select}
              >
                <option value="ALL">All Stations</option>
                {allStationOptions.map((station) => (
                  <option key={station} value={station}>
                    {station}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label style={styles.label}>Training filter</label>
              <select
                value={reportTrainingFilter}
                onChange={(event) => setReportTrainingFilter(event.target.value)}
                style={styles.select}
              >
                <option value="ALL">All Trainings</option>
                {allTrainingOptions.map((training) => (
                  <option key={training} value={training}>
                    {training}
                  </option>
                ))}
              </select>
            </div>
          </section>

          <input
            placeholder="Search station, name, employee number, training..."
            value={reportSearch}
            onChange={(event) => setReportSearch(event.target.value)}
            style={styles.searchInput}
          />

          <div style={styles.headerActions}>
            <button
              style={styles.backButton}
              onClick={() =>
                printRows({
                  title:
                    reportMode === "summary"
                      ? "Training Summary Report"
                      : "Training Not Yet Completed",
                  rows: visibleReportRows,
                  columns: reportColumns,
                })
              }
            >
              🖨️ Print
            </button>

            <button
              style={styles.primaryButton}
              onClick={() =>
                exportRowsToExcel(
                  visibleReportRows,
                  reportMode === "summary" ? "Summary" : "Not Yet Completed",
                  reportMode === "summary"
                    ? `training-summary-${userShip || "ship"}-${monthKey}.xlsx`
                    : `training-not-yet-completed-${userShip || "ship"}-${monthKey}.xlsx`
                )
              }
            >
              📥 Export Excel
            </button>
          </div>

          {!visibleReportRows.length && (
            <p style={styles.emptyText}>
              {reportMode === "summary"
                ? "No completed trainings found for this filter."
                : "No not-yet-completed trainings found for this filter."}
            </p>
          )}

          <div style={styles.equipmentGrid}>
            {visibleReportRows.map((row) => (
              <div
                key={`${reportMode}-${row.Number}-${row.Station}-${row.Name}-${row.Training}`}
                style={{
                  ...styles.equipmentCard,
                  ...(reportMode === "summary"
                    ? styles.countedCard
                    : styles.zeroCountCard),
                }}
              >
                <div style={styles.recipeName}>{row.Name}</div>
                <div style={styles.recipeMeta}>
                  Employee #: {row.EmployeeNumber || "N/A"}
                </div>
                <div style={styles.recipeMeta}>Station: {row.Station || "N/A"}</div>
                <div style={styles.recipeMeta}>Position: {row.Position || "N/A"}</div>
                <div style={styles.recipeMeta}>Training: {row.Training || "N/A"}</div>

                {reportMode === "summary" ? (
                  <>
                    <div style={styles.statusGood}>Completed</div>
                    <div style={styles.recipeMeta}>Date / Time: {row.CompletedAt}</div>
                  </>
                ) : (
                  <div style={styles.statusWarning}>Not Yet Completed</div>
                )}
              </div>
            ))}
          </div>
        </section>
      )}

      {reportMode === "station" && selectedStation && (
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
              Training links are not loaded yet. Global admins can upload the training links file.
            </p>
          )}

          <div style={localStyles.trainingGrid}>
            {filteredTrainingLinks.map((training) => {
              const progress = getTrainingProgressForStation(
                selectedStation,
                training.trainingName
              );

              return (
                <button
                  key={training.trainingName}
                  type="button"
                  style={{
                    ...localStyles.trainingCard,
                    ...(selectedTraining?.trainingName === training.trainingName
                      ? localStyles.trainingCardActive
                      : {}),
                    ...(progress.total > 0 && progress.completed === progress.total
                      ? localStyles.trainingCardComplete
                      : {}),
                    color: "inherit",
                    fontFamily: "inherit",
                  }}
                  onClick={() => {
                    setSelectedTraining(training);
                    setTrainingLaunchModal(training);
                    setTrainingLaunchCrew(null);
                    setTrainingLaunchCrewSearch("");
                  }}
                >
                  <strong>{training.trainingName}</strong>
                  {training.note && <span>{training.note}</span>}
                  <span>
                    {progress.completed} / {progress.total} complete
                  </span>
                  <div style={localStyles.progressOuter}>
                    <div
                      style={{
                        ...localStyles.progressInner,
                        width: `${progress.percent}%`,
                      }}
                    />
                  </div>
                </button>
              );
            })}
          </div>
        </section>
      )}

      {reportMode === "station" && selectedStation && trainingLaunchModal && (
        <div
          style={styles.modalBackdrop}
          onClick={() => {
            setTrainingLaunchModal(null);
            setTrainingLaunchCrew(null);
            setTrainingLaunchCrewSearch("");
          }}
        >
          <div style={styles.modalCard} onClick={(event) => event.stopPropagation()}>
            <button
              type="button"
              style={styles.closeButton}
              onClick={() => {
                setTrainingLaunchModal(null);
                setTrainingLaunchCrew(null);
                setTrainingLaunchCrewSearch("");
              }}
            >
              ✕
            </button>

            <h2 style={styles.productTitle}>{trainingLaunchModal.trainingName}</h2>

            <div style={styles.infoBox}>
              <div>📍 Station: <strong>{selectedStation}</strong></div>
              <div>📚 Training: <strong>{trainingLaunchModal.trainingName}</strong></div>
              <div>👤 Choose the CM name first, then open the training link.</div>
            </div>

            <input
              placeholder="Search crew name, employee number, position..."
              value={trainingLaunchCrewSearch}
              onChange={(event) => setTrainingLaunchCrewSearch(event.target.value)}
              style={styles.searchInput}
            />

            <div style={localStyles.crewPickGrid}>
              {trainingLaunchCrewOptions.map((person) => {
                const complete = isCrewTrainingComplete(person, trainingLaunchModal);
                const selected =
                  trainingLaunchCrew?.crewName === person.crewName &&
                  trainingLaunchCrew?.employeeNumber === person.employeeNumber;

                return (
                  <button
                    key={`${person.station}-${person.crewName}-${person.employeeNumber}-${person.sourceRow}`}
                    type="button"
                    style={{
                      ...localStyles.crewPickCard,
                      ...(selected ? localStyles.crewPickCardActive : {}),
                      ...(complete ? styles.countedCard : {}),
                    }}
                    onClick={() => setTrainingLaunchCrew(person)}
                  >
                    <strong>{person.crewName}</strong>
                    <span>Employee #: {person.employeeNumber || "N/A"}</span>
                    <span>{person.position || "N/A"}</span>
                    <span>{complete ? "Completed" : "Pending"}</span>
                  </button>
                );
              })}
            </div>

            {!trainingLaunchCrew && (
              <p style={styles.warningText}>
                Select a crew member before opening the training link.
              </p>
            )}

            {trainingLaunchCrew && (
              <div style={styles.infoBox}>
                <div>Selected CM: <strong>{trainingLaunchCrew.crewName}</strong></div>
                <div>Employee #: <strong>{trainingLaunchCrew.employeeNumber || "N/A"}</strong></div>
                <div>Position: <strong>{trainingLaunchCrew.position || "N/A"}</strong></div>
              </div>
            )}

            <div style={styles.headerActions}>
              {trainingLaunchCrew ? (
                <a
                  href={trainingLaunchModal.trainingUrl}
                  target="_blank"
                  rel="noreferrer"
                  style={{
                    ...styles.primaryButton,
                    textDecoration: "none",
                    display: "inline-block",
                  }}
                  onClick={() => {
                    setSelectedTraining(trainingLaunchModal);
                    logUsageEvent("training_link_opened", {
                      module: "training",
                      ship: userShip,
                      monthKey,
                      station: selectedStation,
                      crewName: trainingLaunchCrew.crewName,
                      employeeNumber: trainingLaunchCrew.employeeNumber || "",
                      trainingName: trainingLaunchModal.trainingName,
                    });
                  }}
                >
                  Open Training Link
                </a>
              ) : (
                <button type="button" style={styles.backButton} disabled>
                  Open Training Link
                </button>
              )}

              <button
                type="button"
                style={styles.primaryButton}
                disabled={
                  !trainingLaunchCrew ||
                  isCrewTrainingComplete(trainingLaunchCrew, trainingLaunchModal)
                }
                onClick={() => markTrainingComplete(trainingLaunchCrew, trainingLaunchModal)}
              >
                {trainingLaunchCrew &&
                isCrewTrainingComplete(trainingLaunchCrew, trainingLaunchModal)
                  ? "Already Done"
                  : "Mark Done"}
              </button>

              <button
                type="button"
                style={styles.backButton}
                onClick={() => {
                  setTrainingLaunchModal(null);
                  setTrainingLaunchCrew(null);
                  setTrainingLaunchCrewSearch("");
                }}
              >
                Close / Return to CM Names
              </button>
            </div>
          </div>
        </div>
      )}

      {reportMode === "station" && selectedStation && selectedTraining && (
        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 16 }}>
            <div>
              <h2 style={styles.productTitle}>✅ {selectedTraining.trainingName}</h2>
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Crew assigned to {selectedStation}. Open the training link, then tap Mark Done.
              </p>
            </div>

            <div style={styles.shipBadge}>
              {selectedTrainingProgress.completed} / {selectedTrainingProgress.total}
            </div>
          </div>

          <input
            placeholder="Search crew name, employee number, position, nationality..."
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
                  <div style={styles.recipeMeta}>
                    Employee #: {person.employeeNumber || "N/A"}
                  </div>
                  <div style={styles.recipeMeta}>Position: {person.position || "N/A"}</div>
                  <div style={styles.recipeMeta}>
                    Actual Position: {person.actualPosition || "N/A"}
                  </div>
                  <div style={styles.recipeMeta}>
                    Nationality: {person.nationality || "N/A"}
                  </div>
                  <div style={styles.recipeMeta}>Cabin: {person.cabinNo || "N/A"}</div>

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
                      Open Training Link
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
  crewPickGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(180px, 1fr))",
    gap: 10,
    maxHeight: 340,
    overflowY: "auto",
    paddingRight: 4,
  },
  crewPickCard: {
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 10,
    background: "#fafafa",
    display: "grid",
    gap: 4,
    textAlign: "left",
    cursor: "pointer",
    fontFamily: "inherit",
  },
  crewPickCardActive: {
    border: "2px solid #111",
    background: "#f2f2f2",
  },
};

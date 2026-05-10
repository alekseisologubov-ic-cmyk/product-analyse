"use client";

import React, { useEffect, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

const SHIP_DISPLAY_NAMES = {
  BRL: "Brilliant Lady",
  RL: "Resilient Lady",
  SC: "Scarlet Lady",
  VL: "Valiant Lady",
};

const SCHEDULE_ALL_SHIPS = "ALL";

const SCHEDULE_TARGET_SHEETS = ["SEXC", "EXC_EXSC", "Pastry"];

const SCHEDULE_ROTATION_RULES = {
  SEXC: { contractMonths: 3, vacationWeeks: 6, label: "3 month contract / 6 week vacation" },
  EXC_EXSC: { contractMonths: 4, vacationMonths: 2, label: "4 month contract / 2 month rotation" },
  PASTRY: { contractMonths: 4, vacationMonths: 2, label: "4 month contract / 2 month rotation" },
};

const cleanText = (value) =>
  String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

const normalizeShipCode = (value) => {
  const text = cleanText(value).replace(/RESILIANT/g, "RESILIENT");

  if (!text) return "";
  if (text === "BRL" || text.includes("BRILLIANT")) return "BRL";
  if (text === "RL" || text.includes("RESILIENT")) return "RL";
  if (text === "SC" || text.includes("SCARLET")) return "SC";
  if (text === "VL" || text.includes("VALIANT")) return "VL";

  return "";
};

const getShipDisplayName = (shipCode) => SHIP_DISPLAY_NAMES[shipCode] || shipCode || "";

const getScheduleShipDisplayName = (shipCode) =>
  shipCode === SCHEDULE_ALL_SHIPS ? "All Ships" : getShipDisplayName(shipCode);

const getDefaultNextYearStartDate = () => {
  const nextYear = new Date().getFullYear() + 1;
  return `${nextYear}-01-01`;
};

export default function PeopleScheduleModule({ userShip, onBack, styles }) {
  const [scheduleShip, setScheduleShip] = useState(SCHEDULE_ALL_SHIPS);
  const [scheduleDate, setScheduleDate] = useState(getDefaultNextYearStartDate());
  const [scheduleSearch, setScheduleSearch] = useState("");
  const [scheduleMessage, setScheduleMessage] = useState("");
  const [scheduleRows, setScheduleRows] = useState([]);
  const [scheduleViewMode, setScheduleViewMode] = useState("timeline");
  const [scheduleCrewRows, setScheduleCrewRows] = useState([]);
  const [scheduleWorkbookInfo, setScheduleWorkbookInfo] = useState({ sheets: [], loadedAt: "" });

  useEffect(() => {
    try {
      const savedRows = localStorage.getItem("vv_schedule_rows");
      const savedCrewRows = localStorage.getItem("vv_schedule_crew_rows");
      const savedWorkbookInfo = localStorage.getItem("vv_schedule_workbook_info");

      if (savedRows) setScheduleRows(JSON.parse(savedRows));
      if (savedCrewRows) setScheduleCrewRows(JSON.parse(savedCrewRows));
      if (savedWorkbookInfo) setScheduleWorkbookInfo(JSON.parse(savedWorkbookInfo));
    } catch {
      setScheduleRows([]);
      setScheduleCrewRows([]);
      setScheduleWorkbookInfo({ sheets: [], loadedAt: "" });
    }
  }, []);

  useEffect(() => {
    localStorage.setItem("vv_schedule_rows", JSON.stringify(scheduleRows));
  }, [scheduleRows]);

  useEffect(() => {
    localStorage.setItem("vv_schedule_crew_rows", JSON.stringify(scheduleCrewRows));
  }, [scheduleCrewRows]);

  useEffect(() => {
    localStorage.setItem("vv_schedule_workbook_info", JSON.stringify(scheduleWorkbookInfo));
  }, [scheduleWorkbookInfo]);

  useEffect(() => {
    if (!scheduleCrewRows.length) return;

    const timer = window.setTimeout(() => {
      generateSchedule();
    }, 0);

    return () => window.clearTimeout(timer);
  }, [scheduleShip, scheduleDate, scheduleCrewRows.length]);

  const getScheduleRuleKey = (sheetName) => {
    const key = cleanText(sheetName).replace(/[^A-Z0-9]/g, "_");

    if (key === "SEXC") return "SEXC";
    if (key === "EXC_EXSC") return "EXC_EXSC";
    if (key === "PASTRY") return "PASTRY";

    return "EXC_EXSC";
  };

  const getScheduleRuleForSheet = (sheetName) => {
    const key = getScheduleRuleKey(sheetName);
    return SCHEDULE_ROTATION_RULES[key] || SCHEDULE_ROTATION_RULES.EXC_EXSC;
  };

  const parseExcelDate = (value) => {
    if (!value) return null;

    if (value instanceof Date && !isNaN(value.getTime())) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
    }

    const valueText = String(value || "").trim();
    if (!valueText) return null;

    const parsedDate = new Date(valueText);
    if (!isNaN(parsedDate.getTime())) {
      return new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
    }

    return null;
  };

  const formatDate = (date) => {
    const parsed = parseExcelDate(date);
    if (!parsed) return "";

    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const day = String(parsed.getDate()).padStart(2, "0");

    return `${year}-${month}-${day}`;
  };

  const addDays = (date, days) => {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  };

  const addWeeks = (date, weeks) => addDays(date, weeks * 7);

  const addMonths = (date, months) => {
    const result = new Date(date);
    const originalDay = result.getDate();

    result.setDate(1);
    result.setMonth(result.getMonth() + months);

    const lastDayOfTargetMonth = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
    result.setDate(Math.min(originalDay, lastDayOfTargetMonth));

    return result;
  };

  const getDateKey = (date) => {
    const parsed = parseExcelDate(date);
    if (!parsed) return 0;
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()).getTime();
  };

  const getInclusiveDayCount = (startDate, endDate) => {
    const startKey = getDateKey(startDate);
    const endKey = getDateKey(endDate);
    if (!startKey || !endKey || endKey < startKey) return 0;
    return Math.round((endKey - startKey) / (24 * 60 * 60 * 1000)) + 1;
  };

  const isScheduleMetadataRow = (idNumber, name) => {
    const text = cleanText(`${idNumber || ""} ${name || ""}`);

    return (
      !cleanText(name) ||
      text.includes("TODAY DATE") ||
      text.includes("UPDATED / DATE") ||
      text.includes("UPDATED DATE") ||
      text.includes("CONTRACT START") ||
      text.includes("CONTRACT END") ||
      text.includes("TOTAL MONTHS") ||
      text === "ID NAMES" ||
      text === "ID NAME"
    );
  };

  const isFilledCell = (cell) => {
    const fill = cell?.s?.fill;
    if (!fill) return false;

    const colors = [fill.fgColor?.rgb, fill.bgColor?.rgb]
      .filter(Boolean)
      .map((color) => String(color).replace(/^FF/i, "").toUpperCase());

    if (!colors.length && fill.patternType) return fill.patternType !== "none";

    return colors.some((color) => color && color !== "FFFFFF" && color !== "000000" && color !== "FFFFFF00");
  };

  const findHeaderDateForColumn = (rows, rowIndex, colIndex) => {
    for (let r = rowIndex - 1; r >= 0; r -= 1) {
      const possibleDate = parseExcelDate(rows[r]?.[colIndex]);
      if (possibleDate) return possibleDate;
    }

    return null;
  };

  const getHighlightedContractRange = (worksheet, rows, rowIndex) => {
    const startCol = XLSX.utils.decode_col("I");
    const endCol = XLSX.utils.decode_col("IP");
    const dates = [];

    for (let colIndex = startCol; colIndex <= endCol; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = worksheet[address];

      if (!isFilledCell(cell)) continue;

      const headerDate = findHeaderDateForColumn(rows, rowIndex, colIndex);
      const cellDate = parseExcelDate(cell?.v);
      const dateValue = headerDate || cellDate;

      if (dateValue) dates.push(dateValue);
    }

    if (!dates.length) return { start: null, end: null, count: 0 };

    dates.sort((a, b) => a.getTime() - b.getTime());

    return {
      start: dates[0],
      end: dates[dates.length - 1],
      count: dates.length,
    };
  };

  const inferCalendarRangeFromSignDates = (rows, signOnDate, signOffDate) => {
    if (!signOnDate || !signOffDate) return { start: null, end: null, count: 0 };

    const startCol = XLSX.utils.decode_col("I");
    const endCol = XLSX.utils.decode_col("IP");
    const signOnKey = getDateKey(signOnDate);
    const signOffKey = getDateKey(signOffDate);
    const dates = [];

    for (let colIndex = startCol; colIndex <= endCol; colIndex += 1) {
      const headerDate = findHeaderDateForColumn(rows, 6, colIndex) || parseExcelDate(rows[4]?.[colIndex]);
      if (!headerDate) continue;

      const headerKey = getDateKey(headerDate);
      if (headerKey >= signOnKey && headerKey <= signOffKey) dates.push(headerDate);
    }

    if (!dates.length) {
      return {
        start: signOnDate,
        end: signOffDate,
        count: getInclusiveDayCount(signOnDate, signOffDate),
      };
    }

    dates.sort((a, b) => a.getTime() - b.getTime());

    return {
      start: dates[0],
      end: dates[dates.length - 1],
      count: dates.length,
    };
  };

  const parseScheduleWorkbook = (workbook) => {
    const targetSheetKeys = new Set(SCHEDULE_TARGET_SHEETS.map((sheet) => cleanText(sheet)));
    const selectedSheets = workbook.SheetNames.filter((sheetName) => targetSheetKeys.has(cleanText(sheetName)));
    const crewRows = [];

    selectedSheets.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
      const rule = getScheduleRuleForSheet(sheetName);

      let currentShipCode = "";
      let currentShipName = "";
      let currentPosition = "";

      rows.forEach((row, rowIndex) => {
        const shipText = String(row[1] || "").trim();
        const positionText = String(row[2] || "").trim();
        const idNumberRaw = row[3];
        const idNumber = String(idNumberRaw || "").replace(/\.0$/, "").trim();
        const name = String(row[4] || "").trim();
        const signOnDate = parseExcelDate(row[5]);
        const signOffDate = parseExcelDate(row[6]);

        const detectedShipCode = normalizeShipCode(shipText);
        if (detectedShipCode) {
          currentShipCode = detectedShipCode;
          currentShipName = getShipDisplayName(detectedShipCode);
        }

        if (positionText && !cleanText(positionText).includes("POSITION")) {
          currentPosition = positionText;
        }

        // Excel rows 1-6 are title/header rows. Crew rows start after the D/E/F/G header.
        if (rowIndex < 6) return;
        if (isScheduleMetadataRow(idNumber, name)) return;
        if (!name || !currentShipCode) return;

        const highlightedRange = getHighlightedContractRange(worksheet, rows, rowIndex);
        const inferredCalendarRange = inferCalendarRangeFromSignDates(rows, signOnDate, signOffDate);
        const calendarRange = highlightedRange.count > 0 ? highlightedRange : inferredCalendarRange;
        const contractStart = signOnDate || calendarRange.start;
        const calculatedEnd = contractStart ? addDays(addMonths(contractStart, rule.contractMonths), -1) : null;
        const contractEnd = signOffDate || calendarRange.end || calculatedEnd;

        crewRows.push({
          id: `${sheetName}-${currentShipCode}-${currentPosition || "position"}-${idNumber || name}-${rowIndex}`,
          sheetName,
          shipCode: currentShipCode,
          shipName: currentShipName,
          position: currentPosition || "N/A",
          idNumber,
          name,
          signOnDate: formatDate(signOnDate),
          signOffDate: formatDate(signOffDate),
          highlightedStart: formatDate(calendarRange.start),
          highlightedEnd: formatDate(calendarRange.end),
          highlightedDays: calendarRange.count,
          contractStart: formatDate(contractStart),
          contractEnd: formatDate(contractEnd),
          contractMonths: rule.contractMonths,
          vacationWeeks: rule.vacationWeeks || 0,
          vacationMonths: rule.vacationMonths || 0,
          rotationRule: rule.label,
          source: !contractStart
            ? "Needs date review"
            : highlightedRange.count > 0
              ? "Highlighted cells I:IP"
              : "Columns F/G + I:IP calendar dates",
        });
      });
    });

    return {
      crewRows: crewRows.sort((a, b) =>
        a.shipCode.localeCompare(b.shipCode) ||
        a.sheetName.localeCompare(b.sheetName) ||
        a.position.localeCompare(b.position) ||
        a.name.localeCompare(b.name)
      ),
      selectedSheets,
    };
  };

  const generateScheduleRowsFromCrewRows = (crewRowsOverride, shipOverride = scheduleShip || SCHEDULE_ALL_SHIPS) => {
    const selectedShip = shipOverride || SCHEDULE_ALL_SHIPS;
    const planningStart = getPlanningStartDate();
    const planningEnd = getPlanningEndDate();
    const allRows = buildReplacementProjectedScheduleRows(crewRowsOverride, planningStart, planningEnd);
    const rows = selectedShip === SCHEDULE_ALL_SHIPS
      ? allRows
      : allRows.filter((row) => row.ship === selectedShip);

    return {
      rows,
      planningStart,
      planningEnd,
      selectedShip,
      replacementRows: rows.filter((row) => row.replacementFor).length,
      openSlots: rows.filter((row) => row.periodType === "Open Slot").length,
      missingDateRows: rows.filter((row) => row.periodType === "Missing Dates").length,
    };
  };

  const uploadScheduleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const workbook = XLSX.read(evt.target.result, {
          type: "array",
          cellDates: true,
          cellStyles: true,
        });

        const { crewRows, selectedSheets } = parseScheduleWorkbook(workbook);

        setScheduleCrewRows(crewRows);
        setScheduleWorkbookInfo({ sheets: selectedSheets, loadedAt: new Date().toLocaleString() });

        if (!selectedSheets.length) {
          setScheduleRows([]);
          setScheduleMessage("No matching tabs found. Expected tabs: SEXC, EXC_EXSC, Pastry.");
          return;
        }

        const generated = generateScheduleRowsFromCrewRows(crewRows, scheduleShip || SCHEDULE_ALL_SHIPS);
        setScheduleRows(generated.rows);

        setScheduleMessage(
          `Schedule file loaded and next-year schedule generated automatically. ` +
          `${crewRows.length} crew row(s) found from ${selectedSheets.join(", ")}. ` +
          `${generated.rows.length} projected period(s) for ${getScheduleShipDisplayName(generated.selectedShip)} from ${formatDate(generated.planningStart)} to ${formatDate(generated.planningEnd)}. ` +
          `${generated.replacementRows} replacement assignment(s), ${generated.openSlots} open slot(s), ${generated.missingDateRows} needs-date review row(s).`
        );
      } catch (error) {
        setScheduleMessage(`Could not read schedule file: ${error.message}`);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const getPlanningStartDate = () => parseExcelDate(scheduleDate) || new Date();

  const getPlanningEndDate = () => addDays(addMonths(getPlanningStartDate(), 12), -1);

  const getVacationEndDate = (vacationStart, crew) => {
    if (crew.vacationWeeks) return addDays(addWeeks(vacationStart, Number(crew.vacationWeeks || 0)), -1);
    return addDays(addMonths(vacationStart, Number(crew.vacationMonths || 0)), -1);
  };

  const periodOverlaps = (periodStart, periodEnd, planningStart, planningEnd) => {
    return periodStart <= planningEnd && periodEnd >= planningStart;
  };

  const getCrewRoleKey = (crew) => `${getScheduleRuleKey(crew.sheetName)}|${cleanText(crew.position || "")}`;

  const createScheduleRow = ({
    crew,
    periodType,
    startDate,
    endDate,
    shipCode,
    status,
    notes,
    previousShip = "",
    replacementFor = "",
    replacementForId = "",
  }) => ({
    id: `${crew.id}-${periodType}-${shipCode || crew.shipCode}-${formatDate(startDate) || "missing"}-${replacementForId || ""}`,
    ship: shipCode || crew.shipCode,
    shipName: getShipDisplayName(shipCode || crew.shipCode),
    sheetName: crew.sheetName,
    position: crew.position,
    idNumber: crew.idNumber,
    name: crew.name,
    periodType,
    startDate: formatDate(startDate),
    endDate: formatDate(endDate),
    rotationRule: crew.rotationRule,
    status,
    previousShip,
    replacementFor,
    notes,
  });

  const getCrewInitialContractDates = (crew) => {
    const contractMonths = Number(crew.contractMonths || 4);
    let contractStart = parseExcelDate(crew.contractStart || crew.signOnDate);
    let contractEnd = parseExcelDate(crew.contractEnd || crew.signOffDate);

    if (!contractStart) return { contractStart: null, contractEnd: null };

    if (!contractEnd || contractEnd < contractStart) {
      contractEnd = addDays(addMonths(contractStart, contractMonths), -1);
    }

    return { contractStart, contractEnd };
  };

  const getVacationPeriodAfterContract = (crew, contractEnd) => {
    if (!contractEnd) return { vacationStart: null, vacationEnd: null, availableDate: null };

    const vacationStart = addDays(contractEnd, 1);
    const vacationEnd = getVacationEndDate(vacationStart, crew);

    // The rotation sheet uses same-day handover logic: if vacation finishes on June 10,
    // the crew member can replace someone signing off on June 10.
    return { vacationStart, vacationEnd, availableDate: vacationEnd };
  };

  const pickReplacementCandidate = (crewStates, slot) => {
    const slotKey = getDateKey(slot.date);

    const candidates = Object.values(crewStates)
      .filter((state) => {
        const availableKey = getDateKey(state.availableDate);
        return (
          state.status === "available" &&
          state.roleKey === slot.roleKey &&
          state.crew.id !== slot.outgoingCrewId &&
          availableKey &&
          availableKey <= slotKey
        );
      })
      .sort((a, b) => {
        const aExact = getDateKey(a.availableDate) === slotKey ? 1 : 0;
        const bExact = getDateKey(b.availableDate) === slotKey ? 1 : 0;
        if (aExact !== bExact) return bExact - aExact;

        const aDifferentShip = a.lastShip !== slot.shipCode ? 1 : 0;
        const bDifferentShip = b.lastShip !== slot.shipCode ? 1 : 0;
        if (aDifferentShip !== bDifferentShip) return bDifferentShip - aDifferentShip;

        const aWaitingDays = Math.abs(slotKey - getDateKey(a.availableDate));
        const bWaitingDays = Math.abs(slotKey - getDateKey(b.availableDate));
        if (aWaitingDays !== bWaitingDays) return aWaitingDays - bWaitingDays;

        return a.crew.name.localeCompare(b.crew.name);
      });

    return candidates[0] || null;
  };

  const buildReplacementProjectedScheduleRows = (crewRows, planningStart, planningEnd) => {
    const rows = [];
    const crewStates = {};
    const openSlots = [];
    const processingEnd = addMonths(planningEnd, 4);

    crewRows.forEach((crew) => {
      const { contractStart, contractEnd } = getCrewInitialContractDates(crew);
      const roleKey = getCrewRoleKey(crew);

      crewStates[crew.id] = {
        crew,
        roleKey,
        status: "onboard",
        lastShip: crew.shipCode,
        availableDate: null,
        assignmentCount: 0,
      };

      if (!contractStart) {
        rows.push({
          id: `${crew.id}-missing-dates`,
          ship: crew.shipCode,
          shipName: crew.shipName,
          sheetName: crew.sheetName,
          position: crew.position,
          idNumber: crew.idNumber,
          name: crew.name,
          periodType: "Missing Dates",
          startDate: "",
          endDate: "",
          rotationRule: crew.rotationRule,
          status: "Missing sign-on date",
          previousShip: "",
          replacementFor: "",
          notes: "Check column F or highlighted contract cells between I and IP.",
        });
        return;
      }

      if (periodOverlaps(contractStart, contractEnd, planningStart, planningEnd)) {
        rows.push(createScheduleRow({
          crew,
          periodType: "Contract",
          startDate: contractStart,
          endDate: contractEnd,
          shipCode: crew.shipCode,
          status: "On board",
          previousShip: crew.shipCode,
          notes: "Current workbook contract",
        }));
      }

      openSlots.push({
        date: contractEnd,
        shipCode: crew.shipCode,
        roleKey,
        sheetName: crew.sheetName,
        position: crew.position,
        outgoingCrewId: crew.id,
        outgoingName: crew.name,
        outgoingIdNumber: crew.idNumber,
      });

      const { vacationStart, vacationEnd, availableDate } = getVacationPeriodAfterContract(crew, contractEnd);

      if (periodOverlaps(vacationStart, vacationEnd, planningStart, planningEnd)) {
        rows.push(createScheduleRow({
          crew,
          periodType: "Vacation",
          startDate: vacationStart,
          endDate: vacationEnd,
          shipCode: crew.shipCode,
          status: "Off board",
          previousShip: crew.shipCode,
          notes: crew.sheetName === "SEXC" ? "6 week vacation before replacement assignment" : "2 month rotation before replacement assignment",
        }));
      }

      crewStates[crew.id].status = "available";
      crewStates[crew.id].availableDate = availableDate;
    });

    const processedSlotKeys = new Set();
    let guard = 0;

    while (openSlots.length && guard < 2000) {
      openSlots.sort((a, b) => {
        const dateDiff = getDateKey(a.date) - getDateKey(b.date);
        if (dateDiff !== 0) return dateDiff;
        const roleDiff = a.roleKey.localeCompare(b.roleKey);
        if (roleDiff !== 0) return roleDiff;
        return a.shipCode.localeCompare(b.shipCode);
      });

      const slot = openSlots.shift();
      const slotDate = parseExcelDate(slot.date);
      const slotKey = `${formatDate(slotDate)}-${slot.shipCode}-${slot.roleKey}-${slot.outgoingCrewId}`;

      if (!slotDate || slotDate > processingEnd || processedSlotKeys.has(slotKey)) {
        guard += 1;
        continue;
      }

      processedSlotKeys.add(slotKey);

      const candidateState = pickReplacementCandidate(crewStates, slot);

      if (!candidateState) {
        if (periodOverlaps(slotDate, slotDate, planningStart, planningEnd)) {
          rows.push({
            id: `open-slot-${slotKey}`,
            ship: slot.shipCode,
            shipName: getShipDisplayName(slot.shipCode),
            sheetName: slot.sheetName,
            position: slot.position,
            idNumber: "",
            name: "Open Replacement Needed",
            periodType: "Open Slot",
            startDate: formatDate(slotDate),
            endDate: formatDate(slotDate),
            rotationRule: "Needs matching returning crew",
            status: "No returning crew available for this date/position",
            previousShip: "",
            replacementFor: `${slot.outgoingName}${slot.outgoingIdNumber ? ` (${slot.outgoingIdNumber})` : ""}`,
            notes: `Open slot on ${getShipDisplayName(slot.shipCode)} for ${slot.position}.`,
          });
        }

        guard += 1;
        continue;
      }

      const crew = candidateState.crew;
      const previousShip = candidateState.lastShip;
      const contractStart = slotDate;
      const contractEnd = addDays(addMonths(contractStart, Number(crew.contractMonths || 4)), -1);
      const replacementFor = `${slot.outgoingName}${slot.outgoingIdNumber ? ` (${slot.outgoingIdNumber})` : ""}`;

      if (periodOverlaps(contractStart, contractEnd, planningStart, planningEnd)) {
        rows.push(createScheduleRow({
          crew,
          periodType: "Contract",
          startDate: contractStart,
          endDate: contractEnd,
          shipCode: slot.shipCode,
          status: "On board - replacement assignment",
          previousShip,
          replacementFor,
          replacementForId: slot.outgoingCrewId,
          notes: `Replaces ${replacementFor} signing off ${formatDate(slotDate)}. Previous ship: ${getShipDisplayName(previousShip)}.`,
        }));
      }

      const { vacationStart, vacationEnd, availableDate } = getVacationPeriodAfterContract(crew, contractEnd);

      if (periodOverlaps(vacationStart, vacationEnd, planningStart, planningEnd)) {
        rows.push(createScheduleRow({
          crew,
          periodType: "Vacation",
          startDate: vacationStart,
          endDate: vacationEnd,
          shipCode: slot.shipCode,
          status: "Off board",
          previousShip: slot.shipCode,
          notes: crew.sheetName === "SEXC" ? "6 week vacation before next replacement assignment" : "2 month rotation before next replacement assignment",
        }));
      }

      candidateState.status = "available";
      candidateState.lastShip = slot.shipCode;
      candidateState.availableDate = availableDate;
      candidateState.assignmentCount += 1;

      openSlots.push({
        date: contractEnd,
        shipCode: slot.shipCode,
        roleKey: candidateState.roleKey,
        sheetName: crew.sheetName,
        position: crew.position,
        outgoingCrewId: crew.id,
        outgoingName: crew.name,
        outgoingIdNumber: crew.idNumber,
      });

      guard += 1;
    }

    return rows.sort((a, b) => {
      const dateA = parseExcelDate(a.startDate)?.getTime() || 0;
      const dateB = parseExcelDate(b.startDate)?.getTime() || 0;
      if (dateA !== dateB) return dateA - dateB;
      if (a.ship !== b.ship) return String(a.ship).localeCompare(String(b.ship));
      if (a.position !== b.position) return String(a.position).localeCompare(String(b.position));
      return String(a.name).localeCompare(String(b.name));
    });
  };

  const generateSchedule = () => {
    if (!scheduleCrewRows.length) {
      setScheduleMessage("Upload the schedule workbook first. Expected tabs: SEXC, EXC_EXSC, Pastry.");
      return;
    }

    const generated = generateScheduleRowsFromCrewRows(scheduleCrewRows, scheduleShip || SCHEDULE_ALL_SHIPS);

    if (!generated.rows.length) {
      setScheduleRows([]);
      setScheduleMessage(`No projected schedule rows found for ${getScheduleShipDisplayName(generated.selectedShip)} in the selected planning window.`);
      return;
    }

    setScheduleRows(generated.rows);
    setScheduleMessage(
      `Generated ${generated.rows.length} projected period(s) for ${getScheduleShipDisplayName(generated.selectedShip)} from ${formatDate(generated.planningStart)} to ${formatDate(generated.planningEnd)}. ` +
      `${generated.replacementRows} replacement assignment(s), ${generated.openSlots} open slot(s), ${generated.missingDateRows} needs-date review row(s).`
    );
  };

  const clearSchedule = () => {
    const confirmed = window.confirm("Clear the generated schedule?");
    if (!confirmed) return;

    setScheduleRows([]);
    setScheduleMessage("Schedule cleared.");
  };

  const clearScheduleWorkbook = () => {
    const confirmed = window.confirm("Clear uploaded schedule workbook data and generated schedule?");
    if (!confirmed) return;

    setScheduleCrewRows([]);
    setScheduleRows([]);
    setScheduleWorkbookInfo({ sheets: [], loadedAt: "" });
    setScheduleMessage("Schedule workbook data cleared.");
  };

  const getScheduleCrewRowsForSelectedShip = () => {
    const ship = scheduleShip || SCHEDULE_ALL_SHIPS;
    if (ship === SCHEDULE_ALL_SHIPS) return scheduleCrewRows;
    return scheduleCrewRows.filter((crew) => crew.shipCode === ship);
  };

  const getFilteredScheduleCrewRows = () => {
    const searchValue = scheduleSearch.toLowerCase();

    return getScheduleCrewRowsForSelectedShip().filter((crew) =>
      `${crew.shipCode} ${crew.shipName} ${crew.position} ${crew.idNumber} ${crew.name} ${crew.sheetName} ${crew.signOnDate} ${crew.signOffDate} ${crew.rotationRule}`
        .toLowerCase()
        .includes(searchValue)
    );
  };

  const getFilteredScheduleRows = () => {
    const searchValue = scheduleSearch.toLowerCase();

    return scheduleRows.filter((row) =>
      `${row.ship} ${row.shipName} ${row.position} ${row.idNumber} ${row.name} ${row.sheetName} ${row.periodType} ${row.startDate} ${row.endDate} ${row.status} ${row.previousShip} ${row.replacementFor} ${row.notes}`
        .toLowerCase()
        .includes(searchValue)
    );
  };

  const getScheduleTimelineMonths = () => {
    const start = getPlanningStartDate();
    const firstMonth = new Date(start.getFullYear(), start.getMonth(), 1);

    return Array.from({ length: 12 }, (_, index) => {
      const monthStart = addMonths(firstMonth, index);
      const monthEnd = addDays(addMonths(monthStart, 1), -1);

      return {
        start: monthStart,
        end: monthEnd,
        label: monthStart.toLocaleString("en-US", { month: "short", year: "2-digit" }),
      };
    });
  };

  const getScheduleDurationDays = (row) => {
    const start = parseExcelDate(row.startDate);
    const end = parseExcelDate(row.endDate);
    return getInclusiveDayCount(start, end);
  };

  const getScheduleDurationMonths = (row) => {
    const days = getScheduleDurationDays(row);
    return days ? (days / 30).toFixed(1) : "";
  };

  const scheduleRowOverlapsMonth = (row, month) => {
    const start = parseExcelDate(row.startDate);
    const end = parseExcelDate(row.endDate);
    if (!start || !end) return false;
    return start <= month.end && end >= month.start;
  };

  const getScheduleTimelineRows = () => {
    return getFilteredScheduleRows()
      .filter((row) => row.periodType !== "Vacation")
      .sort((a, b) => {
        const dateA = parseExcelDate(a.startDate)?.getTime() || 0;
        const dateB = parseExcelDate(b.startDate)?.getTime() || 0;
        if (String(a.ship) !== String(b.ship)) return String(a.ship).localeCompare(String(b.ship));
        if (String(a.position) !== String(b.position)) return String(a.position).localeCompare(String(b.position));
        if (dateA !== dateB) return dateA - dateB;
        return String(a.name).localeCompare(String(b.name));
      });
  };

  const exportScheduleToExcel = () => {
    if (!scheduleRows.length) return;

    const rows = scheduleRows.map((row) => ({
      ShipCode: row.ship,
      Ship: row.shipName,
      Position: row.position,
      Sheet: row.sheetName,
      IdNumber: row.idNumber,
      Name: row.name,
      PeriodType: row.periodType,
      StartDate: row.startDate,
      EndDate: row.endDate,
      Status: row.status,
      PreviousShip: getShipDisplayName(row.previousShip) || row.previousShip || "",
      ReplacementFor: row.replacementFor || "",
      RotationRule: row.rotationRule,
      Notes: row.notes,
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Year Schedule");
    XLSX.writeFile(wb, `people-schedule-${scheduleShip || SCHEDULE_ALL_SHIPS}-${formatDate(getPlanningStartDate()) || "year"}.xlsx`);
  };

  const printSchedule = () => {
    if (!scheduleRows.length) return;

    const planningStart = getPlanningStartDate();
    const planningEnd = getPlanningEndDate();

    const html = `
      <html>
        <head>
          <title>People Rotation Schedule</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background: #f2f2f2; }
            .contract { color: #2e7d32; font-weight: bold; }
            .vacation { color: #555; font-weight: bold; }
            .missing { color: #b00020; font-weight: bold; }
          </style>
        </head>
        <body>
          <h1>People Rotation Schedule</h1>
          <div><strong>Ship:</strong> ${getScheduleShipDisplayName(scheduleShip || SCHEDULE_ALL_SHIPS)}</div>
          <div><strong>Planning window:</strong> ${formatDate(planningStart)} to ${formatDate(planningEnd)}</div>
          <div><strong>Source tabs:</strong> ${scheduleWorkbookInfo.sheets.join(", ") || "N/A"}</div>
          <div><strong>Printed:</strong> ${new Date().toLocaleString()}</div>

          <table>
            <thead>
              <tr>
                <th>Ship</th>
                <th>Position</th>
                <th>ID</th>
                <th>Name</th>
                <th>Sheet</th>
                <th>Period</th>
                <th>Start</th>
                <th>End</th>
                <th>Status</th>
                <th>Previous Ship</th>
                <th>Replacement For</th>
                <th>Rule</th>
                <th>Notes</th>
              </tr>
            </thead>
            <tbody>
              ${scheduleRows
                .map(
                  (row) => `
                    <tr>
                      <td>${row.shipName || row.ship || ""}</td>
                      <td>${row.position || ""}</td>
                      <td>${row.idNumber || ""}</td>
                      <td>${row.name || ""}</td>
                      <td>${row.sheetName || ""}</td>
                      <td class="${row.periodType === "Contract" ? "contract" : row.periodType === "Vacation" ? "vacation" : "missing"}">${row.periodType}</td>
                      <td>${row.startDate || ""}</td>
                      <td>${row.endDate || ""}</td>
                      <td>${row.status || ""}</td>
                      <td>${getShipDisplayName(row.previousShip) || row.previousShip || ""}</td>
                      <td>${row.replacementFor || ""}</td>
                      <td>${row.rotationRule || ""}</td>
                      <td>${row.notes || ""}</td>
                    </tr>
                  `
                )
                .join("")}
            </tbody>
          </table>
        </body>
      </html>
    `;

    const printWindow = window.open("", "_blank");
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  };

    const selectedScheduleCrewRows = getScheduleCrewRowsForSelectedShip();
    const filteredScheduleCrewRows = getFilteredScheduleCrewRows();
    const filteredScheduleRows = getFilteredScheduleRows();
    const scheduleTimelineRows = getScheduleTimelineRows();
    const scheduleTimelineMonths = getScheduleTimelineMonths();
    const planningStart = getPlanningStartDate();
    const planningEnd = getPlanningEndDate();
    const contractRows = scheduleRows.filter((row) => row.periodType === "Contract");
    const vacationRows = scheduleRows.filter((row) => row.periodType === "Vacation");
    const openSlotRows = scheduleRows.filter((row) => row.periodType === "Open Slot");
    const replacementRows = scheduleRows.filter((row) => row.replacementFor);

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => onBack()}>← Modules</button>
            <div style={styles.shipBadge}>🚢 {getScheduleShipDisplayName(scheduleShip || SCHEDULE_ALL_SHIPS)}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>👥 People & Schedule</h2>

            <label style={styles.label}>Choose ship</label>
            <select value={scheduleShip} onChange={(e) => setScheduleShip(e.target.value)} style={styles.select}>
              <option value={SCHEDULE_ALL_SHIPS}>All Ships</option>
              {SHIPS.map((ship) => <option key={ship} value={ship}>{ship} - {getShipDisplayName(ship)}</option>)}
            </select>

            <label style={styles.label}>Upload schedule workbook</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadScheduleFile} style={styles.fileInput} />
            <p style={styles.emptyText}>Schedule is generated automatically after upload.</p>

            <label style={styles.label}>Planning start date</label>
            <input type="date" value={scheduleDate} onChange={(e) => setScheduleDate(e.target.value)} style={styles.searchInput} />

            <div style={styles.headerActions}>
              <button style={styles.primaryButton} onClick={generateSchedule}>
                ✨ Regenerate Schedule
              </button>
              <button style={styles.backButton} onClick={clearSchedule}>
                Clear Schedule
              </button>
              <button style={styles.deleteButton} onClick={clearScheduleWorkbook}>
                Clear File Data
              </button>
            </div>

            {scheduleMessage && <p style={styles.message}>{scheduleMessage}</p>}

            <div style={styles.infoBox}>
              <div>🚢 Schedule scope: <strong>{getScheduleShipDisplayName(scheduleShip || SCHEDULE_ALL_SHIPS)}</strong></div>
              <div>📅 Planning window: <strong>{formatDate(planningStart)} to {formatDate(planningEnd)}</strong></div>
              <div>📄 Tabs loaded: <strong>{scheduleWorkbookInfo.sheets.join(", ") || "None"}</strong></div>
              <div>👥 Total crew loaded: <strong>{scheduleCrewRows.length}</strong></div>
              <div>👥 Crew for selected ship: <strong>{selectedScheduleCrewRows.length}</strong></div>
              <div>✅ Contract periods generated: <strong>{contractRows.length}</strong></div>
              <div>🌴 Vacation periods generated: <strong>{vacationRows.length}</strong></div>
              <div>🔁 Replacement assignments: <strong>{replacementRows.length}</strong></div>
              <div>⚠️ Open replacement slots: <strong>{openSlotRows.length}</strong></div>
              {scheduleWorkbookInfo.loadedAt && <div>Loaded: <strong>{scheduleWorkbookInfo.loadedAt}</strong></div>}
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📌 Rotation Rules</h2>

            <div style={styles.infoBox}>
              <div><strong>Tabs used:</strong> SEXC, EXC_EXSC, Pastry</div>
              <div><strong>Columns:</strong> D = ID, E = Name, F = Sign on, G = Sign off</div>
              <div><strong>Calendar area:</strong> I to IP highlighted contract dates</div>
              <div><strong>Ship/position:</strong> B = Ship, C = Position</div>
              <div><strong>SEXC:</strong> 3 month contract / 6 week vacation</div>
              <div><strong>EXC_EXSC + Pastry:</strong> 4 month contract / 2 month rotation</div>
              <div><strong>Projection logic:</strong> returning crew replace matching position/sign-off slots across ships.</div>
              <div style={{ color: "#8a5a00" }}>
                Example: if someone finishes vacation on June 10, the app looks for another matching crew member signing off on June 10 and assigns the returning crew to that ship.
              </div>
            </div>

            <label style={styles.label}>Search crew or schedule</label>
            <input
              placeholder="Search ship, position, ID, name, tab, date, contract, vacation..."
              value={scheduleSearch}
              onChange={(e) => setScheduleSearch(e.target.value)}
              style={styles.searchInput}
            />
          </div>
        </section>

        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 20 }}>
            <h2 style={styles.productTitle}>👥 Crew From Workbook</h2>
          </div>

          {scheduleCrewRows.length === 0 && (
            <p style={styles.emptyText}>Upload the workbook with tabs SEXC, EXC_EXSC and Pastry to load crew information.</p>
          )}

          <div style={styles.equipmentGrid}>
            {filteredScheduleCrewRows.map((crew) => (
              <div key={crew.id} style={styles.equipmentCard}>
                <div style={styles.recipeName}>{crew.name}</div>
                <div style={styles.recipeMeta}>Ship: {crew.shipName || crew.shipCode || "N/A"}</div>
                <div style={styles.recipeMeta}>Position: {crew.position || "N/A"}</div>
                <div style={styles.recipeMeta}>ID: {crew.idNumber || "N/A"}</div>
                <div style={styles.recipeMeta}>Tab: {crew.sheetName}</div>
                <div style={styles.recipeMeta}>Sign on: {crew.signOnDate || "N/A"}</div>
                <div style={styles.recipeMeta}>Sign off: {crew.signOffDate || "N/A"}</div>
                {crew.highlightedStart && (
                  <div style={styles.recipeMeta}>Highlighted: {crew.highlightedStart} to {crew.highlightedEnd}</div>
                )}
                <div style={styles.statusNeutral}>{crew.source}</div>
                <div style={styles.statusGood}>{crew.rotationRule}</div>
              </div>
            ))}
          </div>
        </section>

        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 20 }}>
            <h2 style={styles.productTitle}>📅 Next Year Generated Schedule</h2>
            <div style={styles.headerActions}>
              <button
                style={{ ...styles.viewModeButton, ...(scheduleViewMode === "timeline" ? styles.viewModeButtonActive : {}) }}
                onClick={() => setScheduleViewMode("timeline")}
              >
                📊 Timeline
              </button>
              <button
                style={{ ...styles.viewModeButton, ...(scheduleViewMode === "cards" ? styles.viewModeButtonActive : {}) }}
                onClick={() => setScheduleViewMode("cards")}
              >
                🧾 Cards
              </button>
              <button style={styles.backButton} onClick={printSchedule}>🖨️ Print</button>
              <button style={styles.primaryButton} onClick={exportScheduleToExcel}>📥 Export Excel</button>
            </div>
          </div>

          {scheduleRows.length === 0 && (
            <p style={styles.emptyText}>Generated yearly rotation schedule will appear here.</p>
          )}

          {scheduleViewMode === "timeline" ? (
            <div style={styles.timelineScroll}>
              {scheduleTimelineRows.length === 0 ? (
                <p style={styles.emptyText}>No timeline rows found for the selected search/filter.</p>
              ) : (
                <div
                  style={{
                    ...styles.timelineGrid,
                    gridTemplateColumns: `150px 180px 90px 230px 90px 90px 70px 70px repeat(${scheduleTimelineMonths.length}, minmax(68px, 1fr))`,
                  }}
                >
                  <div style={styles.timelineHeaderCell}>Ship</div>
                  <div style={styles.timelineHeaderCell}>Position</div>
                  <div style={styles.timelineHeaderCell}>ID</div>
                  <div style={styles.timelineHeaderCell}>Name</div>
                  <div style={styles.timelineHeaderCell}>Sign On</div>
                  <div style={styles.timelineHeaderCell}>Sign Off</div>
                  <div style={styles.timelineHeaderCell}>Days</div>
                  <div style={styles.timelineHeaderCell}>Months</div>

                  {scheduleTimelineMonths.map((month) => (
                    <div key={month.label} style={styles.timelineHeaderCell}>{month.label}</div>
                  ))}

                  {scheduleTimelineRows.map((row) => (
                    <React.Fragment key={row.id}>
                      <div style={styles.timelineFixedCell}>{row.shipName || row.ship}</div>
                      <div style={styles.timelineFixedCell}>{row.position || "N/A"}</div>
                      <div style={styles.timelineFixedCell}>{row.idNumber || "N/A"}</div>
                      <div style={styles.timelineNameCell}>
                        <strong>{row.name}</strong>
                        {row.replacementFor && <span>Replacing: {row.replacementFor}</span>}
                      </div>
                      <div style={styles.timelineFixedCell}>{row.startDate || "N/A"}</div>
                      <div style={styles.timelineFixedCell}>{row.endDate || "N/A"}</div>
                      <div style={styles.timelineFixedCell}>{getScheduleDurationDays(row) || ""}</div>
                      <div style={styles.timelineFixedCell}>{getScheduleDurationMonths(row) || ""}</div>

                      {scheduleTimelineMonths.map((month) => {
                        const active = scheduleRowOverlapsMonth(row, month);
                        const activeStyle = row.periodType === "Open Slot"
                          ? styles.timelineCellOpen
                          : row.periodType === "Missing Dates"
                            ? styles.timelineCellMissing
                            : styles.timelineCellContract;

                        return (
                          <div
                            key={`${row.id}-${month.label}`}
                            style={{ ...styles.timelineCell, ...(active ? activeStyle : {}) }}
                            title={active ? `${row.periodType}: ${row.startDate} to ${row.endDate}` : ""}
                          />
                        );
                      })}
                    </React.Fragment>
                  ))}
                </div>
              )}
            </div>
          ) : (
            <div style={styles.equipmentGrid}>
              {filteredScheduleRows.map((row) => (
                <div
                  key={row.id}
                  style={{
                    ...styles.equipmentCard,
                    ...(row.periodType === "Contract" ? styles.countedCard : {}),
                    ...(row.periodType === "Open Slot" ? styles.zeroCountCard : {}),
                    ...(row.periodType === "Missing Dates" ? styles.orderWarningCard : {}),
                  }}
                >
                  <div style={styles.recipeName}>{row.name}</div>
                  <div style={styles.recipeMeta}>ID: {row.idNumber || "N/A"}</div>
                  <div style={styles.recipeMeta}>Ship: {row.shipName || row.ship}</div>
                  {row.previousShip && <div style={styles.recipeMeta}>Previous ship: {getShipDisplayName(row.previousShip) || row.previousShip}</div>}
                  {row.replacementFor && <div style={styles.recipeMeta}>Replacing: {row.replacementFor}</div>}
                  <div style={styles.recipeMeta}>Position: {row.position || "N/A"}</div>
                  <div style={styles.recipeMeta}>Tab: {row.sheetName}</div>
                  <div style={row.periodType === "Contract" ? styles.statusGood : row.periodType === "Vacation" ? styles.statusNeutral : row.periodType === "Open Slot" ? styles.statusWarning : styles.statusBad}>
                    {row.periodType}: {row.startDate || "N/A"} to {row.endDate || "N/A"}
                  </div>
                  <div style={styles.recipeMeta}>Status: {row.status}</div>
                  <div style={styles.recipeMeta}>Rule: {row.rotationRule}</div>
                  {row.notes && <div style={styles.recipeMeta}>Notes: {row.notes}</div>}
                </div>
              ))}
            </div>
          )}
        </section>
      </main>
    );
}

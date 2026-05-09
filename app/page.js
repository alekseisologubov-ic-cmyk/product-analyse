"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || "";
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || "";

const supabase =
  supabaseUrl && supabaseAnonKey
    ? createClient(supabaseUrl, supabaseAnonKey)
    : null;

const SHIPS = ["BRL", "RL", "SC", "VL"];

const SHIP_DISPLAY_NAMES = {
  BRL: "Brilliant Lady",
  RL: "Resilient Lady",
  SC: "Scarlet Lady",
  VL: "Valiant Lady",
};

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

const STATIONS = [
  "Galley",
  "Restaurant",
  "Bar",
  "Pantry",
  "Warehouse",
  "Dishwash",
  "Other",
];

const INVENTORY_USERS = [
  "Aleksei",
  "User 1",
  "User 2",
  "User 3",
  "Other",
];

const SCHEDULE_ROLES = [
  "Cook",
  "Steward",
  "Bar",
  "Utility",
  "Supervisor",
  "Host",
  "Runner",
  "Other",
];

const SCHEDULE_SHIFTS = [
  "Breakfast",
  "Lunch",
  "Dinner",
  "Overnight",
  "Custom",
];

const AVAILABILITY_OPTIONS = [
  "Preferred",
  "Available",
  "Emergency Only",
  "Unavailable",
];

const SCHEDULE_TARGET_SHEETS = ["SEXC", "EXC_EXSC", "Pastry"];

const SCHEDULE_ROTATION_RULES = {
  SEXC: { contractMonths: 3, vacationWeeks: 6, label: "3 month contract / 6 week vacation" },
  EXC_EXSC: { contractMonths: 4, vacationMonths: 2, label: "4 month contract / 2 month rotation" },
  PASTRY: { contractMonths: 4, vacationMonths: 2, label: "4 month contract / 2 month rotation" },
};

const ALLERGEN_RULES = [
  { allergen: "Tree Nuts", keywords: ["almond", "walnut", "pecan", "cashew", "hazelnut", "pistachio", "macadamia"] },
  { allergen: "Peanuts", keywords: ["peanut"] },
  { allergen: "Seeds", keywords: ["seed", "seeds", "sunflower seed", "pumpkin seed", "chia", "flax", "hemp seed"], exclude: ["seedless", "seedless cucumber"] },
  { allergen: "Soy", keywords: ["soy", "tofu", "edamame", "miso", "tamari"] },
  { allergen: "Gluten", keywords: ["wheat", "flour", "gluten", "bread", "pasta", "semolina", "barley", "rye", "panko"] },
  { allergen: "Milk / Dairy", keywords: ["milk", "cream", "butter", "cheese", "yogurt", "parmesan", "mozzarella", "ricotta", "cream cheese"] },
  { allergen: "Egg", keywords: ["egg", "eggs", "mayonnaise", "aioli"], exclude: ["eggplant"] },
  { allergen: "Fish", keywords: ["salmon", "tuna", "cod", "anchovy", "fish", "sardine"] },
  { allergen: "Shellfish", keywords: ["shrimp", "crab", "lobster", "mussel", "oyster", "scallop"], exclude: ["clam shell", "clamshell", "packed in a clam shell"] },
  { allergen: "Sesame", keywords: ["sesame", "tahini"] },
  { allergen: "Mustard", keywords: ["mustard"] },
];

const cleanText = (value) =>
  String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

const normalizeVenue = (value) =>
  cleanText(value)
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\s*-\s*VV$/g, "")
    .replace(/\s*VV$/g, "")
    .replace(/\bTHE\s+/g, "")
    .replace(/\bSCL\b/g, "")
    .replace(/\bVAL\b/g, "")
    .replace(/\bRES\b/g, "")
    .replace(/\bBRL\b/g, "")
    .replace(/\bROJO\b/g, "")
    .replace(/\bARIYA\b/g, "")
    .replace(/\bONLY\b/g, "")
    .replace(/\bMANNOR\b/g, "MANOR")
    .replace(/\s+/g, " ")
    .trim();

const formatQty = (value) => Number(value || 0).toFixed(2);

const getImageUrl = (url) => {
  const value = String(url || "").trim();
  if (!value) return "";

  if (value.includes("sharepoint.com") || value.includes("1drv.ms")) {
    return value.includes("?") ? `${value}&download=1` : `${value}?download=1`;
  }

  return value;
};

export default function App() {
  const [consumptionRows, setConsumptionRows] = useState([]);
  const [recipeRows, setRecipeRows] = useState([]);
  const [templateMap, setTemplateMap] = useState({});
  const [templateStatus, setTemplateStatus] = useState("Loading template...");
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [search, setSearch] = useState("");
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);
  const [message, setMessage] = useState("");
  const [viewMode, setViewMode] = useState("all");

  const [module, setModule] = useState("");
  const [equipmentMode, setEquipmentMode] = useState("");

  const [musterItems, setMusterItems] = useState([]);
  const [musterSearch, setMusterSearch] = useState("");
  const [musterMessage, setMusterMessage] = useState("");
  const [selectedEquipment, setSelectedEquipment] = useState(null);

  const [warehouseRows, setWarehouseRows] = useState([]);
  const [warehouseSearch, setWarehouseSearch] = useState("");
  const [warehouseMessage, setWarehouseMessage] = useState("");

  const [inUseRows, setInUseRows] = useState([]);
  const [inUseSearch, setInUseSearch] = useState("");
  const [inUseMessage, setInUseMessage] = useState("");

  const [makeInventoryItems, setMakeInventoryItems] = useState([]);
  const [makeInventorySearch, setMakeInventorySearch] = useState("");
  const [makeInventoryMessage, setMakeInventoryMessage] = useState("");
  const [makeInventoryShip, setMakeInventoryShip] = useState("");
  const [inventoryStation, setInventoryStation] = useState("");
  const [inventoryUserName, setInventoryUserName] = useState("");
  const [customInventoryUserName, setCustomInventoryUserName] = useState("");
  const [currentInventoryItem, setCurrentInventoryItem] = useState(null);
  const [inventoryQty, setInventoryQty] = useState("");
  const [editingInventoryId, setEditingInventoryId] = useState(null);
  const [inventorySummary, setInventorySummary] = useState([]);
  const [inventoryReportMode, setInventoryReportMode] = useState("my");
  const [inventoryLoading, setInventoryLoading] = useState(false);
  const [inventoryError, setInventoryError] = useState("");
  const [showVariance, setShowVariance] = useState(false);

  const [scheduleShip, setScheduleShip] = useState("");
  const [scheduleDate, setScheduleDate] = useState("");
  const [scheduleStation, setScheduleStation] = useState("");
  const [scheduleRole, setScheduleRole] = useState("");
  const [scheduleShiftName, setScheduleShiftName] = useState("Breakfast");
  const [scheduleStartTime, setScheduleStartTime] = useState("08:00");
  const [scheduleEndTime, setScheduleEndTime] = useState("16:00");
  const [scheduleRequiredPeople, setScheduleRequiredPeople] = useState("1");
  const [scheduleSearch, setScheduleSearch] = useState("");
  const [scheduleMessage, setScheduleMessage] = useState("");
  const [schedulePeople, setSchedulePeople] = useState([]);
  const [scheduleRows, setScheduleRows] = useState([]);
  const [scheduleCrewRows, setScheduleCrewRows] = useState([]);
  const [scheduleWorkbookInfo, setScheduleWorkbookInfo] = useState({ sheets: [], loadedAt: "" });
  const [personName, setPersonName] = useState("");
  const [personRole, setPersonRole] = useState("");
  const [personStation, setPersonStation] = useState("");
  const [personShip, setPersonShip] = useState("");
  const [personAvailability, setPersonAvailability] = useState("Available");
  const [personMaxHours, setPersonMaxHours] = useState("8");
  const [personNotes, setPersonNotes] = useState("");

  const shipColumns = { BRL: 8, RL: 11, SC: 14, VL: 17 };

  useEffect(() => {
    loadDefaultTemplate();
  }, []);

  useEffect(() => {
    setMakeInventoryShip(userShip);
    setScheduleShip(userShip);
    setPersonShip(userShip);
  }, [userShip]);

  useEffect(() => {
    try {
      const savedPeople = localStorage.getItem("vv_schedule_people");
      const savedRows = localStorage.getItem("vv_schedule_rows");
      const savedCrewRows = localStorage.getItem("vv_schedule_crew_rows");
      const savedWorkbookInfo = localStorage.getItem("vv_schedule_workbook_info");

      if (savedPeople) setSchedulePeople(JSON.parse(savedPeople));
      if (savedRows) setScheduleRows(JSON.parse(savedRows));
      if (savedCrewRows) setScheduleCrewRows(JSON.parse(savedCrewRows));
      if (savedWorkbookInfo) setScheduleWorkbookInfo(JSON.parse(savedWorkbookInfo));
    } catch {
      setSchedulePeople([]);
      setScheduleRows([]);
      setScheduleCrewRows([]);
      setScheduleWorkbookInfo({ sheets: [], loadedAt: "" });
    }
  }, []);

  useEffect(() => {
    localStorage.setItem("vv_schedule_people", JSON.stringify(schedulePeople));
  }, [schedulePeople]);

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
    if (module === "equipment" && equipmentMode === "makeinventory" && makeInventoryShip) {
      loadInventoryRecords(makeInventoryShip);
    }
  }, [module, equipmentMode, makeInventoryShip]);

  useEffect(() => {
    if (!supabase || !makeInventoryShip || module !== "equipment" || equipmentMode !== "makeinventory") return;

    const channel = supabase
      .channel(`inventory-counts-${makeInventoryShip}`)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "inventory_counts",
          filter: `ship=eq.${makeInventoryShip}`,
        },
        () => {
          loadInventoryRecords(makeInventoryShip);
        }
      )
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [makeInventoryShip, module, equipmentMode]);

  const visibleShips = viewMode === "single" ? [userShip] : SHIPS;

  const buildProductList = (rows) =>
    [...new Set(rows.slice(1).map((r) => String(r[6] || "").trim()).filter(Boolean))].sort();

  const readExcelFile = (file, callback) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: "binary" });
      callback(wb);
    };
    reader.readAsBinaryString(file);
  };

  const workbookToRows = (workbook) => {
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1 });
  };

  const loadDefaultTemplate = async () => {
    try {
      const response = await fetch("/template.xlsx");
      if (!response.ok) {
        setTemplateStatus("Template file not found.");
        return;
      }

      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      setTemplateMap(parseTemplateWorkbook(workbook));
      setTemplateStatus("Template loaded.");
    } catch {
      setTemplateStatus("Could not load template.");
    }
  };

  const parseTemplateWorkbook = (workbook) => {
    const map = {};

    workbook.SheetNames.forEach((sheetName) => {
      const venueKey = normalizeVenue(sheetName);
      if (!venueKey) return;

      if (!map[venueKey]) map[venueKey] = {};

      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      if (!rows.length) return;

      rows.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (cleanText(cell) !== "INGREDIENT NAME") return;

          const templateName =
            String(rows[rowIndex - 1]?.[colIndex] || rows[rowIndex - 1]?.[colIndex - 1] || sheetName || "Template").trim();

          rows.slice(rowIndex + 1).forEach((dataRow) => {
            const product = String(dataRow[colIndex] || "").trim();
            if (!product) return;

            const productKey = cleanText(product);
            if (!productKey) return;

            if (
              productKey === "INGREDIENT NAME" ||
              productKey === "CODE" ||
              productKey === "UM" ||
              productKey.includes("#REF")
            ) {
              return;
            }

            if (!map[venueKey][productKey]) {
              map[venueKey][productKey] = { product, templates: new Set() };
            }

            map[venueKey][productKey].templates.add(templateName);
          });
        });
      });
    });

    Object.keys(map).forEach((venueKey) => {
      Object.keys(map[venueKey]).forEach((productKey) => {
        map[venueKey][productKey].templates = [...map[venueKey][productKey].templates];
      });
    });

    return map;
  };

  const parseMusterWorkbook = (workbook) => {
    const items = [];

    workbook.SheetNames.forEach((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

      rows.slice(1).forEach((row) => {
        const category = String(row[2] || "").trim();
        const code = String(row[3] || "").trim();
        const name = String(row[4] || "").trim();
        const image = String(row[7] || "").trim();

        if (!category || !name) return;

        items.push({ sheetName, category, code, name, image });
      });
    });

    return items;
  };

  const parseWarehouseItems = () => {
    return warehouseRows
      .slice(1)
      .map((row) => {
        const code = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        const par = Number(row[6] || 0);
        const onHand = Number(row[7] || 0);
        const future = Number(row[12] || 0);
        const suggested = Math.max(par - onHand - future, 0);

        return { code, name, par, onHand, future, suggested };
      })
      .filter((item) => item.name || item.code)
      .filter((item) => `${item.code} ${item.name}`.toLowerCase().includes(warehouseSearch.toLowerCase()));
  };

  const parseInUseItems = () => {
    const actualMap = {};

    inUseRows.slice(1).forEach((row) => {
      const code = String(row[0] || "").trim();
      const name = String(row[1] || "").trim();
      const onHand = Number(row[7] || 0);

      if (!code && !name) return;

      actualMap[cleanText(code)] = { code, name, onHand };
    });

    return musterItems
      .map((item) => {
        const actual = actualMap[cleanText(item.code)];
        const onHand = actual ? actual.onHand : 0;

        let status = "Missing";
        if (actual && onHand > 0) status = "In Use";
        if (actual && onHand <= 0) status = "Zero Count";

        return {
          ...item,
          actualName: actual?.name || "",
          onHand,
          status,
        };
      })
      .filter((item) =>
        `${item.sheetName} ${item.category} ${item.code} ${item.name} ${item.status}`
          .toLowerCase()
          .includes(inUseSearch.toLowerCase())
      )
      .sort((a, b) => {
        const order = { Missing: 0, "Zero Count": 1, "In Use": 2 };
        return order[a.status] - order[b.status];
      });
  };

  const getFilteredMakeInventoryItems = () => {
    return makeInventoryItems.filter((item) =>
      `${item.sheetName} ${item.category} ${item.code} ${item.name}`
        .toLowerCase()
        .includes(makeInventorySearch.toLowerCase())
    );
  };

  const getEffectiveInventoryUserName = () => {
    if (inventoryUserName === "Other") return customInventoryUserName.trim();
    return inventoryUserName.trim();
  };

  const getInventoryItemKey = (item) => cleanText(item?.code || item?.name);

  const normalizeInventoryRecord = (record) => ({
    id: record.id,
    ship: record.ship,
    station: record.station,
    userName: record.user_name,
    itemKey: record.item_key,
    code: record.code || "",
    name: record.item_name || "",
    category: record.category || "",
    sheetName: record.sheet_name || "",
    image: record.image || "",
    qty: Number(record.qty || 0),
    confirmedAt: record.updated_at ? new Date(record.updated_at).toLocaleString() : "",
    updatedAt: record.updated_at || "",
  });

  const loadInventoryRecords = async (shipOverride) => {
    const ship = shipOverride || makeInventoryShip || userShip;

    if (!ship) {
      setInventorySummary([]);
      return;
    }

    if (!supabase) {
      setInventoryError("Supabase is not connected. Check NEXT_PUBLIC_SUPABASE_URL and NEXT_PUBLIC_SUPABASE_ANON_KEY.");
      return;
    }

    setInventoryLoading(true);
    setInventoryError("");

    const { data, error } = await supabase
      .from("inventory_counts")
      .select("*")
      .eq("ship", ship)
      .order("updated_at", { ascending: false });

    if (error) {
      setInventoryError(error.message);
      setInventoryLoading(false);
      return;
    }

    setInventorySummary((data || []).map(normalizeInventoryRecord));
    setInventoryLoading(false);
  };

  const getMyInventoryRows = () => {
    const ship = makeInventoryShip || userShip;
    const userName = getEffectiveInventoryUserName();

    return inventorySummary
      .filter(
        (item) =>
          item.ship === ship &&
          item.station === inventoryStation &&
          item.userName === userName
      )
      .sort((a, b) => a.name.localeCompare(b.name));
  };

  const getShipSummaryRows = () => {
    const ship = makeInventoryShip || userShip;
    const grouped = {};

    inventorySummary
      .filter((item) => item.ship === ship)
      .forEach((item) => {
        const key = item.itemKey || cleanText(item.code || item.name);

        if (!grouped[key]) {
          grouped[key] = {
            itemKey: key,
            ship: item.ship,
            code: item.code,
            name: item.name,
            category: item.category,
            sheetName: item.sheetName,
            totalQty: 0,
            stations: new Set(),
            users: new Set(),
            recordCount: 0,
            lastUpdated: "",
          };
        }

        grouped[key].totalQty += Number(item.qty || 0);
        grouped[key].recordCount += 1;

        if (item.station) grouped[key].stations.add(item.station);
        if (item.userName) grouped[key].users.add(item.userName);

        if (!grouped[key].lastUpdated || item.updatedAt > grouped[key].lastUpdated) {
          grouped[key].lastUpdated = item.updatedAt;
        }
      });

    return Object.values(grouped)
      .map((item) => ({
        ...item,
        stations: [...item.stations].sort(),
        users: [...item.users].sort(),
        confirmedAt: item.lastUpdated ? new Date(item.lastUpdated).toLocaleString() : "",
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  };

  const getVisibleInventoryReportRows = () => {
    return inventoryReportMode === "summary" ? getShipSummaryRows() : getMyInventoryRows();
  };

  const getMyInventoryStatusRows = () => {
    const countedMap = {};

    getMyInventoryRows().forEach((item) => {
      const key = item.itemKey || cleanText(item.code || item.name);
      countedMap[key] = item;
    });

    return makeInventoryItems
      .map((item) => {
        const key = getInventoryItemKey(item);
        const counted = countedMap[key];

        return {
          ...item,
          itemKey: key,
          countedQty: counted ? Number(counted.qty || 0) : 0,
          countedAt: counted?.confirmedAt || "",
          status: counted ? "Counted" : "Pending Count",
        };
      })
      .filter((item) =>
        `${item.sheetName} ${item.category} ${item.code} ${item.name} ${item.status}`
          .toLowerCase()
          .includes(makeInventorySearch.toLowerCase())
      );
  };

  const confirmInventoryQty = async () => {
    if (!currentInventoryItem) return;

    const ship = makeInventoryShip || userShip;
    const station = inventoryStation;
    const userName = getEffectiveInventoryUserName();

    if (!supabase) {
      setMakeInventoryMessage("Supabase is not connected. Check your environment variables.");
      return;
    }

    if (!ship || !station || !userName) {
      setMakeInventoryMessage("Choose ship, station, and user name before confirming inventory.");
      return;
    }

    const qty = Number(inventoryQty || 0);
    const itemKey = getInventoryItemKey(currentInventoryItem);

    const payload = {
      ship,
      station,
      user_name: userName,
      item_key: itemKey,
      code: currentInventoryItem.code || "",
      item_name: currentInventoryItem.name || "",
      category: currentInventoryItem.category || "",
      sheet_name: currentInventoryItem.sheetName || "",
      image: currentInventoryItem.image || "",
      qty,
      updated_at: new Date().toISOString(),
    };

    const { error } = await supabase
      .from("inventory_counts")
      .upsert(payload, {
        onConflict: "ship,station,user_name,item_key",
      });

    if (error) {
      setMakeInventoryMessage(`Could not save inventory: ${error.message}`);
      return;
    }

    await loadInventoryRecords(ship);

    setCurrentInventoryItem(null);
    setInventoryQty("");
    setEditingInventoryId(null);
    setMakeInventoryMessage("Inventory quantity saved.");
  };

  const editInventoryItem = (item) => {
    setCurrentInventoryItem({
      code: item.code,
      name: item.name,
      category: item.category,
      sheetName: item.sheetName,
      image: item.image,
    });

    setInventoryQty(String(item.qty ?? ""));
    setEditingInventoryId(item.id || null);

    if (item.station) setInventoryStation(item.station);

    if (item.userName) {
      if (INVENTORY_USERS.includes(item.userName)) {
        setInventoryUserName(item.userName);
        setCustomInventoryUserName("");
      } else {
        setInventoryUserName("Other");
        setCustomInventoryUserName(item.userName);
      }
    }
  };

  const deleteInventoryItem = async (itemToDelete) => {
    if (!supabase || !itemToDelete?.id) return;

    const confirmed = window.confirm(`Delete ${itemToDelete.name}?`);
    if (!confirmed) return;

    const { error } = await supabase
      .from("inventory_counts")
      .delete()
      .eq("id", itemToDelete.id);

    if (error) {
      setMakeInventoryMessage(`Could not delete item: ${error.message}`);
      return;
    }

    await loadInventoryRecords(makeInventoryShip || userShip);
  };

  const clearMyInventory = async () => {
    const ship = makeInventoryShip || userShip;
    const userName = getEffectiveInventoryUserName();

    if (!supabase || !ship || !inventoryStation || !userName) return;

    const confirmed = window.confirm(
      `Clear inventory for ${ship} / ${inventoryStation} / ${userName}?`
    );

    if (!confirmed) return;

    const { error } = await supabase
      .from("inventory_counts")
      .delete()
      .eq("ship", ship)
      .eq("station", inventoryStation)
      .eq("user_name", userName);

    if (error) {
      setMakeInventoryMessage(`Could not clear inventory: ${error.message}`);
      return;
    }

    await loadInventoryRecords(ship);

    setCurrentInventoryItem(null);
    setInventoryQty("");
    setEditingInventoryId(null);
  };

  const clearShipInventory = async () => {
    const ship = makeInventoryShip || userShip;

    if (!supabase || !ship) return;

    const confirmed = window.confirm(
      `Clear ALL inventory records for ${ship} from ALL users and ALL stations?`
    );

    if (!confirmed) return;

    const secondConfirm = window.confirm(
      `This cannot be undone. Confirm again to clear all ${ship} records.`
    );

    if (!secondConfirm) return;

    const { error } = await supabase
      .from("inventory_counts")
      .delete()
      .eq("ship", ship);

    if (error) {
      setMakeInventoryMessage(`Could not clear ship inventory: ${error.message}`);
      return;
    }

    await loadInventoryRecords(ship);

    setCurrentInventoryItem(null);
    setInventoryQty("");
    setEditingInventoryId(null);
  };

  const exportInventorySummaryToExcel = () => {
    const ship = makeInventoryShip || userShip;
    const rows = getVisibleInventoryReportRows();

    if (!rows.length) return;

    const exportRows =
      inventoryReportMode === "summary"
        ? rows.map((item) => ({
            Ship: ship,
            Code: item.code || "",
            Name: item.name || "",
            Category: item.category || "",
            Sheet: item.sheetName || "",
            TotalQuantity: item.totalQty,
            Stations: item.stations.join(", "),
            Users: item.users.join(", "),
            Records: item.recordCount,
            LastUpdated: item.confirmedAt,
          }))
        : rows.map((item) => ({
            Ship: item.ship,
            Station: item.station,
            User: item.userName,
            Code: item.code || "",
            Name: item.name || "",
            Category: item.category || "",
            Sheet: item.sheetName || "",
            Quantity: item.qty,
            Confirmed: item.confirmedAt,
          }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      wb,
      ws,
      inventoryReportMode === "summary" ? "Ship Summary" : "My Report"
    );

    XLSX.writeFile(wb, `inventory-${inventoryReportMode}-${ship || "ship"}.xlsx`);
  };

  const exportInventoryStatusToExcel = () => {
    const ship = makeInventoryShip || userShip;
    const rows = getMyInventoryStatusRows().map((item) => ({
      Ship: ship,
      Station: inventoryStation,
      User: getEffectiveInventoryUserName(),
      Code: item.code || "",
      Name: item.name || "",
      Category: item.category || "",
      Sheet: item.sheetName || "",
      Status: item.status,
      CountedQty: item.countedQty,
      CountedAt: item.countedAt,
    }));

    if (!rows.length) return;

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Inventory Status");
    XLSX.writeFile(wb, `inventory-status-${ship || "ship"}.xlsx`);
  };

  const printInventorySummary = () => {
    const ship = makeInventoryShip || userShip;
    const rows = getVisibleInventoryReportRows();

    if (!rows.length) return;

    const title =
      inventoryReportMode === "summary"
        ? "Ship Inventory Summary Report"
        : "My Inventory Report";

    const tableRows =
      inventoryReportMode === "summary"
        ? rows
            .map(
              (item) => `
                <tr>
                  <td>${item.code || ""}</td>
                  <td>${item.name || ""}</td>
                  <td>${item.category || ""}</td>
                  <td>${item.sheetName || ""}</td>
                  <td>${formatQty(item.totalQty)}</td>
                  <td>${item.stations.join(", ")}</td>
                  <td>${item.users.join(", ")}</td>
                  <td>${item.recordCount}</td>
                  <td>${item.confirmedAt}</td>
                </tr>
              `
            )
            .join("")
        : rows
            .map(
              (item) => `
                <tr>
                  <td>${item.station || ""}</td>
                  <td>${item.userName || ""}</td>
                  <td>${item.code || ""}</td>
                  <td>${item.name || ""}</td>
                  <td>${item.category || ""}</td>
                  <td>${item.sheetName || ""}</td>
                  <td>${formatQty(item.qty)}</td>
                  <td>${item.confirmedAt}</td>
                </tr>
              `
            )
            .join("");

    const tableHeader =
      inventoryReportMode === "summary"
        ? `
          <tr>
            <th>Code</th>
            <th>Name</th>
            <th>Category</th>
            <th>Sheet</th>
            <th>Total Qty</th>
            <th>Stations</th>
            <th>Users</th>
            <th>Records</th>
            <th>Last Updated</th>
          </tr>
        `
        : `
          <tr>
            <th>Station</th>
            <th>User</th>
            <th>Code</th>
            <th>Name</th>
            <th>Category</th>
            <th>Sheet</th>
            <th>Quantity</th>
            <th>Confirmed</th>
          </tr>
        `;

    const html = `
      <html>
        <head>
          <title>${title}</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background: #f2f2f2; }
          </style>
        </head>
        <body>
          <h1>${title}</h1>
          <div><strong>Ship:</strong> ${ship}</div>
          ${
            inventoryReportMode === "my"
              ? `<div><strong>Station:</strong> ${inventoryStation}</div>
                 <div><strong>User:</strong> ${getEffectiveInventoryUserName()}</div>`
              : `<div><strong>Report:</strong> All users and all stations for this ship</div>`
          }
          <div><strong>Printed:</strong> ${new Date().toLocaleString()}</div>

          <table>
            <thead>${tableHeader}</thead>
            <tbody>${tableRows}</tbody>
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

  const printInventoryStatus = () => {
    const ship = makeInventoryShip || userShip;
    const rows = getMyInventoryStatusRows();
    if (!rows.length) return;

    const html = `
      <html>
        <head>
          <title>Inventory Status</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
            th { background: #f2f2f2; }
            .counted { color: #2e7d32; font-weight: bold; }
            .pending { color: #555; font-weight: bold; }
          </style>
        </head>
        <body>
          <h1>Inventory Status</h1>
          <div><strong>Ship:</strong> ${ship}</div>
          <div><strong>Station:</strong> ${inventoryStation}</div>
          <div><strong>User:</strong> ${getEffectiveInventoryUserName()}</div>
          <div><strong>Printed:</strong> ${new Date().toLocaleString()}</div>

          <table>
            <thead>
              <tr>
                <th>Code</th>
                <th>Name</th>
                <th>Category</th>
                <th>Sheet</th>
                <th>Status</th>
                <th>Counted Qty</th>
                <th>Counted At</th>
              </tr>
            </thead>
            <tbody>
              ${rows
                .map(
                  (item) => `
                    <tr>
                      <td>${item.code || ""}</td>
                      <td>${item.name || ""}</td>
                      <td>${item.category || ""}</td>
                      <td>${item.sheetName || ""}</td>
                      <td class="${item.status === "Counted" ? "counted" : "pending"}">${item.status}</td>
                      <td>${formatQty(item.countedQty)}</td>
                      <td>${item.countedAt}</td>
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

  const uploadConsumptionFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      const rows = workbookToRows(workbook);
      setConsumptionRows(rows);
      setProducts(buildProductList(rows));
      setSelectedProduct("");
      setSelectedRecipe(null);
      setMessage("Consumption file loaded.");
    });
  };

  const uploadRecipeFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      setRecipeRows(workbookToRows(workbook));
      setSelectedRecipe(null);
      setMessage("Recipe / location file loaded.");
    });
  };

  const uploadTemplateFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      setTemplateMap(parseTemplateWorkbook(workbook));
      setTemplateStatus("Custom template loaded.");
    });
  };

  const uploadMusterFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      const items = parseMusterWorkbook(workbook);
      setMusterItems(items);
      setSelectedEquipment(null);
      setMusterMessage(`Equipment Muster List loaded from ${workbook.SheetNames.length} sheet(s).`);
    });
  };

  const uploadWarehouseFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      setWarehouseRows(workbookToRows(workbook));
      setWarehouseMessage("Warehouse inventory loaded.");
    });
  };

  const uploadInUseFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      setInUseRows(workbookToRows(workbook));
      setInUseMessage("Inventory in Use file loaded.");
    });
  };

  const uploadMakeInventoryFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      const items = parseMusterWorkbook(workbook);
      setMakeInventoryItems(items);
      setCurrentInventoryItem(null);
      setInventoryQty("");
      setEditingInventoryId(null);
      setShowVariance(false);
      setMakeInventoryMessage(`Master inventory loaded from ${workbook.SheetNames.length} sheet(s).`);
    });
  };

  const consumptionData = useMemo(() => consumptionRows.slice(1), [consumptionRows]);
  const recipeData = useMemo(() => recipeRows.slice(1), [recipeRows]);

  const productMatches = (selectedProductName, row) => {
    const selected = cleanText(selectedProductName);
    const assignedProduct = cleanText(row[12]);
    const productName = cleanText(row[7]);

    if (!selected) return false;
    if (assignedProduct === selected || productName === selected) return true;
    if (assignedProduct.length > 12 && (selected.includes(assignedProduct) || assignedProduct.includes(selected))) return true;
    if (productName.length > 12 && (selected.includes(productName) || productName.includes(selected))) return true;

    return false;
  };

  const templateHasProduct = (venueKey, product) => {
    const selected = cleanText(product);
    const venueTemplates = templateMap[venueKey] || {};

    return Object.entries(venueTemplates).some(([templateProductKey]) => {
      if (templateProductKey === selected) return true;
      if (templateProductKey.length > 12 && (selected.includes(templateProductKey) || templateProductKey.includes(selected))) return true;
      return false;
    });
  };

  const getTemplateMatches = (venueKey, product) => {
    const selected = cleanText(product);
    const venueTemplates = templateMap[venueKey] || {};
    const matches = [];

    Object.entries(venueTemplates).forEach(([templateProductKey, data]) => {
      const isMatch =
        templateProductKey === selected ||
        (templateProductKey.length > 12 && (selected.includes(templateProductKey) || templateProductKey.includes(selected)));

      if (isMatch) matches.push(...data.templates);
    });

    return [...new Set(matches)];
  };

  const getRequiredVenuesForProduct = (product) => {
    const required = {};

    recipeData.forEach((row) => {
      if (!productMatches(product, row)) return;

      const venueRaw = String(row[1] || "").trim();
      const venueKey = normalizeVenue(venueRaw);
      if (!venueKey) return;

      if (!required[venueKey]) {
        required[venueKey] = { displayName: venueRaw || venueKey, recipes: new Set() };
      }

      const recipeCode = String(row[15] || "").trim();
      const recipeName = String(row[16] || "").trim();

      if (recipeCode || recipeName) {
        required[venueKey].recipes.add(`${recipeCode || "N/A"} - ${recipeName || "Unnamed Recipe"}`);
      }
    });

    return required;
  };

  const getConsumptionBreakdown = (product) => {
    let currentVenue = "";
    const result = {};

    consumptionData.forEach((row) => {
      if (row[2]) currentVenue = String(row[2]).trim();

      const venue = currentVenue || "Unknown";
      const venueKey = normalizeVenue(venue);
      const productName = String(row[6] || "").trim();

      if (productName !== product) return;

      if (!result[venueKey]) {
        result[venueKey] = { displayName: venue, ships: {} };
      }

      SHIPS.forEach((ship) => {
        const qty = Number(row[shipColumns[ship]]) || 0;
        result[venueKey].ships[ship] = (result[venueKey].ships[ship] || 0) + qty;
      });
    });

    return result;
  };

  const getCombinedVenueBreakdown = (product) => {
    const actual = getConsumptionBreakdown(product);
    const required = getRequiredVenuesForProduct(product);
    const allVenueKeys = Array.from(new Set([...Object.keys(actual), ...Object.keys(required)])).sort();

    return allVenueKeys.map((venueKey) => {
      const actualVenue = actual[venueKey];
      const requiredVenue = required[venueKey];

      const ships = {};
      SHIPS.forEach((ship) => {
        ships[ship] = actualVenue?.ships?.[ship] || 0;
      });

      const templateMatches = getTemplateMatches(venueKey, product);
      const requiredByRecipe = Boolean(requiredVenue);
      const inTemplate = templateHasProduct(venueKey, product);
      const missingFromTemplate = requiredByRecipe && !inTemplate;

      return {
        venueKey,
        displayName: actualVenue?.displayName || requiredVenue?.displayName || venueKey,
        ships,
        required: requiredByRecipe,
        missingShips: visibleShips.filter((ship) => requiredByRecipe && (ships[ship] || 0) === 0),
        missingFromTemplate,
        templateMatches,
      };
    });
  };

  const getRecipesUsingProduct = (product) => {
    const recipes = {};

    recipeData.forEach((row) => {
      const recipeCode = String(row[15] || "").trim();
      const recipeName = String(row[16] || "").trim();
      const venue = String(row[1] || "").trim();

      if (!recipeCode && !recipeName) return;
      if (recipeName && !isNaN(Number(recipeName))) return;
      if (!productMatches(product, row)) return;

      const key = `${recipeCode || "N/A"} - ${recipeName || "Unnamed Recipe"}`;

      if (!recipes[key]) {
        recipes[key] = {
          key,
          recipeCode: recipeCode || "N/A",
          recipeName: recipeName || "Unnamed Recipe",
          venues: new Set(),
        };
      }

      if (venue) recipes[key].venues.add(venue);
    });

    return Object.values(recipes).map((recipe) => ({ ...recipe, venues: [...recipe.venues] }));
  };

  const getProductsInRecipe = (recipe) => {
    if (!recipe) return [];

    const items = {};

    recipeData.forEach((row) => {
      const recipeCode = String(row[15] || "").trim();
      const recipeName = String(row[16] || "").trim();

      if (recipeCode !== recipe.recipeCode || recipeName !== recipe.recipeName) return;

      const product = String(row[12] || row[7] || "").trim();
      if (!product) return;

      items[product] = true;
    });

    return Object.keys(items).sort();
  };

  const getSubRecipeIngredients = (subRecipeName) => {
    const items = {};
    const cleanSubRecipe = cleanText(subRecipeName);

    recipeData.forEach((row) => {
      const recipeName = cleanText(row[16]);
      const ingredient = String(row[12] || row[7] || "").trim();

      if (!ingredient) return;
      if (recipeName !== cleanSubRecipe) return;
      if (cleanText(ingredient) === cleanSubRecipe) return;

      items[ingredient] = true;
    });

    return Object.keys(items).sort();
  };

  const detectAllergens = (productsInRecipe) => {
    const found = {};

    const checkProductAgainstRules = (product, displayName) => {
      const lowerProduct = String(product || "").toLowerCase();

      ALLERGEN_RULES.forEach((rule) => {
        const isExcluded = rule.exclude?.some((word) => lowerProduct.includes(word));
        const matchedKeyword = !isExcluded && rule.keywords.find((keyword) => lowerProduct.includes(keyword));

        if (matchedKeyword) {
          if (!found[rule.allergen]) found[rule.allergen] = new Set();
          found[rule.allergen].add(displayName);
        }
      });
    };

    productsInRecipe.forEach((product) => {
      checkProductAgainstRules(product, product);

      const subIngredients = getSubRecipeIngredients(product);
      subIngredients.forEach((subItem) => {
        checkProductAgainstRules(subItem, `${product} → ${subItem}`);
      });
    });

    return Object.entries(found).map(([allergen, products]) => ({
      allergen,
      products: [...products].sort(),
    }));
  };

  const parseMusterItems = () => {
    const grouped = {};

    musterItems.forEach((item) => {
      const searchText = `${item.sheetName} ${item.category} ${item.code} ${item.name}`.toLowerCase();
      if (musterSearch && !searchText.includes(musterSearch.toLowerCase())) return;

      const groupKey = `${item.sheetName} / ${item.category}`;
      if (!grouped[groupKey]) grouped[groupKey] = [];
      grouped[groupKey].push(item);
    });

    return grouped;
  };

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
        if (!signOnDate && !signOffDate) return;

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
          source: highlightedRange.count > 0
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
        setScheduleRows([]);
        setScheduleWorkbookInfo({ sheets: selectedSheets, loadedAt: new Date().toLocaleString() });

        if (!selectedSheets.length) {
          setScheduleMessage("No matching tabs found. Expected tabs: SEXC, EXC_EXSC, Pastry.");
          return;
        }

        setScheduleMessage(`Schedule file loaded. ${crewRows.length} crew row(s) found from ${selectedSheets.join(", ")}.`);
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

  const buildCrewRotationRows = (crew, planningStart, planningEnd) => {
    const rows = [];
    let contractStart = parseExcelDate(crew.contractStart || crew.signOnDate);
    let contractEnd = parseExcelDate(crew.contractEnd || crew.signOffDate);
    const contractMonths = Number(crew.contractMonths || 4);

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
        notes: "Check column F or highlighted contract cells between I and IP.",
      });
      return rows;
    }

    if (!contractEnd || contractEnd < contractStart) {
      contractEnd = addDays(addMonths(contractStart, contractMonths), -1);
    }

    let guard = 0;
    while (contractEnd < planningStart && guard < 30) {
      const vacationStart = addDays(contractEnd, 1);
      const vacationEnd = getVacationEndDate(vacationStart, crew);
      contractStart = addDays(vacationEnd, 1);
      contractEnd = addDays(addMonths(contractStart, contractMonths), -1);
      guard += 1;
    }

    guard = 0;
    while (contractStart <= planningEnd && guard < 30) {
      if (periodOverlaps(contractStart, contractEnd, planningStart, planningEnd)) {
        rows.push({
          id: `${crew.id}-contract-${formatDate(contractStart)}`,
          ship: crew.shipCode,
          shipName: crew.shipName,
          sheetName: crew.sheetName,
          position: crew.position,
          idNumber: crew.idNumber,
          name: crew.name,
          periodType: "Contract",
          startDate: formatDate(contractStart),
          endDate: formatDate(contractEnd),
          rotationRule: crew.rotationRule,
          status: "On board",
          notes: crew.source,
        });
      }

      const vacationStart = addDays(contractEnd, 1);
      const vacationEnd = getVacationEndDate(vacationStart, crew);

      if (periodOverlaps(vacationStart, vacationEnd, planningStart, planningEnd)) {
        rows.push({
          id: `${crew.id}-vacation-${formatDate(vacationStart)}`,
          ship: crew.shipCode,
          shipName: crew.shipName,
          sheetName: crew.sheetName,
          position: crew.position,
          idNumber: crew.idNumber,
          name: crew.name,
          periodType: "Vacation",
          startDate: formatDate(vacationStart),
          endDate: formatDate(vacationEnd),
          rotationRule: crew.rotationRule,
          status: "Off board",
          notes: crew.sheetName === "SEXC" ? "6 week vacation" : "2 month rotation",
        });
      }

      contractStart = addDays(vacationEnd, 1);
      contractEnd = addDays(addMonths(contractStart, contractMonths), -1);
      guard += 1;
    }

    return rows;
  };

  const generateSchedule = () => {
    const ship = scheduleShip || userShip;

    if (!ship) {
      setScheduleMessage("Choose ship before generating schedule.");
      return;
    }

    if (!scheduleCrewRows.length) {
      setScheduleMessage("Upload the schedule workbook first. Expected tabs: SEXC, EXC_EXSC, Pastry.");
      return;
    }

    const planningStart = getPlanningStartDate();
    const planningEnd = getPlanningEndDate();
    const selectedCrewRows = scheduleCrewRows.filter((crew) => crew.shipCode === ship);

    if (!selectedCrewRows.length) {
      setScheduleRows([]);
      setScheduleMessage(`No crew rows found for ${getShipDisplayName(ship)} in the uploaded workbook.`);
      return;
    }

    const rows = selectedCrewRows
      .flatMap((crew) => buildCrewRotationRows(crew, planningStart, planningEnd))
      .sort((a, b) => {
        const dateA = parseExcelDate(a.startDate)?.getTime() || 0;
        const dateB = parseExcelDate(b.startDate)?.getTime() || 0;
        if (dateA !== dateB) return dateA - dateB;
        if (a.position !== b.position) return a.position.localeCompare(b.position);
        return a.name.localeCompare(b.name);
      });

    setScheduleRows(rows);
    setScheduleMessage(`Generated ${rows.length} schedule period(s) for ${getShipDisplayName(ship)} from ${formatDate(planningStart)} to ${formatDate(planningEnd)}.`);
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
    const ship = scheduleShip || userShip;
    return scheduleCrewRows.filter((crew) => !ship || crew.shipCode === ship);
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
      `${row.ship} ${row.shipName} ${row.position} ${row.idNumber} ${row.name} ${row.sheetName} ${row.periodType} ${row.startDate} ${row.endDate} ${row.status}`
        .toLowerCase()
        .includes(searchValue)
    );
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
      RotationRule: row.rotationRule,
      Notes: row.notes,
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Year Schedule");
    XLSX.writeFile(wb, `people-schedule-${scheduleShip || userShip || "ship"}-${formatDate(getPlanningStartDate()) || "year"}.xlsx`);
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
          <div><strong>Ship:</strong> ${getShipDisplayName(scheduleShip || userShip)}</div>
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

  const combinedBreakdown = selectedProduct ? getCombinedVenueBreakdown(selectedProduct) : [];
  const recipesForProduct = selectedProduct ? getRecipesUsingProduct(selectedProduct) : [];
  const productsInRecipe = selectedRecipe ? getProductsInRecipe(selectedRecipe) : [];
  const allergenWarnings = selectedRecipe ? detectAllergens(productsInRecipe) : [];
  const filteredProducts = products.filter((p) => p.toLowerCase().includes(search.toLowerCase()));

  const totalConsumption = (() => {
    const totals = { BRL: 0, RL: 0, SC: 0, VL: 0 };

    combinedBreakdown.forEach((venue) => {
      visibleShips.forEach((ship) => {
        totals[ship] += Number(venue.ships[ship] || 0);
      });
    });

    const allShips = visibleShips.reduce((sum, ship) => sum + totals[ship], 0);

    return { totals, allShips };
  })();

  if (!loggedIn) {
    return (
      <main style={styles.page}>
        <section style={styles.loginCard}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
          <h1 style={styles.title}>Virgin Voyages Dashboard</h1>
          <p style={styles.subtitle}>Product, equipment, people and schedule tools</p>

          <label style={styles.label}>🚢 Select your ship</label>
          <select value={userShip} onChange={(e) => setUserShip(e.target.value)} style={styles.select}>
            <option value="">Choose ship</option>
            {SHIPS.map((ship) => <option key={ship}>{ship}</option>)}
          </select>

          <button style={styles.primaryButton} onClick={() => userShip && setLoggedIn(true)}>
            Continue
          </button>
        </section>
      </main>
    );
  }

  if (!module) {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.shipBadge}>🚢 {userShip}</div>
        </header>

        <section style={styles.card}>
          <h2 style={styles.cardTitle}>🧭 Select Module</h2>

          <div style={styles.moduleGrid}>
            <button style={styles.moduleCard} onClick={() => setModule("product")}>
              <div style={styles.moduleIcon}>📦</div>
              <strong>Product Dashboard</strong>
              <span>Consumption, recipes, templates and allergens</span>
            </button>

            <button style={styles.moduleCard} onClick={() => setModule("equipment")}>
              <div style={styles.moduleIcon}>🍽️</div>
              <strong>Equipment</strong>
              <span>Muster list and inventory tools</span>
            </button>

            <button style={styles.moduleCard} onClick={() => setModule("people")}> 
              <div style={styles.moduleIcon}>👥</div>
              <strong>People & Schedule</strong>
              <span>Upload crew workbook and generate yearly contract rotations</span>
            </button>
          </div>
        </section>
      </main>
    );
  }

  if (module === "people") {
    const selectedScheduleCrewRows = getScheduleCrewRowsForSelectedShip();
    const filteredScheduleCrewRows = getFilteredScheduleCrewRows();
    const filteredScheduleRows = getFilteredScheduleRows();
    const planningStart = getPlanningStartDate();
    const planningEnd = getPlanningEndDate();
    const contractRows = scheduleRows.filter((row) => row.periodType === "Contract");
    const vacationRows = scheduleRows.filter((row) => row.periodType === "Vacation");

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(scheduleShip || userShip)}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>👥 People & Schedule</h2>

            <label style={styles.label}>Choose ship</label>
            <select value={scheduleShip} onChange={(e) => setScheduleShip(e.target.value)} style={styles.select}>
              <option value="">Choose ship</option>
              {SHIPS.map((ship) => <option key={ship} value={ship}>{ship} - {getShipDisplayName(ship)}</option>)}
            </select>

            <label style={styles.label}>Upload schedule workbook</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadScheduleFile} style={styles.fileInput} />

            <label style={styles.label}>Planning start date</label>
            <input type="date" value={scheduleDate} onChange={(e) => setScheduleDate(e.target.value)} style={styles.searchInput} />

            <div style={styles.headerActions}>
              <button style={styles.primaryButton} onClick={generateSchedule}>
                ✨ Generate Next Year Schedule
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
              <div>🚢 Ship: <strong>{getShipDisplayName(scheduleShip || userShip) || "Not selected"}</strong></div>
              <div>📅 Planning window: <strong>{formatDate(planningStart)} to {formatDate(planningEnd)}</strong></div>
              <div>📄 Tabs loaded: <strong>{scheduleWorkbookInfo.sheets.join(", ") || "None"}</strong></div>
              <div>👥 Total crew loaded: <strong>{scheduleCrewRows.length}</strong></div>
              <div>👥 Crew for selected ship: <strong>{selectedScheduleCrewRows.length}</strong></div>
              <div>✅ Contract periods generated: <strong>{contractRows.length}</strong></div>
              <div>🌴 Vacation periods generated: <strong>{vacationRows.length}</strong></div>
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
              <div style={{ color: "#8a5a00" }}>
                The workbook highlights I:IP with conditional formatting. The app reads D/E/F/G directly and uses the I:IP calendar window for matching/reporting.
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
              <button style={styles.backButton} onClick={printSchedule}>🖨️ Print</button>
              <button style={styles.primaryButton} onClick={exportScheduleToExcel}>📥 Export Excel</button>
            </div>
          </div>

          {scheduleRows.length === 0 && (
            <p style={styles.emptyText}>Generated yearly rotation schedule will appear here.</p>
          )}

          <div style={styles.equipmentGrid}>
            {filteredScheduleRows.map((row) => (
              <div
                key={row.id}
                style={{
                  ...styles.equipmentCard,
                  ...(row.periodType === "Contract" ? styles.countedCard : {}),
                  ...(row.periodType === "Missing Dates" ? styles.orderWarningCard : {}),
                }}
              >
                <div style={styles.recipeName}>{row.name}</div>
                <div style={styles.recipeMeta}>ID: {row.idNumber || "N/A"}</div>
                <div style={styles.recipeMeta}>Ship: {row.shipName || row.ship}</div>
                <div style={styles.recipeMeta}>Position: {row.position || "N/A"}</div>
                <div style={styles.recipeMeta}>Tab: {row.sheetName}</div>
                <div style={row.periodType === "Contract" ? styles.statusGood : row.periodType === "Vacation" ? styles.statusNeutral : styles.statusBad}>
                  {row.periodType}: {row.startDate || "N/A"} to {row.endDate || "N/A"}
                </div>
                <div style={styles.recipeMeta}>Status: {row.status}</div>
                <div style={styles.recipeMeta}>Rule: {row.rotationRule}</div>
                {row.notes && <div style={styles.recipeMeta}>Notes: {row.notes}</div>}
              </div>
            ))}
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && !equipmentMode) {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setModule("")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>

        <section style={styles.card}>
          <h2 style={styles.cardTitle}>🍽️ Equipment Options</h2>

          <div style={styles.moduleGrid}>
            <button style={styles.moduleCard} onClick={() => setEquipmentMode("muster")}>
              <div style={styles.moduleIcon}>📋</div>
              <strong>Equipment Muster List</strong>
              <span>Grouped by all sheets and sub categories</span>
            </button>

            <button style={styles.moduleCard} onClick={() => setEquipmentMode("inventory")}>
              <div style={styles.moduleIcon}>📊</div>
              <strong>Equipment Inventory</strong>
              <span>Inventory in use, warehouse stock and make inventory</span>
            </button>
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "inventory") {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>

        <section style={styles.card}>
          <h2 style={styles.cardTitle}>📊 Equipment Inventory</h2>

          <div style={styles.moduleGrid}>
            <button style={styles.moduleCard} onClick={() => setEquipmentMode("inuse")}>
              <div style={styles.moduleIcon}>✅</div>
              <strong>Inventory in Use</strong>
              <span>Compare muster list against in-use inventory</span>
            </button>

            <button style={styles.moduleCard} onClick={() => setEquipmentMode("warehouse")}>
              <div style={styles.moduleIcon}>🏬</div>
              <strong>Inventory Warehouse</strong>
              <span>Par, on hand, future order and suggested order</span>
            </button>

            <button style={styles.moduleCard} onClick={() => setEquipmentMode("makeinventory")}>
              <div style={styles.moduleIcon}>📝</div>
              <strong>Make Inventory</strong>
              <span>Multi-user counts, my report and ship summary</span>
            </button>
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "makeinventory") {
    const filteredMakeInventoryItems = getFilteredMakeInventoryItems();
    const myReportRows = getMyInventoryRows();
    const summaryReportRows = getShipSummaryRows();
    const visibleReportRows = getVisibleInventoryReportRows();
    const inventoryStatusRows = getMyInventoryStatusRows();
    const statusCountedItems = inventoryStatusRows.filter((item) => item.status === "Counted");
    const statusPendingItems = inventoryStatusRows.filter((item) => item.status !== "Counted");
    const userName = getEffectiveInventoryUserName();
    const inventoryReady = Boolean(makeInventoryShip && inventoryStation && userName && supabase);
    const countedKeysForMe = new Set(myReportRows.map((item) => item.itemKey || cleanText(item.code || item.name)));

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {makeInventoryShip || userShip}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📝 Make Inventory</h2>

            <label style={styles.label}>Choose ship for this inventory</label>
            <select
              value={makeInventoryShip}
              onChange={(e) => setMakeInventoryShip(e.target.value)}
              style={styles.select}
            >
              <option value="">Choose ship</option>
              {SHIPS.map((ship) => (
                <option key={ship} value={ship}>{ship}</option>
              ))}
            </select>

            <label style={styles.label}>Choose station</label>
            <select
              value={inventoryStation}
              onChange={(e) => setInventoryStation(e.target.value)}
              style={styles.select}
            >
              <option value="">Choose station</option>
              {STATIONS.map((station) => (
                <option key={station} value={station}>{station}</option>
              ))}
            </select>

            <label style={styles.label}>Choose user</label>
            <select
              value={inventoryUserName}
              onChange={(e) => setInventoryUserName(e.target.value)}
              style={styles.select}
            >
              <option value="">Choose user</option>
              {INVENTORY_USERS.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>

            {inventoryUserName === "Other" && (
              <>
                <label style={styles.label}>Enter your name</label>
                <input
                  placeholder="Enter your name..."
                  value={customInventoryUserName}
                  onChange={(e) => setCustomInventoryUserName(e.target.value)}
                  style={styles.searchInput}
                />
              </>
            )}

            <label style={styles.label}>Upload Master Inventory / Muster List</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadMakeInventoryFile} style={styles.fileInput} />

            {makeInventoryMessage && <p style={styles.message}>{makeInventoryMessage}</p>}

            <div style={styles.infoBox}>
              <div>🚢 Inventory ship: <strong>{makeInventoryShip || "Not selected"}</strong></div>
              <div>📍 Station: <strong>{inventoryStation || "Not selected"}</strong></div>
              <div>👤 User: <strong>{userName || "Not selected"}</strong></div>
              <div>📋 Master items loaded: <strong>{makeInventoryItems.length}</strong></div>
              <div>✅ My counted items: <strong>{myReportRows.length}</strong></div>
              <div>🌍 Ship summary records: <strong>{summaryReportRows.length}</strong></div>
              <div>Duplicate entries are automatically updated instead of added twice.</div>
              {inventoryLoading && <div>Loading shared inventory records...</div>}
              {inventoryError && <div style={{ color: "#b00020" }}>{inventoryError}</div>}
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.cardTitle}>🔍 Search Product</h2>

            <input
              placeholder="Search code, name, category or sheet..."
              value={makeInventorySearch}
              onChange={(e) => setMakeInventorySearch(e.target.value)}
              style={styles.searchInput}
            />

            {!inventoryReady && (
              <p style={styles.warningText}>
                Choose ship, station and user before counting. Supabase must be connected.
              </p>
            )}

            <p style={styles.emptyText}>
              Select the correct product from the master list. Green means already counted by this user for this station.
            </p>
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={styles.productTitle}>📦 Select Product for Inventory</h2>

          {makeInventoryItems.length === 0 && (
            <p style={styles.emptyText}>Upload the master inventory file to begin.</p>
          )}

          <div style={styles.equipmentGrid}>
            {filteredMakeInventoryItems.map((item, index) => {
              const itemKey = getInventoryItemKey(item);
              const alreadyCounted = countedKeysForMe.has(itemKey);

              return (
                <button
                  key={`${item.sheetName}-${item.code}-${index}`}
                  style={{ ...styles.equipmentCard, ...(alreadyCounted ? styles.countedCard : {}) }}
                  onClick={() => {
                    if (!inventoryReady) {
                      setMakeInventoryMessage("Choose ship, station and user before counting.");
                      return;
                    }

                    setCurrentInventoryItem(item);
                    setInventoryQty("");
                    setEditingInventoryId(null);
                  }}
                >
                  {item.image ? (
                    <div>
                      <img
                        src={getImageUrl(item.image)}
                        alt={item.name}
                        style={styles.equipmentImage}
                        onError={(e) => {
                          e.currentTarget.style.display = "none";
                          const link = e.currentTarget.nextElementSibling;
                          if (link) link.style.display = "block";
                        }}
                      />
                      <a href={item.image} target="_blank" rel="noreferrer" style={styles.imageLink}>
                        Open Picture
                      </a>
                    </div>
                  ) : (
                    <div style={styles.equipmentNoImage}>No image</div>
                  )}

                  <div style={styles.recipeName}>{item.name}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                  <div style={styles.recipeMeta}>Category: {item.category}</div>
                  {alreadyCounted && <div style={styles.statusGood}>Already Counted By Me</div>}
                </button>
              );
            })}
          </div>
        </section>

        {currentInventoryItem && (
          <section style={styles.card}>
            <h2 style={styles.productTitle}>
              {editingInventoryId ? "✏️ Edit Counted Product" : "✅ Confirm Product"}
            </h2>

            <div style={styles.grid}>
              <div>
                {currentInventoryItem.image ? (
                  <div>
                    <img
                      src={getImageUrl(currentInventoryItem.image)}
                      alt={currentInventoryItem.name}
                      style={styles.modalImage}
                      onError={(e) => {
                        e.currentTarget.style.display = "none";
                        const link = e.currentTarget.nextElementSibling;
                        if (link) link.style.display = "block";
                      }}
                    />
                    <a href={currentInventoryItem.image} target="_blank" rel="noreferrer" style={styles.imageLink}>
                      Open Picture
                    </a>
                  </div>
                ) : (
                  <div style={styles.equipmentNoImage}>No image</div>
                )}
              </div>

              <div>
                <h3>{currentInventoryItem.name}</h3>
                <p><strong>Ship:</strong> {makeInventoryShip || userShip}</p>
                <p><strong>Station:</strong> {inventoryStation || "N/A"}</p>
                <p><strong>User:</strong> {userName || "N/A"}</p>
                <p><strong>Code:</strong> {currentInventoryItem.code || "N/A"}</p>
                <p><strong>Sheet:</strong> {currentInventoryItem.sheetName}</p>
                <p><strong>Category:</strong> {currentInventoryItem.category}</p>

                <p style={styles.warningText}>Is this the correct product?</p>

                <label style={styles.label}>Quantity counted</label>
                <input
                  type="number"
                  min="0"
                  value={inventoryQty}
                  onChange={(e) => setInventoryQty(e.target.value)}
                  style={styles.searchInput}
                  placeholder="Enter quantity..."
                />

                <div style={styles.headerActions}>
                  <button
                    style={styles.backButton}
                    onClick={() => {
                      setCurrentInventoryItem(null);
                      setInventoryQty("");
                      setEditingInventoryId(null);
                    }}
                  >
                    No / Back
                  </button>

                  <button
                    style={styles.primaryButton}
                    onClick={confirmInventoryQty}
                    disabled={!inventoryReady}
                  >
                    {editingInventoryId ? "Update Quantity" : "Confirm Quantity"}
                  </button>
                </div>
              </div>
            </div>
          </section>
        )}

        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 20 }}>
            <h2 style={styles.productTitle}>📄 Inventory Report</h2>

            <div style={styles.headerActions}>
              <button
                style={{
                  ...styles.viewModeButton,
                  ...(inventoryReportMode === "my" ? styles.viewModeButtonActive : {}),
                }}
                onClick={() => setInventoryReportMode("my")}
              >
                👤 My Report
              </button>

              <button
                style={{
                  ...styles.viewModeButton,
                  ...(inventoryReportMode === "summary" ? styles.viewModeButtonActive : {}),
                }}
                onClick={() => setInventoryReportMode("summary")}
              >
                🌍 Summary Report
              </button>

              <button style={styles.backButton} onClick={() => loadInventoryRecords(makeInventoryShip || userShip)}>
                🔄 Refresh
              </button>

              <button style={styles.backButton} onClick={printInventorySummary}>
                🖨️ Print
              </button>

              {inventoryReportMode === "my" && (
                <button style={styles.deleteButton} onClick={clearMyInventory}>
                  🧹 Clear My Report
                </button>
              )}

              {inventoryReportMode === "summary" && (
                <button style={styles.deleteButton} onClick={clearShipInventory}>
                  🧹 Clear Ship Records
                </button>
              )}

              <button style={styles.primaryButton} onClick={exportInventorySummaryToExcel}>
                📥 Export Excel
              </button>
            </div>
          </div>

          <div style={styles.infoBox}>
            <div>🚢 Ship: <strong>{makeInventoryShip || userShip}</strong></div>
            {inventoryReportMode === "my" ? (
              <>
                <div>📍 Station: <strong>{inventoryStation || "Not selected"}</strong></div>
                <div>👤 User: <strong>{userName || "Not selected"}</strong></div>
                <div>✅ My records: <strong>{myReportRows.length}</strong></div>
              </>
            ) : (
              <>
                <div>🌍 Report: <strong>All users and all stations for selected ship</strong></div>
                <div>📦 Summary items: <strong>{summaryReportRows.length}</strong></div>
              </>
            )}
          </div>

          {visibleReportRows.length === 0 && (
            <p style={styles.emptyText}>
              {inventoryReportMode === "summary"
                ? "No ship summary records yet."
                : "Your confirmed quantities will appear here."}
            </p>
          )}

          <div style={styles.equipmentGrid}>
            {inventoryReportMode === "summary"
              ? visibleReportRows.map((item, index) => (
                  <div key={`${item.itemKey}-summary-${index}`} style={styles.equipmentCard}>
                    <div style={styles.recipeName}>{item.name}</div>
                    <div style={styles.recipeMeta}>Ship: {item.ship}</div>
                    <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                    <div style={styles.recipeMeta}>Category: {item.category}</div>
                    <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                    <div style={styles.statusGood}>Total Quantity: {formatQty(item.totalQty)}</div>
                    <div style={styles.recipeMeta}>Stations: {item.stations.join(", ") || "N/A"}</div>
                    <div style={styles.recipeMeta}>Users: {item.users.join(", ") || "N/A"}</div>
                    <div style={styles.recipeMeta}>Records: {item.recordCount}</div>
                    <div style={styles.recipeMeta}>Last Updated: {item.confirmedAt}</div>
                  </div>
                ))
              : visibleReportRows.map((item) => (
                  <div key={item.id} style={styles.equipmentCard}>
                    <div style={styles.recipeName}>{item.name}</div>
                    <div style={styles.recipeMeta}>Ship: {item.ship}</div>
                    <div style={styles.recipeMeta}>Station: {item.station}</div>
                    <div style={styles.recipeMeta}>User: {item.userName}</div>
                    <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                    <div style={styles.recipeMeta}>Category: {item.category}</div>
                    <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                    <div style={styles.statusGood}>Quantity: {formatQty(item.qty)}</div>
                    <div style={styles.recipeMeta}>Confirmed: {item.confirmedAt}</div>

                    <div style={styles.headerActions}>
                      <button style={styles.backButton} onClick={() => editInventoryItem(item)}>
                        ✏️ Edit
                      </button>

                      <button style={styles.deleteButton} onClick={() => deleteInventoryItem(item)}>
                        🗑️ Delete
                      </button>
                    </div>
                  </div>
                ))}
          </div>
        </section>

        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: showVariance ? 20 : 0 }}>
            <h2 style={styles.productTitle}>📊 Count Status</h2>

            <div style={styles.headerActions}>
              <button
                style={styles.backButton}
                onClick={() => setShowVariance((value) => !value)}
              >
                {showVariance ? "Hide Status" : "Open Status"}
              </button>

              {showVariance && (
                <>
                  <button style={styles.backButton} onClick={printInventoryStatus}>
                    🖨️ Print
                  </button>

                  <button style={styles.primaryButton} onClick={exportInventoryStatusToExcel}>
                    📥 Export Excel
                  </button>
                </>
              )}
            </div>
          </div>

          {showVariance && (
            <>
              <div style={styles.infoBox}>
                <div>📋 Master items: <strong>{makeInventoryItems.length}</strong></div>
                <div>✅ Counted by me: <strong>{statusCountedItems.length}</strong></div>
                <div>📋 Remaining for me: <strong>{statusPendingItems.length}</strong></div>
              </div>

              {makeInventoryItems.length === 0 && (
                <p style={styles.emptyText}>Upload the master inventory file to see count status.</p>
              )}

              <h3 style={styles.sectionTitle}>📋 My Master Inventory Status</h3>

              <div style={styles.equipmentGrid}>
                {inventoryStatusRows.map((item, index) => (
                  <div
                    key={`${item.code}-status-${index}`}
                    style={{ ...styles.equipmentCard, ...(item.status === "Counted" ? styles.countedCard : {}) }}
                  >
                    <div style={styles.recipeName}>{item.name}</div>
                    <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                    <div style={styles.recipeMeta}>Category: {item.category}</div>
                    <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>

                    {item.status === "Counted" ? (
                      <>
                        <div style={styles.statusGood}>Counted: {formatQty(item.countedQty)}</div>
                        <div style={styles.recipeMeta}>Counted: {item.countedAt}</div>
                      </>
                    ) : (
                      <div style={styles.statusNeutral}>Pending Count</div>
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

  if (module === "equipment" && equipmentMode === "inuse") {
    const inUseItems = parseInUseItems();
    const missingItems = inUseItems.filter((item) => item.status === "Missing");
    const zeroItems = inUseItems.filter((item) => item.status === "Zero Count");
    const activeItems = inUseItems.filter((item) => item.status === "In Use");

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📤 Upload Inventory in Use</h2>

            <label style={styles.label}>Step 1: Equipment Muster List file</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadMusterFile} style={styles.fileInput} />

            <label style={styles.label}>Step 2: Inventory in Use file</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadInUseFile} style={styles.fileInput} />

            {musterMessage && <p style={styles.message}>{musterMessage}</p>}
            {inUseMessage && <p style={styles.message}>{inUseMessage}</p>}

            <div style={styles.infoBox}>
              <div>❌ Missing from Inventory: <strong>{missingItems.length}</strong></div>
              <div>⚠️ Zero Count: <strong>{zeroItems.length}</strong></div>
              <div>✅ In Use: <strong>{activeItems.length}</strong></div>
              <div>Muster: C = Category, D = Code, E = Name</div>
              <div>In Use: A = Code, B = Name, H = On Hand</div>
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.cardTitle}>🔍 Search Inventory in Use</h2>

            <input
              placeholder="Search code, name, category, sheet or status..."
              value={inUseSearch}
              onChange={(e) => setInUseSearch(e.target.value)}
              style={styles.searchInput}
            />

            <p style={styles.emptyText}>
              Missing items are shown first in red under Missing from Inventory.
            </p>
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={{ ...styles.productTitle, color: "#b00020" }}>❌ Missing from Inventory</h2>

          {musterItems.length === 0 && <p style={styles.emptyText}>Upload the Equipment Muster List first.</p>}
          {inUseRows.length === 0 && <p style={styles.emptyText}>Upload the Inventory in Use file to compare.</p>}
          {musterItems.length > 0 && inUseRows.length > 0 && missingItems.length === 0 && (
            <p style={styles.emptyText}>No missing items found.</p>
          )}

          <div style={styles.equipmentGrid}>
            {missingItems.map((item, i) => (
              <div key={`${item.code}-missing-${i}`} style={{ ...styles.equipmentCard, ...styles.orderWarningCard }}>
                <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                <div style={styles.recipeMeta}>Category: {item.category}</div>
                <div style={styles.statusBad}>Missing</div>
              </div>
            ))}
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={{ ...styles.productTitle, color: "#8a5a00" }}>⚠️ Zero Count</h2>

          <div style={styles.equipmentGrid}>
            {zeroItems.map((item, i) => (
              <div key={`${item.code}-zero-${i}`} style={{ ...styles.equipmentCard, ...styles.zeroCountCard }}>
                <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                <div style={styles.recipeMeta}>Category: {item.category}</div>
                <div style={styles.recipeMeta}>On Hand: {formatQty(item.onHand)}</div>
                <div style={styles.statusWarning}>Zero Count</div>
              </div>
            ))}
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={{ ...styles.productTitle, color: "#2e7d32" }}>✅ In Use</h2>

          <div style={styles.equipmentGrid}>
            {activeItems.map((item, i) => (
              <div key={`${item.code}-active-${i}`} style={styles.equipmentCard}>
                <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                <div style={styles.recipeMeta}>Category: {item.category}</div>
                <div style={styles.recipeMeta}>On Hand: {formatQty(item.onHand)}</div>
                <div style={styles.statusGood}>In Use</div>
              </div>
            ))}
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "warehouse") {
    const warehouseItems = parseWarehouseItems();
    const totalSuggested = warehouseItems.reduce((sum, item) => sum + item.suggested, 0);
    const criticalItems = warehouseItems.filter((item) => item.suggested > 0).length;

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📤 Upload Warehouse Inventory</h2>
            <label style={styles.label}>Warehouse inventory file</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadWarehouseFile} style={styles.fileInput} />

            {warehouseMessage && <p style={styles.message}>{warehouseMessage}</p>}

            <div style={styles.infoBox}>
              <div>📦 Items loaded: <strong>{warehouseItems.length}</strong></div>
              <div>🚨 Items needing order: <strong>{criticalItems}</strong></div>
              <div>🛒 Total suggested order: <strong>{formatQty(totalSuggested)}</strong></div>
              <div>A = Code, B = Name, G = Par, H = On Hand, M = Future Order</div>
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.cardTitle}>🔍 Search Warehouse</h2>
            <input
              placeholder="Search code or equipment name..."
              value={warehouseSearch}
              onChange={(e) => setWarehouseSearch(e.target.value)}
              style={styles.searchInput}
            />
            <p style={styles.emptyText}>Suggested Next Order = Par Level - On Hand - Future Order. Minimum is 0.</p>
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={styles.productTitle}>🏬 Inventory Warehouse</h2>

          {warehouseRows.length === 0 && <p style={styles.emptyText}>Upload the warehouse inventory file to begin.</p>}

          <div style={styles.equipmentGrid}>
            {warehouseItems.map((item, i) => (
              <div key={`${item.code}-${i}`} style={{ ...styles.equipmentCard, ...(item.suggested > 0 ? styles.orderWarningCard : {}) }}>
                <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                <div style={styles.recipeMeta}>Par Level: {formatQty(item.par)}</div>
                <div style={styles.recipeMeta}>On Hand: {formatQty(item.onHand)}</div>
                <div style={styles.recipeMeta}>Future Order: {formatQty(item.future)}</div>
                <div style={item.suggested > 0 ? styles.suggestedOrderBad : styles.suggestedOrderGood}>
                  Suggested Next Order: {formatQty(item.suggested)}
                </div>
              </div>
            ))}
          </div>
        </section>
      </main>
    );
  }

  if (module === "equipment" && equipmentMode === "muster") {
    const groupedMuster = parseMusterItems();
    const totalItems = Object.values(groupedMuster).reduce((sum, items) => sum + items.length, 0);

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {userShip}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📤 Upload Equipment File</h2>
            <label style={styles.label}>Equipment Muster List file</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadMusterFile} style={styles.fileInput} />

            {musterMessage && <p style={styles.message}>{musterMessage}</p>}

            <div style={styles.infoBox}>
              <div>📋 Items loaded: <strong>{totalItems}</strong></div>
              <div>📄 Sheets included: <strong>{[...new Set(musterItems.map((i) => i.sheetName))].length}</strong></div>
              <div>🗂️ Groups: <strong>{Object.keys(groupedMuster).length}</strong></div>
              <div>C = Sub Category, D = Code, E = Name, H = Picture Link</div>
            </div>
          </div>

          <div style={styles.card}>
            <h2 style={styles.cardTitle}>🔍 Search Equipment</h2>
            <input
              placeholder="Search equipment, code, sheet or sub category..."
              value={musterSearch}
              onChange={(e) => setMusterSearch(e.target.value)}
              style={styles.searchInput}
            />
            <p style={styles.emptyText}>Click any equipment card to open the picture and full details.</p>
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={styles.productTitle}>📋 Equipment Muster List</h2>

          {musterItems.length === 0 && <p style={styles.emptyText}>Upload the Equipment Muster List file to begin.</p>}

          {Object.entries(groupedMuster).map(([category, items]) => (
            <div key={category} style={styles.equipmentCategory}>
              <h3 style={styles.sectionTitle}>🗂️ {category}</h3>

              <div style={styles.equipmentGrid}>
                {items.map((item, index) => (
                  <button
                    key={`${item.sheetName}-${item.code}-${index}`}
                    style={styles.equipmentCard}
                    onClick={() => setSelectedEquipment(item)}
                  >
                    {item.image ? (
                      <div>
                        <img
                          src={getImageUrl(item.image)}
                          alt={item.name}
                          style={styles.equipmentImage}
                          onError={(e) => {
                            e.currentTarget.style.display = "none";
                            const link = e.currentTarget.nextElementSibling;
                            if (link) link.style.display = "block";
                          }}
                        />
                        <a href={item.image} target="_blank" rel="noreferrer" style={styles.imageLink}>Open Picture</a>
                      </div>
                    ) : (
                      <div style={styles.equipmentNoImage}>No image</div>
                    )}

                    <div style={styles.recipeName}>{item.name}</div>
                    <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                    <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                    <div style={styles.recipeMeta}>Category: {item.category}</div>
                  </button>
                ))}
              </div>
            </div>
          ))}

          {selectedEquipment && (
            <div style={styles.modalBackdrop} onClick={() => setSelectedEquipment(null)}>
              <div style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
                <button style={styles.closeButton} onClick={() => setSelectedEquipment(null)}>✕</button>

                <h2>{selectedEquipment.name}</h2>
                <p><strong>Code:</strong> {selectedEquipment.code || "N/A"}</p>
                <p><strong>Sheet:</strong> {selectedEquipment.sheetName || "N/A"}</p>
                <p><strong>Category:</strong> {selectedEquipment.category || "N/A"}</p>

                {selectedEquipment.image ? (
                  <div>
                    <img
                      src={getImageUrl(selectedEquipment.image)}
                      alt={selectedEquipment.name}
                      style={styles.modalImage}
                      onError={(e) => {
                        e.currentTarget.style.display = "none";
                        const link = e.currentTarget.nextElementSibling;
                        if (link) link.style.display = "block";
                      }}
                    />
                    <a href={selectedEquipment.image} target="_blank" rel="noreferrer" style={styles.imageLink}>Open Picture</a>
                  </div>
                ) : (
                  <div style={styles.equipmentNoImage}>No image</div>
                )}
              </div>
            </div>
          )}
        </section>
      </main>
    );
  }

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
          <div style={styles.shipBadge}>🚢 {userShip}</div>
        </div>
      </header>

      <div style={styles.viewModeBox}>
        <button onClick={() => setViewMode("single")} style={{ ...styles.viewModeButton, ...(viewMode === "single" ? styles.viewModeButtonActive : {}) }}>
          🚢 {userShip} Only
        </button>

        <button onClick={() => setViewMode("all")} style={{ ...styles.viewModeButton, ...(viewMode === "all" ? styles.viewModeButtonActive : {}) }}>
          🌍 All Ships Overview
        </button>
      </div>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>📤 Upload Files</h2>

          <label style={styles.label}>Step 1: Consumption file</label>
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadConsumptionFile} style={styles.fileInput} />

          <label style={styles.label}>Step 2: Recipe / location file</label>
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadRecipeFile} style={styles.fileInput} />

          <label style={styles.label}>Optional: Replace template file</label>
          <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadTemplateFile} style={styles.fileInput} />

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.infoBox}>
            <div>📦 Products loaded: <strong>{products.length}</strong></div>
            <div>📘 Recipe rows loaded: <strong>{Math.max(recipeRows.length - 1, 0)}</strong></div>
            <div>📋 Template: <strong>{templateStatus}</strong></div>
            <div style={{ color: "#b00020" }}>Red = recipe/location expects usage, but consumption is 0 for visible ship(s).</div>
            <div style={{ color: "#0057b8" }}>Blue = product is in recipe/location, but missing from template for that venue.</div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔍 Select Product</h2>
          <input placeholder="Search product..." value={search} onChange={(e) => setSearch(e.target.value)} style={styles.searchInput} />

          <div style={styles.productList}>
            {filteredProducts.map((product, i) => (
              <button
                key={i}
                onClick={() => {
                  setSelectedProduct(product);
                  setSelectedRecipe(null);
                }}
                style={{ ...styles.productItem, ...(selectedProduct === product ? styles.productItemActive : {}) }}
              >
                {product}
              </button>
            ))}
          </div>
        </div>
      </section>

      {selectedProduct && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>📦 {selectedProduct}</h2>
          <h3 style={styles.sectionTitle}>📊 Total Consumption</h3>

          <div style={styles.totalBox}>
            <div style={styles.totalMain}>
              {viewMode === "single" ? `${userShip} Total: ` : "Total All Ships: "}
              {formatQty(totalConsumption.allShips)}
            </div>

            <div style={styles.totalShipGrid}>
              {visibleShips.map((ship) => (
                <div key={ship} style={styles.totalShipBox}>
                  <span>{ship}</span>
                  <strong>{formatQty(totalConsumption.totals[ship])}</strong>
                </div>
              ))}
            </div>
          </div>

          <h3 style={styles.sectionTitle}>🏢 Consumption by Venue and Ship</h3>

          <div style={styles.venueGrid}>
            {combinedBreakdown.map((venueItem, i) => (
              <div key={i} style={{ ...styles.venueCard, ...(venueItem.missingShips.length > 0 ? styles.venueCardWarning : {}), ...(venueItem.missingFromTemplate ? styles.venueCardTemplateWarning : {}) }}>
                <h4 style={styles.venueTitle}>
                  {venueItem.displayName}
                  <span style={styles.badgeGroup}>
                    {venueItem.missingFromTemplate && <span style={styles.templateBadge}>Missing Template</span>}
                    {venueItem.missingShips.length > 0 && <span style={styles.missingBadge}>Missing: {venueItem.missingShips.join(", ")}</span>}
                  </span>
                </h4>

                {venueItem.templateMatches.length > 0 && <div style={styles.templateFound}>Template/Menu: {venueItem.templateMatches.join(", ")}</div>}

                {venueItem.missingFromTemplate && (
                  <div style={styles.templateWarningText}>
                    Product is used in recipe/location file for this venue but is not found in any template. Product has to be used.
                  </div>
                )}

                <div style={styles.shipGrid}>
                  {visibleShips.map((ship) => {
                    const isMissing = venueItem.required && (venueItem.ships[ship] || 0) === 0;

                    return (
                      <div key={ship} style={{ ...styles.shipBox, ...(ship === userShip ? styles.shipBoxActive : {}), ...(isMissing ? styles.shipBoxMissing : {}) }}>
                        <span style={styles.shipName}>{ship}</span>
                        <strong style={styles.shipQty}>{formatQty(venueItem.ships[ship])}</strong>
                      </div>
                    );
                  })}
                </div>

                {venueItem.missingShips.length > 0 && (
                  <div style={styles.warningSmall}>Product appears in recipe/location file for this venue, but usage is 0 for highlighted ship(s).</div>
                )}
              </div>
            ))}
          </div>

          <h3 style={styles.sectionTitle}>👨‍🍳 Recipes using this product</h3>

          {recipeRows.length === 0 && <p style={styles.emptyText}>Upload the recipe/location file to see recipe details.</p>}
          {recipeRows.length > 0 && recipesForProduct.length === 0 && <p style={styles.emptyText}>No recipes found for this product.</p>}

          <div style={styles.recipeList}>
            {recipesForProduct.map((recipe, i) => (
              <button key={i} onClick={() => setSelectedRecipe(recipe)} style={{ ...styles.recipeCard, ...(selectedRecipe?.key === recipe.key ? styles.recipeCardActive : {}) }}>
                <div style={styles.recipeName}>{recipe.recipeName}</div>
                <div style={styles.recipeMeta}>Code: {recipe.recipeCode}</div>
                <div style={styles.recipeMeta}>Venues: {recipe.venues.length ? recipe.venues.join(", ") : "N/A"}</div>
              </button>
            ))}
          </div>

          {selectedRecipe && (
            <div style={styles.ingredientsCard}>
              <h3 style={styles.sectionTitle}>🧾 Products used in recipe</h3>
              <h4 style={{ marginTop: 0 }}>{selectedRecipe.recipeName} ({selectedRecipe.recipeCode})</h4>

              {productsInRecipe.length === 0 ? (
                <p style={styles.emptyText}>No products found for this recipe.</p>
              ) : (
                <ul>
                  {productsInRecipe.map((product, i) => {
                    const subIngredients = getSubRecipeIngredients(product);

                    return (
                      <li key={i} style={{ marginBottom: 10 }}>
                        <strong>{product}</strong>
                        {subIngredients.length > 0 && (
                          <ul style={styles.subRecipeList}>
                            {subIngredients.map((subItem, j) => <li key={j}>{subItem}</li>)}
                          </ul>
                        )}
                      </li>
                    );
                  })}
                </ul>
              )}

              <h3 style={styles.sectionTitle}>⚠️ Rule-Based Allergen Warning</h3>
              <p style={styles.warningText}>This is a keyword-based warning only. Verify against official allergen data before use.</p>

              {allergenWarnings.length === 0 ? (
                <p style={styles.emptyText}>No likely allergens detected by keyword rules.</p>
              ) : (
                <div style={styles.allergenList}>
                  {allergenWarnings.map((item, i) => (
                    <div key={i} style={styles.allergenCard}>
                      <strong>{item.allergen}</strong>
                      <ul>
                        {item.products.map((product, j) => <li key={j}>{product}</li>)}
                      </ul>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </section>
      )}
    </main>
  );
}

const styles = {
  page: { minHeight: "100vh", padding: 24, background: "#f5f5f5", fontFamily: "Arial, sans-serif", color: "#111" },
  loginCard: { maxWidth: 460, margin: "80px auto", padding: 28, background: "#fff", borderRadius: 16, boxShadow: "0 10px 30px rgba(0,0,0,0.08)", display: "grid", gap: 14 },
  logo: { height: 70, objectFit: "contain", marginBottom: 8 },
  header: { display: "flex", alignItems: "center", justifyContent: "space-between", padding: 18, background: "#fff", borderRadius: 16, boxShadow: "0 4px 18px rgba(0,0,0,0.06)", marginBottom: 20 },
  headerLogo: { height: 54, objectFit: "contain" },
  headerActions: { display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" },
  backButton: { padding: "10px 14px", borderRadius: 999, border: "1px solid #ccc", background: "#fff", cursor: "pointer", fontWeight: "bold" },
  deleteButton: { padding: "10px 14px", borderRadius: 999, border: "1px solid #b00020", background: "#b00020", color: "#fff", cursor: "pointer", fontWeight: "bold" },
  title: { margin: 0, fontSize: 28 },
  subtitle: { margin: 0, color: "#666" },
  label: { fontWeight: "bold", marginTop: 8 },
  select: { padding: 10, borderRadius: 8, border: "1px solid #ccc", width: "100%" },
  primaryButton: { marginTop: 10, padding: 12, borderRadius: 10, border: 0, background: "#111", color: "#fff", fontWeight: "bold", cursor: "pointer" },
  shipBadge: { padding: "10px 14px", borderRadius: 999, background: "#111", color: "#fff", fontWeight: "bold" },
  moduleGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))", gap: 16 },
  moduleCard: { border: "1px solid #ddd", background: "#fafafa", borderRadius: 16, padding: 20, cursor: "pointer", textAlign: "left", display: "grid", gap: 8, boxShadow: "0 4px 12px rgba(0,0,0,0.04)" },
  moduleIcon: { fontSize: 34 },
  viewModeBox: { display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" },
  viewModeButton: { padding: "10px 14px", borderRadius: 999, border: "1px solid #ccc", background: "#fff", cursor: "pointer", fontWeight: "bold" },
  viewModeButtonActive: { background: "#111", color: "#fff", borderColor: "#111" },
  grid: { display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 20, marginBottom: 20 },
  formRow: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  card: { background: "#fff", borderRadius: 16, padding: 20, boxShadow: "0 4px 18px rgba(0,0,0,0.06)" },
  cardTitle: { marginTop: 0 },
  fileInput: { display: "block", margin: "8px 0 16px" },
  message: { color: "#555", fontSize: 14 },
  infoBox: { marginTop: 12, padding: 12, borderRadius: 12, background: "#f2f2f2", display: "grid", gap: 6 },
  searchInput: { width: "100%", padding: 12, borderRadius: 10, border: "1px solid #ccc", marginBottom: 10 },
  productList: { maxHeight: 300, overflowY: "auto", border: "1px solid #ddd", borderRadius: 12 },
  productItem: { width: "100%", display: "block", textAlign: "left", padding: 10, border: 0, borderBottom: "1px solid #eee", background: "#fff", cursor: "pointer" },
  productItemActive: { background: "#eee", fontWeight: "bold" },
  productTitle: { marginTop: 0, fontSize: 24 },
  sectionTitle: { marginTop: 22 },
  totalBox: { background: "#111", color: "#fff", borderRadius: 14, padding: 16, marginBottom: 18 },
  totalMain: { fontSize: 20, fontWeight: "bold" },
  totalShipGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: 10, marginTop: 12 },
  totalShipBox: { background: "#fff", color: "#111", borderRadius: 10, padding: 10, textAlign: "center", display: "grid", gap: 4 },
  venueGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 14 },
  venueCard: { border: "1px solid #ddd", borderRadius: 14, padding: 14, background: "#fafafa" },
  venueCardWarning: { border: "2px solid #b00020", background: "#fff0f0" },
  venueCardTemplateWarning: { border: "2px solid #0057b8", background: "#eef5ff" },
  venueTitle: { marginTop: 0, display: "flex", justifyContent: "space-between", gap: 10, alignItems: "center" },
  badgeGroup: { display: "flex", flexWrap: "wrap", gap: 6, justifyContent: "flex-end" },
  missingBadge: { fontSize: 12, color: "#fff", background: "#b00020", borderRadius: 999, padding: "4px 8px" },
  templateBadge: { fontSize: 12, color: "#fff", background: "#0057b8", borderRadius: 999, padding: "4px 8px" },
  templateFound: { color: "#0057b8", fontSize: 13, fontWeight: "bold", marginBottom: 10 },
  templateWarningText: { color: "#0057b8", fontSize: 13, fontWeight: "bold", marginBottom: 10 },
  shipGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(70px, 1fr))", gap: 6 },
  shipBox: { minWidth: 0, padding: "8px 4px", borderRadius: 10, background: "#fff", border: "1px solid #ddd", display: "grid", gap: 3, textAlign: "center", overflow: "hidden" },
  shipBoxActive: { background: "#111", color: "#fff" },
  shipBoxMissing: { background: "#b00020", color: "#fff", borderColor: "#b00020" },
  shipName: { fontSize: 11, opacity: 0.8 },
  shipQty: { fontSize: 14, lineHeight: 1.1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  warningSmall: { marginTop: 10, color: "#b00020", fontSize: 13, fontWeight: "bold" },
  emptyText: { color: "#777" },
  recipeList: { display: "grid", gap: 10 },
  recipeCard: { width: "100%", textAlign: "left", border: "1px solid #ddd", borderRadius: 12, padding: 12, background: "#fafafa", cursor: "pointer" },
  recipeCardActive: { background: "#eee", borderColor: "#111" },
  recipeName: { fontWeight: "bold" },
  recipeMeta: { color: "#555", fontSize: 14, marginTop: 4 },
  ingredientsCard: { marginTop: 18, border: "1px solid #ddd", borderRadius: 14, padding: 14, background: "#fafafa" },
  subRecipeList: { marginTop: 6, marginBottom: 8, paddingLeft: 24, color: "#333", background: "#f2f2f2", borderRadius: 8, paddingTop: 8, paddingBottom: 8 },
  warningText: { color: "#8a5a00", background: "#fff4d6", padding: 10, borderRadius: 8 },
  allergenList: { display: "grid", gap: 10 },
  allergenCard: { border: "1px solid #e1c16e", background: "#fff9e8", borderRadius: 10, padding: 10 },
  equipmentCategory: { marginBottom: 24 },
  equipmentGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 },
  equipmentCard: { border: "1px solid #ddd", borderRadius: 14, padding: 14, background: "#fafafa", display: "grid", gap: 8, cursor: "pointer", textAlign: "left" },
  countedCard: { border: "2px solid #2e7d32", background: "#f0fff4" },
  equipmentImage: { width: "100%", height: 150, objectFit: "cover", borderRadius: 10, background: "#eee" },
  equipmentNoImage: { height: 150, borderRadius: 10, background: "#eee", display: "flex", alignItems: "center", justifyContent: "center", color: "#777" },
  modalBackdrop: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999, padding: 20 },
  modalCard: { background: "#fff", borderRadius: 18, padding: 22, maxWidth: 760, width: "100%", maxHeight: "90vh", overflowY: "auto", position: "relative", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" },
  modalImage: { width: "100%", maxHeight: "65vh", objectFit: "contain", borderRadius: 14, background: "#f2f2f2" },
  closeButton: { position: "absolute", top: 12, right: 12, border: 0, background: "#111", color: "#fff", borderRadius: 999, width: 34, height: 34, cursor: "pointer", fontWeight: "bold" },
  imageLink: { display: "none", marginTop: 8, padding: 10, borderRadius: 10, background: "#111", color: "#fff", textAlign: "center", textDecoration: "none", fontWeight: "bold" },
  orderWarningCard: { border: "2px solid #b00020", background: "#fff0f0" },
  zeroCountCard: { border: "2px solid #8a5a00", background: "#fff8e1" },
  suggestedOrderBad: { marginTop: 8, padding: 8, borderRadius: 10, background: "#b00020", color: "#fff", fontWeight: "bold", textAlign: "center" },
  suggestedOrderGood: { marginTop: 8, padding: 8, borderRadius: 10, background: "#e8f5e9", color: "#2e7d32", fontWeight: "bold", textAlign: "center" },
  statusGood: { marginTop: 8, padding: 8, borderRadius: 10, background: "#e8f5e9", color: "#2e7d32", fontWeight: "bold", textAlign: "center" },
  statusWarning: { marginTop: 8, padding: 8, borderRadius: 10, background: "#fff4d6", color: "#8a5a00", fontWeight: "bold", textAlign: "center" },
  statusNeutral: { marginTop: 8, padding: 8, borderRadius: 10, background: "#f2f2f2", color: "#555", fontWeight: "bold", textAlign: "center" },
  statusBad: { marginTop: 8, padding: 8, borderRadius: 10, background: "#b00020", color: "#fff", fontWeight: "bold", textAlign: "center" },
};

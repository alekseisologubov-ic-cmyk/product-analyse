"use client";

import React, { Suspense, lazy, useEffect, useMemo, useRef, useState } from "react";
import { createClient } from "@supabase/supabase-js";
import {
  cleanText,
  normalizeVenue,
  productNamesMatch,
  getProductMatchTokens,
  getProductReportKey as getReportProductKey,
  formatQty,
  formatMoney,
  getImageUrl,
  isUsableImageValue,
  escapeHtml,
  toNumber,
} from "./lib/appHelpers";

import {
  readExcelFile,
  workbookToRows,
  parseTemplateWorkbook,
  parseEquipmentMasterFile,
  parseNextOrderWorkbook,
} from "./lib/excelParsers";

import {
  downloadInventoryPdfReport,
  downloadInventoryExcelReportUsingTemplate,
} from "./lib/inventoryReportDownloads";

import {
  downloadIngredientByLocationFileFromStorage,
  uploadIngredientByLocationFileToStorage,
} from "./lib/permanentFiles";

import {
  YEARLY_REGION_ALL,
  parseYearlyRegionalConsumptionWorkbook,
  formatRegionalQty,
} from "./lib/yearlyRegionalConsumption";

import {
  enrichDashboardProductsWithRegionalPar,
} from "./lib/productDashboardRegionalHelpers";

import {
  BAR_STATIONS,
  RESTAURANT_STATIONS,
  CULINARY_STATIONS,
  EQUIPMENT_PICTURE_BUCKET,
  getMasterInventoryScope,
} from "./constants/inventoryConfig";

import {
  buildInventoryQtyMap,
  getEquipmentDepartmentKey,
  getInventoryProductGroupKey,
  inventoryRecordMatchesDepartment,
  normalizeInventoryRecord,
  normalizeInventoryStationStatusRecord,
  normalizeMasterInventoryRecord,
} from "./lib/inventoryHelpers";

import {
  cleanSharedMasterImage,
  getEquipmentImageMatchCode,
  getImageExtensionFromMime,
  getImageMimeFromDataUrl,
  getZipImageExtension,
  getZipImageMimeType,
  makeStorageSafePart,
  normalizeEquipmentPictureCode,
} from "./lib/inventoryImageHelpers";

import MakeInventoryTopBar from "./components/equipment/MakeInventoryTopBar";
import EquipmentDepartmentOptions from "./components/equipment/EquipmentDepartmentOptions";
import EquipmentInventoryOptions from "./components/equipment/EquipmentInventoryOptions";
import EquipmentMusterModule from "./components/equipment/EquipmentMusterModule";
import AdminDashboard from "./components/admin/AdminDashboard";
import { AppProvider } from "./context/AppContext";

const loadPeopleScheduleModule = () => import("./components/PeopleScheduleModule");
const PeopleScheduleModule = lazy(loadPeopleScheduleModule);

const GenerateNextOrder = lazy(() =>
  import("./components/product/GenerateNextOrder")
);

const TemperatureCheckModule = lazy(() =>
  import("./components/temperature/TemperatureCheckModule")
);

const TrainingModule = lazy(() =>
  import("./components/training/TrainingModule")
);

const AllergenModule = lazy(() =>
  import("./components/allergen/AllergenModule")
);
const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL || "";
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY || "";

const supabase =
  supabaseUrl && supabaseAnonKey
    ? createClient(supabaseUrl, supabaseAnonKey)
    : null;

const loadXlsx = async () => {
  const module = await import("xlsx");
  return module.default || module;
};

const getOrCreateVisitorId = () => {
  if (typeof window === "undefined") return "";

  const storageKey = "vv_app_visitor_id";
  const existing = window.localStorage.getItem(storageKey);

  if (existing) return existing;

  const nextId =
    typeof window.crypto !== "undefined" && window.crypto.randomUUID
      ? window.crypto.randomUUID()
      : `visitor-${Date.now()}-${Math.random().toString(16).slice(2)}`;

  window.localStorage.setItem(storageKey, nextId);
  return nextId;
};

const normalizeAppEmail = (value) => String(value || "").trim().toLowerCase();

const isVirginVoyagesEmail = (value) => {
  const email = normalizeAppEmail(value);
  return /^[^\s@]+@virginvoyages\.com$/.test(email);
};

const USER_EMAIL_STORAGE_KEY = "vv_app_user_email";

const ADMIN_EMAILS = new Set([
  "aleksei.sologubov-ic@virginvoyages.com",
  // add more admin emails here
]);

const isAdminEmail = (value) => ADMIN_EMAILS.has(normalizeAppEmail(value));

const SHIPS = ["SC", "VL", "BRL", "RL"];

const SHIP_DISPLAY_NAMES = {
  SC: "Scarlet",
  VL: "Valiant",
  BRL: "Brilliant",
  RL: "Resilient",
};
const getTemplateSheetShipScope = (sheetName) => {
  const text = cleanText(sheetName)
    .replace(/RESILIANT/g, "RESILIENT")
    .replace(/\bV\s*[-]?\s*1\b/g, "V1")
    .replace(/\bS\s*C\s*L\b/g, "SCL");

  const scope = [];

  if (
    /\bSCL\b/.test(text) ||
    /\bSC\b/.test(text) ||
    /\bV1\b/.test(text) ||
    text.includes("SCARLET")
  ) {
    scope.push("SC");
  }

  if (
    /\bVAL\b/.test(text) ||
    /\bVL\b/.test(text) ||
    text.includes("VALIANT")
  ) {
    scope.push("VL");
  }

  if (
    /\bRES\b/.test(text) ||
    /\bRL\b/.test(text) ||
    text.includes("RESILIENT")
  ) {
    scope.push("RL");
  }

  if (
    /\bBRL\b/.test(text) ||
    text.includes("BRILLIANT")
  ) {
    scope.push("BRL");
  }

  return [...new Set(scope)];
};

const SCHEDULE_ALL_SHIPS = "ALL";
const YEARLY_REPORT_ALL_SHIPS = "ALL_SHIPS";

const getShipDisplayName = (shipCode) => SHIP_DISPLAY_NAMES[shipCode] || shipCode || "";

const getScheduleShipDisplayName = (shipCode) =>
  shipCode === SCHEDULE_ALL_SHIPS ? "All Ships" : getShipDisplayName(shipCode);

const normalizeShipCode = (value) => {
  const text = cleanText(value);

  if (!text) return "";

  if (
    text === "SC" ||
    text.includes("SCARLET") ||
    text.includes("SCARLET LADY")
  ) {
    return "SC";
  }

  if (
    text === "VL" ||
    text.includes("VALIANT") ||
    text.includes("VALIANT LADY")
  ) {
    return "VL";
  }

  if (
    text === "BRL" ||
    text === "BR" ||
    text.includes("BRILLIANT") ||
    text.includes("BRILLIANT LADY")
  ) {
    return "BRL";
  }

  if (
    text === "RL" ||
    text === "RES" ||
    text.includes("RESILIENT") ||
    text.includes("RESILIENT LADY")
  ) {
    return "RL";
  }

  return "";
};

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

const cleanTemplateTitle = (value) =>
  String(value || "")
    .replace(/\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/g, "")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*$/g, "")
    .trim();

const cleanTemplateSheetDisplay = (sheetName) =>
  String(sheetName || "")
    .replace(/^\d+\s*[-]?\s*/g, "")
    .replace(/\bSCL\b/gi, "")
    .replace(/\bVAL\b/gi, "")
    .replace(/\bRES\b/gi, "")
    .replace(/\bBRL\b/gi, "")
    .replace(/\bROJO\b/gi, "")
    .replace(/\bARIYA\b/gi, "")
    .replace(/\bONLY\b/gi, "")
    .replace(/\s*-\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim();

const getTemplateSectionName = (templateName) => {
  const cleaned = cleanTemplateTitle(templateName);
  const parts = cleaned.split(/\s*-\s*/).map((part) => part.trim()).filter(Boolean);
  return parts.length > 1 ? parts[parts.length - 1] : cleaned;
};

const getTemplateLocationDisplay = (sheetName, templateName) => {
  const sheetDisplay = cleanTemplateSheetDisplay(sheetName);
  const sectionName = getTemplateSectionName(templateName);

  if (!sheetDisplay && !sectionName) return "Template";
  if (!sheetDisplay) return sectionName;
  if (!sectionName) return sheetDisplay;

  const sheetKey = normalizeVenue(sheetDisplay);
  const sectionKey = normalizeVenue(sectionName);

  if (!sectionKey || sheetKey === sectionKey || sheetKey.includes(sectionKey)) {
    return sheetDisplay;
  }

  return `${sheetDisplay} - ${sectionName}`;
};

const getTemplateLocationKey = (sheetName, templateName) =>
  normalizeVenue(getTemplateLocationDisplay(sheetName, templateName));

const getHistoricalSailorDays = (cellA, cellB) => {
  const a = toNumber(cellA);
  const b = toNumber(cellB);

  if (!a && !b) return 0;
  if (a && !b) return a;
  if (!a && b) return b;

  const low = Math.min(Math.abs(a), Math.abs(b));
  const high = Math.max(Math.abs(a), Math.abs(b));

  // In these order files, one row can be # days and the other can be total sailor-days.
  // If the large value is much bigger than days x 1000, treat it as already total sailor-days.
  if (low > 0 && high > low * 1000) return high;

  // Otherwise treat the two cells as sailors x days.
  return a * b;
};

const excelDateToDate = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
  }

  const text = String(value || "").trim();
  if (!text) return null;

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  return null;
};

const formatDateCell = (value) => {
  const date = excelDateToDate(value);
  if (!date) return String(value || "").trim();
  return date.toLocaleDateString();
};

const getDaysBetweenCells = (startValue, endValue) => {
  const startDate = excelDateToDate(startValue);
  const endDate = excelDateToDate(endValue);

  if (!startDate || !endDate) return 0;

  const startUtc = Date.UTC(startDate.getUTCFullYear(), startDate.getUTCMonth(), startDate.getUTCDate());
  const endUtc = Date.UTC(endDate.getUTCFullYear(), endDate.getUTCMonth(), endDate.getUTCDate());
  const days = Math.round((endUtc - startUtc) / (24 * 60 * 60 * 1000));

  return Number.isFinite(days) && days > 0 ? days : 0;
};

const getDefaultNextYearStartDate = () => {
  const nextYear = new Date().getFullYear() + 1;
  return String(nextYear) + "-01-01";
};

const APP_BROWSER_HISTORY_KEY = "__vv_app_navigation_state__";

const getAppNavigationStateFromHistory = (historyState) => {
  if (!historyState || typeof historyState !== "object") {
    return null;
  }

  return historyState[APP_BROWSER_HISTORY_KEY] || null;
};

const buildAppBrowserHistoryState = (currentHistoryState, navigationState) => ({
  ...(currentHistoryState && typeof currentHistoryState === "object"
    ? currentHistoryState
    : {}),
  [APP_BROWSER_HISTORY_KEY]: navigationState,
});

const getAppNavigationStateKey = (navigationState) =>
  JSON.stringify(navigationState || {});

export default function App() {
  const [consumptionRows, setConsumptionRows] = useState([]);
  const [recipeRows, setRecipeRows] = useState([]);
  const [templateMap, setTemplateMap] = useState({});
  const [templateStatus, setTemplateStatus] = useState(
  "Template will load when Product tools open."
);
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [search, setSearch] = useState("");
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);
  const [welcomeStarted, setWelcomeStarted] = useState(false);
  const [userEmail, setUserEmail] = useState("");
  const [emailConfirmed, setEmailConfirmed] = useState(false);
  const [rememberEmail, setRememberEmail] = useState(false);
  const [emailError, setEmailError] = useState("");
  const [emailOtpCode, setEmailOtpCode] = useState("");
  const [emailCodeSent, setEmailCodeSent] = useState(false);
  const [emailMessage, setEmailMessage] = useState("");
  const [otpLoading, setOtpLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [viewMode, setViewMode] = useState("all");
  const [productCostReportSearch, setProductCostReportSearch] = useState("");
  const [productReportView, setProductReportView] = useState("main");
  const [showProductMissingReport, setShowProductMissingReport] = useState(false);
  const [productMissingReportRows, setProductMissingReportRows] = useState([]);
  const [productMissingReportLoading, setProductMissingReportLoading] = useState(false);
  const [productMissingReportMessage, setProductMissingReportMessage] = useState("");
  const [nextOrderRows, setNextOrderRows] = useState([]);
  const [nextOrderSourceRows, setNextOrderSourceRows] = useState([]);
  const [nextOrderMeta, setNextOrderMeta] = useState({});
  const [nextOrderFileName, setNextOrderFileName] = useState("");
  const [nextOrderTemplateFileName, setNextOrderTemplateFileName] = useState("");
  const [nextOrderSearch, setNextOrderSearch] = useState("");
  const [nextOrderFilter, setNextOrderFilter] = useState("all");
  const [nextOrderLoading, setNextOrderLoading] = useState(false);
  const [nextOrderMessage, setNextOrderMessage] = useState("");
  const [nextOrderView, setNextOrderView] = useState("order");
  const [fmlMissingRows, setFmlMissingRows] = useState([]);
const [fmlMissingSearch, setFmlMissingSearch] = useState("");
const [fmlLowRows, setFmlLowRows] = useState([]);
const [fmlLowSearch, setFmlLowSearch] = useState("");

const [yearlyRegionalConsumption, setYearlyRegionalConsumption] = useState(null);
const [yearlyRegionalFileName, setYearlyRegionalFileName] = useState("");
const [yearlyRegionalMessage, setYearlyRegionalMessage] = useState(
  "Yearly regional file will load when Product tools open."
);
const [selectedRegionalConsumptionRegion, setSelectedRegionalConsumptionRegion] =
  useState("");
const [regionalParBufferPercent, setRegionalParBufferPercent] = useState(0);
  const [yearlyRegionalReportSearch, setYearlyRegionalReportSearch] = useState("");
const [yearlyRegionalReportRegion, setYearlyRegionalReportRegion] =
  useState(YEARLY_REGION_ALL);
const [yearlyRegionalReportShip, setYearlyRegionalReportShip] =
  useState(YEARLY_REPORT_ALL_SHIPS);
const [yearlyRegionalReportMonths, setYearlyRegionalReportMonths] = useState([]);
const [yearlyRegionalReportSort, setYearlyRegionalReportSort] = useState("qty");

const [module, setModule] = useState("");
  const [productMode, setProductMode] = useState("");
  const [equipmentDepartment, setEquipmentDepartment] = useState("");
  const [equipmentMode, setEquipmentMode] = useState("");

  const [musterItems, setMusterItems] = useState([]);
  const [musterMessage, setMusterMessage] = useState("");
  const [selectedEquipment, setSelectedEquipment] = useState(null);

  const [warehouseRows, setWarehouseRows] = useState([]);
  const [warehouseSearch, setWarehouseSearch] = useState("");
  const [warehouseFilter, setWarehouseFilter] = useState("all");
  const [warehouseMessage, setWarehouseMessage] = useState("");

  const [inUseRows, setInUseRows] = useState([]);
  const [inUseSearch, setInUseSearch] = useState("");
  const [inUseMessage, setInUseMessage] = useState("");

  const [makeInventoryItems, setMakeInventoryItems] = useState([]);
  const [masterInventorySource, setMasterInventorySource] = useState("");
  const [masterInventoryLoading, setMasterInventoryLoading] = useState(false);
  const [makeInventorySearch, setMakeInventorySearch] = useState("");
  const [makeInventoryMessage, setMakeInventoryMessage] = useState("");
  const [makeInventoryShip, setMakeInventoryShip] = useState("");
  const [inventoryStation, setInventoryStation] = useState("");
  const [inventoryUserName, setInventoryUserName] = useState("");
  const [inventoryUserPosition, setInventoryUserPosition] = useState("");
  const [currentInventoryItem, setCurrentInventoryItem] = useState(null);
  const [inventoryQty, setInventoryQty] = useState("");
  const [editingInventoryId, setEditingInventoryId] = useState(null);
  const [extraInventoryCode, setExtraInventoryCode] = useState("");
const [extraInventoryName, setExtraInventoryName] = useState("");
const [extraInventoryQty, setExtraInventoryQty] = useState("");
const [extraInventoryPhotoFile, setExtraInventoryPhotoFile] = useState(null);
const [extraInventoryPhotoInputKey, setExtraInventoryPhotoInputKey] = useState(0);
const [extraInventorySaving, setExtraInventorySaving] = useState(false);
const [extraInventoryMessage, setExtraInventoryMessage] = useState("");
  const [inventorySummary, setInventorySummary] = useState([]);
  const [inventoryReportMode, setInventoryReportMode] = useState("my");
  const [summaryStationFilter, setSummaryStationFilter] = useState("ALL");
  const [inventoryLoading, setInventoryLoading] = useState(false);
  const [inventoryError, setInventoryError] = useState("");
  const [showVariance, setShowVariance] = useState(false);
  const [reportBusy, setReportBusy] = useState(false);
  const [inventoryStationStatuses, setInventoryStationStatuses] = useState([]);
  const [inventoryCountSheetTemplateFile, setInventoryCountSheetTemplateFile] = useState(null);
const [inventoryCountSheetTemplateName, setInventoryCountSheetTemplateName] = useState("");
    const [pictureLibraryBusy, setPictureLibraryBusy] = useState(false);
  const [pictureLibraryMessage, setPictureLibraryMessage] = useState("");
  const [drivePictureLibraryByCode, setDrivePictureLibraryByCode] = useState({});

    const realtimeRefreshTimersRef = useRef({});
  const printBusyRef = useRef(false);
  const saveBusyRef = useRef(false);
  const defaultTemplateLoadedRef = useRef(false);
const yearlyRegionalLoadedRef = useRef(false);

  const browserHistoryReadyRef = useRef(false);
  const browserHistoryRestoringRef = useRef(false);
  const browserHistoryLastKeyRef = useRef("");
    const getCurrentAppNavigationState = () => ({
    welcomeStarted,
    emailConfirmed,
    loggedIn,
    userShip,
    module,
    productMode,
    equipmentDepartment,
    equipmentMode,
    productReportView,
  });

  const applyAppNavigationState = (navigationState = {}) => {
    setWelcomeStarted(Boolean(navigationState.welcomeStarted));
    setEmailConfirmed(Boolean(navigationState.emailConfirmed));
    setLoggedIn(Boolean(navigationState.loggedIn));
    setUserShip(navigationState.userShip || "");
    setModule(navigationState.module || "");
    setProductMode(navigationState.productMode || "");
    setEquipmentDepartment(navigationState.equipmentDepartment || "");
    setEquipmentMode(navigationState.equipmentMode || "");
    setProductReportView(navigationState.productReportView || "main");

    // Close detail screens/modals that should not survive a browser-back jump.
    setSelectedProduct("");
    setSelectedRecipe(null);
    setSelectedEquipment(null);
    setCurrentInventoryItem(null);
    setInventoryQty("");
    setEditingInventoryId(null);
  };
    useEffect(() => {
    if (typeof window === "undefined") return;

    const handleBrowserBackForward = (event) => {
      const navigationState = getAppNavigationStateFromHistory(event.state);

      if (!navigationState) {
        return;
      }

      browserHistoryRestoringRef.current = true;
      browserHistoryLastKeyRef.current =
        getAppNavigationStateKey(navigationState);

      applyAppNavigationState(navigationState);
    };

    window.addEventListener("popstate", handleBrowserBackForward);

    return () => {
      window.removeEventListener("popstate", handleBrowserBackForward);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
    useEffect(() => {
    if (typeof window === "undefined") return;

    const navigationState = getCurrentAppNavigationState();
    const navigationStateKey = getAppNavigationStateKey(navigationState);

    if (!browserHistoryReadyRef.current) {
      browserHistoryReadyRef.current = true;
      browserHistoryLastKeyRef.current = navigationStateKey;

      window.history.replaceState(
        buildAppBrowserHistoryState(window.history.state, navigationState),
        "",
        window.location.href
      );

      return;
    }

    if (browserHistoryRestoringRef.current) {
      browserHistoryRestoringRef.current = false;
      browserHistoryLastKeyRef.current = navigationStateKey;
      return;
    }

    if (navigationStateKey === browserHistoryLastKeyRef.current) {
      return;
    }

    browserHistoryLastKeyRef.current = navigationStateKey;

    window.history.pushState(
      buildAppBrowserHistoryState(window.history.state, navigationState),
      "",
      window.location.href
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    welcomeStarted,
    emailConfirmed,
    loggedIn,
    userShip,
    module,
    productMode,
    equipmentDepartment,
    equipmentMode,
    productReportView,
  ]);

  const shipColumns = { BRL: 8, RL: 11, SC: 14, VL: 17 };
  const shipCostColumns = { BRL: 9, RL: 12, SC: 15, VL: 18 };
  const shipUnitPriceColumns = { BRL: 7, RL: 10, SC: 13, VL: 16 };

  const logUsageEvent = async (eventType, details = {}) => {
    try {
      if (!supabase || !eventType) return;

      const activeUserEmail = normalizeAppEmail(details.userEmail || userEmail || "");

      const payload = {
        event_type: eventType,
        ship: details.ship || makeInventoryShip || userShip || "",
        module: details.module || module || "",
        station: details.station || inventoryStation || "",
        user_name: details.userName || getEffectiveInventoryUserName?.() || "",
        user_position: details.userPosition || inventoryUserPosition || "",
        user_email: activeUserEmail,
        visitor_id: getOrCreateVisitorId(),
        page_path: typeof window !== "undefined" ? window.location.pathname : "",
        user_agent: typeof navigator !== "undefined" ? navigator.userAgent : "",
        details: {
          ...details,
          userEmail: activeUserEmail || details.userEmail || "",
        },
      };

      await supabase.from("app_usage_logs").insert(payload);
    } catch {
      // Tracking should never block the app.
    }
  };

  useEffect(() => {
    if (typeof window === "undefined") return;

    const savedEmail = normalizeAppEmail(window.localStorage.getItem(USER_EMAIL_STORAGE_KEY));
    if (isVirginVoyagesEmail(savedEmail)) {
      setUserEmail(savedEmail);
      setRememberEmail(true);
      setEmailConfirmed(true);
      logUsageEvent("remembered_email_loaded", { module: "welcome", userEmail: savedEmail });
    }
  }, []);

  useEffect(() => {
    logUsageEvent("app_opened", { module: "welcome" });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    setMakeInventoryShip(userShip);
  }, [userShip]);

  useEffect(() => {
    if (module !== "equipment") return;

    if (equipmentMode === "muster" || equipmentMode === "makeinventory" || equipmentMode === "warehouse") {
      loadMasterInventoryItems(makeInventoryShip || userShip);
    }

    if (equipmentMode === "makeinventory" && makeInventoryShip) {
  loadInventoryRecords(makeInventoryShip);
  loadInventoryStationStatuses(makeInventoryShip);
}
  }, [module, equipmentMode, makeInventoryShip, userShip]);

  useEffect(() => {
    if (!supabase || module !== "equipment" || (equipmentMode !== "makeinventory" && equipmentMode !== "muster" && equipmentMode !== "warehouse")) return;

    const channels = [];

    const masterScope = getMasterInventoryScope(equipmentDepartment);

    const masterChannel = supabase
      .channel("inventory-master-" + masterScope)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "inventory_master_items",
          filter: "ship=eq." + masterScope,
        },
        () => {
          scheduleRealtimeRefresh("master", makeInventoryShip || userShip);
        }
      )
      .subscribe();

    channels.push(masterChannel);

    if (equipmentMode === "makeinventory" && makeInventoryShip) {
      const countsChannel = supabase
        .channel("inventory-counts-" + makeInventoryShip)
        .on(
          "postgres_changes",
          {
            event: "*",
            schema: "public",
            table: "inventory_counts",
            filter: "ship=eq." + makeInventoryShip,
          },
          (payload) => {
  if (printBusyRef.current) return;

  if (payload.eventType === "DELETE") {
    const deletedId = payload.old?.id;

    setInventorySummary((prev) =>
      prev.filter((item) => item.id !== deletedId)
    );

    return;
  }

  if (!payload.new) return;

  const changedRecord = normalizeInventoryRecord(payload.new);

  if (!inventoryRecordMatchesCurrentDepartment(changedRecord)) return;

  setInventorySummary((prev) => {
    const withoutChanged = prev.filter(
      (item) =>
        item.id !== changedRecord.id &&
        !(
          item.ship === changedRecord.ship &&
          item.station === changedRecord.station &&
          item.userName === changedRecord.userName &&
          item.itemKey === changedRecord.itemKey
        )
    );

    return [changedRecord, ...withoutChanged];
  });
}
        )
        .subscribe();

      channels.push(countsChannel);
    }
          const stationStatusChannel = supabase
        .channel("inventory-station-status-" + makeInventoryShip + "-" + getMasterInventoryScope(equipmentDepartment))
        .on(
          "postgres_changes",
          {
            event: "*",
            schema: "public",
            table: "inventory_station_status",
            filter: "ship=eq." + makeInventoryShip,
          },
          (payload) => {
  if (printBusyRef.current) return;

  const departmentKey = getCurrentEquipmentDepartmentKey();

  if (payload.eventType === "DELETE") {
    const oldRecord = payload.old || {};

    setInventoryStationStatuses((prev) =>
      prev.filter(
        (item) =>
          !(
            item.id === oldRecord.id ||
            (
              item.ship === oldRecord.ship &&
              item.department === oldRecord.department &&
              cleanText(item.station) === cleanText(oldRecord.station)
            )
          )
      )
    );

    return;
  }

  if (!payload.new) return;

  const normalized = normalizeInventoryStationStatusRecord(payload.new);

  if (normalized.department !== departmentKey) return;

  setInventoryStationStatuses((prev) => {
    const withoutCurrent = prev.filter(
      (item) =>
        !(
          item.id === normalized.id ||
          (
            item.ship === normalized.ship &&
            item.department === normalized.department &&
            cleanText(item.station) === cleanText(normalized.station)
          )
        )
    );

    return [...withoutCurrent, normalized].sort((a, b) =>
      a.station.localeCompare(b.station)
    );
  });
}
        )
        .subscribe();

      channels.push(stationStatusChannel);

    return () => {
      channels.forEach((channel) => supabase.removeChannel(channel));
    };
  }, [makeInventoryShip, userShip, module, equipmentMode, equipmentDepartment]);

  useEffect(() => {
    return () => {
      Object.values(realtimeRefreshTimersRef.current).forEach((timer) => {
        window.clearTimeout(timer);
      });
    };
  }, []);

    const scheduleRealtimeRefresh = (type, shipOverride) => {
    if (printBusyRef.current) return;

    const ship = shipOverride || makeInventoryShip || userShip;

    const key =
      type === "master"
        ? "master"
        : type === "stationStatus"
          ? "station-status-" + (ship || "unknown")
          : "counts-" + (ship || "unknown");

    if (realtimeRefreshTimersRef.current[key]) {
      window.clearTimeout(realtimeRefreshTimersRef.current[key]);
    }

    realtimeRefreshTimersRef.current[key] = window.setTimeout(() => {
      delete realtimeRefreshTimersRef.current[key];

      if (type === "master") {
        loadMasterInventoryItems(ship);
        return;
      }

      if (type === "counts" && ship) {
        loadInventoryRecords(ship);
        return;
      }

      if (type === "stationStatus" && ship) {
        loadInventoryStationStatuses(ship);
      }
    }, 900);
  };

    const visibleShips = viewMode === "single" ? [userShip] : SHIPS;

const buildProductList = (rows) =>
  [
    ...new Set(
      rows
        .slice(1)
        .map((r) => String(r[6] || "").trim())
        .filter(Boolean)
    ),
  ].sort();

const uploadNextOrderFile = (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  readExcelFile(file, (workbook) => {
    try {
      const parsed = parseNextOrderWorkbook({
        workbook,
        templateMap,
      });

      setNextOrderFileName(file.name);
      setNextOrderSourceRows(parsed.rows);
      setNextOrderMeta(parsed.meta);
      setNextOrderRows([]);
      setFmlMissingRows(parsed.fmlReportRows || []);
      setFmlLowRows(parsed.fmlRunningLowRows || []);
      setNextOrderSearch("");
      setFmlMissingSearch("");
      setFmlLowSearch("");
      setNextOrderFilter("all");
      setNextOrderView("order");

      setNextOrderMessage(
        "Order file loaded. " +
          parsed.meta.totalItems +
          " product rows found. " +
          parsed.meta.itemsNeedingOrder +
          " need order, " +
          parsed.meta.parCapItems +
          " par cap, " +
          parsed.meta.blueReviewItems +
          " blue review, " +
          parsed.meta.redReviewItems +
          " red review, " +
          parsed.meta.fmlMissingItems +
          " FML not ordered/not used, " +
          parsed.meta.fmlRunningLowItems +
          " FML running low."
      );

      logUsageEvent("next_order_file_uploaded", {
        module: "generate_next_order",
        fileName: file.name,
        sheetName: parsed.meta.sheetName,
        totalItems: parsed.meta.totalItems,
        itemsNeedingOrder: parsed.meta.itemsNeedingOrder,
        parCapItems: parsed.meta.parCapItems,
        blueReviewItems: parsed.meta.blueReviewItems,
        redReviewItems: parsed.meta.redReviewItems,
        fmlMissingItems: parsed.meta.fmlMissingItems,
        fmlRunningLowItems: parsed.meta.fmlRunningLowItems,
      });
    } catch (error) {
      setNextOrderFileName(file.name);
      setNextOrderSourceRows([]);
      setNextOrderMeta({});
      setNextOrderRows([]);
      setNextOrderSearch("");
      setNextOrderFilter("all");
      setNextOrderMessage(error?.message || "Could not read the order file.");
    }
  });
};

  const loadDefaultTemplate = async () => {
  try {
    const XLSX = await loadXlsx();

    const response = await fetch("/template.xlsx");
      if (!response.ok) {
        setTemplateStatus("Template file not found.");
        return;
      }

      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      const parsedTemplate = parseTemplateWorkbook(workbook);
      setTemplateMap(parsedTemplate);
      setTemplateStatus("Default ERP template loaded.");
      setNextOrderTemplateFileName((current) => current || "Default ERP Food ordering template");
    } catch {
      setTemplateStatus("Could not load template.");
    }
  };

  const parseWarehouseItems = () => {
    const masterItems = makeInventoryItems.length ? makeInventoryItems : musterItems;
    const compactText = (value) => cleanText(value).replace(/[^A-Z0-9]/g, "");
    const masterByCode = {};
    const masterByName = {};

    masterItems.forEach((masterItem) => {
      if (!masterItem?.image) return;

      const codeKey = cleanText(masterItem.code);
      const nameKey = compactText(masterItem.name);

      if (codeKey && !masterByCode[codeKey]) masterByCode[codeKey] = masterItem;
      if (nameKey && !masterByName[nameKey]) masterByName[nameKey] = masterItem;
    });

    const tokenizeName = (value) =>
      cleanText(value)
        .replace(/[^A-Z0-9 ]/g, " ")
        .split(" ")
        .map((word) => word.trim())
        .filter((word) => word.length >= 3 && !["FOR", "THE", "AND", "WITH", "EA", "PCS"].includes(word));

    const findPictureMatch = (code, name) => {
      const codeKey = cleanText(code);
      const nameKey = compactText(name);
      const nameTokens = new Set(tokenizeName(name));

      if (codeKey && masterByCode[codeKey]) return masterByCode[codeKey];
      if (nameKey && masterByName[nameKey]) return masterByName[nameKey];

      let bestMatch = null;
      let bestScore = 0;

      masterItems.forEach((masterItem) => {
        if (!masterItem?.image) return;

        const masterCodeKey = cleanText(masterItem.code);
        const masterNameKey = compactText(masterItem.name);
        const masterTokens = tokenizeName(masterItem.name);

        let score = 0;

        if (masterCodeKey && nameKey.includes(masterCodeKey)) score += 80;
        if (masterNameKey.length >= 8 && nameKey.includes(masterNameKey)) score += 70;
        if (masterNameKey.length >= 8 && masterNameKey.includes(nameKey)) score += 65;

        const sharedTokens = masterTokens.filter((token) => nameTokens.has(token));
        const tokenScore = sharedTokens.length * 12;
        score += tokenScore;

        if (score > bestScore) {
          bestScore = score;
          bestMatch = masterItem;
        }
      });

      return bestScore >= 36 ? bestMatch : null;
    };

    return warehouseRows
      .slice(1)
      .map((row) => {
        const code = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        const par = Number(row[6] || 0);
        const onHand = Number(row[7] || 0);
        const future = Number(row[12] || 0);
        const suggested = Math.max(par - onHand - future, 0);
        const pictureMatch = findPictureMatch(code, name);

        return {
          code,
          name,
          par,
          onHand,
          future,
          suggested,
          image: pictureMatch?.image || "",
          imageSource: pictureMatch?.name || "",
          masterSheetName: pictureMatch?.sheetName || "",
          masterCategory: pictureMatch?.category || "",
        };
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
  const searchText = makeInventorySearch.toLowerCase();
  const selectedRestaurantStation = cleanText(inventoryStation);

  return makeInventoryItems.filter((item) => {
    const matchesSearch = `${item.sheetName} ${item.category} ${item.code} ${item.name}`
      .toLowerCase()
      .includes(searchText);

    if (!matchesSearch) return false;

    if (equipmentDepartment !== "restaurant") {
      return true;
    }

    if (!selectedRestaurantStation) {
      return true;
    }

    const itemStation = cleanText(item.stationName || item.sheetName || "");
    const isMasterItem =
      itemStation === "MASTER" ||
      cleanText(item.category) === "MASTER" ||
      cleanText(item.sheetName).includes("MASTER");

    if (isMasterItem) {
      return false;
    }

    return itemStation === selectedRestaurantStation;
  });
};

  const getEffectiveInventoryUserName = () => {
    const name = inventoryUserName.trim();
    const position = inventoryUserPosition.trim();

    if (!name) return "";
    return position ? `${name} - ${position}` : name;
  };

  const getInventoryItemKey = (item) => {
    const departmentKey = cleanText(item?.equipmentDepartment || equipmentDepartment || "culinary")
      .replace(/[^A-Z0-9]/g, "_");

    return cleanText(
      departmentKey + "|" +
        (item?.sheetName || "") + "|" +
        (item?.category || "") + "|" +
        (item?.code || "") + "|" +
        (item?.name || "") + "|" +
        (item?.sourceRow || "")
    );
  };

  const getRestaurantStationList = () => {
  const sourceItems = makeInventoryItems.length ? makeInventoryItems : musterItems;

  const stationsFromTabs = [
    ...new Set(
      sourceItems
        .filter((item) => item?.equipmentDepartment === "restaurant")
        .filter((item) => item?.isVenueSheet !== false)
        .map((item) =>
          String(item.stationName || item.sheetName || "").trim()
        )
        .filter(Boolean)
        .filter((sheetName) => {
          const text = cleanText(sheetName);

          return (
            text !== "MASTER" &&
            !text.includes("MASTER") &&
            !text.includes("SUMMARY") &&
            !text.includes("INDEX") &&
            !text.includes("COVER")
          );
        })
    ),
  ];

  return stationsFromTabs.length ? stationsFromTabs.sort() : ["Restaurant"];
};

const getRestaurantInventoryStationList = () => {
  const sourceItems = makeInventoryItems.length ? makeInventoryItems : musterItems;

  return [
    ...new Set(
      sourceItems
        .filter(
          (item) =>
            String(item?.equipmentDepartment || "").toLowerCase() ===
            "restaurant"
        )
        .map((item) => item.stationName || item.sheetName || "")
        .map((station) => String(station || "").replace(/\s+/g, " ").trim())
        .filter(Boolean)
        .filter((station) => cleanText(station) !== "MASTER")
    ),
  ].sort((a, b) => a.localeCompare(b));
};

const getActiveInventoryStationList = () => {
  if (equipmentDepartment === "bar") return BAR_STATIONS;

  if (equipmentDepartment === "restaurant") {
    return getRestaurantInventoryStationList();
  }

  return CULINARY_STATIONS;
};
  const getInventoryStationLabel = () => {
  if (equipmentDepartment === "bar") return "bar";
  if (equipmentDepartment === "restaurant") return "restaurant";

  return "station";
};

  const loadMasterInventoryItems = async (shipOverride) => {
  if (!supabase) {
    const text =
      "Supabase is not connected. Shared master inventory cannot load.";

    setMakeInventoryMessage(text);
    setMusterMessage(text);
    setMasterInventorySource("Supabase is not connected.");
    return;
  }

  const activeDepartment = equipmentDepartment || "culinary";
  const masterScope = getMasterInventoryScope(activeDepartment);
  const departmentLabel = getEquipmentDepartmentLabelForState(activeDepartment);

  setMasterInventoryLoading(true);
  setMakeInventoryMessage("Loading shared " + departmentLabel + " master list...");

  const loadForScope = async (scope) =>
    supabase
      .from("inventory_master_items")
      .select("*")
      .eq("ship", scope)
      .order("sort_order", { ascending: true });

  try {
    let { data, error } = await loadForScope(masterScope);
    let sourceText =
      "Shared " + departmentLabel + " master list loaded for all users.";

    if (error) {
      throw error;
    }

    if ((!data || data.length === 0) && shipOverride) {
      const legacyResult = await loadForScope(shipOverride);

      if (!legacyResult.error && legacyResult.data?.length) {
        data = legacyResult.data;
        sourceText =
          "Legacy shared master list loaded for " +
shipOverride +
". Upload again from Master List to make it shared by department.";
      }
    }

    let items = (data || []).map(normalizeMasterInventoryRecord);

    setMakeInventoryItems(items);
    setMusterItems(items);
    setMasterInventorySource(
      items.length ? sourceText : "No shared master list uploaded yet."
    );
    setMusterMessage(
      items.length
        ? sourceText + " " + items.length + " item(s) available."
        : "No shared master list uploaded yet."
    );
    setMakeInventoryMessage(
      items.length
        ? sourceText + " " + items.length + " item(s) loaded."
        : "No shared master list uploaded yet."
    );

    // Important: picture sync happens AFTER the master list loads.
    // If picture sync fails, the equipment list still stays loaded.
    if (activeDepartment === "culinary" && items.length > 0) {
      try {
        const pictureMap =
          Object.keys(drivePictureLibraryByCode || {}).length > 0
            ? drivePictureLibraryByCode
            : await loadDrivePictureLibrary({ silent: true });

        if (Object.keys(pictureMap || {}).length > 0) {
          const itemsWithPictures = attachPicturesFromLibraryToItems(
            items,
            pictureMap
          );

          const matchedPictureCount = itemsWithPictures.filter((item) => {
            const match = getPictureMatchForItemFromMap(item, pictureMap);
            return Boolean(match.pictureUrl);
          }).length;

          setMakeInventoryItems(itemsWithPictures);
          setMusterItems(itemsWithPictures);

          setMakeInventoryMessage(
            sourceText +
              " " +
              itemsWithPictures.length +
              " item(s) loaded. Pictures matched: " +
              matchedPictureCount +
              "."
          );

          setMusterMessage(
            sourceText +
              " " +
              itemsWithPictures.length +
              " item(s) loaded. Pictures matched: " +
              matchedPictureCount +
              "."
          );
        }
      } catch (pictureError) {
        setPictureLibraryMessage(
          "Master list loaded, but pictures could not sync: " +
            (pictureError?.message || "Unknown picture sync error")
        );
      }
    }
  } catch (error) {
    const text =
      "Could not load shared " +
      departmentLabel +
      " master inventory: " +
      (error?.message || "Unknown error");

    setMakeInventoryItems([]);
    setMusterItems([]);
    setMasterInventorySource("Master list failed to load.");
    setMakeInventoryMessage(text);
    setMusterMessage(text);
    window.alert(text);
  } finally {
    setMasterInventoryLoading(false);
  }
};
const uploadEquipmentDataImageToStorage = async ({ dataUrl, scope, item, index }) => {
  const value = String(dataUrl || "").trim();

  if (!supabase || !value.startsWith("data:image/")) {
    return "";
  }

  const mime = getImageMimeFromDataUrl(value);
  const extension = getImageExtensionFromMime(mime);

  const codePart = makeStorageSafePart(item?.code || item?.name || index + 1);
  const rowPart = makeStorageSafePart(item?.sourceRow || index + 1);
  const path = `${scope}/${codePart}-${rowPart}.${extension}`;

  const blob = await fetch(value).then((response) => response.blob());

  const { error } = await supabase.storage
    .from(EQUIPMENT_PICTURE_BUCKET)
    .upload(path, blob, {
      contentType: mime,
      upsert: true,
      cacheControl: "31536000",
    });

  if (error) {
    throw error;
  }

  const { data } = supabase.storage
    .from(EQUIPMENT_PICTURE_BUCKET)
    .getPublicUrl(path);

  return data?.publicUrl || "";
};

  const uploadExtraInventoryPhotoToStorage = async ({
  file,
  ship,
  station,
  code,
  name,
}) => {
  if (!supabase || !file) return "";

  const fileName = String(file.name || "").toLowerCase();
  const fileExtension =
    fileName.endsWith(".jpg") || fileName.endsWith(".jpeg")
      ? "jpg"
      : fileName.endsWith(".webp")
      ? "webp"
      : fileName.endsWith(".gif")
      ? "gif"
      : "png";

  const contentType = file.type || "image/" + fileExtension;
  const scope = getMasterInventoryScope(equipmentDepartment);
  const stationPart = makeStorageSafePart(station || "station");
  const itemPart = makeStorageSafePart(code || name || "extra-item");

  const path =
    scope +
    "/extra-inventory/" +
    makeStorageSafePart(ship || "ship") +
    "/" +
    stationPart +
    "/" +
    Date.now() +
    "-" +
    itemPart +
    "." +
    fileExtension;

  const { error } = await supabase.storage
    .from(EQUIPMENT_PICTURE_BUCKET)
    .upload(path, file, {
      contentType,
      upsert: true,
      cacheControl: "31536000",
    });

  if (error) {
    throw error;
  }

  const { data } = supabase.storage
    .from(EQUIPMENT_PICTURE_BUCKET)
    .getPublicUrl(path);

  return data?.publicUrl || "";
};

const getPersistentEquipmentImageUrl = async (item, scope, index) => {
  const imageCandidates = [
    item?.image,
    item?.imageFallback,
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  const directUrl = imageCandidates.find(
    (value) =>
      !value.startsWith("data:image/") &&
      isUsableImageValue(value) &&
      value.length <= 5000
  );

  if (directUrl) {
    return directUrl;
  }

  const dataUrl = imageCandidates.find((value) => value.startsWith("data:image/"));

  if (dataUrl) {
    return uploadEquipmentDataImageToStorage({
      dataUrl,
      scope,
      item,
      index,
    });
  }

  return "";
};

  const getEquipmentDepartmentLabelForState = (department) => {
  const labels = {
    culinary: "Culinary",
    bar: "Bar",
    restaurant: "Restaurant",
  };

  return labels[department] || "Equipment";
};

const getInventoryPictureCandidateCodesForSync = (item) => {
  const candidates = [];

  const addCandidate = (value) => {
    const rawValue = String(value || "").trim();
    if (!rawValue) return;

    const numericKey = normalizeEquipmentPictureCode(rawValue);
    const looseKey = getEquipmentImageMatchCode(rawValue);

    [numericKey, looseKey].forEach((key) => {
      const cleanKey = String(key || "").trim();
      if (!cleanKey) return;
      if (candidates.includes(cleanKey)) return;

      candidates.push(cleanKey);
    });

    const numberMatches = rawValue.match(/\d{4,}/g) || [];

    numberMatches.forEach((numberText) => {
      const cleanNumber = String(numberText || "")
        .replace(/^0+(?=\d)/g, "")
        .trim();

      if (!cleanNumber) return;
      if (candidates.includes(cleanNumber)) return;

      candidates.push(cleanNumber);
    });
  };

  addCandidate(item?.code);
  addCandidate(item?.name);
  addCandidate(item?.pictureFileName);
  addCandidate(item?.pictureMatchCode);
  addCandidate(item?.category);
  addCandidate(item?.sheetName);

  return candidates;
};

const getPictureMatchForItemFromMap = (item, pictureMap = {}) => {
  const candidateCodes = getInventoryPictureCandidateCodesForSync(item);

  for (const codeKey of candidateCodes) {
    const pictureUrl = pictureMap[codeKey];

    if (pictureUrl) {
      return {
        codeKey,
        pictureUrl,
      };
    }
  }

  return {
    codeKey: "",
    pictureUrl: "",
  };
};

const attachPicturesFromLibraryToItems = (items = [], pictureMap = {}) => {
  return (items || []).map((item) => {
    const match = getPictureMatchForItemFromMap(item, pictureMap);

    if (!match.pictureUrl) {
      return item;
    }

    const currentImage = String(item.image || "").trim();
    const currentFallback = String(item.imageFallback || "").trim();

    const shouldReplaceCurrentImage =
      !currentImage || isTemporaryGoogleThumbnail(currentImage);

    return {
      ...item,
      image: shouldReplaceCurrentImage ? match.pictureUrl : currentImage,
      imageFallback:
        !shouldReplaceCurrentImage && currentImage !== match.pictureUrl
          ? currentFallback || match.pictureUrl
          : currentFallback || currentImage || "",
      pictureMatchCode: match.codeKey,
    };
  });
};
const loadExistingMasterImagesByCode = async (scope) => {
  const imageByCode = {};

  if (!supabase || !scope) {
    return imageByCode;
  }

  const { data, error } = await supabase
    .from("inventory_master_items")
    .select("code,image")
    .eq("ship", scope);

  if (error) {
    return imageByCode;
  }

  (data || []).forEach((row) => {
    const codeKey = getEquipmentImageMatchCode(row.code);
    const image = String(row.image || "").trim();

    if (codeKey && image && !imageByCode[codeKey]) {
      imageByCode[codeKey] = image;
    }
  });

  return imageByCode;
};

  const deleteMasterInventoryRowsInBatches = async (scope) => {
    const batchSize = 100;

    while (true) {
      const { data, error } = await supabase
        .from("inventory_master_items")
        .select("id")
        .eq("ship", scope)
        .limit(batchSize);

      if (error) {
        throw error;
      }

      const ids = (data || []).map((row) => row.id).filter(Boolean);

      if (!ids.length) {
        break;
      }

      const deleteResult = await supabase
        .from("inventory_master_items")
        .delete()
        .in("id", ids);

      if (deleteResult.error) {
        throw deleteResult.error;
      }

      if (ids.length < batchSize) {
        break;
      }
    }
  };
  const saveMasterInventoryItems = async (_shipOverride, items) => {
    if (!supabase) {
      const text = "Supabase is not connected. Cannot share the master inventory file.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return false;
    }

    if (!items.length) {
      const text = "No inventory items were found in the uploaded file.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return false;
    }

    setMasterInventoryLoading(true);
    const masterScope = getMasterInventoryScope(equipmentDepartment);
    const departmentLabel = activeEquipmentDepartmentLabel || "Equipment";

    setMakeInventoryMessage("Sharing " + departmentLabel + " master inventory list for all users...");
    setMusterMessage("Sharing " + departmentLabel + " master inventory list for all users...");
    const existingImageByCode = await loadExistingMasterImagesByCode(masterScope);

    try {
  await deleteMasterInventoryRowsInBatches(masterScope);
} catch (error) {
  const text = `Could not replace shared master inventory: ${error.message}`;
  setMasterInventoryLoading(false);
  setMakeInventoryMessage(text);
  setMusterMessage(text);
  window.alert(text);
  return false;
}

    const rowMap = new Map();

for (let index = 0; index < items.length; index += 1) {
  const item = items[index];

  const normalizedItem = {
    ...item,
    equipmentDepartment: item.equipmentDepartment || equipmentDepartment || "culinary",
  };

  const itemKey = getInventoryItemKey(normalizedItem);
  if (!itemKey || rowMap.has(itemKey)) continue;

  if (index > 0 && index % 25 === 0) {
    setMakeInventoryMessage(
      `Saving ${departmentLabel} pictures and master list... ${index} of ${items.length}`
    );
  }

  const persistentImageUrl = await getPersistentEquipmentImageUrl(
    normalizedItem,
    masterScope,
    index
  );

  rowMap.set(itemKey, {
    ship: masterScope,
    item_key: itemKey,
    code: normalizedItem.code || "",
    item_name: normalizedItem.name || "",
    category: normalizedItem.category || "",
    sheet_name: normalizedItem.sheetName || "",
    image:
  persistentImageUrl ||
  cleanSharedMasterImage(normalizedItem.image) ||
  cleanSharedMasterImage(normalizedItem.imageFallback) ||
  existingImageByCode[getEquipmentImageMatchCode(normalizedItem.code)] ||
  "",
    source_row: Number(normalizedItem.sourceRow || index + 1),
    sort_order: index,
    updated_at: new Date().toISOString(),
  });
}
    const rows = [...rowMap.values()];
    const batchSize = 75;

    for (let i = 0; i < rows.length; i += batchSize) {
      const insertResult = await supabase
        .from("inventory_master_items")
        .insert(rows.slice(i, i + batchSize));

      if (insertResult.error) {
        const text = `Could not save shared master inventory: ${insertResult.error.message}`;
        setMasterInventoryLoading(false);
        setMakeInventoryMessage(text);
        setMusterMessage(text);
        window.alert(text);
        return false;
      }
    }

    setMakeInventoryItems(items);
    setMusterItems(items);
    setMasterInventorySource("Shared " + departmentLabel + " master list saved for all users. " + rows.length + " item(s) from all tabs.");
    setMasterInventoryLoading(false);
    setMakeInventoryMessage("Shared " + departmentLabel + " master list saved for all users. " + rows.length + " item(s) from all tabs.");
    setMusterMessage("Shared " + departmentLabel + " master list saved for all users. " + rows.length + " item(s) from all tabs.");
    return true;
  };

    const refreshMakeInventoryData = async (shipOverride) => {
    const ship = shipOverride || makeInventoryShip || userShip;

    await Promise.all([
      ship ? loadInventoryRecords(ship) : Promise.resolve(),
      ship ? loadInventoryStationStatuses(ship) : Promise.resolve(),
      loadMasterInventoryItems(ship),
    ]);
  };

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

  const getCurrentEquipmentDepartmentKey = () =>
  getEquipmentDepartmentKey(equipmentDepartment);

const inventoryRecordMatchesCurrentDepartment = (item) =>
  inventoryRecordMatchesDepartment(item, equipmentDepartment);

  const loadInventoryStationStatuses = async (shipOverride) => {
    const ship = shipOverride || makeInventoryShip || userShip;

    if (!ship) {
      setInventoryStationStatuses([]);
      return;
    }

    if (!supabase) {
      return;
    }

    const departmentKey = getCurrentEquipmentDepartmentKey();

    const { data, error } = await supabase
      .from("inventory_station_status")
      .select("*")
      .eq("ship", ship)
      .eq("department", departmentKey)
      .order("station", { ascending: true });

    if (error) {
      setMakeInventoryMessage(`Could not load station status: ${error.message}`);
      return;
    }

    setInventoryStationStatuses((data || []).map(normalizeInventoryStationStatusRecord));
  };

  const getInventoryStationProgressRows = (statusRows = inventoryStationStatuses) => {
    const ship = makeInventoryShip || userShip;
    const departmentKey = getCurrentEquipmentDepartmentKey();
    const activeStations = getActiveInventoryStationList();

    return activeStations.map((station) => {
      const statusRecord = statusRows.find(
        (item) =>
          item.ship === ship &&
          item.department === departmentKey &&
          cleanText(item.station) === cleanText(station)
      );

      const stationRows = inventorySummary.filter(
        (item) =>
          item.ship === ship &&
          cleanText(item.station) === cleanText(station) &&
          inventoryRecordMatchesCurrentDepartment(item)
      );

      const countedKeys = new Set(
        stationRows.map((item) => item.itemKey || cleanText(item.code || item.name))
      );

      const countedItems = countedKeys.size;
      const totalQty = stationRows.reduce((sum, item) => sum + Number(item.qty || 0), 0);

      const status =
        statusRecord?.status ||
        (countedItems > 0 ? "started" : "not_started");

      const statusLabel =
        status === "submitted"
          ? "Count Submitted"
          : status === "started"
            ? "Count Started"
            : "Not Started";

      return {
        station,
        status,
        statusLabel,
        countedItems,
        totalQty,
        userName: statusRecord?.userName || "",
        userPosition: statusRecord?.userPosition || "",
        startedAt: statusRecord?.startedAt || "",
        submittedAt: statusRecord?.submittedAt || "",
        updatedAt: statusRecord?.updatedAt || "",
      };
    });
  };

  const getCurrentStationProgress = () => {
    const station = inventoryStation;
    if (!station) return null;

    return getInventoryStationProgressRows().find(
      (item) => cleanText(item.station) === cleanText(station)
    );
  };

  const getAllInventoryStationsSubmitted = () => {
    const rows = getInventoryStationProgressRows();
    return rows.length > 0 && rows.every((item) => item.status === "submitted");
  };

  const upsertInventoryStationStatus = async (nextStatus) => {
    if (!supabase) {
      throw new Error("Supabase is not connected.");
    }

    const ship = makeInventoryShip || userShip;
    const station = inventoryStation;
    const userName = getEffectiveInventoryUserName();
    const departmentKey = getCurrentEquipmentDepartmentKey();

    if (!ship || !station || !userName) {
      throw new Error("Choose ship, station, and user before updating station status.");
    }

    const existing = getCurrentStationProgress();
    const now = new Date().toISOString();

    const finalStatus =
      existing?.status === "submitted" && nextStatus === "started"
        ? "submitted"
        : nextStatus;

    const payload = {
      ship,
      department: departmentKey,
      station,
      status: finalStatus,
      user_name: userName,
      user_position: inventoryUserPosition || "",
      updated_at: now,
    };

    if (!existing?.startedAt) {
      payload.started_at = now;
    }

    if (finalStatus === "submitted") {
      payload.submitted_at = now;
    }

    const { data, error } = await supabase
      .from("inventory_station_status")
      .upsert(payload, {
        onConflict: "ship,department,station",
      })
      .select("*")
      .single();

    if (error) {
      throw error;
    }

    const normalized = normalizeInventoryStationStatusRecord(data);

    setInventoryStationStatuses((prev) => {
      const withoutCurrent = prev.filter(
        (item) =>
          !(
            item.ship === normalized.ship &&
            item.department === normalized.department &&
            cleanText(item.station) === cleanText(normalized.station)
          )
      );

      return [...withoutCurrent, normalized].sort((a, b) =>
        a.station.localeCompare(b.station)
      );
    });

    return normalized;
  };

  const submitInventoryStationCount = async () => {
    if (reportBusy) return;

    const ship = makeInventoryShip || userShip;
    const userName = getEffectiveInventoryUserName();

    if (!ship || !inventoryStation || !userName) {
      const text = "Choose ship, station, and user before finishing inventory.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const myRows = getMyInventoryRows();

    const confirmed = window.confirm(
      `Submit inventory count for ${inventoryStation}? You counted ${myRows.length} item(s). After submit, this station will wait until all stations submit.`
    );

    if (!confirmed) return;

    setReportBusy(true);

    try {
      await upsertInventoryStationStatus("submitted");

      setMakeInventoryMessage(
        `${inventoryStation} - Count Submitted. Waiting until all stations submit.`
      );

      logUsageEvent("inventory_station_submitted", {
        module: "make_inventory",
        ship,
        station: inventoryStation,
        userName,
        userPosition: inventoryUserPosition,
        rows: myRows.length,
      });
    } catch (error) {
      const text = error?.message || "Could not submit station inventory.";
      setMakeInventoryMessage(text);
      window.alert(text);
    } finally {
      setReportBusy(false);
    }
  };

  const getFinalStationSummaryRows = () => {
    const ship = makeInventoryShip || userShip;
    const grouped = new Map();

    inventorySummary
      .filter(
        (item) =>
          item.ship === ship &&
          inventoryRecordMatchesCurrentDepartment(item)
      )
      .forEach((item) => {
        const productKey = getInventoryProductGroupKey(item);
        if (!productKey) return;

        const station = item.station || "Unknown Station";
        const key = `${cleanText(station)}__${productKey}`;
        const qty = Number(item.qty || 0);
        const safeQty = Number.isFinite(qty) ? qty : 0;

        if (!grouped.has(key)) {
          grouped.set(key, {
            station,
            code: item.code || "",
            name: item.name || "",
            category: item.category || "",
            sheetName: item.sheetName || "",
            count: 0,
            users: new Set(),
          });
        }

        const existing = grouped.get(key);
        existing.count += safeQty;

        if (item.userName) existing.users.add(item.userName);
      });

    return Array.from(grouped.values())
      .map((item) => ({
        ...item,
        users: [...item.users].sort(),
      }))
      .sort(
        (a, b) =>
          a.station.localeCompare(b.station) ||
          a.name.localeCompare(b.name)
      );
  };

  const generateFinalInventoryReport = async () => {
    if (reportBusy) return;

    if (!getAllInventoryStationsSubmitted()) {
      const text = "Final report is not ready yet. All stations must submit first.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    if (!inventoryCountSheetTemplateFile) {
      const text =
        "Upload the inventory sheet sample first. The final report will use that uploaded file and write counts into column S.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    setReportBusy(true);

    try {
      const rows = getSummaryInventoryRecordsForDownload();
      const stationSummaryRows = getFinalStationSummaryRows();

      const result = await downloadInventoryExcelReportUsingTemplate({
        templateFile: inventoryCountSheetTemplateFile,
        items: rows,
        venueName: getInventoryReportLocationName("summary"),
        reportTitle: "Final Inventory Report",
        stationSummaryRows,
      });

      setInventoryReportMode("summary");

      setMakeInventoryMessage(
        `Final report generated. Uploaded template positions stayed the same. ${result.matchedRows} of ${result.itemRows} item rows matched counts. Station Summary sheet was added.`
      );

      logUsageEvent("final_inventory_report_generated", {
        module: "make_inventory",
        ship: makeInventoryShip || userShip,
        rows: result.itemRows,
        matchedRows: result.matchedRows,
        stationSummaryRows: stationSummaryRows.length,
      });
    } catch (error) {
      const text = error?.message || "Could not generate final inventory report.";
      setMakeInventoryMessage(text);
      window.alert(text);
    } finally {
      setReportBusy(false);
    }
  };

  const getMyInventoryRows = () => {
    const ship = makeInventoryShip || userShip;
    const userName = getEffectiveInventoryUserName();

    return inventorySummary
      .filter(
        (item) =>
          item.ship === ship &&
          item.station === inventoryStation &&
          item.userName === userName &&
          inventoryRecordMatchesCurrentDepartment(item)
      )
      .sort((a, b) => a.name.localeCompare(b.name));
  };

  const getSummaryStationOptions = () => {
    const ship = makeInventoryShip || userShip;
    const stationsFromRecords = inventorySummary
      .filter((item) => item.ship === ship && item.station && inventoryRecordMatchesCurrentDepartment(item))
      .map((item) => item.station);

    const baseStations = getActiveInventoryStationList();

    return [...new Set([...baseStations, ...stationsFromRecords])]
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b));
  };

  const getShipSummaryRows = () => {
    const ship = makeInventoryShip || userShip;
    const grouped = {};
    const selectedStation = summaryStationFilter || "ALL";

    inventorySummary
      .filter(
        (item) =>
          item.ship === ship &&
          inventoryRecordMatchesCurrentDepartment(item) &&
          (selectedStation === "ALL" || item.station === selectedStation)
      )
      .forEach((item) => {
        const key = getInventoryProductGroupKey(item);
        if (!key) return;

        if (!grouped[key]) {
          grouped[key] = {
            itemKey: key,
            ship: item.ship,
            code: item.code || "",
            name: item.name || "",
            categorySet: new Set(),
            sheetSet: new Set(),
            totalQty: 0,
            count: 0,
            qty: 0,
            Quantity: 0,
            stations: new Set(),
            users: new Set(),
            recordCount: 0,
            lastUpdated: "",
          };
        }

        const qty = Number(item.qty || 0);

        grouped[key].totalQty += Number.isFinite(qty) ? qty : 0;
        grouped[key].count = grouped[key].totalQty;
        grouped[key].qty = grouped[key].totalQty;
        grouped[key].Quantity = grouped[key].totalQty;
        grouped[key].recordCount += 1;

        if (item.category) grouped[key].categorySet.add(item.category);
        if (item.sheetName) grouped[key].sheetSet.add(item.sheetName);
        if (item.station) grouped[key].stations.add(item.station);
        if (item.userName) grouped[key].users.add(item.userName);

        if (!grouped[key].lastUpdated || item.updatedAt > grouped[key].lastUpdated) {
          grouped[key].lastUpdated = item.updatedAt;
        }
      });

    return Object.values(grouped)
      .map((item) => {
        const totalQty = Number(item.totalQty || 0);

        return {
          ...item,
          category: [...item.categorySet].sort().join(", "),
          sheetName: [...item.sheetSet].sort().join(", "),
          stationFilter: selectedStation,
          stations: [...item.stations].sort(),
          users: [...item.users].sort(),
          confirmedAt: item.lastUpdated ? new Date(item.lastUpdated).toLocaleString() : "",
          count: totalQty,
          qty: totalQty,
          Quantity: totalQty,
        };
      })
      .sort((a, b) => b.totalQty - a.totalQty || a.name.localeCompare(b.name));
  };
  const getVisibleInventoryReportRows = () => {
    return inventoryReportMode === "summary" ? getShipSummaryRows() : getMyInventoryRows();
  };

      const getInventoryReportLocationName = (modeOverride = inventoryReportMode) => {
    const ship = makeInventoryShip || userShip || "ship";
    const department = activeEquipmentDepartmentLabel || equipmentDepartment || "equipment";

    if (modeOverride === "summary") {
      const stationLabel =
        summaryStationFilter === "ALL"
          ? equipmentDepartment === "bar"
            ? "All Bars"
            : "All Stations"
          : summaryStationFilter;

      return `${ship} - ${department} - ${stationLabel}`;
    }

    const userName = getEffectiveInventoryUserName() || "User";
    const stationName = inventoryStation || "Station";

    return `${ship} - ${department} - ${stationName} - ${userName}`;
  };

    const getMyInventoryExportItems = () => {
    const myRows = getMyInventoryRows();
    const countMap = buildInventoryQtyMap(myRows);

    const masterItems = makeInventoryItems.length ? makeInventoryItems : [];

    const masterKeys = new Set(
      masterItems
        .map((item) => getInventoryProductGroupKey(item))
        .filter(Boolean)
    );

    const extraRows = myRows.filter((item) => {
      const key = getInventoryProductGroupKey(item);

      const isExtra =
        cleanText(item.sheetName) === "EXTRA ITEM" ||
        cleanText(item.category).includes("NOT IN MASTER");

      return isExtra || !masterKeys.has(key);
    });

    const sourceItems = masterItems.length
      ? [...masterItems, ...extraRows]
      : myRows;

    return sourceItems.map((item) => {
      const key = getInventoryProductGroupKey(item);
      const count = key && countMap.has(key) ? countMap.get(key) : 0;

      return {
        ...item,
        code: item.code || "",
        name: item.name || "",
        image: item.image || "",
        count,
        Count: count,
        qty: count,
        quantity: count,
        Quantity: count,
        totalQty: count,
      };
    });
  };

  const getSummaryInventoryRecordsForDownload = () => {
  const ship = makeInventoryShip || userShip;
  const selectedStation = summaryStationFilter || "ALL";

  const summaryCountRecords = inventorySummary
    .filter(
      (item) =>
        item.ship === ship &&
        inventoryRecordMatchesCurrentDepartment(item) &&
        (selectedStation === "ALL" || item.station === selectedStation)
    )
    .map((item) => {
      const qty = Number(item.qty || 0);
      const safeQty = Number.isFinite(qty) ? qty : 0;

      return {
        ...item,
        count: safeQty,
        Count: safeQty,
        qty: safeQty,
        quantity: safeQty,
        Quantity: safeQty,
        totalQty: safeQty,
      };
    });

  const countMap = buildInventoryQtyMap(summaryCountRecords);

  const masterItems = makeInventoryItems.length ? makeInventoryItems : [];
  const masterKeys = new Set(
    masterItems.map((item) => getInventoryProductGroupKey(item)).filter(Boolean)
  );

  const summaryRows = getShipSummaryRows();

  const extraRows = summaryRows.filter((item) => {
    const key = getInventoryProductGroupKey(item);
    const isExtra =
      cleanText(item.sheetName) === "EXTRA ITEM" ||
      cleanText(item.category).includes("NOT IN MASTER");

    return isExtra || !masterKeys.has(key);
  });

  const sourceItems = masterItems.length
    ? [...masterItems, ...extraRows]
    : summaryRows;

  return sourceItems.map((item) => {
    const key = getInventoryProductGroupKey(item);
    const count = key && countMap.has(key) ? countMap.get(key) : 0;

    return {
      ...item,
      code: item.code || "",
      name: item.name || "",
      image: item.image || "",
      count,
      Count: count,
      qty: count,
      quantity: count,
      Quantity: count,
      totalQty: count,
    };
  });
};

  const handleInventoryCountSheetTemplateFile = (event) => {
    const file = event.target.files?.[0];

    if (!file) {
      setInventoryCountSheetTemplateFile(null);
      setInventoryCountSheetTemplateName("");
      return;
    }

    setInventoryCountSheetTemplateFile(file);
    setInventoryCountSheetTemplateName(file.name);

    setMakeInventoryMessage(
      `Inventory sheet sample uploaded: ${file.name}. Export Excel will keep this file's item positions and write counts into column S.`
    );
  };
  const getMyInventoryStatusRows = () => {
    const countedMap = {};

    getMyInventoryRows().forEach((item) => {
      const key = item.itemKey || cleanText(item.code || item.name);
      countedMap[key] = item;
    });

    const selectedRestaurantStation = cleanText(inventoryStation);

const sourceMasterItems =
  equipmentDepartment === "restaurant" && selectedRestaurantStation
    ? makeInventoryItems.filter(
        (item) => cleanText(item.sheetName) === selectedRestaurantStation
      )
    : makeInventoryItems;

return sourceMasterItems
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
    if (saveBusyRef.current) return;

    if (!currentInventoryItem) {
      setMakeInventoryMessage("Select a product before confirming quantity.");
      return;
    }

    const ship = makeInventoryShip || userShip;
    const station = inventoryStation;
    const userName = getEffectiveInventoryUserName();

    if (!supabase) {
      const text = "Supabase is not connected. Check NEXT_PUBLIC_SUPABASE_URL and NEXT_PUBLIC_SUPABASE_ANON_KEY in Vercel.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    if (!ship || !station || !userName) {
      const text = "Choose ship, station, and user name before confirming inventory.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    if (String(inventoryQty).trim() === "") {
      const text = "Enter quantity counted before confirming.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const qty = Number(inventoryQty || 0);

    if (Number.isNaN(qty) || qty < 0) {
      const text = "Quantity must be a valid number.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const currentStationProgress = getCurrentStationProgress();

    if (currentStationProgress?.status === "submitted") {
      const text =
        "This station has already submitted its inventory count. Reset inventory before starting a new count.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    saveBusyRef.current = true;

    const itemSnapshot = { ...currentInventoryItem };
    const qtySnapshot = String(qty);
    const editingIdSnapshot = editingInventoryId;
    const previousInventorySummary = inventorySummary;

    const itemKey = getInventoryItemKey(itemSnapshot);
    const now = new Date().toISOString();

    const optimisticRecord = {
      id: editingIdSnapshot || `optimistic-${ship}-${station}-${userName}-${itemKey}`,
      ship,
      station,
      userName,
      itemKey,
      code: itemSnapshot.code || "",
      name: itemSnapshot.name || "",
      category: itemSnapshot.category || "",
      sheetName: itemSnapshot.sheetName || "",
      image: itemSnapshot.image || "",
      qty,
      confirmedAt: new Date(now).toLocaleString(),
      updatedAt: now,
      optimistic: true,
    };

    setInventorySummary((prev) => {
      const withoutCurrent = prev.filter(
        (item) =>
          item.id !== optimisticRecord.id &&
          !(
            item.ship === ship &&
            item.station === station &&
            item.userName === userName &&
            item.itemKey === itemKey
          )
      );

      return [optimisticRecord, ...withoutCurrent];
    });

    setCurrentInventoryItem(null);
    setInventoryQty("");
    setEditingInventoryId(null);
    setMakeInventoryMessage(`Saved locally: ${itemSnapshot.name} / Qty ${formatQty(qty)}. Syncing...`);

    try {
      const payload = {
        ship,
        station,
        user_name: userName,
        item_key: itemKey,
        code: itemSnapshot.code || "",
        item_name: itemSnapshot.name || "",
        category: itemSnapshot.category || "",
        sheet_name: itemSnapshot.sheetName || "",
        image: itemSnapshot.image || "",
        qty,
        updated_at: now,
      };

      const { data: existingRows, error: findError } = await supabase
        .from("inventory_counts")
        .select("id")
        .eq("ship", ship)
        .eq("station", station)
        .eq("user_name", userName)
        .eq("item_key", itemKey)
        .limit(1);

      if (findError) {
        throw findError;
      }

      const existingId = existingRows?.[0]?.id;
      let saveResult;

      if (existingId) {
        saveResult = await supabase
          .from("inventory_counts")
          .update(payload)
          .eq("id", existingId)
          .select("*")
          .single();
      } else {
        saveResult = await supabase
          .from("inventory_counts")
          .insert(payload)
          .select("*")
          .single();
      }

      if (saveResult.error) {
        throw saveResult.error;
      }

      const savedRecord = normalizeInventoryRecord(saveResult.data);

      setInventorySummary((prev) => {
        const withoutSavedRecord = prev.filter(
          (item) =>
            item.id !== optimisticRecord.id &&
            item.id !== savedRecord.id &&
            !(
              item.ship === savedRecord.ship &&
              item.station === savedRecord.station &&
              item.userName === savedRecord.userName &&
              item.itemKey === savedRecord.itemKey
            )
        );

        return [savedRecord, ...withoutSavedRecord];
      });

      try {
        await upsertInventoryStationStatus("started");
      } catch (statusError) {
        setMakeInventoryMessage(
          `Saved ${itemSnapshot.name} / Qty ${formatQty(qty)}, but station status could not sync: ${statusError.message}`
        );
        return;
      }

      setMakeInventoryMessage(`Saved ${itemSnapshot.name} / Qty ${formatQty(qty)}.`);

      logUsageEvent("inventory_count_saved", {
        module: "make_inventory",
        ship,
        station,
        userName,
        userPosition: inventoryUserPosition,
        itemName: itemSnapshot.name,
        code: itemSnapshot.code,
        category: itemSnapshot.category,
        qty,
      });
    } catch (error) {
      setInventorySummary(previousInventorySummary);
      setCurrentInventoryItem(itemSnapshot);
      setInventoryQty(qtySnapshot);
      setEditingInventoryId(editingIdSnapshot);

      const text = `Sync failed. Quantity was not saved: ${error.message}`;
      setMakeInventoryMessage(text);
      window.alert(text);
    } finally {
      setInventoryLoading(false);
      saveBusyRef.current = false;
    }
  };

  const saveExtraInventoryItem = async () => {
  if (extraInventorySaving) return;

  const ship = makeInventoryShip || userShip;
  const station = inventoryStation;
  const userName = getEffectiveInventoryUserName();

  const code = String(extraInventoryCode || "").trim();
  const name = String(extraInventoryName || "").replace(/\s+/g, " ").trim();
  const qty = Number(extraInventoryQty || 0);

  if (!supabase) {
    const text =
      "Supabase is not connected. Cannot save extra inventory item.";
    setExtraInventoryMessage(text);
    window.alert(text);
    return;
  }

  if (!ship || !station || !userName) {
    const text =
      "Choose ship, station, and user before saving an extra item.";
    setExtraInventoryMessage(text);
    window.alert(text);
    return;
  }

  if (!name) {
    const text = "Enter the item name before saving.";
    setExtraInventoryMessage(text);
    window.alert(text);
    return;
  }

  if (String(extraInventoryQty).trim() === "" || Number.isNaN(qty) || qty < 0) {
    const text = "Enter a valid quantity before saving.";
    setExtraInventoryMessage(text);
    window.alert(text);
    return;
  }

  const currentStationProgress = getCurrentStationProgress();

  if (currentStationProgress?.status === "submitted") {
    const text =
      "This station has already submitted its inventory count. Reset inventory before starting a new count.";
    setExtraInventoryMessage(text);
    window.alert(text);
    return;
  }

  setExtraInventorySaving(true);
  setExtraInventoryMessage("Saving extra item...");

  try {
    const departmentKey = getCurrentEquipmentDepartmentKey();
    const itemKey = cleanText(
      departmentKey +
        "|EXTRA|" +
        (code || "NO-CODE") +
        "|" +
        name
    );

    const imageUrl = extraInventoryPhotoFile
      ? await uploadExtraInventoryPhotoToStorage({
          file: extraInventoryPhotoFile,
          ship,
          station,
          code,
          name,
        })
      : "";

    const now = new Date().toISOString();

    const payload = {
      ship,
      station,
      user_name: userName,
      item_key: itemKey,
      code: code || "EXTRA",
      item_name: name,
      category: "Extra item - not in master list",
      sheet_name: "Extra Item",
      image: imageUrl,
      qty,
      updated_at: now,
    };

    const { data: existingRows, error: findError } = await supabase
      .from("inventory_counts")
      .select("id")
      .eq("ship", ship)
      .eq("station", station)
      .eq("user_name", userName)
      .eq("item_key", itemKey)
      .limit(1);

    if (findError) {
      throw findError;
    }

    const existingId = existingRows?.[0]?.id;
    let saveResult;

    if (existingId) {
      saveResult = await supabase
        .from("inventory_counts")
        .update(payload)
        .eq("id", existingId)
        .select("*")
        .single();
    } else {
      saveResult = await supabase
        .from("inventory_counts")
        .insert(payload)
        .select("*")
        .single();
    }

    if (saveResult.error) {
      throw saveResult.error;
    }

    const savedRecord = normalizeInventoryRecord(saveResult.data);

    setInventorySummary((prev) => {
      const withoutSaved = prev.filter(
        (item) =>
          item.id !== savedRecord.id &&
          !(
            item.ship === savedRecord.ship &&
            item.station === savedRecord.station &&
            item.userName === savedRecord.userName &&
            item.itemKey === savedRecord.itemKey
          )
      );

      return [savedRecord, ...withoutSaved];
    });

    try {
      await upsertInventoryStationStatus("started");
    } catch {}

    setExtraInventoryCode("");
    setExtraInventoryName("");
    setExtraInventoryQty("");
    setExtraInventoryPhotoFile(null);
    setExtraInventoryPhotoInputKey((value) => value + 1);

    setExtraInventoryMessage(
      "Extra item saved: " + name + " / Qty " + formatQty(qty)
    );

    logUsageEvent("extra_inventory_item_saved", {
      module: "make_inventory",
      ship,
      station,
      userName,
      userPosition: inventoryUserPosition,
      code,
      itemName: name,
      qty,
      hasPhoto: Boolean(imageUrl),
    });
  } catch (error) {
    const text = error?.message || "Could not save extra item.";
    setExtraInventoryMessage(text);
    window.alert(text);
  } finally {
    setExtraInventorySaving(false);
  }
};
 const editInventoryItem = (item) => {
  const masterMatch =
    makeInventoryItems.find(
      (masterItem) =>
        getInventoryProductGroupKey(masterItem) ===
        getInventoryProductGroupKey(item)
    ) ||
    musterItems.find(
      (masterItem) =>
        getInventoryProductGroupKey(masterItem) ===
        getInventoryProductGroupKey(item)
    ) ||
    item;

  setCurrentInventoryItem({
    code: item.code,
    name: item.name,
    category: item.category,
    sheetName: item.sheetName,
    image: getEquipmentDisplayImage(masterMatch) || item.image || "",
    imageFallback: getEquipmentFallbackImage(masterMatch),
  });

    setInventoryQty(String(item.qty ?? ""));
    setEditingInventoryId(item.id || null);

    if (item.station) setInventoryStation(item.station);

    if (item.userName) {
      const userParts = String(item.userName).split(" - ");
      setInventoryUserName(userParts[0]?.trim() || "");
      setInventoryUserPosition(userParts.slice(1).join(" - ").trim());
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

    setInventorySummary((prev) => prev.filter((item) => item.id !== itemToDelete.id));
    logUsageEvent("inventory_count_deleted", {
      module: "make_inventory",
      ship: itemToDelete.ship || makeInventoryShip || userShip,
      station: itemToDelete.station || inventoryStation,
      userName: itemToDelete.userName || getEffectiveInventoryUserName(),
      userPosition: inventoryUserPosition,
      code: itemToDelete.code || "",
      itemName: itemToDelete.name || "",
      qty: itemToDelete.qty || 0,
    });
  };

  const clearMyInventory = async () => {
    const ship = makeInventoryShip || userShip;
    const userName = getEffectiveInventoryUserName();

    if (!supabase) {
      const text = "Supabase is not connected. Cannot clear records.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    if (!ship || !inventoryStation || !userName) {
      const text = "Choose ship, station, and user before clearing your report.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const rowsToClear = getMyInventoryRows();
    if (!rowsToClear.length) {
      const text = "There are no records to clear for this ship, station, and user.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const confirmed = window.confirm(
      `Clear ${rowsToClear.length} inventory record(s) for ${ship} / ${inventoryStation} / ${userName}?`
    );

    if (!confirmed) return;

    setInventoryLoading(true);
    setMakeInventoryMessage("Clearing your inventory report...");

    const { error } = await supabase
      .from("inventory_counts")
      .delete()
      .eq("ship", ship)
      .eq("station", inventoryStation)
      .eq("user_name", userName);

    if (error) {
      const text = `Could not clear inventory: ${error.message}`;
      setInventoryError(text);
      setMakeInventoryMessage(text);
      setInventoryLoading(false);
      window.alert(text);
      return;
    }

    setInventorySummary((prev) =>
      prev.filter(
        (item) =>
          !(item.ship === ship && item.station === inventoryStation && item.userName === userName)
      )
    );

    logUsageEvent("inventory_report_cleared", {
      module: "make_inventory",
      ship,
      station: inventoryStation,
      userName,
      userPosition: inventoryUserPosition,
      recordsCleared: rowsToClear.length,
    });

    scheduleRealtimeRefresh("counts", ship);

    setCurrentInventoryItem(null);
    setInventoryQty("");
    setEditingInventoryId(null);
    setInventoryLoading(false);
    setMakeInventoryMessage("My inventory report was cleared.");
  };

  const clearShipInventory = async () => {
    const ship = makeInventoryShip || userShip;

    if (!supabase) {
      const text = "Supabase is not connected. Cannot clear ship records.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    if (!ship) {
      const text = "Choose ship before clearing ship records.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const shipRows = inventorySummary.filter((item) => item.ship === ship);
    if (!shipRows.length) {
      const text = `There are no inventory records to clear for ${ship}.`;
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const confirmed = window.confirm(
      `Clear ALL ${shipRows.length} inventory record(s) for ${ship} from ALL users and ALL stations?`
    );

    if (!confirmed) return;

    const secondConfirm = window.confirm(
      `This cannot be undone. Confirm again to clear all ${ship} records.`
    );

    if (!secondConfirm) return;

    setInventoryLoading(true);
    setMakeInventoryMessage(`Clearing all ${ship} inventory records...`);

    const { error } = await supabase
      .from("inventory_counts")
      .delete()
      .eq("ship", ship);

    if (error) {
      const text = `Could not clear ship inventory: ${error.message}`;
      setInventoryError(text);
      setMakeInventoryMessage(text);
      setInventoryLoading(false);
      window.alert(text);
      return;
    }

    setInventorySummary((prev) => prev.filter((item) => item.ship !== ship));
    logUsageEvent("ship_inventory_cleared", {
      module: "make_inventory",
      ship,
      recordsCleared: shipRows.length,
    });
    scheduleRealtimeRefresh("counts", ship);

    setCurrentInventoryItem(null);
    setInventoryQty("");
    setEditingInventoryId(null);
    setInventoryLoading(false);
    setMakeInventoryMessage(`All ${ship} inventory records were cleared.`);
  };

    const resetInventoryRun = async () => {
    if (reportBusy) return;

    if (!supabase) {
      const text = "Supabase is not connected. Cannot reset inventory.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const ship = makeInventoryShip || userShip;
    const departmentKey = getCurrentEquipmentDepartmentKey();
    const departmentLabel = activeEquipmentDepartmentLabel || equipmentDepartment || "Equipment";

    if (!ship) {
      const text = "Choose ship before resetting inventory.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const rowsToDelete = inventorySummary.filter(
      (item) =>
        item.ship === ship &&
        inventoryRecordMatchesCurrentDepartment(item)
    );

    const firstConfirm = window.confirm(
      `Reset inventory for ${ship} / ${departmentLabel}?\n\nThis will delete all counted quantities and all station started/submitted statuses for this inventory run.\n\nThis cannot be undone.`
    );

    if (!firstConfirm) return;

    const copyConfirm = window.confirm(
      "Before reset: did you make/download a copy of the final report?\n\nClick OK only if you already made a copy.\nClick Cancel to stop reset."
    );

    if (!copyConfirm) {
      setMakeInventoryMessage("Reset cancelled. Please make/download a copy before resetting inventory.");
      return;
    }

    setReportBusy(true);
    setInventoryLoading(true);
    setMakeInventoryMessage("Resetting inventory...");

    try {
      const idsToDelete = rowsToDelete
        .map((item) => item.id)
        .filter(Boolean);

      const batchSize = 500;

      for (let index = 0; index < idsToDelete.length; index += batchSize) {
        const batchIds = idsToDelete.slice(index, index + batchSize);

        if (!batchIds.length) continue;

        const { error } = await supabase
          .from("inventory_counts")
          .delete()
          .in("id", batchIds);

        if (error) {
          throw error;
        }
      }

      const { error: statusError } = await supabase
        .from("inventory_station_status")
        .delete()
        .eq("ship", ship)
        .eq("department", departmentKey);

      if (statusError) {
        throw statusError;
      }

      setInventorySummary((prev) =>
        prev.filter(
          (item) =>
            !(
              item.ship === ship &&
              inventoryRecordMatchesCurrentDepartment(item)
            )
        )
      );

      setInventoryStationStatuses((prev) =>
        prev.filter(
          (item) =>
            !(
              item.ship === ship &&
              item.department === departmentKey
            )
        )
      );

      setCurrentInventoryItem(null);
      setInventoryQty("");
      setEditingInventoryId(null);
      setInventoryReportMode("my");
      setSummaryStationFilter("ALL");
      setShowVariance(false);

      scheduleRealtimeRefresh("counts", ship);
      scheduleRealtimeRefresh("stationStatus", ship);

      setMakeInventoryMessage(
        `${ship} / ${departmentLabel} inventory was reset. You can start a new inventory now.`
      );

      logUsageEvent("inventory_run_reset", {
        module: "make_inventory",
        ship,
        equipmentDepartment,
        departmentKey,
        recordsDeleted: idsToDelete.length,
      });
    } catch (error) {
      const text = error?.message || "Could not reset inventory.";
      setMakeInventoryMessage(text);
      window.alert(text);
    } finally {
      setInventoryLoading(false);
      setReportBusy(false);
    }
  };
  const openPreparedPrintWindow = (html) => {
    if (printBusyRef.current) return;

    const printWindow = window.open("", "_blank");
    if (!printWindow) {
      const text = "Print window was blocked by the browser. Allow popups for this app and try again.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    printBusyRef.current = true;
    setReportBusy(true);
    setMakeInventoryMessage("Preparing print view...");

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();

    window.setTimeout(() => {
      try {
        printWindow.focus();
        printWindow.print();
      } finally {
        window.setTimeout(() => {
          printBusyRef.current = false;
          setReportBusy(false);
          setMakeInventoryMessage("Print view prepared.");
        }, 700);
      }
    }, 250);
  };

        const exportInventorySummaryToExcel = async () => {
    if (reportBusy) return;

    logUsageEvent("export_excel_clicked", {
      module: "make_inventory",
      reportMode: inventoryReportMode,
      ship: makeInventoryShip || userShip,
      station: inventoryReportMode === "summary" ? summaryStationFilter : inventoryStation,
      templateFile: inventoryCountSheetTemplateName || "",
    });

    const ship = makeInventoryShip || userShip;

    if (!ship) {
      const text = "Choose ship before exporting.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    if (!inventoryCountSheetTemplateFile) {
      const text =
        "Upload the inventory sheet sample first. Export Excel will use that uploaded file and write counts into column S.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    setReportBusy(true);

    try {
      if (inventoryReportMode === "summary") {
        const rows = getSummaryInventoryRecordsForDownload();

        const result = await downloadInventoryExcelReportUsingTemplate({
          templateFile: inventoryCountSheetTemplateFile,
          items: rows,
          venueName: getInventoryReportLocationName("summary"),
          reportTitle: "Summary Report",
        });

        setMakeInventoryMessage(
          `Summary Excel report downloaded from uploaded sample. Positions stayed the same. ${result.matchedRows} of ${result.itemRows} item rows matched counts; unmatched rows were set to 0.`
        );

        return;
      }

      const rows = getMyInventoryExportItems();

      const result = await downloadInventoryExcelReportUsingTemplate({
        templateFile: inventoryCountSheetTemplateFile,
        items: rows,
        venueName: getInventoryReportLocationName("my"),
        reportTitle: "Inventory Report",
      });

      setMakeInventoryMessage(
        `Inventory Excel report downloaded from uploaded sample. Positions stayed the same. ${result.matchedRows} of ${result.itemRows} item rows matched counts; unmatched rows were set to 0.`
      );
    } catch (error) {
      const text = error?.message || "Could not export Excel report.";
      setMakeInventoryMessage(text);
      window.alert(text);
    } finally {
      setReportBusy(false);
    }
  };

  const exportInventoryStatusToExcel = async () => {
    logUsageEvent("export_excel_clicked", { module: "make_inventory", reportMode: "count_status", ship: makeInventoryShip || userShip, station: inventoryStation });
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

const XLSX = await loadXlsx();

const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Inventory Status");
    XLSX.writeFile(wb, `inventory-status-${ship || "ship"}.xlsx`);
  };

    const printInventorySummary = async () => {
    if (reportBusy || printBusyRef.current) return;

    logUsageEvent("print_clicked", {
      module: "make_inventory",
      reportMode: inventoryReportMode,
      ship: makeInventoryShip || userShip,
      station: inventoryReportMode === "summary" ? summaryStationFilter : inventoryStation,
    });

    const ship = makeInventoryShip || userShip;

    if (!ship) {
      const text = "Choose ship before creating the PDF report.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    setReportBusy(true);
    printBusyRef.current = true;

    try {
      if (inventoryReportMode === "summary") {
        const rows = getSummaryInventoryRecordsForDownload();

        if (!rows.length) {
          const text = "No summary records to export to PDF for this ship.";
          setMakeInventoryMessage(text);
          window.alert(text);
          return;
        }

        await downloadInventoryPdfReport({
  items: rows,
  venueName: getInventoryReportLocationName("summary"),
  reportTitle: "Summary Report",
  includeCounts: true,
  summary: false,
});

        setMakeInventoryMessage("Summary PDF report downloaded.");
        return;
      }

      const rows = getMyInventoryExportItems();

      if (!rows.length) {
        const text =
          "No inventory items to export to PDF. Upload or refresh the shared master inventory list first.";
        setMakeInventoryMessage(text);
        window.alert(text);
        return;
      }

      await downloadInventoryPdfReport({
        items: rows,
        venueName: getInventoryReportLocationName("my"),
        reportTitle: "Inventory Report",
        includeCounts: true,
        summary: false,
      });

      setMakeInventoryMessage(
        "Inventory PDF report downloaded. Code is in column A, item name is in column F, and count is in column S."
      );
    } catch (error) {
      const text = error?.message || "Could not create PDF report.";
      setMakeInventoryMessage(text);
      window.alert(text);
    } finally {
      printBusyRef.current = false;
      setReportBusy(false);
    }
  };

  const printInventoryStatus = () => {
    logUsageEvent("print_clicked", { module: "make_inventory", reportMode: "count_status", ship: makeInventoryShip || userShip, station: inventoryStation });
    if (reportBusy || printBusyRef.current) return;

    const ship = makeInventoryShip || userShip;
    const rows = getMyInventoryStatusRows().map((item) => ({ ...item }));

    if (!rows.length) {
      const text = "No inventory status rows to print. Upload a master inventory file first.";
      setMakeInventoryMessage(text);
      window.alert(text);
      return;
    }

    const html = `
      <html>
        <head>
          <title>Inventory Status</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .counted { color: #2e7d32; font-weight: bold; }
            .pending { color: #555; font-weight: bold; }
            tr { break-inside: avoid; }
          </style>
        </head>
        <body>
          <h1>Inventory Status</h1>
          <div><strong>Ship:</strong> ${escapeHtml(ship)}</div>
          <div><strong>Station:</strong> ${escapeHtml(inventoryStation)}</div>
          <div><strong>User:</strong> ${escapeHtml(getEffectiveInventoryUserName())}</div>
          <div><strong>Printed:</strong> ${escapeHtml(new Date().toLocaleString())}</div>
          <div><strong>Records:</strong> ${escapeHtml(rows.length)}</div>

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
                      <td>${escapeHtml(item.code || "")}</td>
                      <td>${escapeHtml(item.name || "")}</td>
                      <td>${escapeHtml(item.category || "")}</td>
                      <td>${escapeHtml(item.sheetName || "")}</td>
                      <td class="${item.status === "Counted" ? "counted" : "pending"}">${escapeHtml(item.status)}</td>
                      <td>${escapeHtml(formatQty(item.countedQty))}</td>
                      <td>${escapeHtml(item.countedAt)}</td>
                    </tr>
                  `
                )
                .join("")}
            </tbody>
          </table>
        </body>
      </html>
    `;

    openPreparedPrintWindow(html);
  };

  const uploadConsumptionFile = (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  setMessage("Loading consumption file...");

  readExcelFile(
    file,
    (workbook) => {
      try {
        const rows = workbookToRows(workbook);
        const safeRows = Array.isArray(rows) ? rows : [];
        const productList = buildProductList(safeRows);

        setConsumptionRows(safeRows);
        setProducts(productList);
        setSelectedProduct("");
        setSelectedRecipe(null);
        setProductCostReportSearch("");

        setMessage(
          "Consumption file loaded. " +
            Math.max(safeRows.length - 1, 0) +
            " row(s), " +
            productList.length +
            " product(s)."
        );

        logUsageEvent("product_consumption_file_uploaded", {
          module: "product_dashboard",
          fileName: file.name,
          rowCount: Math.max(safeRows.length - 1, 0),
          products: productList.length,
        });
      } catch (error) {
        setConsumptionRows([]);
        setProducts([]);
        setSelectedProduct("");
        setSelectedRecipe(null);

        const text =
          error?.message ||
          "Could not process the consumption file. Please check the file format.";

        setMessage(text);
        window.alert(text);
      }
    },
    (error) => {
      setConsumptionRows([]);
      setProducts([]);
      setSelectedProduct("");
      setSelectedRecipe(null);

      const text =
        error?.message ||
        "Could not read the consumption file. Please upload a valid Excel file.";

      setMessage(text);
      window.alert(text);
    }
  );

  e.target.value = "";
};
  const loadPermanentYearlyRegionalConsumptionFile = async () => {
  try {
    setYearlyRegionalMessage("Loading permanent yearly regional consumption file...");

    const response = await fetch("/yearly-regional-consumption.xlsx");

    if (!response.ok) {
      setYearlyRegionalConsumption(null);
      setYearlyRegionalFileName("");
      setYearlyRegionalMessage(
        "Permanent yearly regional consumption file was not found. You can still upload it manually."
      );
      return;
    }

    const [XLSX, arrayBuffer] = await Promise.all([
  loadXlsx(),
  response.arrayBuffer(),
]);

const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
    });

    const parsed = parseYearlyRegionalConsumptionWorkbook(workbook);

    setYearlyRegionalConsumption(parsed);
    setYearlyRegionalFileName("yearly-regional-consumption.xlsx");

    // Keep region blank by default so user must choose ATHENS / MIAMI / BARCELONA.
    setSelectedRegionalConsumptionRegion("");

    if (!parsed.regionOptions?.length) {
      setYearlyRegionalMessage(
        "Permanent yearly file loaded, but no regional ship blocks were detected. Check the header rows."
      );
      return;
    }

    setYearlyRegionalMessage(
      "Permanent yearly regional file loaded. " +
        parsed.aggregates.length +
        " regional product records found across " +
        parsed.regionOptions.length +
        " region(s)."
    );

    logUsageEvent("permanent_yearly_regional_consumption_loaded", {
      module: "product_dashboard",
      fileName: "yearly-regional-consumption.xlsx",
      sourceSheet: parsed.sourceSheet,
      regions: parsed.regionOptions || [],
      aggregates: parsed.aggregates.length,
      shipRegionBlocks: parsed.shipRegionBlocks.length,
    });
  } catch (error) {
    setYearlyRegionalConsumption(null);
    setYearlyRegionalFileName("");
    setYearlyRegionalMessage(
      error?.message || "Could not load permanent yearly regional consumption file."
    );
  }
};
  const uploadYearlyRegionalConsumptionFile = async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    setYearlyRegionalMessage("Loading yearly regional consumption file...");

    const arrayBuffer = await file.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
    });

    const parsed = parseYearlyRegionalConsumptionWorkbook(workbook);

    setYearlyRegionalConsumption(parsed);
    setYearlyRegionalFileName(file.name);
    setSelectedRegionalConsumptionRegion("");

    if (!parsed.regionOptions?.length) {
      setYearlyRegionalMessage(
        "Yearly regional file loaded, but no region / home port blocks were detected. Check the header rows."
      );
      return;
    }

    setYearlyRegionalMessage(
      "Yearly regional file loaded. " +
        parsed.aggregates.length +
        " regional product records found across " +
        parsed.regionOptions.length +
        " region(s)."
    );

    logUsageEvent("yearly_regional_consumption_uploaded", {
      module: "product_dashboard",
      fileName: file.name,
      sourceSheet: parsed.sourceSheet,
      regions: parsed.regionOptions || [],
      aggregates: parsed.aggregates.length,
      shipRegionBlocks: parsed.shipRegionBlocks.length,
    });
  } catch (error) {
    setYearlyRegionalConsumption(null);
    setYearlyRegionalFileName("");
    setSelectedRegionalConsumptionRegion("");

    const text =
      error?.message || "Could not load yearly regional consumption file.";

    setYearlyRegionalMessage(text);
    window.alert(text);
  } finally {
    event.target.value = "";
  }
};
  const loadPermanentIngredientByLocationForProductDashboard = async ({ silent = false } = {}) => {
  if (!supabase) {
    if (!silent) {
      const text = "Supabase is not connected. Permanent Ingredient by Location file cannot load.";
      setMessage(text);
      window.alert(text);
    }

    return false;
  }

  try {
    if (!silent) {
      setMessage("Loading permanent Ingredient by Location file...");
    }

    const arrayBuffer = await downloadIngredientByLocationFileFromStorage({ supabase });
    const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const rows = workbookToRows(workbook);

    setRecipeRows(rows);
    setSelectedRecipe(null);

    if (!silent) {
      setMessage(
        `Permanent Ingredient by Location loaded. ${Math.max(rows.length - 1, 0)} recipe/location row(s).`
      );
    }

    return true;
  } catch (error) {
    if (!silent) {
      const text = error?.message || "Could not load permanent Ingredient by Location file.";
      setMessage(text);
      window.alert(text);
    }

    return false;
  }
};
  useEffect(() => {
  if (!loggedIn) return;

  const shouldLoadProductFiles =
    module === "product" ||
    productMode === "dashboard" ||
    productMode === "nextorder";

  if (!shouldLoadProductFiles) return;

  if (!defaultTemplateLoadedRef.current) {
    defaultTemplateLoadedRef.current = true;
    loadDefaultTemplate();
  }

  if (!yearlyRegionalLoadedRef.current) {
    yearlyRegionalLoadedRef.current = true;
    loadPermanentYearlyRegionalConsumptionFile();
  }

  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [loggedIn, module, productMode]);
    useEffect(() => {
  const shouldLoadIngredientFile =
    module === "product" &&
    (productMode === "dashboard" || productMode === "nextorder");

  if (!shouldLoadIngredientFile) return;
  if ((recipeRows || []).length > 1) return;

  loadPermanentIngredientByLocationForProductDashboard({ silent: true });
  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [module, productMode, recipeRows.length]);
  const uploadRecipeFile = async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  if (!isAdmin) {
    const text = "Only admins can replace the permanent Ingredient by Location file.";
    setMessage(text);
    window.alert(text);
    e.target.value = "";
    return;
  }

  if (!supabase) {
    const text = "Supabase is not connected. Cannot save the permanent Ingredient by Location file.";
    setMessage(text);
    window.alert(text);
    e.target.value = "";
    return;
  }

  try {
    setMessage("Saving permanent Ingredient by Location file...");

    await uploadIngredientByLocationFileToStorage({
      supabase,
      file,
    });

    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
    });

    const rows = workbookToRows(workbook);

    setRecipeRows(rows);
    setSelectedRecipe(null);

    setMessage(
      `Permanent Ingredient by Location file updated. ${Math.max(
        rows.length - 1,
        0
      )} recipe/location row(s) loaded.`
    );

    logUsageEvent("permanent_ingredient_by_location_updated", {
      module: "product_dashboard",
      fileName: file.name,
      rowCount: Math.max(rows.length - 1, 0),
      permanent: true,
    });
  } catch (error) {
    const text =
      error?.message || "Could not save permanent Ingredient by Location file.";
    setMessage(text);
    window.alert(text);
  } finally {
    e.target.value = "";
  }
};

  const uploadTemplateFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      const parsedTemplate = parseTemplateWorkbook(workbook);
      setTemplateMap(parsedTemplate);
      setTemplateStatus("Custom ERP template loaded.");
      setNextOrderTemplateFileName(file.name || "Custom ERP template loaded");
      logUsageEvent("product_template_file_uploaded", {
        module: "product_dashboard",
        fileName: file.name,
        sheetCount: workbook.SheetNames.length,
        venueCount: Object.keys(parsedTemplate).length,
      });
    });
  };

  const uploadMusterFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    (async () => {
      try {
        const { workbook, items, sourceSheetName } = await parseEquipmentMasterFile({
  file,
  equipmentDepartment,
});

        setMusterItems(items);
        setMakeInventoryItems(items);
        setSelectedEquipment(null);

        setMusterMessage(
  activeEquipmentDepartmentLabel + " Equipment Master List loaded from " + sourceSheetName + ". Saving shared list..."
);

        logUsageEvent("equipment_muster_file_uploaded", {
          module: "equipment_" + equipmentDepartment + "_muster",
          equipmentDepartment,
          fileName: file.name,
          sheetCount: workbook.SheetNames.length,
          sourceSheetName,
          itemCount: items.length,
        });

        await saveMasterInventoryItems(null, items);
        e.target.value = "";
      } catch (error) {
        setMusterMessage(
          "Could not load " + activeEquipmentDepartmentLabel + " file: " + (error?.message || "Unknown error")
        );
      }
    })();
  };

  const uploadWarehouseFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      const rows = workbookToRows(workbook);
      setWarehouseRows(rows);
      setWarehouseFilter("all");
      setWarehouseMessage("Warehouse inventory loaded.");
      logUsageEvent("warehouse_file_uploaded", {
        module: "inventory_warehouse",
        fileName: file.name,
        rowCount: Math.max(rows.length - 1, 0),
      });
    });
  };

  const uploadInUseFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    readExcelFile(file, (workbook) => {
      const rows = workbookToRows(workbook);
      setInUseRows(rows);
      setInUseMessage("Inventory in Use file loaded.");
      logUsageEvent("inventory_in_use_file_uploaded", {
        module: "inventory_in_use",
        fileName: file.name,
        rowCount: Math.max(rows.length - 1, 0),
      });
    });
  };

const uploadEquipmentPictureZipFile = async (event) => {
  const file = event.target.files?.[0];

  if (!file) return;

  if (!isAdmin) {
    const text = "Only admin can upload the equipment picture ZIP.";
    setPictureLibraryMessage(text);
    window.alert(text);
    event.target.value = "";
    return;
  }

  if (equipmentDepartment !== "culinary") {
    const text = "This picture ZIP upload is for Culinary equipment only.";
    setPictureLibraryMessage(text);
    window.alert(text);
    event.target.value = "";
    return;
  }

  if (!supabase) {
    const text = "Supabase is not connected. Cannot upload picture ZIP.";
    setPictureLibraryMessage(text);
    window.alert(text);
    event.target.value = "";
    return;
  }

  const sourceItems = makeInventoryItems.length ? makeInventoryItems : musterItems;

  if (!sourceItems.length) {
    const text = "Upload or refresh the Culinary master inventory list first, then upload the picture ZIP.";
    setPictureLibraryMessage(text);
    window.alert(text);
    event.target.value = "";
    return;
  }

  setPictureLibraryBusy(true);
setPictureLibraryMessage("Reading picture ZIP...");

try {
  const [{ default: JSZip }, arrayBuffer] = await Promise.all([
    import("jszip"),
    file.arrayBuffer(),
  ]);

  const zip = await JSZip.loadAsync(arrayBuffer);
  const entries = [];

    zip.forEach((relativePath, zipEntry) => {
      if (zipEntry.dir) return;

      const fileName = relativePath.split("/").pop() || "";
      const isImage = /\.(jpg|jpeg|png|webp|gif)$/i.test(fileName);

      if (!isImage) return;

      const codeKey = normalizeEquipmentPictureCode(fileName);

      if (!codeKey) return;

      entries.push({
        relativePath,
        fileName,
        zipEntry,
        codeKey,
      });
    });

    if (!entries.length) {
      throw new Error("No usable image files found in the ZIP. Filenames must contain the equipment code.");
    }

    const masterScope = getMasterInventoryScope("culinary");
    const pictureByCode = {};
    let uploadedCount = 0;
    let skippedDuplicateCount = 0;

    for (let index = 0; index < entries.length; index += 1) {
      const entry = entries[index];

      if (pictureByCode[entry.codeKey]) {
        skippedDuplicateCount += 1;
        continue;
      }

      if (index > 0 && index % 25 === 0) {
        setPictureLibraryMessage(
          `Uploading ZIP pictures... ${index} of ${entries.length}`
        );
      }

      const blob = await entry.zipEntry.async("blob");
      const mimeType = getZipImageMimeType(entry.fileName);
      const extension = getZipImageExtension(entry.fileName);
      const storagePath = `${masterScope}/zip-library/${entry.codeKey}.${extension}`;

      const { error } = await supabase.storage
        .from(EQUIPMENT_PICTURE_BUCKET)
        .upload(storagePath, blob, {
          contentType: mimeType,
          upsert: true,
          cacheControl: "31536000",
        });

      if (error) {
        throw error;
      }

      const { data } = supabase.storage
        .from(EQUIPMENT_PICTURE_BUCKET)
        .getPublicUrl(storagePath);

      const publicUrl = data?.publicUrl || "";

      if (publicUrl) {
        pictureByCode[entry.codeKey] = publicUrl;
        uploadedCount += 1;
      }
    }

    let matchedCount = 0;
    const unmatchedCodes = [];

    const updatedItems = sourceItems.map((item) => {
  const codeKeys =
    typeof getEquipmentPictureCandidateCodes === "function"
      ? getEquipmentPictureCandidateCodes(item)
      : [
          normalizeEquipmentPictureCode(item?.code),
          getEquipmentImageMatchCode(item?.code),
        ].filter(Boolean);

  const matchedCodeKey = codeKeys.find((codeKey) => pictureByCode[codeKey]);
  const pictureUrl = matchedCodeKey ? pictureByCode[matchedCodeKey] : "";

  if (!pictureUrl) {
    if (codeKeys[0]) unmatchedCodes.push(codeKeys[0]);
    return item;
  }

  matchedCount += 1;

  return {
    ...item,
    image: pictureUrl,
    imageFallback: item.imageFallback || item.image || "",
    pictureFileName: matchedCodeKey,
    pictureMatchCode: matchedCodeKey,
  };
});
    setDrivePictureLibraryByCode((current) => ({
      ...current,
      ...pictureByCode,
    }));

    setMakeInventoryItems(updatedItems);
    setMusterItems(updatedItems);

    setPictureLibraryMessage(
      `ZIP pictures uploaded. ${uploadedCount} picture code(s) saved. ${matchedCount} master item(s) matched. Saving master list...`
    );

    await saveMasterInventoryItems(null, updatedItems);

    setPictureLibraryMessage(
      `Picture ZIP sync completed. ${uploadedCount} picture code(s) saved from ZIP. ${matchedCount} item(s) matched. ${unmatchedCodes.length} master item code(s) had no ZIP picture. ${skippedDuplicateCount} duplicate picture file(s) skipped.`
    );

    logUsageEvent("equipment_picture_zip_uploaded", {
      module: "make_inventory",
      ship: makeInventoryShip || userShip,
      equipmentDepartment,
      fileName: file.name,
      uploadedCount,
      matchedCount,
      unmatchedCount: unmatchedCodes.length,
      duplicateCount: skippedDuplicateCount,
    });
  } catch (error) {
    const text = error?.message || "Could not upload equipment picture ZIP.";
    setPictureLibraryMessage(text);
    window.alert(text);
  } finally {
    setPictureLibraryBusy(false);
    event.target.value = "";
  }
};
    const loadDrivePictureLibrary = async ({ silent = false } = {}) => {
  try {
    if (!silent) {
      setPictureLibraryMessage("Loading picture folder...");
    }

    const response = await fetch("/api/drive-picture-library", {
      cache: "no-store",
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(data?.error || "Could not load picture folder.");
    }

    const byCode = {};

    const addPictureKey = (key, pictureUrl) => {
      const rawKey = String(key || "").trim();
      const url = String(pictureUrl || "").trim();

      if (!rawKey || !url) return;

      const cleanKey = rawKey.replace(/^0+(?=\d)/g, "");

      [rawKey, cleanKey].forEach((candidateKey) => {
        if (!candidateKey) return;
        if (byCode[candidateKey]) return;

        byCode[candidateKey] = url;
      });
    };

    (data.files || []).forEach((file) => {
      const pictureUrl =
        file.webViewLink ||
        file.imageUrl ||
        file.thumbnailUrl ||
        "";

      if (!pictureUrl) return;

      (file.numbers || []).forEach((number) => {
        addPictureKey(number, pictureUrl);
      });

      const fileName = String(file.name || "").trim();

      addPictureKey(normalizeEquipmentPictureCode(fileName), pictureUrl);
      addPictureKey(getEquipmentImageMatchCode(fileName), pictureUrl);

      const numberMatches = fileName.match(/\d{4,}/g) || [];

      numberMatches.forEach((numberText) => {
        addPictureKey(numberText, pictureUrl);
      });
    });

    setDrivePictureLibraryByCode(byCode);

    if (!silent) {
      setPictureLibraryMessage(
        "Picture folder loaded. " +
          (data.count || 0) +
          " Drive image(s) found. " +
          Object.keys(byCode).length +
          " picture code key(s) prepared."
      );
    }

    return byCode;
  } catch (error) {
    if (!silent) {
      const text = error?.message || "Could not load picture folder.";
      setPictureLibraryMessage(text);
      window.alert(text);
    }

    return {};
  }
};

const isTemporaryGoogleThumbnail = (value) => {
  const text = String(value || "").trim().toLowerCase();

  return (
    text.includes("googleusercontent.com") ||
    text.includes("thumbnail") ||
    text.includes("drive-thirdparty")
  );
};

const getEquipmentPictureCandidateCodes = (item) => {
  const candidates = [];

  const addCandidate = (value) => {
    const rawValue = String(value || "").trim();
    if (!rawValue) return;

    const numericKey = normalizeEquipmentPictureCode(rawValue);
    const looseKey = getEquipmentImageMatchCode(rawValue);

    [numericKey, looseKey].forEach((key) => {
      const cleanKey = String(key || "").trim();

      if (!cleanKey) return;
      if (candidates.includes(cleanKey)) return;

      candidates.push(cleanKey);
    });

    const numberMatches = rawValue.match(/\d{4,}/g) || [];

    numberMatches.forEach((numberText) => {
      const cleanNumber = String(numberText || "")
        .replace(/^0+(?=\d)/g, "")
        .trim();

      if (!cleanNumber) return;
      if (candidates.includes(cleanNumber)) return;

      candidates.push(cleanNumber);
    });
  };

  addCandidate(item?.code);
  addCandidate(item?.name);
  addCandidate(item?.pictureFileName);
  addCandidate(item?.category);
  addCandidate(item?.sheetName);

  return candidates;
};

const getEquipmentPictureMatchFromMap = (item, pictureMap = {}) => {
  const candidateCodes = getEquipmentPictureCandidateCodes(item);

  for (const codeKey of candidateCodes) {
    const pictureUrl = pictureMap[codeKey];

    if (pictureUrl) {
      return {
        codeKey,
        pictureUrl,
      };
    }
  }

  return {
    codeKey: "",
    pictureUrl: "",
  };
};

const getEquipmentPictureFromLibrary = (item) => {
  return getEquipmentPictureMatchFromMap(
    item,
    drivePictureLibraryByCode
  ).pictureUrl;
};

const getEquipmentImageCandidates = (item) => {
  const driveLibraryImage = getEquipmentPictureFromLibrary(item);

  const rawCandidates = [
    driveLibraryImage,
    item?.image,
    item?.imageFallback,
    item?.photo,
    item?.photoUrl,
    item?.picture,
    item?.pictureUrl,
    item?.imageUrl,
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  const stableCandidates = rawCandidates
    .filter((value) => !isTemporaryGoogleThumbnail(value))
    .filter((value, index, array) => array.indexOf(value) === index)
    .filter((value) => isUsableImageValue(value));

  const temporaryCandidates = rawCandidates
    .filter((value) => isTemporaryGoogleThumbnail(value))
    .filter((value, index, array) => array.indexOf(value) === index)
    .filter((value) => isUsableImageValue(value));

  return [...stableCandidates, ...temporaryCandidates];
};

const getEquipmentDisplayImage = (item) => {
  const candidates = getEquipmentImageCandidates(item);
  return candidates[0] || "";
};

const getEquipmentFallbackImage = (item) => {
  const candidates = getEquipmentImageCandidates(item);
  return candidates[1] || candidates[2] || "";
};

  useEffect(() => {
  if (module !== "equipment") return;
  if (equipmentDepartment !== "culinary") return;
  if (equipmentMode !== "muster" && equipmentMode !== "makeinventory") return;

  loadDrivePictureLibrary({ silent: true });
  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [module, equipmentDepartment, equipmentMode]);
  const syncMasterInventoryPicturesFromDrive = async () => {
  if (!isAdmin) {
    setPictureLibraryMessage("Only admin can sync the picture library.");
    return;
  }

  if (pictureLibraryBusy) return;

  const sourceItems = makeInventoryItems.length ? makeInventoryItems : musterItems;

  if (!sourceItems.length) {
    const text =
      "No master inventory items found. Refresh the shared master list first.";

    setPictureLibraryMessage(text);
    window.alert(text);
    return;
  }

  setPictureLibraryBusy(true);
  setPictureLibraryMessage("Loading picture library and matching pictures...");

  try {
    const pictureMap = await loadDrivePictureLibrary({ silent: true });

    if (!Object.keys(pictureMap || {}).length) {
      throw new Error(
        "Picture library loaded, but no picture code keys were prepared."
      );
    }

    const updatedItems = attachPicturesFromLibraryToItems(
      sourceItems,
      pictureMap
    );

    const matchedCount = updatedItems.filter((item) => {
      const match = getPictureMatchForItemFromMap(item, pictureMap);
      return Boolean(match.pictureUrl);
    }).length;

    setMakeInventoryItems(updatedItems);
    setMusterItems(updatedItems);

    setPictureLibraryMessage(
      "Matched " +
        matchedCount +
        " picture(s). Saving updated master list..."
    );

    await saveMasterInventoryItems(null, updatedItems);

    setPictureLibraryMessage(
      "Picture sync completed. " +
        matchedCount +
        " master item(s) matched with picture library."
    );

    logUsageEvent("equipment_picture_library_synced", {
      module: "make_inventory",
      ship: makeInventoryShip || userShip,
      equipmentDepartment,
      matchedCount,
      pictureKeys: Object.keys(pictureMap || {}).length,
    });
  } catch (error) {
    const text = error?.message || "Could not sync picture library.";
    setPictureLibraryMessage(text);
    window.alert(text);
  } finally {
    setPictureLibraryBusy(false);
  }
};
  const downloadMasterWithDrivePictureLinksInColumnH = async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  if (!isAdmin) {
    const text = "Only admin can create the master file with Drive picture links.";
    setPictureLibraryMessage(text);
    window.alert(text);
    event.target.value = "";
    return;
  }

  if (equipmentDepartment !== "culinary") {
    const text = "This tool is for Culinary master inventory pictures.";
    setPictureLibraryMessage(text);
    window.alert(text);
    event.target.value = "";
    return;
  }

  setPictureLibraryBusy(true);
  setPictureLibraryMessage("Loading Drive picture library and matching links to column H...");

  try {
    const response = await fetch("/api/drive-picture-library");
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data?.error || "Could not load Drive picture library.");
    }

    const pictureByNumber = new Map();

    (data.files || []).forEach((driveFile) => {
      (driveFile.numbers || []).forEach((number) => {
        if (!pictureByNumber.has(number)) {
          pictureByNumber.set(number, driveFile);
        }
      });
    });

    if (!pictureByNumber.size) {
      throw new Error("No picture files with equipment codes were found in the Drive folder.");
    }

    const arrayBuffer = await file.arrayBuffer();
    const isXlsm = /\.xlsm$/i.test(file.name);

    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
      bookVBA: isXlsm,
    });

    const getDrivePictureUrl = (driveFile) =>
      driveFile?.thumbnailUrl ||
      driveFile?.imageUrl ||
      driveFile?.webViewLink ||
      "";

    const findHeaderInfo = (rows) => {
      let headerRowIndex = 0;
      let codeIndex = 3; // D
      let nameIndex = 4; // E

      rows.slice(0, 20).some((row, rowIndex) => {
        const cleanRow = row.map((cell) => cleanText(cell));

        const foundCodeIndex = cleanRow.findIndex(
          (cell) => cell === "CODE" || cell.includes("APOLLO") || cell.includes("VV CODE")
        );

        const foundNameIndex = cleanRow.findIndex(
          (cell) =>
            cell.includes("FINAL DESCRIPTION") ||
            cell.includes("DESCRIPTION") ||
            cell.includes("ITEM NAME") ||
            cell === "NAME"
        );

        if (foundCodeIndex >= 0 && foundNameIndex >= 0) {
          headerRowIndex = rowIndex;
          codeIndex = foundCodeIndex;
          nameIndex = foundNameIndex;
          return true;
        }

        return false;
      });

      return { headerRowIndex, codeIndex, nameIndex };
    };

    const ensureColumnHInRange = (worksheet) => {
      const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:H1");

      if (range.e.c < 7) {
        range.e.c = 7;
        worksheet["!ref"] = XLSX.utils.encode_range(range);
      }
    };

    let matchedRows = 0;
    let unmatchedRows = 0;

    workbook.SheetNames.forEach((sheetName) => {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet || !worksheet["!ref"]) return;

      const rows = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
      });

      if (!rows.length) return;

      const { headerRowIndex, codeIndex, nameIndex } = findHeaderInfo(rows);

      // H header
      worksheet[XLSX.utils.encode_cell({ r: headerRowIndex, c: 7 })] = {
        t: "s",
        v: "Picture Link",
      };

      ensureColumnHInRange(worksheet);

      rows.slice(headerRowIndex + 1).forEach((row, index) => {
        const rowIndex = headerRowIndex + 1 + index;
        const excelRow = rowIndex + 1;

        const code = String(row[codeIndex] || "").trim();
        const name = String(row[nameIndex] || "").trim();

        if (!code && !name) return;
        if (cleanText(code) === "CODE") return;
        if (cleanText(name).includes("FINAL DESCRIPTION")) return;

        const codeKey = normalizeEquipmentPictureCode(code);

        if (!codeKey) {
          unmatchedRows += 1;
          return;
        }

        const driveFile = pictureByNumber.get(codeKey);
        const pictureUrl = getDrivePictureUrl(driveFile);

        if (!pictureUrl) {
          unmatchedRows += 1;
          return;
        }

        const cellAddress = `H${excelRow}`;

        worksheet[cellAddress] = {
          t: "s",
          v: pictureUrl,
          l: {
            Target: pictureUrl,
            Tooltip: "Open equipment picture",
          },
        };

        matchedRows += 1;
      });
    });

    const baseName = file.name.replace(/\.(xlsx|xlsm|xls)$/i, "");
    const outputName = `${baseName}-picture-links-column-H.${isXlsm ? "xlsm" : "xlsx"}`;

    XLSX.writeFile(workbook, outputName, {
      bookType: isXlsm ? "xlsm" : "xlsx",
      bookVBA: isXlsm,
    });

    setPictureLibraryMessage(
      `Picture links written to column H. Matched ${matchedRows} item(s). ${unmatchedRows} item(s) had no picture match. Downloaded: ${outputName}`
    );

    logUsageEvent("equipment_picture_links_written_to_column_h", {
      module: "equipment_picture_links",
      equipmentDepartment,
      fileName: file.name,
      matchedRows,
      unmatchedRows,
      driveFiles: data.count || 0,
    });
  } catch (error) {
    const text = error?.message || "Could not write picture links to column H.";
    setPictureLibraryMessage(text);
    window.alert(text);
  } finally {
    setPictureLibraryBusy(false);
    event.target.value = "";
  }
};
  const writeDrivePictureLinksToGoogleSheetColumnH = async () => {
  if (!isAdmin) {
    const text = "Only admin can write Drive picture links to the Google Sheet.";
    setPictureLibraryMessage(text);
    window.alert(text);
    return;
  }

  if (pictureLibraryBusy) return;

  const adminCode = window.prompt("Enter admin upload code:");

  if (!adminCode) {
    setPictureLibraryMessage("Google Sheet picture link update cancelled.");
    return;
  }

  const confirmed = window.confirm(
    "This will match picture filenames by product code and write the matched picture links into column H of the Google Sheet. Continue?"
  );

  if (!confirmed) return;

  setPictureLibraryBusy(true);
  setPictureLibraryMessage("Writing Drive picture links to Google Sheet column H...");

  try {
    const response = await fetch("/api/admin/write-picture-links-to-sheet", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        userEmail: normalizeAppEmail(userEmail),
        adminCode,
      }),
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      throw new Error(data?.error || "Could not write picture links to Google Sheet.");
    }

    setPictureLibraryMessage(
      `Google Sheet updated. ${data.matchedRows || 0} row(s) matched and written to column H. ` +
        `${data.unmatchedRows || 0} row(s) had no matching picture. ` +
        `${data.imageFileCount || 0} image file(s) checked.`
    );

    logUsageEvent("equipment_picture_links_written_to_google_sheet", {
      module: "equipment_picture_links",
      equipmentDepartment,
      ship: makeInventoryShip || userShip,
      matchedRows: data.matchedRows || 0,
      unmatchedRows: data.unmatchedRows || 0,
      imageFileCount: data.imageFileCount || 0,
      sheetTitle: data.sheetTitle || "",
    });
  } catch (error) {
    const text = error?.message || "Could not write picture links to Google Sheet.";
    setPictureLibraryMessage(text);
    window.alert(text);
  } finally {
    setPictureLibraryBusy(false);
  }
};
  const uploadMakeInventoryFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    (async () => {
      try {
        const { workbook, items, sourceSheetName } = await parseEquipmentMasterFile({
  file,
  equipmentDepartment,
});

        setMakeInventoryItems(items);
        setMusterItems(items);
        setCurrentInventoryItem(null);
        setInventoryQty("");
        setEditingInventoryId(null);
        setShowVariance(false);

        setMasterInventorySource("Uploaded from " + file.name);
        setMakeInventoryMessage(
  activeEquipmentDepartmentLabel + " Master List loaded from " + sourceSheetName + ". Saving shared list for all users..."
);

        logUsageEvent("shared_master_inventory_uploaded", {
          module: "equipment_" + equipmentDepartment + "_make_inventory",
          equipmentDepartment,
          fileName: file.name,
          sheetCount: workbook.SheetNames.length,
          sourceSheetName,
          itemCount: items.length,
        });

        await saveMasterInventoryItems(null, items);
        e.target.value = "";
      } catch (error) {
        setMakeInventoryMessage(
          "Could not load " + activeEquipmentDepartmentLabel + " file: " + (error?.message || "Unknown error")
        );
      }
    })();
  };

    const shouldBuildProductDashboardData =
    module === "product" && productMode === "dashboard";

  const consumptionData = useMemo(
    () => (shouldBuildProductDashboardData ? consumptionRows.slice(1) : []),
    [consumptionRows, shouldBuildProductDashboardData]
  );

  const recipeData = useMemo(
    () => (shouldBuildProductDashboardData ? recipeRows.slice(1) : []),
    [recipeRows, shouldBuildProductDashboardData]
  );

  const productMatches = (selectedProductName, row) => {
    const selected = String(selectedProductName || "").trim();
    const assignedProduct = String(row[12] || "").trim();
    const productName = String(row[7] || "").trim();

    if (!selected) return false;

    return (
      productNamesMatch(selected, assignedProduct) ||
      productNamesMatch(selected, productName)
    );
  };

  const getTemplateVenueKeysForVenue = (venueKey) => {
    const selectedVenue = normalizeVenue(venueKey);
    if (!selectedVenue) return [];

    const templateVenueKeys = Object.keys(templateMap || {});
    if (!templateVenueKeys.length) return [];

    // Exact sheet-name match only. No fuzzy venue matching.
    return templateVenueKeys.filter((key) => key === selectedVenue);
  };

  const templateProductMatches = (templateProductKey, product) => {
    return productNamesMatch(templateProductKey, product);
  };

  const templateLocationExistsForVenue = (venueKey) => {
    const selectedVenue = normalizeVenue(venueKey);
    if (!selectedVenue) return false;

    if (getTemplateVenueKeysForVenue(selectedVenue).length > 0) return true;

    return Object.values(templateMap || {}).some((venueTemplates) =>
      Object.values(venueTemplates || {}).some((data) =>
        (data.templateLocations || []).some((location) => location.locationKey === selectedVenue)
      )
    );
  };

  const templateHasProduct = (venueKey, product) => {
    const selectedVenue = normalizeVenue(venueKey);
    const possibleVenueKeys = getTemplateVenueKeysForVenue(selectedVenue);

    const sheetMatch = possibleVenueKeys.some((templateVenueKey) => {
      const venueTemplates = templateMap[templateVenueKey] || {};
      return Object.keys(venueTemplates).some((templateProductKey) =>
        templateProductMatches(templateProductKey, product)
      );
    });

    if (sheetMatch) return true;

    return Object.values(templateMap || {}).some((venueTemplates) =>
      Object.entries(venueTemplates || {}).some(([templateProductKey, data]) =>
        templateProductMatches(templateProductKey, product) &&
        (data.templateLocations || []).some((location) => location.locationKey === selectedVenue)
      )
    );
  };

  const getTemplateMatches = (venueKey, product) => {
    const selectedVenue = normalizeVenue(venueKey);
    const possibleVenueKeys = getTemplateVenueKeysForVenue(selectedVenue);
    const matches = [];

    possibleVenueKeys.forEach((templateVenueKey) => {
      const venueTemplates = templateMap[templateVenueKey] || {};

      Object.entries(venueTemplates).forEach(([templateProductKey, data]) => {
        if (templateProductMatches(templateProductKey, product)) {
          matches.push(...data.templates);
        }
      });
    });

    Object.values(templateMap || {}).forEach((venueTemplates) => {
      Object.entries(venueTemplates || {}).forEach(([templateProductKey, data]) => {
        if (!templateProductMatches(templateProductKey, product)) return;

        (data.templateLocations || []).forEach((location) => {
          if (location.locationKey === selectedVenue && location.templateName) {
            matches.push(location.templateName);
          }
        });
      });
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

  const getTemplateRequiredVenuesForProduct = (product) => {
    const required = {};

    Object.entries(templateMap || {}).forEach(([venueKey, venueTemplates]) => {
      Object.entries(venueTemplates || {}).forEach(([templateProductKey, data]) => {
        if (!templateProductMatches(templateProductKey, product)) return;

        const locations = data.templateLocations?.length
          ? data.templateLocations
          : (data.templates || []).map((templateName) => ({
              locationKey: getTemplateLocationKey(venueKey, templateName),
              displayName: getTemplateLocationDisplay(venueKey, templateName),
              templateName,
            }));

        locations.forEach((location) => {
          const locationKey = location.locationKey || getTemplateLocationKey(location.sheetName || venueKey, location.templateName);
          if (!locationKey) return;

          if (!required[locationKey]) {
            required[locationKey] = {
              displayName: location.displayName || getTemplateLocationDisplay(location.sheetName || venueKey, location.templateName),
              templateMatches: new Set(),
              templateVenueKey: venueKey,
            };
          }

          if (location.templateName) required[locationKey].templateMatches.add(location.templateName);
        });
      });
    });

    Object.keys(required).forEach((locationKey) => {
      required[locationKey].templateMatches = [...required[locationKey].templateMatches];
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
    const templateRequired = getTemplateRequiredVenuesForProduct(product);
    const allVenueKeys = Array.from(
      new Set([...Object.keys(actual), ...Object.keys(required), ...Object.keys(templateRequired)])
    ).sort();

    return allVenueKeys.map((venueKey) => {
      const actualVenue = actual[venueKey];
      const requiredVenue = required[venueKey];
      const templateRequiredVenue = templateRequired[venueKey];

      const ships = {};
      SHIPS.forEach((ship) => {
        ships[ship] = actualVenue?.ships?.[ship] || 0;
      });

      const templateLoaded = Object.keys(templateMap || {}).length > 0;
      const requiredByRecipe = Boolean(requiredVenue);
      const requiredByTemplate = Boolean(templateRequiredVenue);
      const hasMatchingTemplateVenue = templateLocationExistsForVenue(venueKey);
      const inTemplate = requiredByRecipe && hasMatchingTemplateVenue && templateHasProduct(venueKey, product);
      const recipeTemplateMatches = inTemplate ? getTemplateMatches(venueKey, product) : [];
      const templateChargeMatches = templateRequiredVenue?.templateMatches || [];
      const templateMatches = [...new Set([...recipeTemplateMatches, ...templateChargeMatches])];
      const missingFromTemplate = requiredByRecipe && templateLoaded && hasMatchingTemplateVenue && !inTemplate;
      const requiredForUsage = requiredByRecipe || requiredByTemplate;

      return {
        venueKey,
        displayName: actualVenue?.displayName || requiredVenue?.displayName || templateRequiredVenue?.displayName || venueKey,
        ships,
        required: requiredForUsage,
        requiredByRecipe,
        requiredFromTemplate: requiredByTemplate,
        missingShips: visibleShips.filter((ship) => requiredForUsage && (ships[ship] || 0) === 0),
        missingFromTemplate,
        templateMatches,
      };
    });
  };

  const getProductReportProductList = () => {
    const productMap = new Map();

    const addProduct = (product) => {
      const displayProduct = String(product || "").trim();
      if (!displayProduct) return;

      const tokens = [...new Set(getProductMatchTokens(displayProduct))].sort();
      const key = tokens.length ? tokens.join("|") : cleanText(displayProduct);
      if (!key || productMap.has(key)) return;

      productMap.set(key, displayProduct);
    };

    products.forEach(addProduct);

    recipeData.forEach((row) => {
      addProduct(row[12] || row[7]);
    });

    Object.values(templateMap || {}).forEach((venueTemplates) => {
      Object.values(venueTemplates || {}).forEach((data) => {
        addProduct(data.product);
      });
    });

    return [...productMap.values()].sort((a, b) => a.localeCompare(b));
  };

  const getConsumptionCostReportRows = () => {
    let currentVenue = "";
    const productMap = new Map();

    const getBlankShipCost = () => ({
      qty: 0,
      cost: 0,
      unitPriceSum: 0,
      unitPriceCount: 0,
      unitPrice: 0,
      isLowestUnitPrice: false,
    });

    consumptionData.forEach((row) => {
      if (row[2]) currentVenue = String(row[2]).trim();

      const product = String(row[6] || "").replace(/\s+/g, " ").trim();
      if (!product) return;

      const productKey = getReportProductKey(product);
      const venueDisplay = currentVenue || "Unknown";
      const venueKey = normalizeVenue(venueDisplay);
      if (!productKey || !venueKey) return;

      const code = String(row[5] || "").trim();

      if (!productMap.has(productKey)) {
        productMap.set(productKey, {
          productKey,
          product,
          code,
          totalQty: 0,
          totalCost: 0,
          venuesMap: new Map(),
          firstOrder: productMap.size + 1,
        });
      }

      const productRecord = productMap.get(productKey);
      if (!productRecord.code && code) productRecord.code = code;

      if (!productRecord.venuesMap.has(venueKey)) {
        productRecord.venuesMap.set(venueKey, {
          venueKey,
          location: venueDisplay,
          ships: {},
          totalQty: 0,
          totalCost: 0,
          hasPriceDifference: false,
        });
      }

      const venueRecord = productRecord.venuesMap.get(venueKey);

      SHIPS.forEach((ship) => {
        if (!venueRecord.ships[ship]) venueRecord.ships[ship] = getBlankShipCost();

        const qty = Number(row[shipColumns[ship]] || 0);
        const cost = Number(row[shipCostColumns[ship]] || 0);
        const unitPriceFromFile = Number(row[shipUnitPriceColumns[ship]] || 0);

        venueRecord.ships[ship].qty += qty;
        venueRecord.ships[ship].cost += cost;

        if (unitPriceFromFile > 0) {
          venueRecord.ships[ship].unitPriceSum += unitPriceFromFile;
          venueRecord.ships[ship].unitPriceCount += 1;
        }

        productRecord.totalQty += qty;
        productRecord.totalCost += cost;
      });
    });

    return Array.from(productMap.values())
      .map((productRecord) => {
        const venues = Array.from(productRecord.venuesMap.values())
          .map((venueRecord) => {
            SHIPS.forEach((ship) => {
              const shipData = venueRecord.ships[ship] || getBlankShipCost();

              if (shipData.qty > 0 && shipData.cost > 0) {
                shipData.unitPrice = shipData.cost / shipData.qty;
              } else if (shipData.unitPriceCount > 0) {
                shipData.unitPrice = shipData.unitPriceSum / shipData.unitPriceCount;
              } else {
                shipData.unitPrice = 0;
              }

              venueRecord.ships[ship] = shipData;
            });

            const visibleShipData = visibleShips.map((ship) => venueRecord.ships[ship] || getBlankShipCost());
            const visiblePrices = visibleShipData
              .map((shipData) => Number(shipData.unitPrice || 0))
              .filter((price) => price > 0);

            const minPrice = visiblePrices.length ? Math.min(...visiblePrices) : 0;
            const maxPrice = visiblePrices.length ? Math.max(...visiblePrices) : 0;
            const hasPriceDifference = visiblePrices.length > 1 && maxPrice - minPrice > 0.009;

            visibleShips.forEach((ship) => {
              const shipData = venueRecord.ships[ship] || getBlankShipCost();
              shipData.isLowestUnitPrice = hasPriceDifference && Number(shipData.unitPrice || 0) > 0 && Math.abs(Number(shipData.unitPrice || 0) - minPrice) <= 0.009;
              venueRecord.ships[ship] = shipData;
            });

            venueRecord.totalQty = visibleShips.reduce((sum, ship) => sum + Number(venueRecord.ships[ship]?.qty || 0), 0);
            venueRecord.totalCost = visibleShips.reduce((sum, ship) => sum + Number(venueRecord.ships[ship]?.cost || 0), 0);
            venueRecord.hasPriceDifference = hasPriceDifference;

            return venueRecord;
          })
          .filter((venueRecord) => venueRecord.totalQty > 0 || venueRecord.totalCost > 0);

        const visibleTotalQty = venues.reduce((sum, venue) => sum + Number(venue.totalQty || 0), 0);
        const visibleTotalCost = venues.reduce((sum, venue) => sum + Number(venue.totalCost || 0), 0);

        return {
          ...productRecord,
          venues,
          visibleTotalQty,
          visibleTotalCost,
        };
      })
      .filter((productRecord) => productRecord.venues.length > 0)
      .sort((a, b) => a.firstOrder - b.firstOrder);
  };

  const getTopNotInUseByLocationReport = (limit = 50) => {
    const expectedMap = new Map();
    const usageMap = new Map();
    const templateLocationKeys = new Set();
    const templateProductLocationKeys = new Set();
    const templateLoaded = Object.keys(templateMap || {}).length > 0;
    const delimiter = "|||";

    const makeExpectedKey = (productKey, venueKey) => `${productKey}${delimiter}${venueKey}`;

    const ensureExpected = ({ product, venueKey, displayName, source, templateName }) => {
      const productName = String(product || "").trim();
      const locationKey = normalizeVenue(venueKey);
      if (!productName || !locationKey) return;

      const productKey = getReportProductKey(productName);
      if (!productKey) return;

      const key = makeExpectedKey(productKey, locationKey);

      if (!expectedMap.has(key)) {
        expectedMap.set(key, {
          productKey,
          product: productName,
          venueKey: locationKey,
          location: displayName || venueKey || locationKey,
          sourceRecipe: false,
          sourceTemplate: false,
          templateMatches: new Set(),
        });
      }

      const row = expectedMap.get(key);

      if (source === "recipe") row.sourceRecipe = true;
      if (source === "template") row.sourceTemplate = true;
      if (templateName) row.templateMatches.add(templateName);

      if ((!row.location || row.location === row.venueKey) && displayName) {
        row.location = displayName;
      }
    };

    recipeData.forEach((row) => {
      const product = String(row[12] || row[7] || "").trim();
      const venueRaw = String(row[1] || "").trim();
      const venueKey = normalizeVenue(venueRaw);

      if (!product || !venueKey) return;

      ensureExpected({
        product,
        venueKey,
        displayName: venueRaw || venueKey,
        source: "recipe",
      });
    });

    Object.entries(templateMap || {}).forEach(([templateVenueKey, venueTemplates]) => {
      const normalizedTemplateVenueKey = normalizeVenue(templateVenueKey);
      if (normalizedTemplateVenueKey) templateLocationKeys.add(normalizedTemplateVenueKey);

      Object.entries(venueTemplates || {}).forEach(([templateProductKey, data]) => {
        const product = data.product || templateProductKey;
        const productKey = getReportProductKey(product);
        if (!productKey) return;

        const locations = data.templateLocations?.length
          ? data.templateLocations
          : (data.templates || []).map((templateName) => ({
              locationKey: getTemplateLocationKey(templateVenueKey, templateName),
              displayName: getTemplateLocationDisplay(templateVenueKey, templateName),
              templateName,
              sheetName: templateVenueKey,
            }));

        locations.forEach((location) => {
          const locationKey = normalizeVenue(
            location.locationKey || getTemplateLocationKey(location.sheetName || templateVenueKey, location.templateName)
          );
          if (!locationKey) return;

          templateLocationKeys.add(locationKey);
          templateProductLocationKeys.add(makeExpectedKey(productKey, locationKey));

          ensureExpected({
            product,
            venueKey: locationKey,
            displayName: location.displayName || getTemplateLocationDisplay(location.sheetName || templateVenueKey, location.templateName),
            source: "template",
            templateName: location.templateName,
          });
        });
      });
    });

    let currentVenue = "";

    consumptionData.forEach((row) => {
      if (row[2]) currentVenue = String(row[2]).trim();

      const product = String(row[6] || "").trim();
      const venueKey = normalizeVenue(currentVenue || "Unknown");
      if (!product || !venueKey) return;

      const productKey = getReportProductKey(product);
      if (!productKey) return;

      const key = makeExpectedKey(productKey, venueKey);

      if (!usageMap.has(key)) {
        usageMap.set(key, { BRL: 0, RL: 0, SC: 0, VL: 0 });
      }

      const ships = usageMap.get(key);
      SHIPS.forEach((ship) => {
        ships[ship] += Number(row[shipColumns[ship]] || 0);
      });
    });

    const reportRows = [];

    expectedMap.forEach((expected) => {
      const key = makeExpectedKey(expected.productKey, expected.venueKey);
      const ships = usageMap.get(key) || { BRL: 0, RL: 0, SC: 0, VL: 0 };

      const missingShips = visibleShips.filter((ship) => Number(ships[ship] || 0) === 0);
      if (!missingShips.length) return;

      const visibleTotal = visibleShips.reduce((sum, ship) => sum + Number(ships[ship] || 0), 0);
      const source =
        expected.sourceRecipe && expected.sourceTemplate
          ? "Recipe/Location + Template Charge"
          : expected.sourceTemplate
            ? "Template Charge"
            : "Recipe/Location";

      const missingFromTemplate =
        expected.sourceRecipe &&
        templateLoaded &&
        templateLocationKeys.has(expected.venueKey) &&
        !templateProductLocationKeys.has(key);

      reportRows.push({
        product: expected.product,
        location: expected.location,
        venueKey: expected.venueKey,
        source,
        missingShips,
        visibleTotal,
        ships: { ...ships },
        templateMatches: [...expected.templateMatches].sort(),
        missingFromTemplate,
      });
    });

    return reportRows
      .sort((a, b) => {
        const missingDifference = b.missingShips.length - a.missingShips.length;
        if (missingDifference !== 0) return missingDifference;

        const totalDifference = a.visibleTotal - b.visibleTotal;
        if (totalDifference !== 0) return totalDifference;

        const sourcePriority = (item) => {
          if (item.source === "Template Charge") return 0;
          if (item.source === "Recipe/Location + Template Charge") return 1;
          return 2;
        };

        const sourceDifference = sourcePriority(a) - sourcePriority(b);
        if (sourceDifference !== 0) return sourceDifference;

        const locationDifference = a.location.localeCompare(b.location);
        if (locationDifference !== 0) return locationDifference;

        return a.product.localeCompare(b.product);
      })
      .slice(0, limit);
  };

  const refreshProductMissingReport = () => {
    setProductMissingReportLoading(true);
    setProductMissingReportMessage("Preparing report...");
    setProductMissingReportRows([]);

    window.setTimeout(() => {
      try {
        const rows = getTopNotInUseByLocationReport(50);
        setProductMissingReportRows(rows);
        setProductMissingReportMessage(rows.length ? "" : "No not-in-use records found for the current view.");
      } catch (error) {
        setProductMissingReportRows([]);
        setProductMissingReportMessage(error?.message || "Could not prepare report.");
      } finally {
        setProductMissingReportLoading(false);
      }
    }, 25);
  };

  const toggleProductMissingReport = () => {
    if (showProductMissingReport) {
      setShowProductMissingReport(false);
      return;
    }

    setShowProductMissingReport(true);
    refreshProductMissingReport();
  };

  const exportTopNotInUseByLocationReportToExcel = () => {
    logUsageEvent("export_excel_clicked", { module: "product_dashboard", reportMode: "top_not_in_use_by_location", ship: userShip, viewMode });
    if (productMissingReportLoading) {
      alert("Report is still preparing. Please wait a moment.");
      return;
    }

    const rows = showProductMissingReport ? productMissingReportRows : getTopNotInUseByLocationReport(50);

    if (!rows.length) {
      alert("No not-in-use product/location records found for the current view.");
      return;
    }

    const exportRows = rows.map((item, index) => {
      const shipValues = {};
      visibleShips.forEach((ship) => {
        shipValues[ship] = Number(item.ships?.[ship] || 0);
      });

      return {
        Rank: index + 1,
        Product: item.product,
        Location: item.location,
        Source: item.source,
        MissingShips: item.missingShips.join(", "),
        VisibleTotal: item.visibleTotal,
        TemplateMenu: item.templateMatches.join(", "),
        MissingFromTemplate: item.missingFromTemplate ? "Yes" : "No",
        ...shipValues,
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Top 50 Not In Use");
    XLSX.writeFile(wb, `top-50-not-in-use-${viewMode === "single" ? userShip : "all-ships"}.xlsx`);
  };

  const printTopNotInUseByLocationReport = () => {
    logUsageEvent("print_clicked", { module: "product_dashboard", reportMode: "top_not_in_use_by_location", ship: userShip, viewMode });
    if (productMissingReportLoading) {
      alert("Report is still preparing. Please wait a moment.");
      return;
    }

    const rows = showProductMissingReport ? productMissingReportRows : getTopNotInUseByLocationReport(50);

    if (!rows.length) {
      alert("No not-in-use product/location records found for the current view.");
      return;
    }

    const html = `
      <html>
        <head>
          <title>Top 50 Items Not In Use</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .bad { color: #b00020; font-weight: bold; }
            .blue { color: #0057b8; font-weight: bold; }
          </style>
        </head>
        <body>
          <h1>Top 50 Items Not In Use By Location</h1>
          <div><strong>View:</strong> ${viewMode === "single" ? escapeHtml(userShip) : "All ships"}</div>
          <div><strong>Printed:</strong> ${new Date().toLocaleString()}</div>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Product</th>
                <th>Location</th>
                <th>Source</th>
                <th>Missing Ships</th>
                <th>Ship Usage</th>
                <th>Template/Menu</th>
              </tr>
            </thead>
            <tbody>
              ${rows
                .map(
                  (item, index) => `
                    <tr>
                      <td>${index + 1}</td>
                      <td>${escapeHtml(item.product)}</td>
                      <td>${escapeHtml(item.location)}</td>
                      <td>${escapeHtml(item.source)}</td>
                      <td class="bad">${escapeHtml(item.missingShips.join(", "))}</td>
                      <td>${visibleShips.map((ship) => `${escapeHtml(ship)}: ${formatQty(item.ships?.[ship])}`).join("<br />")}</td>
                      <td class="${item.missingFromTemplate ? "blue" : ""}">${escapeHtml(item.templateMatches.join(", ") || "N/A")}</td>
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
    if (!printWindow) {
      alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  };


  const getMainConsumptionCostReportExportRows = () => {
    const rows = [];

    filteredProductCostReportRows.forEach((item) => {
      item.venues.forEach((venue) => {
        visibleShips.forEach((ship) => {
          const shipData = venue.ships?.[ship] || {};
          const qty = Number(shipData.qty || 0);
          const cost = Number(shipData.cost || 0);
          const unitPrice = Number(shipData.unitPrice || 0);
          const hasData = qty !== 0 || cost !== 0 || unitPrice !== 0;

          if (!hasData) return;

          rows.push({
            Product: item.product,
            Code: item.code || "",
            Venue: venue.location,
            Ship: ship,
            Quantity: qty,
            TotalCost: cost,
            UnitPrice: unitPrice,
            LowestUnitPrice: shipData.isLowestUnitPrice ? "Yes" : "No",
            PriceDifferenceByVenue: venue.hasPriceDifference ? "Yes" : "No",
          });
        });
      });
    });

    return rows;
  };

  const exportMainConsumptionCostReportToExcel = () => {
    logUsageEvent("export_excel_clicked", {
      module: "product_dashboard",
      reportMode: "main_consumption_cost",
      ship: userShip,
      viewMode,
      search: productCostReportSearch,
    });

    const rows = getMainConsumptionCostReportExportRows();

    if (!rows.length) {
      alert("No main consumption/cost report rows to export.");
      return;
    }

    const summaryRows = filteredProductCostReportRows.map((item) => ({
  Product: item.product,
  Code: item.code || "",
  Venues: item.venues.length,
  TotalQuantity: item.visibleTotalQty,
  TotalCost: item.visibleTotalCost,

  Region:
    item.regionalRegion === YEARLY_REGION_ALL
      ? "All regions"
      : item.regionalRegion || "",

  SuggestionBasis: item.suggestionBasis || "",
  RegionalDailyConsumption: Number(item.regionalAvgDailyQty || 0),
  RegionalTotalQty: Number(item.regionalTotalQty || 0),
  RegionalTotalDays: Number(item.regionalTotalDays || 0),
  RegionalEvidenceBlocks: Number(item.regionalEvidenceBlocks || 0),
  SuggestedRegionalParLevel: Number(item.regionalSuggestedParLevel || 0),
  SuggestedParDifference: Number(item.regionalSuggestedParDifference || 0),
  MatchedRegionalProductCode: item.regionalMatchedProductCode || "",
  MatchedRegionalProductName: item.regionalMatchedProductName || "",
}));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "Product Summary");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Venue Ship Detail");
    XLSX.writeFile(wb, "main-consumption-cost-" + (viewMode === "single" ? userShip : "all-ships") + ".xlsx");
  };

  const printMainConsumptionCostReport = () => {
    logUsageEvent("print_clicked", {
      module: "product_dashboard",
      reportMode: "main_consumption_cost",
      ship: userShip,
      viewMode,
      search: productCostReportSearch,
    });

    const rows = getMainConsumptionCostReportExportRows();

    if (!rows.length) {
      alert("No main consumption/cost report rows to print.");
      return;
    }

    const html = "" +
      "<html>" +
      "<head>" +
      "<title>Main Consumption and Cost Report</title>" +
      "<style>" +
      "body{font-family:Arial,sans-serif;padding:24px;}" +
      "h1{margin-bottom:4px;}" +
      "table{width:100%;border-collapse:collapse;margin-top:18px;font-size:11px;}" +
      "th,td{border:1px solid #ccc;padding:6px;text-align:left;vertical-align:top;}" +
      "th{background:#f2f2f2;}" +
      ".low{color:#2e7d32;font-weight:bold;}" +
      "</style>" +
      "</head>" +
      "<body>" +
      "<h1>Main Consumption and Cost Report</h1>" +
      "<div><strong>View:</strong> " + escapeHtml(viewMode === "single" ? userShip : "All ships") + "</div>" +
      "<div><strong>Search:</strong> " + escapeHtml(productCostReportSearch || "All products") + "</div>" +
      "<div><strong>Printed:</strong> " + new Date().toLocaleString() + "</div>" +
      "<table>" +
      "<thead><tr><th>Product</th><th>Code</th><th>Venue</th><th>Ship</th><th>Quantity</th><th>Total Cost</th><th>Unit Price</th><th>Lowest Price</th></tr></thead>" +
      "<tbody>" +
      rows.map((row) => (
        "<tr>" +
        "<td>" + escapeHtml(row.Product) + "</td>" +
        "<td>" + escapeHtml(row.Code) + "</td>" +
        "<td>" + escapeHtml(row.Venue) + "</td>" +
        "<td>" + escapeHtml(row.Ship) + "</td>" +
        "<td>" + formatQty(row.Quantity) + "</td>" +
        "<td>" + formatMoney(row.TotalCost) + "</td>" +
        "<td>" + formatMoney(row.UnitPrice) + "</td>" +
        "<td class=\"" + (row.LowestUnitPrice === "Yes" ? "low" : "") + "\">" + escapeHtml(row.LowestUnitPrice) + "</td>" +
        "</tr>"
      )).join("") +
      "</tbody></table></body></html>";

    const printWindow = window.open("", "_blank");
    if (!printWindow) {
      alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  };

  const getSelectedProductConsumptionReportRows = () => {
    if (!selectedProduct) return [];

    const rows = [];

    combinedBreakdown.forEach((venueItem) => {
      visibleShips.forEach((ship) => {
        const qty = Number(venueItem.ships?.[ship] || 0);
        const missingUsage = Boolean(venueItem.required && qty === 0);

        rows.push({
          Product: selectedProduct,
          Venue: venueItem.displayName,
          Ship: ship,
          Quantity: qty,
          Required: venueItem.required ? "Yes" : "No",
          MissingUsage: missingUsage ? "Yes" : "No",
          MissingFromTemplate: venueItem.missingFromTemplate ? "Yes" : "No",
          TemplateCharge: venueItem.requiredFromTemplate ? "Yes" : "No",
          TemplateMenu: venueItem.templateMatches.join(", "),
          MissingShips: venueItem.missingShips.join(", "),
        });
      });
    });

    return rows;
  };

  const exportSelectedProductConsumptionReportToExcel = () => {
    logUsageEvent("export_excel_clicked", {
      module: "product_dashboard",
      reportMode: "consumption_vs_locations_template",
      ship: userShip,
      viewMode,
      product: selectedProduct,
    });

    const rows = getSelectedProductConsumptionReportRows();

    if (!selectedProduct || !rows.length) {
      alert("Select a product before exporting the consumption report.");
      return;
    }

    const totalRows = visibleShips.map((ship) => ({
      Product: selectedProduct,
      Ship: ship,
      TotalQuantity: totalConsumption.totals[ship],
    }));

    const recipeRowsForExport = recipesForProduct.map((recipe) => ({
      Product: selectedProduct,
      RecipeCode: recipe.recipeCode,
      RecipeName: recipe.recipeName,
      Venues: recipe.venues.join(", "),
    }));

    const allergenRowsForExport = allergenWarnings.flatMap((item) =>
      item.products.map((product) => ({
        Product: selectedProduct,
        Allergen: item.allergen,
        MatchedItem: product,
      }))
    );

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(totalRows), "Totals");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Venue Ship Usage");

    if (recipeRowsForExport.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(recipeRowsForExport), "Recipes");
    }

    if (allergenRowsForExport.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allergenRowsForExport), "Allergens");
    }

    XLSX.writeFile(wb, "consumption-report-" + selectedProduct.replace(/[^a-z0-9]+/gi, "-").toLowerCase().slice(0, 40) + ".xlsx");
  };

  const printSelectedProductConsumptionReport = () => {
    logUsageEvent("print_clicked", {
      module: "product_dashboard",
      reportMode: "consumption_vs_locations_template",
      ship: userShip,
      viewMode,
      product: selectedProduct,
    });

    const rows = getSelectedProductConsumptionReportRows();

    if (!selectedProduct || !rows.length) {
      alert("Select a product before printing the consumption report.");
      return;
    }

    const html = "" +
      "<html>" +
      "<head>" +
      "<title>Consumption Report</title>" +
      "<style>" +
      "body{font-family:Arial,sans-serif;padding:24px;}" +
      "h1{margin-bottom:4px;}" +
      "table{width:100%;border-collapse:collapse;margin-top:18px;font-size:11px;}" +
      "th,td{border:1px solid #ccc;padding:6px;text-align:left;vertical-align:top;}" +
      "th{background:#f2f2f2;}" +
      ".bad{color:#b00020;font-weight:bold;}" +
      ".blue{color:#0057b8;font-weight:bold;}" +
      "</style>" +
      "</head>" +
      "<body>" +
      "<h1>Consumption vs Locations and Template</h1>" +
      "<div><strong>Product:</strong> " + escapeHtml(selectedProduct) + "</div>" +
      "<div><strong>View:</strong> " + escapeHtml(viewMode === "single" ? userShip : "All ships") + "</div>" +
      "<div><strong>Printed:</strong> " + new Date().toLocaleString() + "</div>" +
      "<table>" +
      "<thead><tr><th>Venue</th><th>Ship</th><th>Quantity</th><th>Required</th><th>Missing Usage</th><th>Missing Template</th><th>Template Charge</th><th>Template/Menu</th></tr></thead>" +
      "<tbody>" +
      rows.map((row) => (
        "<tr>" +
        "<td>" + escapeHtml(row.Venue) + "</td>" +
        "<td>" + escapeHtml(row.Ship) + "</td>" +
        "<td>" + formatQty(row.Quantity) + "</td>" +
        "<td>" + escapeHtml(row.Required) + "</td>" +
        "<td class=\"" + (row.MissingUsage === "Yes" ? "bad" : "") + "\">" + escapeHtml(row.MissingUsage) + "</td>" +
        "<td class=\"" + (row.MissingFromTemplate === "Yes" ? "blue" : "") + "\">" + escapeHtml(row.MissingFromTemplate) + "</td>" +
        "<td>" + escapeHtml(row.TemplateCharge) + "</td>" +
        "<td>" + escapeHtml(row.TemplateMenu || "N/A") + "</td>" +
        "</tr>"
      )).join("") +
      "</tbody></table></body></html>";

    const printWindow = window.open("", "_blank");
    if (!printWindow) {
      alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  };

  const sortNextOrderRows = (rows) =>
    [...rows].sort((a, b) => {
      const rowDiff = Number(a.excelRow || 0) - Number(b.excelRow || 0);
      if (rowDiff !== 0) return rowDiff;

      return String(a.product || "").localeCompare(String(b.product || ""));
    });

  const filterNextOrderRows = (rows) => {
    const term = nextOrderSearch.toLowerCase().trim();

    return rows.filter((item) => {
      const matchesFilter =
        nextOrderFilter === "all" ||
        (nextOrderFilter === "needsOrder" && Number(item.suggestedOrder || 0) > 0) ||
        (nextOrderFilter === "noConsumption" && Number(item.pastConsumption || 0) <= 0) ||
        (nextOrderFilter === "noStock" && Number(item.stockOnHand || 0) <= 0);

      if (!matchesFilter) return false;
      if (!term) return true;

      return `${item.code || ""} ${item.product || ""} ${item.uom || ""} ${item.alertLabel || ""} ${item.parLevelNote || ""} ${item.excelRow || ""}`
        .toLowerCase()
        .includes(term);
    });
  };

  const getNextOrderFilterCounts = () => {
    const rows = nextOrderRows.length ? nextOrderRows : sortNextOrderRows(nextOrderSourceRows);

    return {
      all: rows.length,
      needsOrder: rows.filter((item) => Number(item.suggestedOrder || 0) > 0).length,
      noConsumption: rows.filter((item) => Number(item.pastConsumption || 0) <= 0).length,
      noStock: rows.filter((item) => Number(item.stockOnHand || 0) <= 0).length,
    };
  };

  const generateNextOrderReport = () => {
    setNextOrderLoading(true);
    setNextOrderMessage("Generating next order...");
    setNextOrderRows([]);

    window.setTimeout(() => {
      try {
        if (!nextOrderSourceRows.length) {
          setNextOrderMessage("Upload the latest order file first.");
          setNextOrderRows([]);
          return;
        }

        const rows = sortNextOrderRows(nextOrderSourceRows);

        setNextOrderRows(rows.map((item, index) => ({ ...item, orderRank: index + 1 })));
        setNextOrderMessage(`Generated ${rows.length} product lines in the same order as the Excel file. Use filter buttons and search to find products.`);
        logUsageEvent("next_order_generated", {
          module: "generate_next_order",
          fileName: nextOrderFileName,
          rowsGenerated: rows.length,
          itemsNeedingOrder: nextOrderMeta?.itemsNeedingOrder || 0,
          parCapItems: nextOrderMeta?.parCapItems || 0,
          blueReviewItems: nextOrderMeta?.blueReviewItems || 0,
          redReviewItems: nextOrderMeta?.redReviewItems || 0,
        });
      } catch (error) {
        setNextOrderRows([]);
        setNextOrderMessage(error?.message || "Could not generate next order.");
      } finally {
        setNextOrderLoading(false);
      }
    }, 25);
  };

  const getNextOrderRowsForOutput = (respectSearch = true) => {
    const rows = nextOrderRows.length ? nextOrderRows : sortNextOrderRows(nextOrderSourceRows);
    return respectSearch ? filterNextOrderRows(rows) : rows;
  };

  const exportNextOrderToExcel = () => {
    logUsageEvent("export_excel_clicked", { module: "generate_next_order", ship: nextOrderMeta?.shipName || userShip, search: nextOrderSearch, filter: nextOrderFilter });
    if (nextOrderLoading) {
      alert("Next order is still generating. Please wait a moment.");
      return;
    }

    const rows = getNextOrderRowsForOutput();

    if (!rows.length) {
      alert("No next-order lines found. Upload the latest order file and generate the order first.");
      return;
    }

    const exportRows = rows.map((item, index) => ({
      Line: index + 1,
      Code: item.code || "",
      Product: item.product,
      UM: item.uom,
      StockOnHand: Number(item.stockOnHand || 0),
      ParLevel_Q_14Days: Number(item.parLevel || 0),
      FutureOrders_F_to_N: Number(item.futureOrders || 0),
      PastConsumption_AI_to_AN: Number(item.pastConsumption || 0),
      HistoricalSailorDays_AI5_AI6: Number(item.historicalSailorDays || 0),
      AverageConsumptionPerDay: Number(item.averageConsumptionPerDay || 0),
      DaysUntilArrival_B2_to_B3: Number(item.daysUntilArrival || 0),
      ConsumptionUntilArrival: Number(item.consumptionUntilArrival || 0),
      AvailableAtArrival: Number(item.availableAtArrival || 0),
      ProjectedVoyageNeed_B6: Number(item.projectedNeed || 0),
      RawSuggestedBeforePar: Number(item.rawSuggestedOrder || 0),
      ParMaxAllowed_Q_plus_10_percent: Number(item.parMaxAllowed || 0),
      ParCapApplied: item.parCapApplied ? "Yes" : "No",
      SuggestedNextOrder: Number(item.suggestedOrder || 0),
      Alert: item.alertLabel || "",
      ParNote: item.parLevelNote || "",
      Reason: item.orderReason || "Average daily consumption x voyage days, adjusted for stock/future orders until order arrival",
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Next Order");
    XLSX.writeFile(wb, `next-order-${nextOrderMeta?.shipName || userShip || "ship"}.xlsx`);
  };

  const printNextOrder = () => {
    logUsageEvent("print_clicked", { module: "generate_next_order", ship: nextOrderMeta?.shipName || userShip, search: nextOrderSearch, filter: nextOrderFilter });
    if (nextOrderLoading) {
      alert("Next order is still generating. Please wait a moment.");
      return;
    }

    const rows = getNextOrderRowsForOutput();

    if (!rows.length) {
      alert("No next-order lines found. Upload the latest order file and generate the order first.");
      return;
    }

    const html = `
      <html>
        <head>
          <title>Generated Next Order</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            .meta { margin: 2px 0; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .qty { color: #0057b8; font-weight: bold; }
            .blue { color: #0057b8; font-weight: bold; }
            .red { color: #b00020; font-weight: bold; }
          </style>
        </head>
        <body>
          <h1>Generated Next Order</h1>
          <div class="meta"><strong>Source file:</strong> ${escapeHtml(nextOrderFileName || "N/A")}</div>
          <div class="meta"><strong>Ship:</strong> ${escapeHtml(nextOrderMeta?.shipName || userShip || "N/A")}</div>
          <div class="meta"><strong>Order day B2:</strong> ${escapeHtml(nextOrderMeta?.orderDate || "N/A")}</div>
          <div class="meta"><strong>Arrival day B3:</strong> ${escapeHtml(nextOrderMeta?.arrivalDate || "N/A")}</div>
          <div class="meta"><strong>Days until arrival:</strong> ${formatQty(nextOrderMeta?.daysUntilArrival)}</div>
          <div class="meta"><strong>Sailors:</strong> ${formatQty(nextOrderMeta?.targetSailors)}</div>
          <div class="meta"><strong>Days:</strong> ${formatQty(nextOrderMeta?.targetDays)}</div>
          <div class="meta"><strong>Generated:</strong> ${new Date().toLocaleString()}</div>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Code</th>
                <th>Product</th>
                <th>UM</th>
                <th>Stock On Hand</th>
                <th>Par Level Q</th>
                <th>Future Orders F:N</th>
                <th>Past Consumption AI:AN</th>
                <th>Avg / Day</th>
                <th>Use Until Arrival</th>
                <th>Available At Arrival</th>
                <th>Projected Voyage Need</th>
                <th>Suggested Next Order</th>
                <th>Alert</th>
              </tr>
            </thead>
            <tbody>
              ${rows
                .map(
                  (item, index) => `
                    <tr>
                      <td>${index + 1}</td>
                      <td>${escapeHtml(item.code || "")}</td>
                      <td>${escapeHtml(item.product)}</td>
                      <td>${escapeHtml(item.uom)}</td>
                      <td>${formatQty(item.stockOnHand)}</td>
                      <td>${formatQty(item.parLevel)}</td>
                      <td>${formatQty(item.futureOrders)}</td>
                      <td>${formatQty(item.pastConsumption)}</td>
                      <td>${formatQty(item.averageConsumptionPerDay)}</td>
                      <td>${formatQty(item.consumptionUntilArrival)}</td>
                      <td>${formatQty(item.availableAtArrival)}</td>
                      <td>${formatQty(item.projectedNeed)}</td>
                      <td class="qty">${formatQty(item.suggestedOrder)}${item.parCapApplied ? " (Par cap)" : ""}</td>
                      <td class="${item.alertType === "red" ? "red" : item.alertType === "blue" || item.alertType === "order" ? "blue" : ""}">${escapeHtml(item.alertLabel || "")}</td>
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
    if (!printWindow) {
      alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  };

  const getVisibleFmlMissingRows = () => {
    const term = fmlMissingSearch.toLowerCase().trim();
    if (!term) return fmlMissingRows;

    return fmlMissingRows.filter((item) =>
      [
        item.code,
        item.product,
        item.uom,
        item.department,
        item.category,
        item.subCategory,
        item.venueText,
        (item.matchedVenues || []).join(" "),
        (item.templateLocationNames || []).join(" "),
        item.templateShipScopeNote,
        item.reason,
        item.excelRow,
        item.standardOrderRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term)
    );
  };

  const exportFmlMissingReportToExcel = () => {
    const rows = getVisibleFmlMissingRows();

    if (!rows.length) {
      alert("No FML template-matched not-used rows found. The default ERP template is attached; upload the latest order file first.");
      return;
    }

    logUsageEvent("export_excel_clicked", {
      module: "generate_next_order_fml_missing",
      ship: nextOrderMeta?.shipName || userShip,
      search: fmlMissingSearch,
      rows: rows.length,
    });

    const exportRows = rows.map((item, index) => ({
      Line: index + 1,
      FMLRow: item.excelRow,
      StandardOrderRow: item.standardOrderRow || "Not found",
      Code: item.code || "",
      Product: item.product || "",
      UM: item.uom || "",
      MatchedTemplateVenues: (item.matchedVenues || []).join(", "),
      TemplateLocations: (item.templateLocationNames || []).join(", "),
      TemplateSheets: (item.templateSheetNames || []).join(", "),
      TemplateShipScope: item.templateShipScopeNote || "Used by all ships",
      VenuesFromFMLColumnF: item.venueText || "",
      StockOnHand: Number(item.stockOnHand || 0),
      FutureOrders_F_to_N: Number(item.futureOrders || 0),
      PastConsumption_AI_to_AN: Number(item.pastConsumption || 0),
      FoundInStandardOrderTemplate: item.foundInOrderTemplate ? "Yes" : "No",
      FoundInTemplateForShip: item.foundInTemplate ? "Yes" : "No",
      Reason: item.reason || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FML Not Ordered Not Used");
    XLSX.writeFile(wb, "fml-not-ordered-not-used-" + (nextOrderMeta?.shipName || userShip || "ship") + ".xlsx");
  };

  const printFmlMissingReport = () => {
    const rows = getVisibleFmlMissingRows();

    if (!rows.length) {
      alert("No FML template-matched not-used rows found. The default ERP template is attached; upload the latest order file first.");
      return;
    }

    logUsageEvent("print_clicked", {
      module: "generate_next_order_fml_missing",
      ship: nextOrderMeta?.shipName || userShip,
      search: fmlMissingSearch,
      rows: rows.length,
    });

    const html = `
      <html>
        <head>
          <title>FML Not Ordered / Not Used</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            .meta { margin: 2px 0; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .blue { color: #0057b8; font-weight: bold; }
            .warn { color: #8a5a00; font-weight: bold; }
          </style>
        </head>
        <body>
          <h1>FML Products Not Ordered / Not Used</h1>
          <div class="meta"><strong>Source file:</strong> ${escapeHtml(nextOrderFileName || "N/A")}</div>
          <div class="meta"><strong>Ship:</strong> ${escapeHtml(nextOrderMeta?.shipName || userShip || "N/A")}</div>
          <div class="meta"><strong>Rows:</strong> ${rows.length}</div>
          <div class="meta"><strong>Generated:</strong> ${new Date().toLocaleString()}</div>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>FML Row</th>
                <th>Code</th>
                <th>Product</th>
                <th>UM</th>
                <th>Matched Venue(s)</th>
                <th>Template Scope</th>
                <th>Template Location(s)</th>
                <th>Stock</th>
                <th>Future Orders</th>
                <th>Past Consumption</th>
                <th>Reason</th>
              </tr>
            </thead>
            <tbody>
              ${rows
                .map(
                  (item, index) => `
                    <tr>
                      <td>${index + 1}</td>
                      <td>${escapeHtml(item.excelRow || "")}</td>
                      <td>${escapeHtml(item.code || "")}</td>
                      <td>${escapeHtml(item.product || "")}</td>
                      <td>${escapeHtml(item.uom || "")}</td>
                      <td class="blue">${escapeHtml((item.matchedVenues || []).join(", ") || item.venueText || "")}</td>
                      <td>${escapeHtml(item.templateShipScopeNote || "Used by all ships")}</td>
                      <td>${escapeHtml((item.templateLocationNames || []).join(", "))}</td>
                      <td>${formatQty(item.stockOnHand)}</td>
                      <td>${formatQty(item.futureOrders)}</td>
                      <td>${formatQty(item.pastConsumption)}</td>
                      <td class="warn">${escapeHtml(item.reason || "")}</td>
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
    if (!printWindow) {
      alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
  };


  const getVisibleFmlLowRows = () => {
    const term = fmlLowSearch.toLowerCase().trim();
    if (!term) return fmlLowRows;

    return fmlLowRows.filter((item) =>
      [
        item.code,
        item.product,
        item.uom,
        item.department,
        item.category,
        item.subCategory,
        item.venueText,
        (item.matchedVenues || []).join(" "),
        (item.templateLocationNames || []).join(" "),
        item.templateShipScopeNote,
        item.reason,
        item.excelRow,
        item.standardOrderRow,
      ]
        .join(" ")
        .toLowerCase()
        .includes(term)
    );
  };

  const exportFmlLowReportToExcel = () => {
    const rows = getVisibleFmlLowRows();

    if (!rows.length) {
      alert("No FML running-low rows found. The default ERP template is attached; upload the latest order file first.");
      return;
    }

    logUsageEvent("export_excel_clicked", {
      module: "generate_next_order_fml_running_low",
      ship: nextOrderMeta?.shipName || userShip,
      search: fmlLowSearch,
      rows: rows.length,
    });

    const exportRows = rows.map((item, index) => ({
      Line: index + 1,
      FMLRow: item.excelRow,
      StandardOrderRow: item.standardOrderRow || "Not found",
      Code: item.code || "",
      Product: item.product || "",
      UM: item.uom || "",
      MatchedTemplateVenues: (item.matchedVenues || []).join(", "),
      TemplateLocations: (item.templateLocationNames || []).join(", "),
      TemplateSheets: (item.templateSheetNames || []).join(", "),
      TemplateShipScope: item.templateShipScopeNote || "Used by all ships",
      VenuesFromFMLColumnF: item.venueText || "",
      StockOnHand: Number(item.stockOnHand || 0),
      FutureOrders_F_to_N: Number(item.futureOrders || 0),
      PastConsumption_AI_to_AN: Number(item.pastConsumption || 0),
      AverageConsumptionPerDay: Number(item.averageConsumptionPerDay || 0),
      ConsumptionUntilArrival: Number(item.consumptionUntilArrival || 0),
      AvailableAtArrival: Number(item.availableAtArrival || 0),
      DaysCoverAtArrival: Number(item.daysOfCoverAtArrival || 0),
      SuggestedOrder: Number(item.suggestedOrder || 0),
      Reason: item.reason || "",
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FML Running Low");
    XLSX.writeFile(wb, "fml-running-low-" + (nextOrderMeta?.shipName || userShip || "ship") + ".xlsx");
  };

  const printFmlLowReport = () => {
    const rows = getVisibleFmlLowRows();

    if (!rows.length) {
      alert("No FML running-low rows found. The default ERP template is attached; upload the latest order file first.");
      return;
    }

    logUsageEvent("print_clicked", {
      module: "generate_next_order_fml_running_low",
      ship: nextOrderMeta?.shipName || userShip,
      search: fmlLowSearch,
      rows: rows.length,
    });

    const html = `
      <html>
        <head>
          <title>FML Running Low At Arrival</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; }
            h1 { margin-bottom: 4px; }
            .meta { margin: 2px 0; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .blue { color: #0057b8; font-weight: bold; }
            .red { color: #b00020; font-weight: bold; }
            .warn { color: #8a5a00; font-weight: bold; }
          </style>
        </head>
        <body>
          <h1>FML Products Running Low By Arrival</h1>
          <div class="meta"><strong>Source file:</strong> ${escapeHtml(nextOrderFileName || "N/A")}</div>
          <div class="meta"><strong>Ship:</strong> ${escapeHtml(nextOrderMeta?.shipName || userShip || "N/A")}</div>
          <div class="meta"><strong>Rows:</strong> ${rows.length}</div>
          <div class="meta"><strong>Generated:</strong> ${new Date().toLocaleString()}</div>
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>FML Row</th>
                <th>Code</th>
                <th>Product</th>
                <th>UM</th>
                <th>Matched Venue(s)</th>
                <th>Template Scope</th>
                <th>Stock</th>
                <th>Past Consumption</th>
                <th>Avg / Day</th>
                <th>At Arrival</th>
                <th>Days Cover</th>
                <th>Suggested Order</th>
                <th>Reason</th>
              </tr>
            </thead>
            <tbody>
              ${rows
                .map(
                  (item, index) => `
                    <tr>
                      <td>${index + 1}</td>
                      <td>${escapeHtml(item.excelRow || "")}</td>
                      <td>${escapeHtml(item.code || "")}</td>
                      <td>${escapeHtml(item.product || "")}</td>
                      <td>${escapeHtml(item.uom || "")}</td>
                      <td class="blue">${escapeHtml((item.matchedVenues || []).join(", ") || item.venueText || "")}</td>
                      <td>${escapeHtml(item.templateShipScopeNote || "Used by all ships")}</td>
                      <td>${formatQty(item.stockOnHand)}</td>
                      <td>${formatQty(item.pastConsumption)}</td>
                      <td>${formatQty(item.averageConsumptionPerDay)}</td>
                      <td class="${Number(item.availableAtArrival || 0) <= 0 ? "red" : "warn"}">${formatQty(item.availableAtArrival)}</td>
                      <td>${formatQty(item.daysOfCoverAtArrival)}</td>
                      <td>${formatQty(item.suggestedOrder)}</td>
                      <td class="warn">${escapeHtml(item.reason || "")}</td>
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
    if (!printWindow) {
      alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();
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

    const combinedBreakdown =
    shouldBuildProductDashboardData && selectedProduct
      ? getCombinedVenueBreakdown(selectedProduct)
      : [];

  const recipesForProduct =
    shouldBuildProductDashboardData && selectedProduct
      ? getRecipesUsingProduct(selectedProduct)
      : [];

  const productsInRecipe =
    shouldBuildProductDashboardData && selectedRecipe
      ? getProductsInRecipe(selectedRecipe)
      : [];

  const allergenWarnings =
    shouldBuildProductDashboardData && selectedRecipe
      ? detectAllergens(productsInRecipe)
      : [];

  const filteredProducts = shouldBuildProductDashboardData
    ? products.filter((p) => p.toLowerCase().includes(search.toLowerCase()))
    : [];

  const productCostReportRows = useMemo(() => {
  try {
    const rows = getConsumptionCostReportRows();
    return Array.isArray(rows) ? rows : [];
  } catch (error) {
    console.error("Could not build product cost report rows:", error);
    return [];
  }
}, [consumptionData, viewMode, userShip]);

const productCostReportRowsWithRegionalPar = useMemo(() => {
  try {
    return enrichDashboardProductsWithRegionalPar({
      products: Array.isArray(productCostReportRows) ? productCostReportRows : [],
      yearlyRegionalConsumption,
      selectedRegion: selectedRegionalConsumptionRegion,
      voyageDays: 14,
      bufferPercent: regionalParBufferPercent,
    }).map((item) => ({
      ...item,
      venues: Array.isArray(item.venues) ? item.venues : [],
    }));
  } catch (error) {
    console.error("Could not enrich dashboard products with regional par:", error);

    return (Array.isArray(productCostReportRows) ? productCostReportRows : []).map(
      (item) => ({
        ...item,
        venues: Array.isArray(item.venues) ? item.venues : [],
        regionalHasData: false,
        regionalRegion: selectedRegionalConsumptionRegion || "",
        regionalAvgDailyQty: 0,
        regionalSuggestedParLevel: 0,
        regionalSuggestedParDifference: 0,
        regionalEvidenceBlocks: 0,
        suggestionBasis: "Regional calculation error",
      })
    );
  }
}, [
  productCostReportRows,
  yearlyRegionalConsumption,
  selectedRegionalConsumptionRegion,
  regionalParBufferPercent,
]);

const dashboardRegionalMatchedCount = productCostReportRowsWithRegionalPar.filter(
  (item) => item.regionalHasData
).length;

const filteredProductCostReportRows = useMemo(() => {
  const term = productCostReportSearch.toLowerCase().trim();
  const sourceRows = Array.isArray(productCostReportRowsWithRegionalPar)
    ? productCostReportRowsWithRegionalPar
    : [];

  if (!term) return sourceRows;

  return sourceRows.filter((item) => {
    const venues = Array.isArray(item.venues) ? item.venues : [];
    const venueText = venues.map((venue) => venue.location || "").join(" ");

    return (
      String(item.product || "") +
      " " +
      String(item.code || "") +
      " " +
      venueText +
      " " +
      String(item.regionalRegion || "") +
      " " +
      String(item.suggestionBasis || "")
    )
      .toLowerCase()
      .includes(term);
  });
}, [productCostReportRowsWithRegionalPar, productCostReportSearch]);
  const yearlyRegionalMonthOptions = useMemo(() => {
  const monthMap = new Map();

  (yearlyRegionalConsumption?.rows || []).forEach((row) => {
    const monthKey = String(row.monthKey || "").trim();
    const monthName = String(row.monthName || monthKey || "").trim();

    if (!monthKey) return;

    monthMap.set(monthKey, monthName || monthKey);
  });

  return Array.from(monthMap.entries())
    .sort(([leftKey], [rightKey]) => leftKey.localeCompare(rightKey))
    .map(([monthKey, monthName]) => ({
      monthKey,
      monthName,
    }));
}, [yearlyRegionalConsumption]);

const yearlyRegionalActiveMonthKeys = useMemo(() => {
  if (yearlyRegionalReportMonths.length) {
    return yearlyRegionalReportMonths;
  }

  return yearlyRegionalMonthOptions.map((month) => month.monthKey);
}, [yearlyRegionalReportMonths, yearlyRegionalMonthOptions]);

const toggleYearlyRegionalReportMonth = (monthKey) => {
  setYearlyRegionalReportMonths((current) => {
    if (current.includes(monthKey)) {
      return current.filter((key) => key !== monthKey);
    }

    return [...current, monthKey].sort();
  });
};

const yearlyRegionalConsumptionReportRows = useMemo(() => {
  const sourceRows = Array.isArray(yearlyRegionalConsumption?.rows)
    ? yearlyRegionalConsumption.rows
    : [];

  const selectedMonthSet = yearlyRegionalReportMonths.length
    ? new Set(yearlyRegionalReportMonths)
    : null;

  const query = yearlyRegionalReportSearch.toLowerCase().trim();
  const grouped = new Map();

  sourceRows.forEach((row) => {
    const region = String(row.region || "").trim();
    const ship = String(row.ship || "").trim();
    const monthKey = String(row.monthKey || "").trim();

    if (!monthKey) return;

    if (
      yearlyRegionalReportRegion !== YEARLY_REGION_ALL &&
      region !== yearlyRegionalReportRegion
    ) {
      return;
    }

    if (
      yearlyRegionalReportShip !== YEARLY_REPORT_ALL_SHIPS &&
      ship !== yearlyRegionalReportShip
    ) {
      return;
    }

    if (selectedMonthSet && !selectedMonthSet.has(monthKey)) {
      return;
    }

    const qty = Number(row.qty || 0);
    const value = Number(row.value || 0);
    const price = Number(row.price || 0);
    const days = Number(row.days || 0);

    if (!qty && !value) return;

    const productCode = String(row.productCode || "").trim();
    const productName = String(row.productName || "").replace(/\s+/g, " ").trim();
    const productKey =
      productCode ||
      String(row.productKey || "").trim() ||
      cleanText(productName);

    if (!productKey && !productName) return;

    const groupKey = productCode || productKey || cleanText(productName);

    if (!grouped.has(groupKey)) {
      grouped.set(groupKey, {
        productCode,
        productName,
        productKey,
        unitMeasure: row.unitMeasure || "",
        categoryName: row.categoryName || "",
        subCategoryName: row.subCategoryName || "",
        totalQty: 0,
        totalValue: 0,
        totalDays: 0,
        priceSum: 0,
        priceCount: 0,
        months: {},
        regions: {},
        ships: {},
        blocks: 0,
      });
    }

    const item = grouped.get(groupKey);

    item.totalQty += qty;
    item.totalValue += value;
    item.totalDays += days;
    item.blocks += 1;

    if (price > 0) {
      item.priceSum += price;
      item.priceCount += 1;
    }

    if (!item.months[monthKey]) {
      item.months[monthKey] = {
        monthKey,
        monthName: row.monthName || monthKey,
        qty: 0,
        value: 0,
        days: 0,
        blocks: 0,
        regions: new Set(),
        ships: new Set(),
      };
    }

    item.months[monthKey].qty += qty;
    item.months[monthKey].value += value;
    item.months[monthKey].days += days;
    item.months[monthKey].blocks += 1;
    if (region) item.months[monthKey].regions.add(region);
    if (ship) item.months[monthKey].ships.add(ship);

    if (region) {
      if (!item.regions[region]) {
        item.regions[region] = {
          region,
          qty: 0,
          value: 0,
          days: 0,
          blocks: 0,
        };
      }

      item.regions[region].qty += qty;
      item.regions[region].value += value;
      item.regions[region].days += days;
      item.regions[region].blocks += 1;
    }

    if (ship) {
      if (!item.ships[ship]) {
        item.ships[ship] = {
          ship,
          qty: 0,
          value: 0,
          days: 0,
          blocks: 0,
        };
      }

      item.ships[ship].qty += qty;
      item.ships[ship].value += value;
      item.ships[ship].days += days;
      item.ships[ship].blocks += 1;
    }
  });

  return Array.from(grouped.values())
    .map((item) => {
      const months = {};

      Object.entries(item.months || {}).forEach(([monthKey, monthData]) => {
        months[monthKey] = {
          ...monthData,
          regions: Array.from(monthData.regions || []).sort(),
          ships: Array.from(monthData.ships || []).sort(),
          avgDailyQty:
            Number(monthData.days || 0) > 0
              ? Number(monthData.qty || 0) / Number(monthData.days || 0)
              : 0,
        };
      });

      const regionRows = Object.values(item.regions || {})
        .map((regionRow) => ({
          ...regionRow,
          avgDailyQty:
            Number(regionRow.days || 0) > 0
              ? Number(regionRow.qty || 0) / Number(regionRow.days || 0)
              : 0,
        }))
        .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0));

      const shipRows = Object.values(item.ships || {})
        .map((shipRow) => ({
          ...shipRow,
          avgDailyQty:
            Number(shipRow.days || 0) > 0
              ? Number(shipRow.qty || 0) / Number(shipRow.days || 0)
              : 0,
        }))
        .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0));

      return {
        ...item,
        months,
        regionRows,
        shipRows,
        avgDailyQty:
          Number(item.totalDays || 0) > 0
            ? Number(item.totalQty || 0) / Number(item.totalDays || 0)
            : 0,
        avgPrice:
          Number(item.priceCount || 0) > 0
            ? Number(item.priceSum || 0) / Number(item.priceCount || 0)
            : Number(item.totalQty || 0) > 0
            ? Number(item.totalValue || 0) / Number(item.totalQty || 0)
            : 0,
      };
    })
    .filter((item) => {
      if (!query) return true;

      return [
        item.productCode,
        item.productName,
        item.unitMeasure,
        item.categoryName,
        item.subCategoryName,
        item.regionRows.map((regionRow) => regionRow.region).join(" "),
        item.shipRows.map((shipRow) => shipRow.ship).join(" "),
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    })
    .sort((a, b) => {
      if (yearlyRegionalReportSort === "value") {
        return Number(b.totalValue || 0) - Number(a.totalValue || 0);
      }

      if (yearlyRegionalReportSort === "daily") {
        return Number(b.avgDailyQty || 0) - Number(a.avgDailyQty || 0);
      }

      if (yearlyRegionalReportSort === "product") {
        return String(a.productName || "").localeCompare(String(b.productName || ""));
      }

      return Number(b.totalQty || 0) - Number(a.totalQty || 0);
    });
}, [
  yearlyRegionalConsumption,
  yearlyRegionalReportRegion,
  yearlyRegionalReportShip,
  yearlyRegionalReportMonths,
  yearlyRegionalReportSearch,
  yearlyRegionalReportSort,
]);

const exportYearlyRegionalConsumptionReportToExcel = () => {
  if (!yearlyRegionalConsumptionReportRows.length) {
    window.alert("No yearly regional consumption rows to export.");
    return;
  }

  const activeMonthSet = new Set(yearlyRegionalActiveMonthKeys);
  const visibleMonthOptions = yearlyRegionalMonthOptions.filter((month) =>
    activeMonthSet.has(month.monthKey)
  );

  const exportRows = yearlyRegionalConsumptionReportRows.map((item, index) => {
    const row = {
      Number: index + 1,
      Code: item.productCode || "",
      Product: item.productName || "",
      UM: item.unitMeasure || "",
      Category: item.categoryName || "",
      SubCategory: item.subCategoryName || "",
      RegionFilter:
        yearlyRegionalReportRegion === YEARLY_REGION_ALL
          ? "All regions"
          : yearlyRegionalReportRegion,
      ShipFilter:
        yearlyRegionalReportShip === YEARLY_REPORT_ALL_SHIPS
          ? "All ships"
          : yearlyRegionalReportShip,
      TotalQty: Number(item.totalQty || 0),
      TotalValue: Number(item.totalValue || 0),
      TotalDays: Number(item.totalDays || 0),
      AverageDailyQty: Number(item.avgDailyQty || 0),
      AveragePrice: Number(item.avgPrice || 0),
      Blocks: Number(item.blocks || 0),
      Regions: item.regionRows.map((regionRow) => regionRow.region).join(", "),
      Ships: item.shipRows.map((shipRow) => shipRow.ship).join(", "),
    };

    visibleMonthOptions.forEach((month) => {
      const monthData = item.months?.[month.monthKey] || {};

      row[month.monthName + " Qty"] = Number(monthData.qty || 0);
      row[month.monthName + " Value"] = Number(monthData.value || 0);
      row[month.monthName + " Days"] = Number(monthData.days || 0);
      row[month.monthName + " Daily"] = Number(monthData.avgDailyQty || 0);
      row[month.monthName + " Blocks"] = Number(monthData.blocks || 0);
      row[month.monthName + " Regions"] = (monthData.regions || []).join(", ");
      row[month.monthName + " Ships"] = (monthData.ships || []).join(", ");
    });

    return row;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    "Yearly Regional Consumption"
  );

  XLSX.writeFile(workbook, "yearly-regional-consumption-report.xlsx");
};

  const sendAccessCode = () => {
    const email = normalizeAppEmail(userEmail);

    if (!isVirginVoyagesEmail(email)) {
      setEmailError("Please use your Virgin Voyages email ending with @virginvoyages.com.");
      return;
    }

    setUserEmail(email);
    setEmailError("");
    setEmailMessage("Email accepted. Enter the access code to continue.");
    setOtpLoading(false);
    setEmailCodeSent(true);

    logUsageEvent("access_code_screen_opened", {
      module: "welcome",
      userEmail: email,
    });
  };

  const verifyAccessCode = () => {
    const email = normalizeAppEmail(userEmail);
    const token = String(emailOtpCode || "").replace(/\s+/g, "").trim();

    if (!isVirginVoyagesEmail(email)) {
      setEmailError("Please use your Virgin Voyages email ending with @virginvoyages.com.");
      return;
    }

    if (!token) {
      setEmailError("Enter the access code.");
      return;
    }

    if (token !== "1818") {
      setEmailError("Access code is incorrect.");
      return;
    }

    if (typeof window !== "undefined") {
      if (rememberEmail) {
        window.localStorage.setItem(USER_EMAIL_STORAGE_KEY, email);
      } else {
        window.localStorage.removeItem(USER_EMAIL_STORAGE_KEY);
      }
    }

    setUserEmail(email);
    setEmailConfirmed(true);
    setEmailCodeSent(false);
    setEmailOtpCode("");
    setEmailError("");
    setEmailMessage("");
    setOtpLoading(false);

    logUsageEvent("email_verified_with_access_code", {
      module: "welcome",
      userEmail: email,
      remembered: rememberEmail,
    });
  };

  const resetUserEmail = () => {
    try {
      if (typeof window !== "undefined") {
        window.localStorage.removeItem(USER_EMAIL_STORAGE_KEY);
      }
    } catch {}

    setUserEmail("");
    setEmailConfirmed(false);
    setEmailError("");
    setEmailMessage("");
    setEmailOtpCode("");
    setEmailCodeSent(false);
    setRememberEmail(false);
    setUserShip("");
    setLoggedIn(false);
    setWelcomeStarted(true);
  };

  const topNotInUseReport = Array.isArray(productMissingReportRows)
    ? productMissingReportRows
    : [];

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
    const normalizedUserEmail = normalizeAppEmail(userEmail);
  const isAdmin = isAdminEmail(normalizedUserEmail);

  if (!loggedIn && !welcomeStarted) {
    return (
      <main style={styles.welcomePage}>
        <style>{`
          @keyframes vvMarquee {
            0% { transform: translateX(0); }
            100% { transform: translateX(-50%); }
          }

          @keyframes vvGlow {
            0%, 100% { box-shadow: 0 18px 50px rgba(0,0,0,0.18); }
            50% { box-shadow: 0 22px 70px rgba(176,0,32,0.28); }
          }
        `}</style>

        <section style={styles.welcomeHero}>
          <div style={styles.welcomeGlowCard}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.welcomeLogo} />

            <div style={styles.runningLineWrapper}>
              <div style={styles.runningLineTrack}>
                <span style={styles.runningLineText}>Use it • Save Time • Be the Reason Someone Smiles Today • </span>
                <span style={styles.runningLineText}>Use it • Save Time • Be the Reason Someone Smiles Today • </span>
                <span style={styles.runningLineText}>Use it • Save Time • Be the Reason Someone Smiles Today • </span>
                <span style={styles.runningLineText}>Use it • Save Time • Be the Reason Someone Smiles Today • </span>
              </div>
            </div>

            <div style={styles.ahoyStartBox}>
              <button
                style={styles.ahoyStartButton}
                onClick={() => {
                  logUsageEvent("welcome_start_clicked", { module: "welcome" });
                  setWelcomeStarted(true);
                }}
                aria-label="Start"
              >
                AHOY
              </button>
            </div>

            <p style={styles.welcomeSubtitle}>Press AHOY to START.</p>
            <div style={styles.welcomeFooterNote}></div>
          </div>
        </section>
      </main>
    );
  }

  if (!loggedIn && welcomeStarted && !emailConfirmed) {
    return (
      <main style={styles.page}>
        <section style={styles.loginCard}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
          <h1 style={styles.title}>Email Access Code</h1>
          <p style={styles.subtitle}>Enter your Virgin Voyages email and access code.</p>

          <label style={styles.label}>✉️ Virgin Voyages email</label>
          <input
            type="email"
            value={userEmail}
            disabled={otpLoading || emailCodeSent}
            onChange={(e) => {
              setUserEmail(e.target.value);
              setEmailError("");
              setEmailMessage("");
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !emailCodeSent) sendAccessCode();
            }}
            placeholder="name@virginvoyages.com"
            style={styles.searchInput}
            autoComplete="email"
          />

          {emailCodeSent && (
            <>
              <label style={styles.label}>🔐 Access code</label>
              <input
                type="password"
                inputMode="numeric"
                value={emailOtpCode}
                disabled={otpLoading}
                onChange={(e) => {
                  setEmailOtpCode(e.target.value);
                  setEmailError("");
                  setEmailMessage("");
                }}
                onKeyDown={(e) => {
                  if (e.key === "Enter") verifyAccessCode();
                }}
                placeholder="Enter access code..."
                style={styles.searchInput}
                autoComplete="one-time-code"
              />
            </>
          )}

          {emailError && <div style={styles.emailError}>{emailError}</div>}
          {emailMessage && <div style={{ ...styles.infoBox, color: "#2e7d32", fontWeight: "bold" }}>{emailMessage}</div>}

          <label style={styles.checkboxRow}>
            <input
              type="checkbox"
              checked={rememberEmail}
              onChange={(e) => setRememberEmail(e.target.checked)}
            />
            <span>Remember me on this device</span>
          </label>

          <div style={styles.infoBox}>
  <div>
    📦 Products in report:{" "}
    <strong>{filteredProductCostReportRows.length}</strong> / {productCostReportRows.length}
  </div>

  <div>
    🚢 View: <strong>{viewMode === "single" ? userShip : "All Ships"}</strong>
  </div>

  <div>
    🌎 Regional par region:{" "}
    <strong>
      {selectedRegionalConsumptionRegion === YEARLY_REGION_ALL
        ? "All regions"
        : selectedRegionalConsumptionRegion}
    </strong>
  </div>

  <div>
    📈 Regional matches:{" "}
    <strong>
      {dashboardRegionalMatchedCount} / {productCostReportRowsWithRegionalPar.length}
    </strong>
  </div>

  <div>
    📘 Source columns: C Venue, I/L/O/R Quantity, J/M/P/S Total Cost, H/K/N/Q Unit Price.
  </div>
</div>

          {!emailCodeSent ? (
            <button style={styles.primaryButton} onClick={sendAccessCode} disabled={otpLoading}>
              Continue
            </button>
          ) : (
            <>
              <button style={styles.primaryButton} onClick={verifyAccessCode} disabled={otpLoading}>
                Verify Code
              </button>
              <button
                style={styles.backButton}
                onClick={() => {
                  setEmailCodeSent(false);
                  setEmailOtpCode("");
                  setEmailError("");
                  setEmailMessage("");
                }}
                disabled={otpLoading}
              >
                Change Email
              </button>
            </>
          )}

          <button
            style={styles.backButton}
            onClick={() => {
              setWelcomeStarted(false);
              setEmailError("");
              setEmailMessage("");
              setEmailOtpCode("");
              setEmailCodeSent(false);
            }}
            disabled={otpLoading}
          >
            ← Back to AHOY
          </button>
        </section>
      </main>
    );
  }

  if (!loggedIn && welcomeStarted && emailConfirmed) {
    return (
      <main style={styles.page}>
        <section style={styles.loginCard}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.logo} />
          <h1 style={styles.title}>Choose Your Ship</h1>
          <p style={styles.subtitle}>Select your vessel to start the dashboard.</p>

          <div style={styles.infoBox}>
            <div>👤 Signed in as: <strong>{normalizeAppEmail(userEmail)}</strong></div>
            <button type="button" style={styles.inlineLinkButton} onClick={resetUserEmail}>Use different email</button>
          </div>

          <label style={styles.label}>🚢 Select your ship</label>
          <select value={userShip} onChange={(e) => setUserShip(e.target.value)} style={styles.select}>
            <option value="">Choose ship</option>
            {SHIPS.map((ship) => (
  <option key={ship} value={ship}>
    {getShipDisplayName(ship)}
  </option>
))}
          </select>

          <button
            style={styles.primaryButton}
            onClick={() => {
              if (!userShip) return;
              logUsageEvent("ship_selected", { ship: userShip, module: "welcome", userEmail: normalizeAppEmail(userEmail) });
              setLoggedIn(true);
            }}
          >
            Continue
          </button>

          <button style={styles.backButton} onClick={() => setWelcomeStarted(false)}>
            ← Back to Start
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
          <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
        </header>

        <section style={styles.card}>
          <h2 style={styles.cardTitle}>🧭 Select Module</h2>

          <div style={styles.moduleGrid}>
  {isAdmin && (
    <button
      style={styles.moduleCard}
      onClick={() => {
        setModule("admin");
        logUsageEvent("module_opened", {
          module: "admin_dashboard",
          ship: userShip,
        });
      }}
    >
      <div style={styles.moduleIcon}>🛡️</div>
      <strong>Admin Dashboard</strong>
      <span>Usage, inventory status, logs, and admin tools</span>
    </button>
  )}

  <button
    style={styles.moduleCard}
    onClick={() => {
      setModule("equipment");
      setEquipmentDepartment("culinary");
      setEquipmentMode("");
      setProductMode("");
      logUsageEvent("department_opened", {
        module: "department_culinary",
        equipmentDepartment: "culinary",
        ship: userShip,
      });
    }}
  >
    <div style={styles.moduleIcon}>👨‍🍳</div>
    <strong>Culinary</strong>
    <span>Product dashboard, next order, inventory, master list and temperature checks</span>
  </button>

  <button
    style={styles.moduleCard}
    onClick={() => {
      setModule("equipment");
      setEquipmentDepartment("bar");
      setEquipmentMode("");
      setProductMode("");
      logUsageEvent("department_opened", {
        module: "department_bar",
        equipmentDepartment: "bar",
        ship: userShip,
      });
    }}
  >
    <div style={styles.moduleIcon}>🍸</div>
    <strong>Bar</strong>
    <span>Generate next order, inventory and master list for Bar equipment</span>
  </button>

  <button
    style={styles.moduleCard}
    onClick={() => {
      setModule("equipment");
      setEquipmentDepartment("restaurant");
      setEquipmentMode("");
      setProductMode("");
      logUsageEvent("department_opened", {
        module: "department_restaurant",
        equipmentDepartment: "restaurant",
        ship: userShip,
      });
    }}
  >
    <div style={styles.moduleIcon}>🍽️</div>
    <strong>Restaurant</strong>
   <span>Inventory and master list for Restaurant equipment</span>
  </button>
</div>
        </section>
      </main>
    );
  }

  if (module === "admin") {
    if (!isAdmin) {
      return (
        <main style={styles.page}>
          <header style={styles.header}>
            <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
            <div style={styles.headerActions}>
              <button style={styles.backButton} onClick={() => setModule("")}>
                ← Modules
              </button>
              <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
            </div>
          </header>

          <section style={styles.card}>
            <h2 style={styles.cardTitle}>Access denied</h2>
            <p style={styles.emptyText}>
              This dashboard is available only for admin users.
            </p>
          </section>
        </main>
      );
    }

    return (
      <AdminDashboard
        styles={styles}
        supabase={supabase}
        userEmail={normalizedUserEmail}
        userShip={userShip}
        onBack={() => setModule("")}
      />
    );
  }
  if (module === "product" && !productMode) {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
          </div>
        </header>

        <section style={styles.card}>
          <h2 style={styles.cardTitle}>📦 Product Options</h2>
          <p style={styles.emptyText}>Choose whether you want to review products or generate the next order.</p>

          <div style={styles.moduleGrid}>
            <button
              style={styles.moduleCard}
              onClick={() => {
                setProductMode("dashboard");
                setProductReportView("main");
                logUsageEvent("product_option_opened", { module: "product_dashboard", ship: userShip });
              }}
            >
              <div style={styles.moduleIcon}>📊</div>
              <strong>Product Dashboard</strong>
              <span>Use the existing product dashboard with consumption, recipes, templates, allergens and reports.</span>
            </button>

            <button
              style={styles.moduleCard}
              onClick={() => {
                setProductMode("nextorder");
                setNextOrderRows([]);
                setFmlMissingRows([]);
        setFmlLowRows([]);
                setNextOrderSearch("");
                setFmlMissingSearch("");
        setFmlLowSearch("");
                setNextOrderFilter("all");
                setNextOrderView("order");
                setNextOrderMessage("");
                logUsageEvent("product_option_opened", { module: "generate_next_order", ship: userShip });
              }}
            >
              <div style={styles.moduleIcon}>🛒</div>
              <strong>Generate Next Order</strong>
              <span>Use the attached ERP template, upload the latest order workbook, and calculate suggested next-order quantities.</span>
            </button>
          </div>
        </section>
      </main>
    );
  }

  if (module === "product" && productMode === "nextorder") {
    return (
      <Suspense
        fallback={
          <main style={styles.page}>
            <section style={styles.card}>
              <h2 style={styles.cardTitle}>🛒 Loading Generate Next Order...</h2>
              <p style={styles.emptyText}>Preparing order tools.</p>
            </section>
          </main>
        }
      >
        <GenerateNextOrder
  styles={styles}
  userShip={userShip}
  onBack={() => {
    if (equipmentDepartment) {
      setModule("equipment");
      setEquipmentMode("");
      setProductMode("");
      return;
    }

    setProductMode("");
  }}
  logUsageEvent={logUsageEvent}
  yearlyRegionalConsumption={yearlyRegionalConsumption}
  setYearlyRegionalConsumption={setYearlyRegionalConsumption}
  yearlyRegionalFileName={yearlyRegionalFileName}
  setYearlyRegionalFileName={setYearlyRegionalFileName}
  selectedRegionalConsumptionRegion={selectedRegionalConsumptionRegion}
  setSelectedRegionalConsumptionRegion={setSelectedRegionalConsumptionRegion}
  regionalParBufferPercent={regionalParBufferPercent}
  setRegionalParBufferPercent={setRegionalParBufferPercent}
  recipeRows={recipeRows}
/>
      </Suspense>
    );
  }

  if (module === "training") {
    return (
      <Suspense
        fallback={
          <main style={styles.page}>
            <section style={styles.card}>
              <h2 style={styles.cardTitle}>🎓 Loading Training Module...</h2>
              <p style={styles.emptyText}>Preparing station training tracker.</p>
            </section>
          </main>
        }
      >
        <AppProvider
  value={{
    supabase,
    userShip,
    userEmail: normalizeAppEmail(userEmail),
    isAdmin,
    logUsageEvent,
  }}
>
  <TrainingModule
    styles={styles}
    onBack={() => {
      if (equipmentDepartment) {
        setModule("equipment");
        setEquipmentMode("");
        return;
      }

      setModule("");
    }}
  />
</AppProvider>
      </Suspense>
    );
  }
if (module === "allergen") {
  return (
    <Suspense
      fallback={
        <main style={styles.page}>
          <section style={styles.card}>
            <h2 style={styles.cardTitle}>🧬 Loading Allergen Module...</h2>
            <p style={styles.emptyText}>Preparing recipe allergen matrix.</p>
          </section>
        </main>
      }
    >
      <AppProvider
        value={{
          supabase,
          userShip,
          userEmail: normalizeAppEmail(userEmail),
          isAdmin,
          logUsageEvent,
        }}
      >
        <AllergenModule
  key={`${userShip}-${normalizeAppEmail(userEmail)}`}
  styles={styles}
  supabase={supabase}
  userShip={userShip}
  userEmail={normalizeAppEmail(userEmail)}
  isAdmin={isAdmin}
  recipeRows={recipeRows}
  setRecipeRows={setRecipeRows}
  onBack={() => {
    if (equipmentDepartment) {
      setModule("equipment");
      setEquipmentMode("");
      return;
    }

    setModule("");
  }}
  logUsageEvent={logUsageEvent}
/>
      </AppProvider>
    </Suspense>
  );
}
  if (module === "temperature") {
    return (
      <Suspense
        fallback={
          <main style={styles.page}>
            <section style={styles.card}>
              <h2 style={styles.cardTitle}>🌡️ Loading Take Temperature...</h2>
              <p style={styles.emptyText}>Preparing temperature photo tools.</p>
            </section>
          </main>
        }
      >
        <TemperatureCheckModule
          styles={styles}
          supabase={supabase}
          userShip={userShip}
          userEmail={normalizeAppEmail(userEmail)}
isAdmin={isAdmin}
          onBack={() => {
  if (equipmentDepartment) {
    setModule("equipment");
    setEquipmentMode("");
    return;
  }

  setModule("");
}}
          logUsageEvent={logUsageEvent}
        />
      </Suspense>
    );
  }
  if (module === "people") {
    return (
      <Suspense
        fallback={
          <main style={styles.page}>
            <header style={styles.header}>
              <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
              <div style={styles.headerActions}>
                <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
                <div style={styles.shipBadge}>🚢 Loading</div>
              </div>
            </header>

            <section style={styles.card}>
              <h2 style={styles.cardTitle}>👥 Loading People & Schedule...</h2>
              <p style={styles.emptyText}>Preparing the rotation planner only when needed.</p>
            </section>
          </main>
        }
      >
        <PeopleScheduleModule
          userShip={userShip}
          onBack={() => setModule("")}
          styles={styles}
          logUsageEvent={logUsageEvent}
        />
      </Suspense>
    );
  }

  if (module === "equipment" && !equipmentDepartment) {
    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setModule("")}>← Modules</button>
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
          </div>
        </header>

        <section style={styles.card}>
          <h2 style={styles.cardTitle}>🍽️ Equipment Department</h2>
          <p style={styles.emptyText}>Choose which operation area you want to work with.</p>

          <div style={styles.moduleGrid}>
            <button
              style={styles.moduleCard}
              onClick={() => {
                setEquipmentDepartment("culinary");
                setEquipmentMode("");
                logUsageEvent("equipment_department_opened", { module: "equipment_culinary", ship: userShip });
              }}
            >
              <div style={styles.moduleIcon}>👨‍🍳</div>
              <strong>Culinary</strong>
              <span>Current equipment tools: master list, inventory in use, warehouse and make inventory.</span>
            </button>

            <button
              style={styles.moduleCard}
              onClick={() => {
                setEquipmentDepartment("bar");
                setEquipmentMode("");
                logUsageEvent("equipment_department_opened", { module: "equipment_bar", ship: userShip });
              }}
            >
              <div style={styles.moduleIcon}>🍸</div>
              <strong>Bar</strong>
              <span>Master list and inventory tools for Bar equipment.</span>
            </button>

            <button
              style={styles.moduleCard}
              onClick={() => {
                setEquipmentDepartment("restaurant");
                setEquipmentMode("");
                logUsageEvent("equipment_department_opened", { module: "equipment_restaurant", ship: userShip });
              }}
            >
              <div style={styles.moduleIcon}>🍽️</div>
              <strong>Restaurant</strong>
              <span>Master list and inventory tools for Restaurant equipment.</span>
            </button>
          </div>
        </section>
      </main>
    );
  }

  const equipmentDepartmentConfig = {
    culinary: {
      label: "Culinary",
      icon: "👨‍🍳",
    },
    bar: {
      label: "Bar",
      icon: "🍸",
    },
    restaurant: {
      label: "Restaurant",
      icon: "🍽️",
    },
  };

  const activeEquipmentDepartment = equipmentDepartmentConfig[equipmentDepartment];
  const activeEquipmentDepartmentLabel = activeEquipmentDepartment?.label || "Equipment";
  const activeEquipmentDepartmentIcon = activeEquipmentDepartment?.icon || "🍽️";
  const hasEquipmentDepartment = Boolean(activeEquipmentDepartment);

  if (module === "equipment" && hasEquipmentDepartment && !equipmentMode) {
  return (
    <EquipmentDepartmentOptions
      styles={styles}
      userShip={userShip}
      equipmentDepartment={equipmentDepartment}
      activeEquipmentDepartmentLabel={activeEquipmentDepartmentLabel}
      activeEquipmentDepartmentIcon={activeEquipmentDepartmentIcon}
      getShipDisplayName={getShipDisplayName}
      onBackToModules={() => {
        setModule("");
        setEquipmentDepartment("");
        setEquipmentMode("");
        setProductMode("");
      }}
      onOpenProductDashboard={() => {
        setModule("product");
        setProductMode("dashboard");

        logUsageEvent("culinary_option_opened", {
          module: "product_dashboard",
          equipmentDepartment,
          ship: userShip,
        });
      }}
      onOpenNextOrder={() => {
        setModule("product");
        setProductMode("nextorder");
        setNextOrderRows([]);
        setFmlMissingRows([]);
        setFmlLowRows([]);
        setNextOrderSearch("");
        setFmlMissingSearch("");
        setFmlLowSearch("");
        setNextOrderFilter("all");
        setNextOrderView("order");
        setNextOrderMessage("");

        logUsageEvent("department_option_opened", {
          module: "generate_next_order",
          equipmentDepartment,
          ship: userShip,
        });
      }}
            onOpenAllergen={() => {
        setSelectedProduct("");
        setSelectedRecipe(null);
        setSearch("");
        setProductCostReportSearch("");
        setProductReportView("main");

        setModule("allergen");

        logUsageEvent("module_opened", {
          module: "allergen",
          equipmentDepartment,
          ship: userShip,
        });
      }}
      onOpenTraining={() => {
        setModule("training");

        logUsageEvent("module_opened", {
          module: "training",
          equipmentDepartment,
          ship: userShip,
        });
      }}
      onOpenMuster={() => {
        setEquipmentMode("muster");

        logUsageEvent("equipment_option_opened", {
          module: `equipment_${equipmentDepartment}_muster`,
          equipmentDepartment,
          ship: userShip,
        });
      }}
      onOpenInventory={() => {
        setEquipmentMode("inventory");

        logUsageEvent("equipment_option_opened", {
          module: `equipment_${equipmentDepartment}_inventory`,
          equipmentDepartment,
          ship: userShip,
        });
      }}
      onOpenTemperature={() => {
        setModule("temperature");

        logUsageEvent("module_opened", {
          module: "temperature_check",
          equipmentDepartment,
          ship: userShip,
        });
      }}
    />
  );
}
  if (module === "equipment" && hasEquipmentDepartment && equipmentMode === "inventory") {
  return (
    <EquipmentInventoryOptions
      styles={styles}
      userShip={userShip}
      activeEquipmentDepartmentLabel={activeEquipmentDepartmentLabel}
      getShipDisplayName={getShipDisplayName}
      onBack={() => setEquipmentMode("")}
      onOpenInUse={() => {
        setEquipmentMode("inuse");

        logUsageEvent("equipment_inventory_option_opened", {
          module: "inventory_in_use",
          equipmentDepartment,
          ship: userShip,
        });
      }}
      onOpenWarehouse={() => {
        setEquipmentMode("warehouse");

        logUsageEvent("equipment_inventory_option_opened", {
          module: "inventory_warehouse",
          equipmentDepartment,
          ship: userShip,
        });
      }}
      onOpenMakeInventory={() => {
        setEquipmentMode("makeinventory");

        logUsageEvent("equipment_inventory_option_opened", {
          module: "make_inventory",
          equipmentDepartment,
          ship: userShip,
        });
      }}
    />
  );
}

  if (module === "equipment" && hasEquipmentDepartment && equipmentMode === "makeinventory") {
    const filteredMakeInventoryItems = getFilteredMakeInventoryItems();
    const myReportRows = getMyInventoryRows();
    const summaryReportRows = getShipSummaryRows();
        const stationProgressRows = getInventoryStationProgressRows();
    const currentStationProgress = getCurrentStationProgress();
    const allStationsSubmitted = getAllInventoryStationsSubmitted();
    const startedStations = stationProgressRows.filter((item) => item.status === "started").length;
    const submittedStations = stationProgressRows.filter((item) => item.status === "submitted").length;
    const currentStationSubmitted = currentStationProgress?.status === "submitted";
    const visibleReportRows = getVisibleInventoryReportRows();
    const summaryStationOptions = getSummaryStationOptions();
    const activeInventoryStations = getActiveInventoryStationList();
    const inventoryStationLabel = getInventoryStationLabel();
    const selectedSummaryStationLabel =
      summaryStationFilter === "ALL"
        ? equipmentDepartment === "bar"
          ? "All Bars"
          : "All Stations"
        : summaryStationFilter;
    const inventoryStatusRows = getMyInventoryStatusRows();
    const statusCountedItems = inventoryStatusRows.filter((item) => item.status === "Counted");
    const statusPendingItems = inventoryStatusRows.filter((item) => item.status !== "Counted");
    const userName = getEffectiveInventoryUserName();
    const inventoryReady = Boolean(makeInventoryShip && inventoryStation && userName && supabase);
    const countedKeysForMe = new Set(myReportRows.map((item) => item.itemKey || cleanText(item.code || item.name)));
        const countedRecordByKey = new Map(
      myReportRows.map((item) => [
        item.itemKey || cleanText(item.code || item.name),
        item,
      ])
    );
    const currentInventoryDisplayImage = currentInventoryItem
  ? getEquipmentDisplayImage(currentInventoryItem)
  : "";

const currentInventoryFallbackImage = currentInventoryItem
  ? getEquipmentFallbackImage(currentInventoryItem)
  : "";

const currentInventoryOpenImage =
  currentInventoryDisplayImage || currentInventoryFallbackImage || "";

        const selectInventoryItemForCounting = (item) => {
      const itemKey = getInventoryItemKey(item);
      const countedRecord = countedRecordByKey.get(itemKey);

      if (!inventoryReady) {
        setMakeInventoryMessage("Choose ship, station and user before counting.");
        return;
      }

      if (currentStationSubmitted) {
        setMakeInventoryMessage(
          `${inventoryStation} has already submitted count. Waiting for all stations before final report.`
        );
        return;
      }

      const displayImage = getEquipmentDisplayImage(item);
const fallbackImage = getEquipmentFallbackImage(item);

setCurrentInventoryItem({
  ...item,
  image: displayImage,
  imageFallback: fallbackImage,
  itemKey,
});

      setInventoryQty(countedRecord ? String(countedRecord.qty ?? "") : "");
      setEditingInventoryId(countedRecord?.id || null);
      setMakeInventoryMessage(`Selected: ${item.name}`);
    };
    const sortedMakeInventoryItems = filteredMakeInventoryItems
      .map((item, index) => ({
        item,
        index,
        itemKey: getInventoryItemKey(item),
      }))
      .sort((a, b) => {
        const aCounted = countedKeysForMe.has(a.itemKey);
        const bCounted = countedKeysForMe.has(b.itemKey);

        if (aCounted !== bCounted) {
          return aCounted ? 1 : -1;
        }

        return a.index - b.index;
      })
      .map((entry) => entry.item);

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(makeInventoryShip || userShip)}</div>
          </div>
        </header>
      <MakeInventoryTopBar
  styles={styles}
    isAdmin={isAdmin}
  inventoryReportMode={inventoryReportMode}
  setInventoryReportMode={setInventoryReportMode}
  makeInventoryShip={makeInventoryShip}
  userShip={userShip}
  refreshMakeInventoryData={refreshMakeInventoryData}
  inventoryLoading={inventoryLoading}
  reportBusy={reportBusy}
  resetInventoryRun={resetInventoryRun}
  printInventorySummary={printInventorySummary}
  clearMyInventory={clearMyInventory}
  clearShipInventory={clearShipInventory}
  generateFinalInventoryReport={generateFinalInventoryReport}
  exportInventorySummaryToExcel={exportInventorySummaryToExcel}
  allStationsSubmitted={allStationsSubmitted}
  inventoryStation={inventoryStation}
  currentStationProgress={currentStationProgress}
  myReportRows={myReportRows}
  makeInventoryItems={makeInventoryItems}
  startedStations={startedStations}
  submittedStations={submittedStations}
  stationProgressRows={stationProgressRows}
/>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📝 {activeEquipmentDepartmentLabel} Make Inventory</h2>

            <label style={styles.label}>Choose ship for this inventory</label>
            <select
              value={makeInventoryShip}
              onChange={(e) => setMakeInventoryShip(e.target.value)}
              style={styles.select}
            >
              <option value="">Choose ship</option>
              {SHIPS.map((ship) => (
  <option key={ship} value={ship}>
    {getShipDisplayName(ship)}
  </option>
))}
            </select>

            <label style={styles.label}>Choose {inventoryStationLabel}</label>
            <select
              value={inventoryStation}
              onChange={(e) => setInventoryStation(e.target.value)}
              style={styles.select}
            >
              <option value="">Choose {inventoryStationLabel}</option>
              {activeInventoryStations.map((station) => (
                <option key={station} value={station}>{station}</option>
              ))}
            </select>

            <label style={styles.label}>Choose user</label>
            <input
              placeholder="Type your name..."
              value={inventoryUserName}
              onChange={(e) => setInventoryUserName(e.target.value)}
              style={styles.searchInput}
            />

            <label style={styles.label}>Position</label>
            <input
              placeholder="Type your position..."
              value={inventoryUserPosition}
              onChange={(e) => setInventoryUserPosition(e.target.value)}
              style={styles.searchInput}
            />

<button
  type="button"
  style={styles.backButton}
  onClick={() => loadMasterInventoryItems(makeInventoryShip || userShip)}
  disabled={masterInventoryLoading}
>
  {masterInventoryLoading ? "Refreshing..." : "🔄 Refresh Shared Master List"}
</button>
{isAdmin && equipmentDepartment === "culinary" && (
  <button
    type="button"
    style={styles.backButton}
    onClick={syncMasterInventoryPicturesFromDrive}
    disabled={pictureLibraryBusy || masterInventoryLoading}
  >
    {pictureLibraryBusy
      ? "Syncing pictures..."
      : "🖼️ Sync Pictures With Master List"}
  </button>
)}

            {makeInventoryMessage && <p style={styles.message}>{makeInventoryMessage}</p>}

            <div style={styles.infoBox}>
              <div>🚢 Inventory ship: <strong>{makeInventoryShip || "Not selected"}</strong></div>
              <div>📍 {equipmentDepartment === "bar" ? "Bar" : "Station"}: <strong>{inventoryStation || "Not selected"}</strong></div>
              <div>👤 User: <strong>{userName || "Not selected"}</strong></div>
              <div>📋 Shared master items: <strong>{makeInventoryItems.length}</strong></div>
              <div>🖼️ Drive pictures loaded: <strong>{Object.keys(drivePictureLibraryByCode).length}</strong></div>
              <div>📂 Master source: <strong>{masterInventorySource || "Not loaded"}</strong></div>
              {masterInventoryLoading && <div>Loading shared master inventory...</div>}
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
              <div style={styles.card}>
  <h2 style={styles.cardTitle}>➕ Item Not In Master List</h2>

  <p style={styles.emptyText}>
    Use this only when the item is found during inventory but does not exist in the shared master list.
  </p>

  <label style={styles.label}>Code / SKU optional</label>
  <input
    placeholder="Enter code or SKU..."
    value={extraInventoryCode}
    onChange={(event) => setExtraInventoryCode(event.target.value)}
    style={styles.searchInput}
  />

  <label style={styles.label}>Item name</label>
  <input
    placeholder="Enter item name..."
    value={extraInventoryName}
    onChange={(event) => setExtraInventoryName(event.target.value)}
    style={styles.searchInput}
  />

  <label style={styles.label}>Quantity</label>
  <input
    type="number"
    min="0"
    placeholder="Enter quantity..."
    value={extraInventoryQty}
    onChange={(event) => setExtraInventoryQty(event.target.value)}
    style={styles.searchInput}
  />

  <label style={styles.label}>Take / upload photo optional</label>
  <input
    key={extraInventoryPhotoInputKey}
    type="file"
    accept="image/*"
    capture="environment"
    onChange={(event) =>
      setExtraInventoryPhotoFile(event.target.files?.[0] || null)
    }
    style={styles.fileInput}
  />

  {extraInventoryPhotoFile && (
    <div style={styles.statusNeutral}>
      Photo selected: {extraInventoryPhotoFile.name}
    </div>
  )}

  <button
    type="button"
    style={styles.primaryButton}
    onClick={saveExtraInventoryItem}
    disabled={extraInventorySaving || !inventoryReady || currentStationSubmitted}
  >
    {extraInventorySaving ? "Saving..." : "Save Extra Item"}
  </button>

  {extraInventoryMessage && (
    <p style={styles.message}>{extraInventoryMessage}</p>
  )}

  {!inventoryReady && (
    <p style={styles.warningText}>
      Choose ship, station and user before saving an extra item.
    </p>
  )}

  {currentStationSubmitted && (
    <p style={styles.warningText}>
      This station already submitted count. Extra items cannot be added now.
    </p>
  )}
</div>
        </section>

                <section style={styles.card}>
          <h2 style={styles.productTitle}>📦 Select Product for Inventory</h2>

          {makeInventoryItems.length === 0 && (
            <p style={styles.emptyText}>
              Upload the shared master inventory file once for this ship, or click Refresh if another user already uploaded it.
            </p>
          )}

          <div style={styles.equipmentGrid}>
            {sortedMakeInventoryItems.map((item, index) => {
              const itemKey = getInventoryItemKey(item);
              const alreadyCounted = countedKeysForMe.has(itemKey);
              const countedRecord = countedRecordByKey.get(itemKey);
              const displayImage = getEquipmentDisplayImage(item);
              const fallbackImage = getEquipmentFallbackImage(item);

              return (
                <button
                  key={`${item.sheetName}-${item.code}-${index}`}
                  style={{
  ...styles.inventoryItemCard,
  ...(alreadyCounted ? styles.countedCard : {}),
}}
                  onClick={() =>
  selectInventoryItemForCounting({
    ...item,
    image: displayImage,
    imageFallback: fallbackImage,
  })
}
                >
                  {displayImage ? (
  <div style={styles.inventoryImageFrame}>
    <img
  src={getImageUrl(displayImage, "w360")}
  alt={item.name}
  loading="lazy"
  decoding="async"
  style={styles.inventoryCardImage}
  data-fallback-src={
    fallbackImage && fallbackImage !== displayImage
      ? getImageUrl(fallbackImage, "w360")
      : ""
  }
  onError={(e) => {
  const fallbackSrc = e.currentTarget.dataset.fallbackSrc;

  if (fallbackSrc && e.currentTarget.dataset.usedFallback !== "true") {
    e.currentTarget.dataset.usedFallback = "true";
    e.currentTarget.src = fallbackSrc;
    return;
  }

  e.currentTarget.style.display = "none";
  const fallback = e.currentTarget.nextElementSibling;
  if (fallback) fallback.style.display = "flex";
}}
/>
  
    <div style={{ ...styles.inventoryNoImage, display: "none" }}>
      No image
    </div>
  </div>
) : (
  <div style={styles.inventoryNoImage}>No image</div>
)}
                  <div style={styles.recipeName}>{item.name}</div>
                  <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                  <div style={styles.recipeMeta}>Sheet: {item.sheetName}</div>
                  <div style={styles.recipeMeta}>Category: {item.category}</div>

                  {alreadyCounted && (
                    <div style={styles.statusGood}>
                      Already Counted: {formatQty(countedRecord?.qty || 0)}
                    </div>
                  )}
                </button>
              );
            })}
          </div>
                </section>

        {currentInventoryItem ? (
          <div
            style={styles.modalBackdrop}
            onClick={() => {
              setCurrentInventoryItem(null);
              setInventoryQty("");
              setEditingInventoryId(null);
            }}
          >
            <div style={styles.modalCard} onClick={(event) => event.stopPropagation()}>
              <button
                type="button"
                style={styles.closeButton}
                onClick={() => {
                  setCurrentInventoryItem(null);
                  setInventoryQty("");
                  setEditingInventoryId(null);
                }}
              >
                ✕
              </button>

              <h2 style={styles.productTitle}>
                {editingInventoryId ? "✏️ Update Quantity" : "✅ Insert Quantity"}
              </h2>

              <div style={styles.grid}>
                <div>
  {currentInventoryDisplayImage ? (
    <div>
      <div style={styles.modalPictureFrame}>
        <img
          src={getImageUrl(currentInventoryDisplayImage, "w720")}
          alt={currentInventoryItem.name}
          style={styles.modalPreviewImage}
          data-fallback-src={
            currentInventoryFallbackImage &&
            currentInventoryFallbackImage !== currentInventoryDisplayImage
              ? getImageUrl(currentInventoryFallbackImage, "w720")
              : ""
          }
          onError={(e) => {
            const fallbackSrc = e.currentTarget.dataset.fallbackSrc;

            if (
              fallbackSrc &&
              e.currentTarget.dataset.usedFallback !== "true"
            ) {
              e.currentTarget.dataset.usedFallback = "true";
              e.currentTarget.src = fallbackSrc;
              return;
            }

            e.currentTarget.style.display = "none";
            const fallback = e.currentTarget.nextElementSibling;
            if (fallback) fallback.style.display = "flex";
          }}
        />

        <div style={{ ...styles.modalNoImage, display: "none" }}>
          Picture could not be loaded
        </div>
      </div>

      <a
        href={currentInventoryOpenImage}
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
                  <h3 style={{ marginTop: 0 }}>{currentInventoryItem.name}</h3>

                  <p><strong>Ship:</strong> {makeInventoryShip || userShip}</p>
                  <p><strong>Station:</strong> {inventoryStation || "N/A"}</p>
                  <p><strong>User:</strong> {userName || "N/A"}</p>
                  <p><strong>Code:</strong> {currentInventoryItem.code || "N/A"}</p>
                  <p><strong>Sheet:</strong> {currentInventoryItem.sheetName}</p>
                  <p><strong>Category:</strong> {currentInventoryItem.category}</p>

                  {!inventoryReady && (
                    <div style={styles.warningText}>
                      Choose ship, station, and user before confirming quantity.
                    </div>
                  )}

                  {currentStationSubmitted && (
                    <div style={styles.statusWarning}>
                      This station has already submitted count. You can view this item, but cannot update quantity.
                    </div>
                  )}

                  <label style={styles.label}>Insert quantity</label>
                  <input
                    autoFocus
                    type="number"
                    min="0"
                    value={inventoryQty}
                    onChange={(event) => setInventoryQty(event.target.value)}
                    onKeyDown={(event) => {
                      if (
                        event.key === "Enter" &&
                        inventoryReady &&
                        !inventoryLoading &&
                        !currentStationSubmitted
                      ) {
                        confirmInventoryQty();
                      }
                    }}
                    style={styles.searchInput}
                    placeholder="Enter quantity..."
                  />

                  <div style={styles.headerActions}>
                    <button
                      type="button"
                      style={styles.backButton}
                      onClick={() => {
                        setCurrentInventoryItem(null);
                        setInventoryQty("");
                        setEditingInventoryId(null);
                      }}
                    >
                      Cancel
                    </button>

                    <button
                      type="button"
                      style={styles.primaryButton}
                      onClick={confirmInventoryQty}
                      disabled={!inventoryReady || inventoryLoading || currentStationSubmitted}
                    >
                      {inventoryLoading
                        ? "Saving..."
                        : editingInventoryId
                          ? "Update Quantity"
                          : "Confirm Quantity"}
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : null}

        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 20 }}>
  <h2 style={styles.productTitle}>📄 Inventory Report</h2>
</div>
          {inventoryReportMode === "summary" && (
            <div style={styles.reportFilterBox}>
              <label style={styles.label}>📍 {equipmentDepartment === "bar" ? "Bar Filter" : "Station Filter"}</label>
              <select
                value={summaryStationFilter}
                onChange={(e) => setSummaryStationFilter(e.target.value)}
                style={styles.select}
              >
                <option value="ALL">{equipmentDepartment === "bar" ? "All Bars" : "All Stations"}</option>
                {summaryStationOptions.map((station) => (
                  <option key={station} value={station}>{station}</option>
                ))}
              </select>

              <div style={styles.recipeMeta}>
                Managers can filter the ship summary by galley/station to see where equipment is being consumed.
              </div>
            </div>
          )}
<div style={styles.reportFilterBox}>
  <label style={styles.label}>
    📤 Upload Inventory Sheet Sample For Excel Report
  </label>

  <input
    type="file"
    accept=".xlsx,.xlsm"
    onChange={handleInventoryCountSheetTemplateFile}
    style={styles.fileInput}
  />

  <div style={styles.recipeMeta}>
    Export Excel will use this uploaded file as the sample/template. It will keep
    the same item positions and only write counts into column S.
  </div>

  {inventoryCountSheetTemplateName && (
    <div style={styles.statusGood}>
      Uploaded sample: {inventoryCountSheetTemplateName}
    </div>
  )}
</div>

          <div style={styles.infoBox}>
            <div>🚢 Ship: <strong>{makeInventoryShip || userShip}</strong></div>
            {inventoryReportMode === "my" ? (
              <>
                <div>📍 {equipmentDepartment === "bar" ? "Bar" : "Station"}: <strong>{inventoryStation || "Not selected"}</strong></div>
                <div>👤 User: <strong>{userName || "Not selected"}</strong></div>
                <div>✅ My records: <strong>{myReportRows.length}</strong></div>
              </>
            ) : (
              <>
                <div>🌍 Report: <strong>All users for selected ship</strong></div>
                <div>📍 Station Filter: <strong>{selectedSummaryStationLabel}</strong></div>
                <div>📦 Summary items shown: <strong>{summaryReportRows.length}</strong></div>
              </>
            )}
          </div>

          {visibleReportRows.length === 0 && (
            <p style={styles.emptyText}>
              {inventoryReportMode === "summary"
                ? `No ship summary records yet for ${selectedSummaryStationLabel}.`
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
                    {summaryStationFilter !== "ALL" && (
                      <div style={styles.statusWarning}>Filtered Station: {summaryStationFilter}</div>
                    )}
                    <div style={styles.recipeMeta}>Users: {item.users.join(", ") || "N/A"}</div>
                    <div style={styles.recipeMeta}>Records: {item.recordCount}</div>
                    <div style={styles.recipeMeta}>Last Updated: {item.confirmedAt}</div>
                  </div>
                ))
              : visibleReportRows.map((item) => (
                  <div key={item.id} style={styles.equipmentCard}>
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

          <div style={styles.finishBar}>
            {inventoryReportMode === "my" ? (
              <>
                <div>
                  <strong>{inventoryStation || "Station"}</strong>
                  <div style={styles.recipeMeta}>
                    Live count: {myReportRows.length} / {makeInventoryItems.length || 0}
                  </div>
                  <div style={styles.recipeMeta}>
                    Status: {currentStationProgress?.statusLabel || "Not Started"}
                  </div>
                </div>

                {currentStationSubmitted ? (
                  <div style={styles.statusWarning}>
                    ✅ {inventoryStation} - Count Submitted. Waiting for all stations.
                  </div>
                ) : (
                  <button
                    style={styles.primaryButton}
                    onClick={submitInventoryStationCount}
                    disabled={!inventoryReady || reportBusy}
                  >
                    ✅ Finish / Submit Station Count
                  </button>
                )}
              </>
            ) : (
              <>
                <div>
                  <strong>Final Inventory Report</strong>
                  <div style={styles.recipeMeta}>
                    Submitted stations: {submittedStations} / {stationProgressRows.length}
                  </div>
                  <div style={styles.recipeMeta}>
                    Uploaded sample: {inventoryCountSheetTemplateName || "Not uploaded"}
                  </div>
                </div>

                {allStationsSubmitted ? (
                  <button
                    style={styles.primaryButton}
                    onClick={generateFinalInventoryReport}
                    disabled={reportBusy}
                  >
                    📥 Generate Final Report
                  </button>
                ) : (
                  <div style={styles.statusWarning}>
                    ⏳ Final report will unlock after all stations submit.
                  </div>
                )}
              </>
            )}
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

  if (module === "equipment" && hasEquipmentDepartment && equipmentMode === "inuse") {
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
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📤 Upload Inventory in Use</h2>

            <label style={styles.label}>Step 1: Equipment Master List file</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadMusterFile} style={styles.fileInput} />

            <label style={styles.label}>Step 2: Inventory in Use file</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadInUseFile} style={styles.fileInput} />

            {musterMessage && <p style={styles.message}>{musterMessage}</p>}
            {inUseMessage && <p style={styles.message}>{inUseMessage}</p>}

            <div style={styles.infoBox}>
              <div>❌ Missing from Inventory: <strong>{missingItems.length}</strong></div>
              <div>⚠️ Zero Count: <strong>{zeroItems.length}</strong></div>
              <div>✅ In Use: <strong>{activeItems.length}</strong></div>
              <div>Master List: C = Category, D = Code, E = Name</div>
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

          {musterItems.length === 0 && <p style={styles.emptyText}>Upload the Equipment Master List first.</p>}
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

  if (module === "equipment" && hasEquipmentDepartment && equipmentMode === "warehouse") {
    const allWarehouseItems = parseWarehouseItems();
    const isWarehouseOverstock = (item) => Number(item.onHand || 0) - Number(item.par || 0) > 10;
    const needsWarehouseOrder = (item) => Number(item.suggested || 0) > 0;
    const warehouseItems = allWarehouseItems.filter((item) => {
      if (warehouseFilter === "overstock") return isWarehouseOverstock(item);
      if (warehouseFilter === "needsOrder") return needsWarehouseOrder(item);
      return true;
    });
    const totalSuggested = allWarehouseItems.reduce((sum, item) => sum + item.suggested, 0);
    const criticalItems = allWarehouseItems.filter(needsWarehouseOrder).length;
    const warehouseItemsWithPictures = allWarehouseItems.filter((item) => item.image).length;
    const warehouseOverstockItems = allWarehouseItems.filter(isWarehouseOverstock).length;

    return (
      <main style={styles.page}>
        <header style={styles.header}>
          <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={() => setEquipmentMode("inventory")}>← Back</button>
            <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
          </div>
        </header>

        <section style={styles.grid}>
          <div style={styles.card}>
            <h2 style={styles.cardTitle}>📤 Upload Warehouse Inventory</h2>
            <label style={styles.label}>Warehouse inventory file</label>
            <input type="file" accept=".xlsx,.xls,.xlsm" onChange={uploadWarehouseFile} style={styles.fileInput} />

            {warehouseMessage && <p style={styles.message}>{warehouseMessage}</p>}

            <div style={styles.infoBox}>
              <div>📦 Items loaded: <strong>{allWarehouseItems.length}</strong></div>
              <div>👀 Showing now: <strong>{warehouseItems.length}</strong></div>
              <div>🖼️ Pictures matched: <strong>{warehouseItemsWithPictures}</strong></div>
              <div>🔵 Items needing order: <strong>{criticalItems}</strong></div>
              <div>🛒 Total suggested order: <strong>{formatQty(totalSuggested)}</strong></div>
              <div>🔴 Over par by more than 10: <strong>{warehouseOverstockItems}</strong></div>
              <div>A = Code, B = Name, G = Par, H = On Hand, M = Future Order</div>
              <div>Pictures are matched from the shared MEL master list by code/name.</div>
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
            <p style={styles.emptyText}>Blue = suggested order needed. Red = on hand is more than 10 above par level.</p>

            <div style={styles.viewModeBox}>
              <button
                style={{ ...styles.viewModeButton, ...(warehouseFilter === "all" ? styles.viewModeButtonActive : {}) }}
                onClick={() => setWarehouseFilter("all")}
              >
                📋 All ({allWarehouseItems.length})
              </button>

              <button
                style={{ ...styles.viewModeButton, ...(warehouseFilter === "overstock" ? styles.viewModeButtonActive : {}) }}
                onClick={() => setWarehouseFilter("overstock")}
              >
                🔴 Overstock ({warehouseOverstockItems})
              </button>

              <button
                style={{ ...styles.viewModeButton, ...(warehouseFilter === "needsOrder" ? styles.viewModeButtonActive : {}) }}
                onClick={() => setWarehouseFilter("needsOrder")}
              >
                🔵 Needs Order ({criticalItems})
              </button>
            </div>

            <button style={styles.backButton} onClick={() => loadMasterInventoryItems(makeInventoryShip || userShip)}>
              🔄 Refresh Pictures
            </button>
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={styles.productTitle}>🏬 {activeEquipmentDepartmentLabel} Inventory Warehouse</h2>

          {warehouseRows.length === 0 && <p style={styles.emptyText}>Upload the warehouse inventory file to begin.</p>}
          {warehouseRows.length > 0 && warehouseItems.length === 0 && (
            <p style={styles.emptyText}>
              No warehouse items match the current search/filter.
            </p>
          )}

          <div style={styles.equipmentGrid}>
            {warehouseItems.map((item, i) => {
  const displayImage = getEquipmentDisplayImage(item);
  const overstockAmount = Number(item.onHand || 0) - Number(item.par || 0);
  const isOverstock = overstockAmount > 10;

              return (
              <div
                key={`${item.code}-${i}`}
                style={{
                  ...styles.equipmentCard,
                  ...(item.suggested > 0 ? styles.orderNeededCard : {}),
                  ...(isOverstock ? styles.overstockCard : {}),
                }}
              >
                {displayImage ? (
                  <div>
                    <img
                      src={getImageUrl(displayImage)}
                      alt={item.name}
                      style={styles.equipmentImage}
                      onClick={() => setSelectedEquipment({
                        name: item.name,
                        code: item.code,
                        category: item.masterCategory || "Warehouse Item",
                        sheetName: item.masterSheetName || "Warehouse",
                        image: displayImage,
                      })}
                      onError={(e) => {
                        e.currentTarget.style.display = "none";
                        const link = e.currentTarget.nextElementSibling;
                        if (link) link.style.display = "block";
                      }}
                    />
                    <button
                      style={styles.imageButton}
                      onClick={() => setSelectedEquipment({
                        name: item.name,
                        code: item.code,
                        category: item.masterCategory || "Warehouse Item",
                        sheetName: item.masterSheetName || "Warehouse",
                        image: displayImage,
                      })}
                    >
                      View Picture
                    </button>
                    <a href={displayImage} target="_blank" rel="noreferrer" style={styles.imageLink}>Open Picture</a>
                  </div>
                ) : (
                  <div style={styles.equipmentNoImage}>No image matched</div>
                )}

                <div style={styles.recipeName}>{item.name || "Unnamed Item"}</div>
                <div style={styles.recipeMeta}>Code: {item.code || "N/A"}</div>
                <div style={styles.recipeMeta}>Par Level: {formatQty(item.par)}</div>
                <div style={styles.recipeMeta}>On Hand: {formatQty(item.onHand)}</div>
                <div style={styles.recipeMeta}>Future Order: {formatQty(item.future)}</div>
                {item.imageSource && <div style={styles.recipeMeta}>Picture match: {item.imageSource}</div>}
                {isOverstock && (
                  <div style={styles.overstockWarning}>
                    Overstock Alert: {formatQty(overstockAmount)} above par
                  </div>
                )}
                <div style={item.suggested > 0 ? styles.suggestedOrderBad : styles.suggestedOrderGood}>
                  Suggested Next Order: {formatQty(item.suggested)}
                </div>
              </div>
              );
            })}
          </div>

          {selectedEquipment && (
  <div style={styles.modalBackdrop} onClick={() => setSelectedEquipment(null)}>
    <div style={styles.modalCard} onClick={(e) => e.stopPropagation()}>
      <button style={styles.closeButton} onClick={() => setSelectedEquipment(null)}>
        ✕
      </button>

      <h2>{selectedEquipment.name}</h2>
      <p><strong>Code:</strong> {selectedEquipment.code || "N/A"}</p>
      <p><strong>Sheet:</strong> {selectedEquipment.sheetName || "N/A"}</p>
      <p><strong>Category:</strong> {selectedEquipment.category || "N/A"}</p>

      {selectedEquipment.image || selectedEquipment.imageFallback ? (
        <div>
          <img
            src={getImageUrl(selectedEquipment.image || selectedEquipment.imageFallback)}
            alt={selectedEquipment.name}
            style={styles.modalImage}
            onError={(e) => {
              const fallbackSrc = selectedEquipment.imageFallback
                ? getImageUrl(selectedEquipment.imageFallback)
                : "";

              if (fallbackSrc && e.currentTarget.dataset.usedFallback !== "true") {
                e.currentTarget.dataset.usedFallback = "true";
                e.currentTarget.src = fallbackSrc;
                return;
              }

              e.currentTarget.style.display = "none";
              const link = e.currentTarget.nextElementSibling;
              if (link) link.style.display = "block";
            }}
          />

          <a
            href={selectedEquipment.imageFallback || selectedEquipment.image}
            target="_blank"
            rel="noreferrer"
            style={{ ...styles.imageLink, display: "block" }}
          >
            Open Picture
          </a>
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

  if (module === "equipment" && hasEquipmentDepartment && equipmentMode === "muster") {
  return (
    <EquipmentMusterModule
      styles={styles}
      userShip={userShip}
      makeInventoryShip={makeInventoryShip}
      isAdmin={isAdmin}
      equipmentDepartment={equipmentDepartment}
      activeEquipmentDepartmentLabel={activeEquipmentDepartmentLabel}
      getShipDisplayName={getShipDisplayName}
      musterItems={musterItems}
      musterMessage={musterMessage}
      pictureLibraryMessage={pictureLibraryMessage}
      pictureLibraryBusy={pictureLibraryBusy}
      masterInventoryLoading={masterInventoryLoading}
      uploadMusterFile={uploadMusterFile}
      loadMasterInventoryItems={loadMasterInventoryItems}
      syncMasterInventoryPicturesFromDrive={syncMasterInventoryPicturesFromDrive}
      uploadEquipmentPictureZipFile={uploadEquipmentPictureZipFile}
      getEquipmentDisplayImage={getEquipmentDisplayImage}
      getEquipmentFallbackImage={getEquipmentFallbackImage}
      getImageUrl={getImageUrl}
      onBack={() => setEquipmentMode("")}
    />
  );
}
  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img src="/virgin-logo.png" alt="Virgin Voyages" style={styles.headerLogo} />
        <div style={styles.headerActions}>
          <button
  style={styles.backButton}
  onClick={() => {
    if (equipmentDepartment) {
      setModule("equipment");
      setEquipmentMode("");
      setProductMode("");
      return;
    }

    setProductMode("");
  }}
>
  {equipmentDepartment ? `← ${activeEquipmentDepartmentLabel} Options` : "← Product Options"}
</button>
          <div style={styles.shipBadge}>🚢 {getShipDisplayName(userShip)}</div>
        </div>
      </header>

      <div style={styles.viewModeBox}>
        <button onClick={() => setViewMode("single")} style={{ ...styles.viewModeButton, ...(viewMode === "single" ? styles.viewModeButtonActive : {}) }}>
          🚢 {getShipDisplayName(userShip)} Only
        </button>

        <button onClick={() => setViewMode("all")} style={{ ...styles.viewModeButton, ...(viewMode === "all" ? styles.viewModeButtonActive : {}) }}>
          🌍 All Ships Overview
        </button>
      </div>

      <section style={styles.grid}>
  <div style={styles.card}>
    <h2 style={styles.cardTitle}>📤 Upload Files</h2>

    <label style={styles.label}>Step 1: Consumption file</label>
    <input
      type="file"
      accept=".xlsx,.xls,.xlsm"
      onChange={uploadConsumptionFile}
      style={styles.fileInput}
    />

    <label style={styles.label}>Step 2: Permanent Ingredient by Location file</label>

    <div style={styles.infoBox}>
      <div>
        📄 Ingredient by Location file loads automatically for all users.
      </div>
      <div>
        🔒 Only admins can replace the permanent file.
      </div>
    </div>

    <button
      type="button"
      style={styles.backButton}
      onClick={() => loadPermanentIngredientByLocationForProductDashboard()}
    >
      🔄 Reload Permanent Ingredient by Location
    </button>

    {isAdmin && (
      <>
        <label style={styles.label}>
          Admin only: replace permanent Ingredient by Location file
        </label>

        <input
          type="file"
          accept=".xlsx,.xls,.xlsm"
          onChange={uploadRecipeFile}
          style={styles.fileInput}
        />
      </>
    )}

    <label style={styles.label}>Optional: Replace template file</label>

    <input
      type="file"
      accept=".xlsx,.xls,.xlsm"
      onChange={uploadTemplateFile}
      style={styles.fileInput}
    />

    <label style={styles.label}>
      Optional: Yearly regional consumption file May 2025 - April 2026
    </label>

    <input
      type="file"
      accept=".xlsx,.xls,.xlsm"
      onChange={uploadYearlyRegionalConsumptionFile}
      style={styles.fileInput}
    />

    <label style={styles.label}>Region / home port</label>

    <select
  value={selectedRegionalConsumptionRegion}
  onChange={(e) => setSelectedRegionalConsumptionRegion(e.target.value)}
  style={styles.searchInput}
>
  <option value="">Select region / origin</option>
  <option value={YEARLY_REGION_ALL}>All regions</option>

  {(yearlyRegionalConsumption?.regionOptions || []).map((region) => (
    <option key={region} value={region}>
      {region}
    </option>
  ))}
</select>

    <label style={styles.label}>Regional par buffer %</label>

    <input
      type="number"
      min="0"
      step="1"
      value={regionalParBufferPercent}
      onChange={(e) => setRegionalParBufferPercent(Number(e.target.value || 0))}
      style={styles.searchInput}
    />

    {message && <p style={styles.message}>{message}</p>}

    <div style={styles.infoBox}>
      <div>
        📦 Products loaded: <strong>{products.length}</strong>
      </div>

      <div>
        📘 Recipe rows loaded:{" "}
        <strong>{Math.max(recipeRows.length - 1, 0)}</strong>
      </div>

      <div>
        📋 Template: <strong>{templateStatus}</strong>
      </div>

      <div>
  🌎 Yearly regional file:{" "}
  <strong>{yearlyRegionalFileName || "Not uploaded"}</strong>
</div>

<div>
  🧭 Selected region:{" "}
  <strong>
    {!selectedRegionalConsumptionRegion
      ? "Not selected"
      : selectedRegionalConsumptionRegion === YEARLY_REGION_ALL
      ? "All regions"
      : selectedRegionalConsumptionRegion}
  </strong>
</div>

<div>
  📈 Regional matches:{" "}
  <strong>
    {dashboardRegionalMatchedCount} / {productCostReportRowsWithRegionalPar.length}
  </strong>
</div>

      <div>
        🧮 Regional buffer:{" "}
        <strong>{formatQty(regionalParBufferPercent)}%</strong>
      </div>

      <div>{yearlyRegionalMessage}</div>

      <div style={{ color: "#b00020" }}>
        Red = recipe/location or template charge location expects usage, but consumption is 0 for visible ship(s).
      </div>

      <div style={{ color: "#0057b8" }}>
        Blue = product is in recipe/location, but missing from the matching venue template.
      </div>
    </div>
  </div>

  <div style={styles.card}>
    <h2 style={styles.cardTitle}>🧭 Product Report View</h2>
    <p style={styles.emptyText}>Choose the report you want to work with.</p>

    <div style={styles.reportModeGrid}>
      <button
        style={{
          ...styles.reportModeButton,
          ...(productReportView === "main" ? styles.reportModeButtonActive : {}),
        }}
        onClick={() => setProductReportView("main")}
      >
        <strong>💰 Main Report</strong>
        <span>Consumption and cost by product, venue and ship</span>
      </button>
        <button
  style={{
    ...styles.reportModeButton,
    ...(productReportView === "yearlyRegional"
      ? styles.reportModeButtonActive
      : {}),
  }}
  onClick={() => setProductReportView("yearlyRegional")}
>
  <strong>🌎 Yearly Regional Consumption</strong>
  <span>Consumption by month, region / home port, ship, and product</span>
</button>

      <button
        style={{
          ...styles.reportModeButton,
          ...(productReportView === "consumption"
            ? styles.reportModeButtonActive
            : {}),
        }}
        onClick={() => setProductReportView("consumption")}
      >
        <strong>📊 Consumption Report</strong>
        <span>Consumption vs locations and template</span>
      </button>

      <button
        style={{
          ...styles.reportModeButton,
          ...(productReportView === "reports" ? styles.reportModeButtonActive : {}),
        }}
        onClick={() => setProductReportView("reports")}
      >
        <strong>📋 Generate Report</strong>
        <span>Top 50 items not in use by location</span>
      </button>
    </div>
  </div>

  {productReportView === "consumption" && (
    <div style={styles.card}>
      <h2 style={styles.cardTitle}>🔍 Select Product</h2>

      <input
        placeholder="Search product..."
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        style={styles.searchInput}
      />

      <div style={styles.productList}>
        {filteredProducts.map((product, i) => (
          <button
            key={i}
            onClick={() => {
              setSelectedProduct(product);
              setSelectedRecipe(null);
            }}
            style={{
              ...styles.productItem,
              ...(selectedProduct === product ? styles.productItemActive : {}),
            }}
          >
            {product}
          </button>
        ))}
      </div>
    </div>
  )}
</section>

{productReportView === "yearlyRegional" && (
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
        <h2 style={styles.productTitle}>🌎 Yearly Regional Consumption</h2>
        <p style={{ ...styles.emptyText, margin: 0 }}>
          Uses the yearly May 2025 - April 2026 file. Filter by month, region /
          home port, ship, and product.
        </p>
      </div>

      <div style={styles.headerActions}>
        <button
          style={styles.primaryButton}
          onClick={exportYearlyRegionalConsumptionReportToExcel}
        >
          📥 Export Excel
        </button>
      </div>
    </div>

    {!yearlyRegionalConsumption && (
      <p style={styles.emptyText}>
        Upload or load the yearly regional consumption file to use this report.
      </p>
    )}

    <div style={styles.grid}>
      <div style={styles.card}>
        <h3 style={styles.cardTitle}>🔎 Filters</h3>

        <label style={styles.label}>Search product, code, category, region, or ship</label>
        <input
          placeholder="Search yearly regional data..."
          value={yearlyRegionalReportSearch}
          onChange={(event) => setYearlyRegionalReportSearch(event.target.value)}
          style={styles.searchInput}
        />

        <label style={styles.label}>Region / home port</label>
        <select
          value={yearlyRegionalReportRegion}
          onChange={(event) => setYearlyRegionalReportRegion(event.target.value)}
          style={styles.searchInput}
        >
          <option value={YEARLY_REGION_ALL}>All regions</option>

          {(yearlyRegionalConsumption?.regionOptions || []).map((region) => (
            <option key={region} value={region}>
              {region}
            </option>
          ))}
        </select>

        <label style={styles.label}>Ship</label>
        <select
          value={yearlyRegionalReportShip}
          onChange={(event) => setYearlyRegionalReportShip(event.target.value)}
          style={styles.searchInput}
        >
          <option value={YEARLY_REPORT_ALL_SHIPS}>All ships</option>
          <option value="BRL">BRL</option>
          <option value="RL">RL</option>
          <option value="SC">SC</option>
          <option value="VL">VL</option>
        </select>

        <label style={styles.label}>Sort by</label>
        <select
          value={yearlyRegionalReportSort}
          onChange={(event) => setYearlyRegionalReportSort(event.target.value)}
          style={styles.searchInput}
        >
          <option value="qty">Highest quantity</option>
          <option value="value">Highest value</option>
          <option value="daily">Highest daily consumption</option>
          <option value="product">Product name A-Z</option>
        </select>
      </div>

      <div style={styles.card}>
        <h3 style={styles.cardTitle}>📅 Month comparison</h3>

        <div style={styles.viewModeBox}>
          <button
            style={{
              ...styles.viewModeButton,
              ...(yearlyRegionalReportMonths.length === 0
                ? styles.viewModeButtonActive
                : {}),
            }}
            onClick={() => setYearlyRegionalReportMonths([])}
          >
            All months
          </button>

          {yearlyRegionalMonthOptions.map((month) => {
            const active =
              yearlyRegionalReportMonths.length === 0 ||
              yearlyRegionalReportMonths.includes(month.monthKey);

            return (
              <button
                key={month.monthKey}
                style={{
                  ...styles.viewModeButton,
                  ...(active ? styles.viewModeButtonActive : {}),
                }}
                onClick={() => toggleYearlyRegionalReportMonth(month.monthKey)}
              >
                {month.monthName}
              </button>
            );
          })}
        </div>

        <div style={styles.infoBox}>
          <div>
            📄 File:{" "}
            <strong>{yearlyRegionalFileName || "Not loaded"}</strong>
          </div>

          <div>
            📦 Product rows shown:{" "}
            <strong>{yearlyRegionalConsumptionReportRows.length}</strong>
          </div>

          <div>
            🧭 Region filter:{" "}
            <strong>
              {yearlyRegionalReportRegion === YEARLY_REGION_ALL
                ? "All regions"
                : yearlyRegionalReportRegion}
            </strong>
          </div>

          <div>
            🚢 Ship filter:{" "}
            <strong>
              {yearlyRegionalReportShip === YEARLY_REPORT_ALL_SHIPS
                ? "All ships"
                : yearlyRegionalReportShip}
            </strong>
          </div>

          <div>
            📅 Months selected:{" "}
            <strong>
              {yearlyRegionalReportMonths.length === 0
                ? "All months"
                : yearlyRegionalMonthOptions
                    .filter((month) =>
                      yearlyRegionalReportMonths.includes(month.monthKey)
                    )
                    .map((month) => month.monthName)
                    .join(", ")}
            </strong>
          </div>
        </div>
      </div>
    </div>

    {yearlyRegionalConsumption &&
      yearlyRegionalConsumptionReportRows.length === 0 && (
        <p style={styles.emptyText}>
          No yearly regional rows match the current filters.
        </p>
      )}

    <div style={styles.costReportLineList}>
      {yearlyRegionalConsumptionReportRows.map((item, index) => {
        const activeMonthSet = new Set(yearlyRegionalActiveMonthKeys);
        const visibleMonths = yearlyRegionalMonthOptions.filter((month) =>
          activeMonthSet.has(month.monthKey)
        );

        return (
          <div
            key={`${item.productCode || item.productKey || item.productName}-${index}`}
            style={styles.costReportLine}
          >
            <div style={styles.costLineMain}>
              <div style={styles.costLineProduct}>
                {item.productName || "Unnamed product"}
              </div>

              <div style={styles.costLineMeta}>
                {item.productCode ? "Code: " + item.productCode + " • " : ""}
                U/M: {item.unitMeasure || "N/A"}
              </div>

              <div style={styles.costLineMeta}>
                {item.categoryName || "No category"}
                {item.subCategoryName ? " • " + item.subCategoryName : ""}
              </div>

              <div style={styles.statusGood}>
                Daily: {formatRegionalQty(item.avgDailyQty)} • Days:{" "}
                {formatQty(item.totalDays)} • Blocks: {item.blocks}
              </div>
            </div>

            <div style={styles.costLineTotals}>
              <span>Total Qty</span>
              <strong>{formatRegionalQty(item.totalQty)}</strong>
              <span>{formatMoney(item.totalValue)}</span>
              {item.avgPrice > 0 && (
                <small>{formatMoney(item.avgPrice)} / unit</small>
              )}
            </div>

            <div style={styles.costLineVenues}>
              {visibleMonths.map((month) => {
                const monthData = item.months?.[month.monthKey] || {};
                const hasMonthData =
                  Number(monthData.qty || 0) !== 0 ||
                  Number(monthData.value || 0) !== 0;

                return (
                  <div key={month.monthKey} style={styles.costLineVenue}>
                    <div style={styles.costLineVenueTitle}>{month.monthName}</div>

                    {hasMonthData ? (
                      <>
                        <div style={styles.costLineShipChips}>
                          <div style={styles.costLineShipChip}>
                            <span style={styles.costLineShipName}>Qty</span>
                            <strong>{formatRegionalQty(monthData.qty)}</strong>
                          </div>

                          <div style={styles.costLineShipChip}>
                            <span style={styles.costLineShipName}>Value</span>
                            <strong>{formatMoney(monthData.value)}</strong>
                          </div>

                          <div style={styles.costLineShipChip}>
                            <span style={styles.costLineShipName}>Daily</span>
                            <strong>
                              {formatRegionalQty(monthData.avgDailyQty)}
                            </strong>
                          </div>

                          <div style={styles.costLineShipChip}>
                            <span style={styles.costLineShipName}>Days</span>
                            <strong>{formatQty(monthData.days)}</strong>
                          </div>
                        </div>

                        <div style={styles.costLinePriceNote}>
                          Ships: {(monthData.ships || []).join(", ") || "N/A"}
                        </div>

                        <div style={styles.recipeMeta}>
                          Regions: {(monthData.regions || []).join(", ") || "N/A"}
                        </div>
                      </>
                    ) : (
                      <div style={styles.statusNeutral}>No consumption</div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  </section>
)}
      {productReportView === "main" && (
      <section style={styles.card}>
        <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 18 }}>
          <div>
            <h2 style={styles.productTitle}>💰 Main Consumption & Cost Report</h2>
            <p style={{ ...styles.emptyText, margin: 0 }}>
              Automatically generated from the consumption file. Lowest unit price is highlighted when ship prices differ.
            </p>
          </div>

          <div style={styles.headerActions}>
            <button style={styles.backButton} onClick={printMainConsumptionCostReport}>
              🖨️ Print
            </button>
            <button style={styles.primaryButton} onClick={exportMainConsumptionCostReportToExcel}>
              📥 Export Excel
            </button>
          </div>
        </div>

        <div style={styles.infoBox}>
          <div>📦 Products in report: <strong>{filteredProductCostReportRows.length}</strong> / {productCostReportRows.length}</div>
          <div>🚢 View: <strong>{viewMode === "single" ? userShip : "All Ships"}</strong></div>
          <div>📘 Source columns: C Venue, I/L/O/R Quantity, J/M/P/S Total Cost, H/K/N/Q Unit Price.</div>
        </div>

        <input
          placeholder="Search product, code, or venue..."
          value={productCostReportSearch}
          onChange={(e) => setProductCostReportSearch(e.target.value)}
          style={{ ...styles.searchInput, marginTop: 14 }}
        />

        {consumptionRows.length === 0 && (
          <p style={styles.emptyText}>Upload the consumption file to generate this report.</p>
        )}

        {consumptionRows.length > 0 && filteredProductCostReportRows.length === 0 && (
          <p style={styles.emptyText}>No products found for this search.</p>
        )}

        <div style={styles.costReportLineList}>
          {filteredProductCostReportRows.map((item) => {
  const itemVenues = Array.isArray(item.venues) ? item.venues : [];

  return (
            <div key={item.productKey} style={styles.costReportLine}>
              <div style={styles.costLineMain}>
  <div style={styles.costLineProduct}>{item.product}</div>

  <div style={styles.costLineMeta}>
    {item.code ? "Code: " + item.code + " • " : ""}
    {itemVenues.length} venue{itemVenues.length === 1 ? "" : "s"}
  </div>

  {item.regionalHasData ? (
    <div style={styles.statusGood}>
      Region:{" "}
      {item.regionalRegion === YEARLY_REGION_ALL
        ? "All regions"
        : item.regionalRegion}{" "}
      • Daily: {formatRegionalQty(item.regionalAvgDailyQty)} • Suggested Par:{" "}
      {formatRegionalQty(item.regionalSuggestedParLevel)}
    </div>
  ) : yearlyRegionalConsumption ? (
    <div style={styles.statusNeutral}>
      No regional yearly match
    </div>
  ) : null}
</div>

              <div style={styles.costLineTotals}>
                <span>Total</span>
                <strong>{formatQty(item.visibleTotalQty)}</strong>
                <span>{formatMoney(item.visibleTotalCost)}</span>
              </div>

              <div style={styles.costLineVenues}>
                {itemVenues.map((venue) => (
                  <div key={venue.venueKey} style={styles.costLineVenue}>
                    <div style={styles.costLineVenueTitle}>{venue.location}</div>

                    <div style={styles.costLineShipChips}>
                      {visibleShips.map((ship) => {
                        const shipData = venue.ships[ship] || { qty: 0, cost: 0, unitPrice: 0, isLowestUnitPrice: false };
                        const hasData = Number(shipData.qty || 0) !== 0 || Number(shipData.cost || 0) !== 0;

                        if (!hasData) return null;

                        return (
                          <div
                            key={ship}
                            style={{
                              ...styles.costLineShipChip,
                              ...(shipData.isLowestUnitPrice ? styles.costLineShipLowest : {}),
                            }}
                          >
                            <span style={styles.costLineShipName}>{ship}</span>
                            <strong>{formatQty(shipData.qty)}</strong>
                            <span>{formatMoney(shipData.cost)}</span>
                            {shipData.unitPrice > 0 && (
                              <small>{formatMoney(shipData.unitPrice)} / unit</small>
                            )}
                          </div>
                        );
                      })}
                    </div>

                    {venue.hasPriceDifference && (
                      <div style={styles.costLinePriceNote}>Lowest unit price highlighted.</div>
                    )}
                  </div>
                ))}
              </div>
                        </div>
          );
        })}
        </div>

      </section>
      )}

      {productReportView === "reports" && (
      <section style={styles.card}>
        <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: showProductMissingReport ? 18 : 0 }}>
          <div>
            <h2 style={styles.productTitle}>📋 Generate Report</h2>
            <p style={{ ...styles.emptyText, margin: 0 }}>
              Generate the Top 50 items not in use by the location where they should be used.
            </p>
          </div>

          <div style={styles.headerActions}>
            <button
              style={styles.backButton}
              onClick={toggleProductMissingReport}
              disabled={productMissingReportLoading}
            >
              {productMissingReportLoading ? "Preparing..." : showProductMissingReport ? "Hide Report" : "Open Report"}
            </button>

            {showProductMissingReport && (
              <button
                style={styles.backButton}
                onClick={refreshProductMissingReport}
                disabled={productMissingReportLoading}
              >
                🔄 Refresh Report
              </button>
            )}

            <button
              style={styles.backButton}
              onClick={printTopNotInUseByLocationReport}
              disabled={productMissingReportLoading}
            >
              🖨️ Print
            </button>

            <button
              style={styles.primaryButton}
              onClick={exportTopNotInUseByLocationReportToExcel}
              disabled={productMissingReportLoading}
            >
              📥 Export Excel
            </button>
          </div>
        </div>

        {showProductMissingReport && (
          <>
            <div style={styles.infoBox}>
              <div>📋 Report: <strong>Top 50 Not In Use By Location</strong></div>
              <div>🚢 View: <strong>{viewMode === "single" ? userShip : "All Ships"}</strong></div>
              <div>🔎 Rows found: <strong>{topNotInUseReport.length}</strong></div>
              {productMissingReportLoading && <div>Preparing report, please wait...</div>}
              {productMissingReportMessage && <div style={{ color: topNotInUseReport.length ? "#555" : "#8a5a00" }}>{productMissingReportMessage}</div>}
              <div style={{ color: "#b00020" }}>
                Shows products that are expected by recipe/location or template charge location, but usage is 0 for one or more visible ship(s).
              </div>
            </div>

            {!productMissingReportLoading && topNotInUseReport.length === 0 && (
              <p style={styles.emptyText}>
                No not-in-use records found. Upload consumption, recipe/location, and template files, then open the report again.
              </p>
            )}

            <div style={styles.equipmentGrid}>
              {topNotInUseReport.map((item, index) => (
                <div key={`${item.product}-${item.venueKey}-${index}`} style={{ ...styles.equipmentCard, ...styles.orderWarningCard }}>
                  <div style={styles.recipeMeta}>#{index + 1}</div>
                  <div style={styles.recipeName}>{item.product}</div>
                  <div style={styles.recipeMeta}>Location: {item.location}</div>
                  <div style={styles.recipeMeta}>Source: {item.source}</div>

                  {item.templateMatches.length > 0 && (
                    <div style={styles.templateFound}>Template/Menu: {item.templateMatches.join(", ")}</div>
                  )}

                  {item.missingFromTemplate && (
                    <div style={styles.templateWarningText}>Also missing from matching template.</div>
                  )}

                  <div style={styles.shipGrid}>
                    {visibleShips.map((ship) => {
                      const isMissing = item.missingShips.includes(ship);

                      return (
                        <div
                          key={ship}
                          style={{ ...styles.shipBox, ...(ship === userShip ? styles.shipBoxActive : {}), ...(isMissing ? styles.shipBoxMissing : {}) }}
                        >
                          <span style={styles.shipName}>{ship}</span>
                          <strong style={styles.shipQty}>{formatQty(item.ships?.[ship])}</strong>
                        </div>
                      );
                    })}
                  </div>

                  <div style={styles.statusBad}>Missing: {item.missingShips.join(", ")}</div>

                  <button
                    style={styles.backButton}
                    onClick={() => {
                      setProductReportView("consumption");
                      setSelectedProduct(item.product);
                      setSelectedRecipe(null);
                    }}
                  >
                    Open Product
                  </button>
                </div>
              ))}
            </div>
          </>
        )}
      </section>
      )}

      {productReportView === "consumption" && selectedProduct && (
        <section style={styles.card}>
          <div style={{ ...styles.header, boxShadow: "none", padding: 0, marginBottom: 18 }}>
            <div>
              <h2 style={styles.productTitle}>📦 {selectedProduct}</h2>
              <p style={{ ...styles.emptyText, margin: 0 }}>
                Consumption vs locations and template for the selected product.
              </p>
            </div>

            <div style={styles.headerActions}>
              <button style={styles.backButton} onClick={printSelectedProductConsumptionReport}>
                🖨️ Print
              </button>
              <button style={styles.primaryButton} onClick={exportSelectedProductConsumptionReportToExcel}>
                📥 Export Excel
              </button>
            </div>
          </div>

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
                    {venueItem.requiredFromTemplate && <span style={styles.chargeBadge}>Template Charge</span>}
                    {venueItem.missingFromTemplate && <span style={styles.templateBadge}>Missing Template</span>}
                    {venueItem.missingShips.length > 0 && <span style={styles.missingBadge}>Missing: {venueItem.missingShips.join(", ")}</span>}
                  </span>
                </h4>

                {venueItem.requiredFromTemplate && (
                  <div style={styles.templateFound}>
                    Template charge location: {venueItem.templateMatches.join(", ") || "Template"}
                  </div>
                )}

                {venueItem.missingFromTemplate && (
                  <div style={styles.templateWarningText}>
                    Product is used in the recipe/location file for this venue, but it is not found in this venue template.
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
                  <div style={styles.warningSmall}>
                    {venueItem.requiredFromTemplate
                      ? "Product appears in the template charge location, but usage is 0 for highlighted ship(s)."
                      : "Product appears in recipe/location file for this venue, but usage is 0 for highlighted ship(s)."}
                  </div>
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
  page: {
  minHeight: "100vh",
  padding: "clamp(10px, 3vw, 24px)",
  background: "#f5f5f5",
  fontFamily: "Arial, sans-serif",
  color: "#111",
  boxSizing: "border-box",
},
  welcomePage: { minHeight: "100vh", padding: 24, background: "radial-gradient(circle at top left, #ffffff 0%, #f7f7f7 32%, #ececec 100%)", fontFamily: "Arial, sans-serif", color: "#111", display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden" },
  welcomeHero: { width: "100%", maxWidth: 980, margin: "0 auto" },
  welcomeGlowCard: { position: "relative", padding: "34px 28px", background: "rgba(255,255,255,0.94)", borderRadius: 28, border: "1px solid rgba(0,0,0,0.08)", boxShadow: "0 18px 50px rgba(0,0,0,0.18)", display: "grid", gap: 18, textAlign: "center", overflow: "hidden", animation: "vvGlow 4s ease-in-out infinite" },
  welcomeLogo: { height: 120, maxWidth: "100%", objectFit: "contain", margin: "0 auto 2px" },
  runningLineWrapper: { width: "100%", overflow: "hidden", background: "#fff", borderRadius: 999, padding: "10px 0 12px", borderTop: "1px solid rgba(176,0,32,0.10)", borderBottom: "1px solid rgba(176,0,32,0.10)" },
  runningLineTrack: { display: "flex", width: "max-content", whiteSpace: "nowrap", animation: "vvMarquee 30s linear infinite" },
  runningLineText: { color: "#e00000", fontFamily: "Brush Script MT, Segoe Script, Lucida Handwriting, Apple Chancery, cursive", fontSize: 44, fontWeight: 400, letterSpacing: 1.1, paddingRight: 70, textShadow: "0 1px 0 rgba(176,0,32,0.10)", lineHeight: 1.15 },
  welcomeTitle: { margin: "8px 0 0", fontSize: 34, lineHeight: 1.1, fontWeight: 900 },
  welcomeSubtitle: { margin: 0, color: "#555", fontSize: 16 },
  ahoyStartBox: { position: "relative", width: "100%", minHeight: 150, display: "flex", alignItems: "center", justifyContent: "center", background: "transparent", border: 0, boxShadow: "none", borderRadius: 0, overflow: "visible", margin: "4px 0 0" },
  ahoyStartButton: { padding: "8px 22px", border: 0, background: "transparent", color: "#e00000", fontFamily: "Brush Script MT, Segoe Script, Lucida Handwriting, Apple Chancery, cursive", fontSize: "clamp(44px, 7vw, 82px)", fontWeight: 400, cursor: "pointer", lineHeight: 0.9, letterSpacing: 1.2, textShadow: "0 4px 0 rgba(255,255,255,0.95), 0 0 24px rgba(255,255,255,0.95), 0 4px 12px rgba(176,0,32,0.18)" },
  shipPhotoStartBox: { position: "relative", width: "100%", minHeight: 330, borderRadius: 26, overflow: "hidden", border: "1px solid rgba(0,0,0,0.10)", boxShadow: "0 16px 42px rgba(0,0,0,0.18)", background: "#111" },
  shipPhotoStartImage: { width: "100%", height: 350, objectFit: "cover", display: "block", filter: "saturate(1.12) contrast(1.04)" },
  shipPhotoOverlay: { position: "absolute", inset: 0, padding: 22, display: "flex", flexDirection: "column", justifyContent: "space-between", background: "linear-gradient(180deg, rgba(0,0,0,0.16) 0%, rgba(0,0,0,0.05) 43%, rgba(0,0,0,0.62) 100%)" },
  shipPhotoOverlayBottom: { position: "absolute", left: 32, bottom: 28, display: "flex", justifyContent: "flex-start", alignItems: "center", padding: 0 },
  welcomeStartButton: { padding: 0, border: 0, background: "transparent", color: "#e00000", fontFamily: "Brush Script MT, Segoe Script, Lucida Handwriting, Apple Chancery, cursive", fontSize: 78, fontWeight: 400, cursor: "pointer", lineHeight: 1, letterSpacing: 1.1, textShadow: "0 3px 0 rgba(255,255,255,0.95), 0 0 18px rgba(255,255,255,0.95), 0 2px 8px rgba(0,0,0,0.35)" },
  shipPhotoTitle: { alignSelf: "center", padding: "11px 22px", borderRadius: 999, background: "rgba(255,255,255,0.92)", color: "#111", fontSize: 28, fontWeight: 900, boxShadow: "0 8px 20px rgba(0,0,0,0.20)", letterSpacing: 0.2 },
  shipPhotoButtons: { display: "grid", gridTemplateColumns: "repeat(4, minmax(90px, 1fr))", gap: 12, width: "100%" },
  shipPhotoButton: { minHeight: 82, border: "1px solid rgba(255,255,255,0.75)", borderRadius: 18, background: "rgba(17,17,17,0.88)", color: "#fff", cursor: "pointer", display: "grid", alignContent: "center", justifyItems: "center", gap: 5, boxShadow: "0 8px 24px rgba(0,0,0,0.28)", backdropFilter: "blur(4px)", fontWeight: "bold" },
  shipPhotoButtonShip: { fontSize: 22, fontWeight: 900, letterSpacing: 0.8 },
  shipPhotoButtonHint: { fontSize: 12, opacity: 0.82, textTransform: "uppercase", letterSpacing: 1 },
  shipStartGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 14, marginTop: 8 },
  shipStartButton: { minHeight: 142, border: "1px solid #ddd", borderRadius: 22, background: "linear-gradient(180deg, #ffffff 0%, #f7f7f7 100%)", cursor: "pointer", display: "grid", alignContent: "center", justifyItems: "center", gap: 8, boxShadow: "0 8px 22px rgba(0,0,0,0.08)", transition: "transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease", color: "#111" },
  shipStartIcon: { fontSize: 34 },
  shipStartCode: { fontSize: 32, fontWeight: 900, letterSpacing: 1.2 },
  shipStartHint: { fontSize: 13, color: "#666", fontWeight: "bold", textTransform: "uppercase", letterSpacing: 0.8 },
  welcomeFooterNote: { color: "#777", fontSize: 13, fontWeight: "bold", marginTop: 4 },
  loginCard: { maxWidth: 460, margin: "80px auto", padding: 28, background: "#fff", borderRadius: 16, boxShadow: "0 10px 30px rgba(0,0,0,0.08)", display: "grid", gap: 14 },
  logo: { height: 70, objectFit: "contain", marginBottom: 8 },
  header: {
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  gap: 10,
  flexWrap: "wrap",
  padding: "clamp(10px, 3vw, 18px)",
  background: "#fff",
  borderRadius: 16,
  boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
  marginBottom: 16,
},
  headerLogo: {
  height: "clamp(38px, 10vw, 54px)",
  objectFit: "contain",
},
  headerActions: { display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" },
  backButton: { padding: "10px 14px", borderRadius: 999, border: "1px solid #ccc", background: "#fff", cursor: "pointer", fontWeight: "bold" },
  checkboxRow: { display: "flex", alignItems: "center", gap: 10, fontWeight: "bold", marginTop: 4 },
  emailError: { color: "#b00020", background: "#fff0f0", border: "1px solid #f1b8b8", borderRadius: 10, padding: 10, fontWeight: "bold" },
  inlineLinkButton: { border: 0, background: "transparent", color: "#0057b8", cursor: "pointer", fontWeight: "bold", padding: 0, textAlign: "left" },
  deleteButton: { padding: "10px 14px", borderRadius: 999, border: "1px solid #b00020", background: "#b00020", color: "#fff", cursor: "pointer", fontWeight: "bold" },
  title: { margin: 0, fontSize: 28 },
  subtitle: { margin: 0, color: "#666" },
  label: { fontWeight: "bold", marginTop: 8 },
  select: { padding: 10, borderRadius: 8, border: "1px solid #ccc", width: "100%" },
  primaryButton: { marginTop: 10, padding: 12, borderRadius: 10, border: 0, background: "#111", color: "#fff", fontWeight: "bold", cursor: "pointer" },
  shipBadge: { padding: "10px 14px", borderRadius: 999, background: "#111", color: "#fff", fontWeight: "bold" },
  moduleGrid: {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 145px), 1fr))",
  gap: "clamp(8px, 2.5vw, 16px)",
},
  moduleCard: {
  border: "1px solid #ddd",
  background: "#fafafa",
  borderRadius: 16,
  padding: "clamp(12px, 3vw, 20px)",
  minHeight: 118,
  cursor: "pointer",
  textAlign: "center",
  display: "grid",
  alignContent: "center",
  justifyItems: "center",
  gap: 7,
  boxShadow: "0 4px 12px rgba(0,0,0,0.04)",
  fontSize: "clamp(12px, 3.2vw, 14px)",
  lineHeight: 1.25,
},
  moduleIcon: {
  fontSize: "clamp(26px, 8vw, 34px)",
  lineHeight: 1,
},
  viewModeBox: {
  display: "flex",
  gap: 8,
  marginBottom: 16,
  flexWrap: "wrap",
  overflowX: "auto",
  paddingBottom: 2,
},
  viewModeButton: {
  padding: "9px 12px",
  borderRadius: 999,
  border: "1px solid #ccc",
  background: "#fff",
  cursor: "pointer",
  fontWeight: "bold",
  fontSize: "clamp(12px, 3vw, 14px)",
  whiteSpace: "nowrap",
},
  viewModeButtonActive: { background: "#111", color: "#fff", borderColor: "#111" },
  grid: {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 320px), 1fr))",
  gap: "clamp(12px, 3vw, 20px)",
  marginBottom: 16,
},
  formRow: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  card: {
  background: "#fff",
  borderRadius: 16,
  padding: "clamp(12px, 3vw, 20px)",
  boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
  minWidth: 0,
},
  cardTitle: { marginTop: 0 },
  fileInput: { display: "block", margin: "8px 0 16px" },
  message: { color: "#555", fontSize: 14 },
  infoBox: { marginTop: 12, padding: 12, borderRadius: 12, background: "#f2f2f2", display: "grid", gap: 6 },
  reportFilterBox: { marginBottom: 16, padding: 14, borderRadius: 14, background: "#f7fbff", border: "1px solid #cfe4ff", display: "grid", gap: 8 },
  reportModeGrid: { display: "grid", gap: 10 },
  reportModeButton: { border: "1px solid #ddd", background: "#fafafa", borderRadius: 14, padding: 14, cursor: "pointer", textAlign: "left", display: "grid", gap: 4, color: "#111" },
  reportModeButtonActive: { background: "#111", color: "#fff", borderColor: "#111" },
  searchInput: {
  width: "100%",
  padding: "11px 12px",
  borderRadius: 10,
  border: "1px solid #ccc",
  marginBottom: 10,
  boxSizing: "border-box",
  fontSize: 16,
},
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
  chargeBadge: { fontSize: 12, color: "#111", background: "#e8f5e9", borderRadius: 999, padding: "4px 8px", border: "1px solid #2e7d32" },
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
  fmlCompactGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 8 },
  fmlCompactCard: { border: "2px solid #0057b8", borderRadius: 10, padding: 8, background: "#eef5ff", display: "grid", gap: 5, textAlign: "left", fontSize: 11, alignContent: "start", minWidth: 0 },
  fmlCompactName: { fontWeight: "bold", fontSize: 13, lineHeight: 1.12, overflowWrap: "anywhere" },
  fmlCompactMeta: { color: "#444", fontSize: 11, lineHeight: 1.18, overflowWrap: "anywhere" },
  fmlCompactBadge: { padding: "5px 6px", borderRadius: 8, background: "#0057b8", color: "#fff", fontWeight: "bold", textAlign: "center", fontSize: 11, lineHeight: 1.1 },
  fmlCompactReason: { padding: 5, borderRadius: 8, background: "#fff4d6", color: "#8a5a00", fontWeight: "bold", textAlign: "center", fontSize: 10, lineHeight: 1.1 },
  nextOrderGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(185px, 1fr))", gap: 10 },
  nextOrderCard: { border: "1px solid #ddd", borderRadius: 12, padding: 10, background: "#fafafa", display: "grid", gap: 6, textAlign: "left", fontSize: 12 },
  nextOrderCardBlue: { border: "2px solid #0057b8", background: "#eef5ff" },
  nextOrderCardRed: { border: "2px solid #b00020", background: "#fff0f0" },
  nextOrderTopLine: { display: "flex", justifyContent: "space-between", gap: 8, color: "#666", fontSize: 11, fontWeight: "bold" },
  nextOrderName: { fontWeight: "bold", fontSize: 14, lineHeight: 1.15 },
  nextOrderMeta: { color: "#555", fontSize: 12, lineHeight: 1.2 },
  nextOrderMiniGrid: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 6 },
  nextOrderMiniBox: { background: "#fff", border: "1px solid #ddd", borderRadius: 8, padding: 6, display: "grid", gap: 2, textAlign: "center", minWidth: 0 },
  nextOrderMiniBoxNegative: { background: "#b00020", border: "1px solid #b00020", color: "#fff", fontWeight: "bold" },
  nextOrderWarning: { padding: 6, borderRadius: 8, background: "#fff4d6", color: "#8a5a00", fontWeight: "bold", textAlign: "center", fontSize: 12 },
  nextOrderSuggestedBlue: { padding: 7, borderRadius: 8, background: "#0057b8", color: "#fff", fontWeight: "bold", textAlign: "center", fontSize: 13 },
  nextOrderSuggestedNeutral: { padding: 7, borderRadius: 8, background: "#f2f2f2", color: "#555", fontWeight: "bold", textAlign: "center", fontSize: 13 },
  nextOrderStatusBlue: { padding: 6, borderRadius: 8, background: "#0057b8", color: "#fff", fontWeight: "bold", textAlign: "center", fontSize: 12 },
  nextOrderStatusRed: { padding: 6, borderRadius: 8, background: "#b00020", color: "#fff", fontWeight: "bold", textAlign: "center", fontSize: 12 },
  nextOrderStatusNeutral: { padding: 6, borderRadius: 8, background: "#f2f2f2", color: "#555", fontWeight: "bold", textAlign: "center", fontSize: 12 },
  costReportGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(360px, 1fr))", gap: 14, marginTop: 14, alignItems: "stretch" },
  costReportCard: { border: "1px solid #ddd", borderRadius: 14, padding: 12, background: "#fafafa", display: "grid", gap: 10, textAlign: "left", alignContent: "start" },
  costReportHeader: { display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start", minHeight: 58 },
  costReportTotals: { minWidth: 86, padding: 8, borderRadius: 10, background: "#111", color: "#fff", display: "grid", gap: 2, textAlign: "center", fontSize: 12 },
  costVenueBlock: { border: "1px solid #e1e1e1", borderRadius: 12, padding: 10, background: "#fff", display: "grid", gap: 8, alignContent: "start" },
  costVenueTitle: { fontWeight: "bold", fontSize: 14, minHeight: 18 },
  costShipGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(88px, 1fr))", gap: 8, alignItems: "stretch" },
  costShipBox: { border: "1px solid #ddd", borderRadius: 10, padding: 8, background: "#fff", display: "grid", gridTemplateRows: "16px 22px 22px 18px 34px", gap: 4, textAlign: "center", fontSize: 12, minHeight: 136, boxSizing: "border-box", alignContent: "center" },
  costShipLowestPrice: { border: "2px solid #2e7d32", background: "#e8f5e9", color: "#2e7d32", fontWeight: "bold" },
  costShipEmpty: { opacity: 0.5 },
  costUnitPricePlaceholder: { visibility: "hidden" },
  lowestPriceBadge: { marginTop: 4, padding: "4px 6px", borderRadius: 999, background: "#2e7d32", color: "#fff", fontSize: 10, fontWeight: "bold", minHeight: 24, display: "flex", alignItems: "center", justifyContent: "center", lineHeight: 1.1 },
  lowestPriceBadgeSpacer: { minHeight: 32 },
  priceDifferenceNote: { color: "#2e7d32", fontSize: 12, fontWeight: "bold", minHeight: 28 },
  costReportLineList: { display: "grid", gap: 8, marginTop: 14 },
  costReportLine: { border: "1px solid #ddd", borderRadius: 14, padding: 10, background: "#fafafa", display: "grid", gridTemplateColumns: "minmax(220px, 1.1fr) 110px minmax(420px, 2.6fr)", gap: 10, alignItems: "stretch", textAlign: "left" },
  costLineMain: { display: "grid", gap: 5, alignContent: "center", minWidth: 0 },
  costLineProduct: { fontWeight: "bold", fontSize: 14, lineHeight: 1.15, overflowWrap: "anywhere" },
  costLineMeta: { color: "#555", fontSize: 12, lineHeight: 1.2 },
  costLineTotals: { padding: 8, borderRadius: 10, background: "#111", color: "#fff", display: "grid", gap: 2, textAlign: "center", fontSize: 12, alignContent: "center", minHeight: 62 },
  costLineVenues: { display: "flex", gap: 8, overflowX: "auto", alignItems: "stretch", paddingBottom: 2 },
  costLineVenue: { minWidth: 280, border: "1px solid #e1e1e1", borderRadius: 12, padding: 8, background: "#fff", display: "grid", gap: 6, alignContent: "start" },
  costLineVenueTitle: { fontWeight: "bold", fontSize: 12, lineHeight: 1.1 },
  costLineShipChips: { display: "flex", flexWrap: "wrap", gap: 5, alignItems: "stretch" },
  costLineShipChip: { border: "1px solid #ddd", borderRadius: 8, padding: "5px 7px", background: "#fff", display: "grid", gap: 2, textAlign: "center", fontSize: 11, minWidth: 68, alignContent: "center" },
  costLineShipLowest: { border: "2px solid #2e7d32", background: "#e8f5e9", color: "#2e7d32", fontWeight: "bold" },
  costLineShipName: { color: "#555", fontSize: 10, fontWeight: "bold" },
  costLinePriceNote: { color: "#2e7d32", fontSize: 11, fontWeight: "bold", lineHeight: 1.1 },
  equipmentCategory: { marginBottom: 24 },
  equipmentGrid: {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 150px), 1fr))",
  gap: "clamp(8px, 2.5vw, 14px)",
},
  equipmentCard: {
  border: "1px solid #ddd",
  borderRadius: 14,
  padding: "clamp(10px, 2.7vw, 14px)",
  background: "#fafafa",
  display: "grid",
  gap: 7,
  cursor: "pointer",
  textAlign: "left",
  minWidth: 0,
  fontSize: "clamp(12px, 3.2vw, 14px)",
},
    inventoryItemCard: {
    border: "1px solid #ddd",
    borderRadius: 18,
    padding: 14,
    background: "#fff",
    display: "grid",
    gap: 10,
    cursor: "pointer",
    textAlign: "left",
    boxShadow: "0 6px 18px rgba(0,0,0,0.06)",
  },

  inventoryImageFrame: {
    width: "100%",
    height: 170,
    borderRadius: 14,
    overflow: "hidden",
    background: "#f2f2f2",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },

  inventoryCardImage: {
    width: "100%",
    height: "100%",
    objectFit: "cover",
    display: "block",
  },

  inventoryNoImage: {
    height: 170,
    borderRadius: 14,
    background: "#f2f2f2",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#777",
    fontWeight: "bold",
  },
  countedCard: { border: "2px solid #2e7d32", background: "#f0fff4" },
  equipmentImage: { width: "100%", height: 150, objectFit: "cover", borderRadius: 10, background: "#eee" },
  equipmentNoImage: { height: 150, borderRadius: 10, background: "#eee", display: "flex", alignItems: "center", justifyContent: "center", color: "#777" },
  modalBackdrop: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999, padding: 20 },
  modalCard: { background: "#fff", borderRadius: 18, padding: 22, maxWidth: 760, width: "100%", maxHeight: "90vh", overflowY: "auto", position: "relative", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" },
  modalImage: { width: "100%", maxHeight: "65vh", objectFit: "contain", borderRadius: 14, background: "#f2f2f2" },
  modalPictureFrame: {
  width: "100%",
  minHeight: 320,
  borderRadius: 16,
  background: "#f2f2f2",
  border: "1px solid #e2e2e2",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  overflow: "hidden",
},

modalPreviewImage: {
  width: "100%",
  maxHeight: "58vh",
  objectFit: "contain",
  display: "block",
  background: "#f2f2f2",
},

modalNoImage: {
  width: "100%",
  minHeight: 320,
  borderRadius: 16,
  background: "#f2f2f2",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  color: "#777",
  fontWeight: "bold",
},
  closeButton: { position: "absolute", top: 12, right: 12, border: 0, background: "#111", color: "#fff", borderRadius: 999, width: 34, height: 34, cursor: "pointer", fontWeight: "bold" },
  imageButton: { display: "block", width: "100%", marginTop: 8, padding: 10, borderRadius: 10, border: 0, background: "#111", color: "#fff", textAlign: "center", cursor: "pointer", fontWeight: "bold" },
  imageLink: { display: "none", marginTop: 8, padding: 10, borderRadius: 10, background: "#111", color: "#fff", textAlign: "center", textDecoration: "none", fontWeight: "bold" },
  orderWarningCard: { border: "2px solid #b00020", background: "#fff0f0" },
  orderNeededCard: { border: "2px solid #0057b8", background: "#eef5ff" },
  overstockCard: { border: "2px solid #b00020", background: "#fff0f0" },
  overstockWarning: { marginTop: 8, padding: 8, borderRadius: 10, background: "#b00020", color: "#fff", fontWeight: "bold", textAlign: "center" },
  zeroCountCard: { border: "2px solid #8a5a00", background: "#fff8e1" },
  suggestedOrderBad: { marginTop: 8, padding: 8, borderRadius: 10, background: "#0057b8", color: "#fff", fontWeight: "bold", textAlign: "center" },
  suggestedOrderBlue: { marginTop: 8, padding: 8, borderRadius: 10, background: "#0057b8", color: "#fff", fontWeight: "bold", textAlign: "center" },
  suggestedOrderGood: { marginTop: 8, padding: 8, borderRadius: 10, background: "#e8f5e9", color: "#2e7d32", fontWeight: "bold", textAlign: "center" },
  statusGood: { marginTop: 8, padding: 8, borderRadius: 10, background: "#e8f5e9", color: "#2e7d32", fontWeight: "bold", textAlign: "center" },
  statusWarning: { marginTop: 8, padding: 8, borderRadius: 10, background: "#fff4d6", color: "#8a5a00", fontWeight: "bold", textAlign: "center" },
  statusNeutral: { marginTop: 8, padding: 8, borderRadius: 10, background: "#f2f2f2", color: "#555", fontWeight: "bold", textAlign: "center" },
  statusBad: { marginTop: 8, padding: 8, borderRadius: 10, background: "#b00020", color: "#fff", fontWeight: "bold", textAlign: "center" },
              stationStatusGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
    gap: 10,
    marginTop: 14,
  },

  stationStatusCard: {
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 12,
    background: "#f2f2f2",
    display: "grid",
    gap: 5,
    fontSize: 13,
  },

  stationStatusStarted: {
    border: "2px solid #8a5a00",
    background: "#fff4d6",
  },

  stationStatusSubmitted: {
    border: "2px solid #2e7d32",
    background: "#e8f5e9",
  },

  finishBar: {
    marginTop: 18,
    padding: 16,
    borderRadius: 14,
    background: "#f7f7f7",
    border: "1px solid #ddd",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 12,
    flexWrap: "wrap",
  },
  timelineScroll: { width: "100%", overflowX: "auto", border: "1px solid #ddd", borderRadius: 14, background: "#fff" },
  timelineGrid: { display: "grid", minWidth: 1280, alignItems: "stretch" },
  timelineHeaderCell: { position: "sticky", top: 0, zIndex: 2, padding: "10px 8px", background: "#111", color: "#fff", fontSize: 12, fontWeight: "bold", borderRight: "1px solid #333", borderBottom: "1px solid #333", textAlign: "center" },
  timelineFixedCell: { padding: "9px 8px", borderRight: "1px solid #ddd", borderBottom: "1px solid #eee", background: "#fafafa", fontSize: 12, display: "flex", alignItems: "center" },
  timelineNameCell: { padding: "9px 8px", borderRight: "1px solid #ddd", borderBottom: "1px solid #eee", background: "#fafafa", fontSize: 12, display: "grid", gap: 4 },
  timelineCell: { minHeight: 38, borderRight: "1px solid #e5e5e5", borderBottom: "1px solid #eee", background: "#fff" },
  timelineCellContract: { background: "#8a8a8a" },
  timelineCellOpen: { background: "#fff4d6" },
  timelineCellMissing: { background: "#fff0f0" },
};

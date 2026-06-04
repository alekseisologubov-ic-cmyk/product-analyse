"use client";

import React, {
  startTransition,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";

import {
  downloadIngredientByLocationFileFromStorage,
  uploadIngredientByLocationFileToStorage,
} from "../../lib/permanentFiles";

const VENUE_COLUMN_INDEX = 1; // B
const INGREDIENT_NAME_COLUMN_INDEX = 7; // H
const ASSIGNED_PRODUCT_COLUMN_INDEX = 12; // M
const RECIPE_CODE_COLUMN_INDEX = 15; // P
const RECIPE_NAME_COLUMN_INDEX = 16; // Q

const INITIAL_VISIBLE_LIMIT = 100;
const VISIBLE_LIMIT_STEP = 100;

const ALLERGEN_RULES = [
  {
    key: "treeNuts",
    label: "Tree Nuts",
    keywords: [
      "almond",
      "walnut",
      "pecan",
      "cashew",
      "hazelnut",
      "pistachio",
      "macadamia",
      "brazil nut",
      "pine nut",
      "nutella",
      "praline",
      "marzipan",
      "frangipane",
    ],
    exclude: ["coconut", "nutritional yeast"],
  },
  {
    key: "peanuts",
    label: "Peanuts",
    keywords: ["peanut", "groundnut"],
    exclude: [],
  },
  {
    key: "seeds",
    label: "Seeds",
    keywords: [
      "seed",
      "seeds",
      "sunflower seed",
      "pumpkin seed",
      "chia",
      "flax",
      "hemp seed",
      "poppy seed",
      "pepita",
    ],
    exclude: ["seedless", "seedless cucumber"],
  },
  {
    key: "soy",
    label: "Soy",
    keywords: [
      "soy",
      "soya",
      "tofu",
      "edamame",
      "miso",
      "tamari",
      "soybean",
      "soy sauce",
      "teriyaki",
    ],
    exclude: [],
  },
  {
    key: "gluten",
    label: "Gluten",
    keywords: [
      "wheat",
      "flour",
      "gluten",
      "bread",
      "pasta",
      "semolina",
      "barley",
      "rye",
      "panko",
      "couscous",
      "bulgur",
      "farro",
      "orzo",
      "noodle",
      "spaghetti",
      "linguine",
      "macaroni",
      "tortilla flour",
    ],
    exclude: ["rice flour", "corn flour", "almond flour", "coconut flour"],
  },
  {
    key: "milkDairy",
    label: "Milk / Dairy",
    keywords: [
      "milk",
      "cream",
      "butter",
      "cheese",
      "yogurt",
      "yoghurt",
      "parmesan",
      "mozzarella",
      "ricotta",
      "cream cheese",
      "cheddar",
      "feta",
      "mascarpone",
      "ghee",
      "whey",
      "casein",
      "lactose",
      "buttermilk",
      "sour cream",
      "half and half",
    ],
    exclude: ["coconut milk", "almond milk", "oat milk", "soy milk"],
  },
  {
    key: "egg",
    label: "Egg",
    keywords: ["egg", "eggs", "mayonnaise", "mayo", "aioli", "albumen"],
    exclude: ["eggplant"],
  },
  {
    key: "fish",
    label: "Fish",
    keywords: [
      "salmon",
      "tuna",
      "cod",
      "anchovy",
      "fish",
      "sardine",
      "trout",
      "halibut",
      "bass",
      "snapper",
      "mackerel",
      "haddock",
      "sole",
      "tilapia",
      "branzino",
      "sea bass",
      "worcestershire",
      "fish sauce",
    ],
    exclude: [],
  },
  {
    key: "shellfish",
    label: "Shellfish",
    keywords: [
      "shrimp",
      "prawn",
      "crab",
      "lobster",
      "mussel",
      "oyster",
      "scallop",
      "clam",
      "shellfish",
      "crayfish",
      "crawfish",
    ],
    exclude: ["clam shell", "clamshell", "packed in a clam shell"],
  },
  {
    key: "sesame",
    label: "Sesame",
    keywords: ["sesame", "tahini"],
    exclude: [],
  },
  {
    key: "mustard",
    label: "Mustard",
    keywords: ["mustard", "dijon"],
    exclude: [],
  },
];

const localStyles = {
  compactButton: {
    padding: "9px 12px",
    borderRadius: 999,
    border: "1px solid #ccc",
    background: "#fff",
    cursor: "pointer",
    fontWeight: "bold",
    whiteSpace: "nowrap",
  },
  activeButton: {
    background: "#111",
    color: "#fff",
    borderColor: "#111",
  },
  filterGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 220px), 1fr))",
    gap: 10,
  },
  summaryGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 150px), 1fr))",
    gap: 10,
    marginTop: 12,
  },
  summaryTile: {
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 12,
    background: "#fafafa",
    display: "grid",
    gap: 4,
  },
  summaryNumber: {
    fontSize: 24,
    fontWeight: "bold",
  },
  progressBarOuter: {
    width: "100%",
    height: 10,
    borderRadius: 999,
    background: "#eee",
    overflow: "hidden",
    marginTop: 8,
  },
  progressBarInner: {
    height: "100%",
    borderRadius: 999,
    background: "#111",
    transition: "width 0.2s ease",
  },
  tableScroll: {
    width: "100%",
    overflowX: "auto",
    border: "1px solid #ddd",
    borderRadius: 14,
    background: "#fff",
  },
  table: {
    width: "100%",
    borderCollapse: "collapse",
    minWidth: 1120,
    fontSize: 12,
  },
  th: {
    position: "sticky",
    top: 0,
    background: "#111",
    color: "#fff",
    padding: "10px 8px",
    borderRight: "1px solid #333",
    textAlign: "left",
    zIndex: 1,
    whiteSpace: "nowrap",
  },
  td: {
    padding: "9px 8px",
    borderRight: "1px solid #eee",
    borderTop: "1px solid #eee",
    verticalAlign: "top",
  },
  rowButton: {
    border: 0,
    background: "transparent",
    padding: 0,
    margin: 0,
    textAlign: "left",
    cursor: "pointer",
    fontWeight: "bold",
    color: "#111",
  },
  allergenYes: {
    display: "inline-block",
    padding: "5px 7px",
    borderRadius: 999,
    background: "#fff0f0",
    color: "#b00020",
    fontWeight: "bold",
    whiteSpace: "nowrap",
  },
  allergenNo: {
    color: "#777",
    whiteSpace: "nowrap",
  },
  pillWrap: {
    display: "flex",
    flexWrap: "wrap",
    gap: 6,
    marginTop: 6,
  },
  pillBad: {
    display: "inline-block",
    padding: "5px 8px",
    borderRadius: 999,
    background: "#b00020",
    color: "#fff",
    fontSize: 12,
    fontWeight: "bold",
  },
  pillGood: {
    display: "inline-block",
    padding: "5px 8px",
    borderRadius: 999,
    background: "#e8f5e9",
    color: "#2e7d32",
    fontSize: 12,
    fontWeight: "bold",
  },
  modalSection: {
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 12,
    background: "#fafafa",
    marginTop: 12,
  },
  ingredientList: {
    maxHeight: 260,
    overflowY: "auto",
    border: "1px solid #eee",
    borderRadius: 10,
    padding: 10,
    background: "#fff",
  },
};

const waitForNextFrame = () =>
  new Promise((resolve) => {
    if (
      typeof window === "undefined" ||
      typeof window.requestAnimationFrame !== "function"
    ) {
      resolve();
      return;
    }

    window.requestAnimationFrame(() => resolve());
  });

const runAfterPaint = (callback) => {
  if (typeof window === "undefined") {
    callback();
    return;
  }

  window.setTimeout(() => {
    if (typeof window.requestIdleCallback === "function") {
      window.requestIdleCallback(callback, { timeout: 700 });
      return;
    }

    callback();
  }, 0);
};

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const normalizeDisplayText = (value) =>
  String(value || "")
    .replace(/\s+/g, " ")
    .trim();

const normalizeSearchText = (value) =>
  String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();

const normalizeKey = (value) =>
  cleanText(value).replace(/[^A-Z0-9]+/g, " ").trim();

const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

const isNumericOnly = (value) => {
  const text = String(value || "").trim();
  return Boolean(text) && /^-?\d+(\.\d+)?$/.test(text);
};

const getCell = (row, index) => normalizeDisplayText(row?.[index]);

const normalizeAllergenText = (value) =>
  " " +
  String(value || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim() +
  " ";

const keywordFound = (normalizedText, keyword) => {
  const normalizedKeyword = normalizeAllergenText(keyword).trim();

  if (!normalizedKeyword) return false;

  return normalizedText.includes(" " + normalizedKeyword + " ");
};

const getMatchedKeywordForRule = (rule, value) => {
  const normalizedText = normalizeAllergenText(value);

  const excluded = (rule.exclude || []).some((keyword) =>
    keywordFound(normalizedText, keyword)
  );

  if (excluded) return "";

  return (
    (rule.keywords || []).find((keyword) =>
      keywordFound(normalizedText, keyword)
    ) || ""
  );
};

const detectAllergensForIngredient = (ingredientText) => {
  const matches = [];

  ALLERGEN_RULES.forEach((rule) => {
    const matchedKeyword = getMatchedKeywordForRule(rule, ingredientText);

    if (!matchedKeyword) return;

    matches.push({
      key: rule.key,
      label: rule.label,
      keyword: matchedKeyword,
    });
  });

  return matches;
};

const makeRecipeKey = (recipeCode, recipeName) => {
  const code = normalizeKey(recipeCode);
  const name = normalizeKey(recipeName);

  return `${code || "NO-CODE"}__${name || "NO-NAME"}`;
};

const makeIngredientKey = (ingredient) =>
  normalizeKey(
    [
      ingredient?.name,
      ingredient?.rawText,
      ingredient?.sourceProductName,
      ingredient?.sourceAssignedProduct,
    ].join(" ")
  );

const createRecipeGroups = async ({ rawRows, onProgress, shouldCancel }) => {
  const groups = new Map();

  const dataRows = Array.isArray(rawRows) ? rawRows.slice(1) : [];
  const totalRows = dataRows.length;

  for (let index = 0; index < dataRows.length; index += 1) {
    if (shouldCancel?.()) {
      return {
        groups,
        cancelled: true,
      };
    }

    const row = dataRows[index];

    const venue = getCell(row, VENUE_COLUMN_INDEX) || "Unknown Venue";
    const recipeCode = getCell(row, RECIPE_CODE_COLUMN_INDEX);
    const recipeName = getCell(row, RECIPE_NAME_COLUMN_INDEX);

    if (!recipeName || isNumericOnly(recipeName)) {
      continue;
    }

    const sourceProductName = getCell(row, INGREDIENT_NAME_COLUMN_INDEX);
    const sourceAssignedProduct = getCell(row, ASSIGNED_PRODUCT_COLUMN_INDEX);

    const ingredientName = sourceAssignedProduct || sourceProductName;

    if (!ingredientName) {
      continue;
    }

    const recipeKey = makeRecipeKey(recipeCode, recipeName);

    if (!groups.has(recipeKey)) {
      groups.set(recipeKey, {
        recipeKey,
        recipeCode,
        recipeName,
        venues: new Set(),
        ingredientMap: new Map(),
        rowCount: 0,
      });
    }

    const group = groups.get(recipeKey);

    group.venues.add(venue);
    group.rowCount += 1;

    const ingredient = {
      name: ingredientName,
      rawText: [sourceAssignedProduct, sourceProductName]
        .filter(Boolean)
        .join(" / "),
      sourceProductName,
      sourceAssignedProduct,
    };

    const ingredientKey = makeIngredientKey(ingredient);

    if (ingredientKey && !group.ingredientMap.has(ingredientKey)) {
      group.ingredientMap.set(ingredientKey, ingredient);
    }

    if (index % 500 === 0) {
      onProgress?.({
        phase: "Reading recipes",
        done: index,
        total: totalRows,
      });

      await waitForNextFrame();
    }
  }

  onProgress?.({
    phase: "Reading recipes",
    done: totalRows,
    total: totalRows,
  });

  return {
    groups,
    cancelled: false,
  };
};

const buildIngredientsByRecipeName = (recipeGroups) => {
  const map = new Map();

  Array.from(recipeGroups.values()).forEach((group) => {
    const key = normalizeKey(group.recipeName);
    if (!key) return;

    if (!map.has(key)) {
      map.set(key, []);
    }

    map.get(key).push(...Array.from(group.ingredientMap.values()));
  });

  return map;
};

const expandRecipeIngredients = (group, ingredientsByRecipeName) => {
  const output = [];
  const seenOutput = new Set();

  const addOutput = (ingredient, path) => {
    const displayName = [...path, ingredient.name].filter(Boolean).join(" → ");
    const rawText = [displayName, ingredient.rawText].filter(Boolean).join(" ");

    const key = normalizeKey(rawText);

    if (!key || seenOutput.has(key)) return;

    seenOutput.add(key);

    output.push({
      displayName,
      name: ingredient.name,
      rawText,
      sourceProductName: ingredient.sourceProductName || "",
      sourceAssignedProduct: ingredient.sourceAssignedProduct || "",
    });
  };

  const visitIngredient = (ingredient, path = [], depth = 0, visited = new Set()) => {
    addOutput(ingredient, path);

    if (depth >= 2) return;

    const ingredientRecipeKey = normalizeKey(ingredient.name);

    if (!ingredientRecipeKey || visited.has(ingredientRecipeKey)) {
      return;
    }

    const subIngredients = ingredientsByRecipeName.get(ingredientRecipeKey);

    if (!subIngredients?.length) {
      return;
    }

    const nextVisited = new Set(visited);
    nextVisited.add(ingredientRecipeKey);

    subIngredients.forEach((subIngredient) => {
      if (normalizeKey(subIngredient.name) === ingredientRecipeKey) {
        return;
      }

      visitIngredient(
        subIngredient,
        [...path, ingredient.name],
        depth + 1,
        nextVisited
      );
    });
  };

  Array.from(group.ingredientMap.values()).forEach((ingredient) => {
    visitIngredient(ingredient);
  });

  return output;
};

const buildMatrixRowForRecipe = (group, ingredientsByRecipeName) => {
  const expandedIngredients = expandRecipeIngredients(group, ingredientsByRecipeName);

  const allergenMap = {};
  const matchedIngredientSet = new Set();

  ALLERGEN_RULES.forEach((rule) => {
    allergenMap[rule.label] = [];
  });

  expandedIngredients.forEach((ingredient) => {
    const ingredientSearchText = [
      ingredient.displayName,
      ingredient.rawText,
      ingredient.sourceProductName,
      ingredient.sourceAssignedProduct,
    ]
      .filter(Boolean)
      .join(" ");

    const matches = detectAllergensForIngredient(ingredientSearchText);

    matches.forEach((match) => {
      const matchKey = `${match.label}__${ingredient.displayName}__${match.keyword}`;

      if (matchedIngredientSet.has(matchKey)) {
        return;
      }

      matchedIngredientSet.add(matchKey);

      allergenMap[match.label].push({
        ingredient: ingredient.displayName,
        keyword: match.keyword,
      });
    });
  });

  const allergenLabels = ALLERGEN_RULES.map((rule) => rule.label).filter(
    (label) => allergenMap[label]?.length
  );

  const venues = Array.from(group.venues || []).sort();

  const searchText = normalizeSearchText(
    [
      group.recipeCode,
      group.recipeName,
      venues.join(" "),
      expandedIngredients.map((item) => item.displayName).join(" "),
      allergenLabels.join(" "),
      ALLERGEN_RULES.flatMap((rule) =>
        (allergenMap[rule.label] || []).map(
          (match) => `${match.ingredient} ${match.keyword}`
        )
      ).join(" "),
    ].join(" ")
  );

  return {
    recipeKey: group.recipeKey,
    recipeCode: group.recipeCode,
    recipeName: group.recipeName,
    venues,
    ingredients: expandedIngredients,
    ingredientCount: expandedIngredients.length,
    originalRowCount: group.rowCount,
    allergenMap,
    allergenLabels,
    warningCount: allergenLabels.length,
    hasWarnings: allergenLabels.length > 0,
    searchText,
  };
};

const buildAllergenMatrixRows = async ({
  rawRows,
  onProgress,
  shouldCancel,
}) => {
  const groupResult = await createRecipeGroups({
    rawRows,
    onProgress,
    shouldCancel,
  });

  if (groupResult.cancelled) {
    return {
      rows: [],
      cancelled: true,
    };
  }

  const recipeGroups = groupResult.groups;
  const groupList = Array.from(recipeGroups.values()).sort((a, b) => {
    const recipeNameCompare = String(a.recipeName || "").localeCompare(
      String(b.recipeName || "")
    );

    if (recipeNameCompare !== 0) return recipeNameCompare;

    return String(a.recipeCode || "").localeCompare(String(b.recipeCode || ""));
  });

  const ingredientsByRecipeName = buildIngredientsByRecipeName(recipeGroups);
  const outputRows = [];

  for (let index = 0; index < groupList.length; index += 1) {
    if (shouldCancel?.()) {
      return {
        rows: [],
        cancelled: true,
      };
    }

    outputRows.push(buildMatrixRowForRecipe(groupList[index], ingredientsByRecipeName));

    if (index % 35 === 0) {
      onProgress?.({
        phase: "Building matrix",
        done: index,
        total: groupList.length,
      });

      await waitForNextFrame();
    }
  }

  onProgress?.({
    phase: "Building matrix",
    done: groupList.length,
    total: groupList.length,
  });

  return {
    rows: outputRows,
    cancelled: false,
  };
};

const workbookToRows = (XLSX, workbook) => {
  const sheetName = workbook?.SheetNames?.[0] || "";
  const worksheet = workbook?.Sheets?.[sheetName];

  if (!worksheet) {
    return {
      rows: [],
      sheetName,
    };
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  return {
    rows,
    sheetName,
  };
};

const readWorkbookRowsFromArrayBuffer = async (arrayBuffer) => {
  const XLSX = await import("xlsx");

  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
  });

  return {
    workbook,
    ...workbookToRows(XLSX, workbook),
  };
};

const getProgressPercent = (progress) => {
  const total = Number(progress?.total || 0);
  const done = Number(progress?.done || 0);

  if (!total) return 0;

  return Math.max(0, Math.min(100, Math.round((done / total) * 100)));
};

export default function AllergenModule({
  styles,
  supabase,
  userShip,
  userEmail,
  isAdmin,
  onBack,
  logUsageEvent,
  recipeRows = null,
  setRecipeRows = null,
}) {
  const [sourceRows, setSourceRows] = useState([]);
  const [sourceFileName, setSourceFileName] = useState("");
  const [sourceSheetName, setSourceSheetName] = useState("");
  const [sourceMessage, setSourceMessage] = useState("");

  const [matrixRows, setMatrixRows] = useState([]);
  const [matrixLoading, setMatrixLoading] = useState(false);
  const [matrixProgress, setMatrixProgress] = useState({
    phase: "",
    done: 0,
    total: 0,
  });

  const [search, setSearch] = useState("");
  const [debouncedSearch, setDebouncedSearch] = useState("");
  const [selectedAllergen, setSelectedAllergen] = useState("ALL");
  const [selectedVenue, setSelectedVenue] = useState("ALL");
  const [warningFilter, setWarningFilter] = useState("warnings");
  const [visibleLimit, setVisibleLimit] = useState(INITIAL_VISIBLE_LIMIT);
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [adminUploadBusy, setAdminUploadBusy] = useState(false);

  const autoLoadRef = useRef(false);
  const buildJobRef = useRef(0);
  const loggedOpenRef = useRef(false);

  const appStyles = styles || {};

  useEffect(() => {
    if (loggedOpenRef.current) return;

    loggedOpenRef.current = true;

    logUsageEvent?.("allergen_matrix_opened", {
      module: "allergen",
      ship: userShip || "",
      userEmail: userEmail || "",
    });
  }, [logUsageEvent, userShip, userEmail]);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      setDebouncedSearch(search);
    }, 250);

    return () => window.clearTimeout(timer);
  }, [search]);

  useEffect(() => {
    setVisibleLimit(INITIAL_VISIBLE_LIMIT);
  }, [debouncedSearch, selectedAllergen, selectedVenue, warningFilter]);

  useEffect(() => {
    if (Array.isArray(recipeRows) && recipeRows.length > 1) {
      setSourceRows(recipeRows);
      setSourceFileName("Permanent Ingredient by Location");
      setSourceSheetName("Loaded from app");
      setSourceMessage(
        `Ingredient by Location rows received from app. ${Math.max(
          recipeRows.length - 1,
          0
        )} row(s).`
      );
    }
  }, [recipeRows]);

  const loadPermanentIngredientFile = useCallback(
    async ({ silent = false } = {}) => {
      if (!supabase) {
        if (!silent) {
          setSourceMessage(
            "Supabase is not connected. Permanent Ingredient by Location file cannot load."
          );
        }

        return false;
      }

      try {
        if (!silent) {
          setSourceMessage("Loading permanent Ingredient by Location file...");
        }

        const arrayBuffer =
          await downloadIngredientByLocationFileFromStorage({ supabase });

        const parsed = await readWorkbookRowsFromArrayBuffer(arrayBuffer);

        setSourceRows(parsed.rows);
        setSourceFileName("Permanent Ingredient by Location");
        setSourceSheetName(parsed.sheetName || "");

        if (typeof setRecipeRows === "function") {
          setRecipeRows(parsed.rows);
        }

        setSourceMessage(
          `Permanent Ingredient by Location loaded. ${Math.max(
            parsed.rows.length - 1,
            0
          )} row(s).`
        );

        logUsageEvent?.("allergen_permanent_file_loaded", {
          module: "allergen",
          ship: userShip || "",
          sheetName: parsed.sheetName || "",
          rows: Math.max(parsed.rows.length - 1, 0),
        });

        return true;
      } catch (error) {
        const text =
          error?.message ||
          "Could not load permanent Ingredient by Location file.";

        setSourceRows([]);
        setSourceMessage(text);

        if (!silent && typeof window !== "undefined") {
          window.alert(text);
        }

        return false;
      }
    },
    [supabase, setRecipeRows, logUsageEvent, userShip]
  );

  useEffect(() => {
    if (autoLoadRef.current) return;

    if (Array.isArray(recipeRows) && recipeRows.length > 1) {
      return;
    }

    autoLoadRef.current = true;
    loadPermanentIngredientFile({ silent: true });
  }, [recipeRows, loadPermanentIngredientFile]);

  useEffect(() => {
    const rows = Array.isArray(sourceRows) ? sourceRows : [];

    buildJobRef.current += 1;

    const jobId = buildJobRef.current;

    if (rows.length <= 1) {
      setMatrixRows([]);
      setMatrixLoading(false);
      setMatrixProgress({
        phase: "",
        done: 0,
        total: 0,
      });
      return;
    }

    setMatrixLoading(true);
    setMatrixRows([]);
    setMatrixProgress({
      phase: "Preparing",
      done: 0,
      total: Math.max(rows.length - 1, 0),
    });

    runAfterPaint(() => {
      buildAllergenMatrixRows({
        rawRows: rows,
        onProgress: (progress) => {
          if (buildJobRef.current !== jobId) return;

          setMatrixProgress(progress);
        },
        shouldCancel: () => buildJobRef.current !== jobId,
      })
        .then((result) => {
          if (buildJobRef.current !== jobId || result.cancelled) return;

          startTransition(() => {
            setMatrixRows(result.rows);
            setMatrixLoading(false);
            setMatrixProgress({
              phase: "Complete",
              done: result.rows.length,
              total: result.rows.length,
            });
          });
        })
        .catch((error) => {
          if (buildJobRef.current !== jobId) return;

          setMatrixRows([]);
          setMatrixLoading(false);
          setSourceMessage(
            error?.message || "Could not build allergen matrix."
          );
        });
    });

    return () => {
      buildJobRef.current += 1;
    };
  }, [sourceRows]);

  const venueOptions = useMemo(() => {
    const venueSet = new Set();

    matrixRows.forEach((row) => {
      (row.venues || []).forEach((venue) => {
        if (venue) venueSet.add(venue);
      });
    });

    return Array.from(venueSet).sort((a, b) => a.localeCompare(b));
  }, [matrixRows]);

  const allergenCounts = useMemo(() => {
    const counts = {};

    ALLERGEN_RULES.forEach((rule) => {
      counts[rule.label] = 0;
    });

    matrixRows.forEach((row) => {
      ALLERGEN_RULES.forEach((rule) => {
        if (row.allergenMap?.[rule.label]?.length) {
          counts[rule.label] += 1;
        }
      });
    });

    return counts;
  }, [matrixRows]);

  const filteredRows = useMemo(() => {
    const query = normalizeSearchText(debouncedSearch);
    const activeAllergen = selectedAllergen;
    const activeVenue = selectedVenue;

    return matrixRows.filter((row) => {
      if (warningFilter === "warnings" && !row.hasWarnings) return false;
      if (warningFilter === "clear" && row.hasWarnings) return false;

      if (
        activeAllergen !== "ALL" &&
        !row.allergenMap?.[activeAllergen]?.length
      ) {
        return false;
      }

      if (
        activeVenue !== "ALL" &&
        !(row.venues || []).some((venue) => venue === activeVenue)
      ) {
        return false;
      }

      if (query && !row.searchText.includes(query)) return false;

      return true;
    });
  }, [
    matrixRows,
    debouncedSearch,
    selectedAllergen,
    selectedVenue,
    warningFilter,
  ]);

  const visibleRows = useMemo(
    () => filteredRows.slice(0, visibleLimit),
    [filteredRows, visibleLimit]
  );

  const summary = useMemo(() => {
    const recipesWithWarnings = matrixRows.filter(
      (row) => row.hasWarnings
    ).length;

    const ingredientWarnings = matrixRows.reduce((sum, row) => {
      return (
        sum +
        ALLERGEN_RULES.reduce(
          (ruleSum, rule) =>
            ruleSum + Number(row.allergenMap?.[rule.label]?.length || 0),
          0
        )
      );
    }, 0);

    return {
      totalRecipes: matrixRows.length,
      recipesWithWarnings,
      clearRecipes: matrixRows.length - recipesWithWarnings,
      ingredientWarnings,
      filtered: filteredRows.length,
    };
  }, [matrixRows, filteredRows]);

  const handleAdminUpload = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    if (!isAdmin) {
      window.alert("Only admins can replace the permanent Ingredient by Location file.");
      event.target.value = "";
      return;
    }

    if (!supabase) {
      window.alert("Supabase is not connected.");
      event.target.value = "";
      return;
    }

    setAdminUploadBusy(true);
    setSourceMessage("Saving permanent Ingredient by Location file...");

    try {
      await uploadIngredientByLocationFileToStorage({
        supabase,
        file,
      });

      const arrayBuffer = await file.arrayBuffer();
      const parsed = await readWorkbookRowsFromArrayBuffer(arrayBuffer);

      setSourceRows(parsed.rows);
      setSourceFileName(file.name);
      setSourceSheetName(parsed.sheetName || "");

      if (typeof setRecipeRows === "function") {
        setRecipeRows(parsed.rows);
      }

      setSourceMessage(
        `Permanent Ingredient by Location file updated. ${Math.max(
          parsed.rows.length - 1,
          0
        )} row(s) loaded.`
      );

      logUsageEvent?.("allergen_permanent_file_updated", {
        module: "allergen",
        ship: userShip || "",
        fileName: file.name,
        sheetName: parsed.sheetName || "",
        rows: Math.max(parsed.rows.length - 1, 0),
      });
    } catch (error) {
      const text =
        error?.message || "Could not save permanent Ingredient by Location file.";

      setSourceMessage(text);
      window.alert(text);
    } finally {
      setAdminUploadBusy(false);
      event.target.value = "";
    }
  };

  const exportAllergenMatrixToExcel = async () => {
    if (!filteredRows.length) {
      window.alert("No allergen matrix rows to export.");
      return;
    }

    const XLSX = await import("xlsx");

    const exportRows = filteredRows.map((row, index) => {
      const record = {
        Number: index + 1,
        RecipeCode: row.recipeCode || "",
        RecipeName: row.recipeName || "",
        Venues: (row.venues || []).join(", "),
        Status: row.hasWarnings ? "Allergen warning" : "No keyword warning",
        WarningCount: row.warningCount,
        IngredientCount: row.ingredientCount,
        SourceRows: row.originalRowCount,
        Ingredients: (row.ingredients || [])
          .map((ingredient) => ingredient.displayName)
          .join(", "),
      };

      ALLERGEN_RULES.forEach((rule) => {
        const matches = row.allergenMap?.[rule.label] || [];

        record[rule.label] = matches.length
          ? matches
              .map((match) => `${match.ingredient} (${match.keyword})`)
              .join("; ")
          : "";
      });

      return record;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Allergen Matrix");

    XLSX.writeFile(workbook, "allergen-matrix.xlsx");

    logUsageEvent?.("allergen_matrix_exported", {
      module: "allergen",
      ship: userShip || "",
      rows: exportRows.length,
      search: debouncedSearch,
      selectedAllergen,
      selectedVenue,
      warningFilter,
    });
  };

  const printAllergenMatrix = () => {
    if (!filteredRows.length) {
      window.alert("No allergen matrix rows to print.");
      return;
    }

    const rowsForPrint = filteredRows.slice(0, 500);

    const html = `
      <html>
        <head>
          <title>Allergen Matrix</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; color: #111; }
            h1 { margin-bottom: 4px; }
            table { width: 100%; border-collapse: collapse; margin-top: 18px; font-size: 11px; }
            th, td { border: 1px solid #ccc; padding: 6px; text-align: left; vertical-align: top; }
            th { background: #111; color: #fff; }
            .bad { color: #b00020; font-weight: bold; }
            .good { color: #2e7d32; font-weight: bold; }
          </style>
        </head>
        <body>
          <h1>Allergen Matrix</h1>
          <div><strong>Ship:</strong> ${escapeHtml(userShip || "N/A")}</div>
          <div><strong>Source:</strong> ${escapeHtml(sourceFileName || "Ingredient by Location")}</div>
          <div><strong>Rows printed:</strong> ${rowsForPrint.length}</div>
          <div><strong>Generated:</strong> ${escapeHtml(new Date().toLocaleString())}</div>

          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Recipe</th>
                <th>Venues</th>
                <th>Status</th>
                ${ALLERGEN_RULES.map(
                  (rule) => `<th>${escapeHtml(rule.label)}</th>`
                ).join("")}
              </tr>
            </thead>
            <tbody>
              ${rowsForPrint
                .map((row, index) => {
                  return `
                    <tr>
                      <td>${index + 1}</td>
                      <td>
                        <strong>${escapeHtml(row.recipeName)}</strong><br />
                        ${escapeHtml(row.recipeCode || "No code")}
                      </td>
                      <td>${escapeHtml((row.venues || []).join(", "))}</td>
                      <td class="${row.hasWarnings ? "bad" : "good"}">
                        ${row.hasWarnings ? "Warning" : "No keyword warning"}
                      </td>
                      ${ALLERGEN_RULES.map((rule) => {
                        const matches = row.allergenMap?.[rule.label] || [];
                        return `<td>${
                          matches.length
                            ? escapeHtml(
                                matches
                                  .map(
                                    (match) =>
                                      `${match.ingredient} (${match.keyword})`
                                  )
                                  .join("; ")
                              )
                            : ""
                        }</td>`;
                      }).join("")}
                    </tr>
                  `;
                })
                .join("")}
            </tbody>
          </table>
        </body>
      </html>
    `;

    const printWindow = window.open("", "_blank");

    if (!printWindow) {
      window.alert("The print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();

    logUsageEvent?.("allergen_matrix_printed", {
      module: "allergen",
      ship: userShip || "",
      rows: rowsForPrint.length,
    });
  };

  const progressPercent = getProgressPercent(matrixProgress);

  return (
    <main style={appStyles.page}>
      <header style={appStyles.header}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={appStyles.headerLogo}
        />

        <div style={appStyles.headerActions}>
          <button type="button" style={appStyles.backButton} onClick={onBack}>
            ← Back
          </button>

          <div style={appStyles.shipBadge}>🚢 {userShip || "Ship"}</div>
        </div>
      </header>

      <section style={appStyles.grid}>
        <div style={appStyles.card}>
          <h2 style={appStyles.cardTitle}>🧬 Allergen Matrix</h2>

          <p style={appStyles.emptyText}>
            Matrix is built from the permanent Ingredient by Location file.
            It opens first, then builds in chunks so the page does not freeze.
          </p>

          <div style={appStyles.infoBox}>
            <div>
              📄 Source:{" "}
              <strong>{sourceFileName || "Permanent Ingredient by Location"}</strong>
            </div>

            <div>
              📑 Sheet: <strong>{sourceSheetName || "N/A"}</strong>
            </div>

            <div>
              📦 Source rows:{" "}
              <strong>{Math.max((sourceRows || []).length - 1, 0)}</strong>
            </div>

            <div>
              🍽️ Recipes: <strong>{summary.totalRecipes}</strong>
            </div>

            <div>
              ⚠️ Recipes with warnings:{" "}
              <strong>{summary.recipesWithWarnings}</strong>
            </div>

            {matrixLoading && (
              <>
                <div>
                  Building:{" "}
                  <strong>
                    {matrixProgress.phase || "Working"}{" "}
                    {matrixProgress.total
                      ? `${matrixProgress.done} / ${matrixProgress.total}`
                      : ""}
                  </strong>
                </div>

                <div style={localStyles.progressBarOuter}>
                  <div
                    style={{
                      ...localStyles.progressBarInner,
                      width: `${progressPercent}%`,
                    }}
                  />
                </div>
              </>
            )}
          </div>

          {sourceMessage && <p style={appStyles.message}>{sourceMessage}</p>}

          <div style={appStyles.headerActions}>
            <button
              type="button"
              style={appStyles.backButton}
              onClick={() => loadPermanentIngredientFile()}
              disabled={matrixLoading || adminUploadBusy}
            >
              🔄 Reload Permanent File
            </button>

            <button
              type="button"
              style={appStyles.backButton}
              onClick={printAllergenMatrix}
              disabled={matrixLoading || !filteredRows.length}
            >
              🖨️ Print
            </button>

            <button
              type="button"
              style={appStyles.primaryButton}
              onClick={exportAllergenMatrixToExcel}
              disabled={matrixLoading || !filteredRows.length}
            >
              📥 Export Excel
            </button>
          </div>

          {isAdmin && (
            <div style={appStyles.reportFilterBox}>
              <label style={appStyles.label}>
                Admin only: replace permanent Ingredient by Location file
              </label>

              <input
                type="file"
                accept=".xlsx,.xls,.xlsm"
                onChange={handleAdminUpload}
                style={appStyles.fileInput}
                disabled={adminUploadBusy || matrixLoading}
              />

              <div style={appStyles.recipeMeta}>
                This replaces the permanent file used by the Allergen Matrix.
              </div>
            </div>
          )}
        </div>

        <div style={appStyles.card}>
          <h2 style={appStyles.cardTitle}>🔍 Search / Filter</h2>

          <input
            placeholder="Search recipe, ingredient, venue, allergen..."
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            style={appStyles.searchInput}
          />

          <div style={localStyles.filterGrid}>
            <div>
              <label style={appStyles.label}>Warning filter</label>
              <select
                value={warningFilter}
                onChange={(event) => setWarningFilter(event.target.value)}
                style={appStyles.searchInput}
              >
                <option value="all">All recipes</option>
                <option value="warnings">Warnings only</option>
                <option value="clear">No keyword warning</option>
              </select>
            </div>

            <div>
              <label style={appStyles.label}>Allergen</label>
              <select
                value={selectedAllergen}
                onChange={(event) => setSelectedAllergen(event.target.value)}
                style={appStyles.searchInput}
              >
                <option value="ALL">All allergens</option>

                {ALLERGEN_RULES.map((rule) => (
                  <option key={rule.key} value={rule.label}>
                    {rule.label} ({allergenCounts[rule.label] || 0})
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label style={appStyles.label}>Venue</label>
              <select
                value={selectedVenue}
                onChange={(event) => setSelectedVenue(event.target.value)}
                style={appStyles.searchInput}
              >
                <option value="ALL">All venues</option>

                {venueOptions.map((venue) => (
                  <option key={venue} value={venue}>
                    {venue}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div style={localStyles.summaryGrid}>
            <div style={localStyles.summaryTile}>
              <span>Shown</span>
              <strong style={localStyles.summaryNumber}>
                {summary.filtered}
              </strong>
            </div>

            <div style={localStyles.summaryTile}>
              <span>Warnings</span>
              <strong style={localStyles.summaryNumber}>
                {summary.recipesWithWarnings}
              </strong>
            </div>

            <div style={localStyles.summaryTile}>
              <span>Clear</span>
              <strong style={localStyles.summaryNumber}>
                {summary.clearRecipes}
              </strong>
            </div>

            <div style={localStyles.summaryTile}>
              <span>Ingredient hits</span>
              <strong style={localStyles.summaryNumber}>
                {summary.ingredientWarnings}
              </strong>
            </div>
          </div>

          <p style={appStyles.warningText}>
            This is keyword-based guidance only. Always verify against the
            official allergen data before operational use.
          </p>
        </div>
      </section>

      <section style={appStyles.card}>
        <div
          style={{
            ...appStyles.header,
            boxShadow: "none",
            padding: 0,
            marginBottom: 16,
          }}
        >
          <div>
            <h2 style={appStyles.productTitle}>📋 Recipe Allergen Matrix</h2>
            <p style={{ ...appStyles.emptyText, margin: 0 }}>
              Showing {visibleRows.length} of {filteredRows.length} filtered
              recipe row(s).
            </p>
          </div>

          <div style={appStyles.headerActions}>
            <button
              type="button"
              style={{
                ...localStyles.compactButton,
                ...(warningFilter === "warnings"
                  ? localStyles.activeButton
                  : {}),
              }}
              onClick={() => setWarningFilter("warnings")}
            >
              ⚠️ Warnings
            </button>

            <button
              type="button"
              style={{
                ...localStyles.compactButton,
                ...(warningFilter === "all" ? localStyles.activeButton : {}),
              }}
              onClick={() => setWarningFilter("all")}
            >
              📋 All
            </button>

            <button
              type="button"
              style={{
                ...localStyles.compactButton,
                ...(warningFilter === "clear" ? localStyles.activeButton : {}),
              }}
              onClick={() => setWarningFilter("clear")}
            >
              ✅ Clear
            </button>
          </div>
        </div>

        {!sourceRows.length && (
          <p style={appStyles.emptyText}>
            Loading the permanent Ingredient by Location file. Use Reload if it
            does not appear.
          </p>
        )}

        {sourceRows.length > 0 && matrixLoading && (
          <p style={appStyles.emptyText}>
            Building matrix in the background. The app should stay responsive.
          </p>
        )}

        {sourceRows.length > 0 && !matrixLoading && !filteredRows.length && (
          <p style={appStyles.emptyText}>
            No recipes matched the current search/filter.
          </p>
        )}

        {visibleRows.length > 0 && (
          <div style={localStyles.tableScroll}>
            <table style={localStyles.table}>
              <thead>
                <tr>
                  <th style={localStyles.th}>Recipe</th>
                  <th style={localStyles.th}>Venues</th>
                  <th style={localStyles.th}>Ingredients</th>
                  <th style={localStyles.th}>Status</th>

                  {ALLERGEN_RULES.map((rule) => (
                    <th key={rule.key} style={localStyles.th}>
                      {rule.label}
                    </th>
                  ))}
                </tr>
              </thead>

              <tbody>
                {visibleRows.map((row) => (
                  <tr key={row.recipeKey}>
                    <td style={localStyles.td}>
                      <button
                        type="button"
                        style={localStyles.rowButton}
                        onClick={() => setSelectedRecipe(row)}
                      >
                        {row.recipeName || "Unnamed recipe"}
                      </button>

                      <div style={appStyles.recipeMeta}>
                        Code: {row.recipeCode || "N/A"}
                      </div>
                    </td>

                    <td style={localStyles.td}>
                      {(row.venues || []).slice(0, 4).join(", ") || "N/A"}
                      {(row.venues || []).length > 4
                        ? ` +${row.venues.length - 4}`
                        : ""}
                    </td>

                    <td style={localStyles.td}>
                      <strong>{row.ingredientCount}</strong> ingredient(s)
                      <div style={appStyles.recipeMeta}>
                        Source rows: {row.originalRowCount}
                      </div>
                    </td>

                    <td style={localStyles.td}>
                      {row.hasWarnings ? (
                        <span style={localStyles.allergenYes}>
                          ⚠️ {row.warningCount} warning(s)
                        </span>
                      ) : (
                        <span style={localStyles.pillGood}>
                          No keyword warning
                        </span>
                      )}
                    </td>

                    {ALLERGEN_RULES.map((rule) => {
                      const matches = row.allergenMap?.[rule.label] || [];

                      return (
                        <td key={rule.key} style={localStyles.td}>
                          {matches.length ? (
                            <button
                              type="button"
                              style={{
                                ...localStyles.rowButton,
                                color: "#b00020",
                              }}
                              onClick={() => setSelectedRecipe(row)}
                            >
                              Yes ({matches.length})
                            </button>
                          ) : (
                            <span style={localStyles.allergenNo}>—</span>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {filteredRows.length > visibleLimit && (
          <div style={{ marginTop: 14, textAlign: "center" }}>
            <button
              type="button"
              style={appStyles.primaryButton}
              onClick={() =>
                setVisibleLimit((current) => current + VISIBLE_LIMIT_STEP)
              }
            >
              Show More ({Math.min(visibleLimit + VISIBLE_LIMIT_STEP, filteredRows.length)} / {filteredRows.length})
            </button>
          </div>
        )}
      </section>

      {selectedRecipe && (
        <div
          style={appStyles.modalBackdrop}
          onClick={() => setSelectedRecipe(null)}
        >
          <div
            style={appStyles.modalCard}
            onClick={(event) => event.stopPropagation()}
          >
            <button
              type="button"
              style={appStyles.closeButton}
              onClick={() => setSelectedRecipe(null)}
            >
              ✕
            </button>

            <h2>{selectedRecipe.recipeName}</h2>

            <p>
              <strong>Recipe code:</strong>{" "}
              {selectedRecipe.recipeCode || "N/A"}
            </p>

            <p>
              <strong>Venues:</strong>{" "}
              {(selectedRecipe.venues || []).join(", ") || "N/A"}
            </p>

            <p>
              <strong>Status:</strong>{" "}
              {selectedRecipe.hasWarnings
                ? `${selectedRecipe.warningCount} allergen warning(s)`
                : "No keyword warning"}
            </p>

            <div style={localStyles.modalSection}>
              <h3 style={appStyles.sectionTitle}>⚠️ Matched Allergens</h3>

              {!selectedRecipe.hasWarnings && (
                <p style={appStyles.emptyText}>
                  No likely allergens detected by keyword rules.
                </p>
              )}

              {selectedRecipe.hasWarnings &&
                ALLERGEN_RULES.map((rule) => {
                  const matches = selectedRecipe.allergenMap?.[rule.label] || [];

                  if (!matches.length) return null;

                  return (
                    <div key={rule.key} style={{ marginBottom: 12 }}>
                      <strong style={{ color: "#b00020" }}>
                        {rule.label}
                      </strong>

                      <ul>
                        {matches.map((match, index) => (
                          <li key={`${rule.key}-${index}`}>
                            {match.ingredient}{" "}
                            <strong>({match.keyword})</strong>
                          </li>
                        ))}
                      </ul>
                    </div>
                  );
                })}
            </div>

            <div style={localStyles.modalSection}>
              <h3 style={appStyles.sectionTitle}>🧾 Ingredients scanned</h3>

              <div style={localStyles.ingredientList}>
                {(selectedRecipe.ingredients || []).map((ingredient, index) => (
                  <div key={`${ingredient.displayName}-${index}`}>
                    {ingredient.displayName}
                  </div>
                ))}
              </div>
            </div>

            <p style={appStyles.warningText}>
              This is a keyword-based warning only. Verify against official
              allergen data before use.
            </p>
          </div>
        </div>
      )}
    </main>
  );
}

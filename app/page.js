"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

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
  { allergen: "Mustard", keywords: ["mustard"] }
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

export default function App() {
  const [userShip, setUserShip] = useState("");
  const [loggedIn, setLoggedIn] = useState(false);
  const [module, setModule] = useState("");
  const [equipmentMode, setEquipmentMode] = useState("");

  const [consumptionRows, setConsumptionRows] = useState([]);
  const [recipeRows, setRecipeRows] = useState([]);
  const [templateMap, setTemplateMap] = useState({});
  const [templateStatus, setTemplateStatus] = useState("Loading template...");
  const [products, setProducts] = useState([]);
  const [selectedProduct, setSelectedProduct] = useState("");
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [search, setSearch] = useState("");
  const [message, setMessage] = useState("");
  const [viewMode, setViewMode] = useState("all");

  const [musterRows, setMusterRows] = useState([]);
  const [musterSearch, setMusterSearch] = useState("");

  const shipColumns = { BRL: 8, RL: 11, SC: 14, VL: 17 };

  useEffect(() => {
    loadDefaultTemplate();
  }, []);

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

      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (!rows.length) return;

      rows.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          const header = cleanText(cell);
          if (header !== "INGREDIENT NAME") return;

          const templateName =
            String(
              rows[rowIndex - 1]?.[colIndex] ||
              rows[rowIndex - 1]?.[colIndex - 1] ||
              sheetName ||
              "Template"
            ).trim();

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
              map[venueKey][productKey] = {
                product,
                templates: new Set(),
              };
            }

            map[venueKey][productKey].templates.add(templateName);
          });
        });
      });
    });

    Object.keys(map).forEach((venueKey) => {
      Object.keys(map[venueKey]).forEach((productKey) => {
        map[venueKey][productKey].templates = [
          ...map[venueKey][productKey].templates,
        ];
      });
    });

    return map;
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
      const rows = workbookToRows(workbook);
      setMusterRows(rows);
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
        (templateProductKey.length > 12 &&
          (selected.includes(templateProductKey) || templateProductKey.includes(selected)));

      if (isMatch) {
        matches.push(...data.templates);
      }
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
        required[venueKey] = {
          displayName: venueRaw || venueKey,
          recipes: new Set(),
        };
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
        result[venueKey] = {
          displayName: venue,
          ships: {},
        };
      }

      SHIPS.forEach((ship) => {
        const qty = Number(row[shipColumns[ship]]) || 0;
        result[venueKey].ships[ship] =
          (result[venueKey].ships[ship] || 0) + qty;
      });
    });

    return result;
  };

  const getCombinedVenueBreakdown = (product) => {
    const actual = getConsumptionBreakdown(product);
    const required = getRequiredVenuesForProduct(product);

    const allVenueKeys = Array.from(
      new Set([...Object.keys(actual), ...Object.keys(required)])
    ).sort();

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
        displayName:
          actualVenue?.displayName ||
          requiredVenue?.displayName ||
          venueKey,
        ships,
        required: requiredByRecipe,
        missingShips: visibleShips.filter(
          (ship) => requiredByRecipe && (ships[ship] || 0) === 0
        ),
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

      if (!productMatches(product, row)) return;

      const key = `${recipeCode || "N/A"} - ${
        recipeName || "Unnamed Recipe"
      }`;

      if (!recipes[key]) {
        recipes[key] = {
          key,
          recipeCode,
          recipeName,
        };
      }
    });

    return Object.values(recipes);
  };

  const getProductsInRecipe = (recipe) => {
    if (!recipe) return [];

    return recipeData
      .filter(
        (r) =>
          String(r[15]) === recipe.recipeCode &&
          String(r[16]) === recipe.recipeName
      )
      .map((r) => String(r[12] || r[7]))
      .filter(Boolean);
  };

  const detectAllergens = (productsInRecipe) => {
    const found = {};

    productsInRecipe.forEach((product) => {
      const lower = product.toLowerCase();

      ALLERGEN_RULES.forEach((rule) => {
        if (rule.exclude?.some((e) => lower.includes(e))) return;

        if (rule.keywords.some((k) => lower.includes(k))) {
          if (!found[rule.allergen]) found[rule.allergen] = [];
          found[rule.allergen].push(product);
        }
      });
    });

    return found;
  };

  const combinedBreakdown = selectedProduct
    ? getCombinedVenueBreakdown(selectedProduct)
    : [];

  const recipesForProduct = selectedProduct
    ? getRecipesUsingProduct(selectedProduct)
    : [];
  
  const totalConsumption = (() => {
    const totals = { BRL: 0, RL: 0, SC: 0, VL: 0 };

    combinedBreakdown.forEach((venue) => {
      visibleShips.forEach((ship) => {
        totals[ship] += Number(venue.ships[ship] || 0);
      });
    });

    const allShips = visibleShips.reduce(
      (sum, ship) => sum + totals[ship],
      0
    );

    return { totals, allShips };
  })();

  if (!loggedIn) {
    
    if (!module) {
  return (
    <main style={styles.page}>
      <div style={styles.card}>
        <h2>Select Module</h2>

        <button onClick={() => setModule("product")} style={styles.primaryButton}>
          Product Dashboard
        </button>

        <button onClick={() => setModule("equipment")} style={styles.primaryButton}>
          Equipment
        </button>
      </div>
    </main>
  );
}

if (module === "equipment" && !equipmentMode) {
  return (
    <main style={styles.page}>
      <div style={styles.card}>
        <h2>Equipment Options</h2>

        <button onClick={() => setEquipmentMode("muster")} style={styles.primaryButton}>
          Equipment Muster List
        </button>

        <button onClick={() => setEquipmentMode("inventory")} style={styles.primaryButton}>
          Equipment Inventory
        </button>
      </div>
    </main>
  );
}

const uploadMusterFile = (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  readExcelFile(file, (workbook) => {
    const rows = workbookToRows(workbook);
    setMusterRows(rows);
  });
};

if (module === "equipment" && equipmentMode === "muster") {
  const grouped = {};

  musterRows.slice(1).forEach((row) => {
    const category = row[2];
    const code = row[3];
    const name = row[4];
    const img = row[7];

    if (!category || !name) return;

    if (!grouped[category]) grouped[category] = [];
    grouped[category].push({ code, name, img });
  });

  return (
    <main style={styles.page}>
      <div style={styles.card}>
        <h2>Equipment Muster List</h2>

        <input type="file" onChange={uploadMusterFile} />

        {Object.entries(grouped).map(([cat, items]) => (
          <div key={cat}>
            <h3>{cat}</h3>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 12 }}>
              {items.map((i) => (
                <div key={i.code} style={styles.venueCard}>
                  <img src={i.img} style={{ width: "100%", height: 120, objectFit: "cover" }} />
                  <strong>{i.name}</strong>
                  <div>{i.code}</div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </main>
  );
}
    return (
      <div style={{ padding: 40 }}>
        <img src="/virgin-logo.png" style={{ height: 60 }} />
        <h2>Select Ship</h2>

        <select
          onChange={(e) => setUserShip(e.target.value)}
          value={userShip}
        >
          <option value="">Choose ship</option>
          {SHIPS.map((s) => (
            <option key={s}>{s}</option>
          ))}
        </select>

        <button onClick={() => setLoggedIn(true)}>Continue</button>
      </div>
    );
  }

  if (!module) {
    return (
      <div style={{ padding: 40 }}>
        <h2>Select Module</h2>

        <button onClick={() => setModule("product")}>
          Product Dashboard
        </button>

        <button onClick={() => setModule("equipment")}>
          Equipment
        </button>
      </div>
    );
  }

  if (module === "equipment" && !equipmentMode) {
    return (
      <div style={{ padding: 40 }}>
        <h2>Equipment Options</h2>

        <button onClick={() => setEquipmentMode("muster")}>
          Equipment Muster List
        </button>

        <button onClick={() => setEquipmentMode("inventory")}>
          Equipment Inventory
        </button>
      </div>
    );
  }

  if (module === "equipment" && equipmentMode === "muster") {
    const grouped = {};

    musterRows.slice(1).forEach((row) => {
      const category = row[2];
      const code = row[3];
      const name = row[4];
      const img = row[7];

      if (!category || !name) return;

      if (!grouped[category]) grouped[category] = [];
      grouped[category].push({ code, name, img });
    });

    return (
      <div style={{ padding: 20 }}>
        <h2>Equipment Muster List</h2>

        <input type="file" onChange={uploadMusterFile} />

        {Object.entries(grouped).map(([cat, items]) => (
          <div key={cat}>
            <h3>{cat}</h3>

            <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 10 }}>
              {items.map((i) => (
                <div key={i.code} style={{ border: "1px solid #ccc", padding: 10 }}>
                  <img src={i.img} style={{ width: "100%", height: 100, objectFit: "cover" }} />
                  <div>{i.code}</div>
                  <div>{i.name}</div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    );
  }

  
  return (
    <div style={{ padding: 20 }}>
      <h2>Product Dashboard</h2>

      <input type="file" onChange={uploadConsumptionFile} />
      <input type="file" onChange={uploadRecipeFile} />

      <input
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        placeholder="Search product"
      />

      {products
        .filter((p) =>
          p.toLowerCase().includes(search.toLowerCase())
        )
        .map((p) => (
          <div key={p} onClick={() => setSelectedProduct(p)}>
            {p}
          </div>
        ))}

      {selectedProduct && (
        <>
          <h3>{selectedProduct}</h3>

          {combinedBreakdown.map((v) => (
            <div key={v.venueKey}>
              <b>{v.displayName}</b>

              {visibleShips.map((s) => (
                <span
                  key={s}
                  style={{
                    margin: 5,
                    color: v.missingShips.includes(s)
                      ? "red"
                      : v.missingFromTemplate
                      ? "blue"
                      : "black",
                  }}
                >
                  {s}:{formatQty(v.ships[s])}
                </span>
              ))}
            </div>
          ))}
        </>
      )}
    </div>
  );
}

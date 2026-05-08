"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SHIPS = ["BRL", "RL", "SC", "VL"];

const ALLERGEN_RULES = [
  { allergen: "Tree Nuts", keywords: ["almond", "walnut", "pecan", "cashew", "hazelnut", "pistachio", "macadamia"] },
  { allergen: "Peanuts", keywords: ["peanut"] },
  {
    allergen: "Seeds",
    keywords: ["seed", "seeds", "sunflower seed", "pumpkin seed", "chia", "flax", "hemp seed"],
    exclude: ["seedless", "seedless cucumber"]
  },
  { allergen: "Soy", keywords: ["soy", "tofu", "edamame", "miso", "tamari"] },
  { allergen: "Gluten", keywords: ["wheat", "flour", "gluten", "bread", "pasta", "semolina", "barley", "rye", "panko"] },
  { allergen: "Milk / Dairy", keywords: ["milk", "cream", "butter", "cheese", "yogurt", "parmesan", "mozzarella", "ricotta", "cream cheese"] },
  { allergen: "Egg", keywords: ["egg", "eggs", "mayonnaise", "aioli"], exclude: ["eggplant"] },
  { allergen: "Fish", keywords: ["salmon", "tuna", "cod", "anchovy", "fish", "sardine"] },
  {
    allergen: "Shellfish",
    keywords: ["shrimp", "crab", "lobster", "mussel", "oyster", "scallop"],
    exclude: ["clam shell", "clamshell", "packed in a clam shell"]
  },
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


export default function App() {const formatQty = (value) => Number(value || 0).toFixed(2);

const getImageUrl = (url) => {
  const value = String(url || "").trim();
  if (!value) return "";

  if (value.includes("sharepoint.com") || value.includes("1drv.ms")) {
    return value.includes("?")
      ? `${value}&download=1`
      : `${value}?download=1`;
  }

  return value;
};
                               
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

  const parseMusterWorkbook = (workbook) => {
    const items = [];

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      rows.slice(1).forEach((row) => {
        const category = String(row[2] || "").trim(); // C
        const code = String(row[3] || "").trim(); // D
        const name = String(row[4] || "").trim(); // E
        const image = String(row[7] || "").trim(); // H

        if (!category || !name) return;

        items.push({
          sheetName,
          category,
          code,
          name,
          image,
        });
      });
    });

    return items;
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

const uploadWarehouseFile = (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  readExcelFile(file, (workbook) => {
    const rows = workbookToRows(workbook);
    setWarehouseRows(rows);
    setWarehouseMessage("Warehouse inventory loaded.");
  });
};        
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
        result[venueKey] = {
          displayName: venue,
          ships: {},
        };
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

    return Object.values(recipes).map((recipe) => ({
      ...recipe,
      venues: [...recipe.venues],
    }));
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
          <p style={styles.subtitle}>Product and equipment tools</p>

          <label style={styles.label}>🚢 Select your ship</label>
          <select value={userShip} onChange={(e) => setUserShip(e.target.value)} style={styles.select}>
            <option value="">Choose ship</option>
            {SHIPS.map((ship) => (
              <option key={ship}>{ship}</option>
            ))}
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
              <span>Grouped by all sheets and sub categories with code, name and image</span>
            </button>

            <button style={styles.moduleCard} onClick={() => setEquipmentMode("inventory")}>
              <div style={styles.moduleIcon}>📊</div>
              <strong>Equipment Inventory</strong>
              <span>Coming next</span>
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
          <p style={styles.emptyText}>This section is ready to build next.</p>
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
              <div>Column C = Sub Category</div>
              <div>Column D = Code</div>
              <div>Column E = Name</div>
              <div>Column H = Picture Link</div>
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

            <p style={styles.emptyText}>
              Click any equipment card to open the picture and full details.
            </p>
          </div>
        </section>

        <section style={styles.card}>
          <h2 style={styles.productTitle}>📋 Equipment Muster List</h2>

          {musterItems.length === 0 && (
            <p style={styles.emptyText}>Upload the Equipment Muster List file to begin.</p>
          )}

          {musterItems.length > 0 && totalItems === 0 && (
            <p style={styles.emptyText}>No equipment found for this search.</p>
          )}

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

        <a
      href={item.image}
      target="_blank"
      rel="noreferrer"
      style={styles.imageLink}
    >
      Open Picture
    </a>
  </div>
) : (
  <div style={styles.equipmentNoImage}>No image</div>
)}
                    ) : (
                      <div style={styles.equipmentNoImage}>No image</div>
                    )

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
                <button style={styles.closeButton} onClick={() => setSelectedEquipment(null)}>
                  ✕
                </button>

                <h2>{selectedEquipment.name}</h2>
                <p><strong>Code:</strong> {selectedEquipment.code || "N/A"}</p>
                <p><strong>Sheet:</strong> {selectedEquipment.sheetName || "N/A"}</p>
                <p><strong>Category:</strong> {selectedEquipment.category || "N/A"}</p>

              {selectedEquipment.image ? (
  <div>
    <img
      src={selectedEquipment.image}
      alt={selectedEquipment.name}
      style={styles.modalImage}
      onError={(e) => {
        e.currentTarget.style.display = "none";
        const link = e.currentTarget.nextElementSibling;
        if (link) link.style.display = "block";
      }}
    />

    <a
      href={selectedEquipment.image}
      target="_blank"
      rel="noreferrer"
      style={styles.imageLink}
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
        <button
          onClick={() => setViewMode("single")}
          style={{
            ...styles.viewModeButton,
            ...(viewMode === "single" ? styles.viewModeButtonActive : {}),
          }}
        >
          🚢 {userShip} Only
        </button>

        <button
          onClick={() => setViewMode("all")}
          style={{
            ...styles.viewModeButton,
            ...(viewMode === "all" ? styles.viewModeButtonActive : {}),
          }}
        >
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
              <div
                key={i}
                style={{
                  ...styles.venueCard,
                  ...(venueItem.missingShips.length > 0 ? styles.venueCardWarning : {}),
                  ...(venueItem.missingFromTemplate ? styles.venueCardTemplateWarning : {}),
                }}
              >
                <h4 style={styles.venueTitle}>
                  {venueItem.displayName}
                  <span style={styles.badgeGroup}>
                    {venueItem.missingFromTemplate && (
                      <span style={styles.templateBadge}>Missing Template</span>
                    )}
                    {venueItem.missingShips.length > 0 && (
                      <span style={styles.missingBadge}>
                        Missing: {venueItem.missingShips.join(", ")}
                      </span>
                    )}
                  </span>
                </h4>

                {venueItem.templateMatches.length > 0 && (
                  <div style={styles.templateFound}>
                    Template/Menu: {venueItem.templateMatches.join(", ")}
                  </div>
                )}

                {venueItem.missingFromTemplate && (
                  <div style={styles.templateWarningText}>
                    Product is used in recipe/location file for this venue but is not found in any template. Product has to be used.
                  </div>
                )}

                <div style={styles.shipGrid}>
                  {visibleShips.map((ship) => {
                    const isMissing = venueItem.required && (venueItem.ships[ship] || 0) === 0;

                    return (
                      <div
                        key={ship}
                        style={{
                          ...styles.shipBox,
                          ...(ship === userShip ? styles.shipBoxActive : {}),
                          ...(isMissing ? styles.shipBoxMissing : {}),
                        }}
                      >
                        <span style={styles.shipName}>{ship}</span>
                        <strong style={styles.shipQty}>{formatQty(venueItem.ships[ship])}</strong>
                      </div>
                    );
                  })}
                </div>

                {venueItem.missingShips.length > 0 && (
                  <div style={styles.warningSmall}>
                    Product appears in recipe/location file for this venue, but usage is 0 for highlighted ship(s).
                  </div>
                )}
              </div>
            ))}
          </div>

          <h3 style={styles.sectionTitle}>👨‍🍳 Recipes using this product</h3>

          {recipeRows.length === 0 && (
            <p style={styles.emptyText}>Upload the recipe/location file to see recipe details.</p>
          )}

          {recipeRows.length > 0 && recipesForProduct.length === 0 && (
            <p style={styles.emptyText}>No recipes found for this product.</p>
          )}

          <div style={styles.recipeList}>
            {recipesForProduct.map((recipe, i) => (
              <button
                key={i}
                onClick={() => setSelectedRecipe(recipe)}
                style={{
                  ...styles.recipeCard,
                  ...(selectedRecipe?.key === recipe.key ? styles.recipeCardActive : {}),
                }}
              >
                <div style={styles.recipeName}>{recipe.recipeName}</div>
                <div style={styles.recipeMeta}>Code: {recipe.recipeCode}</div>
                <div style={styles.recipeMeta}>
                  Venues: {recipe.venues.length ? recipe.venues.join(", ") : "N/A"}
                </div>
              </button>
            ))}
          </div>

          {selectedRecipe && (
            <div style={styles.ingredientsCard}>
              <h3 style={styles.sectionTitle}>🧾 Products used in recipe</h3>
              <h4 style={{ marginTop: 0 }}>
                {selectedRecipe.recipeName} ({selectedRecipe.recipeCode})
              </h4>

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
                            {subIngredients.map((subItem, j) => (
                              <li key={j}>{subItem}</li>
                            ))}
                          </ul>
                        )}
                      </li>
                    );
                  })}
                </ul>
              )}

              <h3 style={styles.sectionTitle}>⚠️ Rule-Based Allergen Warning</h3>
              <p style={styles.warningText}>
                This is a keyword-based warning only. Verify against official allergen data before use.
              </p>

              {allergenWarnings.length === 0 ? (
                <p style={styles.emptyText}>No likely allergens detected by keyword rules.</p>
              ) : (
                <div style={styles.allergenList}>
                  {allergenWarnings.map((item, i) => (
                    <div key={i} style={styles.allergenCard}>
                      <strong>{item.allergen}</strong>
                      <ul>
                        {item.products.map((product, j) => (
                          <li key={j}>{product}</li>
                        ))}
                      </ul>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </section>
      )}
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
      <button
        onClick={() => setViewMode("single")}
        style={{
          ...styles.viewModeButton,
          ...(viewMode === "single" ? styles.viewModeButtonActive : {}),
        }}
      >
        🚢 {userShip} Only
      </button>

      <button
        onClick={() => setViewMode("all")}
        style={{
          ...styles.viewModeButton,
          ...(viewMode === "all" ? styles.viewModeButtonActive : {}),
        }}
      >
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
        </div>
      </div>

      <div style={styles.card}>
        <h2 style={styles.cardTitle}>🔍 Select Product</h2>

        <input
          placeholder="Search product..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={styles.searchInput}
        />

        <div style={styles.productList}>
          {products.map((product, i) => (
            <button
              key={i}
              onClick={() => setSelectedProduct(product)}
              style={styles.productItem}
            >
              {product}
            </button>
          ))}
        </div>
      </div>
    </section>
  </main>
);
)

const styles = {
  page: {
    minHeight: "100vh",
    padding: 24,
    background: "#f5f5f5",
    fontFamily: "Arial, sans-serif",
    color: "#111",
  },
  loginCard: {
    maxWidth: 460,
    margin: "80px auto",
    padding: 28,
    background: "#fff",
    borderRadius: 16,
    boxShadow: "0 10px 30px rgba(0,0,0,0.08)",
    display: "grid",
    gap: 14,
  },
  logo: { height: 70, objectFit: "contain", marginBottom: 8 },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: 18,
    background: "#fff",
    borderRadius: 16,
    boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
    marginBottom: 20,
  },
  headerLogo: { height: 54, objectFit: "contain" },
  headerActions: {
    display: "flex",
    alignItems: "center",
    gap: 10,
  },
  backButton: {
    padding: "10px 14px",
    borderRadius: 999,
    border: "1px solid #ccc",
    background: "#fff",
    cursor: "pointer",
    fontWeight: "bold",
  },
  title: { margin: 0, fontSize: 28 },
  subtitle: { margin: 0, color: "#666" },
  label: { fontWeight: "bold", marginTop: 8 },
  select: { padding: 10, borderRadius: 8, border: "1px solid #ccc" },
  primaryButton: {
    marginTop: 10,
    padding: 12,
    borderRadius: 10,
    border: 0,
    background: "#111",
    color: "#fff",
    fontWeight: "bold",
    cursor: "pointer",
  },
  shipBadge: {
    padding: "10px 14px",
    borderRadius: 999,
    background: "#111",
    color: "#fff",
    fontWeight: "bold",
  },
  moduleGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
    gap: 16,
  },
  moduleCard: {
    border: "1px solid #ddd",
    background: "#fafafa",
    borderRadius: 16,
    padding: 20,
    cursor: "pointer",
    textAlign: "left",
    display: "grid",
    gap: 8,
    boxShadow: "0 4px 12px rgba(0,0,0,0.04)",
  },
  moduleIcon: {
    fontSize: 34,
  },
  viewModeBox: {
    display: "flex",
    gap: 10,
    marginBottom: 20,
  },
  viewModeButton: {
    padding: "10px 14px",
    borderRadius: 999,
    border: "1px solid #ccc",
    background: "#fff",
    cursor: "pointer",
    fontWeight: "bold",
  },
  viewModeButtonActive: {
    background: "#111",
    color: "#fff",
    borderColor: "#111",
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "1fr 1.4fr",
    gap: 20,
    marginBottom: 20,
  },
  card: {
    background: "#fff",
    borderRadius: 16,
    padding: 20,
    boxShadow: "0 4px 18px rgba(0,0,0,0.06)",
  },
  cardTitle: { marginTop: 0 },
  fileInput: { display: "block", margin: "8px 0 16px" },
  message: { color: "#555", fontSize: 14 },
  infoBox: {
    marginTop: 12,
    padding: 12,
    borderRadius: 12,
    background: "#f2f2f2",
    display: "grid",
    gap: 6,
  },
  searchInput: {
    width: "100%",
    padding: 12,
    borderRadius: 10,
    border: "1px solid #ccc",
    marginBottom: 10,
  },
  productList: {
    maxHeight: 300,
    overflowY: "auto",
    border: "1px solid #ddd",
    borderRadius: 12,
  },
  productItem: {
    width: "100%",
    display: "block",
    textAlign: "left",
    padding: 10,
    border: 0,
    borderBottom: "1px solid #eee",
    background: "#fff",
    cursor: "pointer",
  },
  productItemActive: { background: "#eee", fontWeight: "bold" },
  productTitle: { marginTop: 0, fontSize: 24 },
  sectionTitle: { marginTop: 22 },
  totalBox: {
    background: "#111",
    color: "#fff",
    borderRadius: 14,
    padding: 16,
    marginBottom: 18,
  },
  totalMain: { fontSize: 20, fontWeight: "bold" },
  totalShipGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))",
    gap: 10,
    marginTop: 12,
  },
  totalShipBox: {
    background: "#fff",
    color: "#111",
    borderRadius: 10,
    padding: 10,
    textAlign: "center",
    display: "grid",
    gap: 4,
  },
  venueGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))",
    gap: 14,
  },
  venueCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 14,
    background: "#fafafa",
  },
  venueCardWarning: {
    border: "2px solid #b00020",
    background: "#fff0f0",
  },
  venueCardTemplateWarning: {
    border: "2px solid #0057b8",
    background: "#eef5ff",
  },
  venueTitle: {
    marginTop: 0,
    display: "flex",
    justifyContent: "space-between",
    gap: 10,
    alignItems: "center",
  },
  badgeGroup: {
    display: "flex",
    flexWrap: "wrap",
    gap: 6,
    justifyContent: "flex-end",
  },
  missingBadge: {
    fontSize: 12,
    color: "#fff",
    background: "#b00020",
    borderRadius: 999,
    padding: "4px 8px",
  },
  templateBadge: {
    fontSize: 12,
    color: "#fff",
    background: "#0057b8",
    borderRadius: 999,
    padding: "4px 8px",
  },
  templateFound: {
    color: "#0057b8",
    fontSize: 13,
    fontWeight: "bold",
    marginBottom: 10,
  },
  templateWarningText: {
    color: "#0057b8",
    fontSize: 13,
    fontWeight: "bold",
    marginBottom: 10,
  },
  shipGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(70px, 1fr))",
    gap: 6,
  },
  shipBox: {
    minWidth: 0,
    padding: "8px 4px",
    borderRadius: 10,
    background: "#fff",
    border: "1px solid #ddd",
    display: "grid",
    gap: 3,
    textAlign: "center",
    overflow: "hidden",
  },
  shipBoxActive: { background: "#111", color: "#fff" },
  shipBoxMissing: {
    background: "#b00020",
    color: "#fff",
    borderColor: "#b00020",
  },
  shipName: { fontSize: 11, opacity: 0.8 },
  shipQty: {
    fontSize: 14,
    lineHeight: 1.1,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  warningSmall: {
    marginTop: 10,
    color: "#b00020",
    fontSize: 13,
    fontWeight: "bold",
  },
  emptyText: { color: "#777" },
  recipeList: { display: "grid", gap: 10 },
  recipeCard: {
    width: "100%",
    textAlign: "left",
    border: "1px solid #ddd",
    borderRadius: 12,
    padding: 12,
    background: "#fafafa",
    cursor: "pointer",
  },
  recipeCardActive: { background: "#eee", borderColor: "#111" },
  recipeName: { fontWeight: "bold" },
  recipeMeta: { color: "#555", fontSize: 14, marginTop: 4 },
  ingredientsCard: {
    marginTop: 18,
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 14,
    background: "#fafafa",
  },
  subRecipeList: {
    marginTop: 6,
    marginBottom: 8,
    paddingLeft: 24,
    color: "#333",
    background: "#f2f2f2",
    borderRadius: 8,
    paddingTop: 8,
    paddingBottom: 8,
  },
  warningText: {
    color: "#8a5a00",
    background: "#fff4d6",
    padding: 10,
    borderRadius: 8,
  },
  allergenList: { display: "grid", gap: 10 },
  allergenCard: {
    border: "1px solid #e1c16e",
    background: "#fff9e8",
    borderRadius: 10,
    padding: 10,
  },
  equipmentCategory: {
    marginBottom: 24,
  },
  equipmentGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
    gap: 14,
  },
  equipmentCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 14,
    background: "#fafafa",
    display: "grid",
    gap: 8,
    cursor: "pointer",
    textAlign: "left",
  },
  equipmentImage: {
    width: "100%",
    height: 150,
    objectFit: "cover",
    borderRadius: 10,
    background: "#eee",
  },
  equipmentNoImage: {
    height: 150,
    borderRadius: 10,
    background: "#eee",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#777",
  },
  modalBackdrop: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.55)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 9999,
    padding: 20,
  },
  modalCard: {
    background: "#fff",
    borderRadius: 18,
    padding: 22,
    maxWidth: 760,
    width: "100%",
    maxHeight: "90vh",
    overflowY: "auto",
    position: "relative",
    boxShadow: "0 20px 60px rgba(0,0,0,0.25)",
  },
  modalImage: {
    width: "100%",
    maxHeight: "65vh",
    objectFit: "contain",
    borderRadius: 14,
    background: "#f2f2f2",
  },
  closeButton: {
    position: "absolute",
    top: 12,
    right: 12,
    border: 0,
    background: "#111",
    color: "#fff",
    borderRadius: 999,
    width: 34,
    height: 34,
    cursor: "pointer",
    fontWeight: "bold",
    },
    imageLink: {
    display: "none",
    marginTop: 8,
    padding: 10,
    borderRadius: 10,
    background: "#111",
    color: "#fff",
    textAlign: "center",
    textDecoration: "none",
    fontWeight: "bold",
  },
};

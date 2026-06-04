"use client";

import React, {
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useState,
} from "react";
import * as XLSX from "xlsx";
import { downloadIngredientByLocationFileFromStorage } from "../../lib/permanentFiles";

const ALL_LOCATIONS_KEY = "__ALL_RESTAURANTS__";
const UNKNOWN_LOCATION = "Unknown Location";

const ALLERGEN_RULES = [
  {
    allergen: "Tree Nuts",
    keywords: [
      "almond",
      "walnut",
      "pecan",
      "cashew",
      "hazelnut",
      "pistachio",
      "macadamia",
      "pine nut",
      "brazil nut",
      "nutella",
      "praline",
    ],
  },
  {
    allergen: "Peanuts",
    keywords: ["peanut", "groundnut"],
  },
  {
    allergen: "Seeds",
    keywords: [
      "seed",
      "seeds",
      "sunflower seed",
      "pumpkin seed",
      "chia",
      "flax",
      "hemp seed",
      "poppy seed",
    ],
    exclude: ["seedless", "seedless cucumber"],
  },
  {
    allergen: "Soy",
    keywords: ["soy", "soya", "tofu", "edamame", "miso", "tamari", "soybean"],
  },
  {
    allergen: "Gluten",
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
      "farro",
      "couscous",
      "bulgur",
      "breadcrumbs",
    ],
  },
  {
    allergen: "Milk / Dairy",
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
      "mascarpone",
      "ghee",
      "whey",
      "casein",
      "lactose",
    ],
  },
  {
    allergen: "Egg",
    keywords: ["egg", "eggs", "mayonnaise", "mayo", "aioli", "meringue"],
    exclude: ["eggplant"],
  },
  {
    allergen: "Fish",
    keywords: [
      "salmon",
      "tuna",
      "cod",
      "anchovy",
      "anchovies",
      "fish",
      "sardine",
      "sardines",
      "seabass",
      "sea bass",
      "snapper",
      "halibut",
      "trout",
    ],
  },
  {
    allergen: "Shellfish",
    keywords: [
      "shrimp",
      "prawn",
      "crab",
      "lobster",
      "mussel",
      "oyster",
      "scallop",
      "clam",
      "clams",
      "shellfish",
    ],
    exclude: ["clam shell", "clamshell", "packed in a clam shell"],
  },
  {
    allergen: "Sesame",
    keywords: ["sesame", "tahini"],
  },
  {
    allergen: "Mustard",
    keywords: ["mustard"],
  },
];

const ALLERGEN_NAMES = ALLERGEN_RULES.map((rule) => rule.allergen);

const DEFAULT_VISIBLE_COUNT = 160;
const VISIBLE_STEP = 160;

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const cleanSearchText = (value) =>
  String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();

const normalizeDisplayText = (value) =>
  String(value || "")
    .replace(/\s+/g, " ")
    .trim();

const getLocationKey = (value) => cleanText(value || UNKNOWN_LOCATION);

const makeRecipeKey = ({ recipeCode, recipeName }) => {
  const code = normalizeDisplayText(recipeCode);
  const name = normalizeDisplayText(recipeName);

  if (code || name) {
    return `${cleanText(code)}__${cleanText(name || code)}`;
  }

  return "";
};

const getWorkbookRows = (workbook) => {
  const sheetName = workbook?.SheetNames?.[0];
  const worksheet = sheetName ? workbook.Sheets[sheetName] : null;

  if (!worksheet) {
    return {
      sheetName: "",
      rows: [],
    };
  }

  return {
    sheetName,
    rows: XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
    }),
  };
};

const findHeaderInfo = (rows) => {
  let headerRowIndex = 0;
  let headerRow = rows[0] || [];

  rows.slice(0, 30).some((row, rowIndex) => {
    const cleanRow = row.map((cell) => cleanText(cell));

    const hasRecipe =
      cleanRow.some((cell) => cell.includes("RECIPE")) ||
      cleanRow.some((cell) => cell.includes("MENU ITEM"));

    const hasIngredient =
      cleanRow.some((cell) => cell.includes("INGREDIENT")) ||
      cleanRow.some((cell) => cell.includes("PRODUCT"));

    const hasLocation =
      cleanRow.some((cell) => cell.includes("LOCATION")) ||
      cleanRow.some((cell) => cell.includes("VENUE")) ||
      cleanRow.some((cell) => cell.includes("OUTLET")) ||
      cleanRow.some((cell) => cell.includes("RESTAURANT"));

    if ((hasRecipe && hasIngredient) || (hasLocation && hasIngredient)) {
      headerRowIndex = rowIndex;
      headerRow = row;
      return true;
    }

    return false;
  });

  const cleanHeaders = headerRow.map((cell) => cleanText(cell));

  const findIndex = (patterns, fallback) => {
    const found = cleanHeaders.findIndex((header) =>
      patterns.some(
        (pattern) => header === pattern || header.includes(pattern)
      )
    );

    return found >= 0 ? found : fallback;
  };

  return {
    headerRowIndex,

    venueIndex: findIndex(
      ["VENUE", "LOCATION", "OUTLET", "RESTAURANT"],
      1
    ),

    productIndex: findIndex(
      [
        "ASSIGNED PRODUCT",
        "PRODUCT NAME",
        "PRODUCT",
        "INGREDIENT NAME",
        "INGREDIENT",
        "ITEM NAME",
      ],
      12
    ),

    backupProductIndex: 7,

    recipeCodeIndex: findIndex(
      ["RECIPE CODE", "MENU ITEM CODE", "ITEM CODE"],
      15
    ),

    recipeNameIndex: findIndex(
      ["RECIPE NAME", "MENU ITEM", "MENU ITEM NAME", "DISH", "DISH NAME"],
      16
    ),
  };
};

const ruleMatchesProduct = (rule, product) => {
  const lowerProduct = String(product || "").toLowerCase();

  if (!lowerProduct) return false;

  const excluded = (rule.exclude || []).some((word) =>
    lowerProduct.includes(String(word || "").toLowerCase())
  );

  if (excluded) return false;

  return (rule.keywords || []).some((keyword) =>
    lowerProduct.includes(String(keyword || "").toLowerCase())
  );
};

const detectAllergensForRecipe = ({ ingredients, subIngredientMap }) => {
  const allergenMap = new Map();

  const addMatch = ({ allergen, displayName }) => {
    if (!allergenMap.has(allergen)) {
      allergenMap.set(allergen, new Set());
    }

    allergenMap.get(allergen).add(displayName);
  };

  const checkProduct = ({ product, displayName }) => {
    ALLERGEN_RULES.forEach((rule) => {
      if (ruleMatchesProduct(rule, product)) {
        addMatch({
          allergen: rule.allergen,
          displayName: displayName || product,
        });
      }
    });
  };

  ingredients.forEach((ingredient) => {
    const cleanIngredient = normalizeDisplayText(ingredient);
    if (!cleanIngredient) return;

    checkProduct({
      product: cleanIngredient,
      displayName: cleanIngredient,
    });

    const subIngredients =
      subIngredientMap.get(cleanText(cleanIngredient)) || [];

    subIngredients.forEach((subIngredient) => {
      const cleanSubIngredient = normalizeDisplayText(subIngredient);
      if (!cleanSubIngredient) return;

      checkProduct({
        product: cleanSubIngredient,
        displayName: `${cleanIngredient} → ${cleanSubIngredient}`,
      });
    });
  });

  return Array.from(allergenMap.entries())
    .map(([allergen, products]) => ({
      allergen,
      products: Array.from(products).sort((a, b) => a.localeCompare(b)),
    }))
    .sort((a, b) => a.allergen.localeCompare(b.allergen));
};

const buildAllergenRecipeIndex = (rows) => {
  const safeRows = Array.isArray(rows) ? rows : [];
  const headerInfo = findHeaderInfo(safeRows);
  const dataRows = safeRows.slice(headerInfo.headerRowIndex + 1);

  const recipeMap = new Map();
  const subIngredientMap = new Map();

  // First pass: build sub-recipe ingredient lookup.
  // If an ingredient name is also a recipe name somewhere else,
  // its ingredients will show under that ingredient in the popup.
  dataRows.forEach((row) => {
    const recipeName = normalizeDisplayText(row[headerInfo.recipeNameIndex]);

    const ingredient = normalizeDisplayText(
      row[headerInfo.productIndex] || row[headerInfo.backupProductIndex]
    );

    if (!recipeName || !ingredient) return;
    if (cleanText(recipeName) === cleanText(ingredient)) return;

    const recipeNameKey = cleanText(recipeName);

    if (!subIngredientMap.has(recipeNameKey)) {
      subIngredientMap.set(recipeNameKey, new Set());
    }

    subIngredientMap.get(recipeNameKey).add(ingredient);
  });

  const normalizedSubIngredientMap = new Map();

  subIngredientMap.forEach((items, key) => {
    normalizedSubIngredientMap.set(
      key,
      Array.from(items).sort((a, b) => a.localeCompare(b))
    );
  });

  // Second pass: group recipes and collect every restaurant/location
  // where that recipe appears.
  dataRows.forEach((row, index) => {
    const recipeCode = normalizeDisplayText(row[headerInfo.recipeCodeIndex]);

    const recipeName = normalizeDisplayText(row[headerInfo.recipeNameIndex]);

    const venue =
      normalizeDisplayText(row[headerInfo.venueIndex]) || UNKNOWN_LOCATION;

    const ingredient = normalizeDisplayText(
      row[headerInfo.productIndex] || row[headerInfo.backupProductIndex]
    );

    if (!recipeCode && !recipeName) return;
    if (!ingredient) return;

    const cleanRecipeName = cleanText(recipeName);

    if (
      cleanRecipeName === "RECIPE NAME" ||
      cleanRecipeName === "MENU ITEM" ||
      cleanRecipeName === "MENU ITEM NAME" ||
      cleanRecipeName === "DISH NAME"
    ) {
      return;
    }

    const key = makeRecipeKey({
      recipeCode,
      recipeName,
    });

    if (!key) return;

    if (!recipeMap.has(key)) {
      recipeMap.set(key, {
        key,
        recipeCode: recipeCode || "N/A",
        recipeName: recipeName || recipeCode || "Unnamed Recipe",
        venues: new Set(),
        venueKeys: new Set(),
        ingredients: new Set(),
        sourceRows: [],
      });
    }

    const recipe = recipeMap.get(key);

    if (venue) {
      recipe.venues.add(venue);
      recipe.venueKeys.add(getLocationKey(venue));
    }

    if (ingredient) recipe.ingredients.add(ingredient);

    recipe.sourceRows.push(index + headerInfo.headerRowIndex + 2);
  });

  return Array.from(recipeMap.values())
    .map((recipe) => {
      const ingredients = Array.from(recipe.ingredients).sort((a, b) =>
        a.localeCompare(b)
      );

      const ingredientDetails = ingredients.map((ingredient) => ({
        ingredient,
        subIngredients:
          normalizedSubIngredientMap.get(cleanText(ingredient)) || [],
      }));

      const allergenWarnings = detectAllergensForRecipe({
        ingredients,
        subIngredientMap: normalizedSubIngredientMap,
      });

      const allergenSet = new Set(
        allergenWarnings.map((warning) => warning.allergen)
      );

      const venues = Array.from(recipe.venues).sort((a, b) =>
        a.localeCompare(b)
      );

      const venueKeys = Array.from(recipe.venueKeys);

      const subIngredientSearchText = ingredientDetails
        .flatMap((item) => item.subIngredients)
        .join(" ");

      const searchText = cleanSearchText(
        [
          recipe.recipeCode,
          recipe.recipeName,
          venues.join(" "),
          ingredients.join(" "),
          subIngredientSearchText,
          allergenWarnings
            .map(
              (warning) =>
                warning.allergen + " " + warning.products.join(" ")
            )
            .join(" "),
        ].join(" ")
      );

      return {
        ...recipe,
        venues,
        venueKeys,
        ingredients,
        ingredientDetails,
        allergenWarnings,
        allergenSet,
        allergenCount: allergenWarnings.length,
        matchedIngredientCount: allergenWarnings.reduce(
          (sum, warning) => sum + warning.products.length,
          0
        ),
        searchText,
      };
    })
    .sort((a, b) => {
      const venueDiff = (a.venues[0] || "").localeCompare(b.venues[0] || "");
      if (venueDiff !== 0) return venueDiff;

      return a.recipeName.localeCompare(b.recipeName);
    });
};

const getRecipeHasSelectedAllergen = (recipe, selectedAllergens) => {
  if (!selectedAllergens.length) return false;

  return selectedAllergens.some((allergen) =>
    recipe.allergenSet.has(allergen)
  );
};

const getRecipeIsSafeForSelectedAllergens = (recipe, selectedAllergens) => {
  if (!selectedAllergens.length) return false;

  return !getRecipeHasSelectedAllergen(recipe, selectedAllergens);
};

const getRecipeMatchesLocation = (recipe, selectedLocationKey) => {
  if (!selectedLocationKey) return false;
  if (selectedLocationKey === ALL_LOCATIONS_KEY) return true;

  return (recipe.venueKeys || []).includes(selectedLocationKey);
};

const badgeStyle = {
  display: "inline-flex",
  alignItems: "center",
  gap: 5,
  padding: "5px 8px",
  borderRadius: 999,
  fontSize: 12,
  fontWeight: "bold",
  border: "1px solid #ddd",
  background: "#fff",
};

const allergenBadgeStyle = {
  ...badgeStyle,
  background: "#fff4d6",
  color: "#8a5a00",
  border: "1px solid #e1c16e",
};

const safeBadgeStyle = {
  ...badgeStyle,
  background: "#e8f5e9",
  color: "#2e7d32",
  border: "1px solid #9ccc9c",
};

function AllergenChips({ styles, selectedAllergens, onToggle, onClear }) {
  return (
    <div style={{ display: "grid", gap: 10 }}>
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: 8,
        }}
      >
        {ALLERGEN_NAMES.map((allergen) => {
          const active = selectedAllergens.includes(allergen);

          return (
            <button
              key={allergen}
              type="button"
              style={{
                ...(styles.viewModeButton || {}),
                ...(active ? styles.viewModeButtonActive || {} : {}),
                marginBottom: 0,
              }}
              onClick={() => onToggle(allergen)}
            >
              {active ? "✓ " : ""}
              {allergen}
            </button>
          );
        })}
      </div>

      {selectedAllergens.length > 0 && (
        <button
          type="button"
          style={{
            ...(styles.backButton || {}),
            justifySelf: "start",
          }}
          onClick={onClear}
        >
          Clear selected allergens
        </button>
      )}
    </div>
  );
}

function RecipeAllergenBadges({ recipe }) {
  if (!recipe.allergenWarnings.length) {
    return <div style={safeBadgeStyle}>✅ No keyword allergens detected</div>;
  }

  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
      {recipe.allergenWarnings.map((warning) => (
        <span key={warning.allergen} style={allergenBadgeStyle}>
          ⚠️ {warning.allergen}
        </span>
      ))}
    </div>
  );
}

function LocationCard({
  styles,
  location,
  selectedAllergens,
  onOpen,
}) {
  const isAll = location.locationKey === ALL_LOCATIONS_KEY;

  return (
    <button
      type="button"
      style={{
        ...(styles.equipmentCard || {}),
        width: "100%",
        cursor: "pointer",
        textAlign: "left",
        ...(isAll ? styles.countedCard || {} : {}),
      }}
      onClick={() => onOpen(location)}
    >
      <div
        style={{
          fontSize: 42,
          lineHeight: 1,
          textAlign: "center",
          marginBottom: 4,
        }}
      >
        {isAll ? "🌍" : "🍽️"}
      </div>

      <div style={styles.recipeName}>{location.locationName}</div>

      <div style={styles.recipeMeta}>
        Type: {isAll ? "All Restaurants" : "Restaurant / Location"}
      </div>

      <div style={styles.statusGood}>Dishes: {location.recipeCount}</div>

      <div style={styles.recipeMeta}>
        Dishes with warnings: {location.warningCount}
      </div>

      {selectedAllergens.length > 0 && (
        <>
          <div style={styles.recipeMeta}>
            Safe for selected filter: {location.safeCount}
          </div>

          <div style={styles.recipeMeta}>
            Contains selected allergen: {location.selectedRiskCount}
          </div>
        </>
      )}

      {location.sampleRecipes.length > 0 && (
        <div style={styles.recipeMeta}>
          Examples: {location.sampleRecipes.join(", ")}
        </div>
      )}

      <button
        type="button"
        style={styles.imageButton}
        onClick={(event) => {
          event.stopPropagation();
          onOpen(location);
        }}
      >
        Open Dishes
      </button>
    </button>
  );
}

function RecipeCard({ styles, recipe, selectedAllergens, onOpen }) {
  const selectedRisk = getRecipeHasSelectedAllergen(
    recipe,
    selectedAllergens
  );

  const safeForSelected = getRecipeIsSafeForSelectedAllergens(
    recipe,
    selectedAllergens
  );

  const matchedPreview = recipe.allergenWarnings
    .flatMap((warning) =>
      warning.products.slice(0, 2).map((product) => ({
        allergen: warning.allergen,
        product,
      }))
    )
    .slice(0, 4);

  const subRecipeCount = recipe.ingredientDetails.filter(
    (item) => item.subIngredients.length > 0
  ).length;

  return (
    <button
      type="button"
      style={{
        ...(styles.equipmentCard || {}),
        cursor: "pointer",
        ...(selectedRisk ? styles.orderWarningCard || {} : {}),
        ...(safeForSelected ? styles.countedCard || {} : {}),
      }}
      onClick={() => onOpen(recipe)}
    >
      <div style={styles.recipeName}>
        {recipe.recipeName || "Unnamed Recipe"}
      </div>

      <div style={styles.recipeMeta}>
        Code: {recipe.recipeCode || "N/A"}
      </div>

      <div style={styles.recipeMeta}>
        Restaurant / Location:{" "}
        {recipe.venues.length ? recipe.venues.join(", ") : "N/A"}
      </div>

      <div style={styles.recipeMeta}>
        Ingredients: {recipe.ingredients.length}
      </div>

      {subRecipeCount > 0 && (
        <div style={styles.recipeMeta}>
          Sub-recipes detected: {subRecipeCount}
        </div>
      )}

      <RecipeAllergenBadges recipe={recipe} />

      {selectedAllergens.length > 0 && safeForSelected && (
        <div style={styles.statusGood}>
          Safe for selected allergen filter
        </div>
      )}

      {selectedAllergens.length > 0 && selectedRisk && (
        <div style={styles.statusBad}>Contains selected allergen</div>
      )}

      {matchedPreview.length > 0 && (
        <div
          style={{
            display: "grid",
            gap: 4,
            marginTop: 6,
            textAlign: "left",
          }}
        >
          {matchedPreview.map((match, index) => (
            <div
              key={`${match.allergen}-${match.product}-${index}`}
              style={{
                fontSize: 12,
                color: "#555",
                lineHeight: 1.25,
              }}
            >
              <strong>{match.allergen}:</strong> {match.product}
            </div>
          ))}

          {recipe.matchedIngredientCount > matchedPreview.length && (
            <div style={styles.recipeMeta}>
              +{recipe.matchedIngredientCount - matchedPreview.length} more
              matched item(s)
            </div>
          )}
        </div>
      )}
    </button>
  );
}

function RecipeDetailModal({
  styles,
  recipe,
  selectedAllergens,
  onClose,
}) {
  if (!recipe) return null;

  const selectedRisk = getRecipeHasSelectedAllergen(
    recipe,
    selectedAllergens
  );

  const safeForSelected = getRecipeIsSafeForSelectedAllergens(
    recipe,
    selectedAllergens
  );

  return (
    <div style={styles.modalBackdrop} onClick={onClose}>
      <div
        style={{
          ...(styles.modalCard || {}),
          maxWidth: 980,
        }}
        onClick={(event) => event.stopPropagation()}
      >
        <button type="button" style={styles.closeButton} onClick={onClose}>
          ✕
        </button>

        <h2 style={{ marginTop: 0 }}>{recipe.recipeName}</h2>

        <div style={styles.infoBox}>
          <div>
            <strong>Recipe code:</strong> {recipe.recipeCode || "N/A"}
          </div>

          <div>
            <strong>Restaurant / Location:</strong>{" "}
            {recipe.venues.length ? recipe.venues.join(", ") : "N/A"}
          </div>

          <div>
            <strong>Ingredients:</strong> {recipe.ingredients.length}
          </div>

          <div>
            <strong>Sub-recipes detected:</strong>{" "}
            {
              recipe.ingredientDetails.filter(
                (item) => item.subIngredients.length > 0
              ).length
            }
          </div>

          <div>
            <strong>Allergen groups detected:</strong>{" "}
            {recipe.allergenWarnings.length}
          </div>
        </div>

        {selectedAllergens.length > 0 && (
          <div
            style={{
              ...(safeForSelected ? styles.statusGood : styles.statusBad),
              marginTop: 14,
            }}
          >
            {safeForSelected
              ? "Safe for selected allergens: " + selectedAllergens.join(", ")
              : selectedRisk
              ? "Not safe for selected allergens: " +
                selectedAllergens.join(", ")
              : "No selected allergen conflict detected."}
          </div>
        )}

        <h3 style={styles.sectionTitle}>⚠️ Allergens by matched item</h3>

        {!recipe.allergenWarnings.length ? (
          <p style={styles.emptyText}>
            No likely allergens were detected by keyword rules.
          </p>
        ) : (
          <div style={styles.allergenList}>
            {recipe.allergenWarnings.map((warning) => {
              const selected = selectedAllergens.includes(warning.allergen);

              return (
                <div
                  key={warning.allergen}
                  style={{
                    ...(styles.allergenCard || {}),
                    ...(selected
                      ? {
                          border: "2px solid #b00020",
                          background: "#fff0f0",
                        }
                      : {}),
                  }}
                >
                  <strong>
                    {selected ? "🚫 " : "⚠️ "}
                    {warning.allergen}
                  </strong>

                  <ul>
                    {warning.products.map((product, index) => (
                      <li key={`${warning.allergen}-${product}-${index}`}>
                        {product}
                      </li>
                    ))}
                  </ul>
                </div>
              );
            })}
          </div>
        )}

        <h3 style={styles.sectionTitle}>🧾 Ingredients and sub-recipes</h3>

        {recipe.ingredientDetails.length === 0 ? (
          <p style={styles.emptyText}>No ingredients found for this recipe.</p>
        ) : (
          <ul>
            {recipe.ingredientDetails.map((item, index) => (
              <li key={`${item.ingredient}-${index}`} style={{ marginBottom: 10 }}>
                <strong>{item.ingredient}</strong>

                {item.subIngredients.length > 0 && (
                  <ul style={styles.subRecipeList}>
                    {item.subIngredients.map((subItem, subIndex) => (
                      <li key={`${item.ingredient}-${subItem}-${subIndex}`}>
                        {subItem}
                      </li>
                    ))}
                  </ul>
                )}
              </li>
            ))}
          </ul>
        )}

        <p style={styles.warningText}>
          This is a keyword-based allergen warning. Always verify with the
          official recipe, supplier specification, and onboard allergen process
          before service.
        </p>
      </div>
    </div>
  );
}

function MatrixTable({
  styles,
  recipes,
  selectedAllergens,
  onOpenRecipe,
}) {
  return (
    <div
      style={{
        overflowX: "auto",
        border: "1px solid #ddd",
        borderRadius: 14,
        background: "#fff",
      }}
    >
      <table
        style={{
          width: "100%",
          borderCollapse: "collapse",
          minWidth: 1100,
          fontSize: 13,
        }}
      >
        <thead>
          <tr>
            <th
              style={{
                textAlign: "left",
                padding: 10,
                borderBottom: "1px solid #ddd",
                background: "#111",
                color: "#fff",
                position: "sticky",
                left: 0,
                zIndex: 2,
              }}
            >
              Recipe
            </th>

            <th
              style={{
                textAlign: "left",
                padding: 10,
                borderBottom: "1px solid #ddd",
                background: "#111",
                color: "#fff",
              }}
            >
              Restaurant / Location
            </th>

            {ALLERGEN_NAMES.map((allergen) => {
              const selected = selectedAllergens.includes(allergen);

              return (
                <th
                  key={allergen}
                  style={{
                    padding: 10,
                    borderBottom: "1px solid #ddd",
                    background: selected ? "#b00020" : "#111",
                    color: "#fff",
                    textAlign: "center",
                    whiteSpace: "nowrap",
                  }}
                >
                  {allergen}
                </th>
              );
            })}
          </tr>
        </thead>

        <tbody>
          {recipes.map((recipe) => (
            <tr key={recipe.key}>
              <td
                style={{
                  padding: 10,
                  borderBottom: "1px solid #eee",
                  background: "#fff",
                  position: "sticky",
                  left: 0,
                  zIndex: 1,
                  minWidth: 260,
                }}
              >
                <button
                  type="button"
                  style={{
                    border: 0,
                    background: "transparent",
                    padding: 0,
                    textAlign: "left",
                    cursor: "pointer",
                    fontWeight: "bold",
                    color: "#111",
                  }}
                  onClick={() => onOpenRecipe(recipe)}
                >
                  {recipe.recipeName}
                </button>

                <div style={styles.recipeMeta}>
                  Code: {recipe.recipeCode || "N/A"}
                </div>
              </td>

              <td
                style={{
                  padding: 10,
                  borderBottom: "1px solid #eee",
                  color: "#555",
                  minWidth: 220,
                }}
              >
                {recipe.venues.join(", ") || "N/A"}
              </td>

              {ALLERGEN_NAMES.map((allergen) => {
                const hasAllergen = recipe.allergenSet.has(allergen);
                const selected = selectedAllergens.includes(allergen);

                return (
                  <td
                    key={`${recipe.key}-${allergen}`}
                    style={{
                      padding: 10,
                      borderBottom: "1px solid #eee",
                      textAlign: "center",
                      background:
                        hasAllergen && selected
                          ? "#fff0f0"
                          : hasAllergen
                          ? "#fff4d6"
                          : "#fff",
                      color: hasAllergen ? "#8a5a00" : "#2e7d32",
                      fontWeight: "bold",
                    }}
                  >
                    {hasAllergen ? "⚠️" : "✓"}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function AllergenModule({
  styles = {},
  supabase,
  userShip,
  userEmail,
  isAdmin,
  onBack,
  logUsageEvent,
}) {
  const [recipeRows, setRecipeRows] = useState([]);
  const [fileName, setFileName] = useState("");
  const [sourceSheet, setSourceSheet] = useState("");
  const [message, setMessage] = useState("");
  const [loading, setLoading] = useState(false);

  const [search, setSearch] = useState("");
  const deferredSearch = useDeferredValue(search);

  const [viewMode, setViewMode] = useState("cards");
  const [selectedAllergens, setSelectedAllergens] = useState([]);
  const [selectedRecipe, setSelectedRecipe] = useState(null);
  const [visibleCount, setVisibleCount] = useState(DEFAULT_VISIBLE_COUNT);

  const [selectedLocationKey, setSelectedLocationKey] = useState("");
  const [selectedLocationName, setSelectedLocationName] = useState("");

  const parseArrayBuffer = useCallback(async ({ arrayBuffer, name }) => {
    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
    });

    const { sheetName, rows } = getWorkbookRows(workbook);

    setRecipeRows(rows);
    setFileName(name || "Ingredient by Location");
    setSourceSheet(sheetName || "");
    setVisibleCount(DEFAULT_VISIBLE_COUNT);
    setSelectedLocationKey("");
    setSelectedLocationName("");
    setSelectedRecipe(null);

    return {
      sheetName,
      rows,
    };
  }, []);

  const loadPermanentFile = useCallback(async () => {
    if (!supabase) {
      setMessage(
        "Supabase is not connected. Upload Ingredient by Location manually."
      );
      return;
    }

    setLoading(true);
    setMessage("Loading permanent Ingredient by Location file...");

    try {
      const arrayBuffer =
        await downloadIngredientByLocationFileFromStorage({ supabase });

      const parsed = await parseArrayBuffer({
        arrayBuffer,
        name: "Permanent Ingredient by Location",
      });

      setMessage(
        "Permanent Ingredient by Location loaded. " +
          Math.max(parsed.rows.length - 1, 0) +
          " row(s)."
      );

      if (typeof logUsageEvent === "function") {
        logUsageEvent("allergen_permanent_file_loaded", {
          module: "allergen",
          ship: userShip,
          userEmail,
          rowCount: Math.max(parsed.rows.length - 1, 0),
          sourceSheet: parsed.sheetName,
        });
      }
    } catch (error) {
      setRecipeRows([]);
      setFileName("");
      setSourceSheet("");
      setSelectedLocationKey("");
      setSelectedLocationName("");
      setMessage(
        error?.message ||
          "Could not load permanent Ingredient by Location file. Upload manually."
      );
    } finally {
      setLoading(false);
    }
  }, [supabase, parseArrayBuffer, logUsageEvent, userShip, userEmail]);

  useEffect(() => {
    loadPermanentFile();
  }, [loadPermanentFile]);

  useEffect(() => {
    setVisibleCount(DEFAULT_VISIBLE_COUNT);
  }, [
    deferredSearch,
    viewMode,
    selectedAllergens.join("|"),
    selectedLocationKey,
  ]);

  const uploadIngredientFile = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setMessage("Reading Ingredient by Location file...");

    try {
      const arrayBuffer = await file.arrayBuffer();

      const parsed = await parseArrayBuffer({
        arrayBuffer,
        name: file.name,
      });

      setMessage(
        "Ingredient by Location loaded. " +
          Math.max(parsed.rows.length - 1, 0) +
          " row(s)."
      );

      if (typeof logUsageEvent === "function") {
        logUsageEvent("allergen_file_uploaded", {
          module: "allergen",
          ship: userShip,
          userEmail,
          fileName: file.name,
          rowCount: Math.max(parsed.rows.length - 1, 0),
          sourceSheet: parsed.sheetName,
        });
      }
    } catch (error) {
      setMessage(error?.message || "Could not read the uploaded file.");
      window.alert(error?.message || "Could not read the uploaded file.");
    } finally {
      setLoading(false);
      event.target.value = "";
    }
  };

  const recipeIndex = useMemo(
    () => buildAllergenRecipeIndex(recipeRows),
    [recipeRows]
  );

  const locationCards = useMemo(() => {
    const map = new Map();

    recipeIndex.forEach((recipe) => {
      const venues = recipe.venues.length ? recipe.venues : [UNKNOWN_LOCATION];

      venues.forEach((venue) => {
        const locationKey = getLocationKey(venue);

        if (!map.has(locationKey)) {
          map.set(locationKey, {
            locationKey,
            locationName: venue || UNKNOWN_LOCATION,
            recipes: new Map(),
            warningCount: 0,
            safeCount: 0,
            selectedRiskCount: 0,
            searchText: venue || UNKNOWN_LOCATION,
          });
        }

        const location = map.get(locationKey);

        if (!location.recipes.has(recipe.key)) {
          location.recipes.set(recipe.key, recipe);

          if (recipe.allergenWarnings.length > 0) {
            location.warningCount += 1;
          }

          if (
            selectedAllergens.length > 0 &&
            getRecipeIsSafeForSelectedAllergens(recipe, selectedAllergens)
          ) {
            location.safeCount += 1;
          }

          if (
            selectedAllergens.length > 0 &&
            getRecipeHasSelectedAllergen(recipe, selectedAllergens)
          ) {
            location.selectedRiskCount += 1;
          }

          location.searchText += " " + recipe.searchText;
        }
      });
    });

    const actualLocationCards = Array.from(map.values())
      .map((location) => ({
        ...location,
        recipeCount: location.recipes.size,
        sampleRecipes: Array.from(location.recipes.values())
          .slice(0, 3)
          .map((recipe) => recipe.recipeName),
      }))
      .sort((a, b) => a.locationName.localeCompare(b.locationName));

    const allCard = {
      locationKey: ALL_LOCATIONS_KEY,
      locationName: "All Restaurants",
      recipes: new Map(recipeIndex.map((recipe) => [recipe.key, recipe])),
      recipeCount: recipeIndex.length,
      warningCount: recipeIndex.filter(
        (recipe) => recipe.allergenWarnings.length > 0
      ).length,
      safeCount: selectedAllergens.length
        ? recipeIndex.filter((recipe) =>
            getRecipeIsSafeForSelectedAllergens(recipe, selectedAllergens)
          ).length
        : 0,
      selectedRiskCount: selectedAllergens.length
        ? recipeIndex.filter((recipe) =>
            getRecipeHasSelectedAllergen(recipe, selectedAllergens)
          ).length
        : 0,
      sampleRecipes: recipeIndex.slice(0, 3).map((recipe) => recipe.recipeName),
      searchText: "all restaurants all locations all venues",
    };

    return recipeIndex.length ? [allCard, ...actualLocationCards] : [];
  }, [recipeIndex, selectedAllergens]);

  const availableLocationKeys = useMemo(
    () => new Set(locationCards.map((location) => location.locationKey)),
    [locationCards]
  );

  useEffect(() => {
    if (!selectedLocationKey) return;

    if (!availableLocationKeys.has(selectedLocationKey)) {
      setSelectedLocationKey("");
      setSelectedLocationName("");
    }
  }, [selectedLocationKey, availableLocationKeys]);

  const allergenSummary = useMemo(() => {
    const summary = new Map();

    ALLERGEN_NAMES.forEach((allergen) => {
      summary.set(allergen, {
        allergen,
        recipes: 0,
        matchedItems: 0,
      });
    });

    recipeIndex.forEach((recipe) => {
      recipe.allergenWarnings.forEach((warning) => {
        const current = summary.get(warning.allergen);

        if (!current) return;

        current.recipes += 1;
        current.matchedItems += warning.products.length;
      });
    });

    return Array.from(summary.values());
  }, [recipeIndex]);

  const activeLocationRecipes = useMemo(() => {
    if (!selectedLocationKey) return [];

    return recipeIndex.filter((recipe) =>
      getRecipeMatchesLocation(recipe, selectedLocationKey)
    );
  }, [recipeIndex, selectedLocationKey]);

  const filteredLocationCards = useMemo(() => {
    const query = cleanSearchText(deferredSearch);

    return locationCards.filter((location) => {
      if (!query) return true;

      return cleanSearchText(location.searchText).includes(query);
    });
  }, [locationCards, deferredSearch]);

  const filteredRecipes = useMemo(() => {
    const query = cleanSearchText(deferredSearch);

    return activeLocationRecipes.filter((recipe) => {
      if (query && !recipe.searchText.includes(query)) {
        return false;
      }

      if (viewMode === "safe") {
        if (!selectedAllergens.length) return true;
        return getRecipeIsSafeForSelectedAllergens(recipe, selectedAllergens);
      }

      if (viewMode === "warnings") {
        return recipe.allergenWarnings.length > 0;
      }

      return true;
    });
  }, [
    activeLocationRecipes,
    deferredSearch,
    viewMode,
    selectedAllergens,
  ]);

  const visibleRecipes = useMemo(
    () => filteredRecipes.slice(0, visibleCount),
    [filteredRecipes, visibleCount]
  );

  const safeScopeRecipes = selectedLocationKey
    ? activeLocationRecipes
    : recipeIndex;

  const safeRecipeCount = useMemo(() => {
    if (!selectedAllergens.length) return 0;

    return safeScopeRecipes.filter((recipe) =>
      getRecipeIsSafeForSelectedAllergens(recipe, selectedAllergens)
    ).length;
  }, [safeScopeRecipes, selectedAllergens]);

  const selectedRiskCount = useMemo(() => {
    if (!selectedAllergens.length) return 0;

    return safeScopeRecipes.filter((recipe) =>
      getRecipeHasSelectedAllergen(recipe, selectedAllergens)
    ).length;
  }, [safeScopeRecipes, selectedAllergens]);

  const recipesWithWarnings = useMemo(
    () => recipeIndex.filter((recipe) => recipe.allergenWarnings.length > 0),
    [recipeIndex]
  );

  const recipesWithoutWarnings = recipeIndex.length - recipesWithWarnings.length;

  const activeLocationLabel = selectedLocationName || "Choose restaurant";

  const toggleAllergen = (allergen) => {
    setSelectedAllergens((current) =>
      current.includes(allergen)
        ? current.filter((item) => item !== allergen)
        : [...current, allergen]
    );
  };

  const openLocation = (location) => {
    setSelectedLocationKey(location.locationKey);
    setSelectedLocationName(location.locationName);
    setSearch("");
    setSelectedRecipe(null);
    setVisibleCount(DEFAULT_VISIBLE_COUNT);

    if (typeof logUsageEvent === "function") {
      logUsageEvent("allergen_location_opened", {
        module: "allergen",
        ship: userShip,
        userEmail,
        locationName: location.locationName,
        recipeCount: location.recipeCount,
      });
    }
  };

  const backToLocations = () => {
    setSelectedLocationKey("");
    setSelectedLocationName("");
    setSearch("");
    setSelectedRecipe(null);
    setVisibleCount(DEFAULT_VISIBLE_COUNT);
  };

  const openRecipe = (recipe) => {
    setSelectedRecipe(recipe);

    if (typeof logUsageEvent === "function") {
      logUsageEvent("allergen_recipe_opened", {
        module: "allergen",
        ship: userShip,
        userEmail,
        locationName: activeLocationLabel,
        recipeCode: recipe.recipeCode,
        recipeName: recipe.recipeName,
        allergens: recipe.allergenWarnings.map((item) => item.allergen),
      });
    }
  };

  const viewButtonStyle = (active) => ({
    ...(styles.viewModeButton || {}),
    ...(active ? styles.viewModeButtonActive || {} : {}),
  });

  const searchPlaceholder = selectedLocationKey
    ? "Search dish, recipe code, ingredient, sub-recipe, allergen..."
    : "Search restaurant / location, dish, ingredient, allergen...";

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={styles.headerLogo}
        />

        <div style={styles.headerActions}>
          <button type="button" style={styles.backButton} onClick={onBack}>
            ← Back
          </button>

          <div style={styles.shipBadge}>🚢 {userShip || "Ship"}</div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🧬 Allergen Matrix</h2>

          <p style={styles.emptyText}>
            Choose restaurant / location first, then open the dishes used there.
            Recipe popup shows ingredients and sub-recipes.
          </p>

          <button
            type="button"
            style={styles.backButton}
            onClick={loadPermanentFile}
            disabled={loading}
          >
            {loading
              ? "Loading..."
              : "🔄 Reload Permanent Ingredient by Location"}
          </button>

          <label style={styles.label}>Upload Ingredient by Location file</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm"
            onChange={uploadIngredientFile}
            style={styles.fileInput}
            disabled={loading}
          />

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.infoBox}>
            <div>
              📄 File: <strong>{fileName || "Not loaded"}</strong>
            </div>

            <div>
              📑 Sheet: <strong>{sourceSheet || "N/A"}</strong>
            </div>

            <div>
              🍽️ Restaurants / locations:{" "}
              <strong>
                {locationCards.length ? locationCards.length - 1 : 0}
              </strong>
            </div>

            <div>
              📋 Recipes indexed: <strong>{recipeIndex.length}</strong>
            </div>

            <div>
              ⚠️ Recipes with allergen warnings:{" "}
              <strong>{recipesWithWarnings.length}</strong>
            </div>

            <div>
              ✅ Recipes without keyword warnings:{" "}
              <strong>{recipesWithoutWarnings}</strong>
            </div>

            <div>
              👤 User: <strong>{userEmail || "N/A"}</strong>
              {isAdmin ? " / Admin" : ""}
            </div>
          </div>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>✅ Safe Dish Finder</h2>

          <p style={styles.emptyText}>
            Choose allergens to avoid. Safe dishes will be filtered inside the
            selected restaurant / location.
          </p>

          <AllergenChips
            styles={styles}
            selectedAllergens={selectedAllergens}
            onToggle={toggleAllergen}
            onClear={() => setSelectedAllergens([])}
          />

          <div style={styles.infoBox}>
            <div>
              📍 Scope: <strong>{activeLocationLabel}</strong>
            </div>

            <div>
              🚫 Avoiding:{" "}
              <strong>
                {selectedAllergens.length
                  ? selectedAllergens.join(", ")
                  : "No allergen selected"}
              </strong>
            </div>

            <div>
              ✅ Safe dishes for selected filter:{" "}
              <strong>
                {selectedAllergens.length ? safeRecipeCount : "Select allergen"}
              </strong>
            </div>

            <div>
              ⚠️ Dishes with selected allergen:{" "}
              <strong>
                {selectedAllergens.length
                  ? selectedRiskCount
                  : "Select allergen"}
              </strong>
            </div>
          </div>
        </div>
      </section>

      <section style={styles.card}>
        <div
          style={{
            ...(styles.header || {}),
            boxShadow: "none",
            padding: 0,
            marginBottom: 16,
          }}
        >
          <div>
            <h2 style={styles.productTitle}>
              {selectedLocationKey
                ? `🍽️ ${activeLocationLabel}`
                : "🍽️ Choose Restaurant / Location"}
            </h2>

            <p style={{ ...(styles.emptyText || {}), margin: 0 }}>
              {selectedLocationKey
                ? "Showing dishes used in this restaurant / location."
                : "Select a restaurant / location card first."}
            </p>
          </div>

          <div style={styles.headerActions}>
            {selectedLocationKey && (
              <button
                type="button"
                style={styles.backButton}
                onClick={backToLocations}
              >
                ← Restaurants
              </button>
            )}

            <button
              type="button"
              style={viewButtonStyle(viewMode === "cards")}
              onClick={() => setViewMode("cards")}
            >
              Recipe Cards
            </button>

            <button
              type="button"
              style={viewButtonStyle(viewMode === "matrix")}
              onClick={() => setViewMode("matrix")}
            >
              Matrix
            </button>

            <button
              type="button"
              style={viewButtonStyle(viewMode === "safe")}
              onClick={() => setViewMode("safe")}
            >
              Safe Dishes
            </button>

            <button
              type="button"
              style={viewButtonStyle(viewMode === "warnings")}
              onClick={() => setViewMode("warnings")}
            >
              Warnings Only
            </button>
          </div>
        </div>

        <input
          placeholder={searchPlaceholder}
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          style={styles.searchInput}
        />

        <div style={styles.infoBox}>
          {!selectedLocationKey ? (
            <>
              <div>
                🍽️ Restaurants / locations shown:{" "}
                <strong>{filteredLocationCards.length}</strong>
              </div>

              <div>
                📋 Total restaurants / locations:{" "}
                <strong>
                  {locationCards.length ? locationCards.length - 1 : 0}
                </strong>
              </div>

              <div>
                Open a restaurant / location card to see its dishes.
              </div>
            </>
          ) : (
            <>
              <div>
                📍 Location: <strong>{activeLocationLabel}</strong>
              </div>

              <div>
                👀 Dishes showing: <strong>{visibleRecipes.length}</strong> /{" "}
                {filteredRecipes.length}
              </div>

              <div>
                📋 Dishes in location:{" "}
                <strong>{activeLocationRecipes.length}</strong>
              </div>

              <div>
                🧭 Current view:{" "}
                <strong>
                  {viewMode === "cards"
                    ? "Recipe Cards"
                    : viewMode === "matrix"
                    ? "Matrix"
                    : viewMode === "safe"
                    ? "Safe Dishes"
                    : "Warnings Only"}
                </strong>
              </div>
            </>
          )}
        </div>
      </section>

      {!selectedLocationKey && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>🍽️ Restaurants / Locations</h2>

          {recipeRows.length === 0 && (
            <p style={styles.emptyText}>
              Load or upload the Ingredient by Location file to begin.
            </p>
          )}

          {recipeRows.length > 0 && filteredLocationCards.length === 0 && (
            <p style={styles.emptyText}>
              No restaurant / location matched your search.
            </p>
          )}

          <div style={styles.equipmentGrid}>
            {filteredLocationCards.map((location) => (
              <LocationCard
                key={location.locationKey}
                styles={styles}
                location={location}
                selectedAllergens={selectedAllergens}
                onOpen={openLocation}
              />
            ))}
          </div>
        </section>
      )}

      <section style={styles.card}>
        <h2 style={styles.productTitle}>📊 Allergen Summary</h2>

        <div style={styles.equipmentGrid}>
          {allergenSummary.map((item) => (
            <button
              key={item.allergen}
              type="button"
              style={{
                ...(styles.equipmentCard || {}),
                ...(selectedAllergens.includes(item.allergen)
                  ? styles.orderWarningCard || {}
                  : {}),
              }}
              onClick={() => toggleAllergen(item.allergen)}
            >
              <div style={styles.recipeName}>
                {selectedAllergens.includes(item.allergen) ? "🚫 " : "⚠️ "}
                {item.allergen}
              </div>

              <div style={styles.recipeMeta}>Recipes: {item.recipes}</div>

              <div style={styles.recipeMeta}>
                Matched ingredient items: {item.matchedItems}
              </div>

              <div style={item.recipes > 0 ? allergenBadgeStyle : safeBadgeStyle}>
                {item.recipes > 0 ? "Detected" : "No matches"}
              </div>
            </button>
          ))}
        </div>
      </section>

      {selectedLocationKey && (
        <section style={styles.card}>
          <h2 style={styles.productTitle}>
            {viewMode === "matrix"
              ? "🧬 Allergen Matrix"
              : viewMode === "safe"
              ? "✅ Safe Dish Results"
              : viewMode === "warnings"
              ? "⚠️ Recipes With Allergen Warnings"
              : "🍽️ Recipe Cards"}
          </h2>

          {viewMode === "safe" && !selectedAllergens.length && (
            <p style={styles.warningText}>
              Select one or more allergens in Safe Dish Finder to filter safe
              dishes.
            </p>
          )}

          {filteredRecipes.length === 0 && (
            <p style={styles.emptyText}>
              No dishes match the current search/filter in{" "}
              {activeLocationLabel}.
            </p>
          )}

          {viewMode === "matrix" ? (
            <MatrixTable
              styles={styles}
              recipes={visibleRecipes}
              selectedAllergens={selectedAllergens}
              onOpenRecipe={openRecipe}
            />
          ) : (
            <div style={styles.equipmentGrid}>
              {visibleRecipes.map((recipe) => (
                <RecipeCard
                  key={recipe.key}
                  styles={styles}
                  recipe={recipe}
                  selectedAllergens={selectedAllergens}
                  onOpen={openRecipe}
                />
              ))}
            </div>
          )}

          {filteredRecipes.length > visibleRecipes.length && (
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                marginTop: 18,
              }}
            >
              <button
                type="button"
                style={styles.primaryButton}
                onClick={() =>
                  setVisibleCount((current) => current + VISIBLE_STEP)
                }
              >
                Show More ({visibleRecipes.length} / {filteredRecipes.length})
              </button>
            </div>
          )}
        </section>
      )}

      <RecipeDetailModal
        styles={styles}
        recipe={selectedRecipe}
        selectedAllergens={selectedAllergens}
        onClose={() => setSelectedRecipe(null)}
      />
    </main>
  );
}

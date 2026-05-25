"use client";

import React, { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  downloadIngredientByLocationFileFromStorage,
  uploadIngredientByLocationFileToStorage,
} from "../../lib/permanentFiles";

const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

const safeText = (value) => String(value || "").replace(/\s+/g, " ").trim();

const normalizeCode = (value) => {
  const text = String(value ?? "").trim();
  if (!text) return "";

  const numberValue = Number(text);

  if (
    Number.isFinite(numberValue) &&
    String(Math.trunc(numberValue)) === text.replace(/\.0+$/, "")
  ) {
    return String(Math.trunc(numberValue));
  }

  return text.replace(/\.0+$/, "");
};

const getVenueIcon = (venue) => {
  const text = cleanText(venue);

  if (text.includes("RAZZLE")) return "🎪";
  if (text.includes("WAKE")) return "🌊";
  if (text.includes("PINK AGAVE")) return "🌮";
  if (text.includes("EXTRA VIRGIN")) return "🍝";
  if (text.includes("TEST KITCHEN")) return "🧪";
  if (text.includes("GALLEY")) return "🍽️";
  if (text.includes("DOCK")) return "⚓";
  if (text.includes("GUNBAE")) return "🥢";
  if (text.includes("PIZZA")) return "🍕";
  if (text.includes("PASTRY") || text.includes("BAKERY")) return "🥐";
  if (text.includes("BAR")) return "🍸";

  return "🏢";
};

const VENUE_NAME_COLORS = [
  "#4b5563",
  "#5b4b7a",
  "#6b4e16",
  "#0f5f5c",
  "#7a3e3e",
  "#315d7d",
  "#556b2f",
  "#6b4f5b",
  "#4f5f7a",
  "#5a5f37",
];

const getVenueNameColor = (venue) => {
  const text = cleanText(venue);

  if (!text) return VENUE_NAME_COLORS[0];

  let hash = 0;

  for (let index = 0; index < text.length; index += 1) {
    hash =
      (hash + text.charCodeAt(index) * (index + 1)) %
      VENUE_NAME_COLORS.length;
  }

  return VENUE_NAME_COLORS[hash];
};

const normalizeVenueName = (value) =>
  safeText(value)
    .replace(/\s*-\s*VV$/i, "")
    .replace(/\s+VV$/i, "")
    .replace(/\s+/g, " ")
    .trim();

const normalizeIngredientTextForMatch = (value) =>
  cleanText(value)
    .replace(/[^A-Z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const escapeRegex = (value) =>
  String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const textHasWordOrPhrase = (text, word) => {
  const source = normalizeIngredientTextForMatch(text);
  const target = normalizeIngredientTextForMatch(word);

  if (!source || !target) return false;

  return new RegExp(
    `(^|[^A-Z0-9])${escapeRegex(target)}([^A-Z0-9]|$)`
  ).test(source);
};

const IGNORED_BASIC_INGREDIENTS = [
  "WATER",
  "ICE",
  "SALT",
  "SEA SALT",
  "KOSHER SALT",
  "TABLE SALT",
  "SUGAR",
  "GRANULATED SUGAR",
  "CASTER SUGAR",
  "BROWN SUGAR",
  "ICING SUGAR",
  "POWDERED SUGAR",
  "BLACK PEPPER",
  "WHITE PEPPER",
  "PEPPER",
  "GROUND BLACK PEPPER",
];

const FRESH_HERB_OR_RAW_PRODUCE_WORDS = [
  "HERB",
  "HERBS",
  "MINT",
  "PARSLEY",
  "BASIL",
  "CILANTRO",
  "CORIANDER",
  "DILL",
  "THYME",
  "ROSEMARY",
  "SAGE",
  "OREGANO",
  "MARJORAM",
  "TARRAGON",
  "CHERVIL",
  "CHIVES",
  "CHIVE",
  "BAY LEAF",
  "BAY LEAVES",
  "LEMONGRASS",
  "LEMON GRASS",
  "KAFFIR LIME LEAF",
  "LIME LEAF",
  "LETTUCE",
  "SPINACH",
  "ARUGULA",
  "ROCKET",
  "KALE",
  "CABBAGE",
  "CARROT",
  "ONION",
  "GARLIC",
  "SHALLOT",
  "TOMATO",
  "CUCUMBER",
  "ZUCCHINI",
  "COURGETTE",
  "EGGPLANT",
  "AUBERGINE",
  "POTATO",
  "MUSHROOM",
  "CELERY",
  "APPLE",
  "ORANGE",
  "LEMON",
  "LIME",
  "BANANA",
  "BERRY",
  "BERRIES",
  "STRAWBERRY",
  "BLUEBERRY",
  "RASPBERRY",
  "BLACKBERRY",
];

const PROCESSED_WORDS_THAT_BLOCK_RAW_PRODUCE_SKIP = [
  "SAUCE",
  "DRESSING",
  "MARINADE",
  "PASTE",
  "PESTO",
  "MAYONNAISE",
  "AIOLI",
  "BATTER",
  "BREADING",
  "CRUMB",
  "CRUMBS",
  "STUFFING",
  "COOKIE",
  "CAKE",
  "BISCUIT",
  "CRACKER",
  "BREAD",
  "PASTRY",
  "GLAZE",
  "SPREAD",
];

const isIgnoredBasicIngredient = (value) => {
  const text = normalizeIngredientTextForMatch(value);
  if (!text) return false;

  if (IGNORED_BASIC_INGREDIENTS.includes(text)) return true;

  return [
    /^WATER( |$)/,
    /^(KOSHER |SEA |TABLE )?SALT( |$)/,
    /^SUGAR( |$)/,
    /^(BLACK |WHITE |GROUND )?PEPPER( |$)/,
  ].some((rule) => rule.test(text));
};

const isPlainFreshHerbOrRawProduce = (...values) => {
  const text = values
    .map((value) => normalizeIngredientTextForMatch(value))
    .join(" ");

  if (!text.trim()) return false;

  const hasProcessedWord = PROCESSED_WORDS_THAT_BLOCK_RAW_PRODUCE_SKIP.some(
    (word) => textHasWordOrPhrase(text, word)
  );

  if (hasProcessedWord) return false;

  return FRESH_HERB_OR_RAW_PRODUCE_WORDS.some((word) =>
    textHasWordOrPhrase(text, word)
  );
};

const shouldSkipPossibleAllergensForIngredient = (...values) => {
  const text = values
    .map((value) => normalizeIngredientTextForMatch(value))
    .join(" ");

  if (!text.trim()) return false;

  return isIgnoredBasicIngredient(text) || isPlainFreshHerbOrRawProduce(...values);
};

const ALLERGEN_DISPLAY = {
  "Gluten / Wheat": { icon: "🌾", color: "#8a5a00" },
  Milk: { icon: "🥛", color: "#0057b8" },
  Eggs: { icon: "🥚", color: "#7a4f00" },
  Peanuts: { icon: "🥜", color: "#b00020" },
  "Tree Nuts": { icon: "🌰", color: "#6b3f1d" },
  Soy: { icon: "🫘", color: "#2e7d32" },
  Sesame: { icon: "⚪", color: "#6a4a00" },
  Fish: { icon: "🐟", color: "#005f73" },
  "Crustacean Shellfish": { icon: "🦐", color: "#b00020" },
  Molluscs: { icon: "🦪", color: "#005f73" },
  Mustard: { icon: "🟡", color: "#8a5a00" },
  Celery: { icon: "🥬", color: "#2e7d32" },
  Lupin: { icon: "🌱", color: "#2e7d32" },
  Sulphites: { icon: "⚠️", color: "#8a5a00" },
};

const normalizeAllergenName = (value) => {
  const text = cleanText(value);

  if (!text) return "";

  if (
    [
      "N/A",
      "NA",
      "NONE",
      "NO",
      "NO ALLERGEN",
      "NO ALLERGENS",
      "NULL",
      "NIL",
      "-",
      "--",
      "0",
    ].includes(text)
  ) {
    return "";
  }

  if (
    text.includes("GLUTEN") ||
    text.includes("WHEAT") ||
    text.includes("BARLEY") ||
    text.includes("RYE") ||
    text.includes("OATS")
  ) {
    return "Gluten / Wheat";
  }

  if (text.includes("MILK") || text.includes("LACTOSE") || text.includes("DAIRY")) {
    return "Milk";
  }

  if (text.includes("EGG")) return "Eggs";
  if (text.includes("PEANUT")) return "Peanuts";

  if (
    text.includes("TREE NUT") ||
    text.includes("NUTS") ||
    text.includes("ALMOND") ||
    text.includes("CASHEW") ||
    text.includes("WALNUT") ||
    text.includes("PECAN") ||
    text.includes("PISTACHIO") ||
    text.includes("HAZELNUT") ||
    text.includes("MACADAMIA")
  ) {
    return "Tree Nuts";
  }

  if (text.includes("SOY")) return "Soy";
  if (text.includes("SESAME") || text.includes("TAHINI")) return "Sesame";

  if (
    text.includes("CRUSTACEAN") ||
    text.includes("SHRIMP") ||
    text.includes("PRAWN") ||
    text.includes("CRAB") ||
    text.includes("LOBSTER")
  ) {
    return "Crustacean Shellfish";
  }

  if (
    text.includes("MOLLUSC") ||
    text.includes("MOLLUSK") ||
    text.includes("CLAM") ||
    text.includes("MUSSEL") ||
    text.includes("OYSTER") ||
    text.includes("SCALLOP") ||
    text.includes("SQUID") ||
    text.includes("OCTOPUS")
  ) {
    return "Molluscs";
  }

  if (
    text.includes("FISH") ||
    text.includes("ANCHOV") ||
    text.includes("SALMON") ||
    text.includes("TUNA") ||
    text.includes("COD")
  ) {
    return "Fish";
  }

  if (text.includes("MUSTARD")) return "Mustard";
  if (text.includes("CELERY") || text.includes("CELERIAC")) return "Celery";
  if (text.includes("LUPIN")) return "Lupin";

  if (
    text.includes("SULPH") ||
    text.includes("SULF") ||
    text.includes("METABISULFITE")
  ) {
    return "Sulphites";
  }

  return safeText(value);
};

const splitAllergens = (...values) => {
  const found = new Set();

  values.forEach((value) => {
    String(value || "")
      .split(/[,;|/]+/g)
      .map((part) => normalizeAllergenName(part))
      .filter(Boolean)
      .forEach((item) => found.add(item));
  });

  return [...found].sort();
};

const KEYWORD_RULES = [
  {
    allergen: "Gluten / Wheat",
    words: [
      "wheat",
      "wheat flour",
      "all purpose flour",
      "ap flour",
      "bread flour",
      "cake flour",
      "semolina",
      "barley",
      "rye",
      "malt",
      "panko",
      "breadcrumb",
      "bread crumb",
      "breadcrumbs",
      "bread crumbs",
      "pasta",
      "noodle",
      "noodles",
      "couscous",
      "cracker",
      "crackers",
      "bread",
      "bun",
      "brioche",
      "croissant",
      "pastry",
      "phyllo",
      "filo",
      "tortilla",
      "wrap",
      "soy sauce",
    ],
  },
  {
    allergen: "Milk",
    words: [
      "milk",
      "cream",
      "butter",
      "cheese",
      "yogurt",
      "yoghurt",
      "parmesan",
      "mozzarella",
      "ricotta",
      "mascarpone",
      "whey",
      "casein",
      "lactose",
      "milk chocolate",
      "white chocolate",
    ],
  },
  {
    allergen: "Eggs",
    words: [
      "egg",
      "eggs",
      "mayonnaise",
      "mayo",
      "aioli",
      "meringue",
      "custard",
      "hollandaise",
    ],
  },
  {
    allergen: "Peanuts",
    words: ["peanut", "peanuts", "peanut butter", "peanutbutter", "satay"],
  },
  {
    allergen: "Tree Nuts",
    words: [
      "almond",
      "walnut",
      "pecan",
      "cashew",
      "hazelnut",
      "pistachio",
      "macadamia",
      "brazil nut",
      "pine nut",
      "marzipan",
      "praline",
      "gianduja",
      "nutella",
    ],
  },
  {
    allergen: "Soy",
    words: ["soy", "soya", "tofu", "edamame", "miso", "tamari", "soy sauce", "yuba"],
  },
  {
    allergen: "Sesame",
    words: ["sesame", "tahini", "benne", "gingelly"],
  },
  {
    allergen: "Fish",
    words: [
      "fish",
      "anchovy",
      "anchovies",
      "fish sauce",
      "salmon",
      "tuna",
      "cod",
      "sardine",
      "mackerel",
      "trout",
    ],
  },
  {
    allergen: "Crustacean Shellfish",
    words: ["shrimp", "prawn", "prawns", "crab", "lobster", "crayfish"],
  },
  {
    allergen: "Molluscs",
    words: [
      "clam",
      "clams",
      "mussel",
      "mussels",
      "oyster",
      "oysters",
      "scallop",
      "scallops",
      "squid",
      "octopus",
      "calamari",
    ],
  },
  {
    allergen: "Mustard",
    words: ["mustard", "dijon"],
  },
  {
    allergen: "Celery",
    words: ["celery", "celeriac"],
  },
  {
    allergen: "Lupin",
    words: ["lupin", "lupine"],
  },
  {
    allergen: "Sulphites",
    words: [
      "sulphite",
      "sulphites",
      "sulfite",
      "sulfites",
      "metabisulfite",
      "sulfur dioxide",
      "sulphur dioxide",
    ],
  },
];

const isProcessedOrPreparedItem = (value) => {
  const text = cleanText(value);

  return [
    "SAUCE",
    "DRESSING",
    "MARINADE",
    "PASTE",
    "MIX",
    "SEASONING",
    "COOKIE",
    "CAKE",
    "BISCUIT",
    "CHOCOLATE",
    "CHIPS",
    "CEREAL",
    "CRACKER",
    "BREAD",
    "PASTRY",
    "GLAZE",
    "PUREE",
    "SPREAD",
    "BUTTER",
    "PESTO",
    "MAYONNAISE",
    "AIOLI",
    "HOLLANDAISE",
    "TERIYAKI",
    "HOISIN",
    "WORCESTERSHIRE",
    "OYSTER SAUCE",
    "FISH SAUCE",
  ].some((word) => text.includes(word));
};

const PREMADE_COMMON_ALLERGEN_RULES = [
  {
    words: ["soy sauce", "teriyaki", "hoisin"],
    allergens: ["Soy", "Gluten / Wheat"],
  },
  {
    words: ["fish sauce"],
    allergens: ["Fish"],
  },
  {
    words: ["oyster sauce"],
    allergens: ["Molluscs", "Soy"],
  },
  {
    words: ["worcestershire"],
    allergens: ["Fish"],
  },
  {
    words: [
      "cookie",
      "cookies",
      "biscuit",
      "cake",
      "muffin",
      "brownie",
      "donut",
      "doughnut",
    ],
    allergens: ["Gluten / Wheat", "Milk", "Eggs"],
  },
  {
    words: [
      "bread",
      "bun",
      "brioche",
      "croissant",
      "pastry",
      "tart shell",
      "pie shell",
      "cracker",
    ],
    allergens: ["Gluten / Wheat"],
  },
  {
    words: ["milk chocolate", "white chocolate"],
    allergens: ["Milk"],
  },
  {
    words: ["peanut butter", "peanutbutter", "satay"],
    allergens: ["Peanuts"],
  },
  {
    words: ["pesto"],
    allergens: ["Milk", "Tree Nuts"],
  },
  {
    words: ["mayonnaise", "mayo", "aioli", "hollandaise"],
    allergens: ["Eggs", "Mustard"],
  },
  {
    words: ["granola", "muesli"],
    allergens: ["Gluten / Wheat", "Tree Nuts"],
  },
];

const getPreparedCommonAllergens = (...values) => {
  const text = values
    .map((value) => normalizeIngredientTextForMatch(value))
    .join(" ");

  const found = new Set();

  if (!text.trim()) return [];

  PREMADE_COMMON_ALLERGEN_RULES.forEach((rule) => {
    const matched = rule.words.some((word) => textHasWordOrPhrase(text, word));

    if (matched) {
      rule.allergens.forEach((allergen) => found.add(allergen));
    }
  });

  return [...found].sort();
};

const keywordAllergensForText = (...values) => {
  const text = values
    .map((value) => normalizeIngredientTextForMatch(value))
    .join(" ");

  const found = new Set();

  if (!text.trim()) return [];

  KEYWORD_RULES.forEach((rule) => {
    const matched = rule.words.some((word) => textHasWordOrPhrase(text, word));

    if (matched) found.add(rule.allergen);
  });

  // False-positive prevention.
  if (textHasWordOrPhrase(text, "eggplant") || textHasWordOrPhrase(text, "aubergine")) {
    found.delete("Eggs");
  }

  if (textHasWordOrPhrase(text, "seedless")) {
    found.delete("Sesame");
  }

  // Non-wheat flours should not become gluten/wheat by keyword.
  if (
    textHasWordOrPhrase(text, "rice flour") ||
    textHasWordOrPhrase(text, "corn flour") ||
    textHasWordOrPhrase(text, "chickpea flour") ||
    textHasWordOrPhrase(text, "gram flour") ||
    textHasWordOrPhrase(text, "potato flour") ||
    textHasWordOrPhrase(text, "tapioca flour") ||
    textHasWordOrPhrase(text, "almond flour") ||
    textHasWordOrPhrase(text, "coconut flour")
  ) {
    found.delete("Gluten / Wheat");
  }

  return [...found].sort();
};

const looksGlutenFreeClaim = (...values) => {
  const text = values.map((value) => cleanText(value)).join(" ");
  return (
    /\bGF\b/.test(text) ||
    text.includes("GLUTEN FREE") ||
    text.includes("GLUTEN-FREE")
  );
};

const getHeaderIndexes = (headers) => {
  const cleanHeaders = headers.map((header) => cleanText(header));
  const compactHeaders = cleanHeaders.map((header) =>
    header.replace(/[^A-Z0-9]/g, "")
  );

  const compactName = (value) => cleanText(value).replace(/[^A-Z0-9]/g, "");

  const findExact = (names, fallback = -1) => {
    const wanted = names.map(compactName).filter(Boolean);
    const index = compactHeaders.findIndex((header) => wanted.includes(header));
    return index >= 0 ? index : fallback;
  };

  const ingredientCodeIndex = findExact(
    ["Code", "IngredientCode", "Ingredient Code", "ProductCode", "Product Code"],
    6
  );

  const ingredientNameIndex = findExact(
    ["Name", "IngredientName", "Ingredient Name", "ProductName", "Product Name"],
    7
  );

  return {
    restaurantCode: findExact(["RestaurantCode", "Restaurant Code"], 0),
    restaurantName: findExact(["RestaurantName", "Restaurant Name"], 1),
    menuCode: findExact(["MenuCode", "Menu Code"], 2),
    menuName: findExact(["MenuName", "Menu Name"], 3),
    category: findExact(["Category"], 4),
    subCategory: findExact(["SubCategory", "Sub Category"], 5),

    // These must match the actual ingredient/product columns.
    ingredientCode: ingredientCodeIndex,
    ingredientName: ingredientNameIndex,

    assigned: findExact(["Assigned"], 12),
    assignedType: findExact(["AssignedType", "Assigned Type"], 13),
    recipeCode: findExact(["RecipeCode", "Recipe Code"], 15),
    recipeName: findExact(["RecipeName", "Recipe Name"], 16),
    specialInstructions: findExact(["SpecialInstructions", "Special Instructions"], 17),
    specialInstructions2: findExact(["SpecialInstructions2", "Special Instructions 2"], 18),
    recipeIsBasic: findExact(["RecipeIsBasic", "Recipe Is Basic"], 19),
    hasProductRelation: findExact(["HasProductRelation", "Has Product Relation"], 20),
    ingredientAllergens: findExact(["IngredientAllergens", "Ingredient Allergens"], 21),
    recipeAllergens: findExact(["RecipeAllergens", "Recipe Allergens"], 22),
  };
};

const parseIngredientByLocationWorkbook = (workbook) => {
  const sheetName =
    workbook.SheetNames.find((name) => cleanText(name).includes("SHEET")) ||
    workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    return { rows: [], venues: [], sourceSheet: sheetName || "" };
  }

  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
  });

  const headerRowIndex = rawRows.findIndex((row) => {
    const compactText = row
      .map((cell) => cleanText(cell).replace(/[^A-Z0-9]/g, ""))
      .join("|");

    return (
      compactText.includes("RESTAURANTNAME") &&
      compactText.includes("RECIPENAME") &&
      compactText.includes("INGREDIENTALLERGENS")
    );
  });

  if (headerRowIndex < 0) {
    throw new Error(
      "Could not find Ingredient by Location header row. Expected RestaurantName, RecipeName, Name, IngredientAllergens, RecipeAllergens."
    );
  }

  const headers = rawRows[headerRowIndex];
  const indexes = getHeaderIndexes(headers);
  const parsedRows = [];

  rawRows.slice(headerRowIndex + 1).forEach((row, rowOffset) => {
    const get = (index) => (index >= 0 ? row[index] : "");

    const venueName = normalizeVenueName(get(indexes.restaurantName));
    const menuName = safeText(get(indexes.menuName));
    const recipeName = safeText(get(indexes.recipeName));
    const ingredientName = safeText(get(indexes.ingredientName));

    if (!venueName || !recipeName || !ingredientName) return;

    const assigned = safeText(get(indexes.assigned));
    const ingredientAllergens = splitAllergens(get(indexes.ingredientAllergens));
    const recipeDeclaredAllergens = splitAllergens(get(indexes.recipeAllergens));

    // Important:
    // Ingredient cards should show only allergens that belong to this ingredient.
    // Do NOT merge recipeDeclaredAllergens into every ingredient line.
    // Example: a recipe may contain gluten from bread, but parsley/mint must not show gluten.
    const nameDetectedRealAllergens = keywordAllergensForText(
      ingredientName,
      assigned
    );

    const ignoredBasic = isIgnoredBasicIngredient(ingredientName);

    const plainRawIngredient = isPlainFreshHerbOrRawProduce(
      ingredientName,
      assigned
    );

    const skipPossibleAllergens = shouldSkipPossibleAllergensForIngredient(
      ingredientName,
      assigned
    );

    const processedItem =
      !skipPossibleAllergens &&
      (isProcessedOrPreparedItem(ingredientName) ||
        isProcessedOrPreparedItem(assigned));

    const preparedCommonAllergens = processedItem
      ? getPreparedCommonAllergens(ingredientName, assigned)
      : [];

    // Real allergens shown on the ingredient card.
    // These come from IngredientAllergens and clear ingredient-name detection.
    const explicitAllergens = ignoredBasic
      ? []
      : [...new Set([...ingredientAllergens, ...nameDetectedRealAllergens])].sort();

    const detectedAllergens = ignoredBasic
      ? []
      : [...new Set(nameDetectedRealAllergens)].sort();

    // Fresh herbs, raw produce, water, salt, sugar, pepper do not receive possible hidden allergens.
    const possibleHiddenAllergens =
      ignoredBasic || skipPossibleAllergens
        ? []
        : preparedCommonAllergens.filter(
            (allergen) => !explicitAllergens.includes(allergen)
          );

    const gfClaim = looksGlutenFreeClaim(
      menuName,
      recipeName,
      get(indexes.specialInstructions),
      get(indexes.specialInstructions2)
    );

    const hasIngredientGluten = explicitAllergens.includes("Gluten / Wheat");
    const hiddenWarnings = [];

    if (!ignoredBasic && !skipPossibleAllergens && gfClaim && hasIngredientGluten) {
      hiddenWarnings.push(
        "Ingredient gluten check: recipe/menu says GF or gluten free, but this ingredient shows gluten/wheat."
      );
    }

    if (!ignoredBasic && !skipPossibleAllergens && processedItem) {
      hiddenWarnings.push(
        "Prepared / pre-made item: read the supplier label for may-contain, cross-contact, and full allergen information."
      );
    }

    parsedRows.push({
      id: `${rowOffset}-${venueName}-${recipeName}-${ingredientName}`,
      sourceRow: headerRowIndex + 2 + rowOffset,
      restaurantCode: normalizeCode(get(indexes.restaurantCode)),
      restaurantName: venueName,
      menuCode: normalizeCode(get(indexes.menuCode)),
      menuName,
      category: safeText(get(indexes.category)),
      subCategory: safeText(get(indexes.subCategory)),
      ingredientCode: normalizeCode(get(indexes.ingredientCode)),
      ingredientName,
      assigned,
      assignedType: safeText(get(indexes.assignedType)),
      recipeCode: normalizeCode(get(indexes.recipeCode)),
      recipeName,
      specialInstructions: safeText(get(indexes.specialInstructions)),
      specialInstructions2: safeText(get(indexes.specialInstructions2)),
      recipeIsBasic: safeText(get(indexes.recipeIsBasic)),
      hasProductRelation: safeText(get(indexes.hasProductRelation)),
      ingredientAllergens,
      recipeAllergens: recipeDeclaredAllergens,
      recipeDeclaredAllergens,
      explicitAllergens,
      detectedAllergens,
      possibleHiddenAllergens,
      hiddenWarnings,
      ignoredBasic,
      plainRawIngredient,
      skipPossibleAllergens,
      processedItem,
      gfClaim,
    });
  });

  const venueMap = new Map();

  parsedRows.forEach((row) => {
    const venueKey = cleanText(row.restaurantName);

    if (!venueMap.has(venueKey)) {
      venueMap.set(venueKey, {
        venueKey,
        restaurantName: row.restaurantName,
        restaurantCode: row.restaurantCode,
        icon: getVenueIcon(row.restaurantName),
        recipesMap: new Map(),
        allergens: new Set(),
        possibleHidden: new Set(),
        hiddenWarningCount: 0,
        ingredientCount: 0,
      });
    }

    const venue = venueMap.get(venueKey);

    const recipeKey = cleanText(
      `${row.recipeCode}|${row.recipeName}|${row.menuCode}|${row.menuName}`
    );

    if (!venue.recipesMap.has(recipeKey)) {
      venue.recipesMap.set(recipeKey, {
        recipeKey,
        recipeCode: row.recipeCode,
        recipeName: row.recipeName,
        menuCode: row.menuCode,
        menuName: row.menuName,
        restaurantName: row.restaurantName,
        rows: [],
        ingredients: [],
        subRecipes: [],
        allergens: new Set(),
        possibleHidden: new Set(),
        hiddenWarnings: [],
        gfClaim: false,
      });
    }

    const recipe = venue.recipesMap.get(recipeKey);

    recipe.rows.push(row);
    recipe.gfClaim = recipe.gfClaim || row.gfClaim;

    if (
      row.assignedType === "R" ||
      row.hasProductRelation === "N" ||
      row.recipeIsBasic === "Y"
    ) {
      recipe.subRecipes.push(row);
    } else {
      recipe.ingredients.push(row);
    }

    if (!row.ignoredBasic) venue.ingredientCount += 1;

    // Ingredient-level real allergens.
    row.explicitAllergens.forEach((allergen) => {
      venue.allergens.add(allergen);
      recipe.allergens.add(allergen);
    });

    // Recipe-level declared allergens belong to recipe/venue summaries only.
    // They should not be copied onto every ingredient card.
    (row.recipeDeclaredAllergens || []).forEach((allergen) => {
      venue.allergens.add(allergen);
      recipe.allergens.add(allergen);
    });

    row.possibleHiddenAllergens.forEach((allergen) => {
      venue.possibleHidden.add(allergen);
      recipe.possibleHidden.add(allergen);
    });

    row.hiddenWarnings.forEach((warning) => {
      venue.hiddenWarningCount += 1;

      recipe.hiddenWarnings.push({
        warning,
        ingredientName: row.ingredientName,
        sourceRow: row.sourceRow,
      });
    });
  });

  const venues = [...venueMap.values()]
    .map((venue) => ({
      ...venue,
      allergens: [...venue.allergens].sort(),
      possibleHidden: [...venue.possibleHidden].sort(),
      recipes: [...venue.recipesMap.values()]
        .map((recipe) => ({
          ...recipe,
          allergens: [...recipe.allergens].sort(),
          possibleHidden: [...recipe.possibleHidden].sort(),
          ingredients: recipe.ingredients.sort((a, b) =>
            a.ingredientName.localeCompare(b.ingredientName)
          ),
          subRecipes: recipe.subRecipes.sort((a, b) =>
            a.ingredientName.localeCompare(b.ingredientName)
          ),
        }))
        .sort(
          (a, b) =>
            a.menuName.localeCompare(b.menuName) ||
            a.recipeName.localeCompare(b.recipeName)
        ),
    }))
    .sort((a, b) => a.restaurantName.localeCompare(b.restaurantName));

  return {
    rows: parsedRows,
    venues,
    sourceSheet: sheetName,
  };
};

const allergenBadgeStyle = (allergen, possible = false) => {
  const config = ALLERGEN_DISPLAY[allergen] || { color: "#555" };

  return {
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
    padding: "4px 7px",
    borderRadius: 999,
    background: possible ? "#fff0f0" : "#f2f2f2",
    border: `1px solid ${possible ? "#b00020" : config.color}`,
    color: possible ? "#b00020" : config.color,
    fontWeight: "bold",
    fontSize: 11,
    margin: "2px 3px 2px 0",
  };
};

const AllergenBadges = ({ allergens = [], possible = false }) => {
  if (!allergens.length) {
    return <span style={{ color: "#777", fontSize: 13 }}>None found</span>;
  }

  return (
    <div>
      {allergens.map((allergen) => {
        const icon = ALLERGEN_DISPLAY[allergen]?.icon || "⚠️";

        return (
          <span
            key={`${possible ? "possible" : "allergen"}-${allergen}`}
            style={allergenBadgeStyle(allergen, possible)}
          >
            <span>{icon}</span>
            <span>
              {possible ? "Possible " : ""}
              {allergen}
            </span>
          </span>
        );
      })}
    </div>
  );
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

export default function AllergenModule({
  styles,
  supabase,
  userShip,
  isAdmin = false,
  onBack,
  logUsageEvent = () => {},
}) {
  const [sourceFileName, setSourceFileName] = useState("");
  const [sourceSheetName, setSourceSheetName] = useState("");
  const [rows, setRows] = useState([]);
  const [venues, setVenues] = useState([]);
  const [selectedVenueKey, setSelectedVenueKey] = useState("");
  const [selectedRecipeKey, setSelectedRecipeKey] = useState("");
  const [venueSearch, setVenueSearch] = useState("");
  const [recipeSearch, setRecipeSearch] = useState("");
  const [ingredientSearch, setIngredientSearch] = useState("");
  const [allergenFilter, setAllergenFilter] = useState("ALL");
  const [message, setMessage] = useState("");
  const [loading, setLoading] = useState(false);
  const [permanentFileLoading, setPermanentFileLoading] = useState(false);

  const applyIngredientByLocationArrayBuffer = (
    arrayBuffer,
    fileName = "Permanent Ingredient by Location"
  ) => {
    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellDates: true,
    });

    const parsed = parseIngredientByLocationWorkbook(workbook);

    if (!parsed.rows.length) {
      throw new Error("No usable recipe ingredient rows found in the file.");
    }

    setRows(parsed.rows);
    setVenues(parsed.venues);
    setSourceFileName(fileName);
    setSourceSheetName(parsed.sourceSheet);
    setSelectedVenueKey(parsed.venues[0]?.venueKey || "");
    setSelectedRecipeKey("");
    setIngredientSearch("");

    return parsed;
  };

  const loadPermanentIngredientByLocationFile = async ({ silent = false } = {}) => {
    if (!supabase) {
      if (!silent) {
        const text =
          "Supabase is not connected. Permanent Ingredient by Location file cannot load.";
        setMessage(text);
        window.alert(text);
      }

      return false;
    }

    setPermanentFileLoading(true);

    if (!silent) {
      setMessage("Loading permanent Ingredient by Location file...");
    }

    try {
      const arrayBuffer = await downloadIngredientByLocationFileFromStorage({
        supabase,
      });

      const parsed = applyIngredientByLocationArrayBuffer(
        arrayBuffer,
        "Permanent Ingredient by Location"
      );

      setMessage(
        `Permanent Ingredient by Location loaded. ${parsed.venues.length} venue(s), ${parsed.rows.length} ingredient row(s).`
      );

      return true;
    } catch (error) {
      if (!silent) {
        const text =
          error?.message ||
          "Could not load permanent Ingredient by Location file.";

        setMessage(text);
        window.alert(text);
      }

      return false;
    } finally {
      setPermanentFileLoading(false);
    }
  };

  useEffect(() => {
    loadPermanentIngredientByLocationFile({ silent: true });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [supabase]);

  const uploadIngredientByLocationFile = async (event) => {
    const file = event.target.files?.[0];

    if (!file) return;

    if (!isAdmin) {
      const text =
        "Only admins can replace the permanent Ingredient by Location file.";

      setMessage(text);
      window.alert(text);
      event.target.value = "";
      return;
    }

    if (!supabase) {
      const text =
        "Supabase is not connected. Cannot save the permanent Ingredient by Location file.";

      setMessage(text);
      window.alert(text);
      event.target.value = "";
      return;
    }

    setLoading(true);
    setMessage("Saving permanent Ingredient by Location file...");

    try {
      await uploadIngredientByLocationFileToStorage({
        supabase,
        file,
      });

      const arrayBuffer = await file.arrayBuffer();
      const parsed = applyIngredientByLocationArrayBuffer(arrayBuffer, file.name);

      setMessage(
        `Permanent Ingredient by Location file updated. ${parsed.venues.length} venue(s), ${parsed.rows.length} ingredient row(s) loaded.`
      );

      logUsageEvent("permanent_ingredient_by_location_updated", {
        module: "allergen",
        ship: userShip,
        fileName: file.name,
        permanent: true,
        venues: parsed.venues.length,
        rows: parsed.rows.length,
      });
    } catch (error) {
      const text =
        error?.message ||
        "Could not save permanent Ingredient by Location file.";

      setMessage(text);
      window.alert(text);
    } finally {
      setLoading(false);
      event.target.value = "";
    }
  };

  const allAllergens = useMemo(() => {
    const set = new Set();

    rows.forEach((row) => {
      row.explicitAllergens.forEach((item) => set.add(item));
      (row.recipeDeclaredAllergens || []).forEach((item) => set.add(item));
      row.possibleHiddenAllergens.forEach((item) => set.add(item));
    });

    return [...set].sort();
  }, [rows]);

  const selectedVenue = useMemo(
    () => venues.find((venue) => venue.venueKey === selectedVenueKey) || null,
    [venues, selectedVenueKey]
  );

  const filteredVenues = useMemo(() => {
    const query = venueSearch.toLowerCase().trim();

    return venues.filter((venue) => {
      const matchesSearch =
        !query || venue.restaurantName.toLowerCase().includes(query);

      const matchesAllergen =
        allergenFilter === "ALL" ||
        venue.allergens.includes(allergenFilter) ||
        venue.possibleHidden.includes(allergenFilter);

      return matchesSearch && matchesAllergen;
    });
  }, [venues, venueSearch, allergenFilter]);

  const filteredRecipes = useMemo(() => {
    if (!selectedVenue) return [];

    const query = recipeSearch.toLowerCase().trim();

    return selectedVenue.recipes.filter((recipe) => {
      const matchesSearch =
        !query ||
        [recipe.recipeCode, recipe.recipeName, recipe.menuCode, recipe.menuName]
          .join(" ")
          .toLowerCase()
          .includes(query);

      const matchesAllergen =
        allergenFilter === "ALL" ||
        recipe.allergens.includes(allergenFilter) ||
        recipe.possibleHidden.includes(allergenFilter);

      return matchesSearch && matchesAllergen;
    });
  }, [selectedVenue, recipeSearch, allergenFilter]);

  const selectedRecipe = useMemo(() => {
    if (!selectedVenue) return null;

    return (
      selectedVenue.recipes.find(
        (recipe) => recipe.recipeKey === selectedRecipeKey
      ) || null
    );
  }, [selectedVenue, selectedRecipeKey]);

  const visibleIngredients = useMemo(() => {
    if (!selectedRecipe) return [];

    const query = ingredientSearch.toLowerCase().trim();
    const combined = [...selectedRecipe.ingredients, ...selectedRecipe.subRecipes];

    return combined.filter((item) => {
      if (!query) return true;

      return [
        item.ingredientName,
        item.assigned,
        item.explicitAllergens.join(" "),
        item.possibleHiddenAllergens.join(" "),
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [selectedRecipe, ingredientSearch]);

  const hiddenWarningRows = useMemo(
    () => rows.filter((row) => row.hiddenWarnings.length > 0),
    [rows]
  );

  const exportRecipeMatrix = () => {
    const exportRows = rows.map((row, index) => ({
      Number: index + 1,
      Venue: row.restaurantName,
      MenuCode: row.menuCode,
      MenuName: row.menuName,
      RecipeCode: row.recipeCode,
      RecipeName: row.recipeName,
      IngredientCode: row.ingredientCode,
      IngredientName: row.ingredientName,
      Type: row.assignedType === "R" ? "Sub Recipe" : "Ingredient",
      Category: row.category,
      SubCategory: row.subCategory,
      IngredientDeclaredAllergens: row.ingredientAllergens.join(", "),
      RecipeDeclaredAllergens: (row.recipeDeclaredAllergens || []).join(", "),
      RealAllergensShownOnIngredient: row.explicitAllergens.join(", "),
      NameDetectedAllergens: row.detectedAllergens.join(", "),
      PossibleHiddenAllergens: row.possibleHiddenAllergens.join(", "),
      HiddenWarnings: row.hiddenWarnings.join(" | "),
      IgnoredBasic: row.ignoredBasic ? "Yes" : "No",
      PlainRawIngredient: row.plainRawIngredient ? "Yes" : "No",
      PossibleAllergensSkipped: row.skipPossibleAllergens ? "Yes" : "No",
      SourceRow: row.sourceRow,
    }));

    exportRowsToExcel(
      exportRows,
      "Allergen Matrix",
      `allergen-matrix-${userShip || "ship"}.xlsx`
    );
  };

  return (
    <main style={styles.page}>
      <header style={styles.header}>
        <img
          src="/virgin-logo.png"
          alt="Virgin Voyages"
          style={styles.headerLogo}
        />

        <div style={styles.headerActions}>
          <button style={styles.backButton} onClick={onBack}>
            ← Back
          </button>

          <div style={styles.shipBadge}>
            🧬 Allergens {userShip ? `• ${userShip}` : ""}
          </div>
        </div>
      </header>

      <section style={styles.grid}>
        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🧬 Allergen Matrix</h2>

          <p style={styles.emptyText}>
            Upload the Ingredient by Location workbook. This screen groups recipes
            by venue, then shows ingredients, sub-recipes, declared allergens, and
            possible hidden allergens.
          </p>

          <div style={styles.infoBox}>
            <div>📄 Ingredient by Location file loads automatically for all users.</div>
            <div>🔒 Only admins can replace the permanent file.</div>
          </div>

          <button
            type="button"
            style={styles.backButton}
            onClick={() => loadPermanentIngredientByLocationFile()}
            disabled={loading || permanentFileLoading}
          >
            🔄 Reload Permanent Ingredient File
          </button>

          {isAdmin && (
            <>
              <label style={styles.label}>
                Admin only: replace permanent Ingredient by Location file
              </label>

              <input
                type="file"
                accept=".xlsx,.xls,.xlsm"
                onChange={uploadIngredientByLocationFile}
                style={styles.fileInput}
                disabled={loading || permanentFileLoading}
              />
            </>
          )}

          <div style={styles.infoBox}>
            <div>
              📄 File: <strong>{sourceFileName || "Not uploaded"}</strong>
            </div>
            <div>
              📋 Sheet: <strong>{sourceSheetName || "N/A"}</strong>
            </div>
            <div>
              🏢 Venues: <strong>{venues.length}</strong>
            </div>
            <div>
              🧾 Ingredient rows: <strong>{rows.length}</strong>
            </div>
            <div>
              🚨 Hidden warnings: <strong>{hiddenWarningRows.length}</strong>
            </div>
            {(loading || permanentFileLoading) && <div>Loading...</div>}
          </div>

          {message && <p style={styles.message}>{message}</p>}

          <div style={styles.warningText}>
            This is a support tool only. Ingredient cards now show only allergens
            that belong to that ingredient. Recipe-level allergens are shown in
            recipe summaries but are not copied to every ingredient. Fresh herbs,
            raw produce, water, salt, sugar, and pepper do not receive possible
            hidden allergen warnings. Always verify against official recipe cards
            and supplier specifications before answering a Sailor allergy request.
          </div>

          <button
            style={styles.primaryButton}
            onClick={exportRecipeMatrix}
            disabled={!rows.length}
          >
            📥 Export Full Allergen Matrix
          </button>
        </div>

        <div style={styles.card}>
          <h2 style={styles.cardTitle}>🔎 Filters</h2>

          <label style={styles.label}>Allergen filter</label>

          <select
            value={allergenFilter}
            onChange={(event) => setAllergenFilter(event.target.value)}
            style={styles.select}
          >
            <option value="ALL">All allergens</option>
            {allAllergens.map((allergen) => (
              <option key={allergen} value={allergen}>
                {allergen}
              </option>
            ))}
          </select>

          <label style={styles.label}>Search venue</label>

          <input
            placeholder="Search venue..."
            value={venueSearch}
            onChange={(event) => setVenueSearch(event.target.value)}
            style={styles.searchInput}
          />

          <label style={styles.label}>Search recipe/menu</label>

          <input
            placeholder="Search recipe code, recipe name, menu..."
            value={recipeSearch}
            onChange={(event) => setRecipeSearch(event.target.value)}
            style={styles.searchInput}
          />
        </div>
      </section>

      <section style={styles.card}>
        <h2 style={styles.productTitle}>🏢 Venues</h2>

        {!venues.length && (
          <p style={styles.emptyText}>
            Permanent Ingredient by Location file will load automatically. Admin
            can replace it if needed.
          </p>
        )}

        <div style={localStyles.venueGrid}>
          {filteredVenues.map((venue) => (
            <button
              key={venue.venueKey}
              type="button"
              style={{
                ...localStyles.venueCard,
                ...(selectedVenueKey === venue.venueKey
                  ? localStyles.venueCardActive
                  : {}),
              }}
              onClick={() => {
                setSelectedVenueKey(venue.venueKey);
                setSelectedRecipeKey("");
                setIngredientSearch("");
              }}
            >
              <strong
                style={{
                  ...localStyles.venueName,
                  color: getVenueNameColor(venue.restaurantName),
                }}
              >
                {venue.restaurantName}
              </strong>

              <span style={localStyles.venueRecipeCount}>
                {venue.recipes.length} recipe(s)
              </span>
            </button>
          ))}
        </div>
      </section>

      {selectedVenue && (
        <section style={styles.card}>
          <div
            style={{
              ...styles.header,
              boxShadow: "none",
              padding: 0,
              marginBottom: 16,
            }}
          >
            <div>
              <h2 style={styles.productTitle}>{selectedVenue.restaurantName}</h2>

              <p style={{ ...styles.emptyText, margin: 0 }}>
                Click a recipe to view ingredients, sub-recipes and allergen detail.
              </p>
            </div>

            <div style={styles.shipBadge}>{filteredRecipes.length} recipe(s)</div>
          </div>

          <div style={localStyles.recipeGrid}>
            {filteredRecipes.map((recipe) => (
              <button
                key={recipe.recipeKey}
                type="button"
                style={{
                  ...localStyles.recipeCard,
                  ...(selectedRecipeKey === recipe.recipeKey
                    ? localStyles.recipeCardActive
                    : {}),
                  ...(recipe.hiddenWarnings.length > 0
                    ? localStyles.recipeCardWarning
                    : {}),
                }}
                onClick={() => {
                  setSelectedRecipeKey(recipe.recipeKey);
                  setIngredientSearch("");
                }}
              >
                <strong>{recipe.recipeName}</strong>
                <span>Recipe code: {recipe.recipeCode || "N/A"}</span>
                <span>
                  {recipe.ingredients.length} ingredient(s),{" "}
                  {recipe.subRecipes.length} sub recipe line(s)
                </span>

                <AllergenBadges allergens={recipe.allergens.slice(0, 6)} />

                {recipe.possibleHidden.length > 0 && (
                  <AllergenBadges
                    allergens={recipe.possibleHidden.slice(0, 4)}
                    possible
                  />
                )}

                {recipe.hiddenWarnings.length > 0 && (
                  <div style={styles.statusBad}>Hidden warning</div>
                )}
              </button>
            ))}
          </div>
        </section>
      )}

      {selectedRecipe && (
        <div
          style={styles.modalBackdrop}
          onClick={() => {
            setSelectedRecipeKey("");
            setIngredientSearch("");
          }}
        >
          <div
            style={localStyles.recipeModalCard}
            onClick={(event) => event.stopPropagation()}
          >
            <button
              type="button"
              style={styles.closeButton}
              onClick={() => {
                setSelectedRecipeKey("");
                setIngredientSearch("");
              }}
            >
              ✕
            </button>

            <div style={localStyles.modalHeader}>
              <div>
                <h2 style={{ ...styles.productTitle, marginBottom: 4 }}>
                  🧾 {selectedRecipe.recipeName}
                </h2>

                <p style={{ ...styles.emptyText, margin: 0 }}>
                  Recipe code {selectedRecipe.recipeCode || "N/A"} • Menu{" "}
                  {selectedRecipe.menuName || "N/A"}
                </p>
              </div>

              <div style={styles.statusNeutral}>Allergen Review</div>
            </div>

            <section style={localStyles.modalSummaryGrid}>
              <div style={styles.infoBox}>
                <strong>Recipe / item allergens</strong>
                <AllergenBadges allergens={selectedRecipe.allergens} />
              </div>

              <div style={styles.infoBox}>
                <strong>Possible hidden allergens</strong>
                <AllergenBadges allergens={selectedRecipe.possibleHidden} possible />
              </div>
            </section>

            <input
              placeholder="Search ingredient or allergen..."
              value={ingredientSearch}
              onChange={(event) => setIngredientSearch(event.target.value)}
              style={{ ...styles.searchInput, marginTop: 14 }}
            />

            <div style={localStyles.ingredientGrid}>
              {visibleIngredients.map((item) => (
                <div
                  key={`${item.sourceRow}-${item.ingredientCode}-${item.ingredientName}`}
                  style={{
                    ...localStyles.ingredientCard,
                    ...(item.hiddenWarnings.length > 0
                      ? localStyles.ingredientCardWarning
                      : {}),
                    ...(item.possibleHiddenAllergens.length > 0 &&
                    !item.hiddenWarnings.length
                      ? localStyles.ingredientCardPossible
                      : {}),
                  }}
                >
                  <div style={localStyles.ingredientTitle}>
                    {item.ingredientName}
                  </div>

                  {item.assigned && item.assigned !== item.ingredientName && (
                    <div style={localStyles.compactMeta}>
                      Product / Assigned: {item.assigned}
                    </div>
                  )}

                  <div style={localStyles.typePill}>
                    {item.assignedType === "R" ? "Sub Recipe" : "Ingredient"}
                  </div>

                  {item.ignoredBasic && (
                    <div style={styles.statusNeutral}>Ignored basic item</div>
                  )}

                  {item.plainRawIngredient && !item.ignoredBasic && (
                    <div style={styles.statusNeutral}>
                      Fresh/raw item — possible allergens skipped
                    </div>
                  )}

                  <div>
                    <strong>Declared / real:</strong>
                    <AllergenBadges allergens={item.explicitAllergens} />
                  </div>

                  {item.possibleHiddenAllergens.length > 0 && (
                    <div>
                      <strong>Possible hidden:</strong>
                      <AllergenBadges
                        allergens={item.possibleHiddenAllergens}
                        possible
                      />
                    </div>
                  )}

                  {item.hiddenWarnings.map((warning, index) => (
                    <div key={`${warning}-${index}`} style={styles.statusBad}>
                      ⚠️ {warning}
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </main>
  );
}

const localStyles = {
  venueGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(145px, 1fr))",
    gap: 8,
  },

  venueCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 12,
    background: "#fff",
    display: "grid",
    gap: 6,
    textAlign: "left",
    cursor: "pointer",
    fontFamily: "inherit",
    color: "inherit",
    fontSize: 12,
    alignContent: "center",
    minHeight: 74,
  },

  venueCardActive: {
    border: "2px solid #111",
    background: "#f2f2f2",
  },

  venueCardWarning: {
    border: "2px solid #b00020",
    background: "#fff0f0",
  },

  venueIcon: {
    display: "none",
  },

  venueName: {
    fontSize: 14,
    lineHeight: 1.15,
    fontWeight: 800,
    overflowWrap: "anywhere",
  },

  venueRecipeCount: {
    color: "#555",
    fontSize: 12,
    fontWeight: 700,
  },

  recipeGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(165px, 1fr))",
    gap: 8,
  },

  recipeCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 8,
    background: "#fafafa",
    display: "grid",
    gap: 4,
    textAlign: "left",
    cursor: "pointer",
    fontFamily: "inherit",
    color: "inherit",
    fontSize: 12,
    alignContent: "start",
  },

  recipeCardActive: {
    border: "2px solid #111",
    background: "#f2f2f2",
  },

  recipeCardWarning: {
    border: "2px solid #b00020",
    background: "#fff0f0",
  },

  recipeModalCard: {
    background: "#fff",
    borderRadius: 18,
    padding: 20,
    maxWidth: 1180,
    width: "96%",
    maxHeight: "90vh",
    overflowY: "auto",
    position: "relative",
    boxShadow: "0 20px 60px rgba(0,0,0,0.25)",
  },

  modalHeader: {
    display: "flex",
    justifyContent: "space-between",
    gap: 12,
    alignItems: "flex-start",
    marginBottom: 14,
    paddingRight: 42,
    flexWrap: "wrap",
  },

  modalSummaryGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
    gap: 12,
    marginBottom: 12,
  },

  ingredientGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(165px, 1fr))",
    gap: 8,
  },

  ingredientCard: {
    border: "1px solid #ddd",
    borderRadius: 14,
    padding: 8,
    background: "#fff",
    display: "grid",
    gap: 4,
    fontSize: 11,
    alignContent: "start",
  },

  ingredientCardPossible: {
    border: "1.5px solid #8a5a00",
    background: "#fff8e1",
  },

  ingredientCardWarning: {
    border: "2px solid #b00020",
    background: "#fff0f0",
  },

  ingredientTitle: {
    fontWeight: "bold",
    fontSize: 12.5,
    lineHeight: 1.12,
    overflowWrap: "anywhere",
  },

  compactMeta: {
    color: "#555",
    fontSize: 11,
    lineHeight: 1.2,
    overflowWrap: "anywhere",
  },

  typePill: {
    justifySelf: "start",
    padding: "3px 7px",
    borderRadius: 999,
    background: "#f2f2f2",
    color: "#555",
    fontSize: 11,
    fontWeight: "bold",
  },
};

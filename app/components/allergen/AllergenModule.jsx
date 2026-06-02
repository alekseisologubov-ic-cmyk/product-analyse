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

const escapeHtml = (value) =>
  String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

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

const extractCodesFromText = (value) =>
  String(value || "")
    .match(/\b\d{3,}\b/g)
    ?.map((item) => normalizeCode(item))
    .filter(Boolean) || [];

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

const normalizeRecipeLookupName = (value) =>
  normalizeIngredientTextForMatch(value)
    .replace(/\bVV\b/g, " ")
    .replace(/\bBR\b/g, " ")
    .replace(/\bSC\b/g, " ")
    .replace(/\bVL\b/g, " ")
    .replace(/\bRL\b/g, " ")
    .replace(/\b20\d{2}\b/g, " ")
    .replace(/\b19\d{2}\b/g, " ")
    .replace(/\bSUB RECIPE\b/g, " ")
    .replace(/\bRECIPE\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const getRecipeLookupKeys = (value) =>
  [
    normalizeIngredientTextForMatch(value),
    normalizeRecipeLookupName(value),
  ].filter(Boolean);

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

const isGlutenFreeProductClaim = (...values) => {
  const text = values
    .map((value) => cleanText(value))
    .join(" ")
    .replace(/[^A-Z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!text) return false;

  return (
    /\bGF\b/.test(text) ||
    text.includes("GLUTEN FREE") ||
    text.includes("GLUTENFREE") ||
    text.includes("FREE FROM GLUTEN") ||
    text.includes("NO GLUTEN")
  );
};

const ALLERGEN_ORDER = [
  "Cereals containing gluten",
  "Crustaceans",
  "Eggs",
  "Fish",
  "Peanuts",
  "Soybeans",
  "Milk",
  "Tree Nuts",
  "Celery",
  "Mustard",
  "Sesame seeds",
  "Sulphur dioxide and sulphites",
  "Lupin",
  "Molluscs",
];

const VALID_ALLERGENS = new Set(ALLERGEN_ORDER);

const sortAllergens = (allergens = []) => {
  const unique = [...new Set(allergens)].filter((item) =>
    VALID_ALLERGENS.has(item)
  );

  return unique.sort(
    (a, b) => ALLERGEN_ORDER.indexOf(a) - ALLERGEN_ORDER.indexOf(b)
  );
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
  "ROMAINE",
  "SPINACH",
  "ARUGULA",
  "ROCKET",
  "KALE",
  "CABBAGE",
  "RED CABBAGE",
  "GREEN CABBAGE",
  "NAPA CABBAGE",
  "BOK CHOY",
  "PAK CHOY",
  "CHARD",
  "WATERCRESS",
  "ENDIVE",
  "ESCAROLE",
  "RADICCHIO",

  "CARROT",
  "CARROTS",
  "ONION",
  "ONIONS",
  "RED ONION",
  "WHITE ONION",
  "YELLOW ONION",
  "GARLIC",
  "SHALLOT",
  "SHALLOTS",
  "TOMATO",
  "TOMATOES",
  "CUCUMBER",
  "CUCUMBERS",
  "ZUCCHINI",
  "COURGETTE",
  "EGGPLANT",
  "AUBERGINE",
  "POTATO",
  "POTATOES",
  "SWEET POTATO",
  "SWEET POTATOES",
  "YAM",
  "YAMS",
  "MUSHROOM",
  "MUSHROOMS",
  "BELL PEPPER",
  "BELL PEPPERS",
  "PEPPER",
  "PEPPERS",
  "CAPSICUM",
  "CHILI",
  "CHILLI",
  "JALAPENO",
  "JALAPEÑO",
  "BROCCOLI",
  "CAULIFLOWER",
  "BRUSSELS SPROUT",
  "BRUSSELS SPROUTS",
  "ASPARAGUS",
  "ARTICHOKE",
  "ARTICHOKES",
  "FENNEL",
  "LEEKS",
  "LEEK",
  "RADISH",
  "RADISHES",
  "TURNIP",
  "TURNIPS",
  "PARSNIP",
  "PARSNIPS",
  "RUTABAGA",
  "BEET",
  "BEETS",
  "BEETROOT",
  "BEETROOTS",
  "CORN",
  "SWEET CORN",
  "GREEN BEAN",
  "GREEN BEANS",
  "SNAP PEA",
  "SNAP PEAS",
  "SNOW PEA",
  "SNOW PEAS",
  "PEA",
  "PEAS",
  "OKRA",

  "SQUASH",
  "BUTTERNUT",
  "BUTTERNUT SQUASH",
  "ACORN SQUASH",
  "KABOCHA",
  "PUMPKIN",
  "PUMPKINS",
  "SPAGHETTI SQUASH",
  "YELLOW SQUASH",

  "APPLE",
  "APPLES",
  "ORANGE",
  "ORANGES",
  "LEMON",
  "LEMONS",
  "LIME",
  "LIMES",
  "BANANA",
  "BANANAS",
  "BERRY",
  "BERRIES",
  "STRAWBERRY",
  "STRAWBERRIES",
  "BLUEBERRY",
  "BLUEBERRIES",
  "RASPBERRY",
  "RASPBERRIES",
  "BLACKBERRY",
  "BLACKBERRIES",
  "PINEAPPLE",
  "MANGO",
  "MANGOES",
  "PAPAYA",
  "MELON",
  "WATERMELON",
  "CANTALOUPE",
  "HONEYDEW",
  "GRAPE",
  "GRAPES",
  "PEACH",
  "PEACHES",
  "PEAR",
  "PEARS",
  "PLUM",
  "PLUMS",
  "APRICOT",
  "APRICOTS",
  "KIWI",
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

  return (
    isIgnoredBasicIngredient(text) ||
    isPlainFreshHerbOrRawProduce(...values)
  );
};

const isSubRecipeRow = (item) => {
  const assignedType = cleanText(item?.assignedType);
  const hasProductRelation = cleanText(item?.hasProductRelation);
  const recipeIsBasic = cleanText(item?.recipeIsBasic);
  const nameText = cleanText(
    `${item?.ingredientName || ""} ${item?.assigned || ""}`
  );

  if (assignedType === "R" || assignedType.includes("RECIPE")) {
    return true;
  }

  const looksLikeRecipeName =
    nameText.includes(" - VV") ||
    nameText.includes("(BR)") ||
    nameText.includes("(SC)") ||
    nameText.includes("(VL)") ||
    nameText.includes("(RL)") ||
    nameText.includes(" RECIPE");

  return (
    looksLikeRecipeName &&
    (hasProductRelation === "N" || recipeIsBasic === "Y")
  );
};

const isOnboardRecipeName = (...values) => {
  const text = cleanText(values.filter(Boolean).join(" "));

  if (!text) return false;

  return (
    text.includes(" - VV") ||
    text.includes("(BR)") ||
    text.includes("(SC)") ||
    text.includes("(VL)") ||
    text.includes("(RL)") ||
    text.includes(" RECIPE")
  );
};

const ALLERGEN_DISPLAY = {
  "Cereals containing gluten": { icon: "🌾", color: "#8a5a00" },
  Crustaceans: { icon: "🦐", color: "#b00020" },
  Eggs: { icon: "🥚", color: "#7a4f00" },
  Fish: { icon: "🐟", color: "#005f73" },
  Peanuts: { icon: "🥜", color: "#b00020" },
  Soybeans: { icon: "🫘", color: "#2e7d32" },
  Milk: { icon: "🥛", color: "#0057b8" },
  "Tree Nuts": { icon: "🌰", color: "#6b3f1d" },
  Celery: { icon: "🥬", color: "#2e7d32" },
  Mustard: { icon: "🟡", color: "#8a5a00" },
  "Sesame seeds": { icon: "⚪", color: "#6a4a00" },
  "Sulphur dioxide and sulphites": { icon: "⚠️", color: "#8a5a00" },
  Lupin: { icon: "🌱", color: "#2e7d32" },
  Molluscs: { icon: "🦪", color: "#005f73" },
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
    text.includes("SPELT") ||
    text.includes("KHORASAN") ||
    text.includes("KAMUT") ||
    text.includes("BARLEY") ||
    text.includes("RYE") ||
    text.includes("OAT")
  ) {
    return "Cereals containing gluten";
  }

  if (
    text.includes("CRUSTACEAN") ||
    text.includes("SHRIMP") ||
    text.includes("PRAWN") ||
    text.includes("CRAB") ||
    text.includes("LOBSTER") ||
    text.includes("CRAYFISH") ||
    text.includes("SCAMPI")
  ) {
    return "Crustaceans";
  }

  if (text.includes("EGG")) return "Eggs";

  if (
    text.includes("FISH") ||
    text.includes("ANCHOV") ||
    text.includes("SALMON") ||
    text.includes("TUNA") ||
    text.includes("COD") ||
    text.includes("SARDINE") ||
    text.includes("MACKEREL") ||
    text.includes("TROUT")
  ) {
    return "Fish";
  }

  if (text.includes("PEANUT")) return "Peanuts";

  if (
    text.includes("SOY") ||
    text.includes("SOYA") ||
    text.includes("SOYBEAN") ||
    text.includes("TOFU") ||
    text.includes("EDAMAME") ||
    text.includes("MISO") ||
    text.includes("TAMARI")
  ) {
    return "Soybeans";
  }

  if (
    text.includes("MILK") ||
    text.includes("LACTOSE") ||
    text.includes("DAIRY") ||
    text.includes("CREAM") ||
    text.includes("BUTTER") ||
    text.includes("CHEESE") ||
    text.includes("YOGURT") ||
    text.includes("YOGHURT") ||
    text.includes("WHEY") ||
    text.includes("CASEIN")
  ) {
    return "Milk";
  }

  if (
    text.includes("TREE NUT") ||
    text.includes("ALMOND") ||
    text.includes("HAZELNUT") ||
    text.includes("WALNUT") ||
    text.includes("CASHEW") ||
    text.includes("PECAN") ||
    text.includes("BRAZIL NUT") ||
    text.includes("PISTACHIO") ||
    text.includes("MACADAMIA") ||
    text.includes("QUEENSLAND NUT") ||
    text === "NUT" ||
    text === "NUTS"
  ) {
    return "Tree Nuts";
  }

  if (text.includes("CELERY") || text.includes("CELERIAC")) return "Celery";
  if (text.includes("MUSTARD") || text.includes("DIJON")) return "Mustard";

  if (text.includes("SESAME") || text.includes("TAHINI")) {
    return "Sesame seeds";
  }

  if (
    text.includes("SULPH") ||
    text.includes("SULF") ||
    text.includes("METABISULFITE") ||
    text.includes("SULFUR DIOXIDE") ||
    text.includes("SULPHUR DIOXIDE")
  ) {
    return "Sulphur dioxide and sulphites";
  }

  if (text.includes("LUPIN") || text.includes("LUPINE")) return "Lupin";

  if (
    text.includes("MOLLUSC") ||
    text.includes("MOLLUSK") ||
    text.includes("CLAM") ||
    text.includes("MUSSEL") ||
    text.includes("OYSTER") ||
    text.includes("SCALLOP") ||
    text.includes("SQUID") ||
    text.includes("OCTOPUS") ||
    text.includes("CALAMARI") ||
    text.includes("SNAIL") ||
    text.includes("ESCARGOT")
  ) {
    return "Molluscs";
  }

  return "";
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

  return sortAllergens([...found]);
};

const KEYWORD_RULES = [
  {
    allergen: "Cereals containing gluten",
    words: [
      "wheat",
      "wheat flour",
      "all purpose flour",
      "ap flour",
      "bread flour",
      "cake flour",
      "spelt",
      "khorasan",
      "kamut",
      "semolina",
      "barley",
      "rye",
      "oat",
      "oats",
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
    allergen: "Crustaceans",
    words: [
      "shrimp",
      "prawn",
      "prawns",
      "crab",
      "lobster",
      "crayfish",
      "scampi",
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
      "worcestershire",
    ],
  },
  {
    allergen: "Peanuts",
    words: ["peanut", "peanuts", "peanut butter", "peanutbutter", "satay"],
  },
  {
    allergen: "Soybeans",
    words: [
      "soy",
      "soya",
      "soybean",
      "soybeans",
      "tofu",
      "edamame",
      "miso",
      "tamari",
      "soy sauce",
      "yuba",
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
    allergen: "Tree Nuts",
    words: [
      "almond",
      "hazelnut",
      "walnut",
      "cashew",
      "pecan",
      "brazil nut",
      "pistachio",
      "macadamia",
      "queensland nut",
      "pine nut",
      "marzipan",
      "praline",
      "gianduja",
      "nutella",
    ],
  },
  {
    allergen: "Celery",
    words: ["celery", "celeriac"],
  },
  {
    allergen: "Mustard",
    words: ["mustard", "dijon"],
  },
  {
    allergen: "Sesame seeds",
    words: [
      "sesame",
      "sesame seed",
      "sesame seeds",
      "tahini",
      "benne",
      "gingelly",
    ],
  },
  {
    allergen: "Sulphur dioxide and sulphites",
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
  {
    allergen: "Lupin",
    words: ["lupin", "lupine"],
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
      "snail",
      "escargot",
    ],
  },
];

const isProcessedOrPreparedItem = (value) => {
  const text = normalizeIngredientTextForMatch(value);

  if (!text) return false;

  return [
    "SAUCE",
    "DRESSING",
    "MARINADE",
    "PASTE",
    "MIX",
    "SEASONING",
    "COOKIE",
    "COOKIES",
    "CAKE",
    "BISCUIT",
    "BISCUITS",
    "CHOCOLATE",
    "CHIPS",
    "CEREAL",
    "CRACKER",
    "CRACKERS",
    "BREAD",
    "PASTRY",
    "GLAZE",
    "PUREE",
    "PURÉE",
    "MASH",
    "MASHED",
    "SPREAD",
    "BUTTER",
    "PESTO",
    "MAYONNAISE",
    "MAYO",
    "AIOLI",
    "HOLLANDAISE",
    "TERIYAKI",
    "HOISIN",
    "WORCESTERSHIRE",
    "OYSTER SAUCE",
    "FISH SAUCE",
  ].some((word) => textHasWordOrPhrase(text, word));
};

const PREMADE_COMMON_ALLERGEN_RULES = [
  {
    words: ["soy sauce", "teriyaki", "hoisin"],
    allergens: ["Soybeans", "Cereals containing gluten"],
  },
  {
    words: ["fish sauce"],
    allergens: ["Fish"],
  },
  {
    words: ["oyster sauce"],
    allergens: ["Molluscs", "Soybeans"],
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
    allergens: ["Cereals containing gluten", "Milk", "Eggs"],
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
    allergens: ["Cereals containing gluten"],
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
    allergens: ["Cereals containing gluten", "Tree Nuts"],
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
      rule.allergens.forEach((allergen) => {
        if (VALID_ALLERGENS.has(allergen)) found.add(allergen);
      });
    }
  });

  if (isGlutenFreeProductClaim(...values)) {
    found.delete("Cereals containing gluten");
  }

  return sortAllergens([...found]);
};

const keywordAllergensForText = (...values) => {
  const text = values
    .map((value) => normalizeIngredientTextForMatch(value))
    .join(" ");

  const found = new Set();

  if (!text.trim()) return [];

  KEYWORD_RULES.forEach((rule) => {
    const matched = rule.words.some((word) => textHasWordOrPhrase(text, word));

    if (matched && VALID_ALLERGENS.has(rule.allergen)) {
      found.add(rule.allergen);
    }
  });

  if (
    textHasWordOrPhrase(text, "eggplant") ||
    textHasWordOrPhrase(text, "aubergine")
  ) {
    found.delete("Eggs");
  }

  if (textHasWordOrPhrase(text, "seedless")) {
    found.delete("Sesame seeds");
  }

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
    found.delete("Cereals containing gluten");
  }

  if (
    textHasWordOrPhrase(text, "cream of tartar") ||
    textHasWordOrPhrase(text, "coconut milk") ||
    textHasWordOrPhrase(text, "almond milk") ||
    textHasWordOrPhrase(text, "soy milk") ||
    textHasWordOrPhrase(text, "soya milk") ||
    textHasWordOrPhrase(text, "oat milk")
  ) {
    found.delete("Milk");
  }

  if (isGlutenFreeProductClaim(...values)) {
    found.delete("Cereals containing gluten");
  }

  return sortAllergens([...found]);
};

const looksGlutenFreeClaim = (...values) => {
  const text = values.map((value) => cleanText(value)).join(" ");

  return (
    /\bGF\b/.test(text) ||
    text.includes("GLUTEN FREE") ||
    text.includes("GLUTEN-FREE") ||
    text.includes("GLUTENFREE") ||
    text.includes("FREE FROM GLUTEN") ||
    text.includes("NO GLUTEN")
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
    ingredientCode: ingredientCodeIndex,
    ingredientName: ingredientNameIndex,
    assigned: findExact(["Assigned"], 12),
    assignedType: findExact(["AssignedType", "Assigned Type"], 13),
    recipeCode: findExact(["RecipeCode", "Recipe Code"], 15),
    recipeName: findExact(["RecipeName", "Recipe Name"], 16),
    specialInstructions: findExact(
      ["SpecialInstructions", "Special Instructions"],
      17
    ),
    specialInstructions2: findExact(
      ["SpecialInstructions2", "Special Instructions 2"],
      18
    ),
    recipeIsBasic: findExact(["RecipeIsBasic", "Recipe Is Basic"], 19),
    hasProductRelation: findExact(
      ["HasProductRelation", "Has Product Relation"],
      20
    ),
    ingredientAllergens: findExact(
      ["IngredientAllergens", "Ingredient Allergens"],
      21
    ),
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
    const assignedType = safeText(get(indexes.assignedType));
    const recipeIsBasic = safeText(get(indexes.recipeIsBasic));
    const hasProductRelation = safeText(get(indexes.hasProductRelation));
    const ingredientCode = normalizeCode(get(indexes.ingredientCode));
    const glutenFreeProductClaim = isGlutenFreeProductClaim(ingredientName, assigned);

    const rawIngredientAllergens = splitAllergens(get(indexes.ingredientAllergens));
    const rawRecipeDeclaredAllergens = splitAllergens(get(indexes.recipeAllergens));

    const ingredientAllergens = glutenFreeProductClaim
      ? rawIngredientAllergens.filter(
          (allergen) => allergen !== "Cereals containing gluten"
        )
      : rawIngredientAllergens;

    const recipeDeclaredAllergens = rawRecipeDeclaredAllergens;

    const currentLineIsSubRecipe = isSubRecipeRow({
      ingredientName,
      assigned,
      assignedType,
      recipeIsBasic,
      hasProductRelation,
    });

    const onboardRecipeLine =
      currentLineIsSubRecipe || isOnboardRecipeName(ingredientName, assigned);

    const hasRealItemCode = Boolean(ingredientCode) && !onboardRecipeLine;

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
      hasRealItemCode &&
      !onboardRecipeLine &&
      (isProcessedOrPreparedItem(ingredientName) ||
        isProcessedOrPreparedItem(assigned));

    const preparedCommonAllergens = processedItem
      ? getPreparedCommonAllergens(ingredientName, assigned)
      : [];

    const explicitAllergens = ignoredBasic
      ? []
      : sortAllergens([...ingredientAllergens, ...nameDetectedRealAllergens]);

    const detectedAllergens = ignoredBasic
      ? []
      : sortAllergens(nameDetectedRealAllergens);

    const possibleHiddenAllergens =
      ignoredBasic || skipPossibleAllergens
        ? []
        : sortAllergens(
            preparedCommonAllergens.filter(
              (allergen) => !explicitAllergens.includes(allergen)
            )
          );

    const gfClaim = looksGlutenFreeClaim(
      menuName,
      recipeName,
      get(indexes.specialInstructions),
      get(indexes.specialInstructions2)
    );

    const hasIngredientGluten = explicitAllergens.includes(
      "Cereals containing gluten"
    );

    const hiddenWarnings = [];

    if (!ignoredBasic && !skipPossibleAllergens && gfClaim && hasIngredientGluten) {
      hiddenWarnings.push(
        "Ingredient gluten check: recipe/menu says GF or gluten free, but this ingredient shows cereals containing gluten."
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
      ingredientCode,
      ingredientName,
      assigned,
      assignedType,
      recipeCode: normalizeCode(get(indexes.recipeCode)),
      recipeName,
      specialInstructions: safeText(get(indexes.specialInstructions)),
      specialInstructions2: safeText(get(indexes.specialInstructions2)),
      recipeIsBasic,
      hasProductRelation,
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
      onboardRecipeLine,
      glutenFreeProductClaim,
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

    if (isSubRecipeRow(row)) {
      recipe.subRecipes.push(row);
    } else {
      recipe.ingredients.push(row);
    }

    if (!row.ignoredBasic) venue.ingredientCount += 1;

    row.explicitAllergens.forEach((allergen) => {
      if (VALID_ALLERGENS.has(allergen)) {
        venue.allergens.add(allergen);
        recipe.allergens.add(allergen);
      }
    });

    (row.recipeDeclaredAllergens || []).forEach((allergen) => {
      if (VALID_ALLERGENS.has(allergen)) {
        venue.allergens.add(allergen);
        recipe.allergens.add(allergen);
      }
    });

    row.possibleHiddenAllergens.forEach((allergen) => {
      if (VALID_ALLERGENS.has(allergen)) {
        venue.possibleHidden.add(allergen);
        recipe.possibleHidden.add(allergen);
      }
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
      allergens: sortAllergens([...venue.allergens]),
      possibleHidden: sortAllergens([...venue.possibleHidden]),
      recipes: [...venue.recipesMap.values()]
        .map((recipe) => ({
          ...recipe,
          allergens: sortAllergens([...recipe.allergens]),
          possibleHidden: sortAllergens([...recipe.possibleHidden]),
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
  const visibleAllergens = sortAllergens(allergens);

  if (!visibleAllergens.length) {
    return <span style={{ color: "#777", fontSize: 13 }}>None found</span>;
  }

  return (
    <div>
      {visibleAllergens.map((allergen) => {
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
  const [selectedSubRecipeLine, setSelectedSubRecipeLine] = useState(null);

  const [posterBuilderOpen, setPosterBuilderOpen] = useState(false);
  const [posterSearch, setPosterSearch] = useState("");
  const [selectedPosterRecipeKeys, setSelectedPosterRecipeKeys] = useState([]);

  const [safeDishFinderOpen, setSafeDishFinderOpen] = useState(false);
  const [safeDishSearch, setSafeDishSearch] = useState("");
  const [selectedSafeAllergens, setSelectedSafeAllergens] = useState([]);
  const [
    includePossibleHiddenInSafeFinder,
    setIncludePossibleHiddenInSafeFinder,
  ] = useState(true);

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
    setSelectedSubRecipeLine(null);
    setPosterBuilderOpen(false);
    setPosterSearch("");
    setSelectedPosterRecipeKeys([]);
    setSafeDishFinderOpen(false);
    setSafeDishSearch("");
    setSelectedSafeAllergens([]);
    setIncludePossibleHiddenInSafeFinder(true);
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
  if (!supabase) return;

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

    return sortAllergens([...set]);
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

  const getSubRecipeRowsForLine = (subRecipeLine) => {
    if (!subRecipeLine) return [];

    const currentVenueKey = cleanText(subRecipeLine.restaurantName);

    const sourceParentRecipeKey = cleanText(
      `${subRecipeLine.recipeCode}|${subRecipeLine.recipeName}|${subRecipeLine.menuCode}|${subRecipeLine.menuName}`
    );

    const candidateNameSet = new Set(
      [
        subRecipeLine.ingredientName,
        subRecipeLine.assigned,
      ]
        .flatMap((value) => getRecipeLookupKeys(value))
        .filter(Boolean)
    );

    const candidateCodes = [
      subRecipeLine.ingredientCode,
      subRecipeLine.assigned,
      ...extractCodesFromText(subRecipeLine.ingredientName),
      ...extractCodesFromText(subRecipeLine.assigned),
    ]
      .map((value) => normalizeCode(value))
      .filter(Boolean);

    const candidateCodeSet = new Set(candidateCodes);

    const isNotSameParentRecipe = (row) => {
      const rowRecipeKey = cleanText(
        `${row.recipeCode}|${row.recipeName}|${row.menuCode}|${row.menuName}`
      );

      return rowRecipeKey !== sourceParentRecipeKey;
    };

    const sortBySourceRow = (list) =>
      [...list].sort(
        (a, b) => Number(a.sourceRow || 0) - Number(b.sourceRow || 0)
      );

    const rowNameMatchesCandidate = (value) =>
      getRecipeLookupKeys(value).some((key) => candidateNameSet.has(key));

    const recipeNameMatches = (row) => rowNameMatchesCandidate(row.recipeName);

    const menuNameMatches = (row) => rowNameMatchesCandidate(row.menuName);

    const recipeCodeMatches = (row) =>
      candidateCodeSet.size > 0 &&
      candidateCodeSet.has(normalizeCode(row.recipeCode));

    const sameVenue = (row) => cleanText(row.restaurantName) === currentVenueKey;

    const sameVenueRecipeNameRows = rows.filter(
      (row) =>
        row.sourceRow !== subRecipeLine.sourceRow &&
        isNotSameParentRecipe(row) &&
        sameVenue(row) &&
        recipeNameMatches(row)
    );

    if (sameVenueRecipeNameRows.length) {
      return sortBySourceRow(sameVenueRecipeNameRows);
    }

    const sameVenueRecipeCodeRows = rows.filter(
      (row) =>
        row.sourceRow !== subRecipeLine.sourceRow &&
        isNotSameParentRecipe(row) &&
        sameVenue(row) &&
        recipeCodeMatches(row)
    );

    if (sameVenueRecipeCodeRows.length) {
      return sortBySourceRow(sameVenueRecipeCodeRows);
    }

    const anyVenueRecipeNameRows = rows.filter(
      (row) =>
        row.sourceRow !== subRecipeLine.sourceRow &&
        isNotSameParentRecipe(row) &&
        recipeNameMatches(row)
    );

    if (anyVenueRecipeNameRows.length) {
      return sortBySourceRow(anyVenueRecipeNameRows);
    }

    const anyVenueRecipeCodeRows = rows.filter(
      (row) =>
        row.sourceRow !== subRecipeLine.sourceRow &&
        isNotSameParentRecipe(row) &&
        recipeCodeMatches(row)
    );

    if (anyVenueRecipeCodeRows.length) {
      return sortBySourceRow(anyVenueRecipeCodeRows);
    }

    const sameVenueMenuNameRows = rows.filter(
      (row) =>
        row.sourceRow !== subRecipeLine.sourceRow &&
        isNotSameParentRecipe(row) &&
        sameVenue(row) &&
        menuNameMatches(row)
    );

    if (sameVenueMenuNameRows.length) {
      return sortBySourceRow(sameVenueMenuNameRows);
    }

    return [];
  };

  const selectedSubRecipeRows = useMemo(
    () => getSubRecipeRowsForLine(selectedSubRecipeLine),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [rows, selectedSubRecipeLine]
  );

  const selectedSubRecipeAllergens = useMemo(() => {
    const found = [];

    selectedSubRecipeRows.forEach((row) => {
      found.push(...row.explicitAllergens);
      found.push(...(row.recipeDeclaredAllergens || []));
    });

    return sortAllergens(found);
  }, [selectedSubRecipeRows]);

  const selectedSubRecipePossibleHidden = useMemo(() => {
    const found = [];

    selectedSubRecipeRows.forEach((row) => {
      found.push(...row.possibleHiddenAllergens);
    });

    return sortAllergens(found);
  }, [selectedSubRecipeRows]);

  const hiddenWarningRows = useMemo(
    () => rows.filter((row) => row.hiddenWarnings.length > 0),
    [rows]
  );

  const posterRecipeOptions = useMemo(() => {
    if (!selectedVenue) return [];

    const query = posterSearch.toLowerCase().trim();

    return selectedVenue.recipes.filter((recipe) => {
      if (!query) return true;

      return [
        recipe.recipeCode,
        recipe.recipeName,
        recipe.menuCode,
        recipe.menuName,
        recipe.allergens.join(" "),
        recipe.possibleHidden.join(" "),
      ]
        .join(" ")
        .toLowerCase()
        .includes(query);
    });
  }, [selectedVenue, posterSearch]);

  const selectedPosterRecipes = useMemo(() => {
    if (!selectedVenue) return [];

    const selectedSet = new Set(selectedPosterRecipeKeys);

    return selectedVenue.recipes.filter((recipe) =>
      selectedSet.has(recipe.recipeKey)
    );
  }, [selectedVenue, selectedPosterRecipeKeys]);

  const togglePosterRecipe = (recipeKey) => {
    setSelectedPosterRecipeKeys((current) =>
      current.includes(recipeKey)
        ? current.filter((key) => key !== recipeKey)
        : [...current, recipeKey]
    );
  };

  const selectAllPosterRecipesShown = () => {
    setSelectedPosterRecipeKeys((current) => [
      ...new Set([
        ...current,
        ...posterRecipeOptions.map((recipe) => recipe.recipeKey),
      ]),
    ]);
  };

  const clearPosterRecipes = () => {
    setSelectedPosterRecipeKeys([]);
  };

  const getRecipePosterIngredients = (recipe) => {
    const combined = [...(recipe.ingredients || []), ...(recipe.subRecipes || [])];

    return combined
      .filter((item) => !item.ignoredBasic)
      .map((item) => item.ingredientName)
      .filter(Boolean)
      .slice(0, 10);
  };

  const getPosterBadgeHtml = (allergen, possible = false) => {
    const config = ALLERGEN_DISPLAY[allergen] || {
      icon: "⚠️",
      color: "#555",
    };

    const background = possible ? "#fff0f0" : "#ffffff";
    const borderColor = possible ? "#b00020" : config.color;
    const textColor = possible ? "#b00020" : config.color;

    return `
      <span
        class="poster-badge"
        style="
          border-color: ${escapeHtml(borderColor)};
          color: ${escapeHtml(textColor)};
          background: ${background};
        "
      >
        <span>${escapeHtml(config.icon || "⚠️")}</span>
        <span>${possible ? "Possible " : ""}${escapeHtml(allergen)}</span>
      </span>
    `;
  };

  const printAllergenPoster = () => {
    if (!selectedVenue) {
      window.alert("Choose a venue first.");
      return;
    }

    if (!selectedPosterRecipes.length) {
      window.alert("Choose at least one recipe for the poster.");
      return;
    }

    const allPosterAllergens = sortAllergens(
      selectedPosterRecipes.flatMap((recipe) => [
        ...(recipe.allergens || []),
        ...(recipe.possibleHidden || []),
      ])
    );

    const legendHtml = ALLERGEN_ORDER.map((allergen) => {
      const active = allPosterAllergens.includes(allergen);
      const config = ALLERGEN_DISPLAY[allergen] || {
        icon: "⚠️",
        color: "#555",
      };

      return `
        <div class="legend-pill ${active ? "legend-active" : ""}">
          <span>${escapeHtml(config.icon || "⚠️")}</span>
          <span>${escapeHtml(allergen)}</span>
        </div>
      `;
    }).join("");

    const recipeCardsHtml = selectedPosterRecipes
      .map((recipe, index) => {
        const realAllergens = sortAllergens(recipe.allergens || []);
        const possibleHidden = sortAllergens(recipe.possibleHidden || []);
        const ingredients = getRecipePosterIngredients(recipe);

        const realHtml = realAllergens.length
          ? realAllergens.map((allergen) => getPosterBadgeHtml(allergen)).join("")
          : `<span class="poster-none">No declared allergens found</span>`;

        const possibleHtml = possibleHidden.length
          ? `
            <div class="possible-row">
              <div class="poster-section-label">Possible hidden / check label</div>
              <div>${possibleHidden
                .map((allergen) => getPosterBadgeHtml(allergen, true))
                .join("")}</div>
            </div>
          `
          : "";

        const ingredientsHtml = ingredients.length
          ? ingredients
              .map((ingredient) => `<li>${escapeHtml(ingredient)}</li>`)
              .join("")
          : `<li>No ingredient detail found</li>`;

        return `
          <article class="recipe-poster-card">
            <div class="recipe-number">${index + 1}</div>

            <div class="recipe-card-header">
              <h2>${escapeHtml(recipe.recipeName || "Unnamed Recipe")}</h2>
              <div class="recipe-meta">
                Recipe ${escapeHtml(recipe.recipeCode || "N/A")}
                ${recipe.menuName ? ` • ${escapeHtml(recipe.menuName)}` : ""}
              </div>
            </div>

            <div class="poster-section-label">Allergens</div>
            <div class="badge-wrap">${realHtml}</div>

            ${possibleHtml}

            <div class="ingredients-box">
              <div class="poster-section-label">Key ingredients / sub-recipes</div>
              <ul>${ingredientsHtml}</ul>
            </div>
          </article>
        `;
      })
      .join("");

    const html = `
      <html>
        <head>
          <title>Allergen Poster - ${escapeHtml(selectedVenue.restaurantName)}</title>
          <style>
            @page { size: A4 landscape; margin: 10mm; }
            * { box-sizing: border-box; }
            body { margin: 0; padding: 0; font-family: Arial, sans-serif; color: #111; background: #f3f4f6; }
            .poster-page { min-height: 100vh; padding: 22px; background: radial-gradient(circle at top left, rgba(224,0,0,0.12), transparent 28%), linear-gradient(135deg, #ffffff 0%, #f7f7f7 48%, #ececec 100%); }
            .poster-header { display: grid; grid-template-columns: 1fr auto; gap: 20px; align-items: center; padding: 22px 24px; border-radius: 24px; background: #111; color: #fff; box-shadow: 0 12px 34px rgba(0,0,0,0.18); }
            .poster-header h1 { margin: 0; font-size: 34px; line-height: 1.05; letter-spacing: -0.8px; }
            .poster-subtitle { margin-top: 8px; font-size: 15px; opacity: 0.86; }
            .poster-count { padding: 14px 18px; border-radius: 18px; background: #fff; color: #111; font-weight: 900; text-align: center; min-width: 150px; }
            .poster-count strong { display: block; font-size: 32px; line-height: 1; }
            .legend { margin-top: 16px; padding: 14px; border-radius: 20px; background: rgba(255,255,255,0.9); display: flex; flex-wrap: wrap; gap: 7px; border: 1px solid rgba(0,0,0,0.08); }
            .legend-pill { display: inline-flex; align-items: center; gap: 5px; padding: 6px 9px; border-radius: 999px; background: #f2f2f2; color: #777; font-size: 11px; font-weight: 800; opacity: 0.45; }
            .legend-active { background: #111; color: #fff; opacity: 1; }
            .recipe-grid { margin-top: 18px; display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 14px; }
            .recipe-poster-card { position: relative; break-inside: avoid; min-height: 260px; padding: 18px; border-radius: 22px; background: #fff; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 10px 26px rgba(0,0,0,0.10); overflow: hidden; }
            .recipe-poster-card::before { content: ""; position: absolute; inset: 0 0 auto 0; height: 7px; background: linear-gradient(90deg, #e00000, #111, #e00000); }
            .recipe-number { position: absolute; top: 12px; right: 12px; width: 34px; height: 34px; border-radius: 999px; background: #111; color: #fff; display: grid; place-items: center; font-weight: 900; }
            .recipe-card-header { padding-right: 38px; }
            .recipe-card-header h2 { margin: 0 0 5px; font-size: 21px; line-height: 1.08; }
            .recipe-meta { color: #666; font-size: 12px; font-weight: 700; }
            .poster-section-label { margin: 14px 0 6px; color: #555; font-size: 11px; text-transform: uppercase; letter-spacing: 0.8px; font-weight: 900; }
            .badge-wrap { display: flex; flex-wrap: wrap; gap: 5px; }
            .poster-badge { display: inline-flex; align-items: center; gap: 5px; border: 1.5px solid; border-radius: 999px; padding: 5px 8px; font-size: 11px; font-weight: 900; line-height: 1.1; }
            .poster-none { color: #777; font-size: 12px; font-weight: 700; }
            .possible-row { padding: 9px; border-radius: 14px; background: #fff7f7; border: 1px solid #ffd1d1; margin-top: 10px; }
            .ingredients-box { margin-top: 10px; padding: 10px; border-radius: 15px; background: #f7f7f7; }
            .ingredients-box ul { margin: 0; padding-left: 18px; columns: 2; column-gap: 20px; }
            .ingredients-box li { font-size: 11px; margin-bottom: 3px; break-inside: avoid; }
            .poster-footer { margin-top: 18px; padding: 12px 16px; border-radius: 18px; background: #fff4d6; color: #8a5a00; font-size: 13px; font-weight: 800; border: 1px solid #f1d28a; }
            .no-print { position: fixed; right: 16px; bottom: 16px; display: flex; gap: 8px; }
            .no-print button { border: 0; border-radius: 999px; background: #111; color: #fff; padding: 12px 16px; cursor: pointer; font-weight: 900; box-shadow: 0 8px 24px rgba(0,0,0,0.22); }
            @media print {
              body { background: #fff; }
              .poster-page { padding: 0; background: #fff; }
              .no-print { display: none; }
              .recipe-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); }
              .recipe-poster-card { box-shadow: none; }
            }
          </style>
        </head>

        <body>
          <main class="poster-page">
            <section class="poster-header">
              <div>
                <h1>Allergen Poster</h1>
                <div class="poster-subtitle">
                  ${escapeHtml(selectedVenue.restaurantName)}
                  • Generated ${escapeHtml(new Date().toLocaleString())}
                </div>
              </div>

              <div class="poster-count">
                <strong>${selectedPosterRecipes.length}</strong>
                recipe${selectedPosterRecipes.length === 1 ? "" : "s"}
              </div>
            </section>

            <section class="legend">${legendHtml}</section>
            <section class="recipe-grid">${recipeCardsHtml}</section>

            <section class="poster-footer">
              Support tool only. Always verify against official recipe cards,
              supplier labels, and onboard allergy procedures before answering a
              Sailor allergy request.
            </section>
          </main>

          <div class="no-print">
            <button onclick="window.print()">🖨️ Print Poster</button>
            <button onclick="window.close()">Close</button>
          </div>

          <script>
            window.setTimeout(() => window.print(), 450);
          </script>
        </body>
      </html>
    `;

    const printWindow = window.open("", "_blank");

    if (!printWindow) {
      window.alert("Print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();

    logUsageEvent("allergen_poster_printed", {
      module: "allergen",
      ship: userShip,
      venue: selectedVenue.restaurantName,
      recipeCount: selectedPosterRecipes.length,
      allergens: allPosterAllergens,
    });
  };

  const toggleSafeAllergen = (allergen) => {
    setSelectedSafeAllergens((current) =>
      current.includes(allergen)
        ? current.filter((item) => item !== allergen)
        : [...current, allergen]
    );
  };

  const clearSafeAllergens = () => {
    setSelectedSafeAllergens([]);
    setSafeDishSearch("");
  };

  const selectedSafeAllergenSet = useMemo(
    () => new Set(selectedSafeAllergens),
    [selectedSafeAllergens]
  );

  const safeDishRows = useMemo(() => {
    if (!selectedVenue || !selectedSafeAllergens.length) return [];

    const query = safeDishSearch.toLowerCase().trim();

    return selectedVenue.recipes
      .map((recipe) => {
        const recipeAllergens = sortAllergens(recipe.allergens || []);
        const possibleHidden = sortAllergens(recipe.possibleHidden || []);

        const blockingDeclared = recipeAllergens.filter((allergen) =>
          selectedSafeAllergenSet.has(allergen)
        );

        const blockingPossible = includePossibleHiddenInSafeFinder
          ? possibleHidden.filter((allergen) =>
              selectedSafeAllergenSet.has(allergen)
            )
          : [];

        const blockedAllergens = sortAllergens([
          ...blockingDeclared,
          ...blockingPossible,
        ]);

        const available = blockedAllergens.length === 0;

        return {
          recipe,
          recipeKey: recipe.recipeKey,
          recipeName: recipe.recipeName || "",
          recipeCode: recipe.recipeCode || "",
          menuName: recipe.menuName || "",
          allergens: recipeAllergens,
          possibleHidden,
          blockingDeclared,
          blockingPossible,
          blockedAllergens,
          available,
          ingredientCount:
            Number(recipe.ingredients?.length || 0) +
            Number(recipe.subRecipes?.length || 0),
        };
      })
      .filter((item) => item.available)
      .filter((item) => {
        if (!query) return true;

        return [
          item.recipeName,
          item.recipeCode,
          item.menuName,
          item.allergens.join(" "),
          item.possibleHidden.join(" "),
        ]
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort(
        (a, b) =>
          a.menuName.localeCompare(b.menuName) ||
          a.recipeName.localeCompare(b.recipeName)
      );
  }, [
    selectedVenue,
    selectedSafeAllergens,
    selectedSafeAllergenSet,
    safeDishSearch,
    includePossibleHiddenInSafeFinder,
  ]);

  const blockedDishRows = useMemo(() => {
    if (!selectedVenue || !selectedSafeAllergens.length) return [];

    return selectedVenue.recipes
      .map((recipe) => {
        const recipeAllergens = sortAllergens(recipe.allergens || []);
        const possibleHidden = sortAllergens(recipe.possibleHidden || []);

        const blockingDeclared = recipeAllergens.filter((allergen) =>
          selectedSafeAllergenSet.has(allergen)
        );

        const blockingPossible = includePossibleHiddenInSafeFinder
          ? possibleHidden.filter((allergen) =>
              selectedSafeAllergenSet.has(allergen)
            )
          : [];

        const blockedAllergens = sortAllergens([
          ...blockingDeclared,
          ...blockingPossible,
        ]);

        return {
          recipe,
          recipeName: recipe.recipeName || "",
          recipeCode: recipe.recipeCode || "",
          menuName: recipe.menuName || "",
          blockingDeclared,
          blockingPossible,
          blockedAllergens,
          available: blockedAllergens.length === 0,
        };
      })
      .filter((item) => !item.available)
      .sort(
        (a, b) =>
          a.menuName.localeCompare(b.menuName) ||
          a.recipeName.localeCompare(b.recipeName)
      );
  }, [
    selectedVenue,
    selectedSafeAllergens,
    selectedSafeAllergenSet,
    includePossibleHiddenInSafeFinder,
  ]);

  const exportSafeDishFinderToExcel = () => {
    if (!selectedVenue) {
      window.alert("Choose a venue first.");
      return;
    }

    if (!selectedSafeAllergens.length) {
      window.alert("Choose at least one allergen.");
      return;
    }

    if (!safeDishRows.length) {
      window.alert("No available dishes found for the selected allergens.");
      return;
    }

    const exportRows = safeDishRows.map((item, index) => ({
      Number: index + 1,
      Venue: selectedVenue.restaurantName,
      RecipeCode: item.recipeCode,
      RecipeName: item.recipeName,
      MenuName: item.menuName,
      AvoidingAllergens: selectedSafeAllergens.join(", "),
      Available: "Yes",
      RecipeAllergensFound: item.allergens.join(", "),
      PossibleHiddenAllergensFound: item.possibleHidden.join(", "),
      PossibleHiddenIncludedInCheck: includePossibleHiddenInSafeFinder
        ? "Yes"
        : "No",
      IngredientLines: item.ingredientCount,
    }));

    exportRowsToExcel(
      exportRows,
      "Available Dishes",
      `available-dishes-${selectedVenue.restaurantName
        .replace(/[^a-z0-9]+/gi, "-")
        .toLowerCase()
        .slice(0, 40)}-${userShip || "ship"}.xlsx`
    );

    logUsageEvent("safe_dish_finder_exported", {
      module: "allergen",
      ship: userShip,
      venue: selectedVenue.restaurantName,
      allergens: selectedSafeAllergens,
      availableDishes: safeDishRows.length,
      blockedDishes: blockedDishRows.length,
      includePossibleHidden: includePossibleHiddenInSafeFinder,
    });
  };

  const printSafeDishFinder = () => {
    if (!selectedVenue) {
      window.alert("Choose a venue first.");
      return;
    }

    if (!selectedSafeAllergens.length) {
      window.alert("Choose at least one allergen.");
      return;
    }

    if (!safeDishRows.length) {
      window.alert("No available dishes found for the selected allergens.");
      return;
    }

    const allergenText = selectedSafeAllergens.join(", ");

    const rowsHtml = safeDishRows
      .map(
        (item, index) => `
          <tr>
            <td>${index + 1}</td>
            <td>${escapeHtml(item.recipeName)}</td>
            <td>${escapeHtml(item.recipeCode || "N/A")}</td>
            <td>${escapeHtml(item.menuName || "N/A")}</td>
            <td>${escapeHtml(item.allergens.join(", ") || "None found")}</td>
            <td>${escapeHtml(item.possibleHidden.join(", ") || "None found")}</td>
          </tr>
        `
      )
      .join("");

    const html = `
      <html>
        <head>
          <title>Available Dishes - ${escapeHtml(selectedVenue.restaurantName)}</title>
          <style>
            body { font-family: Arial, sans-serif; padding: 24px; color: #111; }
            h1 { margin-bottom: 4px; }
            .meta { margin: 4px 0; color: #555; font-weight: bold; }
            .warning { margin-top: 14px; padding: 12px; border-radius: 12px; background: #fff4d6; color: #8a5a00; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
            th, td { border: 1px solid #ccc; padding: 7px; text-align: left; vertical-align: top; }
            th { background: #f2f2f2; }
            .good { color: #2e7d32; font-weight: bold; }
          </style>
        </head>

        <body>
          <h1>Available Dishes</h1>
          <div class="meta">Venue: ${escapeHtml(selectedVenue.restaurantName)}</div>
          <div class="meta">Avoiding: ${escapeHtml(allergenText)}</div>
          <div class="meta">Available dishes: ${safeDishRows.length}</div>
          <div class="meta">Blocked dishes: ${blockedDishRows.length}</div>
          <div class="meta">Possible hidden allergens included: ${
            includePossibleHiddenInSafeFinder ? "Yes" : "No"
          }</div>

          <div class="warning">
            Support tool only. Always verify with official recipe cards, supplier labels,
            and onboard allergy procedures before confirming a dish is safe.
          </div>

          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Dish / Recipe</th>
                <th>Recipe Code</th>
                <th>Menu</th>
                <th>Recipe Allergens Found</th>
                <th>Possible Hidden Allergens Found</th>
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </body>
      </html>
    `;

    const printWindow = window.open("", "_blank");

    if (!printWindow) {
      window.alert("Print window was blocked. Allow popups and try again.");
      return;
    }

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.focus();
    printWindow.print();

    logUsageEvent("safe_dish_finder_printed", {
      module: "allergen",
      ship: userShip,
      venue: selectedVenue.restaurantName,
      allergens: selectedSafeAllergens,
      availableDishes: safeDishRows.length,
      blockedDishes: blockedDishRows.length,
      includePossibleHidden: includePossibleHiddenInSafeFinder,
    });
  };

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
      Type: isSubRecipeRow(row) ? "Sub Recipe" : "Ingredient",
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
      GlutenFreeProductClaim: row.glutenFreeProductClaim ? "Yes" : "No",
      PossibleAllergensSkipped: row.skipPossibleAllergens ? "Yes" : "No",
      OnboardRecipeLine: row.onboardRecipeLine ? "Yes" : "No",
      SourceRow: row.sourceRow,
    }));

    exportRowsToExcel(
      exportRows,
      "Allergen Matrix",
      `allergen-matrix-${userShip || "ship"}.xlsx`
    );
  };

  const getProductAllergenReportRows = () => {
    const productMap = new Map();

    rows.forEach((row) => {
      const venue = row.restaurantName || "Unknown Venue";
      const productCode = row.ingredientCode || "";
      const productName = row.ingredientName || "";

      if (!productName) return;

      const productKey = [
        cleanText(venue),
        normalizeCode(productCode) || cleanText(productName),
        cleanText(productName),
      ].join("|");

      if (!productMap.has(productKey)) {
        productMap.set(productKey, {
          Venue: venue,
          ProductCode: productCode,
          ProductName: productName,
          ProductType: isSubRecipeRow(row) ? "Sub Recipe" : "Ingredient",
          realAllergens: new Set(),
          possibleHiddenAllergens: new Set(),
          recipeDeclaredAllergens: new Set(),
          recipes: new Set(),
          menus: new Set(),
          sourceRows: new Set(),
          ignoredBasic: false,
          plainRawIngredient: false,
          glutenFreeProductClaim: false,
          possibleAllergensSkipped: false,
          onboardRecipeLine: false,
        });
      }

      const product = productMap.get(productKey);

      row.explicitAllergens.forEach((allergen) => {
        if (VALID_ALLERGENS.has(allergen)) {
          product.realAllergens.add(allergen);
        }
      });

      row.possibleHiddenAllergens.forEach((allergen) => {
        if (VALID_ALLERGENS.has(allergen)) {
          product.possibleHiddenAllergens.add(allergen);
        }
      });

      (row.recipeDeclaredAllergens || []).forEach((allergen) => {
        if (VALID_ALLERGENS.has(allergen)) {
          product.recipeDeclaredAllergens.add(allergen);
        }
      });

      if (row.recipeCode || row.recipeName) {
        product.recipes.add(
          `${row.recipeCode || "N/A"} - ${row.recipeName || "Unnamed Recipe"}`
        );
      }

      if (row.menuName) product.menus.add(row.menuName);
      if (row.sourceRow) product.sourceRows.add(row.sourceRow);

      product.ignoredBasic = product.ignoredBasic || row.ignoredBasic;
      product.plainRawIngredient =
        product.plainRawIngredient || row.plainRawIngredient;
      product.glutenFreeProductClaim =
        product.glutenFreeProductClaim || row.glutenFreeProductClaim;
      product.possibleAllergensSkipped =
        product.possibleAllergensSkipped || row.skipPossibleAllergens;
      product.onboardRecipeLine =
        product.onboardRecipeLine || row.onboardRecipeLine;
    });

    return [...productMap.values()]
      .map((product, index) => {
        const realAllergens = sortAllergens([...product.realAllergens]);
        const possibleHiddenAllergens = sortAllergens([
          ...product.possibleHiddenAllergens,
        ]);
        const recipeDeclaredAllergens = sortAllergens([
          ...product.recipeDeclaredAllergens,
        ]);

        const productAllergensToReview = sortAllergens([
          ...realAllergens,
          ...possibleHiddenAllergens,
        ]);

        return {
          Number: index + 1,
          Venue: product.Venue,
          ProductCode: product.ProductCode,
          ProductName: product.ProductName,
          ProductType: product.ProductType,
          ProductAllergensToReview: productAllergensToReview.join(", "),
          RealDeclaredOrDetectedAllergens: realAllergens.join(", "),
          PossibleHiddenAllergens: possibleHiddenAllergens.join(", "),
          RecipeLevelDeclaredAllergensForReference:
            recipeDeclaredAllergens.join(", "),
          RecipesUsingProduct: [...product.recipes].sort().join(" | "),
          Menus: [...product.menus].sort().join(" | "),
          SourceRows: [...product.sourceRows].sort((a, b) => a - b).join(", "),
          IgnoredBasic: product.ignoredBasic ? "Yes" : "No",
          FreshRawItemPossibleAllergensSkipped: product.plainRawIngredient
            ? "Yes"
            : "No",
          GlutenFreeProductClaim: product.glutenFreeProductClaim ? "Yes" : "No",
          PossibleAllergensSkipped: product.possibleAllergensSkipped
            ? "Yes"
            : "No",
          OnboardRecipeLine: product.onboardRecipeLine ? "Yes" : "No",
        };
      })
      .sort(
        (a, b) =>
          a.Venue.localeCompare(b.Venue) ||
          a.ProductName.localeCompare(b.ProductName)
      );
  };

  const exportProductAllergenReport = () => {
    const reportRows = getProductAllergenReportRows();

    if (!reportRows.length) {
      window.alert("No product allergen rows to export.");
      return;
    }

    exportRowsToExcel(
      reportRows,
      "Product Allergens",
      `product-allergens-by-location-${userShip || "ship"}.xlsx`
    );

    logUsageEvent("product_allergen_report_exported", {
      module: "allergen",
      ship: userShip,
      rows: reportRows.length,
    });
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
            This support tool lists only the 14 required allergen groups:
            cereals containing gluten, crustaceans, eggs, fish, peanuts,
            soybeans, milk, tree nuts, celery, mustard, sesame seeds, sulphur
            dioxide and sulphites, lupin, and molluscs. Gluten-free product
            names such as “Gluten Free Cookies” or “GF Bread” will not be marked
            as cereals containing gluten from keyword matching. Fresh vegetables,
            fruits, herbs, and produce items such as butternut squash are not
            treated as pre-made products. Always verify against official recipe
            cards, supplier labels, and onboard allergy procedures before
            answering a Sailor allergy request.
          </div>

          <button
            style={styles.primaryButton}
            onClick={exportRecipeMatrix}
            disabled={!rows.length}
          >
            📥 Export Full Allergen Matrix
          </button>

          <button
            style={styles.primaryButton}
            onClick={exportProductAllergenReport}
            disabled={!rows.length}
          >
            📥 Export Product Allergen Report
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
                setSelectedSubRecipeLine(null);
                setPosterBuilderOpen(false);
                setPosterSearch("");
                setSelectedPosterRecipeKeys([]);
                setSafeDishFinderOpen(false);
                setSafeDishSearch("");
                setSelectedSafeAllergens([]);
                setIncludePossibleHiddenInSafeFinder(true);
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

            <div style={styles.headerActions}>
              <button
                type="button"
                style={styles.backButton}
                onClick={() => {
                  setSafeDishFinderOpen((current) => !current);
                  setPosterBuilderOpen(false);
                  setSelectedSubRecipeLine(null);
                }}
              >
                🛡️ Safe Dish Finder
              </button>

              <button
                type="button"
                style={styles.backButton}
                onClick={() => {
                  setPosterBuilderOpen((current) => !current);
                  setSafeDishFinderOpen(false);
                  setSelectedSubRecipeLine(null);
                }}
              >
                🎨 Poster Builder
              </button>

              <div style={styles.shipBadge}>{filteredRecipes.length} recipe(s)</div>
            </div>
          </div>

          {safeDishFinderOpen && (
            <section style={localStyles.safeDishFinderBox}>
              <div style={localStyles.safeDishHeader}>
                <div>
                  <h3 style={localStyles.safeDishTitle}>
                    🛡️ Safe Dish Finder
                  </h3>

                  <div style={localStyles.safeDishSubtext}>
                    Choose allergen(s), then see which dishes in{" "}
                    {selectedVenue.restaurantName} do not contain them.
                  </div>
                </div>

                <div style={localStyles.safeDishCount}>
                  {selectedSafeAllergens.length
                    ? `${safeDishRows.length} available`
                    : "Choose allergen(s)"}
                </div>
              </div>

              <div style={localStyles.safeAllergenGrid}>
                {ALLERGEN_ORDER.map((allergen) => {
                  const selected = selectedSafeAllergens.includes(allergen);
                  const config = ALLERGEN_DISPLAY[allergen] || {
                    icon: "⚠️",
                    color: "#555",
                  };

                  return (
                    <button
                      key={`safe-allergen-${allergen}`}
                      type="button"
                      style={{
                        ...localStyles.safeAllergenButton,
                        ...(selected ? localStyles.safeAllergenButtonActive : {}),
                      }}
                      onClick={() => toggleSafeAllergen(allergen)}
                    >
                      <span>{config.icon}</span>
                      <span>{allergen}</span>
                    </button>
                  );
                })}
              </div>

              <label style={styles.checkboxRow}>
                <input
                  type="checkbox"
                  checked={includePossibleHiddenInSafeFinder}
                  onChange={(event) =>
                    setIncludePossibleHiddenInSafeFinder(event.target.checked)
                  }
                />
                <span>
                  Exclude dishes with possible hidden allergens too. Recommended for allergy requests.
                </span>
              </label>

              <input
                placeholder="Search available dishes..."
                value={safeDishSearch}
                onChange={(event) => setSafeDishSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.headerActions}>
                <button
                  type="button"
                  style={styles.backButton}
                  onClick={clearSafeAllergens}
                  disabled={!selectedSafeAllergens.length && !safeDishSearch}
                >
                  Clear
                </button>

                <button
                  type="button"
                  style={styles.backButton}
                  onClick={printSafeDishFinder}
                  disabled={!safeDishRows.length}
                >
                  🖨️ Print Available Dishes
                </button>

                <button
                  type="button"
                  style={styles.primaryButton}
                  onClick={exportSafeDishFinderToExcel}
                  disabled={!safeDishRows.length}
                >
                  📥 Export Available Dishes
                </button>
              </div>

              {selectedSafeAllergens.length > 0 && (
                <div style={styles.infoBox}>
                  <div>
                    🚫 Avoiding: <strong>{selectedSafeAllergens.join(", ")}</strong>
                  </div>
                  <div>
                    ✅ Available dishes: <strong>{safeDishRows.length}</strong>
                  </div>
                  <div>
                    ❌ Blocked dishes: <strong>{blockedDishRows.length}</strong>
                  </div>
                  <div>
                    🕵️ Possible hidden allergens included in check:{" "}
                    <strong>
                      {includePossibleHiddenInSafeFinder ? "Yes" : "No"}
                    </strong>
                  </div>
                </div>
              )}

              {!selectedSafeAllergens.length && (
                <p style={styles.emptyText}>
                  Select one or more allergens above to see available dishes.
                </p>
              )}

              {selectedSafeAllergens.length > 0 && !safeDishRows.length && (
                <p style={styles.warningText}>
                  No available dishes found for the selected allergen combination.
                  Verify with official recipe cards and onboard allergy procedures.
                </p>
              )}

              <div style={localStyles.safeDishGrid}>
                {safeDishRows.map((item) => (
                  <button
                    key={`safe-dish-${item.recipeKey}`}
                    type="button"
                    style={localStyles.safeDishCard}
                    onClick={() => {
                      setSelectedRecipeKey(item.recipeKey);
                      setSelectedSubRecipeLine(null);
                      setIngredientSearch("");
                    }}
                  >
                    <div style={localStyles.safeDishGoodBadge}>✅ Available</div>

                    <strong>{item.recipeName}</strong>

                    <span>Recipe code: {item.recipeCode || "N/A"}</span>

                    {item.menuName && <span>Menu: {item.menuName}</span>}

                    <div style={localStyles.safeDishMeta}>
                      Allergens found: {item.allergens.join(", ") || "None found"}
                    </div>

                    {includePossibleHiddenInSafeFinder && (
                      <div style={localStyles.safeDishMeta}>
                        Possible hidden:{" "}
                        {item.possibleHidden.join(", ") || "None found"}
                      </div>
                    )}

                    <div style={localStyles.safeDishOpenHint}>
                      Click to open recipe details
                    </div>
                  </button>
                ))}
              </div>
            </section>
          )}

          {posterBuilderOpen && (
            <section style={localStyles.posterBuilderBox}>
              <div style={localStyles.posterBuilderHeader}>
                <div>
                  <h3 style={localStyles.posterBuilderTitle}>
                    🎨 Build Allergen Poster
                  </h3>

                  <div style={localStyles.posterBuilderSubtext}>
                    Select recipes from {selectedVenue.restaurantName}, then print
                    one combined poster.
                  </div>
                </div>

                <div style={localStyles.posterSelectedCount}>
                  {selectedPosterRecipes.length} selected
                </div>
              </div>

              <input
                placeholder="Search recipes for poster..."
                value={posterSearch}
                onChange={(event) => setPosterSearch(event.target.value)}
                style={styles.searchInput}
              />

              <div style={styles.headerActions}>
                <button
                  type="button"
                  style={styles.backButton}
                  onClick={selectAllPosterRecipesShown}
                  disabled={!posterRecipeOptions.length}
                >
                  Select Shown
                </button>

                <button
                  type="button"
                  style={styles.backButton}
                  onClick={clearPosterRecipes}
                  disabled={!selectedPosterRecipeKeys.length}
                >
                  Clear
                </button>

                <button
                  type="button"
                  style={styles.primaryButton}
                  onClick={printAllergenPoster}
                  disabled={!selectedPosterRecipeKeys.length}
                >
                  🖨️ Print Allergen Poster
                </button>
              </div>

              <div style={localStyles.posterRecipePickerGrid}>
                {posterRecipeOptions.map((recipe) => {
                  const selected = selectedPosterRecipeKeys.includes(
                    recipe.recipeKey
                  );

                  const posterAllergens = sortAllergens([
                    ...(recipe.allergens || []),
                    ...(recipe.possibleHidden || []),
                  ]);

                  return (
                    <label
                      key={`poster-${recipe.recipeKey}`}
                      style={{
                        ...localStyles.posterRecipeOption,
                        ...(selected
                          ? localStyles.posterRecipeOptionActive
                          : {}),
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={selected}
                        onChange={() => togglePosterRecipe(recipe.recipeKey)}
                      />

                      <div style={localStyles.posterRecipeText}>
                        <strong>{recipe.recipeName}</strong>
                        <span>Recipe code: {recipe.recipeCode || "N/A"}</span>
                        <span>
                          {posterAllergens.join(", ") || "No allergens found"}
                        </span>
                      </div>
                    </label>
                  );
                })}
              </div>

              {!posterRecipeOptions.length && (
                <p style={styles.emptyText}>No recipes match this poster search.</p>
              )}
            </section>
          )}

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
                  setSelectedSubRecipeLine(null);
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
            setSelectedSubRecipeLine(null);
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
                setSelectedSubRecipeLine(null);
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
                <AllergenBadges
                  allergens={selectedRecipe.possibleHidden}
                  possible
                />
              </div>
            </section>

            <input
              placeholder="Search ingredient or allergen..."
              value={ingredientSearch}
              onChange={(event) => setIngredientSearch(event.target.value)}
              style={{ ...styles.searchInput, marginTop: 14 }}
            />

            <div style={localStyles.ingredientGrid}>
              {visibleIngredients.map((item) => {
                const itemIsSubRecipe = isSubRecipeRow(item);

                return (
                  <div
                    key={`${item.sourceRow}-${item.ingredientCode}-${item.ingredientName}`}
                    role={itemIsSubRecipe ? "button" : undefined}
                    tabIndex={itemIsSubRecipe ? 0 : undefined}
                    onClick={() => {
                      if (itemIsSubRecipe) {
                        setSelectedSubRecipeLine(item);
                      }
                    }}
                    onKeyDown={(event) => {
                      if (
                        itemIsSubRecipe &&
                        (event.key === "Enter" || event.key === " ")
                      ) {
                        event.preventDefault();
                        setSelectedSubRecipeLine(item);
                      }
                    }}
                    style={{
                      ...localStyles.ingredientCard,
                      ...(itemIsSubRecipe
                        ? localStyles.clickableIngredientCard
                        : {}),
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
                      {itemIsSubRecipe ? "Sub Recipe" : "Ingredient"}
                    </div>

                    {itemIsSubRecipe && (
                      <button
                        type="button"
                        style={localStyles.subRecipeButton}
                        onClick={(event) => {
                          event.stopPropagation();
                          setSelectedSubRecipeLine(item);
                        }}
                      >
                        🔎 Open Sub Recipe Ingredients
                      </button>
                    )}

                    {item.ignoredBasic && (
                      <div style={styles.statusNeutral}>Ignored basic item</div>
                    )}

                    {item.plainRawIngredient && !item.ignoredBasic && (
                      <div style={styles.statusNeutral}>
                        Fresh/raw item — possible allergens skipped
                      </div>
                    )}

                    {item.glutenFreeProductClaim && (
                      <div style={styles.statusNeutral}>
                        Gluten-free product claim
                      </div>
                    )}

                    {item.onboardRecipeLine && (
                      <div style={styles.statusNeutral}>
                        Onboard recipe item
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
                );
              })}
            </div>

            {selectedSubRecipeLine && (
              <div
                style={localStyles.subRecipeModalBackdrop}
                onClick={(event) => {
                  event.stopPropagation();
                  setSelectedSubRecipeLine(null);
                }}
              >
                <div
                  style={localStyles.subRecipeModalCard}
                  onClick={(event) => event.stopPropagation()}
                >
                  <button
                    type="button"
                    style={styles.closeButton}
                    onClick={() => setSelectedSubRecipeLine(null)}
                  >
                    ✕
                  </button>

                  <div style={localStyles.modalHeader}>
                    <div>
                      <h2 style={{ ...styles.productTitle, marginBottom: 4 }}>
                        🧾 {selectedSubRecipeLine.ingredientName}
                      </h2>

                      <p style={{ ...styles.emptyText, margin: 0 }}>
                        Sub recipe opened from{" "}
                        <strong>{selectedRecipe.recipeName}</strong>
                      </p>

                      <div style={localStyles.subRecipeInfoLine}>
                        Product code:{" "}
                        <strong>
                          {selectedSubRecipeLine.ingredientCode || "N/A"}
                        </strong>
                      </div>
                    </div>

                    <div style={styles.statusNeutral}>
                      {selectedSubRecipeRows.length} row(s)
                    </div>
                  </div>

                  <section style={localStyles.modalSummaryGrid}>
                    <div style={styles.infoBox}>
                      <strong>Sub recipe allergens</strong>
                      <AllergenBadges allergens={selectedSubRecipeAllergens} />
                    </div>

                    <div style={styles.infoBox}>
                      <strong>Possible hidden allergens</strong>
                      <AllergenBadges
                        allergens={selectedSubRecipePossibleHidden}
                        possible
                      />
                    </div>
                  </section>

                  {!selectedSubRecipeRows.length && (
                    <div style={styles.warningText}>
                      No ingredient rows were found for this sub recipe in the
                      Ingredient by Location file. Check that the sub recipe name
                      or code matches the recipe name/code in the file.
                    </div>
                  )}

                  <div style={localStyles.ingredientGrid}>
                    {selectedSubRecipeRows.map((subItem) => {
                      const subItemIsSubRecipe = isSubRecipeRow(subItem);

                      return (
                        <div
                          key={`${subItem.sourceRow}-${subItem.ingredientCode}-${subItem.ingredientName}`}
                          role={subItemIsSubRecipe ? "button" : undefined}
                          tabIndex={subItemIsSubRecipe ? 0 : undefined}
                          onClick={() => {
                            if (subItemIsSubRecipe) {
                              setSelectedSubRecipeLine(subItem);
                            }
                          }}
                          onKeyDown={(event) => {
                            if (
                              subItemIsSubRecipe &&
                              (event.key === "Enter" || event.key === " ")
                            ) {
                              event.preventDefault();
                              setSelectedSubRecipeLine(subItem);
                            }
                          }}
                          style={{
                            ...localStyles.ingredientCard,
                            ...(subItemIsSubRecipe
                              ? localStyles.clickableIngredientCard
                              : {}),
                            ...(subItem.hiddenWarnings.length > 0
                              ? localStyles.ingredientCardWarning
                              : {}),
                            ...(subItem.possibleHiddenAllergens.length > 0 &&
                            !subItem.hiddenWarnings.length
                              ? localStyles.ingredientCardPossible
                              : {}),
                          }}
                        >
                          <div style={localStyles.ingredientTitle}>
                            {subItem.ingredientName}
                          </div>

                          {subItem.assigned &&
                            subItem.assigned !== subItem.ingredientName && (
                              <div style={localStyles.compactMeta}>
                                Product / Assigned: {subItem.assigned}
                              </div>
                            )}

                          <div style={localStyles.typePill}>
                            {subItemIsSubRecipe
                              ? "Nested Sub Recipe"
                              : "Ingredient"}
                          </div>

                          {subItemIsSubRecipe && (
                            <button
                              type="button"
                              style={localStyles.subRecipeButton}
                              onClick={(event) => {
                                event.stopPropagation();
                                setSelectedSubRecipeLine(subItem);
                              }}
                            >
                              🔎 Open Nested Sub Recipe
                            </button>
                          )}

                          {subItem.ignoredBasic && (
                            <div style={styles.statusNeutral}>
                              Ignored basic item
                            </div>
                          )}

                          {subItem.plainRawIngredient && !subItem.ignoredBasic && (
                            <div style={styles.statusNeutral}>
                              Fresh/raw item — possible allergens skipped
                            </div>
                          )}

                          {subItem.glutenFreeProductClaim && (
                            <div style={styles.statusNeutral}>
                              Gluten-free product claim
                            </div>
                          )}

                          {subItem.onboardRecipeLine && (
                            <div style={styles.statusNeutral}>
                              Onboard recipe item
                            </div>
                          )}

                          <div>
                            <strong>Declared / real:</strong>
                            <AllergenBadges allergens={subItem.explicitAllergens} />
                          </div>

                          {subItem.possibleHiddenAllergens.length > 0 && (
                            <div>
                              <strong>Possible hidden:</strong>
                              <AllergenBadges
                                allergens={subItem.possibleHiddenAllergens}
                                possible
                              />
                            </div>
                          )}

                          {subItem.hiddenWarnings.map((warning, index) => (
                            <div
                              key={`${warning}-${index}`}
                              style={styles.statusBad}
                            >
                              ⚠️ {warning}
                            </div>
                          ))}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            )}
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

  clickableIngredientCard: {
    cursor: "pointer",
    boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
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

  subRecipeButton: {
    border: "1px solid #111",
    borderRadius: 999,
    background: "#111",
    color: "#fff",
    padding: "7px 9px",
    fontSize: 11,
    fontWeight: "bold",
    cursor: "pointer",
    marginTop: 4,
  },

  subRecipeModalBackdrop: {
    position: "fixed",
    inset: 0,
    background: "rgba(0,0,0,0.62)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: 10050,
    padding: 18,
  },

  subRecipeModalCard: {
    background: "#fff",
    borderRadius: 18,
    padding: 20,
    maxWidth: 1080,
    width: "96%",
    maxHeight: "88vh",
    overflowY: "auto",
    position: "relative",
    boxShadow: "0 24px 70px rgba(0,0,0,0.35)",
  },

  subRecipeInfoLine: {
    marginTop: 6,
    color: "#555",
    fontSize: 13,
  },

  posterBuilderBox: {
    marginBottom: 18,
    padding: 16,
    borderRadius: 18,
    background: "linear-gradient(135deg, #ffffff 0%, #f7f7f7 100%)",
    border: "1px solid #ddd",
    boxShadow: "0 8px 22px rgba(0,0,0,0.06)",
  },

  posterBuilderHeader: {
    display: "flex",
    justifyContent: "space-between",
    gap: 12,
    alignItems: "flex-start",
    marginBottom: 12,
    flexWrap: "wrap",
  },

  posterBuilderTitle: {
    margin: 0,
    fontSize: 20,
  },

  posterBuilderSubtext: {
    marginTop: 4,
    color: "#666",
    fontSize: 13,
    fontWeight: "bold",
  },

  posterSelectedCount: {
    padding: "9px 12px",
    borderRadius: 999,
    background: "#111",
    color: "#fff",
    fontWeight: "bold",
  },

  posterRecipePickerGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: 8,
    marginTop: 12,
  },

  posterRecipeOption: {
    border: "1px solid #ddd",
    borderRadius: 14,
    background: "#fff",
    padding: 10,
    display: "grid",
    gridTemplateColumns: "auto 1fr",
    gap: 8,
    cursor: "pointer",
    alignItems: "flex-start",
  },

  posterRecipeOptionActive: {
    border: "2px solid #111",
    background: "#f2f2f2",
  },

  posterRecipeText: {
    display: "grid",
    gap: 3,
    fontSize: 12,
    color: "#555",
  },

  safeDishFinderBox: {
    marginBottom: 18,
    padding: 16,
    borderRadius: 18,
    background: "linear-gradient(135deg, #f0fff4 0%, #ffffff 100%)",
    border: "1px solid #bde5c8",
    boxShadow: "0 8px 22px rgba(0,0,0,0.06)",
  },

  safeDishHeader: {
    display: "flex",
    justifyContent: "space-between",
    gap: 12,
    alignItems: "flex-start",
    marginBottom: 12,
    flexWrap: "wrap",
  },

  safeDishTitle: {
    margin: 0,
    fontSize: 20,
  },

  safeDishSubtext: {
    marginTop: 4,
    color: "#2e7d32",
    fontSize: 13,
    fontWeight: "bold",
  },

  safeDishCount: {
    padding: "9px 12px",
    borderRadius: 999,
    background: "#2e7d32",
    color: "#fff",
    fontWeight: "bold",
  },

  safeAllergenGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))",
    gap: 8,
    marginBottom: 12,
  },

  safeAllergenButton: {
    border: "1px solid #ddd",
    borderRadius: 999,
    background: "#fff",
    padding: "9px 11px",
    display: "flex",
    gap: 7,
    alignItems: "center",
    cursor: "pointer",
    fontWeight: "bold",
    color: "#111",
    textAlign: "left",
  },

  safeAllergenButtonActive: {
    border: "2px solid #2e7d32",
    background: "#e8f5e9",
    color: "#2e7d32",
  },

  safeDishGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fill, minmax(220px, 1fr))",
    gap: 10,
    marginTop: 14,
  },

  safeDishCard: {
    border: "2px solid #2e7d32",
    borderRadius: 16,
    background: "#fff",
    padding: 12,
    display: "grid",
    gap: 5,
    textAlign: "left",
    cursor: "pointer",
    fontFamily: "inherit",
    color: "#111",
    boxShadow: "0 5px 16px rgba(0,0,0,0.06)",
  },

  safeDishGoodBadge: {
    justifySelf: "start",
    padding: "5px 8px",
    borderRadius: 999,
    background: "#e8f5e9",
    color: "#2e7d32",
    fontWeight: "bold",
    fontSize: 12,
  },

  safeDishMeta: {
    color: "#555",
    fontSize: 12,
    lineHeight: 1.25,
  },

  safeDishOpenHint: {
    marginTop: 5,
    color: "#2e7d32",
    fontSize: 12,
    fontWeight: "bold",
  },
};

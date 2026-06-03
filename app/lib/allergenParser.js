import * as XLSX from "xlsx";

export const cleanText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

export const safeText = (value) => String(value || "").replace(/\s+/g, " ").trim();

export const normalizeCode = (value) => {
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

export const extractCodesFromText = (value) =>
  String(value || "")
    .match(/\b\d{3,}\b/g)
    ?.map((item) => normalizeCode(item))
    .filter(Boolean) || [];

export const getVenueIcon = (venue) => {
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

export const getVenueNameColor = (venue) => {
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

export const normalizeVenueName = (value) =>
  safeText(value)
    .replace(/\s*-\s*VV$/i, "")
    .replace(/\s+VV$/i, "")
    .replace(/\s+/g, " ")
    .trim();

export const normalizeIngredientTextForMatch = (value) =>
  cleanText(value)
    .replace(/[^A-Z0-9 ]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const normalizeRecipeLookupName = (value) =>
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

export const getRecipeLookupKeys = (value) =>
  [
    normalizeIngredientTextForMatch(value),
    normalizeRecipeLookupName(value),
  ].filter(Boolean);

const normalizedPhraseCache = new Map();

const getNormalizedPhrase = (word) => {
  const cacheKey = String(word || "");

  if (normalizedPhraseCache.has(cacheKey)) {
    return normalizedPhraseCache.get(cacheKey);
  }

  const normalized = normalizeIngredientTextForMatch(word);
  const phrase = normalized ? ` ${normalized} ` : "";

  normalizedPhraseCache.set(cacheKey, phrase);

  return phrase;
};

export const textHasWordOrPhrase = (text, word) => {
  const source = normalizeIngredientTextForMatch(text);
  const targetPhrase = getNormalizedPhrase(word);

  if (!source || !targetPhrase) return false;

  return ` ${source} `.includes(targetPhrase);
};

export const isGlutenFreeProductClaim = (...values) => {
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

export const ALLERGEN_ORDER = [
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

export const VALID_ALLERGENS = new Set(ALLERGEN_ORDER);

export const sortAllergens = (allergens = []) => {
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

export const isIgnoredBasicIngredient = (value) => {
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

export const isPlainFreshHerbOrRawProduce = (...values) => {
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

export const shouldSkipPossibleAllergensForIngredient = (...values) => {
  const text = values
    .map((value) => normalizeIngredientTextForMatch(value))
    .join(" ");

  if (!text.trim()) return false;

  return (
    isIgnoredBasicIngredient(text) ||
    isPlainFreshHerbOrRawProduce(...values)
  );
};

export const isSubRecipeRow = (item) => {
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

export const isOnboardRecipeName = (...values) => {
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

export const ALLERGEN_DISPLAY = {
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

export const normalizeAllergenName = (value) => {
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

export const splitAllergens = (...values) => {
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
    words: ["shrimp", "prawn", "prawns", "crab", "lobster", "crayfish", "scampi"],
  },
  {
    allergen: "Eggs",
    words: ["egg", "eggs", "mayonnaise", "mayo", "aioli", "meringue", "custard", "hollandaise"],
  },
  {
    allergen: "Fish",
    words: ["fish", "anchovy", "anchovies", "fish sauce", "salmon", "tuna", "cod", "sardine", "mackerel", "trout", "worcestershire"],
  },
  {
    allergen: "Peanuts",
    words: ["peanut", "peanuts", "peanut butter", "peanutbutter", "satay"],
  },
  {
    allergen: "Soybeans",
    words: ["soy", "soya", "soybean", "soybeans", "tofu", "edamame", "miso", "tamari", "soy sauce", "yuba"],
  },
  {
    allergen: "Milk",
    words: ["milk", "cream", "butter", "cheese", "yogurt", "yoghurt", "parmesan", "mozzarella", "ricotta", "mascarpone", "whey", "casein", "lactose", "milk chocolate", "white chocolate"],
  },
  {
    allergen: "Tree Nuts",
    words: ["almond", "hazelnut", "walnut", "cashew", "pecan", "brazil nut", "pistachio", "macadamia", "queensland nut", "pine nut", "marzipan", "praline", "gianduja", "nutella"],
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
    words: ["sesame", "sesame seed", "sesame seeds", "tahini", "benne", "gingelly"],
  },
  {
    allergen: "Sulphur dioxide and sulphites",
    words: ["sulphite", "sulphites", "sulfite", "sulfites", "metabisulfite", "sulfur dioxide", "sulphur dioxide"],
  },
  {
    allergen: "Lupin",
    words: ["lupin", "lupine"],
  },
  {
    allergen: "Molluscs",
    words: ["clam", "clams", "mussel", "mussels", "oyster", "oysters", "scallop", "scallops", "squid", "octopus", "calamari", "snail", "escargot"],
  },
];

export const isProcessedOrPreparedItem = (value) => {
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
    words: ["cookie", "cookies", "biscuit", "cake", "muffin", "brownie", "donut", "doughnut"],
    allergens: ["Cereals containing gluten", "Milk", "Eggs"],
  },
  {
    words: ["bread", "bun", "brioche", "croissant", "pastry", "tart shell", "pie shell", "cracker"],
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

export const getPreparedCommonAllergens = (...values) => {
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

export const keywordAllergensForText = (...values) => {
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

export const looksGlutenFreeClaim = (...values) => {
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

export const parseIngredientByLocationWorkbook = (workbook) => {
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
    .map((venue) => {
      const recipes = [...venue.recipesMap.values()]
        .map((recipe) => ({
          recipeKey: recipe.recipeKey,
          recipeCode: recipe.recipeCode,
          recipeName: recipe.recipeName,
          menuCode: recipe.menuCode,
          menuName: recipe.menuName,
          restaurantName: recipe.restaurantName,
          rows: recipe.rows,
          allergens: sortAllergens([...recipe.allergens]),
          possibleHidden: sortAllergens([...recipe.possibleHidden]),
          hiddenWarnings: recipe.hiddenWarnings,
          gfClaim: recipe.gfClaim,
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
        );

      return {
        venueKey: venue.venueKey,
        restaurantName: venue.restaurantName,
        restaurantCode: venue.restaurantCode,
        icon: venue.icon,
        allergens: sortAllergens([...venue.allergens]),
        possibleHidden: sortAllergens([...venue.possibleHidden]),
        hiddenWarningCount: venue.hiddenWarningCount,
        ingredientCount: venue.ingredientCount,
        recipes,
      };
    })
    .sort((a, b) => a.restaurantName.localeCompare(b.restaurantName));

  return {
    rows: parsedRows,
    venues,
    sourceSheet: sheetName,
    parsedAt: new Date().toISOString(),
    parserVersion: 2,
  };
};

export const parseIngredientByLocationArrayBuffer = (arrayBuffer) => {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    dense: true,
  });

  return parseIngredientByLocationWorkbook(workbook);
};

export const buildSubRecipeIndexes = (rows = []) => {
  const byVenueRecipeName = new Map();
  const byVenueRecipeCode = new Map();
  const byAnyRecipeName = new Map();
  const byAnyRecipeCode = new Map();
  const byVenueMenuName = new Map();

  const add = (map, key, row) => {
    if (!key) return;

    if (!map.has(key)) {
      map.set(key, []);
    }

    map.get(key).push(row);
  };

  rows.forEach((row) => {
    const venueKey = cleanText(row.restaurantName);
    const recipeCode = normalizeCode(row.recipeCode);

    getRecipeLookupKeys(row.recipeName).forEach((nameKey) => {
      add(byVenueRecipeName, `${venueKey}|${nameKey}`, row);
      add(byAnyRecipeName, nameKey, row);
    });

    if (recipeCode) {
      add(byVenueRecipeCode, `${venueKey}|${recipeCode}`, row);
      add(byAnyRecipeCode, recipeCode, row);
    }

    getRecipeLookupKeys(row.menuName).forEach((menuKey) => {
      add(byVenueMenuName, `${venueKey}|${menuKey}`, row);
    });
  });

  const sortMap = (map) => {
    map.forEach((list, key) => {
      map.set(
        key,
        list.sort(
          (a, b) => Number(a.sourceRow || 0) - Number(b.sourceRow || 0)
        )
      );
    });
  };

  [
    byVenueRecipeName,
    byVenueRecipeCode,
    byAnyRecipeName,
    byAnyRecipeCode,
    byVenueMenuName,
  ].forEach(sortMap);

  return {
    byVenueRecipeName,
    byVenueRecipeCode,
    byAnyRecipeName,
    byAnyRecipeCode,
    byVenueMenuName,
  };
};

export const getSubRecipeRowsForLineFromIndexes = ({
  subRecipeLine,
  indexes,
}) => {
  if (!subRecipeLine || !indexes) return [];

  const currentVenueKey = cleanText(subRecipeLine.restaurantName);

  const sourceParentRecipeKey = cleanText(
    `${subRecipeLine.recipeCode}|${subRecipeLine.recipeName}|${subRecipeLine.menuCode}|${subRecipeLine.menuName}`
  );

  const candidateNameKeys = [
    subRecipeLine.ingredientName,
    subRecipeLine.assigned,
  ]
    .flatMap((value) => getRecipeLookupKeys(value))
    .filter(Boolean);

  const candidateCodeKeys = [
    subRecipeLine.ingredientCode,
    subRecipeLine.assigned,
    ...extractCodesFromText(subRecipeLine.ingredientName),
    ...extractCodesFromText(subRecipeLine.assigned),
  ]
    .map((value) => normalizeCode(value))
    .filter(Boolean);

  const isAllowed = (row) => {
    const rowRecipeKey = cleanText(
      `${row.recipeCode}|${row.recipeName}|${row.menuCode}|${row.menuName}`
    );

    return (
      row.sourceRow !== subRecipeLine.sourceRow &&
      rowRecipeKey !== sourceParentRecipeKey
    );
  };

  const collectRowsForKeys = (map, keys) => {
    const seen = new Set();
    const result = [];

    keys.forEach((key) => {
      const list = map.get(key) || [];

      list.forEach((row) => {
        const rowKey = `${row.sourceRow}-${row.recipeCode}-${row.ingredientCode}-${row.ingredientName}`;

        if (!seen.has(rowKey) && isAllowed(row)) {
          seen.add(rowKey);
          result.push(row);
        }
      });
    });

    return result.sort(
      (a, b) => Number(a.sourceRow || 0) - Number(b.sourceRow || 0)
    );
  };

  const sameVenueRecipeNameRows = collectRowsForKeys(
    indexes.byVenueRecipeName,
    candidateNameKeys.map((nameKey) => `${currentVenueKey}|${nameKey}`)
  );

  if (sameVenueRecipeNameRows.length) return sameVenueRecipeNameRows;

  const sameVenueRecipeCodeRows = collectRowsForKeys(
    indexes.byVenueRecipeCode,
    candidateCodeKeys.map((codeKey) => `${currentVenueKey}|${codeKey}`)
  );

  if (sameVenueRecipeCodeRows.length) return sameVenueRecipeCodeRows;

  const anyVenueRecipeNameRows = collectRowsForKeys(
    indexes.byAnyRecipeName,
    candidateNameKeys
  );

  if (anyVenueRecipeNameRows.length) return anyVenueRecipeNameRows;

  const anyVenueRecipeCodeRows = collectRowsForKeys(
    indexes.byAnyRecipeCode,
    candidateCodeKeys
  );

  if (anyVenueRecipeCodeRows.length) return anyVenueRecipeCodeRows;

  const sameVenueMenuNameRows = collectRowsForKeys(
    indexes.byVenueMenuName,
    candidateNameKeys.map((nameKey) => `${currentVenueKey}|${nameKey}`)
  );

  if (sameVenueMenuNameRows.length) return sameVenueMenuNameRows;

  return [];
};

export const SHIPS = ["BRL", "RL", "SC", "VL"];

export const EQUIPMENT_DEPARTMENTS = [
  "Culinary",
  "Bar",
  "Restaurant",
];

export const MAKE_INVENTORY_STATIONS = [
  "VEG PREP",
  "BUTCHER PREP",
  "FISH PREP",
  "BAKERY",
  "Pink Agave",
  "Pastry deck 5",
  "Razzle Dazzle",
  "Kitchen Table",
  "Test Kitchen",
  "Pastry deck 6",
  "The Wake",
  "Garde Manger",
  "Extra Virgin",
  "Manor",
  "The Dock House",
  "Social Club",
  "Pizza Place",
  "The Galley",
  "Gunbae",
  "Sun Club",
  "Locker deck 6",
  "POT WASH DECK 4",
  "POT WASH DECK 5",
  "POT WASH DECK 6",
  "POT WASH DECK 15",
];

export const ALLERGEN_RULES = [
  {
    allergen: "Tree Nuts",
    keywords: ["almond", "walnut", "pecan", "cashew", "hazelnut", "pistachio", "macadamia"],
  },
  {
    allergen: "Peanuts",
    keywords: ["peanut"],
  },
  {
    allergen: "Seeds",
    keywords: ["seed", "seeds", "sunflower seed", "pumpkin seed", "chia", "flax", "hemp seed"],
    exclude: ["seedless", "seedless cucumber"],
  },
  {
    allergen: "Soy",
    keywords: ["soy", "tofu", "edamame", "miso", "tamari"],
  },
  {
    allergen: "Gluten",
    keywords: ["wheat", "flour", "gluten", "bread", "pasta", "semolina", "barley", "rye", "panko"],
  },
  {
    allergen: "Milk / Dairy",
    keywords: ["milk", "cream", "butter", "cheese", "yogurt", "parmesan", "mozzarella", "ricotta", "cream cheese"],
  },
  {
    allergen: "Egg",
    keywords: ["egg", "eggs", "mayonnaise", "aioli"],
    exclude: ["eggplant"],
  },
  {
    allergen: "Fish",
    keywords: ["salmon", "tuna", "cod", "anchovy", "fish", "sardine"],
  },
  {
    allergen: "Shellfish",
    keywords: ["shrimp", "crab", "lobster", "mussel", "oyster", "scallop"],
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

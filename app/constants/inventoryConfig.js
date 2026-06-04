export const EQUIPMENT_PICTURE_BUCKET = "equipment-pictures";

const cleanInventoryText = (value) =>
  String(value || "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();

export const getMasterInventoryScope = (department) => {
  const dept = cleanInventoryText(department || "culinary").replace(
    /[^A-Z0-9]/g,
    "_"
  );

  return "GLOBAL_" + (dept || "CULINARY");
};

export const CULINARY_STATIONS = [
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

export const BAR_STATIONS = [
  "Crew Bar",
  "Crew Shop",
  "Office",
  "IV BAR",
  "KT",
  "Bosun Club",
  "Bosun Club Locker",
  "Pink Agave",
  "Razzle Dazzle",
  "Red Room D6",
  "The Manor",
  "Casino",
  "On the Rocks",
  "On The Rocks Locker",
  "Extra Virgin",
  "Test Kitchen",
  "The Wake",
  "Red Room D7",
  "Sip",
  "Manor Storage Locker",
  "Draught Haus",
  "Grounds Club",
  "Loose Cannon",
  "Social Club",
  "The Dock & Dockhouse",
  "Aquatic / Gym & Tonic",
  "Grounds Club Too",
  "Gunbae",
  "Richard's Rooftop",
  "Sun Club",
  "D16 Storage Locker",
  "Athletic Club Locker",
  "Athletic Club",
  "Crew Lookout",
  "D17 Storage Locker",
];

export const RESTAURANT_STATIONS = [
  "RD",
  "PA",
  "TK",
  "WAKE",
  "EV",
  "GUNBAE",
  "GALLEY",
  "SE",
  "PP",
  "KT",
  "MANOR",
  "DOCK",
  "LTIC",
  "LOCKER 2238Z",
  "LOCKER 4348Z",
  "LOCKER 4350Z",
  "LOCKER 4354Z",
  "LOCKER 5072A",
  "LOCKER 5082Z",
  "LOCKER 5106A",
  "LOCKER 6026P",
  "LOCKER 6066P",
  "LOCKER 7012Z",
  "LOCKER 15276P",
  "ROCKSTAR SUITE",
];

export const getInventoryStationsForDepartment = (department) =>
  department === "bar" ? BAR_STATIONS : CULINARY_STATIONS;

export const getInventoryStationLabelForDepartment = (department) =>
  department === "bar" ? "bar" : "station";

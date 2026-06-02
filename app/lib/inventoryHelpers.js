const cleanInventoryText = (value) =>
  String(value || "").toUpperCase().replace(/\s+/g, " ").trim();

export const getEquipmentDepartmentKey = (department) =>
  cleanInventoryText(department || "culinary").replace(/[^A-Z0-9]/g, "_") ||
  "CULINARY";

export const inventoryRecordMatchesDepartment = (item, department) => {
  const departmentKey = getEquipmentDepartmentKey(department);
  const key = cleanInventoryText(item?.itemKey || "");

  if (!key) return true;
  if (key.startsWith(departmentKey + "|")) return true;

  const hasKnownDepartmentPrefix = ["CULINARY|", "BAR|", "RESTAURANT|"].some(
    (prefix) => key.startsWith(prefix)
  );

  return !hasKnownDepartmentPrefix && departmentKey === "CULINARY";
};

export const normalizeInventoryRecord = (record) => ({
  id: record.id,
  ship: record.ship,
  station: record.station,
  userName: record.user_name,
  itemKey: record.item_key,
  code: record.code || "",
  name: record.item_name || "",
  category: record.category || "",
  sheetName: record.sheet_name || "",
  image: record.image || "",
  qty: Number(record.qty || 0),
  confirmedAt: record.updated_at
    ? new Date(record.updated_at).toLocaleString()
    : "",
  updatedAt: record.updated_at || "",
});

export const normalizeMasterInventoryRecord = (record) => ({
  id: record.id,
  ship: record.ship,
  itemKey: record.item_key,
  code: record.code || "",
  name: record.item_name || "",
  category: record.category || "",
  sheetName: record.sheet_name || "",
  image: record.image || "",
  sourceRow: Number(record.source_row || 0),
  sortOrder: Number(record.sort_order || 0),
  updatedAt: record.updated_at || "",
});

export const normalizeInventoryStationStatusRecord = (record) => ({
  id: record.id,
  ship: record.ship || "",
  department: record.department || "",
  station: record.station || "",
  status: record.status || "not_started",
  userName: record.user_name || "",
  userPosition: record.user_position || "",
  startedAt: record.started_at || "",
  submittedAt: record.submitted_at || "",
  updatedAt: record.updated_at || "",
});

export const getInventoryProductGroupKey = (item) => {
  const codeKey = cleanInventoryText(item?.code || "")
    .replace(/\s+/g, "")
    .replace(/\.0$/, "");

  const nameKey = cleanInventoryText(item?.name || "");

  if (codeKey && nameKey) return `${codeKey}__${nameKey}`;
  if (codeKey) return `CODE__${codeKey}`;
  if (nameKey) return `NAME__${nameKey}`;

  return "";
};

export const buildInventoryQtyMap = (rows = []) => {
  const map = new Map();

  rows.forEach((item) => {
    const key = getInventoryProductGroupKey(item);
    if (!key) return;

    const rawQty =
      item.qty ??
      item.count ??
      item.Count ??
      item.Quantity ??
      item.quantity ??
      item.totalQty ??
      0;

    const qty = Number(rawQty || 0);
    const safeQty = Number.isFinite(qty) ? qty : 0;

    map.set(key, Number(map.get(key) || 0) + safeQty);
  });

  return map;
};

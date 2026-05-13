
export function normalizeCode(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/\.0$/, "");
}

export function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ");
}

export function normalizeKeyText(value) {
  return normalizeText(value).toLowerCase();
}

export function safeFileName(value) {
  return String(value || "inventory")
    .trim()
    .replace(/[\\/:*?"<>|]/g, "-")
    .replace(/\s+/g, "-")
    .toLowerCase();
}

export function getItemCode(item) {
  return (
    item?.code ??
    item?.Code ??
    item?.itemCode ??
    item?.item_code ??
    item?.productCode ??
    item?.product_code ??
    item?.sku ??
    item?.SKU ??
    ""
  );
}

export function getItemName(item) {
  return (
    item?.name ??
    item?.Name ??
    item?.itemName ??
    item?.item_name ??
    item?.productName ??
    item?.product_name ??
    item?.description ??
    item?.Description ??
    item?.itemDescription ??
    item?.item_description ??
    ""
  );
}

export function getItemCount(item) {
  return (
    item?.count ??
    item?.Count ??
    item?.countedQty ??
    item?.countedQuantity ??
    item?.counted_quantity ??
    item?.inventoryCount ??
    item?.inventory_count ??
    item?.reportCount ??
    item?.report_count ??
    item?.qty ??
    item?.Qty ??
    item?.quantity ??
    item?.Quantity ??
    ""
  );
}

export function getItemVenue(item) {
  return (
    item?.venue ??
    item?.Venue ??
    item?.venueName ??
    item?.venue_name ??
    item?.location ??
    item?.Location ??
    ""
  );
}

export function toCountNumber(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  const cleaned = String(value).replace(/,/g, "").trim();

  if (cleaned === "") {
    return "";
  }

  const numberValue = Number(cleaned);

  return Number.isFinite(numberValue) ? numberValue : "";
}

export function makeReportItem(item) {
  const code = normalizeCode(getItemCode(item));
  const name = normalizeText(getItemName(item));
  const count = toCountNumber(getItemCount(item));

  return {
    ...item,
    code,
    name,
    count,
    Count: count,
    quantity: count,
    qty: count,
  };
}

export function prepareInventoryReportItems(items = []) {
  return items
    .map(makeReportItem)
    .filter((item) => item.code || item.name);
}

export function buildSummaryReportItems(items = []) {
  const grouped = new Map();

  prepareInventoryReportItems(items).forEach((item) => {
    const code = normalizeCode(item.code);
    const name = normalizeText(item.name);

    if (!code && !name) return;

    const key = `${code}__${normalizeKeyText(name)}`;
    const count = toCountNumber(item.count);

    if (!grouped.has(key)) {
      grouped.set(key, {
        ...item,
        code,
        name,
        venue: "All Venues",
        Venue: "All Venues",
        count: 0,
        Count: 0,
        quantity: 0,
        qty: 0,
      });
    }

    const existing = grouped.get(key);
    const newTotal = Number(existing.count || 0) + Number(count || 0);

    existing.count = newTotal;
    existing.Count = newTotal;
    existing.quantity = newTotal;
    existing.qty = newTotal;
  });

  return Array.from(grouped.values()).sort((a, b) => {
    const codeA = normalizeCode(a.code);
    const codeB = normalizeCode(b.code);

    if (codeA !== codeB) {
      return codeA.localeCompare(codeB, undefined, { numeric: true });
    }

    return normalizeKeyText(a.name).localeCompare(normalizeKeyText(b.name));
  });
}

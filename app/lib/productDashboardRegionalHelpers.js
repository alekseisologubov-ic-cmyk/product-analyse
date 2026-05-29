import {
  YEARLY_REGION_ALL,
  getRegionalParSuggestion,
  formatRegionalQty,
} from "./yearlyRegionalConsumption";

const toNumber = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const cleaned = String(value ?? "")
    .replace(/,/g, "")
    .replace(/[^0-9.\-]/g, "")
    .trim();

  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
};

const firstValue = (...values) =>
  values.find((value) => value !== undefined && value !== null && value !== "");

export const getDashboardProductCode = (product) =>
  firstValue(
    product?.code,
    product?.itemCode,
    product?.productCode,
    product?.sku,
    product?.erpCode,
    product?.inventoryCode,
    product?.articleCode,
    product?.Code,
    product?.ITEM_CODE,
    product?.PRODUCT_CODE,
    ""
  );

export const getDashboardProductName = (product) =>
  firstValue(
    product?.product,
    product?.productName,
    product?.name,
    product?.itemName,
    product?.description,
    product?.Product,
    product?.PRODUCT,
    product?.PRODUCT_NAME,
    product?.ITEM_NAME,
    ""
  );

export const getDashboardProductUnit = (product) =>
  firstValue(
    product?.unit,
    product?.um,
    product?.uom,
    product?.unitMeasure,
    product?.unitOfMeasure,
    product?.UM,
    product?.UOM,
    product?.UNIT,
    ""
  );

export const getDashboardProductParLevel = (product) =>
  toNumber(
    firstValue(
      product?.parLevel,
      product?.currentParLevel,
      product?.par,
      product?.parQty,
      product?.currentPar,
      product?.currentParQty,
      product?.qPar,
      product?.q,
      product?.PAR,
      product?.PAR_LEVEL,
      product?.CURRENT_PAR,
      product?.CURRENT_PAR_LEVEL,
      product?.par_level,
      product?.current_par,
      product?.current_par_level,
      0
    )
  );

export const enrichDashboardProductsWithRegionalPar = ({
  products,
  yearlyRegionalConsumption,
  selectedRegion = YEARLY_REGION_ALL,
  voyageDays,
  bufferPercent = 0,
}) => {
  return (products || []).map((product) => {
    const productCode = getDashboardProductCode(product);
    const productName = getDashboardProductName(product);
    const unit = getDashboardProductUnit(product);
    const currentParLevel = getDashboardProductParLevel(product);

    const regionalPar = getRegionalParSuggestion({
      yearlyRegionalConsumption,
      productCode,
      productName,
      region: selectedRegion,
      voyageDays,
      bufferPercent,
    });

    const regionalSuggestedParLevel = Number(
      regionalPar.suggestedParLevel || 0
    );

    const regionalSuggestedParDifference =
      regionalSuggestedParLevel - currentParLevel;

    return {
      ...product,

      code: product?.code ?? productCode,
      product: product?.product ?? productName,
      unit: product?.unit ?? unit,

      currentParLevel,

      regionalHasData: Boolean(regionalPar.hasRegionalData),
      regionalRegion: selectedRegion || YEARLY_REGION_ALL,

      regionalTotalQty: Number(regionalPar.totalQty || 0),
      regionalTotalDays: Number(regionalPar.totalDays || 0),
      regionalAvgDailyQty: Number(regionalPar.avgDailyQty || 0),

      regionalSuggestedParLevel,
      regionalSuggestedParDifference,

      regionalEvidenceBlocks: Number(regionalPar.evidenceBlocks || 0),
      regionalMatchedProductCode: regionalPar.matchedProductCode || "",
      regionalMatchedProductName: regionalPar.matchedProductName || "",
      regionalByShip: regionalPar.byShip || [],

      suggestionBasis: regionalPar.hasRegionalData
        ? "Yearly regional consumption"
        : "No yearly regional match",
    };
  });
};

export const buildDashboardRegionalExportRows = (products = []) =>
  products.map((row, index) => ({
    Number: index + 1,
    Code: getDashboardProductCode(row),
    Product: getDashboardProductName(row),
    UM: getDashboardProductUnit(row),

    Region:
      row.regionalRegion === YEARLY_REGION_ALL
        ? "All regions"
        : row.regionalRegion || "",

    SuggestionBasis: row.suggestionBasis || "",

    CurrentParLevel: row.currentParLevel,
    RegionalDailyConsumption: row.regionalAvgDailyQty,
    RegionalTotalQty: row.regionalTotalQty,
    RegionalTotalDays: row.regionalTotalDays,
    RegionalEvidenceBlocks: row.regionalEvidenceBlocks,

    SuggestedRegionalParLevel: row.regionalSuggestedParLevel,
    SuggestedParDifference: row.regionalSuggestedParDifference,

    MatchedRegionalProductCode: row.regionalMatchedProductCode,
    MatchedRegionalProductName: row.regionalMatchedProductName,
  }));

export const getRegionalParCardLines = (row) => {
  if (!row?.regionalHasData) {
    return [
      "Regional yearly match: No",
      "Suggested par: not available",
    ];
  }

  const region =
    row.regionalRegion === YEARLY_REGION_ALL
      ? "All regions"
      : row.regionalRegion || "N/A";

  return [
    `Region: ${region}`,
    `Regional daily consumption: ${formatRegionalQty(row.regionalAvgDailyQty)}`,
    `Current par: ${formatRegionalQty(row.currentParLevel)}`,
    `Suggested par: ${formatRegionalQty(row.regionalSuggestedParLevel)}`,
    `Par difference: ${formatRegionalQty(row.regionalSuggestedParDifference)}`,
    `Evidence: ${row.regionalEvidenceBlocks || 0} block(s), ${formatRegionalQty(
      row.regionalTotalDays
    )} days`,
  ];
};

export const getRegionalParSummaryText = (row) => {
  if (!row?.regionalHasData) {
    return "No yearly regional match";
  }

  return `Regional daily ${formatRegionalQty(
    row.regionalAvgDailyQty
  )}, suggested par ${formatRegionalQty(row.regionalSuggestedParLevel)}`;
};

export const getRegionalParStatus = (row) => {
  if (!row?.regionalHasData) return "missing";

  const difference = Number(row.regionalSuggestedParDifference || 0);

  if (difference > 0) return "increase";
  if (difference < 0) return "decrease";

  return "same";
};

export const sortProductsByRegionalParDifference = (products = []) =>
  [...products].sort((a, b) => {
    const aHasData = a?.regionalHasData ? 1 : 0;
    const bHasData = b?.regionalHasData ? 1 : 0;

    if (aHasData !== bHasData) return bHasData - aHasData;

    const diff =
      Math.abs(Number(b?.regionalSuggestedParDifference || 0)) -
      Math.abs(Number(a?.regionalSuggestedParDifference || 0));

    if (diff !== 0) return diff;

    return String(getDashboardProductName(a)).localeCompare(
      String(getDashboardProductName(b))
    );
  });

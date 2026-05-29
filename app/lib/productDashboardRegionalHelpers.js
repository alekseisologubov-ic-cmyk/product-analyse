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
    ""
  );

export const getDashboardProductName = (product) =>
  firstValue(
    product?.product,
    product?.productName,
    product?.name,
    product?.itemName,
    product?.description,
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

    return {
      ...product,

      regionalHasData: Boolean(regionalPar.hasRegionalData),
      regionalRegion: selectedRegion || YEARLY_REGION_ALL,

      regionalTotalQty: Number(regionalPar.totalQty || 0),
      regionalTotalDays: Number(regionalPar.totalDays || 0),
      regionalAvgDailyQty: Number(regionalPar.avgDailyQty || 0),

      regionalSuggestedParLevel,
      regionalSuggestedParDifference:
        regionalSuggestedParLevel - currentParLevel,

      regionalEvidenceBlocks: Number(regionalPar.evidenceBlocks || 0),
      regionalMatchedProductCode: regionalPar.matchedProductCode || "",
      regionalMatchedProductName: regionalPar.matchedProductName || "",
      regionalByShip: regionalPar.byShip || [],

      currentParLevel,

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

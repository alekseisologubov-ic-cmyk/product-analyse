"use client";

import { buildInventoryWorkbook } from "./inventoryWorkbookBuilder";
import {
  buildSummaryReportItems,
  safeFileName,
} from "./inventoryReportData";

async function getSaveAs() {
  const fileSaver = await import("file-saver");
  return fileSaver.saveAs || fileSaver.default?.saveAs || fileSaver.default;
}

async function saveWorkbookAsExcel(workbook, fileName) {
  const saveAs = await getSaveAs();

  const buffer = await workbook.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  saveAs(blob, fileName);
}

async function saveBlob(blob, fileName) {
  const saveAs = await getSaveAs();
  saveAs(blob, fileName);
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

export async function downloadInventoryCountSheet({
  items = [],
  venueName = "",
} = {}) {
  const workbook = buildInventoryWorkbook({
    items,
    venueName,
    reportTitle: "Inventory Count Sheet",
    includeCounts: false,
  });

  const fileName = `count-sheet-${safeFileName(venueName)}-${today()}.xlsx`;

  await saveWorkbookAsExcel(workbook, fileName);
}

export async function downloadInventoryExcelReport({
  items = [],
  venueName = "",
} = {}) {
  const workbook = buildInventoryWorkbook({
    items,
    venueName,
    reportTitle: "Inventory Report",
    includeCounts: true,
  });

  const fileName = `inventory-report-${safeFileName(venueName)}-${today()}.xlsx`;

  await saveWorkbookAsExcel(workbook, fileName);
}

export async function downloadSummaryExcelReport({
  items = [],
} = {}) {
  const summaryItems = buildSummaryReportItems(items);

  const workbook = buildInventoryWorkbook({
    items: summaryItems,
    venueName: "All Venues",
    reportTitle: "Summary Report",
    includeCounts: true,
  });

  const fileName = `summary-report-all-venues-${today()}.xlsx`;

  await saveWorkbookAsExcel(workbook, fileName);
}

export async function downloadInventoryPdfReport({
  items = [],
  venueName = "",
  reportTitle = "Inventory Report",
  includeCounts = true,
  summary = false,
} = {}) {
  const response = await fetch("/api/inventory-report/pdf", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      items,
      venueName,
      reportTitle,
      includeCounts,
      summary,
    }),
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(message || "Failed to create PDF report.");
  }

  const blob = await response.blob();

  const fileName =
    summary === true
      ? `summary-report-all-venues-${today()}.pdf`
      : `${safeFileName(reportTitle)}-${safeFileName(venueName)}-${today()}.pdf`;

  await saveBlob(blob, fileName);
}

"use client";

import ExcelJS from "exceljs";
import {
  buildSummaryReportItems,
  getItemCode,
  getItemName,
  getItemCount,
  normalizeCode,
  normalizeText,
  prepareInventoryReportItems,
  safeFileName,
  toCountNumber,
} from "./inventoryReportData";

async function getSaveAs() {
  const fileSaver = await import("file-saver");
  return fileSaver.saveAs || fileSaver.default?.saveAs || fileSaver.default;
}

async function getJsPDF() {
  const mod = await import("jspdf");
  return mod.jsPDF || mod.default;
}

async function saveWorkbookAsExcel(workbook, fileName) {
  const saveAs = await getSaveAs();

  const buffer = await workbook.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  saveAs(blob, fileName);
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

/* =========================================================
   PDF EXPORT
   This still creates the app-generated PDF layout.
   Excel export below uses the uploaded sample file.
========================================================= */

function getPdfColumns(pageWidth) {
  const margin = 18;
  const contentWidth = pageWidth - margin * 2;

  const widths = [];

  for (let index = 1; index <= 20; index += 1) {
    if (index === 1) {
      widths.push(58); // A code
    } else if (index >= 2 && index <= 5) {
      widths.push(14); // B-E blank
    } else if (index === 6) {
      widths.push(400); // F item name
    } else if (index >= 7 && index <= 18) {
      widths.push(14); // G-R blank
    } else if (index === 19) {
      widths.push(60); // S count
    } else {
      widths.push(14); // T blank
    }
  }

  const total = widths.reduce((sum, value) => sum + value, 0);
  const scale = contentWidth / total;

  return widths.map((value) => value * scale);
}

function drawGridRow(doc, x, y, rowHeight, columns) {
  let currentX = x;

  columns.forEach((width) => {
    doc.rect(currentX, y, width, rowHeight);
    currentX += width;
  });
}

function getColumnX(startX, columns, columnNumber) {
  let x = startX;

  for (let index = 1; index < columnNumber; index += 1) {
    x += columns[index - 1];
  }

  return x;
}

function drawPdfHeader({
  doc,
  pageWidth,
  margin,
  columns,
  reportTitle,
  venueName,
}) {
  const generatedDate = new Date().toLocaleDateString();

  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text(reportTitle || "Inventory Report", pageWidth / 2, 30, {
    align: "center",
  });

  doc.setFontSize(10);
  doc.text(
    venueName ? `Venue: ${venueName}` : "Venue: All Venues",
    pageWidth / 2,
    48,
    {
      align: "center",
    }
  );

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(`Generated: ${generatedDate}`, pageWidth / 2, 64, {
    align: "center",
  });

  doc.text(
    "Code is in column A. Item name is in column F. Count is in column S.",
    pageWidth / 2,
    80,
    {
      align: "center",
    }
  );

  const headerY = 96;
  const headerHeight = 22;

  doc.setFillColor(237, 237, 237);
  doc.rect(margin, headerY, pageWidth - margin * 2, headerHeight, "F");

  drawGridRow(doc, margin, headerY, headerHeight, columns);

  const codeX = getColumnX(margin, columns, 1);
  const nameX = getColumnX(margin, columns, 6);
  const countX = getColumnX(margin, columns, 19);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.5);

  doc.text("CODE", codeX + 3, headerY + 14);
  doc.text("ITEM NAME", nameX + 3, headerY + 14);
  doc.text("COUNT", countX + columns[18] / 2, headerY + 14, {
    align: "center",
  });

  return headerY + headerHeight;
}

function drawInventoryPdf({
  doc,
  items = [],
  venueName = "",
  reportTitle = "Inventory Report",
  includeCounts = true,
}) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 18;
  const columns = getPdfColumns(pageWidth);

  const rowHeight = 28;
  const bottomMargin = 24;

  let y = drawPdfHeader({
    doc,
    pageWidth,
    margin,
    columns,
    reportTitle,
    venueName,
  });

  const reportItems = prepareInventoryReportItems(items);

  reportItems.forEach((item) => {
    if (y + rowHeight > pageHeight - bottomMargin) {
      doc.addPage();

      y = drawPdfHeader({
        doc,
        pageWidth,
        margin,
        columns,
        reportTitle,
        venueName,
      });
    }

    drawGridRow(doc, margin, y, rowHeight, columns);

    const code = normalizeCode(getItemCode(item));
    const name = normalizeText(getItemName(item));
    const count = toCountNumber(getItemCount(item));

    const codeX = getColumnX(margin, columns, 1);
    const nameX = getColumnX(margin, columns, 6);
    const countX = getColumnX(margin, columns, 19);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(6.8);

    doc.text(String(code || ""), codeX + 3, y + 11);

    const wrappedName = doc.splitTextToSize(String(name || ""), columns[5] - 6);
    doc.text(wrappedName.slice(0, 2), nameX + 3, y + 10);

    if (includeCounts) {
      const countText =
        count === "" || count === null || count === undefined
          ? ""
          : String(count);

      doc.text(countText, countX + columns[18] / 2, y + 17, {
        align: "center",
      });
    }

    y += rowHeight;
  });

  const pageCount = doc.getNumberOfPages();

  for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
    doc.setPage(pageNumber);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7);
    doc.text(
      `Page ${pageNumber} of ${pageCount}`,
      pageWidth / 2,
      pageHeight - 10,
      {
        align: "center",
      }
    );
  }
}

export async function downloadInventoryPdfReport({
  items = [],
  venueName = "",
  reportTitle = "Inventory Report",
  includeCounts = true,
  summary = false,
} = {}) {
  const JsPDF = await getJsPDF();

  const doc = new JsPDF({
    orientation: "landscape",
    unit: "pt",
    format: "letter",
  });

  const finalItems = summary ? buildSummaryReportItems(items) : items;
  const finalVenueName = summary ? venueName || "All Venues" : venueName;
  const finalReportTitle = summary ? "Summary Report" : reportTitle;

  drawInventoryPdf({
    doc,
    items: finalItems,
    venueName: finalVenueName,
    reportTitle: finalReportTitle,
    includeCounts,
  });

  const fileName =
    summary === true
      ? `summary-report-all-venues-${today()}.pdf`
      : `${safeFileName(finalReportTitle)}-${safeFileName(finalVenueName)}-${today()}.pdf`;

  doc.save(fileName);
}

/* =========================================================
   EXCEL EXPORT USING UPLOADED SAMPLE FILE
   This keeps the uploaded file row positions exactly.
   Code = column A
   Name = column F
   Count = column S
========================================================= */

const TEMPLATE_CODE_COL = 1; // A
const TEMPLATE_NAME_COL = 6; // F
const TEMPLATE_COUNT_COL = 19; // S

function getExcelCellText(value) {
  if (value === null || value === undefined) return "";

  if (typeof value === "object") {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || "").join("");
    }

    if (value.text !== undefined) return String(value.text);
    if (value.result !== undefined) return String(value.result);
  }

  return String(value);
}

function makeTemplateReportKey(codeValue, nameValue) {
  const code = normalizeCode(getExcelCellText(codeValue));
  const name = normalizeText(getExcelCellText(nameValue)).toUpperCase();

  if (!code || !name) return "";

  return `${code}__${name}`;
}

function isTemplateItemRow(codeValue, nameValue) {
  const code = normalizeCode(getExcelCellText(codeValue));
  const name = normalizeText(getExcelCellText(nameValue)).toUpperCase();

  if (!code || !name) return false;

  if (code === "CODE") return false;
  if (name === "ITEM NAME") return false;
  if (name === "NAME") return false;
  if (name.includes("ITEM NAME")) return false;

  return true;
}

function buildTemplateCountMap(items = []) {
  const map = new Map();

  prepareInventoryReportItems(items).forEach((item) => {
    const key = makeTemplateReportKey(getItemCode(item), getItemName(item));
    if (!key) return;

    const rawCount = toCountNumber(getItemCount(item));
    const count = rawCount === "" ? 0 : Number(rawCount || 0);
    const safeCount = Number.isFinite(count) ? count : 0;

    map.set(key, Number(map.get(key) || 0) + safeCount);
  });

  return map;
}

export async function downloadInventoryExcelReportUsingTemplate({
  templateFile,
  items = [],
  venueName = "",
  reportTitle = "Inventory Report",
  stationSummaryRows = [],
} = {}) {
  if (!templateFile) {
    throw new Error("Upload the inventory sheet sample first.");
  }

  const fileNameLower = String(templateFile.name || "").toLowerCase();

  if (fileNameLower.endsWith(".xls") && !fileNameLower.endsWith(".xlsx")) {
    throw new Error("Please upload an .xlsx inventory sheet sample. Old .xls files are not supported.");
  }

  const workbook = new ExcelJS.Workbook();
  const arrayBuffer = await templateFile.arrayBuffer();

  await workbook.xlsx.load(arrayBuffer);

  const countMap = buildTemplateCountMap(items);

  let itemRows = 0;
  let matchedRows = 0;

  workbook.eachSheet((worksheet) => {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const codeCell = row.getCell(TEMPLATE_CODE_COL);
      const nameCell = row.getCell(TEMPLATE_NAME_COL);

      const codeText = getExcelCellText(codeCell.value);
      const nameText = getExcelCellText(nameCell.value);

      if (!isTemplateItemRow(codeText, nameText)) return;

      itemRows += 1;

      const key = makeTemplateReportKey(codeText, nameText);
      const hasCount = key && countMap.has(key);
      const count = hasCount ? countMap.get(key) : 0;

      if (hasCount) matchedRows += 1;

      const countCell = row.getCell(TEMPLATE_COUNT_COL);

      countCell.value = count;
      countCell.numFmt = "0";
      countCell.alignment = {
        ...(countCell.alignment || {}),
        horizontal: "center",
        vertical: "middle",
      };
    });
  });

  const outputFileName = `${safeFileName(reportTitle)}-${safeFileName(
    venueName
  )}-from-uploaded-sample-${today()}.xlsx`;

  await saveWorkbookAsExcel(workbook, outputFileName);

  return {
    itemRows,
    matchedRows,
  };
}

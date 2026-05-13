import ExcelJS from "exceljs";
import {
  getItemCode,
  getItemName,
  getItemCount,
  normalizeCode,
  normalizeText,
  prepareInventoryReportItems,
  toCountNumber,
} from "./inventoryReportData";

export const CODE_COL = 1; // A
export const NAME_COL = 6; // F
export const COUNT_COL = 19; // S
export const LAST_COL = 20; // T

export const HEADER_ROW = 15;
export const FIRST_ITEM_ROW = 16;

function applyBorder(cell) {
  cell.border = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
}

function applyBaseCellStyle(cell) {
  applyBorder(cell);

  cell.alignment = {
    vertical: "middle",
    wrapText: true,
  };

  cell.font = {
    size: 10,
  };
}

function applyHeaderCellStyle(cell) {
  applyBorder(cell);

  cell.font = {
    bold: true,
    size: 11,
  };

  cell.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
  };

  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFEDEDED" },
  };
}

function applyTitleStyle(cell, size = 18) {
  cell.font = {
    bold: true,
    size,
  };

  cell.alignment = {
    horizontal: "center",
    vertical: "middle",
  };
}

function setColumnWidths(worksheet) {
  for (let col = 1; col <= LAST_COL; col += 1) {
    worksheet.getColumn(col).width = 5;
  }

  worksheet.getColumn("A").width = 14;
  worksheet.getColumn("B").width = 4;
  worksheet.getColumn("C").width = 4;
  worksheet.getColumn("D").width = 4;
  worksheet.getColumn("E").width = 4;
  worksheet.getColumn("F").width = 62;
  worksheet.getColumn("G").width = 4;
  worksheet.getColumn("H").width = 4;
  worksheet.getColumn("I").width = 4;
  worksheet.getColumn("J").width = 4;
  worksheet.getColumn("K").width = 4;
  worksheet.getColumn("L").width = 4;
  worksheet.getColumn("M").width = 4;
  worksheet.getColumn("N").width = 4;
  worksheet.getColumn("O").width = 4;
  worksheet.getColumn("P").width = 4;
  worksheet.getColumn("Q").width = 4;
  worksheet.getColumn("R").width = 4;
  worksheet.getColumn("S").width = 13;
  worksheet.getColumn("T").width = 4;
}

function addTopSection(worksheet, reportTitle, venueName) {
  worksheet.mergeCells("A1:T1");
  worksheet.getCell("A1").value = reportTitle || "Inventory Report";
  applyTitleStyle(worksheet.getCell("A1"), 18);
  worksheet.getRow(1).height = 30;

  worksheet.mergeCells("A2:T2");
  worksheet.getCell("A2").value = venueName
    ? `Venue: ${venueName}`
    : "Venue: All Venues";
  applyTitleStyle(worksheet.getCell("A2"), 12);
  worksheet.getRow(2).height = 22;

  worksheet.mergeCells("A3:T3");
  worksheet.getCell("A3").value = `Generated: ${new Date().toLocaleDateString()}`;
  worksheet.getCell("A3").alignment = {
    horizontal: "center",
    vertical: "middle",
  };
  worksheet.getRow(3).height = 22;

  worksheet.mergeCells("A5:T5");
  worksheet.getCell("A5").value =
    "Use column S for the inventory count. Item code is in column A. Item name is in column F.";
  worksheet.getCell("A5").font = {
    italic: true,
    size: 10,
  };
  worksheet.getCell("A5").alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
  };
  worksheet.getRow(5).height = 24;
}

function addTableHeader(worksheet) {
  const row = worksheet.getRow(HEADER_ROW);

  row.getCell(CODE_COL).value = "CODE";
  row.getCell(NAME_COL).value = "ITEM NAME";
  row.getCell(COUNT_COL).value = "COUNT";

  for (let col = 1; col <= LAST_COL; col += 1) {
    applyHeaderCellStyle(row.getCell(col));
  }

  row.height = 28;
}

function addItemRows(worksheet, items, includeCounts) {
  const reportItems = prepareInventoryReportItems(items);

  reportItems.forEach((item, index) => {
    const rowNumber = FIRST_ITEM_ROW + index;
    const row = worksheet.getRow(rowNumber);

    const code = normalizeCode(getItemCode(item));
    const name = normalizeText(getItemName(item));
    const count = toCountNumber(getItemCount(item));

    for (let col = 1; col <= LAST_COL; col += 1) {
      applyBaseCellStyle(row.getCell(col));
    }

    row.getCell(CODE_COL).value = code;
    row.getCell(NAME_COL).value = name;
    row.getCell(COUNT_COL).value = includeCounts ? count : "";

    row.getCell(CODE_COL).alignment = {
      horizontal: "left",
      vertical: "middle",
      wrapText: true,
    };

    row.getCell(NAME_COL).alignment = {
      horizontal: "left",
      vertical: "middle",
      wrapText: true,
    };

    row.getCell(COUNT_COL).alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };

    row.getCell(COUNT_COL).numFmt = "0";

    row.height = 38;
  });

  return reportItems.length;
}

function addEmptyRowsIfNeeded(worksheet, itemCount) {
  const minimumRows = 25;

  if (itemCount >= minimumRows) return;

  for (let index = itemCount; index < minimumRows; index += 1) {
    const rowNumber = FIRST_ITEM_ROW + index;
    const row = worksheet.getRow(rowNumber);

    for (let col = 1; col <= LAST_COL; col += 1) {
      applyBaseCellStyle(row.getCell(col));
    }

    row.height = 38;
  }
}

function configureWorksheet(worksheet, totalRows) {
  worksheet.views = [
    {
      state: "frozen",
      ySplit: HEADER_ROW,
    },
  ];

  worksheet.autoFilter = {
    from: {
      row: HEADER_ROW,
      column: 1,
    },
    to: {
      row: HEADER_ROW,
      column: LAST_COL,
    },
  };

  worksheet.pageSetup = {
    orientation: "landscape",
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: `${HEADER_ROW}:${HEADER_ROW}`,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.35,
      bottom: 0.35,
      header: 0.2,
      footer: 0.2,
    },
  };

  worksheet.pageSetup.printArea = `A1:T${totalRows}`;

  worksheet.headerFooter.oddFooter =
    "&LInventory Report&CPage &P of &N&RGenerated by Make Inventory";
}

export function buildInventoryWorkbook({
  items = [],
  venueName = "",
  reportTitle = "Inventory Report",
  includeCounts = true,
} = {}) {
  const workbook = new ExcelJS.Workbook();

  workbook.creator = "Make Inventory";
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet("Report");

  setColumnWidths(worksheet);
  addTopSection(worksheet, reportTitle, venueName);
  addTableHeader(worksheet);

  const itemCount = addItemRows(worksheet, items, includeCounts);
  addEmptyRowsIfNeeded(worksheet, itemCount);

  const totalRows = FIRST_ITEM_ROW + Math.max(itemCount, 25) - 1;

  configureWorksheet(worksheet, totalRows);

  return workbook;
}

import * as XLSX from "xlsx";

export const readExcelFile = async (file, options = {}) => {
  if (!file) {
    throw new Error("No Excel file selected.");
  }

  const arrayBuffer = await file.arrayBuffer();

  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    cellStyles: true,
    cellNF: true,
    cellFormula: true,
    dense: options.dense ?? false,
  });

  return {
    workbook,
    arrayBuffer,
    fileName: file.name || "",
  };
};

export const workbookToRows = (workbook, options = {}) => {
  const {
    sheetName,
    header = 1,
    defval = "",
    raw = false,
  } = options;

  const selectedSheetName = sheetName || workbook.SheetNames?.[0];

  if (!selectedSheetName) {
    return {
      sheetName: "",
      rows: [],
    };
  }

  const worksheet = workbook.Sheets[selectedSheetName];

  if (!worksheet) {
    return {
      sheetName: selectedSheetName,
      rows: [],
    };
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header,
    defval,
    raw,
  });

  return {
    sheetName: selectedSheetName,
    rows,
  };
};

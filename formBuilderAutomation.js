/**
 * Spreadsheet Form Builder & Pagination Engine
 * ------------------------------------------------------------
 * This script processes input data, formats it, and generates
 * paginated output sheets with headers and structured layout.
 *
 * Features:
 * - auto form numbering
 * - input recording
 * - grouped row formatting
 * - pagination with headers
 * - merge-safe rendering
 *
 * NOTE:
 * Replace sheet names with your own.
 */

/***********************
 * CONFIG
 ***********************/
const CONFIG = {
  SHEET_INPUT: "INPUT_SHEET",
  SHEET_OUTPUT_1: "INTERMEDIATE_OUTPUT",
  SHEET_OUTPUT_2: "FINAL_OUTPUT",
  SHEET_HEADER_1: "HEADER_PAGE_1",
  SHEET_HEADER_2: "HEADER_OTHER_PAGES",
  SHEET_RECORDS: "RECORDS",
  SHEET_FORM_NO: "FORM_NUMBER",

  HEADER1_ROWS: 10,
  HEADER2_ROWS: 6,

  PAGE1_ROWS: 30,
  PAGE_N_ROWS: 40,

  COLS: 10
};


/***********************
 * MENU
 ***********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Form Builder")
    .addItem("Generate Form", "runFormBuilder")
    .addToUi();
}


/***********************
 * MAIN ENTRY
 ***********************/
function runFormBuilder() {

  const formNumber = getNextFormNumber_();

  recordInputData_();

  formatIntermediateSheet_();

  SpreadsheetApp.flush();

  buildFinalOutput_(formNumber);

}


/***********************
 * FORM NUMBER
 ***********************/
function getNextFormNumber_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEET_FORM_NO);

  const lastRow = sh.getLastRow();

  let lastNumber = 0;

  if (lastRow > 0) {
    lastNumber = Number(sh.getRange(lastRow, 1).getValue()) || 0;
  }

  const next = lastNumber + 1;

  sh.getRange(lastRow + 1, 1).setValue(next);

  return next;
}


/***********************
 * RECORD INPUT DATA
 ***********************/
function recordInputData_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const input = ss.getSheetByName(CONFIG.SHEET_INPUT);
  const record = ss.getSheetByName(CONFIG.SHEET_RECORDS);

  const values = input.getRange(2, 1, input.getLastRow(), 3).getValues();

  const filtered = values.filter(r => r.some(v => v !== ""));

  if (filtered.length === 0) return;

  const timestamp = new Date();

  const output = filtered.map(r => [
    timestamp,
    ...r
  ]);

  record
    .getRange(record.getLastRow() + 1, 1, output.length, output[0].length)
    .setValues(output);

}


/***********************
 * FORMAT INTERMEDIATE
 ***********************/
function formatIntermediateSheet_() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_OUTPUT_1);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).breakApart();

  // Example grouping logic
  const values = sheet.getRange(2,3,lastRow,1).getValues();

  let start = 2;

  for (let i = 0; i < values.length; i++) {

    const row = i + 2;

    if (values[i][0] !== "" && row !== start) {

      sheet.getRange(start,3,row-start,1).merge();

      start = row;

    }
  }

}


/***********************
 * BUILD FINAL OUTPUT
 ***********************/
function buildFinalOutput_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const src = ss.getSheetByName(CONFIG.SHEET_OUTPUT_1);
  const dst = ss.getSheetByName(CONFIG.SHEET_OUTPUT_2) || ss.insertSheet(CONFIG.SHEET_OUTPUT_2);

  dst.clear();

  const lastRow = src.getLastRow();

  let row = 2;
  let destRow = 1;

  while (row <= lastRow) {

    const maxRows = destRow === 1
      ? CONFIG.PAGE1_ROWS
      : CONFIG.PAGE_N_ROWS;

    const end = Math.min(row + maxRows - 1, lastRow);

    const range = src.getRange(row, 1, end - row + 1, CONFIG.COLS);

    range.copyTo(
      dst.getRange(destRow, 1),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

    destRow += maxRows;
    row = end + 1;

  }

}

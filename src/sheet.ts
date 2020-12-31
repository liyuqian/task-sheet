export default { createSheet, onEdit };

const kTasks = 'Tasks';
const kNextUniqueId = 'NEXT_UNIQUE_ID';

// Start from this 1234567 for easier search and spotting.
const kFirstUniqueId = 1234567;

const kNextUniqueIdRow = 2;

const kIdColName = 'id';
const kTitleColName = 'title';
const kDetailsColName = 'details';
const kDueDateColName = 'due date';
const kLabelsColName = 'labels';
const kDependenciesColName = 'dependencies';
const kCompleteDateColName = 'complete date';
const kObsoleteDateColName = 'obsolete date';
const kColNames = [
  kIdColName,
  kTitleColName,
  kDetailsColName,
  kDueDateColName,
  kLabelsColName,
  kDependenciesColName,
  kCompleteDateColName,
  kObsoleteDateColName,
];
const kColCount = kColNames.length;

// All indices below are 1-based instead 0-based.
const kIdColIndex = kColNames.indexOf(kIdColName) + 1;
const kTitleColIndex = kColNames.indexOf(kTitleColName) + 1;

function createSheet(): void {
  const spreadsheet = SpreadsheetApp.create('tasks');
  const sheet1 = spreadsheet.getActiveSheet();
  const sheet2 = spreadsheet.insertSheet();
  const sheet3 = spreadsheet.insertSheet();
  const sheet4 = spreadsheet.insertSheet();
  const sheet5 = spreadsheet.insertSheet();
  sheet1.setName('Views');
  sheet2.setName(kTasks);
  sheet3.setName('Plan');
  sheet4.setName('Archived-Tasks');
  sheet5.setName('Archived-Plan');
  initTask(sheet2);
  ScriptApp.newTrigger('onEdit').forSpreadsheet(spreadsheet).onEdit().create();
}

function initTask(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.setFrozenRows(1);
  const row1 = sheet.getRange(1, 1, 1, kColCount);
  row1.setHorizontalAlignment('center');
  for (let i = 0; i < kColCount; i += 1) {
    row1.getCell(1, i + 1).setValue(kColNames[i]);
  }
  const row2 = sheet.getRange(2, 1, 1, kColCount);
  row2.getCell(1, kIdColIndex).setValue(kFirstUniqueId);
  row2.getCell(1, kTitleColIndex).setValue(kNextUniqueId);
  if (row2.getRow() !== kNextUniqueIdRow) {
    throw new Error('Mismatched kNextUniqueIdRow '
        + `(${row2.getRow()} != ${kNextUniqueIdRow})`);
  }
  sheet.hideRow(row2);
}

interface EditEvent {
  oldValue: any;
  value: any;
  range: GoogleAppsScript.Spreadsheet.Range;
  source: GoogleAppsScript.Spreadsheet.Spreadsheet;
}

function onEdit(e: EditEvent): void {
  const sheet = e.range.getSheet();
  if (sheet.getName() === kTasks) {
    onTasksEdit(e);
  }
}

function onTasksEdit(e: EditEvent): void {
  const sheet = e.range.getSheet();
  for (let i = 1; i <= e.range.getNumRows(); i += 1) {
    genUniqueIdIfNeeded(sheet, e.range.getRowIndex() + i - 1);
  }
}

function genUniqueIdIfNeeded(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowIndex: number,
): void {
  const fullRow = sheet.getRange(rowIndex, 1, 1, kColCount);
  const titleCell = fullRow.getCell(1, kTitleColIndex);
  const idCell = fullRow.getCell(1, kIdColIndex);
  if (idCell.isBlank() && !titleCell.isBlank()) {
    const nextIdCell = sheet.getRange(kNextUniqueIdRow, kIdColIndex);
    idCell.setValue(nextIdCell.getValue());
    nextIdCell.setValue(parseInt(idCell.getValue(), 10) + 1);
    Logger.log(`Set id ${idCell.getValue()} for row ${fullRow.getRowIndex()}; `
      + `next unique id updated to ${nextIdCell.getValue()}`);
  }
}

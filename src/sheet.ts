export default { createSheet, onEdit };

const kTasks = 'Tasks';

// No collision is expected for ~36^(length / 2) random IDs.
// For length = 8, that's about 1000^2 = 1 million.
const kRandomIdLength = 8;

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
    genIdIfNeeded(sheet, e.range.getRowIndex() + i - 1);
  }
}

function genIdIfNeeded(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowIndex: number,
): void {
  const fullRow = sheet.getRange(rowIndex, 1, 1, kColCount);
  const titleCell = fullRow.getCell(1, kTitleColIndex);
  const idCell = fullRow.getCell(1, kIdColIndex);
  if (idCell.isBlank() && !titleCell.isBlank()) {
    idCell.setValue(genRandomId(kRandomIdLength));
    idCell.setFontFamily('Courier New')
    Logger.log(`Set id ${idCell.getValue()} for row ${fullRow.getRowIndex()}.`);
  }
}

function genRandomId(length: number) {
  let id = '';
  for (let i = 0; i < length; i += 1) {
    id += Math.floor(Math.random() * 36).toString(36);
  }
  return id;
}

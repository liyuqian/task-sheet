export default { createSheet, onEdit };

const kTasks = 'Tasks';
const kPlan = 'Plan';

// No collision is expected for ~36^(length / 2) random IDs.
// For length = 8, that's about 1000^2 = 1 million.
const kRandomIdLength = 8;

const kIdColName = 'id';
const kTitleColName = 'title';
const kDetailsColName = 'details';
const kDueDateColName = 'due date';
const kLabelsColName = 'labels';
const kDependenciesColName = 'dependencies';
const kStartDateColName = 'start date';
const kCompleteDateColName = 'complete date';
const kObsoleteDateColName = 'obsolete date';
const kProgressColName = 'progress';
const kNotesColName = 'notes';

const kTasksColNames = [
  kIdColName,
  kTitleColName,
  kDetailsColName,
  kDueDateColName,
  kLabelsColName,
  kDependenciesColName,
  kStartDateColName,
  kCompleteDateColName,
  kObsoleteDateColName,
];
const kTasksColCount = kTasksColNames.length;

const kPlanColNames = [
  kIdColName,
  kTitleColName,
  kDetailsColName,
  kProgressColName,
  kNotesColName,
];
const kPlanColCount = kPlanColNames.length;
const kCommonColCount = 3; // identical columns between tasks and plan sheets

// All indices below are 1-based instead 0-based.
const kIdColIndex = kTasksColNames.indexOf(kIdColName) + 1;
const kTitleColIndex = kTasksColNames.indexOf(kTitleColName) + 1;
const kStartDateColIndex = kTasksColNames.indexOf(kStartDateColName) + 1;

function createSheet(): void {
  checkIntegrity();
  const spreadsheet = SpreadsheetApp.create('tasks');
  const sheet1 = spreadsheet.getActiveSheet();
  const sheet2 = spreadsheet.insertSheet();
  const sheet3 = spreadsheet.insertSheet();
  const sheet4 = spreadsheet.insertSheet();
  const sheet5 = spreadsheet.insertSheet();
  sheet1.setName('Views');
  sheet2.setName(kTasks);
  sheet3.setName(kPlan);
  sheet4.setName('Archived-Tasks');
  sheet5.setName('Archived-Plan');
  initTasks(sheet2);
  initPlan(sheet3);
  ScriptApp.newTrigger('onEdit').forSpreadsheet(spreadsheet).onEdit().create();
}

function checkIntegrity(): void {
  for (let i = 0; i < kCommonColCount; i += 1) {
    if (kTasksColNames[i] !== kPlanColNames[i]) {
      throw new Error(`Mismatched column ${i}: '
          '${kTasksColNames[i]} != ${kPlanColNames[i]}`);
    }
  }
}

function initTasks(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.setFrozenRows(1);
  const row1 = sheet.getRange(1, 1, 1, kTasksColCount);
  row1.setHorizontalAlignment('center');
  for (let i = 0; i < kTasksColCount; i += 1) {
    row1.getCell(1, i + 1).setValue(kTasksColNames[i]);
  }
}

function initPlan(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.setFrozenRows(1);
  const row1 = sheet.getRange(1, 1, 1, kPlanColCount);
  row1.setHorizontalAlignment('center');
  for (let i = 0; i < kPlanColCount; i += 1) {
    row1.getCell(1, i + 1).setValue(kPlanColNames[i]);
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
    copyToPlanIfStartedToday(sheet, e.range.getRowIndex() + i - 1);
  }
}

function copyToPlanIfStartedToday(
  tasksSheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowIndex: number,
): void {
  const fullRow = tasksSheet.getRange(rowIndex, 1, 1, kTasksColCount);
  const startDateCell = fullRow.getCell(1, kStartDateColIndex);
  const taskId: string = fullRow.getCell(1, kIdColIndex).getValue();
  if (startDateCell.isBlank()) {
    Logger.log(`Skip ${taskId} as start date is blank.`);
    return;
  }
  const startDate = new Date(startDateCell.getValue());
  const today = new Date();
  if (format(startDate) !== format(today)) {
    Logger.log(`Skip ${taskId} as ${format(startDate)} is not today.`);
    return;
  }

  const planSheet = tasksSheet.getParent().getSheetByName(kPlan);
  const planRowCount = planSheet.getDataRange().getNumRows();
  const planTaskIds = planSheet.getRange(1, 1, planRowCount);
  const finder = planTaskIds.createTextFinder(taskId);
  if (finder.findAll().length > 0) {
    Logger.log(`Skip existing ${taskId} (found ${finder.findAll().length})`);
    return;
  }
  const copyRange = tasksSheet.getRange(rowIndex, 1, 1, kCommonColCount);
  copyRange.copyTo(planSheet.getRange(planRowCount + 1, 1));
}

function format(date: Date): string {
  return `${date.getMonth()}/${date.getDay()}/${date.getFullYear()}`;
}

function genIdIfNeeded(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowIndex: number,
): void {
  const fullRow = sheet.getRange(rowIndex, 1, 1, kTasksColCount);
  const titleCell = fullRow.getCell(1, kTitleColIndex);
  const idCell = fullRow.getCell(1, kIdColIndex);
  if (idCell.isBlank() && !titleCell.isBlank()) {
    idCell.setValue(genRandomId(kRandomIdLength));
    idCell.setFontFamily('Courier New');
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

import {
  kTasks,
  kPlan,
  kTasksColNames,
  kTasksColCount,
  kPlanColNames,
  kCommonColCount,
  kIdColIndex,
  kTitleColIndex,
  kStartDateColIndex,
  EditEvent,
  findRowIndexById,
  format,
} from './common';

import { initPlan, onPlanEdit } from './plan';

export {
  createSheet,
  deleteOldTriggers,
  onEdit,
  kRandomIdLength,
};

function createSheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
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
  return spreadsheet;
}

// Be careful, this will remove all triggers associated with this project.
// Hence all previous spreadsheets created by this project will lose their
// triggers.
function deleteOldTriggers(): void {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
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

function onEdit(e: EditEvent): void {
  const sheet = e.range.getSheet();
  if (sheet.getName() === kTasks) {
    onTasksEdit(e);
  } else if (sheet.getName() === kPlan) {
    onPlanEdit(e);
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
  const findResult = findRowIndexById(planSheet, taskId);
  if (findResult !== -1) {
    Logger.log(`Skip existing ${taskId} at row ${findResult}.`);
    return;
  }
  const copyRange = tasksSheet.getRange(rowIndex, 1, 1, kCommonColCount);
  const planRowCount = planSheet.getDataRange().getNumRows();
  copyRange.copyTo(planSheet.getRange(planRowCount + 1, 1));
}

// No collision is expected for ~36^(length / 2) random IDs.
// For length = 8, that's about 1000^2 = 1 million.
const kRandomIdLength = 8;

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

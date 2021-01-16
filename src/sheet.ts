import {
  kTasks,
  kPlan,
  kArchivedPlan,
  kArchivedTasks,
  kTasksColNames,
  kTasksColCount,
  kPlanColNames,
  kCommonColCount,
  kIdColIndex,
  kTitleColIndex,
  EditEvent,
  findRowIndexById,
  copyTo,
  kPlanColCount,
  kProgressColIndex,
  findTaskRowById,
  kObsoleteDateColIndex,
  kCompleteDateColIndex,
  kDueDateColIndex,
} from './common';

import { copyToPlanIfStartedToday, initPlan, onPlanEdit } from './plan';

export {
  createSheet,
  deleteOldTriggers,
  onEdit,
  onOpen,
  archive,
  filterDueSoon,
  kRandomIdLength,
};

// TODO NEXT:
//   sync edits between Tasks and Plan?
function createSheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  checkIntegrity();
  const spreadsheet = SpreadsheetApp.create('tasks');
  const sheet1 = spreadsheet.getActiveSheet();
  const sheet2 = spreadsheet.insertSheet();
  const sheet3 = spreadsheet.insertSheet();
  const sheet4 = spreadsheet.insertSheet();
  sheet1.setName(kTasks);
  sheet2.setName(kPlan);
  sheet3.setName(kArchivedTasks);
  sheet4.setName(kArchivedPlan);
  initTasks(sheet1);
  initPlan(sheet2);
  initTasks(sheet3);
  initPlan(sheet4);
  ScriptApp.newTrigger('onEdit').forSpreadsheet(spreadsheet).onEdit().create();
  ScriptApp.newTrigger('onOpen').forSpreadsheet(spreadsheet).onOpen().create();
  return spreadsheet;
}

function onOpen(): void {
  const menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Archive completed and obsolete tasks', 'archive');
  menu.addItem('Filter due soon tasks', 'filterDueSoon');
  menu.addToUi();
}

function filterDueSoon(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const tasksSheet = spreadsheet.getSheetByName(kTasks);
  const filter = tasksSheet.getDataRange().createFilter();
  const emptyCriteria = SpreadsheetApp
    .newFilterCriteria().whenCellEmpty().build();
  filter.setColumnFilterCriteria(kCompleteDateColIndex, emptyCriteria);
  filter.setColumnFilterCriteria(kObsoleteDateColIndex, emptyCriteria);
  filter.sort(kDueDateColIndex, true);
  spreadsheet.setActiveSheet(tasksSheet);
}

function archive(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const spreadsheet = ss || SpreadsheetApp.getActive();
  const completeCount = archiveCompleted(spreadsheet);
  const obsoleteCount = archiveObsolete(spreadsheet);
  if (SpreadsheetApp.getActive() == null) {
    // There's no active spreadsheet. So don't pop the follwoing alert. (This
    // usually happens during tests.)
    return;
  }
  SpreadsheetApp.getUi().alert(
    `Archived ${completeCount} completed and ${obsoleteCount} obsolete tasks.`,
  );
}

// This function removes completed tasks from tasks and plan sheets, and put
// them into archived tasks and plan sheets. Therefore, rows must be removed
// from bottom to top to avoid index conflicts. We also assume that no other
// edits may happen during this function, or there will be data racing problems.
function archiveCompleted(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
): number {
  Logger.log('Archive completed');
  const tasksSheet = spreadsheet.getSheetByName(kTasks);
  const planSheet = spreadsheet.getSheetByName(kPlan);
  const archivedTasksSheet = spreadsheet.getSheetByName(kArchivedTasks);
  const archivedPlanSheet = spreadsheet.getSheetByName(kArchivedPlan);
  const planRange = planSheet.getDataRange();
  let archiveCount = 0;
  const toBeDeleted = [];
  // Copy in increasing order to maintain the tasks order. Later delete in
  // reversed order to avoid index conflicts as deleted rows may affect the
  // indices of all following rows.
  for (let r = 2; r <= planRange.getNumRows(); r += 1) {
    const row = planSheet.getRange(r, 1, 1, kPlanColCount);
    const progress = parseFloat(row.getCell(1, kProgressColIndex).getValue());
    if (progress === 1) {
      toBeDeleted.push(r);
      const id = row.getCell(1, kIdColIndex).getValue();
      const taskRow = findTaskRowById(id, tasksSheet);
      copyTo(taskRow, archivedTasksSheet);
      copyTo(row, archivedPlanSheet);
      tasksSheet.deleteRow(taskRow.getRowIndex());
      archiveCount += 1;
    }
  }
  for (let i = toBeDeleted.length - 1; i >= 0; i -= 1) {
    planSheet.deleteRow(toBeDeleted[i]);
  }
  return archiveCount;
}

function archiveObsolete(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
): number {
  Logger.log('Archive obsolete');
  const tasksSheet = spreadsheet.getSheetByName(kTasks);
  const planSheet = spreadsheet.getSheetByName(kPlan);
  const archivedTasksSheet = spreadsheet.getSheetByName(kArchivedTasks);
  const archivedPlanSheet = spreadsheet.getSheetByName(kArchivedPlan);
  const tasksRange = tasksSheet.getDataRange();
  let archiveCount = 0;
  const toBeDeleted = [];
  // Copy in increasing order to maintain the tasks order. Later delete in
  // reversed order to avoid index conflicts as deleted rows may affect the
  // indices of all following rows.
  for (let r = 2; r <= tasksRange.getNumRows(); r += 1) {
    const row = tasksSheet.getRange(r, 1, 1, kTasksColCount);
    if (!row.getCell(1, kObsoleteDateColIndex).isBlank()) {
      toBeDeleted.push(r);
      archiveCount += 1;
      const id = row.getCell(1, kIdColIndex).getValue();
      copyTo(row, archivedTasksSheet);
      const planRowIndex = findRowIndexById(planSheet, id);
      if (planRowIndex !== -1) {
        const planRow = planSheet.getRange(planRowIndex, 1, 1, kPlanColCount);
        copyTo(planRow, archivedPlanSheet);
        planSheet.deleteRow(planRowIndex);
      }
    }
  }
  for (let i = toBeDeleted.length - 1; i >= 0; i -= 1) {
    tasksSheet.deleteRow(toBeDeleted[i]);
  }
  return archiveCount;
}

// Be careful, if spreadsheetId isn't given, this will remove all triggers
// associated with this project. Hence all previous spreadsheets created by this
// project will lose their triggers.
function deleteOldTriggers(spreadsheetId: string): void {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (spreadsheetId == null
        || trigger.getTriggerSourceId() === spreadsheetId) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
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

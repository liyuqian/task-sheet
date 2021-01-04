import {
  copyTo,
  EditEvent,
  findRowIndexById,
  format,
  kArchivedPlan,
  kArchivedTasks,
  kCompleteDateColIndex,
  kIdColIndex,
  kPlanColCount,
  kPlanColNames,
  kProgressColIndex,
  kTasks,
  kTasksColCount,
  kTasksColNames,
} from './common';

export { initPlan, onPlanEdit };

function initPlan(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  sheet.setFrozenRows(1);
  const row1 = sheet.getRange(1, 1, 1, kPlanColCount);
  row1.setHorizontalAlignment('center');
  for (let i = 0; i < kPlanColCount; i += 1) {
    row1.getCell(1, i + 1).setValue(kPlanColNames[i]);
  }
}

function onPlanEdit(e: EditEvent): void {
  const sheet = e.range.getSheet();
  for (let i = 1; i <= e.range.getNumRows(); i += 1) {
    markCompletedIfSo(sheet, e.range.getRowIndex() + i - 1);
  }
}

// TODO NEXT:
//   1. test and handle multiple completed tasks and row index changes
//   2. handle obsolete (in the tasks.ts?)
function markCompletedIfSo(
  planSheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowIndex: number,
): void {
  const fullRow = planSheet.getRange(rowIndex, 1, 1, kPlanColCount);
  const progress = parseFloat(fullRow.getCell(1, kProgressColIndex).getValue());
  const id = fullRow.getCell(1, kIdColIndex).getValue();
  if (progress !== 1) {
    Logger.log(`Skip plan update for ${id} as its progress ${progress} != 1`);
    return;
  }
  Logger.log(`Try to mark ${id} as completed.`);
  const tasksSheet = planSheet.getParent().getSheetByName(kTasks);
  const taskRowIndex = findRowIndexById(tasksSheet, id);
  if (taskRowIndex === -1) {
    const message = `Task ${id} not found in the ${kTasksColNames} sheet!`;
    SpreadsheetApp.getUi().alert(message);
    throw new Error(message);
  }
  const taskRow = tasksSheet.getRange(taskRowIndex, 1, 1, kTasksColCount);
  const completeDateCell = taskRow.getCell(1, kCompleteDateColIndex);
  if (!completeDateCell.isBlank()) {
    Logger.log(`Skip row ${taskRowIndex} with existing complete date `
        + `${completeDateCell.getValue()}`);
    return;
  }
  completeDateCell.setValue(format(new Date()));

  const archivedTasksSheet = planSheet
    .getParent().getSheetByName(kArchivedTasks);
  copyTo(taskRow, archivedTasksSheet);

  const archivedPlanSheet = planSheet.getParent().getSheetByName(kArchivedPlan);
  copyTo(fullRow, archivedPlanSheet);

  tasksSheet.deleteRow(taskRowIndex);
  planSheet.deleteRow(rowIndex);
}

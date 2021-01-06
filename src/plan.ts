import {
  copyTo,
  EditEvent,
  findRowIndexById,
  findTaskRowById,
  format,
  kCommonColCount,
  kCompleteDateColIndex,
  kIdColIndex,
  kPlan,
  kPlanColCount,
  kPlanColNames,
  kProgressColIndex,
  kStartDateColIndex,
  kTasks,
  kTasksColCount,
} from './common';

export { initPlan, onPlanEdit, copyToPlanIfStartedToday };

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
  for (let i = e.range.getNumRows(); i >= 1; i -= 1) {
    markCompletedIfSo(sheet, e.range.getRowIndex() + i - 1);
  }
}

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
  const taskRow = findTaskRowById(id, tasksSheet);
  const completeDateCell = taskRow.getCell(1, kCompleteDateColIndex);
  if (!completeDateCell.isBlank()) {
    Logger.log(`Skip row ${taskRow.getRowIndex()} with existing complete date `
        + `${completeDateCell.getValue()}`);
    return;
  }
  completeDateCell.setValue(format(new Date()));
}

function copyToPlanIfStartedToday(
  tasksSheet: GoogleAppsScript.Spreadsheet.Sheet,
  taskRowIndex: number,
): void {
  const fullRow = tasksSheet.getRange(taskRowIndex, 1, 1, kTasksColCount);
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
  const copyRange = tasksSheet.getRange(taskRowIndex, 1, 1, kCommonColCount);
  Logger.log(`Copy ${taskId} from ${kTasks} to ${kPlan}`);
  copyTo(copyRange, planSheet);
}

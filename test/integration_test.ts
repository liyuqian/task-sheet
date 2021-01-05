import {
  EditEvent,
  findRowIndexById,
  format,
  kArchivedPlan,
  kArchivedTasks,
  kCompleteDateColIndex,
  kIdColIndex,
  kPlan,
  kPlanColCount,
  kProgressColIndex,
  kStartDateColIndex,
  kTasks,
  kTasksColCount,
} from '../src/common';

import {
  archiveCompleted,
  createSheet,
  kRandomIdLength,
  onEdit,
} from '../src/sheet';

export default { testAll };

type Range = GoogleAppsScript.Spreadsheet.Range;
type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;

class TestEditEvent implements EditEvent {
  oldValue: any;

  value: any;

  range: Range;

  source: Spreadsheet;

  constructor(oldValue: any, value: any, range: Range, source: Spreadsheet) {
    this.oldValue = oldValue;
    this.value = value;
    this.range = range;
    this.source = source;
  }
}

const kTestTaskTitle = 'test title';
const kDummyTaskRow = [
  '', // id
  kTestTaskTitle,
  '', // description
  '', // due date
  '', // labels
  '', // dependencies
  '', // start date
  '', // complete date
  '', // obsolete date
];
const kDummyTaskValues = [kDummyTaskRow];

function testAll(): void {
  testFlow();
  testMultipleCompletion();
}

function testFlow(): void {
  Logger.log('Start test flow.');
  const spreadsheet = createSheet();
  try {
    // 1. Test id generation.
    const tasksSheet = spreadsheet.getSheetByName(kTasks);
    const taskRow = tasksSheet.getRange(2, 1, 1, kTasksColCount);
    taskRow.setValues(kDummyTaskValues);
    const taskEditEvent = new TestEditEvent(null, null, taskRow, spreadsheet);
    onEdit(taskEditEvent);
    const idCell = taskRow.getCell(1, kIdColIndex);
    if (idCell.isBlank()) {
      throw new Error('The id is unexpectedly blank.');
    }
    if ((idCell.getValue() as string).length !== kRandomIdLength) {
      throw new Error(`The generated id ${idCell.getValue()} has an `
          + `unexpected length that does not match ${kRandomIdLength}.`);
    }
    Logger.log('1. Test id generation passed.');

    // 2. Test copy to plan.
    const today = format(new Date());
    taskRow.getCell(1, kStartDateColIndex).setValue(today);
    onEdit(taskEditEvent);
    const id = idCell.getValue();
    const planSheet = spreadsheet.getSheetByName(kPlan);
    const planRow = planSheet.getRange(2, 1, 1, kPlanColCount);
    const expectedPlanRowValues = [[
      id,
      kTestTaskTitle,
      '', // details
      '', // progress
      '', // notes
    ]];
    expectValuesMatch(planRow.getValues(), expectedPlanRowValues);
    Logger.log('2. Test copy to plan passed.');

    // 3. Test mark as completed.
    planRow.getCell(1, kProgressColIndex).setValue(1);
    const planEditEvent = new TestEditEvent(null, null, planRow, spreadsheet);
    onEdit(planEditEvent);
    const completeDateValue = taskRow
      .getCell(1, kCompleteDateColIndex).getValue();
    if (format(completeDateValue) !== today) {
      throw new Error('Unexpected complete date '
          + `${completeDateValue} != ${today}`);
    }
    Logger.log('3. Test mark as completed passed.');

    // 4. Test archive completed
    archiveCompleted(spreadsheet);
    const archivedTaskRow = spreadsheet
      .getSheetByName(kArchivedTasks).getRange(2, 1, 1, kTasksColCount);
    const archivedTaskId = archivedTaskRow.getCell(1, kIdColIndex).getValue();
    if (archivedTaskId !== id) {
      throw new Error(`Unexpected archivedTaskId ${archivedTaskId} != ${id}`);
    }
    const archivedPlanRow = spreadsheet
      .getSheetByName(kArchivedPlan).getRange(2, 1, 1, kPlanColCount);
    expectedPlanRowValues[0][kProgressColIndex - 1] = 1;
    expectValuesMatch(archivedPlanRow.getValues(), expectedPlanRowValues);
    const taskRowFound = findRowIndexById(tasksSheet, id);
    const planRowFound = findRowIndexById(planSheet, id);
    if (taskRowFound !== -1 || planRowFound !== -1) {
      throw new Error(`Completed rows aren't deleted (${taskRowFound}, `
          + `${planRowFound}).`);
    }
  } finally {
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
  }
}

function testMultipleCompletion(): void {
  Logger.log('Start test multiple completion.');
  const spreadsheet = createSheet();
  try {
    const kTestRowCount = 3;
    const tasksSheet = spreadsheet.getSheetByName(kTasks);
    const rowValues = [...kDummyTaskRow];
    const today = format(new Date());
    rowValues[kStartDateColIndex - 1] = today;
    const taskRows = tasksSheet.getRange(2, 1, kTestRowCount, kTasksColCount);
    taskRows.setValues([rowValues, rowValues, rowValues]);
    const taskEditEvent = new TestEditEvent(null, null, taskRows, spreadsheet);
    onEdit(taskEditEvent);
    const planSheet = spreadsheet.getSheetByName(kPlan);
    const progressRange = planSheet
      .getRange(2, kProgressColIndex, kTestRowCount, 1);
    for (let i = 1; i <= kTestRowCount; i += 1) {
      progressRange.getCell(i, 1).setValue(1);
    }
    const planEditEvent = new TestEditEvent(
      null, null, progressRange, spreadsheet,
    );
    onEdit(planEditEvent);
    archiveCompleted(spreadsheet);
    for (let i = 1; i <= kTestRowCount; i += 1) {
      const completeDate = spreadsheet
        .getSheetByName(kArchivedTasks)
        .getRange(1 + i, kCompleteDateColIndex, 1, 1)
        .getValue();
      if (format(completeDate) !== today) {
        throw new Error(`Unexpected complete date ${completeDate} for row `
          + `${1 + i}`);
      }
    }
    Logger.log('Test multiple completion passed.');
  } finally {
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
  }
}

function expectValuesMatch(actual: any[][], expected: any[][]): void {
  if (!valuesMatch(actual, expected)) {
    throw new Error(`Values don't match ${actual} != ${expected}`);
  }
}

function valuesMatch(actual: any[][], expected: any[][]): boolean {
  if (actual.length !== expected.length) {
    return false;
  }
  for (let i = 0; i < actual.length; i += 1) {
    if (actual[i].length !== expected[i].length) {
      return false;
    }
    for (let j = 0; j < actual[i].length; j += 1) {
      if (actual[i][j] !== expected[i][j]) {
        return false;
      }
    }
  }
  return true;
}

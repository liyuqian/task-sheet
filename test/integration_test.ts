import {
  EditEvent,
  format,
  kCompleteDateColIndex,
  kIdColIndex,
  kPlan,
  kPlanColCount,
  kProgressColIndex,
  kStartDateColIndex,
  kTasks,
  kTasksColCount,
} from '../src/common';

import { createSheet, kRandomIdLength, onEdit } from '../src/sheet';

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
const kDummyTaskValues = [[
  '', // id
  kTestTaskTitle,
  '', // description
  '', // due date
  '', // labels
  '', // dependencies
  '', // start date
  '', // complete date
  '', // obsolete date
]];

function testAll(): void {
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
    if (!rangeValuesMatch(planRow.getValues(), expectedPlanRowValues)) {
      throw new Error(`Values don't match ${planRow.getValues()} != `
          + `${expectedPlanRowValues}`);
    }
    Logger.log('2. Test copy to plan passed.');

    // 3. Test mark as completed.
    planRow.getCell(1, kProgressColIndex).setValue(1);
    const planEditEvent = new TestEditEvent(null, null, planRow, spreadsheet);
    onEdit(planEditEvent);
    const completeDateValue = taskRow.getCell(
      1, kCompleteDateColIndex,
    ).getValue();
    if (format(completeDateValue) !== today) {
      throw new Error('Unexpected complete date '
          + `${completeDateValue} != ${today}`);
    }
    Logger.log('3. Test mark as completed passed.');
  } finally {
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
  }
}

function rangeValuesMatch(a1: any[][], a2: any[][]): boolean {
  if (a1.length !== a2.length) {
    return false;
  }
  for (let i = 0; i < a1.length; i += 1) {
    if (a1[i].length !== a2[i].length) {
      return false;
    }
    for (let j = 0; j < a1[i].length; j += 1) {
      if (a1[i][j] !== a2[i][j]) {
        return false;
      }
    }
  }
  return true;
}

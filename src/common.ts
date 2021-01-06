export {
  kTasks,
  kPlan,
  kArchivedTasks,
  kArchivedPlan,
  kTitleColName,
  kIdColName,
  kDetailsColName,
  kDueDateColName,
  kLabelsColName,
  kDependenciesColName,
  kStartDateColName,
  kCompleteDateColName,
  kObsoleteDateColName,
  kProgressColName,
  kNotesColName,
  kTasksColNames,
  kTasksColCount,
  kPlanColNames,
  kPlanColCount,
  kCommonColCount,
  kIdColIndex,
  kTitleColIndex,
  kDueDateColIndex,
  kStartDateColIndex,
  kCompleteDateColIndex,
  kObsoleteDateColIndex,
  kProgressColIndex,
  EditEvent,
  findRowIndexById,
  findTaskRowById,
  format,
  copyTo,
};

// Sheet names
const kTasks = 'Tasks';
const kPlan = 'Plan';
const kArchivedTasks = 'Archived tasks';
const kArchivedPlan = 'Archived plan';

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
const kDueDateColIndex = kTasksColNames.indexOf(kDueDateColName) + 1;
const kStartDateColIndex = kTasksColNames.indexOf(kStartDateColName) + 1;
const kCompleteDateColIndex = kTasksColNames.indexOf(kCompleteDateColName) + 1;
const kObsoleteDateColIndex = kTasksColNames.indexOf(kObsoleteDateColName) + 1;
const kProgressColIndex = kPlanColNames.indexOf(kProgressColName) + 1;

interface EditEvent {
  oldValue: any;
  value: any;
  range: GoogleAppsScript.Spreadsheet.Range;
  source: GoogleAppsScript.Spreadsheet.Spreadsheet;
}

/// Return -1 if not found
function findRowIndexById(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  id: string,
): number {
  const rowCount = sheet.getDataRange().getNumRows();
  const taskIds = sheet.getRange(1, kIdColIndex, rowCount);
  const finder = taskIds.createTextFinder(id);
  const ranges = finder.findAll();
  if (ranges.length === 0) {
    return -1;
  }
  if (ranges.length > 1) {
    const message = `Duplicate copies (${ranges.length}) of ${id} are found!`;
    SpreadsheetApp.getUi().alert(message);
    throw new Error(message);
  }
  return ranges[0].getRowIndex();
}

function findTaskRowById(
  id: string,
  tasksSheet: GoogleAppsScript.Spreadsheet.Sheet,
): GoogleAppsScript.Spreadsheet.Range {
  const taskRowIndex = findRowIndexById(tasksSheet, id);
  if (taskRowIndex === -1) {
    const message = `Task ${id} not found in the ${kTasks} sheet!`;
    SpreadsheetApp.getUi().alert(message);
    throw new Error(message);
  }
  return tasksSheet.getRange(taskRowIndex, 1, 1, kTasksColCount);
}

function format(date: Date): string {
  return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
}

// Copy a range to a sheet as a new row next to the last row.
function copyTo(
  range: GoogleAppsScript.Spreadsheet.Range,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
): void {
  const rowCount = sheet.getDataRange().getNumRows();
  range.copyTo(sheet.getRange(rowCount + 1, 1));
}

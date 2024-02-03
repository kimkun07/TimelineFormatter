class SheetInfo {
  /**
   * @param {SpreadsheetApp.Range} tableRange
   * @param {SpreadsheetApp.Range[]} projectGroups
   * @param {SpreadsheetApp.Range[]} monthGroups
   */
  constructor(tableRange, projectGroups, monthGroups) {
    this.tableRange = tableRange;
    this.projectGroups = projectGroups;
    this.monthGroups = monthGroups;
  }
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 * @returns {SheetInfo}
 */
function getSheetInfo(sheet) {
  // Get info from NamedRanges
  let projectRange = sheet.getNamedRanges().filter((value, _index, _array) => { return value.getName() == "ProjectStart" }); // NamedRange[]
  if (projectRange.length != 1) {
    throw ("There should be one \"ProjectStart\"");
  }
  let projectRow = projectRange[0].getRange().getRow();
  let projectCol = projectRange[0].getRange().getColumn();

  let monthRange = sheet.getNamedRanges().filter((value, _index, _array) => { return value.getName() == "MonthStart" }); // NamedRange[]
  if (monthRange.length != 1) {
    throw ("There should be one \"MonthStart\"");
  }
  let monthRow = monthRange[0].getRange().getRow();
  let monthCol = monthRange[0].getRange().getColumn();

  // Get table range
  let tableFirstRow = projectRow;
  let tableFirstCol = monthCol;
  let tableLastRow = sheet.getLastRow();
  let tableLastCol = sheet.getLastColumn();
  let tableRange = getRangeInclusive(sheet, tableFirstRow, tableFirstCol, tableLastRow, tableLastCol);

  // Get groups
  let projectGroups = [];
  for (let row = projectRow; row <= tableLastRow;) {
    let cell = getSingleMergedCell(sheet, row, projectCol);
    projectGroups.push(cell);
    row = cell.getLastRow() + 1;
  }

  let monthGroups = [];
  for (let col = monthCol; col <= tableLastCol;) {
    let cell = getSingleMergedCell(sheet, monthRow, col);
    monthGroups.push(cell);
    col = cell.getLastColumn() + 1;
  }

  return new SheetInfo(tableRange, projectGroups, monthGroups);
}

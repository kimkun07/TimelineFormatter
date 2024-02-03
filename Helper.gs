/**
 * @param {int} a
 * @param {int} b
 * @returns {int}
 */
function min(a, b) {
  return (a <= b) ? a : b;
}
/**
 * @param {int} a
 * @param {int} b
 * @returns {int}
 */
function max(a, b) {
  return (a >= b) ? a : b;
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {int} startRow
 * @param {int} startCol
 * @param {int} endRow
 * @param {int} endCol
 * @returns {SpreadsheetApp.Range}
 */
function getRangeInclusive(sheet, startRow, startCol, endRow, endCol) {
  return sheet.getRange(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
}

/**
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {int} row
 * @param {int} col
 * @returns {SpreadsheetApp.Range}
 */
function getSingleMergedCell(sheet, row, col) {
  let singleCell = sheet.getRange(row, col);
  if (!(singleCell.getMergedRanges().length <= 1)) {
    throw ("There are " + singleCell.getMergedRanges().length + " merged ranges for " + singleCell.getA1Notation());
  }
  let cell = (singleCell.getMergedRanges().length != 0) ? singleCell.getMergedRanges()[0] : singleCell;
  return cell;
}

// table_first, table_last: table range
// selection_first, selection_last: selection range

/**
 * @param {SpreadsheetApp.Range|null} selectionRange
 */
function format(selectionRange) {
  let sheet = SpreadsheetApp.getActiveSheet();
  const _sheetInfo = getSheetInfo(sheet);
  if (selectionRange == null) {
    formatTable(sheet, sheetInfo = _sheetInfo);
    selectionRange = _sheetInfo.tableRange;
  }
  formatCell(sheet, sheetInfo = _sheetInfo, selectionRange = selectionRange);
}

function formatAll() {
  format(null);
}

/** Format table in macro perspective
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {SheetInfo} sheetInfo
 */
function formatTable(sheet, sheetInfo) {
  const { tableRange, projectGroups, monthGroups } = sheetInfo;
  const tableFirstRow = tableRange.getRow();
  const tableFirstCol = tableRange.getColumn();
  const tableLastRow = tableRange.getLastRow();
  const tableLastCol = tableRange.getLastColumn();

  // Border
  // 1. Entire Cell: set outer border
  tableRange
    .setBorder(top = true, left = true, bottom = true, right = true, vertical = null, horizontal = null, "black", SpreadsheetApp.BorderStyle.SOLID);
  // 2. Entire Cell: set inner to dashed
  tableRange
    .setBorder(top = null, left = null, bottom = null, right = null, vertical = true, horizontal = true, "black", SpreadsheetApp.BorderStyle.DASHED);
  // 3. Draw row divider for each projecet
  for (let projectRange of projectGroups) {
    let color = projectRange.getFontColorObject().asRgbColor().asHexString();

    let firstRow = projectRange.getRow();
    let lastRow = projectRange.getLastRow();
    getRangeInclusive(sheet, firstRow, 1, lastRow, tableLastCol)
      .setBorder(top = null, left = null, bottom = true, right = null, vertical = null, horizontal = null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  // 4. Draw col divider for each month
  for (let monthRange of monthGroups) {
    let firstCol = monthRange.getColumn();
    let lastCol = monthRange.getLastColumn();
    getRangeInclusive(sheet, 1, firstCol, tableLastRow, lastCol)
      .setBorder(top = null, left = null, bottom = null, right = true, vertical = null, horizontal = null, "black", SpreadsheetApp.BorderStyle.SOLID);
  }
}

/** Format cell within selectionRange
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {SheetInfo} sheetInfo
 * @param {SpreadsheetApp.Range} selectionRange
 */
function formatCell(sheet, sheetInfo, selectionRange) {
  const { tableRange, projectGroups, monthGroups } = sheetInfo;
  const tableFirstRow = tableRange.getRow();
  const tableFirstCol = tableRange.getColumn();
  const tableLastRow = tableRange.getLastRow();
  const tableLastCol = tableRange.getLastColumn();

  const selectionFirstRow = max(selectionRange.getRow(), tableFirstRow);
  const selectionFirstCol = max(selectionRange.getColumn(), tableFirstCol);
  const selectionLastRow = min(selectionRange.getLastRow(), tableLastRow);
  const selectionLastCol = min(selectionRange.getLastColumn(), tableLastCol);

  if (!(selectionFirstCol <= selectionLastCol && selectionFirstRow <= selectionLastRow)) {
    return;
  }

  // For each project, format Empty Cell / Data Cell
  for (let projectRange of projectGroups) {
    let color = projectRange.getFontColorObject().asRgbColor().asHexString();
    let firstRow = max(projectRange.getRow(), selectionFirstRow);
    let lastRow = min(projectRange.getLastRow(), selectionLastRow);
    let firstCol = selectionFirstCol;
    let lastCol = selectionLastCol;

    if (!(firstRow <= lastRow)) {
      continue;
    }

    // Color, border for Empty Cells: white
    getRangeInclusive(sheet, firstRow, selectionFirstCol, lastRow, selectionLastCol)
      .setBackground("white");
    getRangeInclusive(sheet, firstRow, selectionFirstCol, lastRow, selectionLastCol)
      .setBorder(top = null, left = null, bottom = null, right = null, vertical = true, horizontal = true, "black", SpreadsheetApp.BorderStyle.DASHED);

    // Color, Border for Data Cells
    for (let row = firstRow; row <= lastRow; ++row) {
      for (let col = firstCol; col <= lastCol;) {
        let cell = getSingleMergedCell(sheet, row, col);

        if (cell.getValue() != "") {
          // Color
          cell.setBackground(color);

          // Border for right
          if (col + 1 <= tableLastCol && sheet.getRange(row, cell.getLastColumn() + 1).getValue() != "") {
            cell.setBorder(top = null, left = null, bottom = null, right = true, vertical = null, horizontal = null, "white", SpreadsheetApp.BorderStyle.DOUBLE);
          }
          // Border for top
          if (row - 1 >= tableFirstRow) {
            let targetRow = cell.offset(rowOffset = -1, columnOffset = 0);
            let flag = false;
            if (targetRow.getMergedRanges().length != 0) {
              // is merged => has value
              flag = true;
            } else {
              let valueArr = targetRow.getValues()[0];
              if (valueArr.some((value, index, array) => { return value != ""; })) {
                // found value
                flag = true;
              }
            }
            if (flag) {
              cell.setBorder(top = true, left = null, bottom = null, right = null, vertical = null, horizontal = null, "white", SpreadsheetApp.BorderStyle.DOUBLE);
            }
          }
          // Border for bottom
          if (row + 1 <= tableLastRow) {
            let targetRow = cell.offset(rowOffset = +1, columnOffset = 0);
            let flag = false;
            if (targetRow.getMergedRanges().length != 0) {
              // is merged => has value
              flag = true;
            } else {
              let valueArr = targetRow.getValues()[0];
              if (valueArr.some((value, index, array) => { return value != ""; })) {
                // found value
                flag = true;
              }
            }
            if (flag) {
              cell.setBorder(top = null, left = null, bottom = true, right = null, vertical = null, horizontal = null, "white", SpreadsheetApp.BorderStyle.DOUBLE);
            }
          }
        }

        if (col == lastCol) {
          break;
        }
        col = cell.getNextDataCell(direction = SpreadsheetApp.Direction.NEXT).getColumn();
      }
    }
  }
}

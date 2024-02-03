function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Timeline Formatter')
    .addItem('Format All', 'formatAll')
    .addToUi();
}

/**
 * @param {Event} event
 */
function onEdit(event) {
  format(event.range)
}

function applyCellStyle(range, style = {}) {
  const { bg, txt, weight, size, height } = style;

  if (bg) range.setBackground(bg);
  if (txt) range.setFontColor(txt);
  if (weight) range.setFontWeight(weight);
  if (size) range.setFontSize(size);

  range
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, CONFIG.THEME.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

  if (height) range.getSheet().setRowHeight(range.getRow(), height);
}

function renderTextBox(rangeNotation, message, themeItem) {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const range = sheet.getRange(rangeNotation);

  range.breakApart().clearContent();
  range.setDataValidation(null);

  range.merge()
    .setValue(message)
    .setFontFamily(CONFIG.THEME.FONT)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontColor(themeItem.txt)
    .setBackground(themeItem.bg)
    .setFontWeight('bold')
    .setBorder(true, true, true, true, null, null, themeItem.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function addCopyrightFooter() {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  sheet.getRange(CONFIG.RANGES.COPYRIGHT)
    .breakApart()
    .merge()
    .setValue(`© 2026 Developed by TierK • v${CONFIG.VERSION}`)
    .setFontSize(8)
    .setFontColor(CONFIG.THEME.COLOR_FOOTER)
    .setHorizontalAlignment('left');
}

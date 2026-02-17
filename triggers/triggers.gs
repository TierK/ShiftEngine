function onOpen() {
  SpreadsheetApp.getUi().createMenu(getMainMenuName())
    .addItem('נקה מערכת 🧹', 'clearSchedule')
    .addItem('נקה אילוצים שלח מיילים📧', 'sendWeeklyEmails')
    .addItem('עדכן מערכת 📅', 'updateScheduleDropdowns')
    .addItem('ייצוא לקובץ חדש 📤', 'exportSchedule')
    .addItem('הגדר תאריך התחלה 📅', 'setStartDate')
    .addToUi();
}

function installedOnEdit(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === CONFIG.SHEETS.RESPONSES && range.getRow() === 1) return;

  range.setFontFamily(CONFIG.THEME.FONT);

  if (sheetName === CONFIG.SHEETS.SCHEDULE) {
    syncActualShiftsAndStatus();
    addCopyrightFooter();
  }
}

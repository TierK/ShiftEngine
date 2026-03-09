function exportSchedule() {
  const startDateStr = getScriptProperty('START_DATE');
  if (!startDateStr) {
    SpreadsheetApp.getUi().alert("נא להגדיר תאריך התחלה (באמצעות 'נקה מערכת') לפני הייצוא");
    return;
  }

  const baseDate = DateService.parse(startDateStr);
  const fileName = getFormattedFileName(baseDate);
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);
  
  const titleW1 = schedSheet.getRange('B1').getValue() || 'שבוע 1';
  const titleW2 = schedSheet.getRange('B15').getValue() || 'שבוע 2';

  // Extract colors directly from the Responses sheet headers
  // Fallback to default pastel colors if the sheet isn't formatted yet
  let colorW1 = '#e7f3ff'; 
  let colorW2 = '#e6f4ea';
  if (respSheet && respSheet.getLastColumn() >= 16) {
    colorW1 = respSheet.getRange(1, 3).getBackground();
    colorW2 = respSheet.getRange(1, 10).getBackground();
  }

  const newSS = SpreadsheetApp.create(fileName);
  const newSheet = newSS.getSheets()[0];
  newSheet.setName(CONFIG.SHEETS.EXPORT_NAME).setRightToLeft(true).setHiddenGridlines(true);

  const targetFolderId = getOrCreateFolder(startDateStr);
  DriveApp.getFileById(newSS.getId()).moveTo(DriveApp.getFolderById(targetFolderId));

  buildWeekBlock(newSheet, schedSheet, {
    targetRow: 2,
    sourceRange: CONFIG.RANGES.GRID_W1.str,
    startDate: baseDate,
    title: titleW1,
    weekColor: colorW1 // Pass extracted color
  });

  const dateW2 = DateService.addDays(baseDate, 7);
  buildWeekBlock(newSheet, schedSheet, {
    targetRow: 11,
    sourceRange: CONFIG.RANGES.GRID_W2.str,
    startDate: dateW2,
    title: titleW2,
    weekColor: colorW2 // Pass extracted color
  });

  newSheet.setColumnWidth(1, 85);
  newSheet.setColumnWidths(2, 7, 110);

  buildAnalytics(newSS, baseDate);
  showSuccessDialog(newSS.getUrl());
}

function buildWeekBlock(targetSheet, sourceSheet, config) {
  const { targetRow, sourceRange, startDate, title, weekColor } = config;
  const data = sourceSheet.getRange(sourceRange).getValues();

  const rangeStr = DateService.getRangeStr(startDate).replace(' - ', ' עד ');
  const fullTitle = `${title.split(':')[0]} מ-${rangeStr}`;

  const titleRange = targetSheet.getRange(targetRow - 1, 1, 1, 8);
  titleRange.merge().setValue(fullTitle);
  applyCellStyle(titleRange, { bg: CONFIG.THEME.EXPORT_HEADER_BG, txt: CONFIG.THEME.EXPORT_HEADER_TXT, weight: 'bold', size: 12 });

  const dateRow = CONFIG.CONSTANTS.DAYS_HEADER.map((day, i) => {
    const d = DateService.addDays(startDate, i);
    return `${day}\n${d.getDate()}.${d.getMonth() + 1}`;
  });

  const headerRange = targetSheet.getRange(targetRow, 2, 1, 7);
  headerRange.setValues([dateRow]);
  applyCellStyle(headerRange, { bg: weekColor, weight: 'bold', height: 45 });

  CONFIG.CONSTANTS.SHIFT_TYPES.forEach((type, index) => {
    const rowOffset = targetRow + 1 + index * 2;
    const labelCell = targetSheet.getRange(rowOffset, 1, 2, 1);
    labelCell.merge().setValue(type);
    applyCellStyle(labelCell, { bg: CONFIG.THEME.EXPORT_SUBHEADER_BG, weight: 'bold' });

    const shiftRowRange = targetSheet.getRange(rowOffset, 1, 2, 8);
    shiftRowRange.setBackground(index % 2 === 0 ? '#ffffff' : '#f1f3f4');
  });

  const dataRange = targetSheet.getRange(targetRow + 1, 2, 6, 7);
  dataRange.setValues(data);
  applyCellStyle(dataRange);

  targetSheet.getRange(targetRow - 1, 1, 8, 8).setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  targetSheet.getRange(targetRow, 1, 7, 8).setBorder(null, null, null, null, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
}

function getFormattedFileName(startDate) {
  const pad = (n) => String(n).padStart(2, '0');
  const startDay = pad(startDate.getDate());
  const startMonth = pad(startDate.getMonth() + 1);
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 13);
  const endDay = pad(endDate.getDate());
  const endMonth = pad(endDate.getMonth() + 1);
  const endYear = endDate.getFullYear();
  return `שיבוץ: ${startDay}.${startMonth}-${endDay}.${endMonth}.${endYear}`;
}

function getOrCreateFolder(dateStr) {
  const rootName = 'שיבוץ מנהרת תשתיות';
  const parts = dateStr.split('.');
  const folderName = `${parts[1]}.${parts[2]}`;

  const root = DriveApp.getFoldersByName(rootName).hasNext()
    ? DriveApp.getFoldersByName(rootName).next()
    : DriveApp.createFolder(rootName);

  const sub = root.getFoldersByName(folderName).hasNext()
    ? root.getFoldersByName(folderName).next()
    : root.createFolder(folderName);

  return sub.getId();
}

function showSuccessDialog(url) {
  const html = HtmlService.createHtmlOutput(
    `<div dir="rtl" style="font-family: Arial; padding: 15px; text-align: center; color: #3c4043;">
      <h3 style="color: #1e8e3e; margin-top: 0;">הייצוא הושלם! ✨</h3>
      <p style="font-size: 14px; margin-bottom: 20px;">הקובץ מוכן ב-Drive שלך.</p>
      <a href="${url}" target="_blank" style="background: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">לפתיחת הקובץ 🚀</a>
    </div>`
  ).setWidth(300).setHeight(170);

  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function updateScheduleDateHeaders() {
  const startDateStr = getScriptProperty('START_DATE');
  if (!startDateStr) return;

  const baseDate = DateService.parse(startDateStr);
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);

  const titleW1 = `שבוע 1: ${DateService.getRangeStr(baseDate)}`;
  const dateW2 = new Date(baseDate);
  dateW2.setDate(baseDate.getDate() + 7);
  const titleW2 = `שבוע 2: ${DateService.getRangeStr(dateW2)}`;

  sheet.getRange(CONFIG.RANGES.HEADER_W1).setValue(titleW1);
  sheet.getRange(CONFIG.RANGES.HEADER_W2).setValue(titleW2);
}

function syncActualShiftsAndStatus() {
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  if (staffSheet.getLastRow() < 2) return;

  const w1Data = schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).getValues().flat();
  const w2Data = schedSheet.getRange(CONFIG.RANGES.GRID_W2.str).getValues().flat();

  const countShifts = (arr) => {
    const counts = {};
    arr.forEach((val) => {
      const name = val.toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, '').trim();
      if (name && name !== CONFIG.EMPTY_CELL) counts[name] = (counts[name] || 0) + 1;
    });
    return counts;
  };

  const c1 = countShifts(w1Data);
  const c2 = countShifts(w2Data);

  const staffRange = staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 7);
  const updated = staffRange.getValues().map((row) => {
    const name = row[CONFIG.STAFF.IDX_NAME].toString().trim();
    const target = row[CONFIG.STAFF.IDX_TARGET];

    row[CONFIG.STAFF.IDX_ACTUAL_W1] = c1[name] || 0;
    row[CONFIG.STAFF.IDX_STATUS_W1] = isShiftCountValid(target, row[CONFIG.STAFF.IDX_ACTUAL_W1]) ? 'OK' : 'Check';
    row[CONFIG.STAFF.IDX_ACTUAL_W2] = c2[name] || 0;
    row[CONFIG.STAFF.IDX_STATUS_W2] = isShiftCountValid(target, row[CONFIG.STAFF.IDX_ACTUAL_W2]) ? 'OK' : 'Check';
    return row;
  });

  staffRange.setValues(updated);
  applyGreenIfAllOk();
}

function cleanupDuplicateResponses() {
  const sheet = getSheet(CONFIG.SHEETS.RESPONSES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1).filter((row) => row[1] && row[1].toString().trim() !== '');
  if (rows.length === 0) return;

  rows.sort((a, b) => {
    const timeA = a[0] ? new Date(a[0]).getTime() : 0;
    const timeB = b[0] ? new Date(b[0]).getTime() : 0;
    return timeA - timeB;
  });

  const unique = {};
  rows.forEach((row) => {
    const name = row[1].toString().trim();
    unique[name] = row;
  });

  const finalData = Object.values(unique);
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.getRange(2, 1, maxRows - 1, sheet.getMaxColumns()).clearContent();
  }

  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, header.length)
      .setValues(finalData)
      .setFontFamily(CONFIG.THEME.FONT);
  }
}

function setStartDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'הגדרת סבב שיבוץ חדש',
    'נא להזין את תאריך יום ראשון הקרוב (תחילת הסבב):\nפורמט: DD.MM.YYYY (לדוגמה 22.02.2026)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const dateText = response.getResponseText().trim();
    if (dateText.split('.').length === 3) {
      setScriptProperty('START_DATE', dateText);
      updateScheduleDateHeaders();
      ui.alert('✅ התאריך עודכן והכותרות רועננו');
    } else {
      ui.alert('⚠️ פורמט לא תקין. נא להשתמש בנקודות: DD.MM.YYYY');
    }
  }
}

function clearSchedule() {
  setStartDate();
  setScriptProperty('emailsSent', 'false');

  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  [CONFIG.RANGES.GRID_W1.str, CONFIG.RANGES.GRID_W2.str].forEach((r) => {
    sheet.getRange(r)
      .setDataValidation(null)
      .clearContent()
      .setBackground(CONFIG.THEME.DEFAULT.bg)
      .setValue(CONFIG.EMPTY_CELL);
  });

  syncActualShiftsAndStatus();
}

/**
 * --- ShiftEngine  ---
 * Automated engine for bi-weekly labor scheduling.
 * Features: Algorithmic shift balancing, real-time conflict validation,
 * and automated export of schedules to a standalone spreadsheet
 * including generated charts and analytics.
 * * @version 0.14
 * @developer [TierK](https://github.com/TierK)
* @license MIT
 * Copyright (c) 2026 Kim
 */
/** --- GLOBAL CONFIGURATION --- */

const CONFIG = {
  VERSION: "0.14",
  SPREADSHEET_ID: '',
  FORM_URL: "",
  EMPTY_CELL: 'אין אילוץ',

  SHEETS: {
    SCHEDULE: "Schedule",
    STAFF: "Staff",
    RESPONSES: "Responses"
  },

  STAFF: {
    IDX_NAME: 0,
    IDX_EMAIL: 1,
    IDX_TARGET: 2,
    IDX_ACTUAL_W1: 3,
    IDX_STATUS_W1: 4,
    IDX_ACTUAL_W2: 5,
    IDX_STATUS_W2: 6
  },

  RANGES: {
    GRID_W1: { startRow: 3, startCol: 2, numRows: 6, numCols: 7, str: "B3:H8" },
    GRID_W2: { startRow: 17, startCol: 2, numRows: 6, numCols: 7, str: "B17:H22" },
    STATUS_W1: "B10:H10",
    STATUS_W2: "B24:H24",
    WARNING_W1: "B11:H11",
    WARNING_W2: "B25:H25",
    COPYRIGHT: "B30:H30",
    EXPORT_ZONE: "B27:H28"
  },

  SHIFT_TYPES: ["בוקר", "צהריים", "לילה"],
  
  DAYS_HEADER: ["יום ראשון", "יום שני", "יום שלישי", "יום רביעי", "יום חמישי", "יום שישי", "שבת"],

  THEME: {
    DEFAULT: { bg: "#f8f9fa", txt: "#5f6368", border: "#dadce0" },
    VALID:   { bg: "#e7f3ff", txt: "#1a73e8", border: "#aecbfa" },
    ERROR:   { bg: "#fce8e6", txt: "#d93025", border: "#fad2cf" },
    SUCCESS: { bg: "#e6f4ea", txt: "#1e8e3e", border: "#ceead6" },
    WARNING: { bg: "#fef7e0", txt: "#ea8600", border: "#feefc3" },
    INFO:    { bg: "#f1f3f4", txt: "#3c4043", border: "#bdc1c6" },
    FONT: "Varela Round",
    BORDER_COLOR: "#bdc1c6"
  }
};

const MESSAGES = {
  NEW_WEEK: "🧹 הטבלה נקיה. שלחו מיילים והמתינו שכולם יגישו אילוצים",
  WAITING_FOR_DATA: "📩 המיילים נשלחו. ממתינים לתגובות: ",
  DATA_READY: "✅ כל האילוצים התקבלו! לחצו על 'עדכן מערכת שעות' בתפריט 'פקודות✨'",
  NAME_ERROR: "⚠️ אופס! אחד השמות לא מזוהה. כדאי לבדוק בגיליון Staff ולתקן",
  FINISHED: "✅ כל הכבוד! הסידור מוכן וכולם מרוצים 🏆",
  IN_PROGRESS: "🛠️ השיבוץ מתקדם מצוין! וודאו שכל הסטטוסים בגיליון Staff הופכים ל-OK",
  DOUBLE_SHIFT: "⚠️ שים לב: עובד שובץ ליותר ממשמרת אחת באותו יום",
  NIGHT_MORNING: "⚠️ שים לב: עובד שובץ למשמרת בוקר מיד אחרי משמרת לילה",
  STATUS_OK: "OK",
  STATUS_CHECK: "Check"
};

const SHIFT_ROWS_MAP = { "בוקר": [0, 1], "צהריים": [2, 3], "לילה": [4, 5] };
const AUTO_MARKER = '\u200B'; 

/** --- CORE UTILITIES --- */

/**
 * Gets a sheet by name from the configured spreadsheet
 * @param {string} name - Sheet name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
const getSheet = (name) => SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(name);

/**
 * Validates if the assigned shifts match the staff target
 * @param {string|number} target - Expected shifts (can be a range like "3-4")
 * @param {number} actual - Current assigned shifts
 * @returns {boolean}
 */
const isShiftCountValid = (target, actual) => {
  if (target === undefined || target === null || target === "") return false;
  let targetStr = target.toString().trim();
  const currentActual = Number(actual);
  if (targetStr.includes('-')) {
    const parts = targetStr.split('-').map(Number);
    return currentActual >= parts[0] && currentActual <= parts[1];
  }
  return Number(targetStr) === currentActual;
};

/**
 * Shuffles an array in place
 * @param {Array} array
 * @returns {Array}
 */
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

/**
 * Identifies unique and duplicate employee names in responses
 * @returns {Object} {unique, duplicates}
 */
function getUniqueResponseNames() {
  const sheet = getSheet(CONFIG.SHEETS.RESPONSES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { unique: [], duplicates: [] };
  const namesRaw = sheet.getRange(2, 2, lastRow - 1, 1).getValues().map(r => r[0].toString().trim()).filter(String);
  const unique = [], duplicates = [], seen = {};
  namesRaw.forEach(name => {
    if (seen[name]) { if (!duplicates.includes(name)) duplicates.push(name); }
    else { seen[name] = true; unique.push(name); }
  });
  return { unique, duplicates };
}

/**
 * Universal styling helper for spreadsheet ranges
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {Object} style - Style properties (bg, txt, weight, size, height)
 */
function applyCellStyle(range, style = {}) {
  const { bg, txt, weight, size, height } = style;
  
  if (bg) range.setBackground(bg);
  if (txt) range.setFontColor(txt);
  if (weight) range.setFontWeight(weight);
  if (size) range.setFontSize(size);
  
  range.setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setBorder(true, true, true, true, true, true, CONFIG.THEME.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);

  if (height) {
    const sheet = range.getSheet();
    sheet.setRowHeight(range.getRow(), height);
  }
}

/** --- DATE MANAGEMENT --- */

/**
 * Prompts user to enter the start date for the new 14-day cycle
 */
function setStartDate() {
  const ui = SpreadsheetApp.getUi();
  // Improved UI message for clarity and Israeli date format preference
  const response = ui.prompt(
    'הגדרת סבב שיבוץ חדש', 
    'נא להזין את תאריך יום ראשון הקרוב (תחילת הסבב):\nפורמט: DD.MM.YYYY (לדוגמה 22.02.2026)', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const dateText = response.getResponseText().trim();
    // Supporting dot separator as requested
    const dateParts = dateText.split('.');
    
    if (dateParts.length === 3) {
      PropertiesService.getScriptProperties().setProperty('START_DATE', dateText);
      updateHeadersWithDates();
      ui.alert('✅ התאריך עודכן והכותרות רועננו');
    } else {
      ui.alert('⚠️ פורמט לא תקין. נא להשתמש בנקודות: DD.MM.YYYY');
    }
  }
}

/**
 * Recalculates and updates Hebrew headers on the Schedule sheet
 */
function updateHeadersWithDates() {
  const startDateStr = PropertiesService.getScriptProperties().getProperty('START_DATE');
  if (!startDateStr) return;

  const parts = startDateStr.split('.');
  // Note: parts[1]-1 because JS months are 0-indexed
  const date = new Date(parts[2], parts[1] - 1, parts[0]);
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const months = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];

  const formatDateRange = (d) => {
    let dEnd = new Date(d);
    dEnd.setDate(d.getDate() + 6);
    return `שבוע ${d.getDate()} - ${dEnd.getDate()} ל${months[d.getMonth()]} ${d.getFullYear()}`;
  };

  sheet.getRange("B1").setValue(formatDateRange(date));
  
  let dateW2 = new Date(date);
  dateW2.setDate(date.getDate() + 7);
  sheet.getRange("B15").setValue(formatDateRange(dateW2));
}

/** --- MAIN LOGIC & ASSIGNMENT --- */

/**
 * Main algorithm to assign shifts based on staff preferences and constraints
 */
/**
 * Updates the schedule dropdowns and performs auto-assignment.
 * Logic: Prioritizes shifts with fewer available candidates (Rare First) 
 * and balances based on employee targets.
 */
function updateScheduleDropdowns() {
  cleanupDuplicateResponses();
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);

  const staffData = staffSheet.getDataRange().getValues().slice(1);
  const staffTargets = {};
  staffData.forEach(row => { 
    staffTargets[row[CONFIG.STAFF.IDX_NAME].toString().trim()] = row[CONFIG.STAFF.IDX_TARGET]; 
  });

  const lastRow = respSheet.getLastRow();
  if (lastRow < 2) return;

  const responses = respSheet.getRange(2, 2, lastRow - 1, 15).getValues(); 
  let candidatesMap = {};

  responses.forEach(row => {
    const name = row[0].trim();
    if (!name) return;
    for (let i = 1; i <= 14; i++) {
      if (row[i]) row[i].split(",").forEach(s => {
        let key = `${i-1}-${s.trim()}`;
        if (!candidatesMap[key]) candidatesMap[key] = [];
        candidatesMap[key].push(name);
      });
    }
  });

  let nightWorkersGlobal = Array(14).fill().map(() => []);
  let dailyAssignments = Array.from({ length: 14 }, () => []);

  const processWeek = (weekOffset, config) => {
    let assignmentCount = {}; 
    responses.forEach(r => { if(r[0]) assignmentCount[r[0].trim()] = 0; });

    for (let dayLocal = 0; dayLocal < 7; dayLocal++) {
      const globalDayIndex = weekOffset + dayLocal;

      CONFIG.SHIFT_TYPES.forEach(type => {
        let pool = candidatesMap[`${globalDayIndex}-${type}`] || [];
        const localRows = SHIFT_ROWS_MAP[type];

        localRows.forEach(localRowIdx => {
          const actualRow = config.startRow + localRowIdx;
          const actualCol = config.startCol + dayLocal;
          const cell = schedSheet.getRange(actualRow, actualCol);

          let validCandidates = pool.filter(name => {
            if (dailyAssignments[globalDayIndex].includes(name)) return false;
            if (type === "בוקר" && globalDayIndex > 0) {
              if (nightWorkersGlobal[globalDayIndex - 1].includes(name)) return false;
            }
            return true;
          });

          validCandidates = shuffleArray(validCandidates).sort((a, b) => {
            const aTarget = staffTargets[a];
            const bTarget = staffTargets[b];

            const aDeficit = (typeof aTarget === 'string' && aTarget.includes('-')) 
                             ? Number(aTarget.split('-')[0]) - assignmentCount[a]
                             : Number(aTarget) - assignmentCount[a];
                             
            const bDeficit = (typeof bTarget === 'string' && bTarget.includes('-')) 
                             ? Number(bTarget.split('-')[0]) - assignmentCount[b]
                             : Number(bTarget) - assignmentCount[b];

            if (aDeficit !== bDeficit) return bDeficit - aDeficit;
            return assignmentCount[a] - assignmentCount[b];
          });

          if (validCandidates.length > 0) {
            const selected = validCandidates[0];
            cell.setValue(selected + AUTO_MARKER);
            if (pool.length > 1) {
              const rule = SpreadsheetApp.newDataValidation().requireValueInList(pool).build();
              cell.setDataValidation(rule);
            }
            assignmentCount[selected]++;
            dailyAssignments[globalDayIndex].push(selected);
            if (type === "לילה") nightWorkersGlobal[globalDayIndex].push(selected);
          } else {
            cell.setDataValidation(null).setValue(CONFIG.EMPTY_CELL);
          }
        });
      });
    }
  };

  schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).clearContent();
  schedSheet.getRange(CONFIG.RANGES.GRID_W2.str).clearContent();

  processWeek(0, CONFIG.RANGES.GRID_W1);
  processWeek(7, CONFIG.RANGES.GRID_W2);
  
  syncActualShiftsAndStatus();
}

/**
 * Helper to calculate how many shifts an employee still needs to reach their minimum target.
 */
function getDeficit(name, targets, counts) {
  const targetStr = targets[name] || "0";
  const minTarget = targetStr.toString().includes('-') 
                    ? Number(targetStr.split('-')[0]) 
                    : Number(targetStr);
  return minTarget - (counts[name] || 0);
}

function applyGreenIfAllOk() {
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const staffData = staffSheet.getLastRow() > 1 ? staffSheet.getDataRange().getValues().slice(1) : [];
  const staffNames = staffData.map(r => r[CONFIG.STAFF.IDX_NAME].toString().trim());

  const { unique: responseNames, duplicates } = getUniqueResponseNames();
  const emailsSent = PropertiesService.getScriptProperties().getProperty('emailsSent') === 'true';

  const w1Values = schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).getValues();
  const w1SatNight = [
    w1Values[4][6].toString().replace(AUTO_MARKER, "").trim(),
    w1Values[5][6].toString().replace(AUTO_MARKER, "").trim()
  ];

  const validateWeek = (config, isWeek2, statusRange, warningRange) => {
    const range = schedSheet.getRange(config.str);
    const gridValues = range.getValues();
    const validations = range.getDataValidations();
    
    let state = { isEmpty: true, hasNameError: false, hasDoubleShift: false, hasNightMorning: false, hasEmptyCells: false };
    let bgColors = Array.from({ length: 6 }, () => Array(7).fill(CONFIG.THEME.DEFAULT.bg));

    const statusColIdx = isWeek2 ? CONFIG.STAFF.IDX_STATUS_W2 : CONFIG.STAFF.IDX_STATUS_W1;
    const isWeekStaffStatusOk = staffData.every(r => r[statusColIdx] === "OK");

    for (let col = 0; col < 7; col++) {
      let namesThisDay = [];
      for (let row = 0; row < 6; row++) {
        let val = gridValues[row][col].toString();
        let cleanVal = val.replace(AUTO_MARKER, "").trim();

        if (!val || val.trim() === "" || cleanVal === CONFIG.EMPTY_CELL) {
          if (!val || val.trim() === "") state.hasEmptyCells = true;
        } else {
          state.isEmpty = false;
          if (!staffNames.includes(cleanVal)) { bgColors[row][col] = CONFIG.THEME.ERROR.bg; state.hasNameError = true; }
          else if (namesThisDay.includes(cleanVal)) { bgColors[row][col] = CONFIG.THEME.ERROR.bg; state.hasDoubleShift = true; }
          else if (col > 0 && (row === 0 || row === 1)) {
             const prevNight = [
               gridValues[4][col-1].toString().replace(AUTO_MARKER, "").trim(), 
               gridValues[5][col-1].toString().replace(AUTO_MARKER, "").trim()
             ];
             if (prevNight.includes(cleanVal)) { bgColors[row][col] = CONFIG.THEME.ERROR.bg; state.hasNightMorning = true; }
          }
          else if (isWeek2 && col === 0 && (row === 0 || row === 1)) {
             if (w1SatNight.includes(cleanVal)) { bgColors[row][col] = CONFIG.THEME.ERROR.bg; state.hasNightMorning = true; }
          }
          namesThisDay.push(cleanVal);
        }
      }
    }

    const shouldBeGreen = !state.isEmpty && isWeekStaffStatusOk && !state.hasEmptyCells && !state.hasNameError && !state.hasDoubleShift && !state.hasNightMorning;

    for (let col = 0; col < 7; col++) {
      for (let row = 0; row < 6; row++) {
        let val = gridValues[row][col].toString();
        let cleanVal = val.replace(AUTO_MARKER, "").trim();
        if (bgColors[row][col] !== CONFIG.THEME.ERROR.bg) {
          if (shouldBeGreen) bgColors[row][col] = CONFIG.THEME.SUCCESS.bg;
          else if (cleanVal === CONFIG.EMPTY_CELL || cleanVal === "") bgColors[row][col] = CONFIG.THEME.DEFAULT.bg;
          else {
            let isAmbiguous = (val.indexOf(AUTO_MARKER) !== -1 && validations[row][col]?.getCriteriaValues()[0]?.length > 2);
            bgColors[row][col] = isAmbiguous ? CONFIG.THEME.WARNING.bg : CONFIG.THEME.VALID.bg;
          }
        }
      }
    }
    range.setBackgrounds(bgColors);
    determineFinalStatus(statusRange, warningRange, state, responseNames.length, staffNames.length, emailsSent, duplicates, shouldBeGreen);
  };

  validateWeek(CONFIG.RANGES.GRID_W1, false, CONFIG.RANGES.STATUS_W1, CONFIG.RANGES.WARNING_W1);
  validateWeek(CONFIG.RANGES.GRID_W2, true, CONFIG.RANGES.STATUS_W2, CONFIG.RANGES.WARNING_W2);
}

function determineFinalStatus(statusRange, warningRange, state, count, total, emailsSent, duplicates, isWeekReady) {
  let mainMsg = "", mainTheme = CONFIG.THEME.INFO;
  const isAllIn = count >= total;

  if (isWeekReady) {
    mainMsg = MESSAGES.FINISHED;
    mainTheme = CONFIG.THEME.SUCCESS;
  }
  else if (state.isEmpty) {
    if (isAllIn && total > 0) { mainMsg = MESSAGES.DATA_READY; mainTheme = CONFIG.THEME.SUCCESS; }
    else if (emailsSent) { mainMsg = `${MESSAGES.WAITING_FOR_DATA} (${count}/${total})`; mainTheme = CONFIG.THEME.WARNING; }
    else { mainMsg = MESSAGES.NEW_WEEK; mainTheme = CONFIG.THEME.SUCCESS; }
  }
  else if (state.hasNameError) { mainMsg = MESSAGES.NAME_ERROR; mainTheme = CONFIG.THEME.ERROR; }
  else if (state.hasDoubleShift) { mainMsg = MESSAGES.DOUBLE_SHIFT; mainTheme = CONFIG.THEME.ERROR; }
  else if (state.hasNightMorning) { mainMsg = MESSAGES.NIGHT_MORNING; mainTheme = CONFIG.THEME.ERROR; }
  else { mainMsg = MESSAGES.IN_PROGRESS; mainTheme = CONFIG.THEME.INFO; }

  renderTextBox(statusRange, mainMsg, mainTheme);
  
  if (duplicates && duplicates.length) {
    renderTextBox(warningRange, `⚠️ כפילויות: ${duplicates.join(", ")}`, CONFIG.THEME.WARNING);
  } else {
    getSheet(CONFIG.SHEETS.SCHEDULE).getRange(warningRange).breakApart().clear().setBorder(false, false, false, false, false, false);
  }
}

/**
 * Syncs the actual shift counts to the Staff sheet
 */
function syncActualShiftsAndStatus() {
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  if (staffSheet.getLastRow() < 2) return;

  const w1Data = schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).getValues().flat();
  const w2Data = schedSheet.getRange(CONFIG.RANGES.GRID_W2.str).getValues().flat();

  const countShifts = (arr) => {
    let counts = {};
    arr.forEach(val => {
      const name = val.toString().replace(AUTO_MARKER, "").trim();
      if (name && name !== CONFIG.EMPTY_CELL) counts[name] = (counts[name] || 0) + 1;
    });
    return counts;
  };

  const c1 = countShifts(w1Data);
  const c2 = countShifts(w2Data);

  const staffRange = staffSheet.getRange(2, 1, staffSheet.getLastRow() - 1, 7);
  const updated = staffRange.getValues().map(row => {
    const name = row[CONFIG.STAFF.IDX_NAME].toString().trim();
    const target = row[CONFIG.STAFF.IDX_TARGET];
    
    row[CONFIG.STAFF.IDX_ACTUAL_W1] = c1[name] || 0;
    row[CONFIG.STAFF.IDX_STATUS_W1] = isShiftCountValid(target, row[CONFIG.STAFF.IDX_ACTUAL_W1]) ? "OK" : "Check";
    row[CONFIG.STAFF.IDX_ACTUAL_W2] = c2[name] || 0;
    row[CONFIG.STAFF.IDX_STATUS_W2] = isShiftCountValid(target, row[CONFIG.STAFF.IDX_ACTUAL_W2]) ? "OK" : "Check";
    return row;
  });

  staffRange.setValues(updated);
  applyGreenIfAllOk();
}

/** --- SYSTEM OPERATIONS --- */

/**
 * Removes duplicate form responses, keeping only the most recent one
 */
function cleanupDuplicateResponses() {
  const sheet = getSheet(CONFIG.SHEETS.RESPONSES);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const header = data[0];
  const rows = data.slice(1).filter(row => row[1] && row[1].toString().trim() !== "");
  rows.sort((a, b) => new Date(a[0]).getTime() - new Date(b[0]).getTime());
  const unique = {};
  rows.forEach(row => { unique[row[1].toString().trim()] = row; });
  sheet.clearContents().appendRow(header);
  const finalData = Object.values(unique);
  if (finalData.length > 0) sheet.getRange(2, 1, finalData.length, header.length).setValues(finalData);
}

/**
 * Clears the schedule grid and resets metadata
 */
function clearSchedule() {
  setStartDate(); // Prompt for new date immediately
  PropertiesService.getScriptProperties().setProperty('emailsSent', 'false');
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  [CONFIG.RANGES.GRID_W1.str, CONFIG.RANGES.GRID_W2.str].forEach(r => {
    sheet.getRange(r).setDataValidation(null).clearContent().setBackground(CONFIG.THEME.DEFAULT.bg).setValue(CONFIG.EMPTY_CELL);
  });
  syncActualShiftsAndStatus();
}

/**
 * Sends notification emails to staff and resets response collection
 */
function sendWeeklyEmails() {
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const staff = staffSheet.getDataRange().getValues();
  staff.slice(1).forEach(row => {
    const name = row[0], email = row[1];
    if (email) {
      try {
        MailApp.sendEmail({
          to: email,
          subject: "🗓️ שיבוץ עבודה: מילוי אילוצים לשבועיים הקרובים",
          htmlBody: `<div dir="rtl" style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.5; color: #3c4043;">היי <b>${name}</b>! 👋<br><br>הגיע הזמן למלא את משמרות <b>לשבועיים הקרובים</b>. ✨<br><br>נא למלא אילוצים בקישור הבא: 🚀<br><a href="${CONFIG.FORM_URL}" style="font-size: 18px; color: #1a73e8; font-weight: bold;">לחץ כאן למעבר לטופס האילוצים</a><br><br>תזכורת: נא לשלוח את האילוצים לשבועיים הקרובים <span style="text-decoration: underline; font-weight: bold;">לכל המאוחר עד יום שלישי בסוף היום</span>. ⏳<br><br>תודה מראש,<br><b>מרתה</b> 🌷</div>`
        });
      } catch (e) { console.error("Email error: " + email); }
    }
  });
  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);
  if (respSheet.getLastRow() > 1) respSheet.getRange(2, 1, respSheet.getLastRow() - 1, respSheet.getLastColumn()).clearContent();
  PropertiesService.getScriptProperties().setProperty('emailsSent', 'true');
  applyGreenIfAllOk();
  Browser.msgBox("✨ המיילים נשלחו!");
}

/**
 * Standard menu for the spreadsheet
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('פקודות✨')
    .addItem('נקה אילוצים שלח מיילים📧', 'sendWeeklyEmails')
    .addItem('עדכן מערכת 📅', 'updateScheduleDropdowns')
    .addItem('ייצוא לקובץ חדש 📤', 'exportSchedule') // It's back
    .addItem('נקה מערכת 🧹', 'clearSchedule')
    .addItem('הגדר תאריך התחלה 📅', 'setStartDate') // Added separate option for date management
    .addToUi();
}

/**
 * Triggered on cell edits
 */
function installedOnEdit(e) {
  if (e && e.range.getSheet().getName() === CONFIG.SHEETS.SCHEDULE) {
     syncActualShiftsAndStatus();
     addCopyrightFooter();
  }
}

/**
 * Adds a system footer with versioning
 */
function addCopyrightFooter() {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  sheet.getRange(CONFIG.RANGES.COPYRIGHT).breakApart().merge()
    .setValue(`© 2026 Developed by TierK • v${CONFIG.VERSION}`)
    .setFontSize(8).setFontColor("#d1d5db").setHorizontalAlignment("left");
}

/**
 * Renders a stylized message box in the spreadsheet
 */
function renderTextBox(rangeNotation, message, themeItem) {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  sheet.getRange(rangeNotation).breakApart().merge().setValue(message).setFontFamily(CONFIG.THEME.FONT).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setFontSize(10).setFontColor(themeItem.txt).setBackground(themeItem.bg).setFontWeight("bold")
    .setBorder(true, true, true, true, null, null, themeItem.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/** * --- ANALYTICS ENGINE v0.14 ---
 * Refactored to handle dynamic row positioning and invisible markers.
 */

// Configuration for colors to ensure table and charts match perfectly
const COLORS = {
  PIE_BG: '#D9EAD3',    // Light Green (Matches "Total" column)
  NIGHT_BG: '#E1D5E7',  // Light Purple (Matches "Nights" column)
  SHABBAT_BG: '#FFF2CC',// Light Yellow (Matches "Shabbat" column)
  TABLE_HEADER: '#4c11a1',
  TABLE_TEXT: '#ffffff',
/** --- DARK MODE ANALYTICS CONFIG --- */
  PAGE_BG: '#1c1c1c',     // Deep dark background for the whole sheet
  CHART_BG: '#2d2d2d',    // Slightly lighter dark for chart boxes
  TEXT_MAIN: '#e0e0e0',   // Off-white for readability
  TABLE_HEADER: '#4c11a1',
  TABLE_TEXT: '#ffffff',
  GRID_LINES: '#444444'   // Subtle grid for the table
};

const ANALYTICS_PALETTE = [
  '#FFD1DC', '#B3E5FC', '#C8E6C9', '#FFF9C4', '#FFCCBC', 
  '#D1C4E9', '#F0F4C3', '#FFE0B2', '#CFD8DC', '#E1BEE7',
  '#B2EBF2', '#F8BBD0', '#DCEDC8', '#FFECB3', '#D7CCC8',
  '#B2DFDB', '#E0E0E0', '#BBDEFB', '#FFF59D', '#F48FB1'
];

/**
 * Main function to export schedule to a new spreadsheet.
 * Coordinates data placement, folder organization, and analytics.
 */
function exportSchedule() {
  const scriptProps = PropertiesService.getScriptProperties();
  const startDateStr = scriptProps.getProperty('START_DATE');
  
  if (!startDateStr) {
    SpreadsheetApp.getUi().alert("נא להגדיר תאריך התחלה (באמצעות 'נקה מערכת') לפני הייצוא");
    return;
  }

  // 1. Setup Dates
  const parts = startDateStr.split('.');
  const baseDate = new Date(parts[2], parts[1] - 1, parts[0]);
  
  // 2. Generate Dynamic File Name
  const fileName = getFormattedFileName(baseDate);

  // 3. Setup Sheets and Content
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const titleW1 = schedSheet.getRange("B1").getValue() || "שבוע 1";
  const titleW2 = schedSheet.getRange("B15").getValue() || "שבוע 2";

  const newSS = SpreadsheetApp.create(fileName);
  const newSheet = newSS.getSheets()[0];
  newSheet.setName("שעות")
          .setRightToLeft(true)
          .setHiddenGridlines(true);

  // 4. Organize in Folders
  const targetFolderId = getOrCreateFolder(startDateStr);
  const file = DriveApp.getFileById(newSS.getId());
  const targetFolder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(targetFolder);

  // 5. Build Weeks
  buildWeekBlock(newSheet, schedSheet, {
    targetRow: 2,
    sourceRange: CONFIG.RANGES.GRID_W1.str,
    startDate: baseDate,
    title: titleW1
  });

  const dateW2 = new Date(baseDate);
  dateW2.setDate(baseDate.getDate() + 7);
  
  buildWeekBlock(newSheet, schedSheet, {
    targetRow: 11,
    sourceRange: CONFIG.RANGES.GRID_W2.str,
    startDate: dateW2,
    title: titleW2
  });

  // Final visual adjustments for the main sheet
  newSheet.setColumnWidth(1, 85);
  newSheet.setColumnWidths(2, 7, 110);

  // 6. Generate Analytics Module
  buildAnalytics(newSS, baseDate);

  showSuccessDialog(newSS.getUrl());
}

/**
 * Service: Formats filename as "שיבוץ: dd.mm-dd.mm.yyyy"
 */
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

/**
 * Orchestrates the creation of the Analytics dashboard.
 * Corrected: Added extra safety checks for data availability.
 */
function buildAnalytics(spreadsheet, baseDate) {
  const sourceSheet = spreadsheet.getSheetByName("שעות");
  if (!sourceSheet) return;

  // Ensure data is flushed to the sheet before reading for analytics
  SpreadsheetApp.flush();

  const analyticsSheet = spreadsheet.insertSheet("אנליטיקה");
  analyticsSheet.setRightToLeft(true).setHiddenGridlines(true);

  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const lastRow = staffSheet.getLastRow();
  const staffNames = lastRow > 1 
    ? staffSheet.getRange(2, CONFIG.STAFF.IDX_NAME + 1, lastRow - 1, 1).getValues().flat().filter(String)
    : [];

  const stats = processScheduleData(sourceSheet, staffNames);

  // Render Table
  renderAnalyticsTable(analyticsSheet, stats);

  // Layout Setup
  analyticsSheet.setColumnWidth(5, 30);      
  analyticsSheet.setColumnWidth(6, 60);      
  analyticsSheet.setRowHeights(1, 20, 25); 

  // Create Charts
  createVisualCharts(analyticsSheet, stats.length);
}

/**
 * Renders the stats table in Dark Mode.
 * Ensures high contrast between pastel name tags and the dark theme.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The analytics sheet.
 * @param {Array} stats - Array of employee statistics objects.
 */
function renderAnalyticsTable(sheet, stats) {
  // 1. Prepare the Dark Mode Canvas (extended to row 25)
  sheet.getRange(1, 1, 25, 10).setBackground(COLORS.PAGE_BG);
  
  // 2. Headers
  const headers = [["עובד", "סה\"כ", "לילות", "שבתות"]];
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setValues(headers)
             .setBackground(COLORS.TABLE_HEADER)
             .setFontColor(COLORS.TABLE_TEXT)
             .setFontWeight("bold")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");

  if (stats.length === 0) return;

  // 3. Data Styling
  const rows = stats.map(s => [s.name, s.total, s.nights, s.shabbat]);
  const dataRange = sheet.getRange(2, 1, stats.length, 4);
  dataRange.setValues(rows)
        .setFontColor(COLORS.TEXT_MAIN)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setFontFamily(CONFIG.THEME.FONT)
        .setBorder(true, true, true, true, true, true, COLORS.GRID_LINES, SpreadsheetApp.BorderStyle.SOLID);

  // 4. Column-specific Colors (Matching Charts)
  stats.forEach((s, i) => {
    const rowNum = i + 2;
    
    // Column 1: Name (Pastel)
    sheet.getRange(rowNum, 1).setBackground(s.color).setFontColor('#000000').setFontWeight("bold");
    
    // Column 2: Total (Match Pie BG)
    sheet.getRange(rowNum, 2).setBackground(COLORS.CHART_BG).setFontColor('#ffffff');
    
    // Column 3: Nights (Match Night Chart Bars)
    sheet.getRange(rowNum, 3).setBackground('#bb86fc').setFontColor('#000000');
    
    // Column 4: Shabbat (Match Shabbat Chart Bars)
    sheet.getRange(rowNum, 4).setBackground('#fbc02d').setFontColor('#000000');
  });

  sheet.setColumnWidths(1, 4, 95);
  sheet.setRowHeights(1, 23, 28); // Standardize row heights for the view
}
/**
 * Creates visual dashboard charts.
 * Pie (Green) -> Left Top
 * Night (Purple) -> Right Bottom (Under Table)
 * Shabbat (Yellow) -> Left Bottom (Under Pie)
 */
/**
 * Creates visual dashboard charts in Dark Mode.
 * Fixed the .setOptions error by using a consistent configuration approach.
 */
function createVisualCharts(sheet, rowCount) {
  const CHART_SIZE = {
    PIE: { WITH: 460, HEIGHT: 300 },
    GRAPH: { WITH: 430, HEIGHT: 220 }
  };

  // 1. Workload Distribution (Pie/Donut)
  const pieChart = sheet.newChart()
    .asPieChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 2))
    .setOption('title', 'חלוקת משמרות כללית')
    .setOption('pieHole', 0.4)
    .setOption('width', CHART_SIZE.PIE.WITH)
    .setOption('height', CHART_SIZE.PIE.HEIGHT)
    .setOption('backgroundColor', COLORS.CHART_BG)
    .setOption('colors', ANALYTICS_PALETTE.slice(0, rowCount))
    .setOption('titleTextStyle', { color: COLORS.TEXT_MAIN, fontSize: 14 })
    .setOption('legend', { textStyle: { color: COLORS.TEXT_MAIN } })
    .setPosition(1, 6, 0, 0)
    .build();

  // 2. Night Shift Analytics (Column)
  const nightChart = sheet.newChart()
    .asColumnChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 1)) // Names
    .addRange(sheet.getRange(1, 3, rowCount + 1, 1)) // Nights Data
    .setOption('title', 'ניתוח לילות')
    .setOption('width', CHART_SIZE.GRAPH.WITH)
    .setOption('height', CHART_SIZE.GRAPH.HEIGHT)
    .setOption('backgroundColor', COLORS.CHART_BG)
    .setOption('colors', ['#bb86fc']) // Electric purple
    .setOption('legend', { position: 'none' })
    .setOption('titleTextStyle', { color: COLORS.TEXT_MAIN })
    .setOption('hAxis', { textStyle: { color: COLORS.TEXT_MAIN } })
    .setOption('vAxis', { textStyle: { color: COLORS.TEXT_MAIN }, gridlines: { color: '#444' } })
    .setPosition(16, 1, 0, 0)
    .build();

  // 3. Shabbat Shift Analytics (Column)
  const shabbatChart = sheet.newChart()
    .asColumnChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 1)) // Names
    .addRange(sheet.getRange(1, 4, rowCount + 1, 1)) // Shabbat Data
    .setOption('title', 'ניתוח שבתות')
    .setOption('width', CHART_SIZE.GRAPH.WITH)
    .setOption('height', CHART_SIZE.GRAPH.HEIGHT)
    .setOption('backgroundColor', COLORS.CHART_BG)
    .setOption('colors', ['#fbc02d']) // Bright gold
    .setOption('legend', { position: 'none' })
    .setOption('titleTextStyle', { color: COLORS.TEXT_MAIN })
    .setOption('hAxis', { textStyle: { color: COLORS.TEXT_MAIN } })
    .setOption('vAxis', { textStyle: { color: COLORS.TEXT_MAIN }, gridlines: { color: '#444' } })
    .setPosition(16, 6, 0, 0)
    .build();

  sheet.insertChart(pieChart);
  sheet.insertChart(nightChart);
  sheet.insertChart(shabbatChart);
}

/**
 * Logic: Aggregates counts and assigns unique colors.
 * Optimized to detect shift types by label instead of hardcoded row numbers.
 */
function processScheduleData(sheet, names) {
  // Taking 20 rows to cover both weeks including headers
  const data = sheet.getRange(1, 1, 25, 8).getValues(); 
  const statsMap = {};
  
  // Initialize map with clean names
  names.forEach((name, index) => {
    const cleanName = name.trim();
    statsMap[cleanName] = { 
        total: 0, 
        nights: 0, 
        shabbat: 0, 
        color: ANALYTICS_PALETTE[index % ANALYTICS_PALETTE.length] 
    };
  });

  data.forEach((row) => {
    const shiftLabel = String(row[0]).trim(); // Column A: "בוקר", "צהריים", "לילה"
    const isNightShift = (shiftLabel === "לילה");
    
    // colIndex 1-7 are Sunday-Saturday
    row.forEach((cell, colIndex) => {
      if (colIndex === 0) return; // Skip labels column
      
      // Clean data from AUTO_MARKER and spaces
      const rawValue = String(cell).replace(/\u200B/g, "").trim();
      
      if (rawValue && statsMap[rawValue]) {
        statsMap[rawValue].total++;
        if (isNightShift) statsMap[rawValue].nights++;
        
        // Shabbat is the last column (index 7 in a 0-8 array)
        // Friday is index 6. Usually Shabbat shifts are Friday Night + Saturday
        if (colIndex === 7 || (colIndex === 6 && isNightShift)) {
          statsMap[rawValue].shabbat++;
        }
      }
    });
  });

  return names.map(name => {
    const n = name.trim();
    return {
      name: n, 
      total: statsMap[n].total, 
      nights: statsMap[n].nights, 
      shabbat: statsMap[n].shabbat, 
      color: statsMap[n].color
    };
  });
}

//////////////////////////////////////////////////////////////////////////////
/**
 * Service: Handles folder hierarchy on Google Drive
 */
function getOrCreateFolder(dateStr) {
  const rootName = "שיבוץ מנהרת תשתיות";
  const parts = dateStr.split('.');
  const folderName = `${parts[1]}.${parts[2]}`; // MM.YYYY

  let root = DriveApp.getFoldersByName(rootName).hasNext() 
    ? DriveApp.getFoldersByName(rootName).next() 
    : DriveApp.createFolder(rootName);

  let sub = root.getFoldersByName(folderName).hasNext()
    ? root.getFoldersByName(folderName).next()
    : root.createFolder(folderName);

  return sub.getId();
}

/**
 * Service: Builds a specific week block
 */
function buildWeekBlock(targetSheet, sourceSheet, config) {
  const { targetRow, sourceRange, startDate, title } = config;
  const data = sourceSheet.getRange(sourceRange).getValues();
  const weekColor = getRandomPastelHex(); 

  // 1. Header Title
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + 6);
  const dateRangeStr = `${startDate.getDate()}.${startDate.getMonth() + 1} עד ${endDate.getDate()}.${endDate.getMonth() + 1}`;
  const fullTitle = `${title} מ-${dateRangeStr}`;

  const titleRange = targetSheet.getRange(targetRow - 1, 1, 1, 8);
  titleRange.merge().setValue(fullTitle);
  applyCellStyle(titleRange, { bg: "#4c11a1", txt: "#ffffff", weight: "bold", size: 12 });

// 2. Dates Row
  const dateRow = CONFIG.DAYS_HEADER.map((day, i) => {
    let d = new Date(startDate);
    d.setDate(startDate.getDate() + i);
    return `${day}\n${d.getDate()}.${d.getMonth() + 1}`;
  });

  const headerRange = targetSheet.getRange(targetRow, 2, 1, 7);
  headerRange.setValues([dateRow]);
  applyCellStyle(headerRange, { bg: weekColor, weight: "bold", height: 45 });

  // 3. Shift Labels & Merging
  const shiftTypes = CONFIG.SHIFT_TYPES; // ["בוקר", "צהריים", "לילה"]
  shiftTypes.forEach((type, index) => {
    const rowOffset = targetRow + 1 + (index * 2);
    const labelCell = targetSheet.getRange(rowOffset, 1, 2, 1);
    labelCell.merge().setValue(type);
    applyCellStyle(labelCell, { bg: "#f8f9fa", weight: "bold" });
    
    const shiftRowRange = targetSheet.getRange(rowOffset, 1, 2, 8);
    if (index % 2 === 0) {
      shiftRowRange.setBackground("#ffffff");
    } else {
      shiftRowRange.setBackground("#f1f3f4");
    }
  });

  // 4. Data Grid
  const dataRange = targetSheet.getRange(targetRow + 1, 2, 6, 7);
  dataRange.setValues(data);
  applyCellStyle(dataRange);

  // 5. Borders
  const fullWeekRange = targetSheet.getRange(targetRow - 1, 1, 8, 8);
  // Set outside thick border
  fullWeekRange.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // Set inside thin grid lines
  const gridRange = targetSheet.getRange(targetRow, 1, 7, 8);
  gridRange.setBorder(null, null, null, null, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Service: UI Dialog
 */
function showSuccessDialog(url) {
  const html = HtmlService.createHtmlOutput(
    `<div dir="rtl" style="font-family: Arial; padding: 15px; text-align: center; color: #3c4043;">
      <h3 style="color: #1e8e3e; margin-top: 0;">הייצוא הושלם! ✨</h3>
      <p style="font-size: 14px; margin-bottom: 20px;">הקובץ מוכן ב-Drive שלך.</p>
      <a href="${url}" target="_blank" 
         style="background: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">
         לפתיחת הקובץ 🚀
      </a>
    </div>`
  ).setWidth(300).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Generates a vibrant pastel hex color, 
 * avoiding gray, black, and pure white.
 */
function getRandomPastelHex() {
// Hue: full circle
  const h = Math.floor(Math.random() * 360);
  // Saturation: 70-90% (avoids gray/faded colors)
  const s = Math.floor(Math.random() * 20) + 70;
  // Lightness: 75-85% (bright enough to be pastel, but not white)
  const l = Math.floor(Math.random() * 10) + 75;

  return hslToHex(h, s, l);
}

function hslToHex(h, s, l) {
  l /= 100;
  const a = s * Math.min(l, 1 - l) / 100;
  const f = n => {
    const k = (n + h / 30) % 12;
    const color = l - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);
    return Math.round(255 * color).toString(16).padStart(2, '0');
  };
  return `#${f(0)}${f(8)}${f(4)}`;
}
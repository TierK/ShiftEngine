/**
 * --- ShiftEngine v0.15.2 ---
 * Automated engine for bi-weekly labor scheduling.
 * * * New in v0.15.2:
 * - Refactored project structure: decoupled CONFIG from logic.
 * - Centralized DateService: synced date formatting between Schedule and Export.
 * - UI Automation: Dynamic random pastel coloring for Responses sheet (sync with Export style).
 * - Global Validation Cleanup: Automated removal of "Invalid" red flags.
 * - Code Standardization: Strict English documentation/comments.
 * Final UI adjustment: Fixed copyright position and resolved merging conflicts.
 * * @developer TierK
 * @license MIT
 * Copyright (c) 2026 Kim
 */

/** --- GLOBAL CONFIGURATION --- */
const CONFIG = {
  VERSION: "0.15.2",
  SPREADSHEET_ID: '',
  FORM_URL: "",
  EMPTY_CELL: 'אין אילוץ',

  SHEETS: {
    SCHEDULE: "Schedule",
    STAFF: "Staff",
    RESPONSES: "Responses",
    ANALYTICS: "אנליטיקה",
    EXPORT_NAME: "שעות"
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
    GRID_W1: { str: "B3:H8", startRow: 3, startCol: 2, numRows: 6, numCols: 7 },
    GRID_W2: { str: "B17:H22", startRow: 17, startCol: 2, numRows: 6, numCols: 7 },
    HEADER_W1: "A1:H1", 
    HEADER_W2: "A15:H15",
    STATUS_W1: "B10:H10",
    STATUS_W2: "B24:H24",
    WARNING_W1: "B11:H11",
    WARNING_W2: "B25:H25",
    COPYRIGHT: "A30:H30",
  },

  THEME: {
    FONT: "Varela Round",
    FONT_SIZE_MAIN: 10,
    FONT_SIZE_FOOTER: 8,
    COLOR_FOOTER: "#d1d5db",
    BORDER_COLOR: "#bdc1c6",
    
    // Status Colors
    DEFAULT: { bg: "#f8f9fa", txt: "#5f6368", border: "#dadce0" },
    VALID:   { bg: "#e7f3ff", txt: "#1a73e8", border: "#aecbfa" },
    ERROR:   { bg: "#fce8e6", txt: "#d93025", border: "#fad2cf" },
    SUCCESS: { bg: "#e6f4ea", txt: "#1e8e3e", border: "#ceead6" },
    WARNING: { bg: "#fef7e0", txt: "#ea8600", border: "#feefc3" },
    INFO:    { bg: "#f1f3f4", txt: "#3c4043", border: "#bdc1c6" },

    // Export & Analytics Theme
    EXPORT_HEADER_BG: "#4c11a1",
    EXPORT_HEADER_TXT: "#ffffff",
    EXPORT_SUBHEADER_BG: "#f8f9fa",
    ANALYTICS_DARK: {
      PAGE_BG: '#1c1c1c',
      CHART_BG: '#2d2d2d',
      TEXT_MAIN: '#e0e0e0',
      GRID_LINES: '#444444'
    }
  },

  ANALYTICS: {
    PALETTE: [
      '#FFD1DC', '#B3E5FC', '#C8E6C9', '#FFF9C4', '#FFCCBC', 
      '#D1C4E9', '#F0F4C3', '#FFE0B2', '#CFD8DC', '#E1BEE7',
      '#B2EBF2', '#F8BBD0', '#DCEDC8', '#FFECB3', '#D7CCC8',
      '#B2DFDB', '#E0E0E0', '#BBDEFB', '#FFF59D', '#F48FB1'
    ],
    COLOR_NIGHT: '#bb86fc',
    COLOR_SHABBAT: '#fbc02d',
    CHART_SIZE: {
      PIE: { WIDTH: 460, HEIGHT: 300 },
      GRAPH: { WIDTH: 430, HEIGHT: 220 }
    }
  },

  CONSTANTS: {
    MONTHS: ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'],
    DAYS_HEADER: ["יום ראשון", "יום שני", "יום שלישי", "יום רביעי", "יום חמישי", "יום שישי", "שבת"],
    SHIFT_TYPES: ["בוקר", "צהריים", "לילה"],
    SHIFT_ROWS_MAP: { "בוקר": [0, 1], "צהריים": [2, 3], "לילה": [4, 5] },
    AUTO_MARKER: '\u200B'
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

/** --- CORE TRIGGERS --- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('פקודות✨')
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
    
  // Global Font Sync
  e.range.setFontFamily(CONFIG.THEME.FONT);

  if (sheetName === CONFIG.SHEETS.SCHEDULE) {
    syncActualShiftsAndStatus();
    addCopyrightFooter();
  }
}

/** --- MAIN BUSINESS LOGIC --- */

/**
 * Main orchestrator: Syncs headers, cleans data, processes shifts, and formats UI.
 */
function updateScheduleDropdowns() {
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).clearContent().setDataValidation(null);
  schedSheet.getRange(CONFIG.RANGES.GRID_W2.str).clearContent().setDataValidation(null);
  updateScheduleDateHeaders();
  cleanupDuplicateResponses();
  
  performAutoAssignment(schedSheet);
  formatResponsesUI();
  syncActualShiftsAndStatus();
  addCopyrightFooter();
}

/**
 * Sends weekly emails and resets the responses sheet.
 */
function sendWeeklyEmails() {
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const staff = staffSheet.getDataRange().getValues();
  
  staff.slice(1).forEach(row => {
    const name = row[CONFIG.STAFF.IDX_NAME], email = row[CONFIG.STAFF.IDX_EMAIL];
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
  const maxRows = respSheet.getMaxRows();
  if (maxRows > 1) {
    respSheet.deleteRows(2, maxRows - 1);
  }
  PropertiesService.getScriptProperties().setProperty('emailsSent', 'true');
  applyGreenIfAllOk();
  Browser.msgBox("✨ המיילים נשלחו!");
}

/**
 * Core assignment logic wrapper.
 * Adapts existing assignment logic to the new CONFIG structure.
 */
function performAutoAssignment(schedSheet) {
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

  const processWeek = (weekOffset, rangeConfig) => {
    let assignmentCount = {};
    responses.forEach(r => { if(r[0]) assignmentCount[r[0].trim()] = 0; });

    for (let dayLocal = 0; dayLocal < 7; dayLocal++) {
      const globalDayIndex = weekOffset + dayLocal;
      CONFIG.CONSTANTS.SHIFT_TYPES.forEach(type => {
        let pool = candidatesMap[`${globalDayIndex}-${type}`] || [];
        const localRows = CONFIG.CONSTANTS.SHIFT_ROWS_MAP[type];

        localRows.forEach(localRowIdx => {
          const cell = schedSheet.getRange(rangeConfig.startRow + localRowIdx, rangeConfig.startCol + dayLocal);

          let validCandidates = pool.filter(name => {
             if (dailyAssignments[globalDayIndex].includes(name)) return false;
             if (type === "בוקר" && globalDayIndex > 0) {
               if (nightWorkersGlobal[globalDayIndex - 1].includes(name)) return false;
             }
             return true;
          });

          validCandidates = shuffleArray(validCandidates).sort((a, b) => {
            const aDeficit = getDeficit(a, staffTargets, assignmentCount);
            const bDeficit = getDeficit(b, staffTargets, assignmentCount);
            return bDeficit - aDeficit || assignmentCount[a] - assignmentCount[b];
          });

          if (validCandidates.length > 0) {
            const selected = validCandidates[0];
            cell.setValue(selected + CONFIG.CONSTANTS.AUTO_MARKER);
            
            // RESTORED DROPDOWN:
            if (pool.length > 0) {
              const rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(pool)
                .setAllowInvalid(true) // Allows manual override
                .build();
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

  schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).clearContent().setDataValidation(null);
  schedSheet.getRange(CONFIG.RANGES.GRID_W2.str).clearContent().setDataValidation(null);
  processWeek(0, CONFIG.RANGES.GRID_W1);
  processWeek(7, CONFIG.RANGES.GRID_W2);
}

/** --- VALIDATION & STATUS LOGIC --- */

function applyGreenIfAllOk() {
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const staffData = staffSheet.getLastRow() > 1 ? staffSheet.getDataRange().getValues().slice(1) : [];
  const staffNames = staffData.map(r => r[CONFIG.STAFF.IDX_NAME].toString().trim());
  const { unique: responseNames, duplicates } = getUniqueResponseNames();
  const emailsSent = PropertiesService.getScriptProperties().getProperty('emailsSent') === 'true';

  const w1Values = schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).getValues();
  const w1SatNight = [
    w1Values[4][6].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, "").trim(),
    w1Values[5][6].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, "").trim()
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
        let cleanVal = val.replace(CONFIG.CONSTANTS.AUTO_MARKER, "").trim();

        if (!val || val.trim() === "" || cleanVal === CONFIG.EMPTY_CELL) {
          state.hasEmptyCells = true; 
        } else {
          state.isEmpty = false;
          if (!staffNames.includes(cleanVal)) { bgColors[row][col] = CONFIG.THEME.ERROR.bg; state.hasNameError = true; }
          else if (namesThisDay.includes(cleanVal)) { bgColors[row][col] = CONFIG.THEME.ERROR.bg; state.hasDoubleShift = true; }
          else if (col > 0 && (row === 0 || row === 1)) {
             const prevNight = [
               gridValues[4][col-1].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, "").trim(), 
               gridValues[5][col-1].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, "").trim()
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
        let cleanVal = val.replace(CONFIG.CONSTANTS.AUTO_MARKER, "").trim();
        if (bgColors[row][col] !== CONFIG.THEME.ERROR.bg) {
          if (shouldBeGreen) bgColors[row][col] = CONFIG.THEME.SUCCESS.bg;
          else if (cleanVal === CONFIG.EMPTY_CELL || cleanVal === "") bgColors[row][col] = CONFIG.THEME.DEFAULT.bg;
          else {
            let isAmbiguous = (val.indexOf(CONFIG.CONSTANTS.AUTO_MARKER) !== -1 && validations[row][col]?.getCriteriaValues()[0]?.length > 2);
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

function syncActualShiftsAndStatus() {
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  if (staffSheet.getLastRow() < 2) return;

  const w1Data = schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).getValues().flat();
  const w2Data = schedSheet.getRange(CONFIG.RANGES.GRID_W2.str).getValues().flat();

  const countShifts = (arr) => {
    let counts = {};
    arr.forEach(val => {
      const name = val.toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, "").trim();
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

/** --- UI & FORMATTING SERVICES --- */

/**
 * Removes all red "Invalid" validation errors across all sheets.
 */
function clearAllValidationErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => {
    // Clears validation to remove red flags
    sheet.getDataRange().setDataValidation(null);
  });
}

/**
 * Formats the Responses sheet with dynamic random colors and global font.
 */
function formatResponsesUI() {
  const sheet = getSheet(CONFIG.SHEETS.RESPONSES);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (lastRow < 2) return;
  

  // 1. Reset and set global font
  const bodyRange = sheet.getRange(2, 1, maxRows - 1, sheet.getMaxColumns());
  
  bodyRange.setFontFamily(CONFIG.THEME.FONT)
           .setHorizontalAlignment("center")
           .setVerticalAlignment("middle")
           .setFontWeight("normal") 
           .setFontColor("#000000"); 

  // 2. Dynamic Header Coloring (Sync with Export style)
  if (lastCol >= 16) {
    sheet.getRange(1, 3, 1, 7).setBackground(getRandomPastelHex()).setFontWeight("bold");

    sheet.getRange(1, 10, 1, 7).setBackground(getRandomPastelHex()).setFontWeight("bold");
  }

  // 3. Zebra striping
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.setBackground("#ffffff"); // Чистим фон тела

  for (let i = 2; i <= lastRow; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, lastCol).setBackground("#f8f9fa");
    }
  }
}

/**
 * Updates Schedule headers using shared DateService logic.
 */
function updateScheduleDateHeaders() {
  const startDateStr = PropertiesService.getScriptProperties().getProperty('START_DATE');
  if (!startDateStr) return;

  const baseDate = DateService.parse(startDateStr);
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE)

  const titleW1 = `שבוע 1: ${DateService.getRangeStr(baseDate)}`;
  const dateW2 = new Date(baseDate);
  dateW2.setDate(baseDate.getDate() + 7);
  const titleW2 = `שבוע 2: ${DateService.getRangeStr(dateW2)}`;

  sheet.getRange(CONFIG.RANGES.HEADER_W1).setValue(titleW1);
  sheet.getRange(CONFIG.RANGES.HEADER_W2).setValue(titleW2);
}

/** --- UTILITY SERVICES --- */

/**
 * Shared Date operations to ensure consistency.
 */
const DateService = {
  parse: (str) => {
    const p = str.split('.');
    return new Date(p[2], p[1] - 1, p[0]);
  },
  getRangeStr: (d) => {
    let dEnd = new Date(d);
    dEnd.setDate(d.getDate() + 6);
    return `${d.getDate()}.${d.getMonth() + 1} - ${dEnd.getDate()}.${dEnd.getMonth() + 1}`;
  },
  addDays: (date, days) => {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  },
  format: (date) => {
    if (!date || !(date instanceof Date)) return "";
    return `${date.getDate()}.${date.getMonth() + 1}`;
  },
};

/**
 * Calculates shift deficit for assignment logic.
 */
function getDeficit(name, targets, counts) {
  const targetStr = targets[name] || "0";
  const minTarget = targetStr.toString().includes('-') 
                    ? Number(targetStr.split('-')[0]) 
                    : Number(targetStr);
  return minTarget - (counts[name] || 0);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function isShiftCountValid(target, actual) {
  if (target === undefined || target === null || target === "") return false;
  let targetStr = target.toString().trim();
  const currentActual = Number(actual);
  if (targetStr.includes('-')) {
    const parts = targetStr.split('-').map(Number);
    return currentActual >= parts[0] && currentActual <= parts[1];
  }
  return Number(targetStr) === currentActual;
}

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

function cleanupDuplicateResponses() {
  const sheet = getSheet(CONFIG.SHEETS.RESPONSES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getDataRange().getValues();
  const header = data[0];

  const rows = data.slice(1).filter(row => row[1] && row[1].toString().trim() !== "");
  
  if (rows.length === 0) return;

  rows.sort((a, b) => {
    const timeA = a[0] ? new Date(a[0]).getTime() : 0;
    const timeB = b[0] ? new Date(b[0]).getTime() : 0;
    return timeA - timeB;
  });

  const unique = {};
  rows.forEach(row => { 
    const name = row[1].toString().trim();
    unique[name] = row; 
  });

  const finalData = Object.values(unique);

  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, sheet.getMaxColumns()).clearContent();
  }

  sheet.getRange(2, 1, finalData.length, data[0].length)
        .setValues(finalData)
        .setFontFamily(CONFIG.THEME.FONT);

  if (finalData.length > 0) {
    sheet.getRange(2, 1, finalData.length, header.length) // используем header.length для точности
         .setValues(finalData)
         .setFontFamily(CONFIG.THEME.FONT);
  }
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function getRandomPastelHex() {
  const h = Math.floor(Math.random() * 360);
  const s = Math.floor(Math.random() * 20) + 70;
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

/** --- SYSTEM OPS: EXPORT & ANALYTICS --- */

function exportSchedule() {
  const scriptProps = PropertiesService.getScriptProperties();
  const startDateStr = scriptProps.getProperty('START_DATE');
  if (!startDateStr) {
    SpreadsheetApp.getUi().alert("נא להגדיר תאריך התחלה (באמצעות 'נקה מערכת') לפני הייצוא");
    return;
  }

  const baseDate = DateService.parse(startDateStr);
  const fileName = getFormattedFileName(baseDate);
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const titleW1 = schedSheet.getRange("B1").getValue() || "שבוע 1";
  const titleW2 = schedSheet.getRange("B15").getValue() || "שבוע 2";

  const newSS = SpreadsheetApp.create(fileName);
  const newSheet = newSS.getSheets()[0];
  newSheet.setName(CONFIG.SHEETS.EXPORT_NAME)
          .setRightToLeft(true)
          .setHiddenGridlines(true);

  const targetFolderId = getOrCreateFolder(startDateStr);
  DriveApp.getFileById(newSS.getId()).moveTo(DriveApp.getFolderById(targetFolderId));

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

  newSheet.setColumnWidth(1, 85);
  newSheet.setColumnWidths(2, 7, 110);

  buildAnalytics(newSS, baseDate);
  showSuccessDialog(newSS.getUrl());
}

function buildWeekBlock(targetSheet, sourceSheet, config) {
  const { targetRow, sourceRange, startDate, title } = config;
  const data = sourceSheet.getRange(sourceRange).getValues();
  const weekColor = getRandomPastelHex(); 

  // Header Title
  const rangeStr = DateService.getRangeStr(startDate).replace(' - ', ' עד '); // Adjusting format specifically for Export header
  const fullTitle = `${title.split(':')[0]} מ-${rangeStr}`;
  
  const titleRange = targetSheet.getRange(targetRow - 1, 1, 1, 8);
  titleRange.merge().setValue(fullTitle);
  applyCellStyle(titleRange, { bg: CONFIG.THEME.EXPORT_HEADER_BG, txt: CONFIG.THEME.EXPORT_HEADER_TXT, weight: "bold", size: 12 });

  // Dates Row
  const dateRow = CONFIG.CONSTANTS.DAYS_HEADER.map((day, i) => {
    let d = new Date(startDate);
    d.setDate(startDate.getDate() + i);
    return `${day}\n${d.getDate()}.${d.getMonth() + 1}`;
  });
  const headerRange = targetSheet.getRange(targetRow, 2, 1, 7);
  headerRange.setValues([dateRow]);
  applyCellStyle(headerRange, { bg: weekColor, weight: "bold", height: 45 });

  // Shift Labels
  CONFIG.CONSTANTS.SHIFT_TYPES.forEach((type, index) => {
    const rowOffset = targetRow + 1 + (index * 2);
    const labelCell = targetSheet.getRange(rowOffset, 1, 2, 1);
    labelCell.merge().setValue(type);
    applyCellStyle(labelCell, { bg: CONFIG.THEME.EXPORT_SUBHEADER_BG, weight: "bold" });
    
    const shiftRowRange = targetSheet.getRange(rowOffset, 1, 2, 8);
    shiftRowRange.setBackground(index % 2 === 0 ? "#ffffff" : "#f1f3f4");
  });

  // Data
  const dataRange = targetSheet.getRange(targetRow + 1, 2, 6, 7);
  dataRange.setValues(data);
  applyCellStyle(dataRange);

  // Borders
  targetSheet.getRange(targetRow - 1, 1, 8, 8).setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  targetSheet.getRange(targetRow, 1, 7, 8).setBorder(null, null, null, null, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
}

function buildAnalytics(spreadsheet, baseDate) {
  const sourceSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.EXPORT_NAME);
  if (!sourceSheet) return;
  SpreadsheetApp.flush();

  const analyticsSheet = spreadsheet.insertSheet(CONFIG.SHEETS.ANALYTICS);
  analyticsSheet.setRightToLeft(true).setHiddenGridlines(true);

  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const lastRow = staffSheet.getLastRow();
  const staffNames = lastRow > 1 
    ? staffSheet.getRange(2, CONFIG.STAFF.IDX_NAME + 1, lastRow - 1, 1).getValues().flat().filter(String)
    : [];

  const stats = processScheduleData(sourceSheet, staffNames);
  renderAnalyticsTable(analyticsSheet, stats);
  
  analyticsSheet.setColumnWidth(5, 30);      
  analyticsSheet.setColumnWidth(6, 60);      
  analyticsSheet.setRowHeights(1, 20, 25);
  createVisualCharts(analyticsSheet, stats.length);
}

function renderAnalyticsTable(sheet, stats) {
  const colors = CONFIG.THEME.ANALYTICS_DARK;
  sheet.getRange(1, 1, 25, 10).setBackground(colors.PAGE_BG);
  
  const headers = [["עובד", "סה\"כ", "לילות", "שבתות"]];
  sheet.getRange(1, 1, 1, 4).setValues(headers)
       .setBackground(CONFIG.THEME.EXPORT_HEADER_BG)
       .setFontColor(CONFIG.THEME.EXPORT_HEADER_TXT)
       .setFontWeight("bold")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");

  if (stats.length === 0) return;

  const rows = stats.map(s => [s.name, s.total, s.nights, s.shabbat]);
  sheet.getRange(2, 1, stats.length, 4).setValues(rows)
       .setFontColor(colors.TEXT_MAIN)
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle")
       .setFontFamily(CONFIG.THEME.FONT)
       .setBorder(true, true, true, true, true, true, colors.GRID_LINES, SpreadsheetApp.BorderStyle.SOLID);

  stats.forEach((s, i) => {
    const r = i + 2;
    sheet.getRange(r, 1).setBackground(s.color).setFontColor('#000000').setFontWeight("bold");
    sheet.getRange(r, 2).setBackground(colors.CHART_BG).setFontColor('#ffffff');
    sheet.getRange(r, 3).setBackground(CONFIG.ANALYTICS.COLOR_NIGHT).setFontColor('#000000');
    sheet.getRange(r, 4).setBackground(CONFIG.ANALYTICS.COLOR_SHABBAT).setFontColor('#000000');
  });

  sheet.setColumnWidths(1, 4, 95);
  sheet.setRowHeights(1, 23, 28);
}

function createVisualCharts(sheet, rowCount) {
  const colors = CONFIG.THEME.ANALYTICS_DARK;
  const size = CONFIG.ANALYTICS.CHART_SIZE;

  const pieChart = sheet.newChart().asPieChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 2))
    .setOption('title', 'חלוקת משמרות כללית')
    .setOption('pieHole', 0.4)
    .setOption('width', size.PIE.WIDTH).setOption('height', size.PIE.HEIGHT)
    .setOption('backgroundColor', colors.CHART_BG)
    .setOption('colors', CONFIG.ANALYTICS.PALETTE.slice(0, rowCount))
    .setOption('titleTextStyle', { color: colors.TEXT_MAIN, fontSize: 14 })
    .setOption('legend', { textStyle: { color: colors.TEXT_MAIN } })
    .setPosition(1, 6, 0, 0).build();

  const nightChart = sheet.newChart().asColumnChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 1))
    .addRange(sheet.getRange(1, 3, rowCount + 1, 1))
    .setOption('title', 'ניתוח לילות')
    .setOption('width', size.GRAPH.WIDTH).setOption('height', size.GRAPH.HEIGHT)
    .setOption('backgroundColor', colors.CHART_BG)
    .setOption('colors', [CONFIG.ANALYTICS.COLOR_NIGHT])
    .setOption('legend', { position: 'none' })
    .setOption('titleTextStyle', { color: colors.TEXT_MAIN })
    .setOption('hAxis', { textStyle: { color: colors.TEXT_MAIN } })
    .setOption('vAxis', { textStyle: { color: colors.TEXT_MAIN }, gridlines: { color: '#444' } })
    .setPosition(16, 1, 0, 0).build();

  const shabbatChart = sheet.newChart().asColumnChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 1))
    .addRange(sheet.getRange(1, 4, rowCount + 1, 1))
    .setOption('title', 'ניתוח שבתות')
    .setOption('width', size.GRAPH.WIDTH).setOption('height', size.GRAPH.HEIGHT)
    .setOption('backgroundColor', colors.CHART_BG)
    .setOption('colors', [CONFIG.ANALYTICS.COLOR_SHABBAT])
    .setOption('legend', { position: 'none' })
    .setOption('titleTextStyle', { color: colors.TEXT_MAIN })
    .setOption('hAxis', { textStyle: { color: colors.TEXT_MAIN } })
    .setOption('vAxis', { textStyle: { color: colors.TEXT_MAIN }, gridlines: { color: '#444' } })
    .setPosition(16, 6, 0, 0).build();

  sheet.insertChart(pieChart);
  sheet.insertChart(nightChart);
  sheet.insertChart(shabbatChart);
}

function processScheduleData(sheet, names) {
  const data = sheet.getRange(1, 1, 25, 8).getValues(); 
  const statsMap = {};shuffleArray
  names.forEach((name, index) => {
    const cleanName = name.trim();
    statsMap[cleanName] = { 
        total: 0, nights: 0, shabbat: 0, 
        color: CONFIG.ANALYTICS.PALETTE[index % CONFIG.ANALYTICS.PALETTE.length] 
    };
  });

  data.forEach((row) => {
    const shiftLabel = String(row[0]).trim();
    const isNightShift = (shiftLabel === "לילה");
    row.forEach((cell, colIndex) => {
      if (colIndex === 0) return;
      const rawValue = String(cell).replace(/\u200B/g, "").trim();
      if (rawValue && statsMap[rawValue]) {
        statsMap[rawValue].total++;
        if (isNightShift) statsMap[rawValue].nights++;
        if (colIndex === 7 || (colIndex === 6 && isNightShift)) {
          statsMap[rawValue].shabbat++;
        }
      }
    });
  });

  return names.map(name => {
    const n = name.trim();
    return { name: n, ...statsMap[n] };
  });
}

function getFormattedFileName(startDate) {
  const pad = (n) => String(n).padStart(2, '0');
  const startDay = pad(startDate.getDate()), startMonth = pad(startDate.getMonth() + 1);
  const endDate = new Date(startDate); endDate.setDate(startDate.getDate() + 13);
  const endDay = pad(endDate.getDate()), endMonth = pad(endDate.getMonth() + 1), endYear = endDate.getFullYear();
  return `שיבוץ: ${startDay}.${startMonth}-${endDay}.${endMonth}.${endYear}`;
}

function getOrCreateFolder(dateStr) {
  const rootName = "שיבוץ מנהרת תשתיות";
  const parts = dateStr.split('.');
  const folderName = `${parts[1]}.${parts[2]}`;
  let root = DriveApp.getFoldersByName(rootName).hasNext() ? DriveApp.getFoldersByName(rootName).next() : DriveApp.createFolder(rootName);
  let sub = root.getFoldersByName(folderName).hasNext() ? root.getFoldersByName(folderName).next() : root.createFolder(folderName);
  return sub.getId();
}

function applyCellStyle(range, style = {}) {
  const { bg, txt, weight, size, height } = style;
  if (bg) range.setBackground(bg);
  if (txt) range.setFontColor(txt);
  if (weight) range.setFontWeight(weight);
  if (size) range.setFontSize(size);
  range.setHorizontalAlignment("center").setVerticalAlignment("middle")
       .setBorder(true, true, true, true, true, true, CONFIG.THEME.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  if (height) range.getSheet().setRowHeight(range.getRow(), height);
}

function renderTextBox(rangeNotation, message, themeItem) {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const range = sheet.getRange(rangeNotation);
  
  range.breakApart().clearContent();
  range.setDataValidation(null); // Clear validation ONLY here
  
  range.merge().setValue(message)
    .setFontFamily(CONFIG.THEME.FONT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontColor(themeItem.txt)
    .setBackground(themeItem.bg)
    .setFontWeight("bold")
    .setBorder(true, true, true, true, null, null, themeItem.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
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

function addCopyrightFooter() {
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  sheet.getRange(CONFIG.RANGES.COPYRIGHT).breakApart().merge()
    .setValue(`© 2026 Developed by TierK • v${CONFIG.VERSION}`)
    .setFontSize(8).setFontColor(CONFIG.THEME.COLOR_FOOTER).setHorizontalAlignment("left");
}

function setStartDate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('הגדרת סבב שיבוץ חדש', 'נא להזין את תאריך יום ראשון הקרוב (תחילת הסבב):\nפורמט: DD.MM.YYYY (לדוגמה 22.02.2026)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const dateText = response.getResponseText().trim();
    if (dateText.split('.').length === 3) {
      PropertiesService.getScriptProperties().setProperty('START_DATE', dateText);
      updateScheduleDateHeaders();
      ui.alert('✅ התאריך עודכן והכותרות רועננו');
    } else {
      ui.alert('⚠️ פורמט לא תקין. נא להשתמש בנקודות: DD.MM.YYYY');
    }
  }
}

function clearSchedule() {
  setStartDate();
  PropertiesService.getScriptProperties().setProperty('emailsSent', 'false');
  const sheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  [CONFIG.RANGES.GRID_W1.str, CONFIG.RANGES.GRID_W2.str].forEach(r => {
    sheet.getRange(r).setDataValidation(null).clearContent().setBackground(CONFIG.THEME.DEFAULT.bg).setValue(CONFIG.EMPTY_CELL);
  });
  syncActualShiftsAndStatus();
}

/**
 * Optimized UI Cleanup. 
 * Only clears validation where it's NOT needed (Messages & Responses).
 */
/**
 * Optimized UI Cleanup. 
 * Only clears validation where it's NOT needed (Messages & Responses).
 */
function cleanupUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear only Responses (we don't need dropdowns there)
  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);
  if (respSheet) respSheet.getDataRange().setDataValidation(null);
  
  // Clear Message Areas in Schedule (so they don't have "Invalid" flags)
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  if (schedSheet) {
    [CONFIG.RANGES.STATUS_W1, CONFIG.RANGES.STATUS_W2, 
     CONFIG.RANGES.WARNING_W1, CONFIG.RANGES.WARNING_W2].forEach(r => {
       schedSheet.getRange(r).setDataValidation(null);
    });
  }
}

/**
 * Modified Assignment Logic to keep Dropdowns but hide Red Flags.
 */
function updateValidationWithSafeValue(cell, pool, selectedName) {
  if (!pool || pool.length === 0) {
    cell.setDataValidation(null);
    return;
  }
  
  // If manual edit or auto-assign puts a name not in pool, 
  // we add it to a temporary pool for this cell to hide the red arrow.
  let safePool = [...pool];
  if (selectedName && !safePool.includes(selectedName)) {
    safePool.push(selectedName);
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(safePool)
    .setAllowInvalid(true) // Key: allows manual entry without blocking
    .build();
  updateValidationWithSafeValue(cell, pool, selected);
}
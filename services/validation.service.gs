function applyGreenIfAllOk() {
  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const staffData = staffSheet.getLastRow() > 1 ? staffSheet.getDataRange().getValues().slice(1) : [];
  const staffNames = staffData.map((r) => r[CONFIG.STAFF.IDX_NAME].toString().trim());
  const { unique: responseNames, duplicates } = getUniqueResponseNames();
  const emailsSent = getScriptProperty('emailsSent') === 'true';

  const w1Values = schedSheet.getRange(CONFIG.RANGES.GRID_W1.str).getValues();
  const w1SatNight = [
    w1Values[4][6].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, '').trim(),
    w1Values[5][6].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, '').trim()
  ];

  const validateWeek = (config, isWeek2, statusRange, warningRange) => {
    const range = schedSheet.getRange(config.str);
    const gridValues = range.getValues();
    const validations = range.getDataValidations();

    const state = { isEmpty: true, hasNameError: false, hasDoubleShift: false, hasNightMorning: false, hasEmptyCells: false };
    const bgColors = Array.from({ length: 6 }, () => Array(7).fill(CONFIG.THEME.DEFAULT.bg));

    const statusColIdx = isWeek2 ? CONFIG.STAFF.IDX_STATUS_W2 : CONFIG.STAFF.IDX_STATUS_W1;
    const isWeekStaffStatusOk = staffData.every((r) => r[statusColIdx] === 'OK');

    for (let col = 0; col < 7; col++) {
      const namesThisDay = [];
      for (let row = 0; row < 6; row++) {
        const val = gridValues[row][col].toString();
        const cleanVal = val.replace(CONFIG.CONSTANTS.AUTO_MARKER, '').trim();

        if (!val || val.trim() === '' || cleanVal === CONFIG.EMPTY_CELL) {
          state.hasEmptyCells = true;
        } else {
          state.isEmpty = false;
          if (!staffNames.includes(cleanVal)) {
            bgColors[row][col] = CONFIG.THEME.ERROR.bg;
            state.hasNameError = true;
          } else if (namesThisDay.includes(cleanVal)) {
            bgColors[row][col] = CONFIG.THEME.ERROR.bg;
            state.hasDoubleShift = true;
          } else if (col > 0 && (row === 0 || row === 1)) {
            const prevNight = [
              gridValues[4][col - 1].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, '').trim(),
              gridValues[5][col - 1].toString().replace(CONFIG.CONSTANTS.AUTO_MARKER, '').trim()
            ];
            if (prevNight.includes(cleanVal)) {
              bgColors[row][col] = CONFIG.THEME.ERROR.bg;
              state.hasNightMorning = true;
            }
          } else if (isWeek2 && col === 0 && (row === 0 || row === 1) && w1SatNight.includes(cleanVal)) {
            bgColors[row][col] = CONFIG.THEME.ERROR.bg;
            state.hasNightMorning = true;
          }
          namesThisDay.push(cleanVal);
        }
      }
    }

    const shouldBeGreen = !state.isEmpty && isWeekStaffStatusOk && !state.hasEmptyCells && !state.hasNameError && !state.hasDoubleShift && !state.hasNightMorning;

    for (let col = 0; col < 7; col++) {
      for (let row = 0; row < 6; row++) {
        const val = gridValues[row][col].toString();
        const cleanVal = val.replace(CONFIG.CONSTANTS.AUTO_MARKER, '').trim();
        if (bgColors[row][col] === CONFIG.THEME.ERROR.bg) continue;

        if (shouldBeGreen) {
          bgColors[row][col] = CONFIG.THEME.SUCCESS.bg;
        } else if (cleanVal === CONFIG.EMPTY_CELL || cleanVal === '') {
          bgColors[row][col] = CONFIG.THEME.DEFAULT.bg;
        } else {
          const isAmbiguous = val.indexOf(CONFIG.CONSTANTS.AUTO_MARKER) !== -1 && validations[row][col] && validations[row][col].getCriteriaValues()[0] && validations[row][col].getCriteriaValues()[0].length > 2;
          bgColors[row][col] = isAmbiguous ? CONFIG.THEME.WARNING.bg : CONFIG.THEME.VALID.bg;
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
  let mainMsg = '';
  let mainTheme = CONFIG.THEME.INFO;
  const isAllIn = count >= total;

  if (isWeekReady) {
    mainMsg = MESSAGES.FINISHED;
    mainTheme = CONFIG.THEME.SUCCESS;
  } else if (state.isEmpty) {
    if (isAllIn && total > 0) {
      mainMsg = MESSAGES.DATA_READY;
      mainTheme = CONFIG.THEME.SUCCESS;
    } else if (emailsSent) {
      mainMsg = `${MESSAGES.WAITING_FOR_DATA} (${count}/${total})`;
      mainTheme = CONFIG.THEME.WARNING;
    } else {
      mainMsg = MESSAGES.NEW_WEEK;
      mainTheme = CONFIG.THEME.SUCCESS;
    }
  } else if (state.hasNameError) {
    mainMsg = MESSAGES.NAME_ERROR;
    mainTheme = CONFIG.THEME.ERROR;
  } else if (state.hasDoubleShift) {
    mainMsg = MESSAGES.DOUBLE_SHIFT;
    mainTheme = CONFIG.THEME.ERROR;
  } else if (state.hasNightMorning) {
    mainMsg = MESSAGES.NIGHT_MORNING;
    mainTheme = CONFIG.THEME.ERROR;
  } else {
    mainMsg = MESSAGES.IN_PROGRESS;
    mainTheme = CONFIG.THEME.INFO;
  }

  renderTextBox(statusRange, mainMsg, mainTheme);

  if (duplicates && duplicates.length) {
    renderTextBox(warningRange, `⚠️ כפילויות: ${duplicates.join(', ')}`, CONFIG.THEME.WARNING);
  } else {
    getSheet(CONFIG.SHEETS.SCHEDULE)
      .getRange(warningRange)
      .breakApart()
      .clear()
      .setBorder(false, false, false, false, false, false);
  }
}

function clearAllValidationErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach((sheet) => {
    sheet.getDataRange().setDataValidation(null);
  });
}

function cleanupUI() {
  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);
  if (respSheet) respSheet.getDataRange().setDataValidation(null);

  const schedSheet = getSheet(CONFIG.SHEETS.SCHEDULE);
  if (schedSheet) {
    [CONFIG.RANGES.STATUS_W1, CONFIG.RANGES.STATUS_W2, CONFIG.RANGES.WARNING_W1, CONFIG.RANGES.WARNING_W2].forEach((r) => {
      schedSheet.getRange(r).setDataValidation(null);
    });
  }
}

function updateValidationWithSafeValue(cell, pool, selectedName) {
  if (!pool || pool.length === 0) {
    cell.setDataValidation(null);
    return;
  }

  const safePool = [...pool];
  if (selectedName && !safePool.includes(selectedName)) {
    safePool.push(selectedName);
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(safePool)
    .setAllowInvalid(true)
    .build();

  cell.setDataValidation(rule);
}

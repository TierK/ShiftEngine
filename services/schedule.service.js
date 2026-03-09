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

function performAutoAssignment(schedSheet) {
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);

  const staffData = staffSheet.getDataRange().getValues().slice(1);
  const staffTargets = {};
  staffData.forEach((row) => {
    staffTargets[row[CONFIG.STAFF.IDX_NAME].toString().trim()] = row[CONFIG.STAFF.IDX_TARGET];
  });

  const lastRow = respSheet.getLastRow();
  if (lastRow < 2) return;

  const responses = respSheet.getRange(2, 2, lastRow - 1, 15).getValues();
  const candidatesMap = {};

  responses.forEach((row) => {
    const name = row[0].trim();
    if (!name) return;

    for (let i = 1; i <= 14; i++) {
      if (!row[i]) continue;
      row[i].split(',').forEach((shift) => {
        const key = `${i - 1}-${shift.trim()}`;
        if (!candidatesMap[key]) candidatesMap[key] = [];
        candidatesMap[key].push(name);
      });
    }
  });

  const nightWorkersGlobal = Array(14).fill().map(() => []);
  const dailyAssignments = Array.from({ length: 14 }, () => []);

  const processWeek = (weekOffset, rangeConfig) => {
    const assignmentCount = {};
    responses.forEach((r) => {
      if (r[0]) assignmentCount[r[0].trim()] = 0;
    });

    for (let dayLocal = 0; dayLocal < 7; dayLocal++) {
      const globalDayIndex = weekOffset + dayLocal;
      CONFIG.CONSTANTS.SHIFT_TYPES.forEach((type) => {
        const pool = candidatesMap[`${globalDayIndex}-${type}`] || [];
        const localRows = CONFIG.CONSTANTS.SHIFT_ROWS_MAP[type];

        localRows.forEach((localRowIdx) => {
          const cell = schedSheet.getRange(rangeConfig.startRow + localRowIdx, rangeConfig.startCol + dayLocal);

          let validCandidates = pool.filter((name) => {
            if (dailyAssignments[globalDayIndex].includes(name)) return false;
            if (type === 'בוקר' && globalDayIndex > 0 && nightWorkersGlobal[globalDayIndex - 1].includes(name)) return false;
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
            updateValidationWithSafeValue(cell, pool, selected);

            assignmentCount[selected]++;
            dailyAssignments[globalDayIndex].push(selected);
            if (type === 'לילה') nightWorkersGlobal[globalDayIndex].push(selected);
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

function getDeficit(name, targets, counts) {
  const targetStr = targets[name] || '0';
  const minTarget = targetStr.toString().includes('-') ? Number(targetStr.split('-')[0]) : Number(targetStr);
  return minTarget - (counts[name] || 0);
}

function getUniqueResponseNames() {
  const sheet = getSheet(CONFIG.SHEETS.RESPONSES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { unique: [], duplicates: [] };

  const namesRaw = sheet.getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .map((r) => r[0].toString().trim())
    .filter(String);

  const unique = [];
  const duplicates = [];
  const seen = {};

  namesRaw.forEach((name) => {
    if (seen[name]) {
      if (!duplicates.includes(name)) duplicates.push(name);
    } else {
      seen[name] = true;
      unique.push(name);
    }
  });

  return { unique, duplicates };
}

function formatResponsesUI() {
  const sheet = getSheet(CONFIG.SHEETS.RESPONSES);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  if (lastRow < 2) return;

  const bodyRange = sheet.getRange(2, 1, maxRows - 1, sheet.getMaxColumns());
  bodyRange
    .setFontFamily(CONFIG.THEME.FONT)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('normal')
    .setFontColor('#000000');

  if (lastCol >= 16) {
    sheet.getRange(1, 3, 1, 7).setBackground(getRandomPastelHex()).setFontWeight('bold');
    sheet.getRange(1, 10, 1, 7).setBackground(getRandomPastelHex()).setFontWeight('bold');
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.setBackground('#ffffff');

  for (let i = 2; i <= lastRow; i++) {
    if (i % 2 === 0) {
      sheet.getRange(i, 1, 1, lastCol).setBackground('#f8f9fa');
    }
  }
}

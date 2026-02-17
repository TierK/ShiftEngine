function buildAnalytics(spreadsheet) {
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

  const headers = [['עובד', 'סה"כ', 'לילות', 'שבתות']];
  sheet.getRange(1, 1, 1, 4)
    .setValues(headers)
    .setBackground(CONFIG.THEME.EXPORT_HEADER_BG)
    .setFontColor(CONFIG.THEME.EXPORT_HEADER_TXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (stats.length === 0) return;

  const rows = stats.map((s) => [s.name, s.total, s.nights, s.shabbat]);
  sheet.getRange(2, 1, stats.length, 4)
    .setValues(rows)
    .setFontColor(colors.TEXT_MAIN)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontFamily(CONFIG.THEME.FONT)
    .setBorder(true, true, true, true, true, true, colors.GRID_LINES, SpreadsheetApp.BorderStyle.SOLID);

  stats.forEach((s, i) => {
    const r = i + 2;
    sheet.getRange(r, 1).setBackground(s.color).setFontColor('#000000').setFontWeight('bold');
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
    .setOption('width', size.PIE.WIDTH)
    .setOption('height', size.PIE.HEIGHT)
    .setOption('backgroundColor', colors.CHART_BG)
    .setOption('colors', CONFIG.ANALYTICS.PALETTE.slice(0, rowCount))
    .setOption('titleTextStyle', { color: colors.TEXT_MAIN, fontSize: 14 })
    .setOption('legend', { textStyle: { color: colors.TEXT_MAIN } })
    .setPosition(1, 6, 0, 0)
    .build();

  const nightChart = sheet.newChart().asColumnChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 1))
    .addRange(sheet.getRange(1, 3, rowCount + 1, 1))
    .setOption('title', 'ניתוח לילות')
    .setOption('width', size.GRAPH.WIDTH)
    .setOption('height', size.GRAPH.HEIGHT)
    .setOption('backgroundColor', colors.CHART_BG)
    .setOption('colors', [CONFIG.ANALYTICS.COLOR_NIGHT])
    .setOption('legend', { position: 'none' })
    .setOption('titleTextStyle', { color: colors.TEXT_MAIN })
    .setOption('hAxis', { textStyle: { color: colors.TEXT_MAIN } })
    .setOption('vAxis', { textStyle: { color: colors.TEXT_MAIN }, gridlines: { color: '#444' } })
    .setPosition(16, 1, 0, 0)
    .build();

  const shabbatChart = sheet.newChart().asColumnChart()
    .addRange(sheet.getRange(1, 1, rowCount + 1, 1))
    .addRange(sheet.getRange(1, 4, rowCount + 1, 1))
    .setOption('title', 'ניתוח שבתות')
    .setOption('width', size.GRAPH.WIDTH)
    .setOption('height', size.GRAPH.HEIGHT)
    .setOption('backgroundColor', colors.CHART_BG)
    .setOption('colors', [CONFIG.ANALYTICS.COLOR_SHABBAT])
    .setOption('legend', { position: 'none' })
    .setOption('titleTextStyle', { color: colors.TEXT_MAIN })
    .setOption('hAxis', { textStyle: { color: colors.TEXT_MAIN } })
    .setOption('vAxis', { textStyle: { color: colors.TEXT_MAIN }, gridlines: { color: '#444' } })
    .setPosition(16, 6, 0, 0)
    .build();

  sheet.insertChart(pieChart);
  sheet.insertChart(nightChart);
  sheet.insertChart(shabbatChart);
}

function processScheduleData(sheet, names) {
  const data = sheet.getRange(1, 1, 25, 8).getValues();
  const statsMap = {};

  names.forEach((name, index) => {
    const cleanName = name.trim();
    statsMap[cleanName] = {
      total: 0,
      nights: 0,
      shabbat: 0,
      color: CONFIG.ANALYTICS.PALETTE[index % CONFIG.ANALYTICS.PALETTE.length]
    };
  });

  data.forEach((row) => {
    const shiftLabel = String(row[0]).trim();
    const isNightShift = shiftLabel === 'לילה';

    row.forEach((cell, colIndex) => {
      if (colIndex === 0) return;
      const rawValue = String(cell).replace(/\u200B/g, '').trim();
      if (!rawValue || !statsMap[rawValue]) return;

      statsMap[rawValue].total++;
      if (isNightShift) statsMap[rawValue].nights++;
      if (colIndex === 7 || (colIndex === 6 && isNightShift)) {
        statsMap[rawValue].shabbat++;
      }
    });
  });

  return names.map((name) => {
    const n = name.trim();
    return { name: n, ...statsMap[n] };
  });
}

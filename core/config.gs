const CONFIG = {
  VERSION: '0.16.1',
  SPREADSHEET_ID: '',
  FORM_URL: '',
  EMPTY_CELL: 'אין אילוץ',

  SHEETS: {
    SCHEDULE: 'Schedule',
    STAFF: 'Staff',
    RESPONSES: 'Responses',
    ANALYTICS: 'אנליטיקה',
    EXPORT_NAME: 'שעות'
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
    GRID_W1: { str: 'B3:H8', startRow: 3, startCol: 2, numRows: 6, numCols: 7 },
    GRID_W2: { str: 'B17:H22', startRow: 17, startCol: 2, numRows: 6, numCols: 7 },
    HEADER_W1: 'A1:H1',
    HEADER_W2: 'A15:H15',
    STATUS_W1: 'B10:H10',
    STATUS_W2: 'B24:H24',
    WARNING_W1: 'B11:H11',
    WARNING_W2: 'B25:H25',
    COPYRIGHT: 'A30:H30'
  },

  THEME: {
    FONT: 'Varela Round',
    FONT_SIZE_MAIN: 10,
    FONT_SIZE_FOOTER: 8,
    COLOR_FOOTER: '#d1d5db',
    BORDER_COLOR: '#bdc1c6',

    DEFAULT: { bg: '#f8f9fa', txt: '#5f6368', border: '#dadce0' },
    VALID: { bg: '#e7f3ff', txt: '#1a73e8', border: '#aecbfa' },
    ERROR: { bg: '#fce8e6', txt: '#d93025', border: '#fad2cf' },
    SUCCESS: { bg: '#e6f4ea', txt: '#1e8e3e', border: '#ceead6' },
    WARNING: { bg: '#fef7e0', txt: '#ea8600', border: '#feefc3' },
    INFO: { bg: '#f1f3f4', txt: '#3c4043', border: '#bdc1c6' },

    EXPORT_HEADER_BG: '#4c11a1',
    EXPORT_HEADER_TXT: '#ffffff',
    EXPORT_SUBHEADER_BG: '#f8f9fa',
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
    DAYS_HEADER: ['יום ראשון', 'יום שני', 'יום שלישי', 'יום רביעי', 'יום חמישי', 'יום שישי', 'שבת'],
    SHIFT_TYPES: ['בוקר', 'צהריים', 'לילה'],
    SHIFT_ROWS_MAP: { בוקר: [0, 1], צהריים: [2, 3], לילה: [4, 5] },
    AUTO_MARKER: '\u200B'
  }
};

const MESSAGES = {
  NEW_WEEK: '🧹 הטבלה נקיה. שלחו מיילים והמתינו שכולם יגישו אילוצים',
  WAITING_FOR_DATA: '📩 המיילים נשלחו. ממתינים לתגובות: ',
  DATA_READY: "✅ כל האילוצים התקבלו! לחצו על 'עדכן מערכת שעות' בתפריט 'פקודות✨'",
  NAME_ERROR: '⚠️ אופס! אחד השמות לא מזוהה. כדאי לבדוק בגיליון Staff ולתקן',
  FINISHED: '✅ כל הכבוד! הסידור מוכן וכולם מרוצים 🏆',
  IN_PROGRESS: '🛠️ השיבוץ מתקדם מצוין! וודאו שכל הסטטוסים בגיליון Staff הופכים ל-OK',
  DOUBLE_SHIFT: '⚠️ שים לב: עובד שובץ ליותר ממשמרת אחת באותו יום',
  NIGHT_MORNING: '⚠️ שים לב: עובד שובץ למשמרת בוקר מיד אחרי משמרת לילה',
  STATUS_OK: 'OK',
  STATUS_CHECK: 'Check'
};

const DateService = {
  parse: (str) => {
    const p = str.split('.');
    return new Date(p[2], p[1] - 1, p[0]);
  },

  getRangeStr: (d) => {
    const dEnd = new Date(d);
    dEnd.setDate(d.getDate() + 6);
    return `${d.getDate()}.${d.getMonth() + 1} - ${dEnd.getDate()}.${dEnd.getMonth() + 1}`;
  },

  addDays: (date, days) => {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  },

  format: (date) => {
    if (!date || !(date instanceof Date)) return '';
    return `${date.getDate()}.${date.getMonth() + 1}`;
  },
  
  getHebrewTwoWeeks: function(startDate) {
    const endDate = this.addDays(startDate, 13);
    const startMonth = CONFIG.CONSTANTS.MONTHS[startDate.getMonth()];
    const endMonth = CONFIG.CONSTANTS.MONTHS[endDate.getMonth()];
    
    return `מ${startDate.getDate()} ל${startMonth} עד ${endDate.getDate()} ל${endMonth}`;
  }
};

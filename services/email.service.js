function sendWeeklyEmails() {
  const startDateStr = getScriptProperty('START_DATE');
  const baseDate = DateService.parse(startDateStr);
  
  if (!baseDate || isNaN(baseDate.getTime())) {
    Browser.msgBox('⚠️ שגיאה: תאריך התחלה לא הוגדר או לא תקין. נא לבצע הגדרת סבב חדש.');
    return;
  }

  const two_weeks = DateService.getHebrewTwoWeeks(baseDate);

  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const staff = staffSheet.getDataRange().getValues();

  staff.slice(1).forEach((row) => {
    const name = row[CONFIG.STAFF.IDX_NAME];
    const email = row[CONFIG.STAFF.IDX_EMAIL];
    if (!email) return;

    try {
      MailApp.sendEmail({
        to: email,
        subject: `🗓 מילוי אילוצים לשיבוץ: ${two_weeks}`,
        htmlBody: `<div dir="rtl" style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.5; color: #3c4043;">
          היי <b>${name}</b>! 👋<br><br>
          נא למלא אילוצים לשבועיים <b>${two_weeks}</b>: 🚀<br>
          <a href="${CONFIG.MAIL_FORM_URL}" style="font-size: 18px; color: #1a73e8; font-weight: bold;">לחץ כאן למעבר לטופס האילוצים</a><br><br>
          תזכורת: נא לשלוח את האילוצים לשבועיים הקרובים <span style="text-decoration: underline; font-weight: bold;">לכל המאוחר עד יום שלישי בסוף היום</span>. ⏳<br><br>
          תודה מראש,<br>
          <b>מרתה</b> 🖤</div>`
      });
    } catch (error) {
      console.error(`Email error for ${email}: ${error}`);
    }
  });

  // Очистка старых ответов
  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);
  if (respSheet.getLastRow() > 1) {
    respSheet.deleteRows(2, respSheet.getLastRow() - 1);
  }

  setScriptProperty('emailsSent', 'true');
  applyGreenIfAllOk();
  Browser.msgBox('✨ המיילים נשלחו!');
}
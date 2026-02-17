function sendWeeklyEmails() {
  const staffSheet = getSheet(CONFIG.SHEETS.STAFF);
  const staff = staffSheet.getDataRange().getValues();

  staff.slice(1).forEach((row) => {
    const name = row[CONFIG.STAFF.IDX_NAME];
    const email = row[CONFIG.STAFF.IDX_EMAIL];
    if (!email) return;

    try {
      MailApp.sendEmail({
        to: email,
        subject: '🗓️ שיבוץ עבודה: מילוי אילוצים לשבועיים הקרובים',
        htmlBody: `<div dir="rtl" style="font-family: Arial, sans-serif; font-size: 16px; line-height: 1.5; color: #3c4043;">היי <b>${name}</b>! 👋<br><br>הגיע הזמן למלא את משמרות <b>לשבועיים הקרובים</b>. ✨<br><br>נא למלא אילוצים בקישור הבא: 🚀<br><a href="${CONFIG.FORM_URL}" style="font-size: 18px; color: #1a73e8; font-weight: bold;">לחץ כאן למעבר לטופס האילוצים</a><br><br>תזכורת: נא לשלוח את האילוצים לשבועיים הקרובים <span style="text-decoration: underline; font-weight: bold;">לכל המאוחר עד יום שלישי בסוף היום</span>. ⏳<br><br>תודה מראש,<br><b>מרתה</b> 🌷</div>`
      });
    } catch (error) {
      console.error(`Email error for ${email}: ${error}`);
    }
  });

  const respSheet = getSheet(CONFIG.SHEETS.RESPONSES);
  const maxRows = respSheet.getMaxRows();
  if (maxRows > 1) {
    respSheet.deleteRows(2, maxRows - 1);
  }

  setScriptProperty('emailsSent', 'true');
  applyGreenIfAllOk();
  Browser.msgBox('✨ המיילים נשלחו!');
}

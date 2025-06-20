function writeAiSummaryToColumnP(summaryText) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Summary");
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy"); // matches your sheet format
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0];
    const rowStr = rowDate instanceof Date
      ? Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "M/d/yyyy")
      : rowDate.toString().trim();

    if (rowStr === todayStr) {
      sheet.getRange(i + 1, 16).setValue(summaryText); // Column P = col 16
      Logger.log(`✅ AI summary written to row ${i + 1}, column P.`);
      return;
    }
  }

  Logger.log("⚠️ Could not find matching date in Daily Summary tab to write AI summary.");
}

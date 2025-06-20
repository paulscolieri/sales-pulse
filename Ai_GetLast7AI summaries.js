function getLast7AiSummaries() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Summary");
  const data = sheet.getDataRange().getValues();
  const summaries = [];

  for (let i = data.length - 1; i >= 1; i--) {
    const date = data[i][0];
    const summary = data[i][15]; // Column P (index 15)
    if (date && summary) {
      const formattedDate = date instanceof Date
        ? Utilities.formatDate(date, Session.getScriptTimeZone(), "M/d")
        : date.toString().trim();

      summaries.push(`- ${formattedDate}: ${summary}`);
    }
    if (summaries.length >= 7) break;
  }

  return summaries.length > 0
    ? `📅 *Previous 7 Days of AI Summaries:*\n${summaries.join("\n")}`
    : "";
}

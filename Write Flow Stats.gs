function writeFlowEmailStatsToSheet() {
  const sheetName = "Flow Email Stats";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  // Write headers if not already present
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date",
      "Flow ID",
      "Flow Message ID",
      "Recipients",
      "Opens",
      "Opens Unique",
      "Open Rate",
      "Clicks",
      "Clicks Unique",
      "Click Rate",
      "Click-to-Open Rate",
      "Conversions",
      "Conversion Uniques",
      "Conversion Rate",
      "Conversion Value",
      "Revenue per Recipient",
      "Unsubscribes",
      "Unsubscribe Rate",
      "Spam Complaints",
      "Spam Complaint Rate"
    ]);
  }

  const results = getYesterdayFlowEmailStats(); // pulls and returns array of stat rows
  if (!results || results.length === 0) return;

  const dateStr = new Date().toISOString().split("T")[0]; // today's date (for yesterday's stats)

  results.forEach(({ flow_id, flow_message_id, stats }) => {
    sheet.appendRow([
      dateStr,
      flow_id,
      flow_message_id,
      stats.recipients || 0,
      stats.opens || 0,
      stats.opens_unique || 0,
      stats.open_rate || 0,
      stats.clicks || 0,
      stats.clicks_unique || 0,
      stats.click_rate || 0,
      stats.click_to_open_rate || 0,
      stats.conversions || 0,
      stats.conversion_uniques || 0,
      stats.conversion_rate || 0,
      stats.conversion_value || 0,
      stats.revenue_per_recipient || 0,
      stats.unsubscribes || 0,
      stats.unsubscribe_rate || 0,
      stats.spam_complaints || 0,
      stats.spam_complaint_rate || 0
    ]);
  });
}



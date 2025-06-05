function writeFlowEmailStatsToSheet() {
  // Step 1: Refresh Flow Reference sheet
  writeKlaviyoFlowEmailDataToSheet();

  // Step 2: Set up target sheet
  const sheetName = "Flow Email Stats";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  // Write headers if not already present
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date",
      "Flow ID",
      "Flow Name",
      "Flow Message ID",
      "Message Name",
      "Subject Line",
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

  // Step 3: Load flow message metadata from Flow Reference tab
  const flowMessageMetaMap = buildFlowMessageReferenceMap();

  // Step 4: Fetch stats
  const results = getYesterdayFlowEmailStats();
  if (!results || results.length === 0) return;

  const dateStr = new Date().toISOString().split("T")[0];

  // Step 5: Write rows
  results.forEach(({ flow_id, flow_message_id, stats }) => {
    const meta = flowMessageMetaMap[flow_message_id] || {};

    sheet.appendRow([
      dateStr,
      flow_id,
      meta.flowName || "",
      flow_message_id,
      meta.messageName || "",
      meta.subjectLine || "",
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

  Logger.log("Flow Email Stats written successfully.");
}

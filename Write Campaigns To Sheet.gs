function writeCampaignReferenceToSheet() {
  
  const results = getYesterdayCampaignStats(); // or swap with getCampaignsLastXDays() for testing

  if (!results || results.length === 0) {
    Logger.log("No campaigns found.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Campaign Reference");
  if (!sheet) throw new Error("Campaign Reference sheet not found");

  results.forEach(row => {
    sheet.appendRow([
      row.campaign_name || "",
      row.campaign_id || "",
      row.subject_line || "",
      row.send_time || "",
      row.status || "",
      row.send_channel || ""
    ]);
  });

  Logger.log(`Wrote ${results.length} campaigns to Campaign Reference.`);
}

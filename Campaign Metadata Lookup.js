function backfillCampaignMetadataInEmailLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Email Log");
  const refSheet = ss.getSheetByName("Campaign Reference");

  if (!logSheet || !refSheet) throw new Error("Required sheet not found.");

  const logData = logSheet.getDataRange().getValues();
  const refData = refSheet.getDataRange().getValues();

  const refMap = {};
  const refHeaders = refData[0];

  for (let i = 1; i < refData.length; i++) {
    const row = refData[i];
    const campaignId = row[1];
    if (campaignId) {
      refMap[campaignId] = {
        name: row[0],
        subject: row[2],
        preview: row[3],
        sendTime: row[4]
      };
    }
  }

  const logHeaders = logData[0];
  const nameCol = logHeaders.indexOf("Campaign Name");
  const idCol = logHeaders.indexOf("Campaign ID");
  const subjectCol = logHeaders.indexOf("Subject Line");
  const previewCol = logHeaders.indexOf("Preview Text");
  const sendTimeCol = logHeaders.indexOf("Send Time");

  const updated = [];

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const campaignId = row[idCol];

    if (campaignId && refMap[campaignId]) {
      const meta = refMap[campaignId];
      if (!row[nameCol]) logSheet.getRange(i + 1, nameCol + 1).setValue(meta.name);
      if (!row[subjectCol]) logSheet.getRange(i + 1, subjectCol + 1).setValue(meta.subject);
      if (!row[previewCol]) logSheet.getRange(i + 1, previewCol + 1).setValue(meta.preview);
      if (!row[sendTimeCol]) logSheet.getRange(i + 1, sendTimeCol + 1).setValue(meta.sendTime);
      updated.push(campaignId);
    }
  }

  Logger.log(`Updated ${updated.length} Email Log rows with metadata.`);
}

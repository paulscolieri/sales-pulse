function addNewDateRowToDailySummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Summary");
  if (!sheet) {
    Logger.log("❌ Could not find the 'Daily Summary' sheet.");
    return;
  }

  // 🔎 Scan column A from bottom up for the last non-empty date
  const columnA = sheet.getRange("A:A").getValues();
  let lastDateRow = -1;
  for (let i = columnA.length - 1; i >= 0; i--) {
    if (columnA[i][0] instanceof Date || (typeof columnA[i][0] === "string" && columnA[i][0].trim() !== "")) {
      lastDateRow = i + 1; // 0-index to 1-index
      break;
    }
  }

  if (lastDateRow === -1) {
    Logger.log("❌ No date found in column A.");
    return;
  }

  const lastDateValue = sheet.getRange(lastDateRow, 1).getValue();
  const lastDate = lastDateValue instanceof Date ? lastDateValue : new Date(lastDateValue);

  if (isNaN(lastDate)) {
    Logger.log(`⚠️ Last date in column A is invalid: ${lastDateValue}`);
    return;
  }

  const nextDate = new Date(lastDate);
  nextDate.setDate(nextDate.getDate() + 1);

  sheet.insertRowAfter(lastDateRow);

  const newRow = lastDateRow + 1;
  sheet.getRange(newRow, 1).setValue(nextDate);

  const formulaRange = sheet.getRange(lastDateRow, 2, 1, 14);
  const targetRange = sheet.getRange(newRow, 2, 1, 14);
  formulaRange.copyTo(targetRange, { contentsOnly: false });

  Logger.log(`✅ Added new date row: ${nextDate.toLocaleDateString()} and copied formulas.`);
}
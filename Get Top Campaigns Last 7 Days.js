function getTopCampaignLast7Days(sheetName = "Email Log") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const header = data[0];
  const dateIdx = header.indexOf("Date");
  const campaignNameIdx = header.indexOf("Campaign Name");
  const revenueIdx = header.indexOf("Revenue");

  const revenueByCampaign = {};
  const today = new Date();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const campaign = row[campaignNameIdx] || "Unnamed Campaign";
    const revenue = parseFloat(row[revenueIdx] || 0);
    let rowDate;

    const rawDate = row[dateIdx];
    if (rawDate instanceof Date) {
      rowDate = new Date(rawDate);
    } else if (typeof rawDate === "string" && rawDate.includes("-")) {
      rowDate = new Date(rawDate);
    } else {
      continue;
    }

    const daysAgo = (today - rowDate) / (1000 * 60 * 60 * 24);
    if (daysAgo > 7) continue;

    revenueByCampaign[campaign] = (revenueByCampaign[campaign] || 0) + revenue;
  }

  const topEntry = Object.entries(revenueByCampaign).sort((a, b) => b[1] - a[1])[0];
  if (!topEntry) return null;

  return {
    campaign: topEntry[0],
    revenue: parseFloat(topEntry[1].toFixed(2))
  };
}

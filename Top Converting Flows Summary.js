function getTopConvertingFlowLast7Days(sheetName = "Flow Email Stats") {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  const today = new Date();
  const header = data[0];
  const dateIdx = header.indexOf("Date");
  const flowNameIdx = header.indexOf("Flow Name");
  const conversionsIdx = header.indexOf("Conversions");
  const conversionValueIdx = header.indexOf("Conversion Value");

  const conversionsByFlow = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawDate = row[dateIdx];
    const flowName = row[flowNameIdx] || "Unknown Flow";
    const conversions = parseInt(row[conversionsIdx] || 0);
    const conversionValue = parseFloat(row[conversionValueIdx] || 0);

    let rowDate;
    if (rawDate instanceof Date) {
      rowDate = new Date(rawDate);
    } else if (typeof rawDate === "string" && rawDate.includes("/")) {
      const [mm, dd, yyyy] = rawDate.split("/").map(s => parseInt(s, 10));
      rowDate = new Date(yyyy, mm - 1, dd);
    } else if (typeof rawDate === "string" && rawDate.includes("-")) {
      rowDate = new Date(rawDate);
    } else {
      continue; // invalid date
    }

    const daysAgo = (today - rowDate) / (1000 * 60 * 60 * 24);
    if (daysAgo > 7) continue;

    if (!conversionsByFlow[flowName]) {
      conversionsByFlow[flowName] = { conversions: 0, conversionValue: 0 };
    }

    conversionsByFlow[flowName].conversions += conversions;
    conversionsByFlow[flowName].conversionValue += conversionValue;
  }

  const topEntry = Object.entries(conversionsByFlow)
    .sort((a, b) => b[1].conversions - a[1].conversions)[0];

  if (!topEntry) {
    Logger.log("📭 No flow conversions found in the last 7 days.");
    return null;
  }

  const [flowName, metrics] = topEntry;
  const result = {
    flowName,
    conversions: metrics.conversions,
    conversionValue: parseFloat(metrics.conversionValue.toFixed(2))
  };

  Logger.log("📊 Top converting flow in last 7 days:");
  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

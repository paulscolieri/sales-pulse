function buildFlowMessageReferenceMap() {
  const map = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Flow Reference");
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const messageIdIndex = headers.indexOf("Message ID");
  const flowNameIndex = headers.indexOf("Flow Name");
  const messageNameIndex = headers.indexOf("Message Name");
  const subjectLineIndex = headers.indexOf("Subject Line");

  rows.forEach(row => {
    const messageId = row[messageIdIndex];
    if (messageId) {
      map[messageId] = {
        flowName: row[flowNameIndex] || "",
        messageName: row[messageNameIndex] || "",
        subjectLine: row[subjectLineIndex] || ""
      };
    }
  });

  return map;
}

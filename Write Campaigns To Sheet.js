function writeKlaviyoCampaignReferenceToSheet() {
  const sheetName = "Campaign Reference";
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty("KLAVIYO_API_KEY");
  const revision = "2025-04-15";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // Add headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Campaign Name",
      "Campaign ID",
      "Subject Line",
      "Preview Text",
      "Send Time",
      "Status",
      "Send Channel",
      "Included Audience IDs"
    ]);
  }

  // 🧠 Build list of existing Campaign IDs to avoid duplicates
  const existingData = sheet.getDataRange().getValues();
  const existingIds = new Set(existingData.slice(1).map(row => row[1])); // column B = Campaign ID

  // ⏱ Build 30-day created_at filter
  const today = new Date();
  const start = new Date(today);
  start.setDate(start.getDate() - 30);
  start.setHours(0, 0, 0, 0);
  const startISO = start.toISOString();

  const rawFilter = `equals(messages.channel,"email"),greater-or-equal(created_at,${startISO})`;
  const filter = encodeURIComponent(rawFilter);
  let nextPage = `https://a.klaviyo.com/api/campaigns?filter=${filter}&include=campaign-messages`;

  // 🧾 Fetch all paginated results
  while (nextPage) {
    const res = UrlFetchApp.fetch(nextPage, {
      method: "get",
      headers: {
        "Authorization": `Klaviyo-API-Key ${apiKey}`,
        "Accept": "application/json",
        "revision": revision
      }
    });

    const data = JSON.parse(res.getContentText());
    const campaigns = data.data || [];

    // 📬 Map campaign-message.id → subject/preview
    const campaignMessages = data.included?.filter(item => item.type === "campaign-message") || [];
    const messageMap = {};
    campaignMessages.forEach(msg => {
      const content = msg.attributes?.definition?.content || {};
      messageMap[msg.id] = {
        subject: content.subject || "",
        preview_text: content.preview_text || ""
      };
    });

    // 🖊 Write new campaign rows
    campaigns.forEach(campaign => {
      const attrs = campaign.attributes || {};
      const id = campaign.id;

      if (!existingIds.has(id)) {
        const msgId = campaign.relationships?.["campaign-messages"]?.data?.[0]?.id;
        const messageMeta = messageMap[msgId] || {};
        const includedAudiences = (attrs.audiences?.included || []).join(", ");

        sheet.appendRow([
          attrs.name || "",
          id,
          messageMeta.subject || "",
          messageMeta.preview_text || "",
          attrs.send_time || "",
          attrs.status || "",
          attrs.send_channel || "",
          includedAudiences
        ]);

        existingIds.add(id); // avoid re-adding
      }
    });

    nextPage = data.links?.next || null;
  }

  Logger.log("✅ Campaign Reference updated with new campaigns.");
}

function generateAndSendDailySummary() {
  const shopifySummary = writeLineItemsShopifyStyle();
  const flowSummary = writeFlowEmailStatsToSheet();
  const campaignSummary = writeCampaignDailyStatsToLog();
  const adsSummary = writeMetaAdInsightsToSheet();

  const report = {
    date: new Date().toISOString().split("T")[0],
    shopify: shopifySummary,
    klaviyo_flows: flowSummary,
    klaviyo_campaigns: campaignSummary,
    meta_ads: adsSummary
  };

  const prompt = `
You're an ecommerce performance analyst. Here's a report of yesterday's performance data from multiple sources. Write a short, insightful executive summary (under 120 words) that highlights:
- Key wins and underperformance
- Notable trends or top-performing campaigns/products
- ROAS context for ads
- Anything worth flagging for tomorrow

Be specific. Make it clear and concise. Here's the data:
${JSON.stringify(report, null, 2)}
`;

  const payload = {
    model: "gpt-4o",
    messages: [
      { role: "system", content: "You are an ecommerce performance analyst. You summarize data for execs." },
      { role: "user", content: prompt }
    ]
  };

  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload)
  });

  const gptResponse = JSON.parse(response.getContentText());
  const summary = gptResponse.choices[0].message.content;

  Logger.log("🧠 Daily AI Summary:");
  Logger.log(summary);

  // Optional: send via email
  // MailApp.sendEmail({
  //   to: "you@example.com",
  //   subject: `📈 Daily Exec Summary – ${report.date}`,
  //   htmlBody: `<pre>${summary}</pre>`
  // });
  Logger.log("🧠 Daily AI Summary:");
  Logger.log(summary);

  return summary;
}

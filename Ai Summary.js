function generateAndSendDailySummary() {
  // Step 1: Pull summaries
  const shopifySummary = getShopifySummaryFromSheet();
  const flowSummary = getFlowSummaryFromSheet();
  const campaignSummary = getCampaignSummaryFromSheet();
  const adsSummary = getMetaAdSummaryFromSheet();
  const context = getAiContextFromSheet();

  // Step 2: Combine report
  const report = {
    date: new Date().toISOString().split("T")[0],
    shopify: shopifySummary,
    klaviyo_flows: flowSummary,
    klaviyo_campaigns: campaignSummary,
    meta_ads: adsSummary
  };

  // ✅ Add this line here
  const trailingSummaryContext = getLast7AiSummaries();

  // Step 3: Construct GPT prompt using the 5-question structure
  const prompt = `
You're an ecommerce performance analyst. Write a short, clear executive summary (under 150 words) of yesterday’s performance for the brand "${context["Brand Name"] || "our brand"}".

Tone: ${context["Brand Voice"] || "concise and direct"}  
Business Model: ${context["Business Model"] || "ecommerce"}  
Key Products: ${context["Key Products"] || "N/A"}  
Focus this week: ${context["Strategic Focus This Week"] || "N/A"}  
Current promo: ${context["Current Offers/Promos"] || "None"}  
Known external factors: ${context["Known External Factors"] || "None"}  

Use the following structure to answer these 5 exec-level questions:

1️⃣ **Did we make money yesterday?**  
- Report daily revenue, profit, and margin. Compare to goals if available.

2️⃣ **What worked or didn’t?**  
- Highlight standout flows, campaigns, ads, or products. Flag underperformers or anomalies.

3️⃣ **How does it compare to the past week?**  
- Show sales, margin, or ROAS trendlines. Mention any big shifts in performance.

4️⃣ **What needs attention?**  
- Flag high ad spend with low ROAS, unsub spikes, flows with no conversions, etc.  
- If relevant, include this note: ${context["⚠️ Anything to Watch"] || "N/A"}

5️⃣ **Are we on pace to hit our goals?**  
- Use sales goal (${context["Sales Goal (Daily)"] || "N/A"}) and margin goal (${context["Gross Margin Target"] || "N/A"}) to evaluate progress.

${trailingSummaryContext ? "\n\n📅 Previous 7 Days of Summaries:\n" + trailingSummaryContext : ""}

Here is the performance data:
${JSON.stringify(report, null, 2)}
`;

  // Step 4: Call OpenAI
  const payload = {
    model: "gpt-4o",
    messages: [
      { role: "system", content: "You are an ecommerce performance analyst. You summarize performance reports for executives." },
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

  writeAiSummaryToColumnP(summary);


  return summary;
}

function getAiContextFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AI");
  if (!sheet) {
    Logger.log("⚠️ AI sheet not found.");
    return {};
  }

  const data = sheet.getDataRange().getValues();
  const context = {};

  data.forEach(row => {
    const label = row[0]?.toString().trim();
    const value = row[1]?.toString().trim();
    if (label && value) {
      context[label] = value;
    }
  });

  return context;
}

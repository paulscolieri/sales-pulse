function generateAndSendDailySummary() {

  try {
    // Step 1: Pull summaries
    const shopifySummary = getShopifySummaryFromSheet();
    const flowSummary = getFlowSummaryFromSheet();
    const campaignSummary = getCampaignSummaryFromSheet();
    const adsSummary = getMetaAdSummaryFromSheet();
    const context = getAiContextFromSheet();

    const campaignContext = `
    📧 *Klaviyo Campaign Overview*
    - Campaigns sent in last 7 days: ${campaignSummary.numCampaigns7d}
    - Last campaign sent: ${campaignSummary.lastCampaignDate}
    - Best campaign (7d): ${campaignSummary.topCampaign7d} – Revenue: $${campaignSummary.topCampaignRevenue7d}
    - Average Open Rate (7d): ${campaignSummary.avgOpenRate7d}%
    - Average Click Rate (7d): ${campaignSummary.avgClickRate7d}%
    `;

    // Step 2: Combine report
    const report = {
      date: new Date().toISOString().split("T")[0],
      shopify: shopifySummary,
      campaign_context: campaignContext,
      klaviyo_flows: flowSummary,
      klaviyo_campaigns: campaignSummary,
      meta_ads: adsSummary
    };

    // Trailing 7days of AI summaries
    //const trailingSummaryContext = getLast7AiSummaries();

    // Step 3: Construct GPT prompt using the 5-question structure
    const prompt = `
    You're an expert business coach, digital marketer, and ecommerce performance analyst. Write a clear, actionable executive summary (under 300 words) of yesterday’s performance for the brand "${context["Brand Name"] || "our brand"}".

    Tone: ${context["Brand Voice"] || "concise and direct"}  
    Business Model: ${context["Business Model"] || "ecommerce"}  
    Key Products: ${context["Key Products"] || "N/A"}  
    Focus this week: ${context["Strategic Focus This Week"] || "N/A"}  
    Current promo: ${context["Current Offers/Promos"] || "None"}  
    Known external factors: ${context["Known External Factors"] || "None"}
    Other Context: ${context["Other Context"] || "None"}    

    Use the following structure to answer these 5 exec-level questions. **In your analysis, prioritize recent trends and performance drivers over single-day anomalies. Include actionable suggestions for each relevant section.**

    **Special instructions:**
    - Focus your Klaviyo analysis on engagement trends and overall revenue performance over the past 7 days, including number of campaigns sent (numCampaigns7d), average open rate (avgOpenRate7d), average click rate (avgClickRate7d), last campaign date (lastCampaignDate), and top-performing campaign revenue. 
    - Do not evaluate Klaviyo performance based solely on single-day revenue when no campaign was sent. Instead, compare recent 7-day performance to engagement and conversion goals.
    - If open rates are high but click rates or conversions are low across recent campaigns, recommend specific improvements to offers, creative, or segmentation.
    - Provide clear, direct next steps as action items based on weekly trends.
    - Do not add a sign-off or closing; end your response after the final recommendations.

    1️⃣ **Did we make money yesterday?**  
    - Report daily revenue, profit, and margin. Compare to goals if available.

    2️⃣ **What worked or didn’t?**  
    - Highlight standout flows, campaigns, ads, or products. Include best-performing Klaviyo campaigns or engagement trends.

    3️⃣ **How does it compare to the past week?**  
    - Analyze sales, margin, ROAS, Klaviyo engagement trends (open/click rates), and campaign cadence. Identify shifts or patterns.

    4️⃣ **What needs attention?**  
    - Provide actionable advice on flows, campaigns, or ads. If there are repeated days of low conversions, call this out. Suggest improvements to CTAs, segmentation, or creative.

    5️⃣ **Are we on pace to hit our goals?**  
    - Evaluate performance relative to sales (${context["Sales Goal (Daily)"] || "N/A"}) and margin goals (${context["Gross Margin Target"] || "N/A"}). Recommend promotional opportunities to get back on track if needed.

    **Important:**
    - Do not invent or assume data beyond what is included in the performance report.
    - If specific metrics (like inventory levels or unrelated brand names) are not present in the data, do not mention them or speculate.

    **Important:**
    - Limit your analysis strictly to the data included in the performance report JSON.
    - Do not invent data, mention unrelated brands, or speculate on inventory levels if inventory metrics are not explicitly provided.
    - Base all recommendations only on the structured performance data supplied.

    Here is the performance data:
    ${JSON.stringify(report, null, 2)}
    `;

    // Step 4: Call OpenAI
    const payload = {
      model: "gpt-4o",
      messages: [
        { role: "system", content: "You are a helpful, intelligent, world-class ecommerce performance analyst. You summarize performance reports for executives with key KPIs and business drivers. You provide suggestions on where to look if performance needs to be improved" },
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
    sendSummaryToSlack(summary);

    return summary;
  } catch (error) {
  Logger.log(`❌ Error in generateAndSendDailySummary: ${error.message}`);
  sendSummaryToSlack(`❌ *AI Summary Error:*\n\`\`\`${error.stack}\`\`\``);
  throw error; // Optional: rethrow for Apps Script visibility
  }
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

function sendSummaryToSlack(summaryText) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty("SLACK_WEBHOOK_URL");

  if (!webhookUrl) {
    Logger.log("⚠️ Slack webhook URL not found in script properties.");
    return;
  }

  const payload = {
    text: `📈 *Daily Exec Summary – ${new Date().toLocaleDateString()}*\n\n${summaryText}`
  };

  UrlFetchApp.fetch(webhookUrl, {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });

  Logger.log("✅ AI summary sent to Slack.");
  
}


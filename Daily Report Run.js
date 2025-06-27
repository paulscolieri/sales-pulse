function runDailyReporting() {
  Logger.log("Starting daily Klaviyo email reporting...");
  try {
    
    //Adding new date row and formulas to summary sheet
    addNewDateRowToDailySummary();

    // 1. Update Campaign Reference with any new metadata
    writeKlaviyoCampaignReferenceToSheet();
    Utilities.sleep(5000); // small delay to ensure data is written

    // 2. Write campaign performance stats to Email Log
    writeCampaignDailyStatsToLog();
    Utilities.sleep(2000); // optional delay

    // 3. Backfill any missing metadata in Email Log
    backfillCampaignMetadataInEmailLog();
    Utilities.sleep(2000); // optional delay

    // 4. Write Flow Stats
    writeFlowEmailStatsToSheet();

    // 5. Write Shopify Orders
    writeLineItemsShopifyStyle();

    //6. Write Meta Ads Stats
    writeMetaAdInsightsToSheet();

    // 7. Log shopify trends for AI
    const shopifyTrends = getShopifyKpiTrends();
    Logger.log(shopifyTrends);

    // 🧠 Generate AI Summary
    const summary = generateAndSendDailySummary();
    Logger.log("✅ Summary complete.");
    Logger.log(summary);

    Logger.log("✅ Daily email reporting complete.");

  } catch (error) {
    Logger.log(`❌ Error in runDailyReporting: ${error.message}`);
    sendSummaryToSlack(`❌ *Daily Reporting Error:*\n\`\`\`${error.stack}\`\`\``);
    throw error; // Optional: rethrow if you want Apps Script to fail the execution visibly
  }
}

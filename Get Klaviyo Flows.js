function getYesterdayFlowEmailStats() {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');
  const revision = '2025-04-15';
  const conversionMetricId = getConversionMetricId();
  const url = 'https://a.klaviyo.com/api/flow-values-reports';

  const payload = {
    data: {
      type: "flow-values-report",
      attributes: {
        timeframe: { key: "yesterday" },
        statistics: [
          "recipients",
          "opens",
          "opens_unique",
          "open_rate",
          "clicks",
          "clicks_unique",
          "click_rate",
          "click_to_open_rate",
          "conversions",
          "conversion_uniques",
          "conversion_value",
          "conversion_rate",
          "revenue_per_recipient",
          "unsubscribes",
          "unsubscribe_rate",
          "spam_complaints",
          "spam_complaint_rate"
        ],
        conversion_metric_id: conversionMetricId,
        filter: 'equals(send_channel,"email")'
      }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/vnd.api+json",
    headers: {
      "Authorization": `Klaviyo-API-Key ${apiKey}`,
      "Accept": "application/vnd.api+json",
      "revision": revision
    },
    payload: JSON.stringify(payload)
  });

  const data = JSON.parse(response.getContentText());
  const results = data?.data?.attributes?.results || [];

  if (results.length === 0) {
    Logger.log("No flow email stats for yesterday.");
    return [];
  }

  // Optional: validate message IDs against known reference sheet
  const referenceMap = buildFlowMessageReferenceMap();
  const missingMessages = [];

  const finalStats = results.map(item => {
    const flow_id = item.groupings.flow_id;
    const flow_message_id = item.groupings.flow_message_id;

    if (!referenceMap[flow_message_id]) {
      missingMessages.push(flow_message_id);
    }

    return {
      flow_id,
      flow_message_id,
      stats: item.statistics
    };
  });

  if (missingMessages.length > 0) {
    Logger.log(`Missing metadata for ${missingMessages.length} message IDs:`);
    Logger.log(missingMessages.join(", "));
  }

  return finalStats;
}

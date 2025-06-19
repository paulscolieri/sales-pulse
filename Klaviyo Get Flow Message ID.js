function getFlowAndMessageName(flowMessageId) {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');
  const revision = '2025-04-15';

  const url = `https://a.klaviyo.com/api/flow-messages/${flowMessageId}?include=flow`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': `Klaviyo-API-Key ${apiKey}`,
      'Accept': 'application/json',
      'revision': revision
    }
  });

  const json = JSON.parse(response.getContentText());
  const messageName = json.data?.attributes?.name || "";
  const flowName = json.included?.[0]?.attributes?.name || "";

  return { messageName, flowName };
}

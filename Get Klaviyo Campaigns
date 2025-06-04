function getYesterdayCampaigns() {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');

  const today = new Date();
  const startDate = new Date(today);
  startDate.setDate(today.getDate() - 1); // yesterday start
  startDate.setHours(0, 0, 0, 0);

  const endDate = new Date(today);
  endDate.setDate(today.getDate() - 1); // yesterday end
  endDate.setHours(23, 59, 59, 999);

  const startISOString = startDate.toISOString();
  const endISOString = endDate.toISOString();

  const rawFilter = `equals(messages.channel,"email"),greater-or-equal(created_at,${startISOString}),less-than(created_at,${endISOString})`;
  const encodedFilter = encodeURIComponent(rawFilter);
  const url = `https://a.klaviyo.com/api/campaigns?filter=${encodedFilter}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'Authorization': `Klaviyo-API-Key ${apiKey}`,
      'Accept': 'application/json',
      'revision': '2025-04-15'
    }
  });

  const data = JSON.parse(response.getContentText());
  Logger.log(data);
}

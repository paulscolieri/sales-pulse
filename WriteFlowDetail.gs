function writeKlaviyoFlowEmailDataToSheet() {
  const scriptProps = PropertiesService.getScriptProperties();
  const apiKey = scriptProps.getProperty('KLAVIYO_API_KEY');
  const revision = '2025-04-15';

  const sheetName = 'Flow Messages';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear(); // clear old data
  }

  // Write header
  sheet.appendRow(['Flow Name', 'Action Index', 'Email Name', 'Subject Line']);

  let nextPage = 'https://a.klaviyo.com/api/flows?page[size]=50';

  while (nextPage) {
    const flowRes = UrlFetchApp.fetch(nextPage, {
      method: 'get',
      headers: {
        'Authorization': `Klaviyo-API-Key ${apiKey}`,
        'Accept': 'application/json',
        'revision': revision
      }
    });

    const flowData = JSON.parse(flowRes.getContentText());

    (flowData.data || []).forEach(flow => {
      const flowId = flow.id;
      const flowName = flow.attributes.name;

      const defUrl = `https://a.klaviyo.com/api/flows/${flowId}/?additional-fields[flow]=definition`;

      try {
        const defRes = UrlFetchApp.fetch(defUrl, {
          method: 'get',
          headers: {
            'Authorization': `Klaviyo-API-Key ${apiKey}`,
            'Accept': 'application/json',
            'revision': revision
          }
        });

        const defData = JSON.parse(defRes.getContentText());
        const actions = defData.data.attributes.definition.actions;

        (actions || []).forEach((action, index) => {
          if (action.type === 'send-email' && action.data?.message) {
            const msg = action.data.message;
            const msgName = msg.name || '';
            const msgSubject = msg.subject_line || '';
            sheet.appendRow([flowName, index, msgName, msgSubject]);
          }
        });

      } catch (e) {
        Logger.log(`Error fetching flow definition for ${flowName}: ${e.message}`);
      }
    });

    nextPage = flowData.links?.next || null;
  }

  Logger.log('Flow message info written to sheet.');
}

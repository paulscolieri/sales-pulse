function setCredentials_template() {
  PropertiesService.getScriptProperties().setProperty('SHOPIFY_ADMIN_TOKEN', 'your-access-token-here');
  PropertiesService.getScriptProperties().setProperty('SHOPIFY_DOMAIN', 'your-store.myshopify.com');
  PropertiesService.getScriptProperties().setProperty('KLAVIYO_API_KEY', 'pk_your-klaviyo-private-api-key');
  PropertiesService.getScriptProperties().setProperty('META_ACCESS_TOKEN', 'YOUR_META_TOKEN');
  PropertiesService.getScriptProperties().setProperty('META_AD_ACCOUNT_ID', 'act_{META_ACCOUNT_ID}');
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', 'YOUR_OPEN_AI_API');
  
}


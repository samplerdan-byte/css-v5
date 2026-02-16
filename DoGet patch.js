// ============================================================
// 🔧 PATCH: Update doGet() in WebApp.gs to route portal pages
// ============================================================
// Replace your existing doGet() with this version:

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'main';
  
  // Client Status Portal — token-based access
  if (page === 'portal') {
    return HtmlService.createTemplateFromFile('ClientPortal')
      .evaluate()
      .setTitle('CSS — Order Status Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // Legacy scanner URL
  if (page === 'scanner') {
    return HtmlService.createTemplateFromFile('WebAppScanner')
      .evaluate()
      .setTitle('CSS Scanner')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
if (page === 'portal') {
  return HtmlService.createTemplateFromFile('ClientPortal')
    .evaluate()
    .setTitle('CSS — Order Status Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

  
  // Default: full web app
  return HtmlService.createTemplateFromFile('WebAppScanner')
    .evaluate()
    .setTitle('CSS V3 — Commodity Sampler Services')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

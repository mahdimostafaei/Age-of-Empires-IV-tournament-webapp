// =========================================================================
// WebApp.gs - The Master HTML Server
// =========================================================================

function doGet(e) {
  // 1. Grab the master Index shell
  const template = HtmlService.createTemplateFromFile('Index');
  
  // 2. Evaluate and staple everything together, then serve it as one webpage
  return template.evaluate()
      .setTitle('Age of Empires IV - Tournament Hub')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) // Allows embedding if needed
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Makes it mobile-friendly
}

// This is the stapler function. The HTML file calls this to pull in other files.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
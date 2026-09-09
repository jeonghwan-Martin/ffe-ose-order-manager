// Run manually in the STORAGE Apps Script project. Does not alter properties,
// call CalendarApp or send email. Includes orphan keys omitted by listProjects.
function exportOrderManagerProperties() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const selected = {};
  Object.keys(properties).forEach(function (key) {
    if (key === 'project-index' || key === 'item-catalog' || key.indexOf('project-data:') === 0) {
      selected[key] = properties[key];
    }
  });
  const exportData = {format:'ffe-ose-properties-export-v1', exportedAt:new Date().toISOString(), properties:selected};
  const file = DriveApp.createFile('ffe-ose-properties-' + Date.now() + '.json', JSON.stringify(exportData,null,2), MimeType.PLAIN_TEXT);
  Logger.log(file.getUrl());
}

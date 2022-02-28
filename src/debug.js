/**
 * Delete all user properties to reproduce what first-time users will see.
 */
function resetUserProperties() {
  PropertiesService.getUserProperties().deleteAllProperties();
}

if (typeof module === 'object') {
  module.exports = { resetUserProperties };
}

/* exported resetUserProperties */

function resetUserProperties() {
  PropertiesService.getUserProperties().deleteAllProperties();
}
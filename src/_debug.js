/* global UP_KEY_CREATED_DRAFT_IDS, UP_KEY_PREV_CONFIG, UP_KEY_USER_CONFIG */
/* exported resetUserProperties */

function resetUserProperties() {
  PropertiesService.getUserProperties()
    .deleteProperty(UP_KEY_USER_CONFIG)
    .deleteProperty(UP_KEY_PREV_CONFIG)
    .deleteProperty(UP_KEY_CREATED_DRAFT_IDS);
}
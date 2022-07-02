// Mock GAS classes

/**
 * @see https://developers.google.com/apps-script/reference/properties/properties-service?hl=en
 */
class PropertiesService {
  static getUserProperties() {
    return new UserProperties();
  }
}

class UserProperties {
  constructor(propertiesString = '{}') {
    this.userProperties = JSON.parse(propertiesString);
  }
  getProperty(key) {
    return this.userProperties[key];
  }
  /**
   *
   * @param {String} key
   * @param {String} value
   * @see
   */
  setProperty(key, value) {
    this.userProperties[key] = value;
  }
}

module.exports = {
  PropertiesService,
};

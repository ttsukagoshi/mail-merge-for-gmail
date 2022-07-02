// Mock GAS classes

const EMAIL = 'myEmail@example.com';

/**
 * @see https://developers.google.com/apps-script/reference/base/session?hl=en
 */
class Session {
  static getActiveUser() {
    return new User();
  }
}

class User {
  constructor() {
    this.email = EMAIL;
  }
  getEmail() {
    return this.email;
  }
}

module.exports = {
  Session,
};

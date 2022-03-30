import * as Msal from 'msal';

export default class AuthService {
  constructor() {
    let PROD_REDIRECT_URI = 'http://localhost:3000/';
    let redirectUri = window.location.origin;
    if (window.location.hostname !== '127.0.0.1') {
      redirectUri = PROD_REDIRECT_URI;
    }
    this.applicationConfig = {
      clientID: '79e112ac-472c-47c6-9935-08542f1ae5f5',
      graphScopes: ['user.read']
    };
    this.app = new Msal.UserAgentApplication(
      this.applicationConfig.clientID,
      '',
      () => {
        // callback for login redirect
      },
      {
        redirectUri
      }
    );
  }
  login = () => {
    return this.app.loginPopup(this.applicationConfig.graphScopes).then(
      idToken => {
        const user = this.app.getUser();
        if (user) {
          return user;
        } else {
          return null;
        }
      },
      () => {
        return null;
      }
    );
  };
  logout = () => {
    this.app.logout();
  };
  getToken = () => {
    return this.app.acquireTokenSilent(this.applicationConfig.graphScopes).then(
      accessToken => {
        return accessToken;
      },
      error => {
        return this.app
          .acquireTokenPopup(this.applicationConfig.graphScopes)
          .then(
            accessToken => {
              return accessToken;
            },
            err => {
              console.error(err);
            }
          );
      }
    );
  };
}

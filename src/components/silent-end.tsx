import React, { Component } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { AuthenticationContext, AdalConfig } from "react-adal";
import { clientId } from "../utils/config";

// ADAL.js configuration
const config: AdalConfig = {
  clientId: clientId,
  redirectUri: window.location.origin + "/start",
  cacheLocation: "localStorage",
  navigateToLoginRequestUrl: false,
};

export class SilentEnd extends Component {
  componentDidMount() {
    microsoftTeams.initialize();

    const authContext = new AuthenticationContext(config);
    if (authContext.isCallback(window.location.hash)) {
      authContext.handleWindowCallback(window.location.hash);

      // Only call notifySuccess or notifyFailure if this page is in the authentication popup
      if (window.opener) {
        if (authContext.getCachedUser()) {
          microsoftTeams.authentication.notifySuccess();
        } else {
          microsoftTeams.authentication.notifyFailure(authContext.getLoginError());
        }
      }
    }
  }

  render() {
    return <h3>Done!</h3>;
  }
}

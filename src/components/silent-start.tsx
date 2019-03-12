import React, { Component } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { AuthenticationContext, AdalConfig } from "react-adal";
import { clientId } from "../utils/config";

// ADAL.js configuration
const config: AdalConfig = {
  clientId: clientId,
  redirectUri: window.location.origin + "/end",
  cacheLocation: "localStorage",
  navigateToLoginRequestUrl: false,
};

export class SilentStart extends Component {
  componentDidMount() {
    microsoftTeams.initialize();
    microsoftTeams.getContext((context: microsoftTeams.Context) => {

      // Setup extra query parameters for ADAL
      // - openid and profile scope adds profile information to the id_token
      // - login_hint provides the expected user name
      if (context.upn) {
        config.extraQueryParameter = "scope=openid+profile&login_hint=" + encodeURIComponent(context.upn);
      } else {
        config.extraQueryParameter = "scope=openid+profile";
      }

      // Navigate to the AzureAD login page
      // @ts-ignore
      const authContext = new AuthenticationContext(config);
      authContext.login();
    });
  }

  render() {
    return <h3>Authenticating...</h3>;
  }
}

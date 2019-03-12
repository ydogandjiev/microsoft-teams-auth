import React, { Component, RefObject } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { AuthenticationContext, AdalConfig } from "react-adal";
import { clientId } from "../utils/config";

// ADAL.js configuration
const config: AdalConfig = {
  clientId: clientId,
  redirectUri: window.location.origin,
  cacheLocation: "localStorage",
  navigateToLoginRequestUrl: false,
  extraQueryParameter: "",
};

export class Silent extends Component {
  btnLoginRef = React.createRef<HTMLButtonElement>();
  divError = React.createRef<HTMLDivElement>();
  divProfile = React.createRef<HTMLDivElement>();
  profileDisplayName = React.createRef<HTMLSpanElement>();
  profileUpn = React.createRef<HTMLSpanElement>();
  profileObjectId = React.createRef<HTMLSpanElement>();

  componentDidMount() {
    microsoftTeams.initialize();
    microsoftTeams.getContext((context: microsoftTeams.Context) => {

      // Loads data for the given user
      this.loadData(context.upn);
    });
  }

  render() {
    return (
      <>
        <p>This sample demonstrates silent authentication in a Microsoft Teams tab.</p>
        <p>
          The tab will try to get an id token for the user silently, validate the token and show the profile information decoded from the token. The "Login" button will appear only if silent
					authentication failed.
				</p>

        <button ref={this.btnLoginRef} id="btnLogin" onClick={this.login} style={{ display: "none" }}>
          Login to Azure AD
				</button>

        <div>
          <div ref={this.divError} id="divError" style={{ display: "none" }} />
          <div ref={this.divProfile} id="divProfile" style={{ display: "none" }}>
            <div>
              <b>Name:</b> <span ref={this.profileDisplayName} id="profileDisplayName" />
            </div>
            <div>
              <b>UPN:</b> <span ref={this.profileUpn} id="profileUpn" />
            </div>
            <div>
              <b>Object id:</b> <span ref={this.profileObjectId} id="profileObjectId" />
            </div>
          </div>
        </div>
      </>
    );
  }

  loadData = (upn: string | undefined) => {
    // Setup extra query parameters for ADAL
    // - openid and profile scope adds profile information to the id_token
    // - login_hint provides the expected user name
    if (upn) {
      config.extraQueryParameter = "scope=openid+profile&login_hint=" + encodeURIComponent(upn);
    } else {
      config.extraQueryParameter = "scope=openid+profile";
    }

    const authContext = new AuthenticationContext(config);

    // See if there's a cached user and it matches the expected user
    const user = authContext.getCachedUser();
    if (user) {
      if (user.userName !== upn) {
        // User doesn't match, clear the cache
        authContext.clearCache();
      }
    }

    // Get the id token (which is the access token for resource = clientId)
    const token = authContext.getCachedToken(config.clientId);
    if (token) {
      // showProfileInformation(token);
      console.log("token ", token);
    } else {
      // No token, or token is expired
      (authContext as any)._renewIdToken((err: string, idToken: string) => {
        if (err) {
          console.log("Renewal failed: " + err);

          // Failed to get the token silently; show the login button
          this.btnLoginRef.current!.setAttribute("style", "display: ''");

          // You could attempt to launch the login popup here, but in browsers this could be blocked by
          // a popup blocker, in which case the login attempt will fail with the reason FailedToOpenWindow.
        } else {
          // showProfileInformation(idToken);
          console.log("token ", idToken);
        }
      });
    }
  };

  login = () => {
    this.divError.current!.setAttribute("style", "display: 'none'");
    this.divError.current!.setAttribute("text", "");
    this.divProfile.current!.setAttribute("style", "display: 'none'");

    microsoftTeams.authentication.authenticate({
      url: window.location.origin + "/start",
      width: 600,
      height: 535,

      successCallback: result => {
        const authContext = new AuthenticationContext(config);
        const idToken = authContext.getCachedToken(config.clientId);
        if (idToken) {
          // showProfileInformation(idToken);
        } else {
          console.error("Error getting cached id token. This should never happen.");

          // At this point we have to get the user involved, so show the login button
          this.btnLoginRef.current!.setAttribute("style", "display: ''");
        }
      },

      failureCallback: reason => {
        if (reason === "CancelledByUser" || reason === "FailedToOpenWindow") {
          console.log("Login was blocked by popup blocker or canceled by user.");
        } else {
          console.log("Login failed: " + reason);
        }

        // At this point we have to get the user involved, so show the login button
        this.btnLoginRef.current!.setAttribute("style", "display: ''");
        this.divError.current!.setAttribute("style", "display: ''");
        this.divError.current!.setAttribute("text", reason!);
        this.divProfile.current!.setAttribute("style", "display: 'none'");
      },
    });
  };
}

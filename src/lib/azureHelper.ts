import {
  getUserDetails,
  getEvents,
  getUsersByDisplayName,
} from "./graphService";
import { msalInstance, GRAPH_REQUESTS /* , isIE */ } from "./authUtils";

//const useRedirectFlow = isIE();
const useRedirectFlow = true;

class AzureHelper {
  LoginRedirect() {
    msalInstance.loginRedirect(GRAPH_REQUESTS.LOGIN);
  }

  SignOut() {
    msalInstance.logout();
  }

  async getUserProfile() {
    try {
      var accessToken = await this.getAccessToken(
        GRAPH_REQUESTS.PROFILE,
        useRedirectFlow
      );

      if (accessToken) {
        // Get the user's profile from Graph
        var user = await getUserDetails(accessToken);
        return user;
      }
    } catch (err) {
      throw err;
    }
  }

  async getUserEvents() {
    try {
      var accessToken = await this.getAccessToken(
        GRAPH_REQUESTS.CALENDAR,
        useRedirectFlow
      );

      if (accessToken) {
        // Get the user's profile from Graph
        var events = await getEvents(accessToken);
        return events;
      }
    } catch (err) {
      throw err;
    }
  }

  async getUsersByName(name: string) {
    try {
      var accessToken = await this.getAccessToken(
        GRAPH_REQUESTS.CALENDAR,
        useRedirectFlow
      );

      if (accessToken) {
        // Get the user's profile from Graph
        var users = await getUsersByDisplayName(accessToken, name);
        return users;
      }
    } catch (err) {
      throw err;
    }
  }

  /* async getAccessToken(request: any, redirect: any) {
    return msalInstance.acquireTokenSilent(request).catch((error) => {
      // Call acquireTokenPopup (popup window) in case of acquireTokenSilent failure
      // due to consent or interaction required ONLY
      if (requiresInteraction(error.errorCode)) {
        return redirect ? msalInstance.acquireTokenRedirect(request) : msalInstance.acquireTokenPopup(request);
      } else {
        console.error('Non-interactive error:', error.errorCode);
      }
    });
  } */

  async getAccessToken(scopes: any, redirect: any): Promise<string> {
    //debugger;
    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token
      var silentResult = await msalInstance.acquireTokenSilent(scopes);

      return silentResult.accessToken;
    } catch (err) {
      // If a silent request fails, it may be because the user needs
      // to login or grant consent to one or more of the requested scopes
      if (this.isInteractionRequired(err)) {
        var interactiveResult = await msalInstance.acquireTokenPopup(scopes);

        return interactiveResult.accessToken;
      } else {
        throw err;
      }
    }
  }

  normalizeError(error: string | Error): any {
    var normalizedError = {};
    if (typeof error === "string") {
      var errParts = error.split("|");
      normalizedError =
        errParts.length > 1
          ? { message: errParts[1], debug: errParts[0] }
          : { message: error };
    } else {
      normalizedError = {
        message: error.message,
        debug: JSON.stringify(error),
      };
    }
    return normalizedError;
  }

  isInteractionRequired(error: Error): boolean {
    if (!error.message || error.message.length <= 0) {
      return false;
    }

    return (
      error.message.indexOf("consent_required") > -1 ||
      error.message.indexOf("interaction_required") > -1 ||
      error.message.indexOf("login_required") > -1
    );
  }

  /* async onSignIn(redirect) {
    redirect = true;
    if (redirect) {
        return msalApp.loginRedirect(GRAPH_REQUESTS.LOGIN);
    }

    const loginResponse = await msalApp
        .loginPopup(GRAPH_REQUESTS.LOGIN)
        .catch(error => {
            this.setState({
                error: error.message
            });
        });

    if (loginResponse) {
        this.setState({
            account: loginResponse.account,
            error: null
        });

        const tokenResponse = await this.acquireToken(
            GRAPH_REQUESTS.LOGIN
        ).catch(error => {
            this.setState({
                error: error.message
            });
        });

        if (tokenResponse) {
            const graphProfile = await fetchMsGraph(
                GRAPH_ENDPOINTS.ME,
                tokenResponse.accessToken
            ).catch(() => {
                this.setState({
                    error: "Unable to fetch Graph profile."
                });
            });

            if (graphProfile) {
                this.setState({
                    graphProfile
                });
            }

            if (
                tokenResponse.scopes.indexOf(GRAPH_SCOPES.MAIL_READ) > 0
            ) {
                return this.readMail(tokenResponse.accessToken);
            }
        }
    }
  } */
}
export default new AzureHelper();

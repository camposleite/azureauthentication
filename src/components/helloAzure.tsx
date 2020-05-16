import * as React from "react";
import azureHelper from "../lib/azureHelper";
import { msalInstance } from "../lib/authUtils";

import Calendar from "./calendar";

class HelloAzure extends React.Component {
  state = {
    userName: "anonymous",
    userIsAuthenticated: false,
    error: "",
    events: [],
  };

  handleLogin = () => {
    azureHelper.LoginRedirect();
  };

  handleSignOut = () => {
    azureHelper.SignOut();
  };

  handleGetCalendar = async () => {
    var events = await azureHelper.getUserEvents();
    this.setState({ events: events.value });
  };

  async componentWillMount() {
    msalInstance.handleRedirectCallback((error: any) => {
      if (error) {
        const errorMessage = error.errorMessage
          ? error.errorMessage
          : "Unable to acquire access token.";
        // setState works as long as navigateToLoginRequestUrl: false
        this.setState({
          error: errorMessage,
        });
      }
    });

    const user = msalInstance.getAccount();

    if (user != null) {
      this.setState({ userName: user.userName, userIsAuthenticated: true });
    }
  }

  public render() {
    return (
      <div>
        Hello, {this.state.userName}
        <div>
          {this.state.userIsAuthenticated ? (
            <div>
              <input
                type="button"
                value="Sign out"
                onClick={this.handleSignOut}
              />
              <hr />
              <input
                type="button"
                value="Get calendar events"
                onClick={this.handleGetCalendar}
              />

              <Calendar events={this.state.events} />
            </div>
          ) : (
            <input type="button" value="Sign in" onClick={this.handleLogin} />
          )}
        </div>
      </div>
    );
  }
}

export default HelloAzure;

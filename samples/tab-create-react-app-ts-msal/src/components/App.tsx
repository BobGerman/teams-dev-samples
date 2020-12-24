import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";
import { HashRouter as Router, Route } from "react-router-dom";
import PrivacyPage from "./PrivacyPage";
import TermsOfUsePage from "./TermsOfUsePage";
import TabPage from './TabPage';
import TabConfigPage from "./TabConfigPage";
import WebPage from "./WebPage";
import AuthService from '../services/AuthService/MsalAuthService';

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export interface IAppProps { };
export interface IAppState {
  authInitialized: boolean;
}

export default class App extends React.Component<IAppProps, IAppState> {

  constructor(props: IAppProps) {
    super(props);
    this.state = {
      authInitialized: false
    }
  }

  componentDidMount() {
    // React routing and OAuth don't play nice together
    // Take care of the OAuth fun before routing
    AuthService.init().then(() => {
      this.setState({
        authInitialized: true
      });
    })
  }

  render() {

    if (microsoftTeams) {

      // Set up routes that don't use the Teams SDK
      if (window.parent === window.self) {
        return (
          <div className="App">
            <Router>
              <Route exact path="/privacy" component={PrivacyPage} />
              <Route exact path="/termsofuse" component={TermsOfUsePage} />
              <Route exact path="/web" component={WebPage} />
              <Route exact path="/" component={WebPage} />
              <Route exact path="/tab" component={TeamsHostError} />
              <Route exact path="/config" component={TeamsHostError} />
            </Router>
          </div>
        );
      }

      // Initialize the Microsoft Teams SDK
      microsoftTeams.initialize(window as any);

      // Set up routes that use the Teams SDK
      return (
        <div className="App">
          <Router>
            <Route exact path="/tab" component={TabPage} />
            <Route exact path="/config" component={TabConfigPage} />
          </Router>
        </div>
      );
    }

    // Error when the Microsoft Teams SDK is not found
    // in the project.
    return (
      <h3>Microsoft Teams SDK not found.</h3>
    );

    // Thiscomponent displays an error message when a route
    // requiring Teams is run outside of Teams
    function TeamsHostError() {
      return (
        <div>
          <h3 className="Error">This page is for use in Microsoft Teams.</h3>
        </div>
      );
    }
  }

}

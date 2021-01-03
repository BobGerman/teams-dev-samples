import React from 'react';
import { IConfig, ConfigService } from '../../services/ConfigService/ConfigService';
import ThemeService from '../../services/ThemeService/ThemeService';
import * as microsoftTeams from "@microsoft/teams-js";
import { Provider, Header, ThemePrepared } from "@fluentui/react-northstar";
import IAuthService from '../../services/AuthService/IAuthService';
import AuthService from '../../services/AuthService/TeamsAuthService';
import * as MicrosoftGraphClient from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

/**
 * The web UI to display in the Teams UI
 */
export interface ITabPageProps { }
export interface ITabPageState {
  config?: IConfig;
  teamsContext?: microsoftTeams.Context;
  theme: ThemePrepared;
  graphService?: MicrosoftGraphClient.Client;
  messages: MicrosoftGraph.Message[];
  error: string;
}

export default class TabPage extends React.Component<ITabPageProps, ITabPageState> {

  constructor(props: ITabPageProps) {
    super(props);
    this.state = {
      config: undefined,
      teamsContext: undefined,
      theme: ThemeService.getFluentTheme(),
      graphService: undefined,
      messages: [],
      error: ""
    }
  }

  async componentDidMount() {

    // 1. Get tab configuration information
    let { config, teamsContext } = await ConfigService.getContextAndConfig();
    this.setState({
      config: config,
      teamsContext: teamsContext,
      theme: ThemeService.getFluentTheme(teamsContext.theme)
    });

    // 2. Handle theme changes
    ThemeService.registerOnThemeChangeHandler((newTheme) => {
      this.setState({
        theme: newTheme
      });
    });

    // 3. Try to silently get messages
    await this.getMessages();

    // Per https://stackoverflow.com/questions/63765776/personal-tab-renders-fine-then-a-few-seconds-later-shows-there-was-a-problem-r/64048235#64048235
    // need to call both notifyAppLoaded() and notifySuccess() or Teams will error out after a few seconds
    microsoftTeams.appInitialization.notifyAppLoaded();
    microsoftTeams.appInitialization.notifySuccess();

  }

  render() {

    if (!this.state.messages.length) {

      // Earlier attempt to log in failed - show the button
      return (
        <Provider theme={this.state.theme}>
          <Header>{process.env.REACT_APP_MANIFEST_NAME}</Header>
          <p>Version {process.env.REACT_APP_MANIFEST_APP_VERSION}</p>
          { this.state.error ? <p>Error: {this.state.error}</p> : null}
          <button onClick={async () => {
            await this.getMessages();
          }}>Log in</button>
        </Provider>
      );

    } else {

      let key = 0;
      return (
        <Provider theme={this.state.theme}>
          <Header>{process.env.REACT_APP_MANIFEST_NAME}</Header>
          <p>Version {process.env.REACT_APP_MANIFEST_APP_VERSION}</p>
          { this.state.error ? <p>Error: {this.state.error}</p> : null}
          <p>{this.state.teamsContext?.teamName ?
            `You are in ${this.state.teamsContext?.teamName}` :
            `You are not in a Team`
          }</p>
          <p>Your app is running in the Teams UI</p>
          <p>Your short message is {this.state.config?.shortMessage}</p>
          <p>Username: {AuthService.getUsername()}</p>
          <ol>
            {
              this.state.messages.map(message => (
                <li key={key++}>EMAIL: {message.receivedDateTime}<br />{message.subject}
                </li>
              ))
            }
          </ol>
        </Provider>
      );
    }
  }

  private async getMessages(): Promise<void> {

    try {
      let client = await this.GraphClientFactory(AuthService);
      let messages = await this.getMessagesFromGraph(client);
      this.setState({
        messages: messages,
        error: ""
      });
    }
    catch (error) {
      this.setState({
        messages: [],
        error: error
      });
    }
  }

  // TO MOVE TO GRAPH SERVICE
  private async GraphClientFactory(authService: IAuthService): Promise<MicrosoftGraphClient.Client> {

    let result: MicrosoftGraphClient.Client;

    let scopes = process.env.REACT_APP_AAD_GRAPH_DELEGATED_SCOPES?.split(',') || [];

    // Ensure we are logged in
    if (!authService.isLoggedIn()) {

      await authService.login(scopes);

    }

    // Initialize a new Graph client
    result = MicrosoftGraphClient.Client.init({

      authProvider: async (done: MicrosoftGraphClient.AuthProviderCallback) => {
        const token = await AuthService.getAccessToken(scopes);
        done(null, token);
      }
    });

    return result;
  }

  private async getMessagesFromGraph(client: MicrosoftGraphClient.Client): Promise<MicrosoftGraph.Message[]> {

    return new Promise<MicrosoftGraph.Message[]>((resolve, reject) => {

      client
        .api("me/mailFolders/inbox/messages")
        .select(["receivedDateTime", "subject"])
        .top(15)
        .get(async (error: MicrosoftGraphClient.GraphError, response: any) => {
          if (!error) {
            resolve(response.value as MicrosoftGraph.Message[]);
          } else {
            reject(error);
          }
        });

    });
  }
}

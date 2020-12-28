import React from 'react';
import { IConfig, ConfigService } from '../services/ConfigService';
import ThemeService from '../services/ThemeService';
import * as microsoftTeams from "@microsoft/teams-js";
import { Provider, Header, ThemePrepared } from "@fluentui/react-northstar";
import TeamsAuthService from '../services/AuthService/TeamsAuthService';
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
  accessToken?: string;
  messages: MicrosoftGraph.Message[];
  error: string;
}

export default class WebPage extends React.Component<ITabPageProps, ITabPageState> {

  constructor(props: ITabPageProps) {
    super(props);
    this.state = {
      config: undefined,
      teamsContext: undefined,
      theme: ThemeService.getFluentTheme(),
      messages: [],
      error: ""
    }
  }

  async componentDidMount() {
    let { config, teamsContext } = await ConfigService.getContextAndConfig();
    this.setState({
      config: config,
      teamsContext: teamsContext,
      theme: ThemeService.getFluentTheme(teamsContext.theme)
    });
    ThemeService.registerOnThemeChangeHandler((newTheme) => {
      this.setState({
        theme: newTheme
      });
    });
    microsoftTeams.appInitialization.notifyAppLoaded();
  }

  render() {

    if (!this.state.config) {
      return <div>loading...</div>
    } else {

      let key = 0;
      return (
        <Provider theme={this.state.theme}>
          <Header>{process.env.REACT_APP_MANIFEST_NAME}</Header>
          <p>Version {process.env.REACT_APP_MANIFEST_APP_VERSION}</p>
          <p>{this.state.teamsContext?.teamName ?
            `You are in ${this.state.teamsContext?.teamName}` :
            `You are not in a Team`
          }</p>
          <p>Your app is running in the Teams UI</p>
          <p>Your short message is {this.state.config.shortMessage}</p>
          <button onClick={this.getMessages.bind(this)}>Get Mail</button>
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

  private msGraphClient?: MicrosoftGraphClient.Client;

  private async getMessages() {

    let scopes = process.env.REACT_APP_AAD_GRAPH_DELEGATED_SCOPES?.split(',') || [];
    const token = await TeamsAuthService.getAccessToken(scopes, microsoftTeams);

    if (! this.msGraphClient) {

      this.msGraphClient = MicrosoftGraphClient.Client.init({
        authProvider: async (done) => {
            done(null, token);
        }
    });

      this.msGraphClient
      .api("me/mailFolders/inbox/messages")
      .select(["receivedDateTime", "subject"])
      .top(15)
      .get(async (error: MicrosoftGraphClient.GraphError, response: any) => {
        if (!error) {
          this.setState(Object.assign({}, this.state, {
            messages: response.value as MicrosoftGraph.Message[]
          }));
        } else {
          this.setState({
            error: error.message
          });
        }
      });
    }
  }
}

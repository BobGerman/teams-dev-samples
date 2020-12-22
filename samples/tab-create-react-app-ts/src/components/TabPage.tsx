import React from 'react';
import { IConfig, ConfigService } from '../services/ConfigService';
import ThemeService from '../services/ThemeService';
import * as microsoftTeams from "@microsoft/teams-js";
import { Provider, Header, ThemePrepared } from "@fluentui/react-northstar";

/**
 * The web UI to display in the Teams UI
 */
export interface ITabPageProps { }
export interface ITabPageState {
  config?: IConfig;
  teamsContext?: microsoftTeams.Context;
  theme: ThemePrepared;
}

export default class WebPage extends React.Component<ITabPageProps, ITabPageState> {

  constructor(props: ITabPageProps) {
    super(props);
    this.state = {
      config: undefined,
      teamsContext: undefined,
      theme: ThemeService.getFluentTheme()
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

      return (
        <Provider theme={this.state.theme}>
          <Header>{process.env.REACT_APP_MANIFEST_NAME}</Header>
          <p>{this.state.teamsContext?.teamName ?
            `You are in ${this.state.teamsContext?.teamName}` :
            `You are not in a Team`
          }</p>
          <p>Your app is running in the Teams UI</p>
          <p>Your short message is {this.state.config.shortMessage}</p>
          <p></p>
        </Provider>
      );
    }
  }

}

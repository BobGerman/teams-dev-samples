import React from 'react';
import { IConfigInfo, ConfigService } from '../services/ConfigService';
import ThemeService from '../services/ThemeService';
import * as microsoftTeams from "@microsoft/teams-js";
import { Provider, Header, ThemePrepared } from "@fluentui/react-northstar";

/**
 * The web UI to display in the Teams UI
 */
export interface ITabPageProps { }
export interface ITabPageState {
  config?: IConfigInfo;
  theme: ThemePrepared;
}

export default class WebPage extends React.Component<ITabPageProps, ITabPageState> {

  constructor(props: ITabPageProps) {
    super(props);
    this.state = {
      config: undefined,
      theme: ThemeService.getFluentTheme()
    }
  }

  async componentDidMount() {
    let configInfo = await ConfigService.getConfigInfo();
    this.setState({
      config: configInfo,
      theme: ThemeService.getFluentTheme(configInfo.teamsContext?.theme)
    });
    ThemeService.registerOnThemeChangeHandler((newTheme) => {
      this.setState({
        theme: newTheme
      });
    });
    microsoftTeams.appInitialization.notifyAppLoaded();
  }

  render() {

    if (!this.state.config?.teamsContext) {
      return <div>loading...</div>
    } else {

      return (
        <Provider theme={this.state.theme}>
          <Header>{process.env.REACT_APP_MANIFEST_NAME}</Header>
          <p>{this.state.config.teamsContext.teamName ?
            `You are in ${this.state.config.teamsContext.teamName}` :
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

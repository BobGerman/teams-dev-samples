import React from 'react';
import { IConfigInfo, ConfigService } from '../services/ConfigService';
import * as microsoftTeams from "@microsoft/teams-js";
import { Provider, Header, ThemePrepared } from "@fluentui/react-northstar";
import { teamsTheme, teamsDarkTheme, teamsHighContrastTheme } from '@fluentui/react-northstar';

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
      theme: teamsTheme
    }
  }

  async componentDidMount() {
    let configInfo = await ConfigService.getConfigInfo();
    this.setState({
      config: configInfo,
      theme: this.getTheme(configInfo.teamsContext?.theme)
    });
    microsoftTeams.registerOnThemeChangeHandler((theme: string = "default"): void => {

      this.setState({
        theme: this.getTheme(theme)
      });
    });
    microsoftTeams.appInitialization.notifyAppLoaded();
  }

  private getTheme(theme?: string): ThemePrepared {
    let result = teamsTheme;
    if (theme === 'dark') result = teamsDarkTheme;
    if (theme === 'contrast') result = teamsHighContrastTheme
    return result
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

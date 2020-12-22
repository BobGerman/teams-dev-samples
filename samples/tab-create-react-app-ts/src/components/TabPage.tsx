import React from 'react';
import { IConfigInfo, ConfigService } from '../services/ConfigService';
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * The web UI to display in the Teams UI
 */
export interface ITabPageProps{ }
export interface ITabPageState {
  config?: IConfigInfo;
}

export default class WebPage extends React.Component<ITabPageProps, ITabPageState> {

  constructor (props: ITabPageProps) {
    super(props);
    this.state = {
      config: undefined
    }
  }

  async componentDidMount () {
    let configInfo = await ConfigService.getConfigInfo();
    this.setState({
      config: configInfo
    });
    microsoftTeams.appInitialization.notifyAppLoaded();
  }

  render() {

    if (!this.state.config?.teamsContext) {
      return <div>loading...</div>
    } else {
      return (
        <div>
          <h1>{process.env.REACT_APP_MANIFEST_NAME}</h1>
          <p>{this.state.config.teamsContext.teamName ?
            `You are in ${this.state.config.teamsContext.teamName}` :
            `You are not in a Team`
          }</p>
          <p>Your app is running in the Teams UI</p>
          <p>Your short message is {this.state.config.shortMessage}</p>
          <p></p>
        </div>
      );
      }
  }

}

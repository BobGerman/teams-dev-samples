import React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { ConfigService } from '../services/ConfigService';

export interface ITabConfigPageProps { };
export interface ITabConfigPageState {
  tabName: string;
  shortMessage: string;
  firstRun: boolean;
}
export default class TabConfig extends React.Component<ITabConfigPageProps, ITabConfigPageState> {

  constructor(props: ITabConfigPageProps) {
    super(props);
    this.state = {
      tabName: "My Tab",
      shortMessage: "",
      firstRun: false
    }
  }

  async componentDidMount() {
    const configInfo = await ConfigService.getConfigInfo();
    this.setState({
      shortMessage: configInfo.shortMessage,
      firstRun: !configInfo.shortMessage
    });
    microsoftTeams.appInitialization.notifySuccess();
  }

  render() {

    // Assume the tab is hosted in the same location as this page
    const baseUrl = window.location.href.substring(0, window.location.href.indexOf("#/") + 1);

    microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {

      // This will run when the user clicks the "Save" button in the Teams config popup.
      microsoftTeams.settings.setSettings({
        suggestedDisplayName: this.state.tabName,
        entityId: ConfigService.getEntityId({
          shortMessage: this.state.shortMessage
        }),
        contentUrl: `${baseUrl}/Tab`,
        websiteUrl: `${baseUrl}/Web`
      });
      saveEvent.notifySuccess();

    });

    this.checkValidityState();
    return (
      <div>
        <h1>Tab Configuration</h1>
        <table>
          { this.state.firstRun ?
            <tr>
              <td>Tab name: </td>
              <td>
                <input onChange={this.handleTabNameChange.bind(this)}
                  value={this.state.tabName}
                />
              </td>
            </tr>
            : null }
          <tr>
            <td>Short message: </td>
            <td>
              <input onChange={this.handleShortMessageChange.bind(this)}
                value={this.state.shortMessage}
              />
            </td>
          </tr>
        </table>
      </div>
    );
  }

  private handleTabNameChange(event: React.FormEvent<HTMLInputElement>) {
    this.setState({ tabName: event.currentTarget.value });
    this.checkValidityState();
  }

  private handleShortMessageChange(event: React.FormEvent<HTMLInputElement>) {
    this.setState({ shortMessage: event.currentTarget.value });
    this.checkValidityState();
  }

  private checkValidityState() {
    let result = false;
    if (this.state.tabName && this.state.shortMessage) {
      result = true;
    }
    microsoftTeams.settings.setValidityState(result);
  }

}

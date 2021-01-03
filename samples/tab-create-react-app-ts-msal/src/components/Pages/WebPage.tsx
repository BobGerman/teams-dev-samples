import React from 'react';
import AuthService from '../../services/AuthService/MsalAuthService';
import * as MicrosoftGraphClient from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

/**
 * The web UI used when Teams pops out a browser window
 */
export interface IWebPageProps { }
export interface IWebPageState {
  accessToken: string;
  messages: MicrosoftGraph.Message[];
  error: string;
}
export default class WebPage extends React.Component<IWebPageProps, IWebPageState> {

  constructor(props: IWebPageProps) {
    super(props);
    this.state = {
      accessToken: "",
      messages: [],
      error: ""
    }
  }

  async componentDidMount() {

    this.getMessages();

  }

  render() {

    let key = 0;
    return (
      <div>
        <h1>{process.env.REACT_APP_MANIFEST_NAME}</h1>
        <p>Version {process.env.REACT_APP_MANIFEST_APP_VERSION}</p>
        <p>Your app is running in a stand-alone web page</p>
        <p>Your short message is not available</p>

        <p>Username: {AuthService.getUsername()}</p>
        <ol>
          {
            this.state.messages.map(message => (
              <li key={key++}>EMAIL: {message.receivedDateTime}<br />{message.subject}
              </li>
            ))
          }
        </ol>


      </div>
    );
  }

  private msGraphClient?: MicrosoftGraphClient.Client;
  private async getMessages(): Promise<MicrosoftGraph.Message[]> {

    let result: MicrosoftGraph.Message[] = [];

    if (!this.msGraphClient) {

      // Set up the Graph client
      let scopes = process.env.REACT_APP_AAD_GRAPH_DELEGATED_SCOPES?.split(',') || [];

      // Ensure we are logged in
      if (!AuthService.isLoggedIn()) {

        await AuthService.login(scopes);

      }

      // Initialize a new Graph client
      this.msGraphClient = MicrosoftGraphClient.Client.init({

        authProvider: async (done: MicrosoftGraphClient.AuthProviderCallback) => {
          if (!this.state.accessToken) {
            // Might redirect the browser and not return; will redirect back when done
            const token = await AuthService.getAccessToken(scopes);
            this.setState({
              accessToken: token
            });
          }
          done(null, this.state.accessToken);
        }
      });
    }

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

    return result;
  }

}

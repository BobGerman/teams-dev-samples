import React from 'react';
import IAuthService from '../../services/AuthService/IAuthService';
import AuthService from '../../services/AuthService/MsalAuthService';
import * as MicrosoftGraphClient from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

/**
 * The web UI used when Teams pops out a browser window
 */
export interface IWebPageProps { }
export interface IWebPageState {
  messages: MicrosoftGraph.Message[];
  error: string;
}
export default class WebPage extends React.Component<IWebPageProps, IWebPageState> {

  constructor(props: IWebPageProps) {
    super(props);
    this.state = {
      messages: [],
      error: ""
    }
  }

  async componentDidMount() {

    await this.getMessages();

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

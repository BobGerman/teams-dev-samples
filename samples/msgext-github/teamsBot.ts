import {
  TeamsActivityHandler,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
} from "botbuilder";
import issuesME from "./messageExtensions/issuesME";
import pullRequestME from "./messageExtensions/pullRequestME";

export interface DataInterface {
  likeCount: number;
}

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {

    switch (query.parameters[0].name) {
      case issuesME.ME_NAME:
        return await issuesME.handleTeamsMessagingExtensionQuery(context, query);
      case pullRequestME.ME_NAME:
        return await pullRequestME.handleTeamsMessagingExtensionQuery(context, query);
      default:
        return null;
    }
  }

  public async handleTeamsMessagingExtensionSelectItem(
    context: TurnContext,
    item: any
  ): Promise<MessagingExtensionResponse> {
    switch (item.queryType) {
      case issuesME.ME_NAME:
        return await issuesME.handleTeamsMessagingExtensionSelectItem(context, item); 
      case pullRequestME.ME_NAME:
        return await pullRequestME.handleTeamsMessagingExtensionSelectItem(context, item);
      default:
        return null;
    }
  }
}

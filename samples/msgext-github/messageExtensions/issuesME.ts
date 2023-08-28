import { default as axios } from "axios";
import * as ACData from "adaptivecards-templating";
import { 
    CardFactory,
    TurnContext,
    MessagingExtensionQuery,
    MessagingExtensionResponse
} from "botbuilder";
import GithubIssue from "../model/githubIssue";

class IssuesME {

    public readonly ME_NAME = "issuesQuery";

    // Get suppliers given a query
    async handleTeamsMessagingExtensionQuery (
        context: TurnContext,
        query: MessagingExtensionQuery
        ): Promise<MessagingExtensionResponse> {

        try {
            const response = await axios.get(
                `https://api.github.com/repos/pnp/teams-dev-samples/issues`
            );

            const attachments = [];
            const results = response.data.filter(i => i.title.toLowerCase().includes(query.parameters[0].value.toLowerCase()));
            results.forEach((issue: GithubIssue) => {

                const itemAttachment = CardFactory.heroCard(issue.title);
                let previewAttachment = CardFactory.thumbnailCard(issue.title);
                previewAttachment.content.tap = {
                    type: "invoke",
                    value: {    // Values passed to selectItem when an item is selected
                        queryType: this.ME_NAME,
                        githubIssue: issue
                    }
                };
                const attachment = { ...itemAttachment, preview: previewAttachment };
                attachments.push(attachment);
            });

            return {
                composeExtension: {
                    type: "result",
                    attachmentLayout: "list",
                    attachments: attachments,
                }
            };

        } catch (error) {
            console.log(error);
        }
    };

    async handleTeamsMessagingExtensionSelectItem (context: TurnContext, selectedValue: GithubIssue): Promise<MessagingExtensionResponse>  {

        const templateJson = require('./issuesCard.json');
        const template = new ACData.Template(templateJson);
        const card = template.expand({
            $root: selectedValue
        });

        const resultCard = CardFactory.adaptiveCard(card);

        return {
            composeExtension: {
                type: "result",
                attachmentLayout: "list",
                attachments: [resultCard]
            }
        };

    };

}

export default new IssuesME();
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
                let previewAttachment = CardFactory.thumbnailCard(issue.title, [issue.user.avatar_url]);

                // Clean up the value for presentation
                issue.created_at = new Date(issue.created_at).toLocaleDateString();
                issue.updated_at = issue.updated_at ? new Date(issue.updated_at).toLocaleDateString() : "n/a";
                issue.closed_at = issue.closed_at ? new Date(issue.closed_at).toLocaleDateString() : "n/a";
                issue.body = issue.body.length > 100 ? issue.body.substring(0, 100) + "..." : issue.body;

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

    async handleTeamsMessagingExtensionSelectItem (context: TurnContext, selectedValue: any): Promise<MessagingExtensionResponse>  {

        const issue: GithubIssue = selectedValue.githubIssue;
        const templateJson = issue.pull_request ?
            require('./IssuesWithPR.json') :
            require('./issuesCard.json');
        const template = new ACData.Template(templateJson);
        const card = template.expand({
            $root: issue
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
import { default as axios } from "axios";
import * as ACData from "adaptivecards-templating";
import { 
    CardFactory,
    TurnContext,
    MessagingExtensionQuery,
    MessagingExtensionResponse
} from "botbuilder";
import GithubPullRequest from "../model/githubPullRequest";

class PullRequestsME {

    public readonly ME_NAME = "pullRequestsQuery";

    // Get suppliers given a query
    async handleTeamsMessagingExtensionQuery (
        context: TurnContext,
        query: MessagingExtensionQuery
        ): Promise<MessagingExtensionResponse> {

        try {
            const response = await axios.get(
                `https://api.github.com/repos/pnp/teams-dev-samples/pulls`
            );

            const attachments = [];
            const results = response.data.filter(i => i.title.toLowerCase().includes(query.parameters[0].value.toLowerCase()));
            results.forEach((pr: GithubPullRequest) => {

                const itemAttachment = CardFactory.heroCard(pr.title);
                let previewAttachment = CardFactory.thumbnailCard(pr.title, [pr.user.avatar_url]);

                // Clean up the value for presentation
                pr.created_at = new Date(pr.created_at).toLocaleDateString();
                pr.updated_at = pr.updated_at ? new Date(pr.updated_at).toLocaleDateString() : "n/a";
                pr.closed_at = pr.closed_at ? new Date(pr.closed_at).toLocaleDateString() : "n/a";
                pr.body = pr.body.length > 100 ? pr.body.substring(0, 100) + "..." : pr.body;

                previewAttachment.content.tap = {
                    type: "invoke",
                    value: {    // Values passed to selectItem when an item is selected
                        queryType: this.ME_NAME,
                        githubPullRequest: pr
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

        const pullRequest = selectedValue.githubPullRequest;
        const templateJson = require('./pullRequestCard.json');
        const template = new ACData.Template(templateJson);
        const card = template.expand({
            $root: pullRequest
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

export default new PullRequestsME();
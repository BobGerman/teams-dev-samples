import { default as axios } from "axios";
import ACData from "adaptivecards-templating";
import { 
    CardFactory,
    TurnContext,
    MessagingExtensionQuery,
    MessagingExtensionResponse
} from "botbuilder";
import GithubIssueResponse from "../model/githubIssueResponse";

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
            results.forEach((issue: GithubIssueResponse) => {

                const itemAttachment = CardFactory.heroCard(issue.title);
                let previewAttachment = CardFactory.thumbnailCard(issue.title);
                previewAttachment.content.tap = {
                    type: "invoke",
                    value: {    // Values passed to selectItem when an item is selected
                        queryType: this.ME_NAME,
                        id: issue.id,
                        url: issue.url,

                        SupplierID: 1,
                        flagUrl: "",
                        imageUrl: "",
                        Address:  "",
                        City:  "",
                        CompanyName: issue.title,
                        ContactName: "",
                        ContactTitle: "",
                        Country: "",
                        Fax: "",
                        Phone: "",
                        PostalCode: "",
                        Region: ""
                    },
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

    async handleTeamsMessagingExtensionSelectItem (context, selectedValue): Promise<MessagingExtensionResponse>  {

        // Read card from JSON file
        const templateJson = require('../cards/supplierCard.json');
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
import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/spaceLaunchesMessageExtension/config.html")
export default class SpaceLaunchesMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: "Headline"
                    },
                    {
                        type: "TextBlock",
                        text: "Description"
                    },
                    {
                        type: "Image",
                        url: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`
                    },
                    {
                        type: "ActionSet",
                        actions: [
                            {
                                type: "Action.Execute",
                                title: "More details",
                                data: {
                                    action: "moreDetails",
                                    id: "1234-5678"
                                },
                                fallback: "Action.Submit"
                            }
                        ]
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.4"
            });
        const preview = {
            contentType: "application/vnd.microsoft.card.thumbnail",
            content: {
                title: "Headline",
                text: "Description",
                images: [
                    {
                        url: `https://${process.env.PUBLIC_HOSTNAME}/assets/icon.png`
                    }
                ]
            }
        };

        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    { ...card, preview }
                ]
            } as MessagingExtensionResult);
        } else {
            // the rest
            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    { ...card, preview }
                ]
            } as MessagingExtensionResult);
        }
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        // Handle the Action.Submit action on the adaptive card
        if (value.action === "moreDetails") {
            log(`I got this ${value.id}`);
        }
        return Promise.resolve();
    }

}

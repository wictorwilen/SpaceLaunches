import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import Axios from "axios";

// Initialize debug logging module
const log = debug("msteams");

const getLaunches = (text: string): Promise<any> => {
    return new Promise<any>((resolve, reject) => {
        Axios.get(`https://spacelaunchnow.me/api/3.3.0/launch/upcoming/?search=${text}`, { headers: { accept: "application/json" } }).then(response => {
            resolve(response.data.results);
        });

    });
};

const launchToAttachment = (launch: any): any => {
    const card = CardFactory.adaptiveCard(
        {
            type: "AdaptiveCard",
            body: [
                {
                    type: "TextBlock",
                    size: "Large",
                    text: launch.name
                },
                {
                    type: "TextBlock",
                    text: launch.mission ? launch.mission.description : "TBD"
                },
                {
                    type: "Image",
                    url: launch.image_url ? launch.image_url : `https://${process.env.PUBLIC_HOSTNAME}/assets/icon-color.png`
                },
                {
                    type: "ActionSet",
                    actions: [
                        {
                            type: "Action.OpenUrl",
                            title: "More details",
                            url: launch.url
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
            title: launch.name,
            text: launch.mission ? launch.mission.description : "TBD",
            images: [
                {
                    url: launch.image_url ? launch.image_url : `https://${process.env.PUBLIC_HOSTNAME}/assets/icon-color.png`
                }
            ]
        }
    };
    return { ...card, preview };
};

@PreventIframe("/spaceLaunchesMessageExtension/config.html")
export default class SpaceLaunchesMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {
        let launches = [];
        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run
            launches = await getLaunches("");

        } else {
            launches = await getLaunches(query.parameters![0].value);
        }
        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: launches.map(launch => launchToAttachment(launch))
        } as MessagingExtensionResult);
    }


}

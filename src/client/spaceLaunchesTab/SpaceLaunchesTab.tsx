import * as React from "react";
import { Provider, Flex, Text, Grid, Header, Card, CardHeader, CardBody, Image } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app } from "@microsoft/teams-js";
import Axios from "axios";

/**
 * Implementation of the Space Launces content page
 */
export const SpaceLaunchesTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [launches, setLaunches] = useState<any[]>([]);
    const [host, setHost] = useState<string>("");

    useEffect(() => {
        if (inTeams === true) {
            app.notifySuccess();
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.page.id);
            setHost(context.app.host.name + "/" + context.app.host.clientType);
        }
    }, [context]);

    useEffect(() => {
        Axios.get("https://spacelaunchnow.me/api/3.3.0/launch/upcoming/", { headers: { accept: "application/json" } }).then(response => {
            setLaunches(response.data.results);
        });
    }, []);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content={`🚀 Upcoming Rocket launches hosted in ${host}`} />
                </Flex.Item>
                <Flex.Item>
                    <Grid columns={4}>
                        {launches.map((launch: any) => {
                            return <Card key={launch.id} elevated fluid>
                                <CardHeader>
                                    <Flex gap="gap.small" >
                                        <Text content="🚀" />
                                        <Text content={launch.name} weight="semibold"/>
                                    </Flex>
                                </CardHeader>
                                <CardBody>
                                    <Flex column gap="gap.small">
                                        <Image src={launch.image_url} style={{ height: "100px", objectFit: "cover" }} />
                                        <Text content={launch.mission ? launch.mission.description : "TBD"} weight="light"/>
                                    </Flex>
                                </CardBody>
                            </Card>;
                        })}
                    </Grid>

                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}><div>
                        <Text size="smaller" >
                            (C) Copyright Wictor Wilén<br/>
                            <a href="https://www.flaticon.com/free-icons/rocket" title="rocket icons">Rocket icons created by Freepik - Flaticon</a>
                        </Text>
                    </div>
                </Flex.Item>
            </Flex>
        </Provider >
    );
};

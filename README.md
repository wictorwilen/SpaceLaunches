# Space Launches - Demo of a multi-host Teams

This demo shows a personal tab and a messaging extension that can be used in Microsoft Teams, Office.com and Outlook.

This project is built using the **[yo teams](https://aka.ms/yoteams)** application generator that creates Node.js based Teams, Office and Outlook applications.

## Configure the example

In order to run this example you need to

1. Register a new *Azure bot* service (choose single or multi-tenant)
2. Add a new secret to the application for the bot
3. Create a `.env` file in the project root folder with the contents as per below and replace the Bot App Id and Bot App Secret with your values
4. Enable the *Microsoft Teams* and *Outlook* channels for the Azure Bot. (Important - don't forget to add the Outlook channel before you test it in Outlook)


``` env
PUBLIC_HOSTNAME=launches.azurewebsites.net

# Id of the Microsoft Teams application
APPLICATION_ID=b94e79f0-dab0-11ec-a4db-c3d63815d46b
# Package name of the Microsoft Teams application
PACKAGE_NAME=launches

# App Id and App Password for the Bot Framework bot
MICROSOFT_APP_ID=REPLACE_WITH_YOUR_AZURE_BOT_APP_ID
MICROSOFT_APP_PASSWORD=REPLACE_WITH_YOUR_AZURE_BOT_APP_SECRET

# Port for local debugging
PORT=3007

# Debug settings, default logging "msteams"
DEBUG=msteams
```
 

## Run the example

To run this example follow these steps:

1. Type `gulp ngrok-serve` or `gulp ngrok-serve --debug` in your console.
2. Update your Azure bot registration with the temporary ngrok URL printed in the console
3. Then upload the generated package (`/package/launches.zip`) to your Microsoft Teams app catalog
4. Try it out in Teams, Outlook and Office.com and have fun

> NOTE: if you in Outlook on the web get an `errorCode` with the `BotNotProperlyConfigured` value then you either has not registered the Outlook channel or the feature has not yet rolled out to you. Eventually after adding the channel it will work.

**/GLHF WW**
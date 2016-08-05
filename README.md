# OutlookConnectorSDK
Outlook Group Connector SDK simplifies your code in C# to send comprehensive Json payload (Canvas) that contains activities, tables, images and action links.

This solution contains the following:

## 1) CloudServices Folder
This folder contains webrole in a cloud service that sends messages to an outlook group for a configured webhookurl.
### Project Names: 
    a) OGCCloudSvc: Cloud service project to be deployed to Azure (if needed) or run locally.
    b) OGCWorkerRole: worker role project to sends messages to outlook group. 
    This worker role is a sample app in case you want to have a continuously running backend service 
    that pushes messages to outlook groups based on your logic in the worker role.

## 2) Microservices Folder (Using service fabric)
This folder contains service fabric web api that sends a message to an outlook group for a given webhook url.
### Project Names:
    a) OGCSFApp: Service Fabric Application project.
    b) OGCWebApi: Service Fabric web api project.


## 3) Office365OutlookConnectorSDK
This project contains the source code for published "Outlook365OutlookConnectorSDK" NuGet package.

## 4) webapp45: 
A sample asp.net web application runs on .NET 4.5 and uses the SDK NuGet package.

## 5) OGCSDKConsoleClient: 
A sample console application that uses the SDK NuGet package.


** How to use this SDK in your applications: **

1) Install Office365OutlookConnectorSDK nuget package from Visual Studio into your project.
This is the [SDK NuGet package](https://www.nuget.org/packages/Office365ConnectorSDK/) download link.

2) Here is how to construct a comprehensive message with images, tables and activity details and send it to an outlook group:

                '''
                // activity section
                Section activitySection = new Section()
                {
                    activityTitle = "mostafa elzoghbi commented",
                    activitySubtitle = "On Project Office Connector",
                    activityText = "\"Here are the designs docs \"",
                    activityImage = "http://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg"
                };
                Message message = new Message()
                {
                    summary = "This is the subject for the sent message to an outlook group",
                    title = msg
                };
                message.AddSection(section activity);
                message.AddFacts("Facts", facts);
                message.AddImages("Images", images);
                message.AddAction("check details here", "http://mostafaelzoghbi.com");


                var result = await message.Send(webhookUrl);

                '''

This is how the outlook group connector message looks like:

Please let me know if you want to contribute or add features to this sdk.

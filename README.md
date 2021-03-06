# OutlookConnectorSDK
Outlook Group Connector SDK simplifies your code in C# to send comprehensive Json payload (Canvas) that contains activities, tables, images and action links.

This solution contains the following:

## 1) Office 365 Group Connector SDK
This project contains the source code for published [Outlook365OutlookConnectorSDK](https://www.nuget.org/packages/Office365ConnectorSDK/) NuGet package.

## 2) Console App: 
A sample console application project (OGCSDKConsoleClient) that uses the SDK NuGet package.

## 4) ASP.NET 4.5 Web App: 
A sample ASP.NET web application runs on .NET 4.5 (webapp45) and uses the SDK NuGet package.

## 3) Cloud Services Folder
This folder contains webrole in a cloud service that sends messages to an outlook group for a configured webhookurl.
### Project Names: 
    a) OGCCloudSvc: Cloud service project to be deployed to Azure (if needed) or run locally.
    b) OGCWorkerRole: worker role project to sends messages to outlook group. 
    This worker role is a sample app in case you want to have a continuously running backend service 
    that pushes messages to outlook groups based on your logic in the worker role.

## 4) Microservices Folder (Using Azure Service Fabric)
This folder contains service fabric web api that sends a message to an outlook group for a given webhook url.
### Project Names:
    a) OGCSFApp: Service Fabric Application project.
    b) OGCWebApi: Service Fabric web api project.



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
![ogcmessagescreenshot](https://cloud.githubusercontent.com/assets/11993393/17452149/906877e4-5b3a-11e6-94d8-d28c38fcf663.PNG)

## A live demo for group connectors integration web app
I have built this application that demonstrates outlook connector integration showcase that includes an integration for "Connect to Office 365" button in a third party website and how to send a detailed canvas messages to a group.

### How to Use it:

•Outlook Connector landing page: Click on "Enterprise" menu item, install our connector into one of your office 365 groups.
•Send a message to any group: Click on "Send Message" menu item, set a title message and group name and click on Send button. Check your group and you will be notified with a full detailed canvas message.

Application Url: http://outlookconnectorwebapp.azurewebsites.net/


Please let me know if you want to contribute or add features to this sdk.

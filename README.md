# OutlookConnectorSDK
Outlook Group Connector SDK simplifies your code in C# to send comprehensive Json payload (Canvas) that contains activities, tables, images and action links.

This SDK has been packaged and deployed to NuGet website, here is the link for the nuget package: 
https://www.nuget.org/packages/Office365ConnectorSDK/

How to use this SDK in your applications:

1) Install Office365OutlookConnectorSDK nuget package from Visual Studio into your project.

2) Here is how to construct a sample comprehensive message with images, tables and activity details:


                Message message = new Message()
                {
                    summary = "This is the subject for the sent message to an outlook group",
                    title = msg
                };
                message.AddSection(nSec1);
                message.AddFacts("Facts", facts);
                message.AddImages("Images", images);
                message.AddAction("check details here", "http://mostafaelzoghbi.com");


                var result = await message.Send(webhookUrl);


Please let me know if you want to contribute or add features to this sdk.

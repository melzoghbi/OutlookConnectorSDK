using Office365ConnectorSDK;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OGCSDKConsoleClient
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter a message title");
            SendMessage(Console.ReadLine()).Wait();
            Console.WriteLine("A message has been sent successfully!");
            Console.ReadLine();
        }

        static async Task<string> SendMessage(string msg)
        {
            string webhookUrl = ConfigurationManager.AppSettings["webhookUrl"];
            // section
            Section nSec1 = new Section()
            {
                activityTitle = "mostafa elzoghbi commented",
                activitySubtitle = "On Project Office Connector",
                activityText = "\"Here are the designs docs \"",
                activityImage = "http://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg"
            };
            // facts table
            List<Fact> facts = new List<Fact>() {
                    new Fact("Labels", "Designs, redlines"),
                    new Fact("Due date", "July 13, 2016"),
                    new Fact("Attachments", "[final.jpg](http://connectorsdemo.azurewebsites.net/images/WIN14_Jan_04.jpg)")
            };
            // images list
            List<Image> images = new List<Image>() {
                new Image("http://connectorsdemo.azurewebsites.net/images/MicrosoftSurface_024_Cafe_OH-06315_VS_R1c.jpg"),
                new Image("http://connectorsdemo.azurewebsites.net/images/WIN12_Scene_01.jpg"),
                new Image("http://connectorsdemo.azurewebsites.net/images/WIN12_Anthony_02.jpg")

            };

            if (!string.IsNullOrEmpty(webhookUrl))
            {
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

                return "A message has been sent!";
            }
            else
                return "Please set webhook url in the config!";
        }
    }
}

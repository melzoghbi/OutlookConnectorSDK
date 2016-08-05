using Microsoft.Owin;
using Office365ConnectorSDK;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;
using System.Web.Http;

namespace OGCWebApi.Controllers
{
    public class ValuesController : ApiController
    {
        [HttpGet]
        public async Task<string> SendOutlookGroupMessage()
        {
            // Get webhook url for the selected group
           string webhookUrl = ConfigurationManager.AppSettings["webhookUrl"];

            return await SendOutlookGroupMessage(webhookUrl);
        }
        /// <summary>
        /// This uses NuGet SDK v1.1. A simpler implementation for sending messages to outlook connectors
        /// </summary>
        /// <param name="webhookUrl"></param>
        /// <returns></returns>
        [HttpGet]
        public async Task<string> SendOutlookGroupMessage(string webhookUrl)
        {
          
            // section
            Section activitySection = new Section()
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
                    summary = "Mostafa commented on a post",
                    title = "Service Fabric web api has sent this message"
                };

                message.AddSection(activitySection);
                message.AddFacts("facts", facts);
                message.AddImages("images", images);
                message.AddAction("check details here", "http://mostafaelzoghbi.com");


                var result = await message.Send(webhookUrl);

                return "This message has been sent to the specified office 365 group";
            }
            else
                return "This group doesn't exist";

        }

      
    }
}

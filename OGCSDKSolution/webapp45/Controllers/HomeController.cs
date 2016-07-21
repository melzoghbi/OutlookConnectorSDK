using Newtonsoft.Json.Linq;
using Office365ConnectorSDK;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace webapp45.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Redirect(string state, string webhook_url, string group_name)
        {
            string responseString = string.Empty;


            var restClient = new RestClient(webhook_url);
            var restRequest = new RestRequest(Method.POST);
            // add header format json
            restRequest.AddHeader("Content-Type", "application/json");
            restRequest.RequestFormat = DataFormat.Json;

            #region one way to add simple body as in anonymous object
            restRequest.AddBody(new
            {
                title = "Thank you for installing our app",
                text = "Welcome to our community! we will keep you posted with up to date information for our platform",
                themeColor = "#3479BF"
            });
            #endregion
                    

            // We need to store this group to push future messages
            DBManager dbMgr = new DBManager();
            dbMgr.AddGroup(group_name, webhook_url);

            var response = restClient.Execute(restRequest);

            ViewData["Message"] = "You have successfully subscribed to our Outlook connector";
            ViewData["Details"] = "A welcome message has been sent to your group, and you will start receiving updates to this group";

            return View();
        }

        public ActionResult Enterprise()
        {
#if DEBUG
            ViewData["url"] = "https://outlook.office.com/connectors/Connect?state=mywebsite&app_id=62e489e3-df21-48f5-8d24-a7388ef2ae3b&callback_url=https://localhost:44397/Home/redirect" ;
#else
            ViewData["url"] = "https://outlook.office.com/connectors/Connect?state=myAppsState&app_id=24fa062c-dd1a-418e-98dc-3818005b32f1&callback_url=http://outlookconnectorwebapp.azurewebsites.net/Home/redirect" ;

#endif
            return View();
        }

        public ActionResult PushMessage()
        {
#if DEBUG
            ViewData["groupName"] = "OnlyMo";
#endif
            return View();
        }
        /// <summary>
        /// This action uses old SDK v1.0 nuget package, check out action PushMessage2
        /// NuGet v1.1 is compatible for both implementation PushMessage & PushMessage2 actions.
        /// </summary>
        /// <param name="formCols"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<ActionResult> PushMessage(FormCollection formCols)
        {
            string responseString = string.Empty;

            // Get webhook url for the selected group
            string groupName = formCols["txtGroupName"];
            DBManager dbMgr = new DBManager();
            string webhookUrl = dbMgr.GetWebhookUrl(groupName);

            string msg = formCols["txtMessage"];

            if (!string.IsNullOrEmpty(webhookUrl))
            {
                Message message = new Message()
                {
                    summary = "Mostafa commented on a post",
                    title = msg,
                    sections = new List<Section>() {
                        new Section() {
                            activityTitle = "mostafa elzoghbi commented",
                            activitySubtitle = "On Project Office Connector",
                            activityText = "\"Here are the designs docs \"",
                            activityImage = "http://connectorsdemo.azurewebsites.net/images/MSC12_Oscar_002.jpg"
                },
                       new Section() {
                title = "Details",
                facts = new List<Fact>() {
                    new Fact("Labels", "Designs, redlines"),
                    new Fact("Due date", "July 13, 2016"),
                    new Fact("Attachments", "[final.jpg](http://connectorsdemo.azurewebsites.net/images/WIN14_Jan_04.jpg)")
            }
        },
                        new Section() {
            title = "Images",
            images = new List<Image>() {
                new Image("http://connectorsdemo.azurewebsites.net/images/MicrosoftSurface_024_Cafe_OH-06315_VS_R1c.jpg"),
                new Image("http://connectorsdemo.azurewebsites.net/images/WIN12_Scene_01.jpg"),
                new Image("http://connectorsdemo.azurewebsites.net/images/WIN12_Anthony_02.jpg")
            }
        }
                    },
                    potentialAction = new List<PotentialAction>() {
        new PotentialAction()
        {
            name = "check details here",
            target = new List<string> { "http://mostafaelzoghbi.com" }
        }
                    }
                };
                var result = await message.Send(webhookUrl);

                ViewData["Message"] = "This message has been sent to the specified office 365 group";
            }
            else
                ViewData["Message"] = "This group doesn't exist";

            return View();
        }
        
        /// <summary>
        /// This uses NuGet SDK v1.1. A simpler implementation for sending messages to outlook connectors
        /// </summary>
        /// <param name="formCols"></param>
        /// <returns></returns>
        [HttpPost]
        public async Task<ActionResult> PushMessage2(FormCollection formCols)
        {
            string responseString = string.Empty;

            // Get webhook url for the selected group
            string groupName = formCols["txtGroupName"];
            DBManager dbMgr = new DBManager();
            string webhookUrl = dbMgr.GetWebhookUrl(groupName);

            string msg = formCols["txtMessage"];

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
                    title = msg
                };

                message.AddSection(activitySection);
                message.AddFacts("facts", facts);
                message.AddImages("images", images);
                message.AddAction("check details here", "http://mostafaelzoghbi.com");
      

        var result = await message.Send(webhookUrl);

                ViewData["Message"] = "This message has been sent to the specified office 365 group";
            }
            else
                ViewData["Message"] = "This group doesn't exist";

            return View("PushMessage");
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
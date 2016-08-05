using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.Diagnostics;
using Microsoft.WindowsAzure.ServiceRuntime;
using Microsoft.WindowsAzure.Storage;
using System.Configuration;
using Office365ConnectorSDK;

namespace OGCWorkerRole
{
    public class WorkerRole : RoleEntryPoint
    {
        private readonly CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        private readonly ManualResetEvent runCompleteEvent = new ManualResetEvent(false);

        public override void Run()
        {
            Trace.TraceInformation("OGCWorkerRole is running");

            try
            {
                this.RunAsync(this.cancellationTokenSource.Token).Wait();
            }
            finally
            {
                this.runCompleteEvent.Set();
            }
        }

        public override bool OnStart()
        {
            // Set the maximum number of concurrent connections
            ServicePointManager.DefaultConnectionLimit = 12;

            // For information on handling configuration changes
            // see the MSDN topic at http://go.microsoft.com/fwlink/?LinkId=166357.

            bool result = base.OnStart();

            Trace.TraceInformation("OGCWorkerRole has been started");

            return result;
        }

        public override void OnStop()
        {
            Trace.TraceInformation("OGCWorkerRole is stopping");

            this.cancellationTokenSource.Cancel();
            this.runCompleteEvent.WaitOne();

            base.OnStop();

            Trace.TraceInformation("OGCWorkerRole has stopped");
        }

        private async Task<bool> RunAsync(CancellationToken cancellationToken)
        {
            // TODO: Replace the following with your own logic.
            while (!cancellationToken.IsCancellationRequested)
            {
                Trace.TraceInformation("OGC worker role is working");

                // you can check groups you stored and send daily, weekly or monthly messages to subscribed outlook groups
                string webhookUrl = ConfigurationManager.AppSettings["webhookUrl"];
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
                        summary = "This is the subject for the sent message to an outlook group",
                        title = "This is a message that has been sent from a worker role"
                    };
                    message.AddSection(activitySection);
                    message.AddFacts("Facts", facts);
                    message.AddImages("Images", images);
                    message.AddAction("check details here", "http://mostafaelzoghbi.com");


                    var result = await message.Send(webhookUrl);
                }
                else
                    return Task.FromResult<bool>(false).Result;

            }
            return Task.FromResult<bool>(true).Result;
        }
    }
}

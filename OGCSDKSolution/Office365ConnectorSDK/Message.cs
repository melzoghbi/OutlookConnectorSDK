using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Office365ConnectorSDK
{
    public class Message
    {
        public string summary { get; set; }
        public string text { get; set; }
        public string title { get; set; }
        public string themeColor { get; set; }
        public List<Section> sections { get; set; }
        public List<PotentialAction> potentialAction { get; set; }

        public string ToJson()
        {
            return JsonConvert.SerializeObject(this, new JsonSerializerSettings() { NullValueHandling = NullValueHandling.Ignore });
        }
        public async Task<bool> Send(string webhook_uri)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var content = new StringContent(this.ToJson(), System.Text.Encoding.UTF8, "application/json");
            using (var response = await client.PostAsync(webhook_uri, content))
            {
                return response.IsSuccessStatusCode;
            }
        }

        #region Extended methods  - Section
        public void AddSection(Section section)
        {
            if (this.sections == null)
                this.sections = new List<Section>();          

            this.sections.Add(section);
        }
        #endregion

        #region Extended methods - Activity
        public void AddActivity(string activityTitle, string activitySubTitle, string activityText, string activityImageUrl)
        {
            if (sections == null)
                sections = new List<Section>();

            Section newSection = new Section();
            newSection.activityTitle = activityTitle;
            newSection.activitySubtitle = activitySubTitle;
            newSection.activityText = activityText;
            newSection.activityImage = activityImageUrl;
            sections.Add(newSection);

            this.sections.Add(newSection);
        }
        #endregion

        #region Extended methods  - Facts Table
        public void AddFacts(string title, List<Fact> facts)
        {
            if (sections == null)
                sections = new List<Section>();
            if (facts == null)
                return;

            Section newSection = new Section();
            newSection.title = title;
            newSection.facts = new List<Fact>();
            newSection.facts.AddRange(facts);

            this.sections.Add(newSection);
        }
        #endregion
        
        #region Extended methods  - Images
        public void AddImages(string title, List<Image> images)
        {
            if (sections == null)
                sections = new List<Section>();
            if (images == null)
                return;

            Section newSection = new Section();
            newSection.title = title;
            newSection.images = new List<Image>();
            newSection.images.AddRange(images);

            this.sections.Add(newSection);
        }
        public void AddImages(string title, List<string> imageUrls)
        {
            if (sections == null)
                sections = new List<Section>();
            if (imageUrls == null)
                return;

            Section newSection = new Section();
            newSection.title = title;
            foreach (var item in imageUrls)
            {
                newSection.images.Add(new Image(item));
            }

            this.sections.Add(newSection);

        }
        #endregion

        #region Extended methods  - Actions
        public void AddAction(string actionName, string targetUrl)
        {
            if (potentialAction == null)
                potentialAction = new List<PotentialAction>();

            this.potentialAction.Add(new PotentialAction(actionName, targetUrl));
        }
        #endregion

       
    }
}

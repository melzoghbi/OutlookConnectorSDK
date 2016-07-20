using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365ConnectorSDK
{
    public class PotentialAction
    {
        public PotentialAction()
        {
            context = "http://schema.org";
            type = "ViewAction";
        }
        public PotentialAction(string name, string target)
        {
            context = "http://schema.org";
            type = "ViewAction";
            this.name = name;
            this.target = new List<string> { target };
        }
        [JsonProperty(PropertyName = "@context")]
        public string context { get; set; }
        [JsonProperty(PropertyName = "@type")]
        public string type { get; set; }
        public string name { get; set; }
        public List<string> target { get; set; }
    }
}

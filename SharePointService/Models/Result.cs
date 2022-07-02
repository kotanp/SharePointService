using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SharePointService.Models
{
    public class Result
    {
        [JsonProperty(PropertyName = "sharingLinkWrite")]
        public string SharingLinkWrite { get; set; }

        [JsonProperty(PropertyName = "sharingLinkRead")]
        public string SharingLinkRead { get; set; }
    }
}

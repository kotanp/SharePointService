using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SharePointService.Models
{
    public class Settings
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string Scopes { get; set; }
        public string BaseUrl { get; set; }
        public string TokenEndPoint { get; set; }
        public string SiteUrl { get; set; }
        public string FilePath { get; set; }
    }
}

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SharePointService.Models
{
    public class PdfResult
    {
        [JsonProperty(PropertyName = "pdfBytes")]
        public string pdfBytes { get; set; }
    }
}

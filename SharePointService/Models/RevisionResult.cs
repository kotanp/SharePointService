using Newtonsoft.Json;

namespace SharePointService.Models
{
    /**
     * Revision model
     */
    public class RevisionResult
    {
        [JsonProperty(PropertyName = "revisionCount")]
        public int RevisionCount { get; set; }
    }
}

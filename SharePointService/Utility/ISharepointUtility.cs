using Microsoft.Graph;
using SharePointService.Models;

namespace SharePointService.Utility
{
    public interface ISharepointUtility
    {
        /**
         * Downloads sharepoint file, then stores it into model class
         */
        public SharepointItem downloadSharepointItem(string sharepointUrl, GraphServiceClient client, string sharedItemId);
    }
}

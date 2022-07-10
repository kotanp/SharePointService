using Microsoft.Graph;
using SharePointService.Models;

namespace SharePointService.Utility
{
    public interface ISharepointUtility
    {
        /**
         * Downloads sharepoint file, then stores it into model class
         */
        public SharepointItem DownloadSharepointItem(string sharepointUrl, GraphServiceClient client, string sharedItemId);

        /**
         * Extracts the sharepoint file's extension
         */
        public string GetFileExtensionOfSharepointItem(SharepointItem sharepointItem);

        /**
         * Returns true if the fileExtension is xls(x)
         */
        public bool IsExtensionXlsx(string fileExtension);

        /**
         * Returns true if the fileExtension is doc(x)
         */
        public bool IsExtensionDocx(string fileExtension);
    }
}

using Microsoft.Graph;
using SharePointService.Models;
using System.Collections.Generic;
using System.Net.Http;

namespace SharePointService.Utility
{
    public class SharepointUtility : ISharepointUtility
    {
        public SharepointItem DownloadSharepointItem(string sharepointUrl, GraphServiceClient client, string sharedItemId)
        {
            SharepointItem sharepointItem = new SharepointItem();
            var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("format", "pdf")
                    };
            var name = client.Shares[sharedItemId].DriveItem.Request().GetAsync().GetAwaiter().GetResult();
            sharepointItem.Name = name.Name;
            var requestUrl = $"{client.BaseUrl}/shares/{sharedItemId}/driveitem/content";
            var message = new HttpRequestMessage(HttpMethod.Get, requestUrl);
            client.AuthenticationProvider.AuthenticateRequestAsync(message);
            var response = client.HttpProvider.SendAsync(message).GetAwaiter().GetResult();
            byte[] downloadedByteArray = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
            sharepointItem.Data = downloadedByteArray;
            return sharepointItem;
        }

        public string GetFileExtensionOfSharepointItem(SharepointItem sharepointItem)
        {
            return sharepointItem.Name.Substring(sharepointItem.Name.IndexOf('.') + 1);
        }

        public bool IsExtensionDocx(string fileExtension)
        {
            return fileExtension.Equals("docx") || fileExtension.Equals("doc") || fileExtension.Equals(".docx") || fileExtension.Equals(".doc");
        }

        public bool IsExtensionXlsx(string fileExtension)
        {
            return fileExtension.Equals("xlsx") || fileExtension.Equals("xls") || fileExtension.Equals(".xlsx") || fileExtension.Equals(".xls");
        }
    }
}

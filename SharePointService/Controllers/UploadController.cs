using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using SharePointService.Models;
using Microsoft.Extensions.Options;
using System.Net.Http.Headers;
using Microsoft.Extensions.Logging;
using System.IO;
using Newtonsoft.Json;
using System.Diagnostics;

namespace SharePointService.Controllers
{
    public class UploadController : Controller
    {
        private readonly Settings _settings;
        private readonly ILogger _logger;
        private GraphServiceClient client;
        //private FolderStructure _folderstructure;
        //private DataModel data;
        public UploadController(IOptions<Settings> options, ILoggerFactory logFactory)
        {
            _settings = options.Value;
            _logger = logFactory.CreateLogger<UploadController>();
            client = GraphClient(_settings.ClientId, _settings.ClientSecret, new[] { _settings.Scopes }, _settings.BaseUrl, _settings.TokenEndPoint);
        }
        [HttpPost]
        public IActionResult Index(string originalFileName, string filePath, string uuid)
        {
            // string baseUrl = _settings.BaseUrl;
            // var clientId = _settings.ClientId;
            // var clientSecret = _settings.ClientSecret;
            // var scopes = new[] { _settings.Scopes };
            // string settingsPath = _settings.FilePath;
            // string tokenEndpoint = _settings.TokenEndPoint;
            string siteUrl = _settings.SiteUrl;
            // GraphServiceClient client = GraphClient(clientId, clientSecret, scopes, baseUrl, tokenEndpoint);
            var idCollection = client.Sites.Request().GetAsync().GetAwaiter().GetResult();
            var siteId = idCollection.FirstOrDefault(x => x.WebUrl == siteUrl)?.Id;
            string[] strucutre = filePath.Split('/');
            string filename = originalFileName;
            string itempath = "";
            for (int i = 4; i < strucutre.Length - 1; i++)
            {
                itempath += "/" + strucutre[i];
            }
            byte[] byteArray;
            using (var stream = new MemoryStream())
            {
                Request.Body.CopyToAsync(stream).GetAwaiter().GetResult();
                byteArray = stream.ToArray();
            }
            CreateFolder(client, siteId, _logger, strucutre[3]).Wait(TimeSpan.FromSeconds(10));
            UploadFile(client, siteId, _logger, byteArray, strucutre[3], itempath.Substring(1), filename, uuid);
            var links = CreateOrganizationSharingLink(client, siteId, strucutre[3], itempath, filename);
            //string sharinglink = link.Result;
            //byte[] pdfbytes = ConvertToPdf(client, link.Result);
            //var intArray = pdfbytes.Select(b => (int)b).ToArray();
            //var uuid = columnDefinition(client, links.Result[0]);
            Result result = new Result();
            result.SharingLinkWrite = links.Result[0];
            result.SharingLinkRead = links.Result[1];
            //result.UUID = uuid;

            var json = JsonConvert.SerializeObject(result);
            return Content(json, "application/json");
        }

        [HttpDelete]
        public IActionResult Delete(string fileFullUrl)
        {
            var sharedItemId = UrlToSharingToken(fileFullUrl);
            client.Shares[sharedItemId].DriveItem.Request().DeleteAsync();
            return Content("OK");
        }

        [HttpGet]
        public IActionResult ConvertToPdf(string fileFullUrl)
        {
            var sharedItemId = UrlToSharingToken(fileFullUrl);
            var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("format", "pdf")
                    };

            Stream driveItem = client.Shares[sharedItemId].DriveItem.Content.Request(queryOptions).GetAsync().GetAwaiter().GetResult();
            byte[] byteArray;
            using (var memoryStream = new MemoryStream())
            {
                driveItem.CopyTo(memoryStream);
                byteArray = memoryStream.ToArray();
            }
            var intArray = byteArray.Select(b => (int)b).ToArray();
            PdfResult result = new PdfResult();
            result.pdfBytes = String.Join(" ", intArray);
            var json = JsonConvert.SerializeObject(result);
            return Content(json, "application/json");
        }

        public static string columnDefinition(GraphServiceClient client, string fileFullUrl)
        {
            var sharedItemId = UrlToSharingToken(fileFullUrl);
            var driveItem3 = client.Shares[sharedItemId].DriveItem.ListItem.Fields.Request().Select("UUID").GetAsync().GetAwaiter().GetResult();
            string uuid = "";
            foreach (var item in driveItem3.AdditionalData)
            {
                uuid = item.Value.ToString();
            }
            return uuid;
        }

        public static byte[] ConvertToPdf(GraphServiceClient client, string fileFullUrl)
        {

            var sharedItemId = UrlToSharingToken(fileFullUrl);
            //var name = client.Shares[sharedItemId].DriveItem.Request().GetAsync().GetAwaiter().GetResult();
            var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("format", "pdf")
                    };

            Stream driveItem = client.Shares[sharedItemId].DriveItem.Content.Request(queryOptions).GetAsync().GetAwaiter().GetResult();

            using (var memoryStream = new MemoryStream())
            {
                driveItem.CopyTo(memoryStream);
                var byteArray = memoryStream.ToArray();
                return byteArray;
            }

        }

        static string UrlToSharingToken(string inputUrl)
        {
            string sharingUrl = inputUrl;
            string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(sharingUrl));
            string encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
            return encodedUrl;
        }


        private static void UploadFile(GraphServiceClient client, string siteid, ILogger logger, byte[] byteArray, string listname, string itempath, string filename, string uuid)
        {
            var drives = client.Sites[siteid].Drives.Request().GetAsync().GetAwaiter().GetResult();
            var driveId = drives.FirstOrDefault(x => x.Name == listname)?.Id;;
            var stream = new MemoryStream(byteArray);
            if (!String.IsNullOrEmpty(driveId))
            {
                try
                {
                    client.Sites[siteid].Drives[driveId].Root.ItemWithPath(itempath + "/" + filename).Content.Request().PutAsync<DriveItem>(stream).Wait(5000);
                }
                catch (Exception se)
                {
                    logger.LogInformation("File feltöltési hiba: {0}", se);
                }
                var fieldValueSet = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"UUID", uuid}
                    }
                };
                try
                {
                    client.Sites[siteid].Drives[driveId].Root.ItemWithPath(itempath + "/" + filename).ListItem.Fields.Request().UpdateAsync(fieldValueSet).Wait(5000);   
                }
                catch (Exception e)
                {
                    logger.LogInformation("File paraméter frissítési hiba: {0}", e);
                }
            }
        }


        public static async Task CreateFolder(GraphServiceClient client, string siteid, ILogger logger, string listname)
        {
            var list = new Microsoft.Graph.List
            {
                Columns = new ListColumnsCollectionPage()
                {
                    new ColumnDefinition
                    {
                        Name = "UUID",
                        Text = new TextColumn
                        {
                        },
                        IsDeletable = false,
                        Hidden  = true
                    }
                },
                DisplayName = listname,
                ListInfo = new ListInfo
                {
                    Hidden = false,
                    ContentTypesEnabled = false,
                    Template = "documentLibrary"
                }
            };
            try
            {
                await client.Sites[siteid].Lists.Request().AddAsync(list);
            }
            catch (ServiceException se)
            {
                if (se.Error.Code != "nameAlreadyExists")
                {
                    logger.LogInformation("Főmappa létrehozási hiba: {0}", se.Message);
                }
            }
        }

        private static async Task<List<string>> CreateOrganizationSharingLink(GraphServiceClient client, string siteid, string listname, string itempath, string filename)
        {
            List<string> sharingLinks = new List<string>();
            var drives = client.Sites[siteid].Drives.Request().GetAsync().GetAwaiter().GetResult();
            var driveId = drives.FirstOrDefault(x => x.Name == listname)?.Id;
            var fileid = client.Sites[siteid].Drives[driveId].Items["root:" + itempath + ":"].Children.Request().Filter($"name eq '" + filename + "'").GetAsync().GetAwaiter().GetResult().Select(x => x.Id).FirstOrDefault();
            //var fileid = fileids.Where(x => x.Name == "teszt.docx").Select(x => x.Id).FirstOrDefault();
            var type = "edit";
            var scope = "organization";
            //var shareinglink = client.Sites[siteid].Drives[driveId].Items[fileid].CreateLink(type, scope, null, null, null).Request().PostAsync().GetAwaiter().GetResult();
            var shareinglinkEdit =  client.Sites[siteid].Drives[driveId].Items[fileid].CreateLink(type, scope, null, null, null).Request().PostAsync().GetAwaiter().GetResult();
            sharingLinks.Add(shareinglinkEdit.Link.WebUrl);
            type = "view";
            var shareinglinkView = client.Sites[siteid].Drives[driveId].Items[fileid].CreateLink(type, scope, null, null, null).Request().PostAsync().GetAwaiter().GetResult();
            sharingLinks.Add(shareinglinkView.Link.WebUrl);
            return sharingLinks;
        }

        public static GraphServiceClient GraphClient(string clientId, string clientSecret, string[] scopes, string baseUrl, string tokenEndpoint)
        {
            Task<string> accessToken = AccessToken(clientId, clientSecret, scopes, tokenEndpoint);
            GraphServiceClient client = new GraphServiceClient(baseUrl, new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken.Result);

                        return Task.FromResult(0);
                    }));
            return client;
        }

        public static async Task<string> AccessToken(string clientId, string clientSecret, string[] scopes, string endpoint)
        {
            IConfidentialClientApplication app;
            app = ConfidentialClientApplicationBuilder.Create(clientId)
                                                      .WithClientSecret(clientSecret)
                                                      .WithAuthority(new Uri(endpoint))
                                                      .Build();
            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            return result.AccessToken;
        }
    }
}

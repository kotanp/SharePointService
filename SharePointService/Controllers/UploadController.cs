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
using SharePointService.Service;
using System.Net.Http;
using SharePointService.Utility;

namespace SharePointService.Controllers
{
    public class UploadController : Controller
    {

        private IConverterService converterService;
        private readonly Settings _settings;
        private readonly ILogger _logger;
        private GraphServiceClient client;
        private ISharepointUtility sharepointUtility;

        public UploadController(IConverterService converterService, ISharepointUtility sharepointUtility,IOptions<Settings> options, ILoggerFactory logFactory)
        {
            this.converterService = converterService;
            this.sharepointUtility = sharepointUtility;
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
            var siteId = idCollection.Where(x => x.WebUrl == siteUrl).FirstOrDefault().Id;
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
            CreateFolder(client, siteId, _logger, strucutre[3]).Wait(TimeSpan.FromSeconds(1));
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
            SharepointItem sharepointItem = sharepointUtility.downloadSharepointItem(fileFullUrl, client, UrlToSharingToken(fileFullUrl));
/*            var sharedItemId = UrlToSharingToken(fileFullUrl);
            var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("format", "pdf")
                    };
            var name = client.Shares[sharedItemId].DriveItem.Request().GetAsync().GetAwaiter().GetResult();
            var requestUrl = $"{client.BaseUrl}/shares/{sharedItemId}/driveitem/content";
            var message = new HttpRequestMessage(HttpMethod.Get, requestUrl);
            client.AuthenticationProvider.AuthenticateRequestAsync(message);
            var response =  client.HttpProvider.SendAsync(message).GetAwaiter().GetResult();
            byte[] downloadedByteArray =  response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();*/
            //Stream driveItem = client.Shares[sharedItemId].DriveItem.Content.Request(queryOptions).GetAsync().GetAwaiter().GetResult();
            //var items = client.Drives[sharedItemId].Items.Request();
            //byte[] byteArray;
            //using (var memoryStream = new MemoryStream())
            //{
            //    driveItem.CopyTo(memoryStream);
            //    byteArray = memoryStream.ToArray();
            //}
            //var intArray = byteArray.Select(b => b).ToArray();
            //PdfResult result = new PdfResult();
            //result.pdfBytes = String.Join(" ", intArray);
            string fileExtension;
            if (!String.IsNullOrEmpty(sharepointItem.Name))
            {
                fileExtension = sharepointItem.Name.Substring(sharepointItem.Name.IndexOf('.') + 1);
            } else
            {
                throw new Exception("File extension cannot be extracted!");
            }
            var json = this.converterService.ConvertToPdf(sharepointItem.Data, fileExtension);
            return Content(json, "application/json");
        }

        [HttpPost]
        public IActionResult TestInterop(string filePath, string fileExtension)
        {
            byte[] doc = System.IO.File.ReadAllBytes(filePath);
            var json = this.converterService.ConvertToPdf(doc, fileExtension);
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


        public static async void UploadFile(GraphServiceClient client, string siteid, ILogger logger, byte[] byteArray, string listname, string itempath, string filename, string uuid)
        {
            var drives = client.Sites[siteid].Drives.Request().GetAsync().GetAwaiter().GetResult();
            var driveId = drives.Where(x => x.Name == listname).FirstOrDefault().Id;;
            var stream = new MemoryStream(byteArray);
            try
            {
                await client.Sites[siteid].Drives[driveId].Root.ItemWithPath(itempath + "/" + filename).Content.Request().PutAsync<DriveItem>(stream);
            }
            catch (ServiceException se)
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
            await client.Sites[siteid].Drives[driveId].Root.ItemWithPath(itempath + "/" + filename).ListItem.Fields.Request().UpdateAsync(fieldValueSet);
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

        public static async Task<List<string>> CreateOrganizationSharingLink(GraphServiceClient client, string siteid, string listname, string itempath, string filename)
        {
            List<string> sharingLinks = new List<string>();
            var drives = client.Sites[siteid].Drives.Request().GetAsync().GetAwaiter().GetResult();
            var driveId = drives.Where(x => x.Name == listname).FirstOrDefault().Id;
            var fileid = client.Sites[siteid].Drives[driveId].Items["root:" + itempath + ":"].Children.Request().Filter($"name eq '" + filename + "'").GetAsync().GetAwaiter().GetResult().Select(x => x.Id).FirstOrDefault();
            //var fileid = fileids.Where(x => x.Name == "teszt.docx").Select(x => x.Id).FirstOrDefault();
            var type = "edit";
            var scope = "organization";
            //var shareinglink = client.Sites[siteid].Drives[driveId].Items[fileid].CreateLink(type, scope, null, null, null).Request().PostAsync().GetAwaiter().GetResult();
            var shareinglinkEdit = await client.Sites[siteid].Drives[driveId].Items[fileid].CreateLink(type, scope, null, null, null).Request().PostAsync();
            sharingLinks.Add(shareinglinkEdit.Link.WebUrl);
            type = "view";
            var shareinglinkView = await client.Sites[siteid].Drives[driveId].Items[fileid].CreateLink(type, scope, null, null, null).Request().PostAsync();
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

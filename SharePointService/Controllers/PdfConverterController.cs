using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using SharePointService.Models;
using System.IO;
using System.Net.Http.Headers;
using Microsoft.Extensions.Options;

namespace SharePointService.Controllers
{
    public class PdfConverterController : Controller
    {
        private readonly Settings _settings;
        public PdfConverterController(IOptions<Settings> options)
        {
            _settings = options.Value;
        }
        public IActionResult Index()
        {
            string baseUrl = _settings.BaseUrl;
            var clientId = _settings.ClientId;
            var clientSecret = _settings.ClientSecret;
            var scopes = new[] { _settings.Scopes };
            string tokenEndpoint = _settings.TokenEndPoint;

            GraphServiceClient client = GraphClient(clientId, clientSecret, scopes, baseUrl, tokenEndpoint);
            ConvertToPdf(client);
            string sharinglink = "https://mikrodat.sharepoint.com/:w:/s/teszt80/EdggayXe5h1BtCuCdD4Zf2gB7gtSVN7FCKA1E7vbtfBCzw?e=uFZZV4";
            byte[] pdfbytes = ConvertToPdf(client);
            var intArray = pdfbytes.Select(b => (int)b).ToArray();
            Result result = new Result();
            //result.SharingLink = sharinglink;
            //result.PdfBytes = intArray;
            var json = JsonConvert.SerializeObject(result);
            return Content(json, "application/json");
        }

        public static byte[] ConvertToPdf(GraphServiceClient client)
        {
            string fileFullUrl = "https://mikrodat.sharepoint.com/:w:/s/teszt80/EdggayXe5h1BtCuCdD4Zf2gB7gtSVN7FCKA1E7vbtfBCzw?e=uFZZV4";

            var sharedItemId = UrlToSharingToken(fileFullUrl);
            //var name = client.Shares[sharedItemId].DriveItem.Request().GetAsync().GetAwaiter().GetResult();

            var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("format", "pdf")
                    };

            Stream driveItem = client.Shares[sharedItemId].DriveItem.Content.Request(queryOptions).GetAsync().GetAwaiter().GetResult();
            //using (var fileStream = new FileStream(filePath, FileMode.Create))
            //{
            //    driveItem.CopyTo(fileStream);
            //}
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

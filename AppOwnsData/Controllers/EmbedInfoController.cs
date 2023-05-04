// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

namespace AppOwnsData.Controllers
{
    using AppOwnsData.Models;
    using AppOwnsData.Services;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Options;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;
    using Microsoft.PowerBI.Api.Models;
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text.Json;
    using System.Threading;
    using System.Threading.Tasks;

   
    public class EmbedInfoController : Controller
    {
        private readonly PbiEmbedService pbiEmbedService;
        private readonly IOptions<AzureAd> azureAd;
        private readonly IOptions<PowerBI> powerBI;

        public EmbedInfoController(PbiEmbedService pbiEmbedService, IOptions<AzureAd> azureAd, IOptions<PowerBI> powerBI)
        {
            this.pbiEmbedService = pbiEmbedService;
            this.azureAd = azureAd;
            this.powerBI = powerBI;
        }

        /// <summary>
        /// Returns Embed token, Embed URL, and Embed token expiry to the client
        /// </summary>
        /// <returns>JSON containing parameters for embedding</returns>
        [HttpGet]
        public async Task<string> GetEmbedInfoAsync()
        {
            try
            {
                // Validate whether all the required configurations are provided in appsettings.json
                string configValidationResult = ConfigValidatorService.ValidateConfig(azureAd, powerBI);
                if (configValidationResult != null)
                {
                    HttpContext.Response.StatusCode = 400;
                    return configValidationResult;
                }

                EmbedParams embedParams = await pbiEmbedService.GetEmbedParamsAsync(new Guid(powerBI.Value.WorkspaceId), new Guid(powerBI.Value.ReportId));
                return JsonSerializer.Serialize<EmbedParams>(embedParams);
            }
            catch (Exception ex)
            {
                HttpContext.Response.StatusCode = 500;
                return ex.Message + "\n\n" + ex.StackTrace;
            }
        }

        public async Task<string> GetPDF()

        {
            string PaginatedReportName1 = "Test" + DateTime.Now.Ticks.ToString();
      //      var PaginatedReportParameters1 = new List<ParameterValue>() {
      //  new ParameterValue { Name="Profit Centre", Value="3279" },
      //  new ParameterValue { Name="Year", Value="2023" },
      //  new ParameterValue { Name="Period", Value="4" }
      ////  new ParameterValue { Name="Week End Date", Value="08-01-2023" }
      //};
           var res= await pbiEmbedService.ExportPaginatedReportAsync(new Guid(powerBI.Value.WorkspaceId), new Guid(powerBI.Value.ReportId), PaginatedReportName1, FileFormat.PDF, Parameters:null);
            return res;
        }

        public async void CallRefreshAsync()
        {
            HttpClient client = new HttpClient();
            //AAS template
            //client.BaseAddress = new Uri("https://<rollout>.asazure.windows.net/servers/<serverName>/models/<resource>/");

            //PBI template
            client.BaseAddress = new Uri("https://api.powerbi.com/v1.0/myorg/groups/a69c954e-bf9a-4a74-b81b-662429ce84bc/datasets/eaf53f8b-f058-4090-a77c-0bfdcf8cbcd9/");

            // Send refresh request
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await UpdateToken());

            RefreshRequest refreshRequest = new RefreshRequest()
            {
                type = "full",
                maxParallelism = 10
            };

            HttpResponseMessage response = await client.PostAsJsonAsync("refreshes", refreshRequest);
            string content = await response.Content.ReadAsStringAsync();
            response.EnsureSuccessStatusCode();
            Uri location = response.Headers.Location;
            Console.WriteLine(response.Headers.Location);

            // Check the response
            while (response.IsSuccessStatusCode) // Will exit while loop when exit Main() method (it's running asynchronously)
            {
                string output = "";

                // Refresh token if required
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await UpdateToken());

                response = await client.GetAsync(location);
                if (response.IsSuccessStatusCode)
                {
                    output = await response.Content.ReadAsStringAsync();
                }

              //  Console.Clear();
             //   Console.WriteLine(output);

                Thread.Sleep(5000);
            }
        }


        private static async Task<string> UpdateToken()
        {

            // AAS REST API Inputs:
            // string resourceURI = "https://*.asazure.windows.net";
            // string authority = "https://login.windows.net/<TenantID>/oauth2/authorize";
            // AuthenticationContext ac = new AuthenticationContext(authority);

            // PBI REST API Inputs:
            string resourceURI = "https://analysis.windows.net/powerbi/api";
            string TenantID = "abd1682b-8c75-4e59-b3b5-42c43ec58fb2";
            string authority = "https://login.microsoftonline.com/"+ TenantID;
            string[] scopes = new string[] { $"{resourceURI}/.default" };



            #region Use Interactive or username/password

            //string clientID = "<App ID>"; // Native app with necessary API permissions

            //Interactive login if not cached:
            //AuthenticationContext ac = new AuthenticationContext(authority);
            //AuthenticationResult ar = await ac.AcquireTokenAsync(resourceURI, clientID, new Uri("urn:ietf:wg:oauth:2.0:oob"), new PlatformParameters(PromptBehavior.SelectAccount));

            // Username/password:
            // AuthenticationContext ac = new AuthenticationContext(authority);
            // UserPasswordCredential cred = new UserPasswordCredential("<User ID (UPN e-mail format)>", "<Password>");
            // AuthenticationResult ar = await ac.AcquireTokenAsync(resourceURI, clientID, cred);

            #endregion

            // AAS Service Principal:
            // ClientCredential cred = new ClientCredential("<App ID>", "<App Key>");
            // AuthenticationResult ar = await ac.AcquireTokenAsync(resourceURI, cred);


            // PBI Service Principal: 
            AuthenticationContext ac = new AuthenticationContext(authority);
            ClientCredential cred = new ClientCredential("36d62223-af28-4421-bc5e-5967203b6244", "hbd8Q~cQGu76t_sWEj.5DWRl1ZhOYaxNtSslMaRH");
            AuthenticationResult ar = await ac.AcquireTokenAsync(resourceURI, cred);

            return ar.AccessToken;
        }
        class RefreshRequest
        {
            public string type { get; set; }
            public int maxParallelism { get; set; }
        }
    }
}

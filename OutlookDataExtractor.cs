using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;

namespace OutlookDataLabellingTool
{
    class OutlookDataExtractor
    {
        PublicClientApplication _clientApplication = null;

        public OutlookDataExtractor(string clientId)
        {
            _clientApplication = new PublicClientApplication(clientId);
        }
  
        public async Task<string> GetAccessToken(IEnumerable<string> scopes)
        {
            // Acquire an access token for the given scope.
            var authenticationResult = await _clientApplication.AcquireTokenAsync(scopes);
            return authenticationResult.AccessToken;
        }

        public async Task GetOutlookDataAsText(string accessToken, string url, Func<dynamic, Task> handler)
        {
            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Accept.Clear();
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                // Get Outlook body content data as text instead of default HTML
                httpClient.DefaultRequestHeaders.Add("Prefer", "outlook.body-content-type=\"text\"");

                do
                {
                    var httpResponse = await httpClient.GetAsync(url);
                    
                    if (!httpResponse.IsSuccessStatusCode)
                    {
                        throw new HttpException((int)httpResponse.StatusCode, httpResponse.ReasonPhrase);
                    }

                    var httpResponseContent = await httpResponse.Content.ReadAsStringAsync();

                    // use dynamic since we don't know format (can also using returned Newtonsoft JObject)
                    dynamic graphResponse = JsonConvert.DeserializeObject(httpResponseContent);
                    var outlookRecords = graphResponse.value;
                    foreach (var outlookRecord in outlookRecords)
                    {
                        await handler(outlookRecord);
                    }
                    // Continue getting paged data from query 
                    url = graphResponse["@odata.nextLink"];
                }
                while (url != null);

            }
        }
    }
}

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ERPPlus.IntegrationTests.Config;

namespace IntegrationTests.Helper
{
    public static class TokenHelper
    {
        public static async Task<string> GetAccessToken(HttpClient client)
        {
            var body = new[]
            {
                new KeyValuePair<string, string>("grant_type", "client_credentials"),
                new KeyValuePair<string, string>("client_id", AppConfig.ClientId),
                new KeyValuePair<string, string>("client_secret", AppConfig.ClientSecret),
                new KeyValuePair<string, string>("scope", "api_integration"),
            };

            var request = new HttpRequestMessage(HttpMethod.Post, AppConfig.TokenUrl)
            {
                Content = new FormUrlEncodedContent(body)
            };

            var response = await client.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new Exception("Token request failed: " + responseContent);

            var json = JObject.Parse(responseContent);
            return json["access_token"]?.ToString();
        }
    }
}

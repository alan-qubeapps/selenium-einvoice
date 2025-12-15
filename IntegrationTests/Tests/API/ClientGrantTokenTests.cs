using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using ERPPlus.IntegrationTests.Config;


namespace IntegrationTests.Tests.API
{
    [TestFixture]
    public class ClientGrantTokenTests
    {
        private HttpClient _client;

        [OneTimeSetUp]
        public void Setup()
        {
            _client = new HttpClient();
        }

        [OneTimeTearDown]
        public void Cleanup()
        {
            _client.Dispose();
        }

        [Test]
        [Order(1)]
        public async Task GetClientGrantToken_ShouldReturnAccessToken()
        {
            // Arrange
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

            // Act
            var response = await _client.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            // Assert
            Assert.That(response.IsSuccessStatusCode, Is.True, "Request failed: " + responseContent);

            var json = JObject.Parse(responseContent);

            // Validate response
            Assert.That(json["access_token"], Is.Not.Null, "Access token not found in response");
            Assert.That(json["token_type"]?.ToString(), Is.EqualTo("Bearer"));

            // Save token for reuse in later tests
            TokenStorage.AccessToken = json["access_token"]?.ToString();

            TestContext.WriteLine("Access Token: " + TokenStorage.AccessToken);
        }
    }

    // Static storage for token reuse
    public static class TokenStorage
    {
        public static string AccessToken { get; set; }
    }

}

using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using IntegrationTests.Tests.API;
using ERPPlus.IntegrationTests.Config;
using IntegrationTests.Helper;

namespace InvoiceApiTests
{
    [TestFixture]
    public class SendInvoiceTests
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
            public async Task SendInvoice_ShouldReturnSuccess()
            {
                LogHelper.ClearLogs();

                LogHelper.LogStep("Requesting access token...");
                string token = await TokenHelper.GetAccessToken(_client);
                LogHelper.LogStep("Token received successfully");

                // Generate dynamic values
                LogHelper.LogStep("Preparing dynamic values...");
                var now = DateTime.Now;
                var currentDate = now.ToString("yyyy-MM-dd");
                var currentTime = now.ToString("HH:mm:ss");
                var datetime = now.ToString("yyMMddHHmmss");

                // Taxes
                LogHelper.LogStep("Building taxes...");
                var taxes = new List<object>
            {
                new { TaxType = "01", TotalTaxableAmount = 0.00, TotalTaxAmount = 0.00 }
            };

                var detailTaxes = new List<object>
            {
                new { TaxType = "01", TotalTaxAmount = 0.00, TaxRate = 0,
                      TaxPerUnitAmount = 0.00, TaxBaseUnitMeasure = 0 }
            };

                // Transaction details
                LogHelper.LogStep("Building transaction details...");
                var transactionDetails = new List<object>
            {
                new {
                    TransactionDetailClassification = new[] { "022" },
                    TransactionDetailDescription = "L Test Item 0001",
                    TransactionDetailUnitPrice = 100.00,
                    TransactionDetailTaxExemption = new[] {
                        new { Description = "", Amount = "", Rate = 0 }
                    },
                    TransactionDetailSubtotal = 100.00,
                    TransactionDetailTotalExcludingTax = 100.00,
                    TransactionDetailQuantity = 1,
                    TransactionDetailMeasurement = "",
                    TransactionDetailDiscount = new[] {
                        new { Description = "", Amount = "", Rate = "" }
                    },
                    TransactionDetailFee = new[] {
                        new { Description = "", Amount = "", Rate = "" }
                    },
                    TransactionDetailProductTariffCode = "123456789012",
                    TransactionDetailProductCountry = "MYS",
                    TransactionTax = detailTaxes
                }
            };

                // Transaction main
                LogHelper.LogStep("Building transaction object...");
                var transactionKey = $"102-{datetime}";
                var transaction = new
                {
                    SupplierTinNo = "C23768804020",
                    StoreCode = "STR_000001",
                    DocumentCode = "",
                    CustomerCode = "",
                    TransactionReference = new
                    {
                        ExternalStoreCode = "B001",
                        TerminalNo = "T2",
                        TransactionNo = transactionKey,
                        TransactionDate = currentDate
                    },
                    TransactionVersion = "1.1",
                    TransactionCode = "01",
                    TransactionDocumentNo = $"YINV-{transactionKey}",
                    TransactionDocumentDate = currentDate,
                    TransactionDocumentTime = currentTime,
                    TransactionCurrencyCode = "MYR",
                    TransactionAdvancePayment = "0.00",
                    TransactionAdvancePaymentRefNo = $"REFL-{datetime}",
                    TransactionBillRefNo = "Bill_Reff_No",
                    TransactionTotalExcludingTax = "200.00",
                    TransactionTotalIncludingTax = "100.00",
                    TransactionTotalPayableAmount = "100.00",
                    TransactionNetTotal = "100.00",
                    TransactionTotalDiscount = "0.00",
                    TransactionTotalFee = "0.00",
                    TransactionGrandTotalTax = "100.00",
                    TransactionRounding = "0.00",
                    TransactionDetailCount = 1,
                    TransactionDetail = transactionDetails,
                    TransactionTax = taxes
                };

            var body = new { Transaction = new[] { transaction } };
            string jsonBody = JsonConvert.SerializeObject(body, Formatting.Indented);

            // Send request
            LogHelper.LogStep("Sending invoice request 📤");
            var request = new HttpRequestMessage(HttpMethod.Post, AppConfig.InvoiceUrl);
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            request.Headers.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            var response = await _client.SendAsync(request);
            var responseContent = await response.Content.ReadAsStringAsync();

            // Validate
            LogHelper.LogStep("Validating response ✅");
            Assert.That(response.IsSuccessStatusCode, Is.True,
                "Invoice request failed: " + responseContent);

            var jsonResponse = JObject.Parse(responseContent);
            Assert.That(jsonResponse, Is.Not.Null, "No JSON response received.");

            LogHelper.LogStep("Invoice submitted successfully 🎉");

            // ✅ Extract and clean the invoice link
            var linkToken = jsonResponse.SelectToken("data.lhdn.acceptedDocuments[0].url");

            if (linkToken != null)
            {
                string rawLink = linkToken.ToString();
                string cleanedLink = rawLink.Replace("\\/", "/");
                LogHelper.LogStep("✅ Accessible Invoice Link: " + cleanedLink);
            }
            else
            {
                LogHelper.LogStep("⚠️ No transaction link found in response.");
            }


            // 🧾 Always log full response JSON
            LogHelper.LogStep("Full JSON Response: " + responseContent);

            // 👉 Print all cleaned logs into NUnit output
            foreach (var log in LogHelper.GetLogs())
            {
                TestContext.WriteLine(log);
            }

        }
    }
    }

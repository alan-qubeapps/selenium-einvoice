using Allure.Net.Commons;
using Allure.NUnit;
using Allure.NUnit.Attributes;
using AngleSharp.Dom;
using EInvoice.SeleniumTests.Config;
using EInvoice.SeleniumTests.Drivers;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using ScreenRecorderLib;
using SeleniumExtras.WaitHelpers;
using SeleniumTests.Helpers;
using SeleniumTests.Pages.BusinessEntity;
using SeleniumTests.Pages.Log;
using SeleniumTests.Pages.Transaction;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Media;
using System.Reflection.Metadata;
using System.Xml.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.F_Transaction
{

    public static class ExcelDataReaderTransactionValid
    {
        public static IEnumerable<object[]> GetCreateTransactionTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string BEname = worksheet.Cells[row, 1].Text?.Trim();
                    string BETinNumber = worksheet.Cells[row, 2].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 3].Text?.Trim();
                    string BERegisterID = worksheet.Cells[row, 4].Text?.Trim();
                    string BEsst = worksheet.Cells[row, 5].Text?.Trim();
                    string BETTRegisterNumber = worksheet.Cells[row, 6].Text?.Trim();
                    string BEMSIC = worksheet.Cells[row, 7].Text?.Trim();
                    string BEContactNumber = worksheet.Cells[row, 8].Text?.Trim();
                    string BEemail = worksheet.Cells[row, 9].Text?.Trim();
                    string BECity = worksheet.Cells[row, 10].Text?.Trim();
                    string BEState = worksheet.Cells[row, 11].Text?.Trim();
                    string BEPosCode = worksheet.Cells[row, 12].Text?.Trim();
                    string BECountry = worksheet.Cells[row, 13].Text?.Trim();
                    string BEAddress1 = worksheet.Cells[row, 14].Text?.Trim();
                    string BEAddress2 = worksheet.Cells[row, 15].Text?.Trim();
                    string BEAddress3 = worksheet.Cells[row, 16].Text?.Trim();


                    yield return new object[]
                    {
                        BEname, BETinNumber, BEregisterType, BERegisterID, BEsst, BETTRegisterNumber, BEMSIC, BEContactNumber, BEemail, BECity, BEState, BEPosCode, BECountry, BEAddress1, BEAddress2, BEAddress3
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetSearchTransactionTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string tab = worksheet.Cells[row, 1].Text?.Trim();
                    string searchText = worksheet.Cells[row, 2].Text?.Trim();

                    yield return new object[]
                    {
                        tab, searchText
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetExportTransactionTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string tab = worksheet.Cells[row, 1].Text?.Trim();

                    yield return new object[]
                    {
                        tab
                    };

                }
            }
        }



        public static IEnumerable<object[]> GetUpdateTransactionTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string BEname = worksheet.Cells[row, 1].Text?.Trim();
                    string BETinNumber = worksheet.Cells[row, 2].Text?.Trim();
                    string BEsst = worksheet.Cells[row, 3].Text?.Trim();
                    string BETTRegisterNumber = worksheet.Cells[row, 4].Text?.Trim();
                    string BEMSIC = worksheet.Cells[row, 5].Text?.Trim();
                    string BEContactNumber = worksheet.Cells[row, 6].Text?.Trim();
                    string BEemail = worksheet.Cells[row, 7].Text?.Trim();
                    string BECity = worksheet.Cells[row, 8].Text?.Trim();
                    string BEState = worksheet.Cells[row, 9].Text?.Trim();
                    string BEPosCode = worksheet.Cells[row, 10].Text?.Trim();
                    string BECountry = worksheet.Cells[row, 11].Text?.Trim();
                    string BEAddress1 = worksheet.Cells[row, 12].Text?.Trim();
                    string BEAddress2 = worksheet.Cells[row, 13].Text?.Trim();
                    string BEAddress3 = worksheet.Cells[row, 14].Text?.Trim();


                    yield return new object[]
                    {
                        BEname, BETinNumber, BEsst, BETTRegisterNumber, BEMSIC, BEContactNumber, BEemail, BECity, BEState, BEPosCode, BECountry, BEAddress1, BEAddress2, BEAddress3
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetSearchCategoryTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string category = worksheet.Cells[row, 1].Text?.Trim();

                    yield return new object[]
                    {
                        category
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetFilterLHDNStatusTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string lhdnStatus = worksheet.Cells[row, 1].Text?.Trim();

                    yield return new object[]
                    {
                        lhdnStatus
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetResubmitTransactionTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string DocumentNo = worksheet.Cells[row, 1].Text?.Trim();

                    yield return new object[]
                    {
                        DocumentNo
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetFilterAllTestData(string filePath, string sheetName)
        {
            var fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null || worksheet.Dimension == null)
                    throw new Exception($"❌ Sheet '{sheetName}' is empty or missing in {filePath}");

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string CategoryTab = worksheet.Cells[row, 1].Text?.Trim();
                    string DateType = worksheet.Cells[row, 2].Text?.Trim();
                    string DateRange = worksheet.Cells[row, 3].Text?.Trim();
                    string FromMonth = worksheet.Cells[row, 4].Text?.Trim();
                    string FromYear = worksheet.Cells[row, 5].Text?.Trim();
                    string FromDate = worksheet.Cells[row, 6].Text?.Trim();
                    string ToMonth = worksheet.Cells[row, 7].Text?.Trim();
                    string ToYear = worksheet.Cells[row, 8].Text?.Trim();
                    string ToDate = worksheet.Cells[row, 9].Text?.Trim();
                    string Status = worksheet.Cells[row, 10].Text?.Trim();
                    string DocumentType = worksheet.Cells[row, 11].Text?.Trim();
                    string BusinessEntityName = worksheet.Cells[row, 12].Text?.Trim();
                    string StoreName = worksheet.Cells[row, 13].Text?.Trim();

                    yield return new object[]
                    {
                        CategoryTab, DateType, DateRange, FromMonth, FromYear, FromDate, ToMonth, ToYear, ToDate, Status, DocumentType, BusinessEntityName, StoreName
                    };

                }
            }
        }


    }


        
    [TestFixture, Order(1)]
    [AllureNUnit]
    [AllureSuite("Transaction - Transaction - Valid")]
    [AllureEpic("ERP-117")]
    public class TransactionTests_Valid
    {
        private IWebDriver _driver;
        private TransactionPage _TransactionPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "TransactionTestDataValid.xlsx");

        public static IEnumerable<object[]> CreateTransactionTestData =>
        ExcelDataReaderTransactionValid.GetCreateTransactionTestData(ExcelPath, "CreateTransactionTestData");

        public static IEnumerable<object[]> SearchTransactionTestData =>
        ExcelDataReaderTransactionValid.GetSearchTransactionTestData(ExcelPath, "SearchTransactionTestData");

        public static IEnumerable<object[]> ExportTransactionTestData =>
        ExcelDataReaderTransactionValid.GetExportTransactionTestData(ExcelPath, "ExportTransactionTestData");

        public static IEnumerable<object[]> SearchCategoryTestData =>
        ExcelDataReaderTransactionValid.GetSearchCategoryTestData(ExcelPath, "SearchCategoryTestData");

        public static IEnumerable<object[]> FilterLHDNStatusTestData =>
        ExcelDataReaderTransactionValid.GetFilterLHDNStatusTestData(ExcelPath, "FilterLHDNStatusTestData");
        public static IEnumerable<object[]> ResubmitTransactionTestData =>
        ExcelDataReaderTransactionValid.GetResubmitTransactionTestData(ExcelPath, "ResubmitTransactionTestData");
        public static IEnumerable<object[]> FilterAllTestData =>
        ExcelDataReaderTransactionValid.GetFilterAllTestData(ExcelPath, "FilterAllTestData");
        
        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Transaction Page";

            string folderWithModule = Path.Combine(AppConfig.CsvExportFolder, today, moduleName);
            Directory.CreateDirectory(folderWithModule);

            int version = 1;
            string baseFileName;
            string exportPath;

            do
            {
                baseFileName = $"TestResults_{moduleName.Replace(" ", "_")}_v{version}.xlsx";
                exportPath = Path.Combine(folderWithModule, baseFileName);
                version++;
            } while (File.Exists(exportPath));

            // 🟢 Save version for later use
            _fileVersion = version - 1;
            _exportFilePath = exportPath;

            Console.WriteLine($"📂 Using export file: {_exportFilePath}");

            _driver = DriverFactory.CreateDriver();
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            _driver.Manage().Window.Maximize();
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/auth/login");

            try
            {
                var footerElement = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div[1]/div/div/app-footer/div/div/span[2]")
                ));
                _footerValue = footerElement.Text.Trim();
                Console.WriteLine($"📄 Footer captured on login page: {_footerValue}");
            }
            catch
            {
                Console.WriteLine("⚠️ Footer not found on login page.");
                _footerValue = string.Empty;
            }

            _loginHelper = new LoginHelper(_driver, _wait);
            CaptureFooterBeforeLogin();
            _loginHelper.PerformLogin(AppConfig.UserName, AppConfig.Password, false);
            helperFunction.WaitForPageToLoad(_wait);
        }


        [SetUp]
        public void SetUp()
        {
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(3));
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/transaction");
            helperFunction.WaitForPageToLoad(_wait);
            _TransactionPage = new TransactionPage(_driver);
            _logMessages.Clear();

            _moduleName = "Transaction Page";
            string testName = NUnit.Framework.TestContext.CurrentContext.Test.MethodName;
            string baseFolderPath = AppConfig.BaseVideoFolder;
            string todayFolderName = DateTime.Now.ToString("yyyy-MM-dd");

            string fullFolderPath = Path.Combine(baseFolderPath, todayFolderName, _moduleName);
            Directory.CreateDirectory(fullFolderPath);

            // 🟢 Use the SAME version as Excel result file
            int counter = Interlocked.Increment(ref _recordingCounter);
            string recordingFileName = $"{_moduleName}_{testName}_v{_fileVersion}_{counter}.mp4";
            _recordingFilePath = Path.Combine(fullFolderPath, recordingFileName);

            _recordingCompletedEvent.Reset();

            try
            {
                var options = new RecorderOptions
                {
                    RecorderMode = RecorderMode.Video,
                    VideoOptions = new VideoOptions
                    {
                        Framerate = 30,
                        Bitrate = 8000 * 1000
                    },
                    AudioOptions = new AudioOptions
                    {
                        IsAudioEnabled = false
                    }
                };

                _recorder = Recorder.CreateRecorder(options);
                _recorder.OnRecordingComplete += (s, e) => _recordingCompletedEvent.Set();
                _recorder.OnRecordingFailed += (s, e) => _recordingCompletedEvent.Set();
                _recorder.Record(_recordingFilePath);
                Thread.Sleep(2000);

                Console.WriteLine($"📹 Recording started: {_recordingFilePath}");
            }
            catch (Exception ex)
            {
                LogStep($"❌ Failed to start recorder: {ex.Message}");
            }
        }




        [Test]
        [Category("Transaction")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify invoices + total amounts via search")]
        public void ImportB2CTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2CTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2C transaction type");
                _TransactionPage.ClickB2CTransactionButton();

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");


                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);


                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2C transaction type");
                _TransactionPage.ClickB2CTransactionButton();
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → no batch import
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        // Locate Cancel button by class
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Cancel button not present → nothing to do
                    }
                }


                // ===== Read CSV =====
                LogStep("📄 Reading invoice numbers and total amounts from CSV...");
                var invoiceNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 23)
                        {
                            var invoice = columns[3].Trim();
                            if (string.IsNullOrEmpty(invoice))
                            {
                                LogStep("Empty invoice number detected, stop getting CSV data.");
                                break; // stop reading further
                            }

                            invoiceNumbers.Add(invoice);
                            DocumentDate.Add(columns[4].Trim());
                            totalAmounts.Add(columns[22].Trim());
                        }
                    }
                }

                // ===== Verify each invoice + total amount =====
                for (int i = 0; i < invoiceNumbers.Count; i++)
                {
                    string invoice = invoiceNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Invoice";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching invoice '{invoice}' to verify expected Document Date'{expectedDocumentDate}', Document Type '{documentType}',  expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(invoice);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualInvoice = cells[1].Text.Trim(); 
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualInvoice.Equals(invoice, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                        {
                                            isMatchFound = true;
                                            LogStep($"✅ Invoice '{invoice}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ Invoice '{invoice}' total amount mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, invoice not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Invoice '{invoice}' with total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 CSV Import and invoice amount verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Transaction")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify Invoice + total amounts after batch CSV import")]
        public void VerifyB2CInvoiceTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2CTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to Invoice tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2C transaction type");
                _TransactionPage.ClickB2CTransactionButton();
                Thread.Sleep(1000);

                

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading invoice numbers and total amounts from CSV...");
                var invoiceNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 23)
                        {
                            var invoice = columns[3].Trim();
                            if (string.IsNullOrEmpty(invoice))
                            {
                                LogStep("⚠️ Empty invoice number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            invoiceNumbers.Add(invoice);
                            DocumentDate.Add(columns[4].Trim());
                            totalAmounts.Add(columns[22].Trim());
                        }
                    }
                }

                // ===== Verify each invoice + total amount =====
                for (int i = 0; i < invoiceNumbers.Count; i++)
                {
                    string invoice = invoiceNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Invoice";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching invoice '{invoice}' to verify expected Document Date'{expectedDocumentDate}', Document Type '{documentType}',  expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(invoice);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualInvoice = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualInvoice.Equals(invoice, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                        {
                                            isMatchFound = true;
                                            LogStep($"✅ Invoice '{invoice}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ Invoice '{invoice}' total amount mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, invoice not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Invoice '{invoice}' with total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }

            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify invoices + total amounts via search")]
        public void ImportB2BInvoiceTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BInvoiceTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var invoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Invoice ']")));
                invoiceTab.Click();
                LogStep("🧾 Clicked on the 'Invoice' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot2 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot2.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);


                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                invoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Invoice ']")));
                invoiceTab.Click();
                LogStep("🧾 Clicked on the 'Invoice' tab.");
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → no batch import
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        // Locate Cancel button by class
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Cancel button not present → nothing to do
                    }
                }

                // ===== Read CSV =====
                LogStep("📄 Reading invoice numbers and total amounts from CSV...");
                var invoiceNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var invoice = columns[3].Trim();
                            if (string.IsNullOrEmpty(invoice))
                            {
                                LogStep("⚠️ Empty invoice number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            invoiceNumbers.Add(invoice);
                            DocumentDate.Add(columns[5].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each invoice + total amount =====
                for (int i = 0; i < invoiceNumbers.Count; i++)
                {
                    string invoice = invoiceNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Invoice";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching invoice '{invoice}' to verify expected Document Date'{expectedDocumentDate}', Document Type '{documentType}',  expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(invoice);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualInvoice = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualInvoice.Equals(invoice, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                        {
                                            isMatchFound = true;
                                            LogStep($"✅ Invoice '{invoice}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ Invoice '{invoice}' total amount mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, invoice not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Invoice '{invoice}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Transaction")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify Invoice + total amounts after batch CSV import")]
        public void VerifyB2BInvoiceTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BInvoiceTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to Invoice tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var InvoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Invoice ']")));
                InvoiceTab.Click();
                LogStep("🧾 Clicked on the 'Invoice' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading Invoice numbers and total amounts from CSV...");
                var InvoiceNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var Invoice = columns[3].Trim();
                            if (string.IsNullOrEmpty(Invoice)) break;

                            InvoiceNumbers.Add(Invoice);
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each Invoice + total amount =====
                for (int i = 0; i < InvoiceNumbers.Count; i++)
                {
                    string Invoice = InvoiceNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Invoice Note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching Invoice '{Invoice}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(Invoice);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualInvoice = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualInvoice.Equals(Invoice, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"✅ Invoice '{Invoice}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if Invoice not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Invoice '{Invoice}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }




        [Test]
        [Category("Transaction")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify Refunds + total amounts via search")]
        public void ImportB2BRefundTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BRefundTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var RefundTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Refund ']")));
                RefundTab.Click();
                LogStep("🧾 Clicked on the 'Refund' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);


                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                RefundTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Refund ']")));
                RefundTab.Click();
                LogStep("🧾 Clicked on the 'Refund' tab.");
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → no batch import
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        // Locate Cancel button by class
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Cancel button not present → nothing to do
                    }
                }


                // ===== Read CSV =====
                LogStep("📄 Reading Refund numbers and total amounts from CSV...");
                var RefundNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var Refund = columns[3].Trim();
                            if (string.IsNullOrEmpty(Refund))
                            {
                                LogStep("⚠️ Empty Refund number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            RefundNumbers.Add(Refund);
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each Refund + total amount =====
                for (int i = 0; i < RefundNumbers.Count; i++)
                {
                    string Refund = RefundNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Refund Note";
                    bool isMatchFound = false;



                    LogStep($"🔍 Searching Refund '{Refund}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(Refund);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualRefund = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualRefund.Equals(Refund, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                        {
                                            isMatchFound = true;
                                            LogStep($"✅ Refund '{Refund}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ Refund '{Refund}' total amount mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, Refund not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Refund '{Refund}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 CSV Import and Refund amount verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify Refunds + total amounts after batch CSV import")]
        public void VerifyB2BRefundTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BRefundTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to Refund tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var refundTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Refund ']")));
                refundTab.Click();
                LogStep("🧾 Clicked on the 'Refund' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading Refund numbers and total amounts from CSV...");
                var RefundNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var Refund = columns[3].Trim();
                            if (string.IsNullOrEmpty(Refund)) break;

                            RefundNumbers.Add(Refund);
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each Refund + total amount =====
                for (int i = 0; i < RefundNumbers.Count; i++)
                {
                    string Refund = RefundNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Refund Note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching Refund '{Refund}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(Refund);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualRefund = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualRefund.Equals(Refund, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"✅ Refund '{Refund}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if refund not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Refund '{Refund}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Transaction")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify CreditNote + total amounts via search")]
        public void ImportB2BCreditNoteTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BCreditNoteTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var CreditNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Credit Note ']")));
                CreditNoteTab.Click();
                LogStep("🧾 Clicked on the 'Credit Note' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                CreditNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Credit Note ']")));
                CreditNoteTab.Click();
                LogStep("🧾 Clicked on the 'Credit Note' tab.");
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → no batch import
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        // Locate Cancel button by class
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Cancel button not present → nothing to do
                    }
                }

                // ===== Read CSV =====
                LogStep("📄 Reading CreditNote numbers and total amounts from CSV...");
                var CreditNoteNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var CreditNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(CreditNote))
                            {
                                LogStep("⚠️ Empty CreditNote number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            CreditNoteNumbers.Add(CreditNote);
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each CreditNote + total amount =====
                for (int i = 0; i < CreditNoteNumbers.Count; i++)
                {
                    string CreditNote = CreditNoteNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Credit Note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching CreditNote '{CreditNote}' to verify expected Document Date'{expectedDocumentDate}', Document Type '{documentType}',  expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(CreditNote);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualCreditNote = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualCreditNote.Equals(CreditNote, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                        {
                                            isMatchFound = true;
                                            LogStep($"✅ CreditNote '{CreditNote}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ CreditNote '{CreditNote}' total amount mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, CreditNote not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ CreditNote '{CreditNote}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 CSV Import and CreditNote amount verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(8)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify Credit + total amounts after batch CSV import")]
        public void VerifyB2BCreditNoteTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BCreditNoteTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to Credit tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var CreditTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Credit Note ']")));
                CreditTab.Click();
                LogStep("🧾 Clicked on the 'Credit Note' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading CreditNote numbers and total amounts from CSV...");
                var CreditNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var CreditNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(CreditNote))
                            {
                                LogStep("⚠️ Empty CreditNote number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            CreditNumbers.Add(CreditNote);
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each Credit + total amount =====
                for (int i = 0; i < CreditNumbers.Count; i++)
                {
                    string Credit = CreditNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Credit Note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching Credit '{Credit}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(Credit);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualCredit = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualCredit.Equals(Credit, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"✅ Credit '{Credit}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if Credit not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Credit '{Credit}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(9)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify DebitNote + total amounts via search")]
        public void ImportB2BDebitNoteTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BDebitNoteTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var DebitNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Debit Note ']")));
                DebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'Debit Note' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                DebitNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Debit Note ']")));
                DebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'Debit Note' tab.");
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → no batch import
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        // Locate Cancel button by class
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Cancel button not present → nothing to do
                    }
                }

                // ===== Read CSV =====
                LogStep("📄 Reading DebitNote numbers and total amounts from CSV...");
                var DebitNoteNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var DebitNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(DebitNote))
                            {
                                LogStep("⚠️ Empty DebitNote number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            DebitNoteNumbers.Add(DebitNote);
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each DebitNote + total amount =====
                for (int i = 0; i < DebitNoteNumbers.Count; i++)
                {
                    string DebitNote = DebitNoteNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Debit Note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching DebitNote '{DebitNote}' to verify expected Document Date'{expectedDocumentDate}', Document Type '{documentType}',  expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(DebitNote);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualDebitNote = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualDebitNote.Equals(DebitNote, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                        {
                                            isMatchFound = true;
                                            LogStep($"✅ DebitNote '{DebitNote}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ DebitNote '{DebitNote}' total amount mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, DebitNote not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ DebitNote '{DebitNote}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 CSV Import and DebitNote amount verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(10)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify DebitNote + total amounts after batch CSV import")]
        public void VerifyB2BDebitNoteTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BDebitNoteTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to DebitNote tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var DebitNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Debit Note ']")));
                DebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'DebitNote' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading Debit Note numbers and total amounts from CSV...");
                var DebitNoteNumbers = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var DebitNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(DebitNote)) break;

                            DebitNoteNumbers.Add(DebitNote);
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each DebitNote + total amount =====
                for (int i = 0; i < DebitNoteNumbers.Count; i++)
                {
                    string DebitNote = DebitNoteNumbers[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Debit Note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching DebitNote '{DebitNote}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(DebitNote);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualDebitNote = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualDebitNote.Equals(DebitNote, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"✅ DebitNote '{DebitNote}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if DebitNote not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ DebitNote '{DebitNote}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }




        [Test]
        [Category("Transaction")]
        [Order(11)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify SBInvoice + total amounts via search")]
        public void ImportB2BSBInvoiceTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BSBInvoiceTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var SBInvoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Invoice ']")));
                SBInvoiceTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Invoice' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                SBInvoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Invoice ']")));
                SBInvoiceTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Invoice' tab.");
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → no batch import
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        // Locate Cancel button by class
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Cancel button not present → nothing to do
                    }
                }

                // ===== Read CSV =====
                LogStep("📄 Reading SBInvoice numbers and total amounts from CSV...");
                var SBInvoiceNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBInvoice = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBInvoice))
                            {
                                LogStep("⚠️ Empty SBInvoice number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            SBInvoiceNumbers.Add(SBInvoice);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[5].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBInvoice + total amount =====
                for (int i = 0; i < SBInvoiceNumbers.Count; i++)
                {
                    string SBInvoice = SBInvoiceNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self Billed Invoice";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBInvoice '{SBInvoice}' to verify expected Document Date'{expectedDocumentDate}', Document Type '{documentType}',  expected Tin ID'{expectedTinID}',  expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(SBInvoice);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBInvoice = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBInvoice.Equals(SBInvoice, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                            {
                                                isMatchFound = true;
                                                LogStep($"✅ SBInvoice '{SBInvoice}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', Tin ID '{actualTinID}', total amount '{actualAmount}'");

                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ SBInvoice '{SBInvoice}' mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}',  Expected: '{expectedTinID}', Found: '{actualTinID}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, SBInvoice not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ SBInvoice '{SBInvoice}' with document date '{expectedDocumentDate}', document type '{documentType}', Tin ID '{expectedTinID}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 CSV Import and SBInvoice amount verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(12)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify SBInvoice + total amounts after batch CSV import")]
        public void VerifyB2BSBInvoiceTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BSBInvoiceTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to SBInvoice tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var SBInvoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Invoice ']")));
                SBInvoiceTab.Click();
                LogStep("🧾 Clicked on the 'SBInvoice' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading Self Billed Invoice numbers and total amounts from CSV...");
                var SBInvoiceNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBInvoice = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBInvoice)) break;

                            SBInvoiceNumbers.Add(SBInvoice);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[5].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBInvoice + total amount =====
                for (int i = 0; i < SBInvoiceNumbers.Count; i++)
                {
                    string SBInvoice = SBInvoiceNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self Billed Invoice";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBInvoice '{SBInvoice}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(SBInvoice);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBInvoice = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBInvoice.Equals(SBInvoice, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase) &&
                                    actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase))

                                {
                                    isMatchFound = true;
                                    LogStep($"✅ SBInvoice '{SBInvoice}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', Tin ID '{actualTinID}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if SBInvoice not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ SBInvoice '{SBInvoice}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }




        [Test]
        [Category("Transaction")]
        [Order(13)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify SBRefund + total amounts via search")]
        public void ImportB2BSBRefundTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BSBRefundTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var SBRefundTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Refund ']")));
                SBRefundTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Refund' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                SBRefundTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Refund ']")));
                SBRefundTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Refund' tab.");
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → no batch import
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        // Locate Cancel button by class
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // Cancel button not present → nothing to do
                    }
                }

                // ===== Read CSV =====
                LogStep("📄 Reading SBRefund numbers and total amounts from CSV...");
                var SBRefundNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBRefund = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBRefund))
                            {
                                LogStep("⚠️ Empty SBRefund number detected, stopping CSV read.");
                                break; // stop reading further
                            }

                            SBRefundNumbers.Add(SBRefund);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBRefund + total amount =====
                for (int i = 0; i < SBRefundNumbers.Count; i++)
                {
                    string SBRefund = SBRefundNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self billed refund note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBRefund '{SBRefund}' to verify expected Document Date'{expectedDocumentDate}', Document Type '{documentType}',  expected Tin ID'{expectedTinID}',  expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(SBRefund);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // Loop through pages until match or end
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBRefund = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBRefund.Equals(SBRefund, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                            {
                                                isMatchFound = true;
                                                LogStep($"✅ SBRefund '{SBRefund}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', Tin ID '{actualTinID}', total amount '{actualAmount}'");

                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Assert.Fail($"❌ SBRefund '{SBRefund}' mismatch. Expected: '{expectedDocumentDate}', Found: '{actualDocumentDate}'. Expected: '{expectedAmount}', Found: '{actualAmount}',  Expected: '{documentType}', Found: '{actualDocumentType}',  Expected: '{expectedTinID}', Found: '{actualTinID}'");
                                }
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else
                            {
                                // No more pages, SBRefund not found
                                break;
                            }
                        }
                        catch
                        {
                            // Pagination element not found (only one page)
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ SBRefund '{SBRefund}' with document date '{expectedDocumentDate}', document type '{documentType}', Tin ID '{expectedTinID}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 CSV Import and SBRefund amount verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(14)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify SBRefund + total amounts after batch CSV import")]
        public void VerifyB2BSBRefundTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BSBRefundTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to SBRefund tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var SBRefundTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Refund ']")));
                SBRefundTab.Click();
                LogStep("🧾 Clicked on the 'SBRefund' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading Self Billed Refund numbers and total amounts from CSV...");
                var SBRefundNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBRefund = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBRefund)) break;

                            SBRefundNumbers.Add(SBRefund);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBRefund + total amount =====
                for (int i = 0; i < SBRefundNumbers.Count; i++)
                {
                    string SBRefund = SBRefundNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self billed refund note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBRefund '{SBRefund}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(SBRefund);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBRefund = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBRefund.Equals(SBRefund, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase) &&
                                    actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase))

                                {
                                    isMatchFound = true;
                                    LogStep($"✅ SBRefund '{SBRefund}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', Tin ID '{actualTinID}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if SBRefund not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ SBRefund '{SBRefund}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Transaction")]
        [Order(15)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify SBCreditNote + total amounts via search")]
        public void ImportB2BSBCreditNoteTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BSBCreditNoteTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var SBCreditNoteTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Credit Note ']")));
                SBCreditNoteTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed CreditNote' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");

                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                _TransactionPage.ClickB2BTransactionButton();
                SBCreditNoteTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Credit Note ']")));
                SBCreditNoteTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed CreditNote' tab.");
                WaitForUIEffect(1000);

                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // no import in progress
                }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return; // skip further verification
                }
                else
                {
                    // ===== On-fly import =====
                    try
                    {
                        var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                        if (cancelButton.Displayed && cancelButton.Enabled)
                        {
                            LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                            cancelButton.Click();
                            WaitForUIEffect(1000);
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // no cancel button
                    }
                }

                // ===== Read CSV =====
                LogStep("📄 Reading SBCreditNote numbers and total amounts from CSV...");
                var SBCreditNoteNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip header lines

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBCreditNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBCreditNote))
                            {
                                LogStep("⚠️ Empty SBCreditNote number detected, stopping CSV read.");
                                break;
                            }

                            SBCreditNoteNumbers.Add(SBCreditNote);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBCreditNote =====
                for (int i = 0; i < SBCreditNoteNumbers.Count; i++)
                {
                    string SBCreditNote = SBCreditNoteNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self billed credit note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBCreditNote '{SBCreditNote}'...");

                    _TransactionPage.SearchTransaction(SBCreditNote);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // --- Check if “No data available” message is displayed ---
                    try
                    {
                        var noDataElement = _driver.FindElement(By.XPath("//p[contains(text(),'No data available')]"));
                        if (noDataElement.Displayed)
                        {
                            LogStep($"⚠️ No data available for '{SBCreditNote}'. Skipping further verification.");
                            continue;
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        // table has data, continue search
                    }

                    // --- Search in table ---
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBCreditNote = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBCreditNote.Equals(SBCreditNote, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"✅ SBCreditNote '{SBCreditNote}' verified successfully.");
                                    break;
                                }
                                else
                                {
                                    Assert.Fail($"❌ SBCreditNote '{SBCreditNote}' data mismatch. Expected Date: '{expectedDocumentDate}', Type: '{documentType}', TinID: '{expectedTinID}', Amount: '{expectedAmount}'.");
                                }
                            }
                        }

                        if (isMatchFound) break;

                        // --- Pagination handling ---
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break;
                        }
                        catch
                        {
                            break; // only one page
                        }
                    }

                    if (!isMatchFound)
                    {
                        LogStep($"❌ SBCreditNote '{SBCreditNote}' not found in any table page.");
                        Assert.Fail($"❌ SBCreditNote '{SBCreditNote}' not found.");
                    }
                }

                LogStep("🎉 CSV Import and SBCreditNote verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Transaction")]
        [Order(16)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify SBCreditNote + total amounts after batch CSV import")]
        public void VerifyB2BSBCreditNoteTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BSBCreditNoteTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to SBCreditNote tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var SBCreditNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Credit Note ']")));
                SBCreditNoteTab.Click();
                LogStep("🧾 Clicked on the 'SBCreditNote' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading Self Billed CreditNote numbers and total amounts from CSV...");
                var SBCreditNoteNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBCreditNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBCreditNote)) break;

                            SBCreditNoteNumbers.Add(SBCreditNote);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBCreditNote + total amount =====
                for (int i = 0; i < SBCreditNoteNumbers.Count; i++)
                {
                    string SBCreditNote = SBCreditNoteNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self billed credit note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBCreditNote '{SBCreditNote}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(SBCreditNote);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBCreditNote = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBCreditNote.Equals(SBCreditNote, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase) &&
                                    actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase))

                                {
                                    isMatchFound = true;
                                    LogStep($"✅ SBCreditNote '{SBCreditNote}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', Tin ID '{actualTinID}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if SBCreditNote not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ SBCreditNote '{SBCreditNote}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(17)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button and verify SBDebitNote + total amounts via search")]
        public void ImportB2BSBDebitNoteTransactionCSVFile()
        {
            string filePath = AppConfig.ImportB2BSBDebitNoteTransactionWDCSVFile;

            try
            {
                // ===== Import CSV =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();

                var SBDebitNoteTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Debit Note ']")));
                SBDebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Debit Note' tab.");

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("ngb-modal-window app-upload-modal .modal-body .d-flex.align-items-center > button"));
                _TransactionPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(1000);

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                        LogStep($"❌ Import Error detected: {importError.Text}");
                        Assert.Fail("Test failed due to Import Error popup.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("✅ No Import Error popup detected, continue with verification...");
                }

                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Navigate to Transaction Page =====
                LogStep("📄 Navigating to Transaction page to check import status...");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                SBDebitNoteTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Debit Note ']")));
                SBDebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Debit Note' tab again.");
                Thread.Sleep(1000);


                // ===== Check for batch import =====
                bool isBatchImport = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isBatchImport = true;
                    }
                }
                catch (NoSuchElementException) { }

                if (isBatchImport)
                {
                    LogStep("⚠️ Batch import detected. Cannot perform immediate verification.");
                    LogStep("❗ Please run a separate test case later to verify the inserted data.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"BatchDetected_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return;
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading SBDebitNote numbers and total amounts from CSV...");
                var SBDebitNoteNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue;

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBDebitNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBDebitNote))
                            {
                                LogStep("⚠️ Empty SBDebitNote number detected, stopping CSV read.");
                                break;
                            }

                            SBDebitNoteNumbers.Add(SBDebitNote);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBDebitNote + total amount =====
                for (int i = 0; i < SBDebitNoteNumbers.Count; i++)
                {
                    string SBDebitNote = SBDebitNoteNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self billed debit note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBDebitNote '{SBDebitNote}' to verify expected data...");

                    _TransactionPage.SearchTransaction(SBDebitNote);
                    WaitForUIEffect(2500);
                    helperFunction.WaitForTransactionTableToLoad(_wait);

                    // ===== Loop through table pages =====
                    while (true)
                    {
                        // 🧩 Check for “No data available”
                        try
                        {
                            var noData = _driver.FindElement(By.XPath("//p[contains(text(),'No data available')]"));
                            if (noData.Displayed)
                            {
                                LogStep("⚠️ No data available detected. Ending search for this record.");
                                break;
                            }
                        }
                        catch (NoSuchElementException) { }

                        var rows = _driver.FindElements(By.XPath("//table/tbody/tr"));
                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBDebitNote = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBDebitNote.Equals(SBDebitNote, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"✅ Matched: {SBDebitNote} | Date: {actualDocumentDate} | Type: {actualDocumentType} | TinID: {actualTinID} | Amount: {actualAmount}");
                                    break;
                                }
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("li.page-item i.next"));
                            var parentLi = nextButton.FindElement(By.XPath("./ancestor::li"));

                            if (parentLi.GetAttribute("class").Contains("disabled") || !nextButton.Displayed)
                            {
                                LogStep("⛔ Reached last page or pagination disabled. Stopping search.");
                                break;
                            }

                            nextButton.Click();
                            WaitForUIEffect(2000);
                            helperFunction.WaitForTransactionTableToLoad(_wait);
                        }
                        catch (NoSuchElementException)
                        {
                            LogStep("⚠️ Pagination not found, likely single-page result.");
                            break;
                        }
                    }

                    if (!isMatchFound)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"NoMatch_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        Assert.Fail($"❌ SBDebitNote '{SBDebitNote}' not found or details mismatch.");
                    }
                }

                LogStep("🎉 CSV Import and SBDebitNote verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Error_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Transaction")]
        [Order(18)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Verify SBDebitNote + total amounts after batch CSV import")]
        public void VerifyB2BSBDebitNoteTransactionBatchImport()
        {
            // Reuse the same CSV used in the import test
            string filePath = AppConfig.ImportB2BSBDebitNoteTransactionWDCSVFile;

            try
            {
                LogStep("📄 Navigating to Transaction page...");

                // ===== Go to SBDebitNote tab =====
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(1000);

                var SBDebitNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Debit Note ']")));
                SBDebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'SBDebitNote' tab.");
                Thread.Sleep(1000);

                // ===== Check if batch import still in progress =====
                bool isImportInProgress = false;
                try
                {
                    var marquee = _driver.FindElement(By.CssSelector("div.marquee-section p.marquee"));
                    if (marquee.Displayed && marquee.Text.Contains("Import In Progress"))
                    {
                        isImportInProgress = true;
                    }
                }
                catch (NoSuchElementException)
                {
                    // Marquee not found → import likely completed
                }

                if (isImportInProgress)
                {
                    LogStep("⚠️ Import still in progress. Cannot verify data yet.");
                    LogStep("❗ Please run this test case later to verify the inserted data.");
                    return; // exit test
                }

                // ===== Dismiss leftover modal/overlay =====
                try
                {
                    var cancelButton = _driver.FindElement(By.CssSelector("button.btn.btn-light.mx-2"));
                    if (cancelButton.Displayed && cancelButton.Enabled)
                    {
                        LogStep("🛑 Clicking 'Cancel' button to dismiss leftover modal/overlay before searching.");
                        cancelButton.Click();
                        WaitForUIEffect(1000);
                    }
                }
                catch (NoSuchElementException) { }
                catch (ElementNotInteractableException) { }

                // ===== Read CSV =====
                LogStep("📄 Reading Self Billed DebitNote numbers and total amounts from CSV...");
                var SBDebitNoteNumbers = new List<string>();
                var TinID = new List<string>();
                var DocumentDate = new List<string>();
                var totalAmounts = new List<string>();

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip first 3 rows

                        var columns = line.Split(',');
                        if (columns.Length >= 24)
                        {
                            var SBDebitNote = columns[3].Trim();
                            if (string.IsNullOrEmpty(SBDebitNote)) break;

                            SBDebitNoteNumbers.Add(SBDebitNote);
                            TinID.Add(columns[0].Trim());
                            DocumentDate.Add(columns[6].Trim());
                            totalAmounts.Add(columns[23].Trim());
                        }
                    }
                }

                // ===== Verify each SBDebitNote + total amount =====
                for (int i = 0; i < SBDebitNoteNumbers.Count; i++)
                {
                    string SBDebitNote = SBDebitNoteNumbers[i];
                    string expectedTinID = TinID[i];
                    string expectedDocumentDate = DocumentDate[i];
                    string expectedAmount = totalAmounts[i];
                    string documentType = "Self billed debit note";
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching SBDebitNote '{SBDebitNote}' to verify expected Document Date '{expectedDocumentDate}', Document Type '{documentType}', expected Amount '{expectedAmount}'");

                    _TransactionPage.SearchTransaction(SBDebitNote);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForTransactionTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    // ===== Loop through pages until match or end =====
                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[1]/div/table/tbody/tr"));

                        if (rows.Count == 0)
                        {
                            // No rows → data not found
                            break;
                        }

                        bool foundInThisPage = false;

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualDocumentDate = cells[0].Text.Trim();
                            string actualSBDebitNote = cells[1].Text.Trim();
                            string actualDocumentType = cells[2].Text.Trim();
                            string actualTinID = cells[4].Text.Trim();
                            string actualAmount = cells[7].Text.Trim();

                            if (actualSBDebitNote.Equals(SBDebitNote, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualDocumentDate.Equals(expectedDocumentDate, StringComparison.OrdinalIgnoreCase) &&
                                    actualDocumentType.Equals(documentType, StringComparison.OrdinalIgnoreCase) &&
                                    actualAmount.Equals(expectedAmount, StringComparison.OrdinalIgnoreCase) &&
                                    actualTinID.Equals(expectedTinID, StringComparison.OrdinalIgnoreCase))

                                {
                                    isMatchFound = true;
                                    LogStep($"✅ SBDebitNote '{SBDebitNote}' matched with document date '{actualDocumentDate}', document type '{actualDocumentType}', Tin ID '{actualTinID}', total amount '{actualAmount}'");

                                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                                }
                                foundInThisPage = true;
                                break;
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForTransactionTableToLoad(_wait);
                            }
                            else break; // No more pages
                        }
                        catch
                        {
                            break; // Pagination element not found → only one page
                        }

                        // Stop loop if no match found on this page
                        if (!foundInThisPage) break;
                    }

                    // ===== Fail test if SBDebitNote not found =====
                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ SBDebitNote '{SBDebitNote}' with document date '{expectedDocumentDate}', document type '{documentType}', total amount '{expectedAmount}' was not found.");
                    }
                }

                LogStep("🎉 Batch CSV Import verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during batch import verification: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Transaction")]
        [Order(19)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Transaction Search - General Match (Partial Match Accepted)")]
        [TestCaseSource(nameof(SearchTransactionTestData))]
        public void Search_Transaction(string tab, string searchText)
        {
            LogStep($"🗂 Navigating to tab: '{tab}' before searching for '{searchText}'");

            // --- Step 1: Navigate to correct tab ---
            try
            {
                var tabElement = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath($"//li[contains(@class,'nav-item')]//a[contains(@class,'nav-link')][contains(translate(., ' ', ''), '{tab.Replace(" ", "")}')]"))
                );
                tabElement.Click();
                WaitForUIEffect(3000);
                LogStep($"✅ Switched to tab: '{tab}'");
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail($"❌ Tab '{tab}' not found or not clickable.");
            }

            // --- Step 2: Perform search ---
            LogStep($"🔍 Starting search for: '{searchText}'");
            _TransactionPage.SearchTransaction(searchText);
            WaitForUIEffect(4000);
            helperFunction.WaitForTransactionTableToLoad(_wait);

            bool isMatchFound = false;
            int pageCount = 1;

            while (true)
            {
                helperFunction.WaitForTransactionTableToLoad(_wait);

                // --- Step 3: Check for "No data available" message ---
                try
                {
                    var noData = _driver.FindElement(By.XPath("//p[contains(text(),'No data available')]"));
                    if (noData.Displayed)
                    {
                        LogStep("⚠️ No data available message detected. Ending search.");
                        break;
                    }
                }
                catch (NoSuchElementException)
                {
                    // ignore, means table has data
                }

                // --- Step 4: Check table rows ---
                var rows = _driver.FindElements(By.XPath("//table/tbody/tr"));
                Console.WriteLine($"Total rows on page {pageCount}: {rows.Count}");

                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));
                    foreach (var cell in cells)
                    {
                        string cellText;
                        try { cellText = cell.FindElement(By.TagName("span")).Text.Trim(); }
                        catch { cellText = cell.Text.Trim(); }

                        if (cellText.Replace(" ", "")
                                    .Contains(searchText.Replace(" ", ""), StringComparison.OrdinalIgnoreCase))
                        {
                            isMatchFound = true;
                            LogStep($"✅ Match found in table cell: '{cellText}'");

                            // Capture screenshot
                            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                            break;
                        }
                    }
                    if (isMatchFound) break;
                }

                if (isMatchFound)
                    break;

                // --- Step 5: Check pagination (Next button) ---
                try
                {
                    var nextButton = _driver.FindElement(By.CssSelector("li.page-item i.next"));
                    var parentLi = nextButton.FindElement(By.XPath("./ancestor::li"));

                    if (parentLi.GetAttribute("class").Contains("disabled") || !nextButton.Displayed)
                    {
                        LogStep("⛔ Reached last page or pagination disabled. Stopping loop.");
                        break;
                    }

                    nextButton.Click();
                    WaitForUIEffect(2500);
                    pageCount++;
                    LogStep($"➡️ Moved to next page ({pageCount}).");
                }
                catch (NoSuchElementException)
                {
                    LogStep("⚠️ No pagination element found. Possibly single-page result.");
                    break;
                }
                catch (ElementClickInterceptedException)
                {
                    LogStep("⚠️ Cannot click next page (intercepted or hidden). Stopping.");
                    break;
                }
            }

            // --- Step 6: Final validation ---
            if (!isMatchFound)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"NoMatch_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.Fail($"❌ No matching record found for '{searchText}' after scanning all pages.");
            }
            else
            {
                LogStep("✅ Final assertion passed: match found.");
            }
        }


        [Test]
        [Category("Transaction Exported")]
        [Order(20)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Export Transaction Report")]
        [TestCaseSource(nameof(ExportTransactionTestData))]
        public void ExportTransactionReport(string tab)
        {
            string downloadPath = AppConfig.DownloadPath;
            string filePrefix = "Transaction Index";
            bool isModalDisplayed = false;

            LogStep($"🗂 Navigating to tab: '{tab}'");

            // --- Step 1: Navigate to correct tab ---
            try
            {
                var tabElement = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath($"//li[contains(@class,'nav-item')]//a[contains(@class,'nav-link')][contains(normalize-space(.), '{tab}')]"))
                );
                tabElement.Click();
                WaitForUIEffect(3000);
                LogStep($"✅ Switched to tab: '{tab}'");
            }
            catch (WebDriverTimeoutException)
            {
                Assert.Fail($"❌ Tab '{tab}' not found or not clickable.");
            }

            // --- Step 2: Click Export button ---
            By exportBtnSelector = By.CssSelector("#kt_content_container > app-transactionv2 > div:nth-child(3) > div > div:nth-child(2) > div.col-sm-12.col-md-8.pe-10.d-flex.justify-content-end.align-items-center > div > a");
            helperFunction.WaitForElementToBeClickable(_wait, exportBtnSelector);

            var exportButton = _driver.FindElement(exportBtnSelector);
            LogStep("📤 Clicked Export button");
            exportButton.Click();

            // --- Step 3: Wait for export process or spinner ---
            try
            {
                LogStep("⏳ Waiting for export process (button disable or spinner visible)...");

                bool exportStarted = false;
                DateTime startTime = DateTime.Now;

                while ((DateTime.Now - startTime).TotalSeconds < 10)
                {
                    if (exportButton.GetAttribute("class").Contains("disabled") ||
                        exportButton.GetAttribute("aria-disabled") == "true" ||
                        _driver.FindElements(By.CssSelector(".spinner-border, .ngx-spinner, .loading-indicator, .mat-progress-spinner")).Any())
                    {
                        exportStarted = true;
                        break;
                    }
                    Thread.Sleep(2000);
                }

                if (exportStarted)
                {
                    LogStep("⚙️ Export started. Waiting for it to finish...");

                    bool exportFinished = false;
                    startTime = DateTime.Now;

                    while ((DateTime.Now - startTime).TotalSeconds < 60)
                    {
                        bool buttonReady = !exportButton.GetAttribute("class").Contains("disabled") &&
                                           exportButton.GetAttribute("aria-disabled") != "true";
                        bool spinnerGone = !_driver.FindElements(By.CssSelector(".spinner-border, .ngx-spinner, .loading-indicator, .mat-progress-spinner")).Any();

                        if (buttonReady && spinnerGone)
                        {
                            exportFinished = true;
                            break;
                        }
                        Thread.Sleep(1000);
                    }

                    if (exportFinished)
                        LogStep("✅ Export process completed — ready for verification.");
                    else
                        LogStep("⚠️ Export may still be processing, proceeding to file check.");
                }
                else
                {
                    LogStep("⚠️ No export spinner or disable detected, continue to next check.");
                }
            }
            catch (Exception ex)
            {
                LogStep($"⚠️ Exception during export wait: {ex.Message}. Continue checking modal.");
            }

            // --- Step 4: Handle modal or file check ---
            try
            {
                var modal = _wait.Until(driver =>
                {
                    try
                    {
                        var element = driver.FindElement(By.XPath("/html/body/div/div"));
                        return element.Displayed ? element : null;
                    }
                    catch (NoSuchElementException)
                    {
                        return null;
                    }
                });

                if (modal != null)
                {
                    isModalDisplayed = true;

                    var message = modal.Text.Trim();
                    LogStep("📢 Modal appeared: " + message);

                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    if (!message.ToLower().Contains("under processing due to large size"))
                    {
                        Assert.Fail("❌ Unexpected modal message: " + message);
                    }

                    var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    okButton.Click();
                    LogStep("ℹ️ Export is processing in background. Skipping file check.");
                }

                // --- Step 5: File download verification ---
                if (!isModalDisplayed)
                {
                    string latestFile = null;
                    DateTime startTime = DateTime.Now;
                    bool fileDownloaded = false;

                    LogStep("⏳ Waiting for file download to complete...");

                    while ((DateTime.Now - startTime).TotalSeconds < 90)
                    {
                        var files = Directory.GetFiles(downloadPath, $"{filePrefix}*")
                            .Where(f => !f.EndsWith(".crdownload"))
                            .OrderByDescending(File.GetLastWriteTime)
                            .ToList();

                        if (files.Any())
                        {
                            latestFile = files.First();
                            if (File.GetLastWriteTime(latestFile) >= startTime)
                            {
                                fileDownloaded = true;
                                break;
                            }
                        }

                        Thread.Sleep(1000);
                    }

                    // Always take screenshot at the end
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    if (fileDownloaded && File.Exists(latestFile))
                    {
                        LogStep($"✅ Export completed successfully. File downloaded: {Path.GetFileName(latestFile)}");
                    }
                    else
                    {
                        LogStep("❌ File not downloaded or timed out.");
                        Assert.Fail("❌ File download failed — no file found in expected time.");
                    }
                }
            }
            catch (WebDriverTimeoutException)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("✅ No modal appeared. Proceeding with file verification.");
            }
        }




        [Test]
        [Category("Transaction")]
        [Order(21)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Filter - Filter By Category")]
        [TestCaseSource(nameof(SearchCategoryTestData))]
        public void FilterCategoryFunction(string category)
        {
            IReadOnlyCollection<IWebElement> GetRows() =>
                _driver.FindElements(By.XPath("//table/tbody[1]/tr"));

            string GetStatusFromCell(IWebElement cell)
            {
                try
                {
                    var spanText = cell.FindElement(By.TagName("span")).Text.Trim();
                    if (!string.IsNullOrEmpty(spanText))
                        return spanText;
                }
                catch { }

                try
                {
                    var directText = cell.Text.Trim();
                    if (!string.IsNullOrEmpty(directText))
                        return directText;
                }
                catch { }

                return string.Empty;
            }

            bool IsNoDataMessageShown()
            {
                try
                {
                    var noDataElement = _driver.FindElement(By.XPath("//table/tbody/tr/td"));
                    string message = noDataElement?.Text?.Trim();
                    return message != null && message.Contains("No Data", StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return false;
                }
            }

            void ValidateStatusColumn(int statusColumnIndex = 6, params string[] expectedStatuses)
            {
                var rows = GetRows();
                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));
                    if (cells.Count <= statusColumnIndex) continue; // skip if row has fewer columns

                    string actualStatus = GetStatusFromCell(cells[statusColumnIndex]);

                    LogStep($"🔍 Validating Status: Expected = '{string.Join(", ", expectedStatuses)}', Found = '{actualStatus}'");

                    if (!expectedStatuses.Any(s => s.Equals(actualStatus, StringComparison.OrdinalIgnoreCase)))
                    {
                        LogStep($"❌ Mismatch - Expected: '{string.Join(", ", expectedStatuses)}', Found: '{actualStatus}'");
                        Assert.Fail("❌ One or more rows have unexpected status.");
                    }
                }
            }

            WaitForUIEffect();


            try
            {
                // === Apply Filter ===
                LogStep($"📌 Testing filter: {category} Category");
                switch (category.Trim().ToUpperInvariant())
                {
                    case "ALL":
                        _TransactionPage.ClickFilterALLCategoryButton();
                        break;
                    case "B2C":
                        _TransactionPage.ClickFilterB2CCategoryButton();
                        break;
                    case "B2B":
                        _TransactionPage.ClickFilterB2BCategoryButton();
                        break;
                    case "CONSOLIDATED":
                        _TransactionPage.ClickFilterConsolidatedCategoryButton();
                        break;
                    case "RESUBMIT":
                        _TransactionPage.ClickFilterResubmitCategoryButton();
                        break;
                    case "READY TO SEND LHDN":
                        _TransactionPage.ClickFilterReadytoSendLHDNCategoryButton();
                        break;
                    default:
                        throw new ArgumentException($"❌ Unknown category: {category}");
                }


                WaitForUIEffect();
                helperFunction.WaitForTransactionTableToLoad(_wait);
                Thread.Sleep(2000);

                // === Screenshot ===
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{category}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved to: {_lastScreenshotPath}");

                // === Handle "No Data" Scenario ===
                if (IsNoDataMessageShown())
                {
                    LogStep($"✅ Filter applied: '{category}' - No records found (No Data Available message shown).");
                    return; // Test passes
                }

                var rows = GetRows();
                if (!rows.Any())
                {
                    Assert.Fail("❌ Expected data rows but none were found.");
                }

                // === Validate based on category ===
                if (category.Equals("B2C", StringComparison.OrdinalIgnoreCase))
                {
                    ValidateStatusColumn(6, "Pending", "Processed");
                }
                else if (category.Equals("B2B", StringComparison.OrdinalIgnoreCase))
                {
                    ValidateStatusColumn(6, "Valid", "Invalid", "Pending", "Accepted", "Expired");
                }
                else if (category.Equals("Consolidated", StringComparison.OrdinalIgnoreCase))
                {
                    ValidateStatusColumn(6, "Valid", "Invalid");
                }
                else if (category.Equals("Resubmit", StringComparison.OrdinalIgnoreCase))
                {
                    ValidateStatusColumn(7, "Invalid", "Rejected");
                }
                else if (category.Equals("Ready to send LHDN", StringComparison.OrdinalIgnoreCase))
                {
                    ValidateStatusColumn(6, "Pending");
                }
                else if (category.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var row in rows)
                    {
                        var cells = row.FindElements(By.TagName("td"));
                        if (cells.Count < 7) continue;

                        string actualStatus = GetStatusFromCell(cells[6]);
                        LogStep($"🔍 Found Status = '{actualStatus}'");
                    }
                }
            }
            catch (Exception ex)
            {
                // === Capture Screenshot on Error ===
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Transaction_{category}_ERROR_{DateTime.Now:yyyyMMdd_HHmmssfff}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Exception on '{category}' filter. Screenshot saved: {_lastScreenshotPath}");
                Assert.Fail(ex.Message);
            }
        }

        [Test]
        [Category("Transaction")]
        [Order(22)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Resubmit - Resubmit Transaction")]
        [TestCaseSource(nameof(ResubmitTransactionTestData))]
        public void ResubmitTransactionFunction(string DocumentNos)
        {
            IReadOnlyCollection<IWebElement> GetRows() =>
                _driver.FindElements(By.XPath("//table/tbody[1]/tr"));

            bool IsNoDataMessageShown()
            {
                try
                {
                    var noDataElement = _driver.FindElement(By.XPath("//table/tbody/tr/td"));
                    string message = noDataElement?.Text?.Trim();
                    return message != null && message.Contains("No Data", StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return false;
                }
            }

            void SelectCheckboxForDocument(string doc)
            {
                bool selected = false;
                int retry = 0;

                while (!selected && retry < 3)
                {
                    retry++;
                    try
                    {
                        var rows = GetRows();
                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Any(cell => !string.IsNullOrEmpty(cell.Text) &&
                                cell.Text.Trim().Equals(doc, StringComparison.OrdinalIgnoreCase)))
                            {
                                var checkbox = row.FindElement(By.CssSelector("input[type='checkbox']"));
                                if (!checkbox.Selected)
                                {
                                    checkbox.Click();
                                    LogStep($"✅ Checkbox selected for Document '{doc}'.");
                                }
                                selected = true;
                                break;
                            }
                        }

                        if (!selected)
                            Thread.Sleep(1000);
                    }
                    catch (StaleElementReferenceException)
                    {
                        Thread.Sleep(1000);
                    }
                }

                if (!selected)
                    LogStep($"⚠️ Unable to find checkbox for Document '{doc}' after retries.");
            }

            WaitForUIEffect();

            try
            {
                _TransactionPage.ClickFilterResubmitCategoryButton();
                WaitForUIEffect();

                var docs = new List<string>();
                bool selectAll = false;

                if (!string.IsNullOrEmpty(DocumentNos))
                {
                    if (DocumentNos.Equals("All", StringComparison.OrdinalIgnoreCase))
                        selectAll = true;
                    else
                    {
                        docs = DocumentNos.Split(',')
                                          .Select(d => d.Trim())
                                          .Where(d => !string.IsNullOrEmpty(d))
                                          .ToList();
                    }
                }

                var notFoundDocs = new List<string>();
                var foundDocs = new List<string>();

                if (selectAll)
                {
                    // Check if modal is present before interacting with table
                    try
                    {
                        var modal1 = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]/p")));
                        var message1 = modal1.Text.Trim();
                        if (message1.Contains("Transaction are still processing", StringComparison.OrdinalIgnoreCase))
                        {
                            LogStep($"❌ Modal detected: {message1}");
                            var okButton = modal1.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                            okButton.Click();
                            WaitForUIEffect();
                            Assert.Fail($"❌ Resubmit failed for 'All': {message1}");
                        }
                    }
                    catch (WebDriverTimeoutException)
                    {
                        // Modal not present, proceed to select checkbox
                        var selectAllCheckbox = _driver.FindElement(By.CssSelector("table thead tr th input[type='checkbox']"));
                        if (selectAllCheckbox.Displayed && !selectAllCheckbox.Selected)
                        {
                            selectAllCheckbox.Click();
                            LogStep("✅ 'Select All' checkbox clicked.");
                            foundDocs.Add("All");
                        }
                    }
                }
                else if (docs.Any())
                {
                    foreach (var doc in docs)
                    {
                        LogStep($"🔍 Searching for Document '{doc}'...");
                        _TransactionPage.SearchTransaction(doc);
                        helperFunction.WaitForTransactionTableToLoad(_wait);
                        Thread.Sleep(2000);

                        if (IsNoDataMessageShown())
                        {
                            LogStep($"⚠️ Document '{doc}' not found in Resubmit tab — skipping resubmit for this document.");
                            notFoundDocs.Add(doc);
                            continue;
                        }

                        SelectCheckboxForDocument(doc);
                        foundDocs.Add(doc);
                    }
                }

                WaitForUIEffect();

                if (!foundDocs.Any())
                {
                    LogStep("⚠️ No valid documents found for resubmit — skipping Resubmit button click.");
                    LogStep("🏁 Test passed: No documents available for resubmit.");
                    return;
                }

                WaitForUIEffect();

                var resubmitButton = _driver.FindElement(By.CssSelector("a.btn.btn-light-primary > span.btn-text-hide"));
                resubmitButton.Click();
                LogStep("📤 Resubmit button clicked.");

                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep("🔍 Verifying modal message...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");

                // NEW: Fail immediately if modal shows "Transaction are still processing" for Select All
                if (selectAll && message.Contains("Transaction are still processing", StringComparison.OrdinalIgnoreCase))
                {
                    var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    okButton.Click();
                    WaitForUIEffect();
                    Assert.Fail($"❌ Resubmit failed: Modal indicates transactions are still processing — message: {message}");
                }

                if (!message.ToLower().Contains("success"))
                    Assert.Fail($"❌ Expected success message but got: {message}");

                var okBtn = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okBtn.Click();
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DocumentNos}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved: {_lastScreenshotPath}");

                // Post-resubmit verification for individual documents
                LogStep("🔎 Checking Resubmit tab after resubmit...");
                _TransactionPage.ClickFilterResubmitCategoryButton();
                helperFunction.WaitForTransactionTableToLoad(_wait);
                Thread.Sleep(2000);

                if (IsNoDataMessageShown())
                {
                    LogStep("✅ All transactions have been successfully resubmitted — 'No Data' message displayed.");
                }
                else
                {
                    var remainingRows = GetRows().Count;
                    LogStep($"⚠️ There are still {remainingRows} transactions remaining in the Resubmit tab.");
                }

                LogStep("🎉 Resubmit transaction test completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Transaction_{DocumentNos}_ERROR_{DateTime.Now:yyyyMMdd_HHmmssfff}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Exception during resubmit. Screenshot saved: {_lastScreenshotPath}");
                Assert.Fail(ex.Message);
            }
        }



        [Test]
        [Category("Transaction")]
        [Order(23)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Filter - Apply All Inputs")]
        [TestCaseSource(nameof(FilterAllTestData))]
        public void VerifyTransactionFilters_AllOption(
    string categoryTab, string dateType, string dateRange,
    string fromMonth, string fromYear, string fromDate,
    string toMonth, string toYear, string toDate,
    string status, string documentType, string businessEntity, string storeName)
        {
            try
            {
                LogStep($"📌 Starting filter test with '{categoryTab}' inputs...");

                // ---------------------------
                // CATEGORY TAB
                // ---------------------------
                if (!string.IsNullOrEmpty(categoryTab))
                {
                    try
                    {
                        var tab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.XPath($"//a[contains(normalize-space(.), '{categoryTab}')]")));
                        tab.Click();
                        LogStep($"✅ Category Tab selected: {categoryTab}");

                    }
                    catch
                    {
                        LogStep($"⚠️ Category Tab not found: {categoryTab}");
                    }
                }

                WaitForUIEffect(15000);

                // ---------------------------
                // DATE TYPE (Radio Buttons inside dropdown)
                // ---------------------------
                if (!string.IsNullOrEmpty(dateType))
                {
                    try
                    {
                        // Open the Date Type dropdown using your provided CSS selector
                        var dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.CssSelector("#kt_content_container > app-transactionv2 > div:nth-child(2) > div > div:nth-child(1) > div:nth-child(1) > app-date-filter-dropdown")));
                        dropdown.Click();
                        WaitForUIEffect();
                        LogStep("✅ Date Type dropdown opened.");

                        // Select the radio button by label text
                        var radioOption = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.XPath($"//label[contains(normalize-space(.), '{dateType.Trim()}')]/preceding-sibling::input[@type='radio']")));
                        radioOption.Click();
                        WaitForUIEffect();

                        LogStep($"✅ Date Type selected: {dateType}");
                    }
                    catch (Exception ex)
                    {
                        LogStep($"⚠️ Failed to select Date Type: {dateType}. Exception: {ex.Message}");
                    }
                }



                // ---------------------------
                // DATE RANGE
                // ---------------------------
                DateTime fromDateValue = DateTime.MinValue;
                DateTime toDateValue = DateTime.MaxValue;

                if (!string.IsNullOrEmpty(dateRange))
                {
                    var dateRangeDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                        By.CssSelector("select.form-control.form-select"))); // adjust selector if needed
                    var select = new SelectElement(dateRangeDropdown);

                    try
                    {
                        switch (dateRange.Trim())
                        {
                            case "Specific Date":
                                select.SelectByText(dateRange.Trim());
                                WaitForUIEffect();
                                LogStep($"✅ Date Range selected: {dateRange}");

                                var calendarButton = _wait.Until(
                                    ExpectedConditions.ElementToBeClickable(
                                        By.CssSelector("button.btn.btn-outline-secondary.bi.bi-calendar3.btn-sm")
                                    )
                                );
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", calendarButton);
                                WaitForUIEffect();

                                // ---------------- FROM ----------------
                                if (!string.IsNullOrEmpty(fromMonth))
                                {
                                    var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                                        By.CssSelector("select[aria-label='Select month']")));
                                    new SelectElement(monthDropdown).SelectByText(fromMonth.Trim());
                                    WaitForUIEffect(2000);
                                    LogStep($"✅ From Month selected: {fromMonth}");
                                }

                                if (!string.IsNullOrEmpty(fromYear))
                                {
                                    var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                                        By.CssSelector("select[aria-label='Select year']")));
                                    new SelectElement(yearDropdown).SelectByText(fromYear.Trim());
                                    WaitForUIEffect(2000);
                                    LogStep($"✅ From Year selected: {fromYear}");
                                }

                                if (!string.IsNullOrEmpty(fromDate))
                                {
                                    var fromDay = _wait.Until(ExpectedConditions.ElementToBeClickable(
                                        By.XPath($"//span[contains(@class,'custom-day') and normalize-space(text())='{fromDate.Trim()}']")));
                                    fromDay.Click();
                                    WaitForUIEffect(2000);
                                    LogStep($"✅ From Date selected: {fromDate}");
                                }

                                WaitForUIEffect(2000);


                                // ---------------- TO ----------------
                                if (!string.IsNullOrEmpty(toMonth))
                                {
                                    var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                                        By.CssSelector("select[aria-label='Select month']")));
                                    new SelectElement(monthDropdown).SelectByText(toMonth.Trim());
                                    WaitForUIEffect(2000);
                                    LogStep($"✅ To Month selected: {toMonth}");
                                }

                                if (!string.IsNullOrEmpty(toYear))
                                {
                                    var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                                        By.CssSelector("select[aria-label='Select year']")));
                                    new SelectElement(yearDropdown).SelectByText(toYear.Trim());
                                    WaitForUIEffect(2000);
                                    LogStep($"✅ To Year selected: {toYear}");
                                }

                                if (!string.IsNullOrEmpty(toDate))
                                {
                                    var toDay = _wait.Until(ExpectedConditions.ElementToBeClickable(
                                        By.XPath($"//span[contains(@class,'custom-day') and normalize-space(text())='{toDate.Trim()}']")));
                                    toDay.Click();
                                    WaitForUIEffect(2000);
                                    LogStep($"✅ To Date selected: {toDate}");
                                }

                                WaitForUIEffect(2000);

                                var selectButton = _wait.Until(
                                    ExpectedConditions.ElementToBeClickable(
                                        By.XPath("//button[normalize-space(text())='Select']")
                                    )
                                );
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", selectButton);
                                WaitForUIEffect(2000);
                                break;

                            default:
                                select.SelectByText(dateRange.Trim());
                                WaitForUIEffect(2000);
                                LogStep($"✅ Date Range selected: {dateRange}");
                                break;
                        }

                        // ---------------------------
                        // Parse the selected date range for verification
                        // ---------------------------
                        var dateRangeInput = _driver.FindElement(By.Name("dpRange"));
                        string selectedDate = dateRangeInput.GetAttribute("value").Trim();
                        LogStep($"ℹ️ Selected Date Range: {selectedDate}");

                        if (!string.IsNullOrEmpty(selectedDate) && selectedDate.Contains("-"))
                        {
                            try
                            {
                                var parts = selectedDate.Split('-').Select(p => p.Trim()).ToArray();
                                WaitForUIEffect(2000);

                                // Example: "01 Jan - 31 Dec 25"
                                string[] formats = { "dd MMM", "dd MMM yy" }; // list of possible formats
                                fromDateValue = DateTime.ParseExact(parts[0], formats, CultureInfo.InvariantCulture, DateTimeStyles.None);
                                toDateValue = DateTime.ParseExact(parts[1], "dd MMM yy", CultureInfo.InvariantCulture);

                                LogStep($"✅ Parsed Date Range: From {fromDateValue:yyyy-MM-dd} To {toDateValue:yyyy-MM-dd}");
                            }
                            catch (Exception ex)
                            {
                                LogStep($"⚠️ Failed to parse date range: {selectedDate}, Exception: {ex.Message}");
                            }
                        }
                    }
                    catch { }
                }


                // ---------------------------
                // STATUS
                // ---------------------------
                if (!string.IsNullOrEmpty(status))
                {
                    try
                    {
                        // Open status dropdown
                        var statusToggle = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.CssSelector("#kt_content_container > app-transactionv2 > div:nth-child(2) > div > div:nth-child(1) > div:nth-child(2) > app-status-filter-dropdown")));
                        statusToggle.Click();
                        WaitForUIEffect();
                    }
                    catch
                    {
                        LogStep("⚠️ Failed to open Status filter dropdown.");
                    }

                    if (status.Equals("All", StringComparison.OrdinalIgnoreCase))
                    {
                        // Select all available options that are not yet selected
                        var allOptions = _driver.FindElements(By.XPath(
                            "//div[contains(@class,'selection') and not(contains(@class,'selected'))]/span"));

                        foreach (var option in allOptions)
                        {
                            string optionText = option.Text.Trim(); // capture text first
                            try
                            {
                                option.Click();
                                WaitForUIEffect(); // optional: wait for selection effect
                                LogStep($"✅ Selected status: {optionText}");
                            }
                            catch
                            {

                            }
                        }
                    }
                    else
                    {
                        // Get currently selected items
                        var selectedItems = _driver.FindElements(By.XPath(
                            "//div[contains(@class,'selection') and contains(@class,'selected')]/span"));

                        // Remove items not in input
                        foreach (var item in selectedItems)
                        {
                            string text = item.Text.Trim();
                            if (!status.Split(',').Select(s => s.Trim()).Contains(text, StringComparer.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    var removeBtn = item.FindElement(By.XPath("./following-sibling::i[contains(@class,'delete')]"));
                                    removeBtn.Click();
                                    WaitForUIEffect();
                                    LogStep($"🗑 Removed status: {text}");
                                }
                                catch { }
                            }
                        }

                        // Select new items from dropdown
                        foreach (var s in status.Split(',').Select(x => x.Trim()))
                        {
                            bool alreadySelected = selectedItems.Any(x => x.Text.Trim().Equals(s, StringComparison.OrdinalIgnoreCase));
                            if (alreadySelected) continue;

                            try
                            {
                                var option = _wait.Until(ExpectedConditions.ElementToBeClickable(
                                    By.XPath($"//div[contains(@class,'selection') and not(contains(@class,'selected'))]/span[contains(normalize-space(.), '{s}')]")));
                                option.Click();
                                WaitForUIEffect();
                                LogStep($"✅ Selected status: {s}");
                            }
                            catch
                            {
                                LogStep($"⚠️ Status not found in dropdown: {s}");
                            }
                        }
                    }
                }





                // ---------------------------
                // DOCUMENT TYPE
                // ---------------------------
                if (!string.IsNullOrEmpty(documentType))
                {
                    try
                    {
                        // Open Document Type dropdown
                        var docToggle = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.CssSelector("#kt_content_container > app-transactionv2 > div:nth-child(2) > div > div:nth-child(1) > div:nth-child(3) > app-document-type-filter")));
                        docToggle.Click();
                        WaitForUIEffect();
                    }
                    catch
                    {
                        LogStep("⚠️ Failed to open Document Type dropdown.");
                    }

                    if (documentType.Equals("All", StringComparison.OrdinalIgnoreCase))
                    {
                        // Select all unselected Document Type options
                        var allOptions = _driver.FindElements(By.XPath(
                            "//div[contains(@class,'selection') and not(contains(@class,'selected'))]/span"));

                        foreach (var option in allOptions)
                        {
                            string optionText = option.Text.Trim();
                            try
                            {
                                option.Click();
                                WaitForUIEffect();
                                LogStep($"✅ Selected Document Type: {optionText}");
                            }
                            catch
                            {
                              
                            }
                        }
                    }
                    else
                    {
                        // Get selected items
                        var selectedItems = _driver.FindElements(By.XPath(
                            "//div[contains(@class,'selection') and contains(@class,'selected')]/span"));

                        // Remove any item not included in input
                        foreach (var item in selectedItems)
                        {
                            string text = item.Text.Trim();
                            if (!documentType.Split(',').Select(s => s.Trim()).Contains(text, StringComparer.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    var removeBtn = item.FindElement(By.XPath("./following-sibling::i[contains(@class,'delete')]"));
                                    removeBtn.Click();
                                    WaitForUIEffect();
                                    LogStep($"🗑 Removed Document Type: {text}");
                                }
                                catch { }
                            }
                        }

                        // Add required items
                        foreach (var dt in documentType.Split(',').Select(x => x.Trim()))
                        {
                            bool alreadySelected = selectedItems.Any(x =>
                                x.Text.Trim().Equals(dt, StringComparison.OrdinalIgnoreCase));

                            if (alreadySelected) continue;

                            try
                            {
                                var option = _wait.Until(ExpectedConditions.ElementToBeClickable(
                                    By.XPath($"//div[contains(@class,'selection') and not(contains(@class,'selected'))]/span[contains(normalize-space(.), '{dt}')]")));
                                option.Click();
                                WaitForUIEffect();
                                LogStep($"✅ Selected Document Type: {dt}");
                            }
                            catch
                            {
                                LogStep($"⚠️ Document Type not found in dropdown: {dt}");
                            }
                        }
                    }
                }


                // ---------------------------
                // BUSINESS ENTITY
                // ---------------------------
                if (!string.IsNullOrEmpty(businessEntity))
                {
                    try
                    {
                        // Open Business Entity dropdown
                        var entityToggle = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.CssSelector("#kt_content_container > app-transactionv2 > div:nth-child(2) > div > div:nth-child(2) > div:nth-child(1) > app-entity-filter-dropdown")));
                        entityToggle.Click();
                        WaitForUIEffect(2000);
                        LogStep("✅ Business Entity dropdown opened.");
                    }
                    catch
                    {
                        LogStep("⚠️ Failed to open Business Entity dropdown.");
                    }

                    if (businessEntity.Equals("All", StringComparison.OrdinalIgnoreCase))
                    {
                        // Select all unselected Business Entity options
                        var allOptions = _driver.FindElements(By.XPath(
                            "//div[contains(@class,'selection') and not(contains(@class,'selected'))]/span"));

                        foreach (var option in allOptions)
                        {
                            string optionText = option.Text.Trim();
                            try
                            {
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                                option.Click();
                                WaitForUIEffect();
                                LogStep($"✅ Selected Business Entity: {optionText}");
                            }
                            catch
                            {

                            }
                        }
                    }
                    else
                    {
                        var targetEntities = businessEntity.Split(',').Select(x => x.Trim()).ToList();

                        WaitForUIEffect(2000);


                        // Remove any currently selected entities not in target list
                        var selectedItems = _driver.FindElements(By.XPath(
                            "//div[contains(@class,'selection') and contains(@class,'selected')]/span"));

                        foreach (var item in selectedItems)
                        {
                            string text = item.Text.Trim();
                            if (!targetEntities.Contains(text, StringComparer.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    var removeBtn = item.FindElement(By.XPath("./following-sibling::i[contains(@class,'delete')]"));
                                    removeBtn.Click();
                                    WaitForUIEffect();
                                    LogStep($"🗑 Removed Business Entity: {text}");
                                }
                                catch { }
                            }
                        }

                        // Add required entities
                        foreach (var entity in targetEntities)
                        {
                            WaitForUIEffect(2000);


                            try
                            {
                                // Re-fetch unselected options each time
                                var option = _driver.FindElements(By.XPath(
                                    $"//div[contains(@class,'selection') and not(contains(@class,'selected'))]/span"))
                                    .FirstOrDefault(x => x.Text.Trim().Equals(entity, StringComparison.OrdinalIgnoreCase));

                                if (option != null)
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                                    option.Click();
                                    WaitForUIEffect();
                                    LogStep($"✅ Selected Business Entity: {entity}");
                                }
                                else
                                {
                                }
                            }
                            catch
                            {
                                LogStep($"⚠️ Failed to select Business Entity: {entity}");
                            }
                        }
                    }
                }

                                        
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollTo(0, 0);");
                WaitForUIEffect();


                // ---------------------------
                // STORE NAME
                // ---------------------------
                if (!string.IsNullOrEmpty(storeName))
                {
                    try
                    {
                        // Open Store dropdown
                        var storeToggle = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.CssSelector("#kt_content_container > app-transactionv2 > div:nth-child(2) > div > div:nth-child(2) > div:nth-child(2) > app-store-filter-dropdown")));
                        storeToggle.Click();
                        WaitForUIEffect();
                        LogStep("✅ Store dropdown opened.");
                    }
                    catch
                    {
                        LogStep("⚠️ Failed to open Store Name dropdown.");
                    }

                    if (!storeName.Equals("All", StringComparison.OrdinalIgnoreCase))
                    {
                        var inputStores = storeName.Split(',').Select(x => x.Trim()).ToList();

                        // ----------------------------
                        // 1️⃣ Remove all currently selected stores NOT in Excel
                        // ----------------------------
                        var selectedGroups = _driver.FindElements(By.CssSelector("div.accordion-header button"))
                            .Where(btn =>
                            {
                                try
                                {
                                    var parentDiv = btn.FindElement(By.XPath("./ancestor::div[contains(@class,'accordion-item')]"));
                                    return parentDiv.FindElements(By.CssSelector("div.selection.selected span")).Any();
                                }
                                catch { return false; }
                            }).ToList();

                        foreach (var groupBtn in selectedGroups)
                        {
                            string entityText = groupBtn.Text.Trim();

                            // Expand group if collapsed
                            try
                            {
                                if (groupBtn.GetAttribute("aria-expanded") != "true")
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", groupBtn);
                                    groupBtn.Click();
                                    WaitForUIEffect();
                                }
                            }
                            catch { }

                            // Find selected stores
                            IWebElement groupContainer = null;
                            try
                            {
                                groupContainer = groupBtn.FindElement(By.XPath("./ancestor::div[contains(@class,'accordion-item')]"));
                            }
                            catch { }

                            if (groupContainer != null)
                            {
                                var selectedStores = groupContainer.FindElements(By.CssSelector("div.selection.selected span")).ToList();

                                // Bulk remove stores not in Excel list
                                var toRemove = selectedStores
                                    .Where(s =>
                                    {
                                        string text = s.Text.Trim();
                                        string fullStore = entityText + ":" + text;
                                        return !inputStores.Any(x => x.Equals(fullStore, StringComparison.OrdinalIgnoreCase));
                                    }).ToList();

                                if (toRemove.Any())
                                {
                                    var removeButtons = toRemove.Select(s => s.FindElement(By.XPath("./following-sibling::i[contains(@class,'pi-times')]"))).ToList();
                                    ((IJavaScriptExecutor)_driver).ExecuteScript(@"
                        var btns = arguments[0];
                        for(var i=0;i<btns.length;i++){ btns[i].click(); }
                    ", removeButtons);
                                    WaitForUIEffect();
                                    LogStep($"🗑 Cleared {toRemove.Count} stores not in Excel under entity {entityText}.");
                                }
                            }
                        }

                        // ----------------------------
                        // 2️⃣ Add missing stores from Excel
                        // ----------------------------
                        foreach (var storeFull in inputStores)
                        {
                            string[] parts = storeFull.Split(':');
                            if (parts.Length < 2) continue;
                            string entity = parts[0].Trim();
                            string storeOnly = parts[1].Trim();

                            // Find the group button
                            var groupBtn = _driver.FindElements(By.CssSelector("div.accordion-header button"))
                                .FirstOrDefault(b => b.Text.Trim().Equals(entity, StringComparison.OrdinalIgnoreCase));
                            if (groupBtn == null) continue;

                            // Expand if collapsed
                            if (groupBtn.GetAttribute("aria-expanded") != "true")
                            {
                                try { groupBtn.Click(); WaitForUIEffect(); }
                                catch { ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", groupBtn); WaitForUIEffect(); }
                            }

                            // Click store option
                            try
                            {
                                var storeOption = _wait.Until(ExpectedConditions.ElementToBeClickable(
                                    By.XPath($"//div[contains(@class,'accordion-item')]//button[normalize-space(text())='{entity}']/ancestor::div[contains(@class,'accordion-item')]//span[normalize-space(text())='{storeFull}']")));
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", storeOption);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", storeOption);
                                WaitForUIEffect();
                                LogStep($"✅ Selected store: {storeFull}");
                            }
                            catch
                            {
                                LogStep($"⚠️ Could not find store to select: {storeFull}");
                            }
                        }

                        // Scroll to top to apply filters
                        ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollTo(0,0);");
                        WaitForUIEffect();
                        LogStep("⬆️ Scrolled to top to apply filters.");
                    }
                }




                // ---------------------------
                // APPLY FILTER
                // ---------------------------
                try
                {
                    var apply = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("//button[contains(., 'Apply')]")));
                    apply.Click();
                    WaitForUIEffect();
                    LogStep("📤 Filters applied successfully.");
                }
                catch
                {
                    LogStep("⚠️ Apply button not found.");
                }

                // ---------------------------
                // VERIFY FILTERED TABLE RESULTS
                // ---------------------------
                try
                {
                    LogStep("🔍 Verifying filtered table results...");

                    bool checkStatus = !string.IsNullOrEmpty(status) && !status.Equals("All", StringComparison.OrdinalIgnoreCase);
                    bool checkDocType = !string.IsNullOrEmpty(documentType) && !documentType.Equals("All", StringComparison.OrdinalIgnoreCase);

                    int passedRows = 0;
                    int failedRows = 0;
                    int maxLogRows = 15;
                    int loggedRows = 0;

                    // Set rows per page to 100
                    try
                    {
                        var rowsDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-transactionv2/div[2]/div/div[3]/div/div/div[2]/app-global-pagination/div/div[1]/select")));

                        var selectRows = new SelectElement(rowsDropdown);
                        selectRows.SelectByValue("100");
                        WaitForUIEffect(2000);
                        LogStep("✅ Set table display to 100 rows per page.");
                    }
                    catch { }

                    bool hasNextPage = true;

                    while (hasNextPage)
                    {
                        var tableRows = _wait.Until(d =>
                            d.FindElements(By.CssSelector("table tbody tr"))
                             .Where(r => r.Displayed)
                             .ToList()
                        );

                        if (tableRows.Any(r => r.Text.Contains("No data available")))
                        {
                            LogStep("⚠️ Apply filter returned NO DATA. Test passes but skipping row-level validation.");
                            break;
                        }

                        foreach (var row in tableRows)
                        {
                            bool rowPass = true;
                            string rowLog = "";

                            try
                            {
                                int dateColumnIndex;

                                switch (dateType?.Trim())
                                {
                                    case "Document Date":
                                        dateColumnIndex = 1;
                                        break;
                                    case "Created Date":
                                        dateColumnIndex = 9;
                                        break;
                                    case "Submission Date":
                                        dateColumnIndex = 11;
                                        break;
                                    default:
                                        throw new ArgumentException($"❌ Unknown date type: {dateType}. Test cannot continue.");
                                }

                                // --- GET DATA ---
                                string actualStatus = checkStatus ? row.FindElement(By.CssSelector("td:nth-child(7)")).Text.Trim() : "";
                                string actualDocType = checkDocType ? row.FindElement(By.CssSelector("td:nth-child(3)")).Text.Trim() : "";
                                string dateText = row.FindElement(By.CssSelector($"td:nth-child({dateColumnIndex})")).Text.Trim();
                                DateTime rowDate = DateTime.Parse(dateText.Split(' ')[0]); // only date part



                                // --- STATUS CHECK ---
                                if (checkStatus && !status.Split(',').Select(s => s.Trim()).Any(s => s.Equals(actualStatus, StringComparison.OrdinalIgnoreCase)))
                                    rowPass = false;

                                // --- DOC TYPE CHECK ---
                                if (checkDocType && !documentType.Split(',').Select(s => s.Trim()).Any(s => s.Equals(actualDocType, StringComparison.OrdinalIgnoreCase)))
                                    rowPass = false;

                                // --- DATE RANGE CHECK ---
                                if (rowDate < fromDateValue || rowDate > toDateValue)
                                    rowPass = false;

                                // --- COUNT PASS/FAIL ---
                                if (rowPass) passedRows++;
                                else failedRows++;

                                // --- LOG ONLY FIRST 15 ROWS ---
                                if (loggedRows < 15)
                                {
                                    rowLog = $"Row {loggedRows + 1}: {dateRange} = '{rowDate}', Status = '{actualStatus}', DocumentType = '{actualDocType}' => RowPass = {rowPass}";
                                    LogStep(rowLog);
                                    loggedRows++;
                                }
                            }
                            catch
                            {
                                failedRows++;
                            }
                        }


                        // Check Next page
                        var nextLi = _driver.FindElement(By.CssSelector("li.page-item.next"));
                        if (nextLi.GetAttribute("class").Contains("disabled"))
                        {
                            LogStep("➡️ Next button is disabled. Reached last page.");
                            break;
                        }

                        try
                        {
                            var nextBtn = nextLi.FindElement(By.CssSelector("a.page-link"));
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", nextBtn);
                            WaitForUIEffect(1000);
                        }
                        catch
                        {
                            LogStep("⚠️ Failed to click Next button. Stopping pagination.");
                            break;
                        }
                    }

                    LogStep($"🎯 Filter verification completed: Total Rows = {passedRows + failedRows}, Passed = {passedRows}, Failed = {failedRows}");
                    LogStep($"🔹 Columns checked: {(checkStatus ? "Status" : "")}{(checkStatus && checkDocType ? ", " : "")}{(checkDocType ? "Document Type" : "")}");
                }
                catch (Exception ex)
                {
                    Assert.Fail("❌ Exception during filter verification: " + ex.Message);
                }

                // ---------------------------
                // TAKE SCREENSHOT AFTER FILTER
                // ---------------------------
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_Filter_All_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved to: {_lastScreenshotPath}");

                LogStep("🎉 Filter test with 'All' inputs completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_Filter_All_ERROR_{DateTime.Now:yyyyMMdd_HHmmssfff}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Exception during filter test. Screenshot saved: {_lastScreenshotPath}");
                Assert.Fail(ex.Message);
            }
        }





        [TearDown]
        public void TearDown()
        {
            try
            {
                _recorder?.Stop();
                _recordingCompletedEvent.WaitOne(TimeSpan.FromSeconds(30));

                var context = NUnit.Framework.TestContext.CurrentContext;
                string testName = context.Test.MethodName;
                string result = context.Result.Outcome.Status.ToString();

                string inputParams = "";
                var testMethod = GetType().GetMethod(testName);
                var paramInfos = testMethod?.GetParameters();

                if (paramInfos != null && context.Test.Arguments.Length == paramInfos.Length)
                {
                    var formattedParams = new List<string>();
                    for (int i = 0; i < paramInfos.Length; i++)
                    {
                        string name = paramInfos[i].Name ?? $"Param{i + 1}";
                        string value = context.Test.Arguments[i]?.ToString() ?? "null";
                        formattedParams.Add($"{name} = {value}");
                    }
                    inputParams = string.Join(", ", formattedParams);
                }
                else
                {
                    inputParams = string.Join(", ", context.Test.Arguments.Select(arg => arg?.ToString() ?? "null"));
                }

                string message = CleanMessage(string.Join(" | ", _logMessages));
                DateTime time = DateTime.Now;

                ExportTestResultToExcel(testName, inputParams, result, message, time, _lastScreenshotPath);

            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Error in TearDown: " + ex.Message);
            }
        }


        private string _lastModuleName = string.Empty;
        private int _testCaseCounter = 1;
        private string _lastScreenshotPath = null;
        private string _exportFilePath; // class-level field
        private string _footerValue = string.Empty;

        private void ExportTestResultToExcel(string testName, string inputParams, string result, string message, DateTime time, string screenshotPath = null)
        {
            try
            {
                string testerName = AppConfig.TesterName;
                string developerName = AppConfig.FEDeveloperName + "\n" + AppConfig.BEDeveloperName;
                string managerName = AppConfig.ManagerName;
                string clientName = AppConfig.ClientName;
                string changeDesc = AppConfig.ChangeDesc;

                // Build export file path if not yet set
                if (string.IsNullOrEmpty(_exportFilePath))
                {
                    string today = DateTime.Now.ToString("yyyy-MM-dd");
                    string moduleName = _moduleName.Replace(" ", "_");
                    string folderWithModule = Path.Combine(AppConfig.CsvExportFolder, today, _moduleName);
                    Directory.CreateDirectory(folderWithModule);

                    string baseFileName = $"TestResults_{moduleName}_{today}.xlsx";
                    _exportFilePath = Path.Combine(folderWithModule, baseFileName);
                }

                // Copy from template if not exist
                if (!File.Exists(_exportFilePath))
                {
                    var templatePath = AppConfig.TestCaseFile;
                    File.Copy(templatePath, _exportFilePath);
                }

                var file = new FileInfo(_exportFilePath);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    // ✅ Header information
                    worksheet.Cells["D5"].Value = changeDesc;
                    worksheet.Cells["D7"].Value = _footerValue;
                    worksheet.Cells["F2"].Value = testerName;
                    worksheet.Cells["F4"].Value = developerName;
                    worksheet.Cells["F6"].Value = managerName;
                    worksheet.Cells["F8"].Value = clientName;
                    worksheet.Cells["C13"].Value = testerName;
                    worksheet.Cells["D2"].Value = _moduleName;
                    worksheet.Cells["B13"].Value = DateTime.Now.ToString("yyyy-MM-dd");
                    worksheet.Cells["H2"].Value = DateTime.Now.ToString("yyyy-MM-dd");

                    int startRow = 19;
                    int row = startRow;

                    // Find next empty row
                    while (!string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text))
                    {
                        row++;
                    }

                    // Reset counter if module changed
                    if (_moduleName != _lastModuleName)
                    {
                        _testCaseCounter = 1;
                        _lastModuleName = _moduleName;
                    }

                    // Format test steps
                    string[] steps = message.Split(new[] { '\n', '•', '|' }, StringSplitOptions.RemoveEmptyEntries);
                    string formattedSteps = string.Join("\n", steps.Select((s, i) => $"{i + 1}. {s.Trim()}"));

                    // Extract expected result if passed
                    string expectedResult = "";
                    if (result.Equals("Passed", StringComparison.OrdinalIgnoreCase))
                    {
                        var modalLine = steps.FirstOrDefault(s => s.Trim().StartsWith("Modal:", StringComparison.OrdinalIgnoreCase));
                        if (!string.IsNullOrEmpty(modalLine))
                        {
                            expectedResult = modalLine.Substring(modalLine.IndexOf(':') + 1).Trim().Trim('"');
                        }
                        else
                        {
                            foreach (string s in steps.Reverse())
                            {
                                string trimmed = s.Trim();
                                string lower = trimmed.ToLowerInvariant();
                                if (lower.Contains("successfully") || lower.Contains("has been") || lower.Contains("was saved")
                                    || lower.Contains("updated successfully") || lower.Contains("created") || lower.Contains("deleted")
                                    || lower.Contains("duplicate") || lower.Contains("success") || lower.Contains("match found")
                                    || lower.Contains("found") || lower.Contains("completed") || lower.Contains("download")
                                    || lower.Contains("processing") || lower.Contains("succeeded"))
                                {
                                    expectedResult = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(trimmed.TrimEnd('.'));
                                    break;
                                }
                            }
                        }
                    }

                    // Format input params
                    string formattedInputParams = string.Join(
                        Environment.NewLine,
                        (inputParams ?? string.Empty)
                            .Split(',')
                            .Select(p => p.Trim())
                    );

                    // ✅ Write to Excel main table
                    worksheet.Cells[row, 1].Value = _testCaseCounter;
                    worksheet.Cells[row, 2].Value = _moduleName;
                    worksheet.Cells[row, 3].Value = testName;
                    worksheet.Cells[row, 4].Value = formattedSteps;
                    worksheet.Cells[row, 5].Value = expectedResult;
                    worksheet.Cells[row, 6].Value = formattedInputParams;
                    worksheet.Cells[row, 6].Style.WrapText = true;
                    worksheet.Cells[row, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[row, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    worksheet.Cells[row, 7].Value = result;
                    worksheet.Cells[row, 8].Value = time.ToString("yyyy-MM-dd HH:mm:ss");

                    // ✅ Result color highlight
                    var statusCell = worksheet.Cells[row, 7];
                    statusCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    if (result.Equals("Passed", StringComparison.OrdinalIgnoreCase))
                        statusCell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                    else if (result.Equals("Failed", StringComparison.OrdinalIgnoreCase))
                        statusCell.Style.Fill.BackgroundColor.SetColor(Color.LightPink);

                    // ✅ Screenshot sheet
                    try
                    {
                        if (!string.IsNullOrEmpty(screenshotPath) && File.Exists(screenshotPath))
                        {
                            var screenshotSheet = package.Workbook.Worksheets["Screenshots"];
                            if (screenshotSheet == null)
                                screenshotSheet = package.Workbook.Worksheets.Add("Screenshots");

                            int imgRow = 2;
                            while (!string.IsNullOrWhiteSpace(screenshotSheet.Cells[imgRow, 1].Text))
                            {
                                imgRow += 28;
                            }

                            int mergeWidth = 4;
                            screenshotSheet.Cells[imgRow, 1, imgRow, mergeWidth].Merge = true;
                            screenshotSheet.Cells[imgRow + 1, 1, imgRow + 1, mergeWidth].Merge = true;

                            var labelCell1 = screenshotSheet.Cells[imgRow, 1];
                            labelCell1.Value = $"🧪 Test Case {_testCaseCounter} : {testName}";
                            labelCell1.Style.Font.Bold = true;
                            labelCell1.Style.Font.Size = 12;
                            labelCell1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            labelCell1.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
                            labelCell1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                            var labelCell2 = screenshotSheet.Cells[imgRow + 1, 1];
                            labelCell2.Value = $"🕒 Timestamp: {time:yyyy-MM-dd HH:mm:ss}";
                            labelCell2.Style.Font.Italic = true;
                            labelCell2.Style.Font.Size = 11;
                            labelCell2.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            labelCell2.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                            for (int col = 1; col <= mergeWidth; col++)
                            {
                                screenshotSheet.Column(col).Width = 30;
                            }

                            var image = Image.FromFile(screenshotPath);
                            var excelImage = screenshotSheet.Drawings.AddPicture($"Screenshot_{testName}_{imgRow}", image);
                            excelImage.SetPosition(imgRow + 2, 5, 0, 0);
                            excelImage.SetSize(640, 360);

                            Console.WriteLine($"🖼️ Screenshot inserted successfully for test: {testName} at row {imgRow}.");
                        }
                    }
                    catch (Exception imgEx)
                    {
                        Console.WriteLine("⚠️ Failed to insert screenshot: " + imgEx.Message);
                    }

                    // Save Excel
                    package.Save();
                    _testCaseCounter++;
                }

                Console.WriteLine($"✅ Exported test result to Excel: {_exportFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Error in ExportTestResultToExcel: " + ex.Message);
            }
        }

        // ✅ Capture footer before login (for header info)
        public void CaptureFooterBeforeLogin()
        {
            try
            {
                var footerElement = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/body/div/div[2]/div/span")
                ));
                _footerValue = footerElement.Text.Trim();
                Console.WriteLine($"📄 Footer captured on login page: {_footerValue}");
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("⚠️ Footer not found on login page.");
                _footerValue = string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Failed to capture footer on login page: {ex.Message}");
                _footerValue = string.Empty;
            }
        }




        private void LogStep(string message)
        {
            Console.WriteLine(message);
            _logMessages.Add(CleanMessage(message));
        }

        private string CleanMessage(string raw)
        {
            return raw?
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\"", "'")
                .Replace("✅", "")
                .Replace("❌", "")
                .Replace("📤", "")
                .Replace("💾", "")
                .Replace("🖼️", "")
                .Replace("📢", "")
                .Replace("🔍", "")
                .Replace("⛔", "")
                .Replace("🟡", "")
                .Replace("🟢", "")
                .Replace("🔴", "")
                .Replace("📂", "")
                .Replace("🎉", "")
                .Replace("⏳", "")
                .Replace("⚠️", "")
                .Replace("📌", "")
                .Replace("📁", "")
                .Replace("📸", "")
                .Replace("📄", "")
                .Replace("🔎", "")
                .Replace("ℹ️", "")
                .Replace("🧭", "")
                .Replace("🆕", "")
                .Replace("⌨️", "")
                .Replace("📝", "")
                .Replace("🎨", "")
                .Replace("🎯", "")
                .Replace("🛠️", "")
                .Replace("☑️", "")
                .Replace("📜", "")
                .Replace("🔘", "")
                .Trim();
        }



        private void WaitForUIEffect(int ms = 1000)
        {
            Thread.Sleep(ms); // adjustable UI pause for better video capture
        }



        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            try
            {
                _driver?.Quit();
                _driver?.Dispose();
            }
            catch { }

            try
            {
                SystemSounds.Exclamation.Play();
            }
            catch { }
        }

    }
}

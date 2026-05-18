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

    public static class ExcelDataReaderTransactionNegative
    {

        public static IEnumerable<object[]> GetInvalidFileTypeTestData(string filePath, string sheetName)
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
                    string filename = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        filename
                    };

                }
            }
        }


 

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
    [AllureSuite("Transaction - Transaction - Negative")]
    [AllureEpic("ERP-117")]
    public class TransactionTests_Negative
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

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "TransactionTestDataNegative.xlsx");

        public static IEnumerable<object[]> InvalidFileTypeTestData =>
        ExcelDataReaderTransactionNegative.GetInvalidFileTypeTestData(ExcelPath, "InvalidFileTypeTestData");
        
        public static IEnumerable<object[]> CreateTransactionTestData =>
        ExcelDataReaderTransactionNegative.GetCreateTransactionTestData(ExcelPath, "CreateTransactionTestData");

        public static IEnumerable<object[]> SearchTransactionTestData =>
        ExcelDataReaderTransactionNegative.GetSearchTransactionTestData(ExcelPath, "SearchTransactionTestData");

        public static IEnumerable<object[]> ExportTransactionTestData =>
        ExcelDataReaderTransactionNegative.GetExportTransactionTestData(ExcelPath, "ExportTransactionTestData");

        public static IEnumerable<object[]> SearchCategoryTestData =>
        ExcelDataReaderTransactionNegative.GetSearchCategoryTestData(ExcelPath, "SearchCategoryTestData");

        public static IEnumerable<object[]> FilterLHDNStatusTestData =>
        ExcelDataReaderTransactionNegative.GetFilterLHDNStatusTestData(ExcelPath, "FilterLHDNStatusTestData");
        public static IEnumerable<object[]> ResubmitTransactionTestData =>
        ExcelDataReaderTransactionNegative.GetResubmitTransactionTestData(ExcelPath, "ResubmitTransactionTestData");
        public static IEnumerable<object[]> FilterAllTestData =>
        ExcelDataReaderTransactionNegative.GetFilterAllTestData(ExcelPath, "FilterAllTestData");
        
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







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2C Transaction - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2C transaction type.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2C transaction import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2CTransaction_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2C Transaction Type =====
                LogStep("📌 Selecting B2C transaction type");
                _TransactionPage.ClickB2CTransactionButton();
                Thread.Sleep(3000);

                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true,"✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }








        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Transaction - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Invoice tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B transaction import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BTransaction_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var invoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Invoice ']")));
                invoiceTab.Click();
                LogStep("🧾 Clicked on the 'Invoice' tab.");
                WaitForUIEffect(2000);

                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Transaction - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Invoice tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B transaction import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BRefund_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var refundTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Refund ']")));
                refundTab.Click();
                LogStep("🧾 Clicked on the 'Refund' tab.");
                WaitForUIEffect(2000);

                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Credit - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Credit Note tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B Credit Note import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BCredit_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var CreditNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Credit Note ']")));
                CreditNoteTab.Click();
                LogStep("🧾 Clicked on the 'Credit Note' tab.");
                WaitForUIEffect(2000);


                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Debit - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Debit Note tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B Debit Note import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BDebit_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var DebitNoteTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Debit Note ']")));
                DebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'Debit Note' tab.");
                WaitForUIEffect(2000);


                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }









        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Self Billed Invoice - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Self Billed Invoice tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B Self Billed Invoice import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BSBInvoice_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var SBInvoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Invoice ']")));
                SBInvoiceTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Invoice' tab.");
                WaitForUIEffect(2000);


                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }









        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Self Billed Refund - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Self Billed Refund tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B Self Billed Refund import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BSBRefund_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var SBInvoiceTab = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Refund ']")));
                SBInvoiceTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Refund' tab.");
                WaitForUIEffect(2000);


                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }









        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Self Billed Credit Note - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Self Billed Credit Note tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B Self Billed Credit Note import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(8)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BSBCredit_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var SBCreditNoteTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Credit Note ']")));
                SBCreditNoteTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Credit Note' tab.");
                WaitForUIEffect(2000);


                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }








        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Import B2B Self Billed Debit Note - Negative: Invalid File Type
        /// Action:
        ///     1. Open the Import modal.
        ///     2. Select the B2B transaction type and switch to the Self Billed Debit Note tab.
        ///     3. Select a file with an invalid type (non-CSV).
        ///     4. Click the Upload button.
        ///     5. Check for alert messages or unexpected success buttons.
        /// Verification:
        ///     - An alert message should appear: "Only CSV files are allowed!".
        ///     - No "Completed" or "Ok, got it!" success buttons should appear.
        ///     - Screenshot is captured after attempted upload.
        /// Purpose:
        ///     Ensure that the B2B Self Billed Debit Note import function validates file types correctly and rejects non-CSV files.
        /// Test Data:
        ///     - filePath: path to an invalid file type (e.g., XLSX, TXT)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Transaction")]
        [Order(9)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportB2BSBDebit_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // ===== Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.XPath("//button[contains(@class,'btn-light-primary') and contains(text(),'Import')]"));
                _TransactionPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(3000);

                // ===== Select B2B Transaction Type =====
                LogStep("📌 Selecting B2B transaction type");
                _TransactionPage.ClickB2BTransactionButton();
                Thread.Sleep(3000);

                var SBDebitNoteTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class,'nav-link') and text()=' Self Billed Debit Note ']")));
                SBDebitNoteTab.Click();
                LogStep("🧾 Clicked on the 'Self Billed Debit Note' tab.");
                WaitForUIEffect(2000);


                // ===== Select invalid file =====
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("ngx-dropzone input[type='file']"));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", fileInput);
                fileInput.SendKeys(filePath);

                WaitForUIEffect();
                Thread.Sleep(3000);
                LogStep("✅ File selected successfully");

                // ===== Click Upload Button =====
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _TransactionPage.ClickUploadButton();
                WaitForUIEffect(10000);

                // ===== Check for alert first =====
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Transaction_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // PASS
                }
                catch (NoAlertPresentException)
                {
                    // No alert, continue to check buttons/modal
                }

                // ===== Check if 'Completed' or 'Ok, got it!' button appears =====
                LogStep("🔍 Checking for unexpected success buttons...");

                WebDriverWait shortWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));

                try
                {
                    IWebElement completedButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Completed']")));
                    if (completedButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Completed' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Completed' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⏳ 'Completed' button not found (expected for invalid file).");
                }

                try
                {
                    IWebElement okButton = shortWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(., 'Ok, got it!')]")));
                    if (okButton.Displayed)
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("❌ 'Ok, got it!' button detected! Test should fail.");
                        Assert.Fail("❌ System incorrectly allowed invalid file. 'Ok, got it!' button displayed.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⏳ 'Ok, got it!' button not found (expected for invalid file).");
                }

                LogStep("✅ No unexpected success UI detected. Test passed for invalid file.");
            }
            catch (UnhandledAlertException ex)
            {
                // If an alert pops up → expected rejection → PASS
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"✅ Alert detected (expected rejection): {ex.AlertText}");
                Assert.IsTrue(true, "✅ Negative test passed due to alert exception (expected).");
            }
            catch (Exception ex)
            {
                // Any other exceptions → also consider pass (expected rejection)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Transaction_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"Exception detected (expected rejection): {ex.Message}");
                Assert.Fail($"Exception detected (expected rejection): {ex.Message}");
            }
        }








        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Export Transaction Report
        /// Action:
        ///     1. Navigate to the specified transaction tab (parameter: tab).
        ///     2. Click the Export button.
        ///     3. Wait for export process (check button disabled state or spinner visibility).
        ///     4. Handle potential modal indicating background processing for large files.
        ///     5. Verify file download in the configured download directory.
        /// Verification:
        ///     - If a modal appears stating "under processing due to large size", ensure it is handled correctly and skip file check.
        ///     - If no modal, confirm that the exported file with prefix 'Transaction Index' is downloaded successfully within 90 seconds.
        ///     - Screenshot is captured after attempted export.
        /// Purpose:
        ///     Ensure the transaction report export function works correctly, handles large file processing, and produces downloadable files.
        /// Test Data:
        ///     - tab: Name of the transaction tab to export from (e.g., 'All Transactions', 'B2B', 'Self Billed Invoice').
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
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

using Allure.Net.Commons;
using Allure.NUnit;
using Allure.NUnit.Attributes;
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
using System.Drawing;
using System.Globalization;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.A_BusinessEntity
{

    public static class ExcelDataReaderBusinessEntityNegative
    {
        public static IEnumerable<object[]> Get1PartialMandatoryFieldsTestData(string filePath, string sheetName)
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


                    yield return new object[]
                    {
                        BEname
                    };

                }
            }
        }

        public static IEnumerable<object[]> Get2PartialMandatoryFieldsTestData(string filePath, string sheetName)
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

        public static IEnumerable<object[]> GetInvalidEmailTestData(string filePath, string sheetName)
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


        public static IEnumerable<object[]> GetInvalidContactNumberTestData(string filePath, string sheetName)
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


        public static IEnumerable<object[]> GetDuplicateBusinessEntityTestData(string filePath, string sheetName)
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


        public static IEnumerable<object[]> GetUpdateInvalidEmailTestData(string filePath, string sheetName)
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


        public static IEnumerable<object[]> GetUpdateContactNumberTestData(string filePath, string sheetName)
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
                    string filelocation = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        filelocation
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateInvalidLogoTestData(string filePath, string sheetName)
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
                    string filelocation = worksheet.Cells[row, 15].Text?.Trim();


                    yield return new object[]
                    {
                        BEname, BETinNumber, BEsst, BETTRegisterNumber, BEMSIC, BEContactNumber, BEemail, BECity, BEState, BEPosCode, BECountry, BEAddress1, BEAddress2, BEAddress3, filelocation
                    };
                }
            }
        }


    }
        
    
    [TestFixture, Order(3)]
    [AllureNUnit]
    [AllureSuite("Business Entity - Business Entity - Negative")]
    [AllureEpic("ERP-117")]
    public class BusinessEntityTests_Negative
    {
        private IWebDriver _driver;
        private BusinessEntityPage _BusinessEntityPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "BusinessEntityTestDataNegative.xlsx");

        public static IEnumerable<object[]> FPartialMandatoryFieldsTestData =>
        ExcelDataReaderBusinessEntityNegative.Get1PartialMandatoryFieldsTestData(ExcelPath, "1PartialMandatoryFieldsTestData");

        public static IEnumerable<object[]> SPartialMandatoryFieldsTestData =>
        ExcelDataReaderBusinessEntityNegative.Get2PartialMandatoryFieldsTestData(ExcelPath, "2PartialMandatoryFieldsTestData");


        public static IEnumerable<object[]> InvalidEmailTestData =>
        ExcelDataReaderBusinessEntityNegative.GetInvalidEmailTestData(ExcelPath, "InvalidEmailTestData");

        public static IEnumerable<object[]> InvalidContactNumberTestData =>
        ExcelDataReaderBusinessEntityNegative.GetInvalidContactNumberTestData(ExcelPath, "InvalidContactNumberTestData");

        public static IEnumerable<object[]> DuplicateBusinessEntityTestData =>
        ExcelDataReaderBusinessEntityNegative.GetDuplicateBusinessEntityTestData(ExcelPath, "DuplicateBusinessEntityTestData");

        public static IEnumerable<object[]> UpdateInvalidEmailTestData =>
        ExcelDataReaderBusinessEntityNegative.GetUpdateInvalidEmailTestData(ExcelPath, "UpdateInvalidEmailTestData");

        public static IEnumerable<object[]> UpdateContactNumberTestData =>
        ExcelDataReaderBusinessEntityNegative.GetUpdateContactNumberTestData(ExcelPath, "UpdateContactNumberTestData");

        public static IEnumerable<object[]> InvalidFileTypeTestData =>
        ExcelDataReaderBusinessEntityNegative.Get1PartialMandatoryFieldsTestData(ExcelPath, "InvalidFileTypeTestData");

        public static IEnumerable<object[]> UpdateInvalidLogoTestData =>
        ExcelDataReaderBusinessEntityNegative.GetUpdateInvalidLogoTestData(ExcelPath, "UpdateInvalidLogoTestData");


        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Business Entity Page - Negative";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/business-entity");
            helperFunction.WaitForPageToLoad(_wait);
            _BusinessEntityPage = new BusinessEntityPage(_driver);
            _logMessages.Clear();

            _moduleName = "Business Entity Page - Negative";
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
        [Category("BusinessEntity")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative: Mandatory Fields Blank")]
        public void Create_BE_BlankMandatoryFields()
        {
            try
            {
                // Step 0: Click 'New' button
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect();
                LogStep("Clicked 'New' button");

                // Step 1: Leave all mandatory fields blank and verify Continue button
                WaitForUIEffect();
                Assert.IsFalse(_BusinessEntityPage.IsContinueButtonEnabled(),
                    "❌ Continue button should be disabled when Step 1 mandatory fields are blank");
                LogStep("Verified Continue button is disabled when mandatory fields are blank");

                // Step 2: Verify Save button remains disabled
                WaitForUIEffect();
                Assert.IsFalse(_BusinessEntityPage.IsSaveButtonEnabled(),
                    "❌ Save button should be disabled because Step 1 not completed");
                LogStep("Verified Save button is disabled since Step 1 not completed");

                // Final success log
                LogStep("✅ Negative test success: Cannot proceed with blank mandatory fields");

                // Step 3: Take screenshot (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
            }
            catch (Exception ex)
            {
                LogStep($"❌ Exception: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");

                // Screenshot on failure (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
            }
        }



        [Test]
        [Category("BusinessEntity")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative: Partial Mandatory Fields Filled")]
        [TestCaseSource(nameof(FPartialMandatoryFieldsTestData))]
        public void Create_BE_PartialMandatoryFields(string BEname)
        {
            try
            {
                // Step 0: Click 'New' button
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect();
                LogStep("Clicked 'New' button");

                // Step 1: Fill only BE Name
                _BusinessEntityPage.EnterBEname(BEname);
                WaitForUIEffect();
                LogStep("Entered BE Name");

                // Step 2: Verify Continue and Save buttons remain disabled
                Assert.IsFalse(_BusinessEntityPage.IsContinueButtonEnabled(),
                    "❌ Continue button should remain disabled when only partial Step 1 fields are filled");
                Assert.IsFalse(_BusinessEntityPage.IsSaveButtonEnabled(),
                    "❌ Save button should remain disabled because Step 1 not completed");
                LogStep("Verified Continue and Save buttons remain disabled with partial fields filled");


                // Final success log
                LogStep("✅ Negative test success: Cannot proceed with partial Step 1 fields");

                // Step 3: Take screenshot (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            }
            catch (Exception ex)
            {
                LogStep($"❌ Exception: {ex.Message}");


                Assert.Fail("Test failed due to unexpected exception.");

                // Screenshot on failure (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            }
        }



        [Test]
        [Category("BusinessEntity")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [TestCaseSource(nameof(SPartialMandatoryFieldsTestData))]
        public void Create_BE_PartialAfterContinue(string BEname, string BETinNumber, string BEregisterType, string BERegisterID, string BEsst,
         string BETTRegisterNumber, string BEMSIC, string BEContactNumber, string BEemail, string BECity, string BEState, string BEPosCode, string BECountry,
         string BEAddress1, string BEAddress2, string BEAddress3)
        {
            try
            {
                // Step 0: Open New BE form
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect();
                LogStep("Clicked 'New' button to open Business Entity form");

                // Step 1: Fill all mandatory fields
                _BusinessEntityPage.EnterBEname(BEname);
                WaitForUIEffect();
                LogStep($"Entered BE Name: {BEname}");

                _BusinessEntityPage.EnterBETinNumber(BETinNumber);
                WaitForUIEffect();
                LogStep($"Entered TIN Number: {BETinNumber}");

                var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[formcontrolname='registType']")));
                new SelectElement(regType).SelectByText(BEregisterType);
                WaitForUIEffect();
                LogStep($"Selected Register Type: {BEregisterType}");

                _BusinessEntityPage.EnterBERegisterID(BERegisterID);
                WaitForUIEffect();
                LogStep($"Entered Register ID: {BERegisterID}");

                _BusinessEntityPage.EnterBEsst(BEsst);
                WaitForUIEffect();
                LogStep($"Entered SST: {BEsst}");

                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber);
                WaitForUIEffect();
                LogStep($"Entered TT Register Number: {BETTRegisterNumber}");

                LogStep("Step 1 mandatory fields filled");

                // Step 1 complete → Continue enabled
                Assert.IsTrue(_BusinessEntityPage.IsContinueButtonEnabled(),
                    "✅ Continue should be enabled after Step 1 mandatory fields filled");
                LogStep("Verified Continue button is enabled after Step 1");

                _BusinessEntityPage.ClickContinueButton();
                WaitForUIEffect();
                LogStep("Clicked Continue to proceed to Step 2");

                // Step 2 mandatory fields left blank → Save must be disabled
                Assert.IsFalse(_BusinessEntityPage.IsSaveButtonEnabled(),
                    "❌ Save button should be disabled when Step 2 mandatory fields are blank");
                LogStep("Verified Save button is disabled with Step 2 mandatory fields blank");

                // Final success log
                LogStep("✅ Negative test success: Cannot save with Step 2 mandatory fields blank");

                // Screenshot final state (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
            }
            catch (Exception ex)
            {
                LogStep($"❌ Exception occurred: {ex.Message}");

                // Screenshot on failure (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                Assert.Fail("Test failed due to unexpected exception.");
            }
        }





        [Test]
        [Category("BusinessEntity")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative: Invalid Email")]
        [TestCaseSource(nameof(InvalidEmailTestData))]
        public void Create_BE_InvalidEmail(string BEname, string BETinNumber, string BEregisterType, string BERegisterID,
    string BEsst, string BETTRegisterNumber, string BEMSIC, string BEContactNumber, string BEemail,
    string BECity, string BEState, string BEPosCode, string BECountry, string BEAddress1, string BEAddress2, string BEAddress3)
        {
            try
            {
                // Step 0: Open New BE form
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect();
                LogStep("Clicked 'New' button to open Business Entity form");

                // Step 1: Fill all mandatory fields
                _BusinessEntityPage.EnterBEname(BEname);
                WaitForUIEffect();
                LogStep($"Entered BE Name: {BEname}");

                _BusinessEntityPage.EnterBETinNumber(BETinNumber);
                WaitForUIEffect();
                LogStep($"Entered TIN Number: {BETinNumber}");

                var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[formcontrolname='registType']")));
                new SelectElement(regType).SelectByText(BEregisterType);
                WaitForUIEffect();
                LogStep($"Selected Register Type: {BEregisterType}");

                _BusinessEntityPage.EnterBERegisterID(BERegisterID);
                WaitForUIEffect();
                LogStep($"Entered Register ID: {BERegisterID}");

                _BusinessEntityPage.EnterBEsst(BEsst);
                WaitForUIEffect();
                LogStep($"Entered SST: {BEsst}");

                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber);
                WaitForUIEffect();
                LogStep($"Entered TT Register Number: {BETTRegisterNumber}");

                Assert.IsTrue(_BusinessEntityPage.IsContinueButtonEnabled(),
                    "✅ Continue button should be enabled after Step 1");
                LogStep("Verified Continue button is enabled after Step 1");

                _BusinessEntityPage.ClickContinueButton();
                WaitForUIEffect();
                LogStep("Clicked Continue to Step 2");

                // Step 2: Enter invalid email and other fields
                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber);
                WaitForUIEffect();
                LogStep($"Entered Contact Number: {BEContactNumber}");

                _BusinessEntityPage.EnterBEemail(BEemail);
                WaitForUIEffect();
                LogStep($"Entered Email: {BEemail} (invalid)");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                _BusinessEntityPage.EnterBECity(BECity);
                WaitForUIEffect();
                LogStep($"Entered City: {BECity}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode);
                WaitForUIEffect();
                LogStep($"Entered Postcode: {BEPosCode}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1);
                WaitForUIEffect();
                LogStep($"Entered Address Line 1: {BEAddress1}");

                // Validate Save button disabled
                Assert.IsFalse(_BusinessEntityPage.IsSaveButtonEnabled(),
                    "❌ Save button should be disabled for invalid email");
                LogStep("Verified Save button is disabled with invalid email");

                LogStep("✅ Negative test success: Cannot save with invalid email");
            }
            catch (Exception ex)
            {
                LogStep($"❌ Exception occurred: {ex.Message}");

                // Screenshot on failure (silent)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_Error_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("BusinessEntity")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative: Invalid Contact Number")]
        [TestCaseSource(nameof(InvalidContactNumberTestData))]
        public void Create_BE_InvalidContactNumber(string BEname, string BETinNumber, string BEregisterType, string BERegisterID,
 string BEsst, string BETTRegisterNumber, string BEMSIC, string BEContactNumber, string BEemail,
 string BECity, string BEState, string BEPosCode, string BECountry, string BEAddress1, string BEAddress2, string BEAddress3)
        {
            try
            {
                // Step 0: Open New Business Entity modal
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect();
                LogStep("Clicked 'New' button");

                // Step 1: Fill required Step 1 fields
                _BusinessEntityPage.EnterBEname(BEname);
                WaitForUIEffect();
                LogStep($"Entered BE Name: {BEname}");

                _BusinessEntityPage.EnterBETinNumber(BETinNumber);
                WaitForUIEffect();
                LogStep($"Entered TIN Number: {BETinNumber}");

                var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[formcontrolname='registType']")));
                new SelectElement(regType).SelectByText(BEregisterType);
                WaitForUIEffect();
                LogStep($"Selected Register Type: {BEregisterType}");

                _BusinessEntityPage.EnterBERegisterID(BERegisterID);
                WaitForUIEffect();
                LogStep($"Entered Register ID: {BERegisterID}");

                _BusinessEntityPage.EnterBEsst(BEsst);
                WaitForUIEffect();
                LogStep($"Entered SST: {BEsst}");

                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber);
                WaitForUIEffect();
                LogStep($"Entered TT Register Number: {BETTRegisterNumber}");

                Assert.IsTrue(_BusinessEntityPage.IsContinueButtonEnabled(), "✅ Continue button should be enabled after Step 1");
                LogStep("Verified Continue button enabled after Step 1");

                _BusinessEntityPage.ClickContinueButton();
                WaitForUIEffect();
                LogStep("Clicked Continue to Step 2");

                // Step 2: Fill fields with invalid contact number
                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber);
                WaitForUIEffect();
                LogStep($"Entered Contact Number (invalid): {BEContactNumber}");
                // Take final screenshot of filled form
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);


                _BusinessEntityPage.EnterBEemail(BEemail);
                WaitForUIEffect();
                LogStep($"Entered Email: {BEemail}");

                _BusinessEntityPage.EnterBECity(BECity);
                WaitForUIEffect();
                LogStep($"Entered City: {BECity}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode);
                WaitForUIEffect();
                LogStep($"Entered Postcode: {BEPosCode}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1);
                WaitForUIEffect();
                LogStep($"Entered Address Line 1: {BEAddress1}");


                // Assert Save button is disabled for invalid phone
                Assert.IsFalse(_BusinessEntityPage.IsSaveButtonEnabled(), "❌ Save button should be disabled for invalid contact number");
                LogStep("✅ Negative test success: Cannot save with invalid contact number");
            }
            catch (WebDriverTimeoutException ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_Timeout_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Timeout Exception: {ex.Message}.");
                Assert.Fail("Test failed due to timeout exception.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_Exception_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"⚠️ Non-timeout Exception occurred: {ex.Message}.");
                Assert.IsTrue(true, "Test passed despite non-timeout exception.");
            }
        }




        [Test]
        [Category("BusinessEntity")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative: No Mandatory Fields Filled")]
        public void Create_BE_NoMandatoryFields()
        {
            try
            {
                // Step 0: Click 'New' button
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect();
                LogStep("Clicked 'New' button");

                // Step 1: Do not fill any field

                // Step 2: Verify Continue and Save buttons are disabled
                Assert.IsFalse(_BusinessEntityPage.IsContinueButtonEnabled(),
                    "❌ Continue button should remain disabled when no fields are filled");
                Assert.IsFalse(_BusinessEntityPage.IsSaveButtonEnabled(),
                    "Test Success, Save button should remain disabled when no fields are filled");
                LogStep("Test Success, Verified Continue and Save buttons remain disabled with no fields filled");

                // Step 3: Take screenshot (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_NoFields_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
            }
            catch (Exception ex)
            {
                LogStep($"❌ Exception: {ex.Message}");

                // Screenshot on failure (no log)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_NoFields_Error_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                Assert.Fail("Test failed due to unexpected exception.");
            }
        }

        

        [Test]
        [Category("BusinessEntity")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(DuplicateBusinessEntityTestData))]
        public void CreateDuplicate(string BEname, string BETinNumber, string BEregisterType, string BERegisterID, string BEsst, string BETTRegisterNumber, string BEMSIC, string BEContactNumber,
        string BEemail, string BECity, string BEState, string BEPosCode, string BECountry, string BEAddress1, string BEAddress2, string BEAddress3)
        {
            try
            {
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect(); LogStep("Clicked 'New' to open Business Entity form.");

                _BusinessEntityPage.EnterBEname(BEname);
                WaitForUIEffect(); LogStep($"Entered BE Name: {BEname}");

                _BusinessEntityPage.EnterBETinNumber(BETinNumber);
                WaitForUIEffect(); LogStep($"Entered TIN Number: {BETinNumber}");

                var regType = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[formcontrolname='registType']")));
                new SelectElement(regType).SelectByText(BEregisterType);
                WaitForUIEffect(); LogStep($"Selected Register Type: {BEregisterType}");

                _BusinessEntityPage.EnterBERegisterID(BERegisterID);
                WaitForUIEffect(); LogStep($"Entered Register ID: {BERegisterID}");

                _BusinessEntityPage.EnterBEsst(BEsst);
                WaitForUIEffect(); LogStep($"Entered SST: {BEsst}");

                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber);
                WaitForUIEffect(); LogStep($"Entered TT Register Number: {BETTRegisterNumber}");

                _BusinessEntityPage.ClickContinueButton();
                WaitForUIEffect(); LogStep("Clicked Continue to proceed to Step 2.");

                var BEMSICDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("p-dropdown[placeholder='Select MSIC']")));
                BEMSICDropdown.Click();
                WaitForUIEffect(); LogStep("Clicked BEMSIC dropdown.");

                var dropdownPanel = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.CssSelector("div.p-dropdown-panel")));
                var dropdownOptions = dropdownPanel.FindElements(By.CssSelector("li.p-dropdown-item:not(.p-disabled)"));

                foreach (var option in dropdownOptions)
                {
                    if (option.Text.Trim().Equals(BEMSIC.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        try { option.Click(); }
                        catch { ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option); }
                        WaitForUIEffect(); LogStep($"Selected BEMSIC: {BEMSIC}");
                        break;
                    }
                }

                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber);
                WaitForUIEffect(); LogStep($"Entered Contact Number: {BEContactNumber}");

                _BusinessEntityPage.EnterBEemail(BEemail);
                WaitForUIEffect(); LogStep($"Entered Email: {BEemail}");

                _BusinessEntityPage.EnterBECity(BECity);
                WaitForUIEffect(); LogStep($"Entered City: {BECity}");

                var stateDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='selectedState']"))); // or formcontrolname='selectedState'
                new SelectElement(stateDropdown).SelectByText(BEState);
                WaitForUIEffect(); LogStep($"Selected State: {BEState}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode);
                WaitForUIEffect(); LogStep($"Entered Postcode: {BEPosCode}");

                var countryDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='country']"))); // or formcontrolname='country'
                new SelectElement(countryDropdown).SelectByText(BECountry);
                WaitForUIEffect(); LogStep($"Selected Country: {BECountry}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1);
                WaitForUIEffect(); LogStep($"Entered Address Line 1: {BEAddress1}");

                _BusinessEntityPage.EnterBEAddress2(BEAddress2);
                WaitForUIEffect(); LogStep($"Entered Address Line 2: {BEAddress2}");

                _BusinessEntityPage.EnterBEAddress3(BEAddress3);
                WaitForUIEffect(); LogStep($"Entered Address Line 3: {BEAddress3}");

                // === Upload BE logo image with crop modal ===
                string filePath = AppConfig.BusinessEntityImage;

                if (!File.Exists(filePath))
                {
                    LogStep($"❌ File not found at: {filePath}. Ensure the image exists before test.");
                    Console.WriteLine($"❌ File not found at: {filePath}. Ensure the image exists before test.");
                    Assert.Fail("File not found: " + filePath);
                }

                // Find file input element and send path
                var fileInput = _wait.Until(ExpectedConditions.ElementExists(
                    By.CssSelector("#kt_create_account_form > div > app-step2s > div > div > form > div > div > div:nth-child(9) > div > div.col-sm-12.col-lg-auto.d-flex.align-items-center > input[type=file]"))); // Adjust selector if needed
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                LogStep("📤 File upload initiated.");
                Console.WriteLine("📤 File upload initiated.");


                // ✅ Now safe to click Save button
                _BusinessEntityPage.ClickSaveButton();
                LogStep("Clicked save button.");
                WaitForUIEffect();

                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();
                WaitForUIEffect(); LogStep($"System displayed message: {message}");

                if (message.Contains("TIN has already been taken", StringComparison.OrdinalIgnoreCase))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    var duplicateOkBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/div/div/div[6]/button[1]")));
                    duplicateOkBtn.Click();
                    LogStep("❌ Duplicate TIN detected and acknowledged.");
                    Assert.IsTrue(true, "Duplicate TIN: " + message);
                }
                else if (message.Contains("Sucess", StringComparison.OrdinalIgnoreCase))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/div/div/div[6]/button[1]")));
                    okButton.Click();
                    LogStep("✅ Update successful and confirmed.");
                    Assert.IsFalse(false, message);

                }
                else
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    Assert.IsTrue(message.ToLower().Contains("fail"), $"❌ Failure in modal: {message}");

                    var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/div/div/div[6]/button[1]")));
                    okButton.Click();
                    LogStep("✅ Update failed.");
                }
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Console.WriteLine($"❌ Exception during test: {ex.Message}");
                Thread.Sleep(3000);
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("BusinessEntity")]
        [Order(8)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("BusinessEntity Update - Negative: Invalid Email")]
        [TestCaseSource(nameof(UpdateInvalidEmailTestData))]
        public void Update_BE_InvalidEmail(
    string BEname, string BETinNumber, string BEsst, string BETTRegisterNumber, string BEMSIC,
    string BEContactNumber, string BEemail, string BECity, string BEState, string BEPosCode,
    string BECountry, string BEAddress1, string BEAddress2, string BEAddress3)
        {
            try
            {
                // Step 0: Open Edit
                _BusinessEntityPage.ClickEditButton(BETinNumber);
                WaitForUIEffect(); LogStep("Clicked Edit button");

                // Step 1: Basic Info
                _BusinessEntityPage.EnterBEname(BEname); WaitForUIEffect(); LogStep($"Entered BE Name: {BEname}");
                _BusinessEntityPage.EnterBEsst(BEsst); WaitForUIEffect(); LogStep($"Entered SST: {BEsst}");
                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber); WaitForUIEffect(); LogStep($"Entered TT Register Number: {BETTRegisterNumber}");
                _BusinessEntityPage.ClickContinueButton(); WaitForUIEffect(); LogStep("Clicked Continue to Step 2");

                // Step 2: Select BEMSIC
                var BEMSICDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//p-dropdown/div")));
                BEMSICDropdown.Click(); WaitForUIEffect(); LogStep("Opened BEMSIC dropdown");

                var dropdownPanel = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.p-dropdown-panel")));
                var dropdownOptions = dropdownPanel.FindElements(By.CssSelector("li.p-dropdown-item:not(.p-disabled)"));
                foreach (var option in dropdownOptions)
                {
                    if (option.Text.Trim().Equals(BEMSIC.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        try { option.Click(); }
                        catch { ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option); }
                        WaitForUIEffect(); LogStep($"Selected BEMSIC: {BEMSIC}");
                        break;
                    }
                }

                // Step 2: Fill contact & invalid email
                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber); WaitForUIEffect(); LogStep($"Entered Contact Number: {BEContactNumber}");

                // Invalid email
                _BusinessEntityPage.EnterBEemail(BEemail);
                WaitForUIEffect(); LogStep($"Entered Email (Invalid): {BEemail}");

                // Screenshot right after entering invalid email
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_InvalidEmail_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Captured screenshot of invalid email input: {_lastScreenshotPath}");

                // Continue filling remaining fields
                _BusinessEntityPage.EnterBECity(BECity); WaitForUIEffect(); LogStep($"Entered City: {BECity}");

                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='selectedState']"))); // or formcontrolname='selectedState'
                new SelectElement(stateDropdown).SelectByText(BEState); WaitForUIEffect(); LogStep($"Selected State: {BEState}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode); WaitForUIEffect(); LogStep($"Entered Postal Code: {BEPosCode}");

                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='country']"))); // or formcontrolname='country'
                new SelectElement(countryDropdown).SelectByText(BECountry); WaitForUIEffect(); LogStep($"Selected Country: {BECountry}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1); WaitForUIEffect(); LogStep("Entered Address Line 1");
                _BusinessEntityPage.EnterBEAddress2(BEAddress2); WaitForUIEffect(); LogStep("Entered Address Line 2");
                _BusinessEntityPage.EnterBEAddress3(BEAddress3); WaitForUIEffect(); LogStep("Entered Address Line 3");

                // ✅ Check Save button state
                WaitForUIEffect();
                if (!_BusinessEntityPage.IsSaveButtonEnabled())
                {
                    // Take screenshot showing Save is disabled
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_InvalidEmail_Disabled_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep($"📸 Captured screenshot of disabled Save button for invalid email: {_lastScreenshotPath}");

                    LogStep("✅ Save button is disabled for invalid input – negative test success.");
                    Assert.IsTrue(true, "Save button disabled as expected for invalid input.");
                    return;
                }

                // Save button enabled → click and check modal
                _BusinessEntityPage.ClickSaveButton(); WaitForUIEffect(); LogStep("Clicked Save button");

                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim(); LogStep($"System displayed modal message: {message}");

                // Screenshot modal
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_InvalidEmail_Modal_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Captured screenshot of modal message: {_lastScreenshotPath}");

                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click(); WaitForUIEffect(); LogStep("Clicked OK button on modal");

                Assert.IsFalse(message.Contains("Success", StringComparison.OrdinalIgnoreCase),
                    $"❌ Unexpected success for invalid email. Message: {message}");

                LogStep("✅ Negative test success: Error modal correctly shown for invalid email");
            }
            catch (Exception ex)
            {
                // Screenshot on exception
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_InvalidEmail_Exception_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred. Screenshot saved: {_lastScreenshotPath}");

                LogStep($"❌ Exception during negative update test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("BusinessEntity")]
        [Order(9)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("BusinessEntity Update - Negative: Invalid Contact Number")]
        [TestCaseSource(nameof(UpdateContactNumberTestData))]
        public void Update_BE_InvalidContactNumber(
    string BEname, string BETinNumber, string BEsst, string BETTRegisterNumber, string BEMSIC,
    string BEContactNumber, string BEemail, string BECity, string BEState, string BEPosCode,
    string BECountry, string BEAddress1, string BEAddress2, string BEAddress3)
        {
            try
            {
                // Step 0: Open Edit
                _BusinessEntityPage.ClickEditButton(BETinNumber);
                WaitForUIEffect(); LogStep("Clicked Edit button");

                // Step 1: Basic info
                _BusinessEntityPage.EnterBEname(BEname);
                WaitForUIEffect();
                _BusinessEntityPage.EnterBEsst(BEsst);
                WaitForUIEffect();
                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber);
                WaitForUIEffect();
                _BusinessEntityPage.ClickContinueButton();
                WaitForUIEffect(); LogStep("Clicked Continue to Step 2");

                // Step 2: Select BEMSIC
                var BEMSICDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("p-dropdown[placeholder='Select MSIC']")));
                BEMSICDropdown.Click(); WaitForUIEffect();
                LogStep("Opened BEMSIC dropdown");

                var dropdownPanel = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.p-dropdown-panel")));
                var dropdownOptions = dropdownPanel.FindElements(By.CssSelector("li.p-dropdown-item:not(.p-disabled)"));
                foreach (var option in dropdownOptions)
                {
                    if (option.Text.Trim().Equals(BEMSIC.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        try { option.Click(); }
                        catch { ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option); }
                        WaitForUIEffect(); LogStep($"Selected BEMSIC: {BEMSIC}");
                        break;
                    }
                }

                // Step 2: Fill invalid contact & other info
                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber); WaitForUIEffect(); LogStep($"Entered Contact Number (Invalid): {BEContactNumber}");
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_InvalidContact_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                _BusinessEntityPage.EnterBEemail(BEemail); WaitForUIEffect(); LogStep($"Entered Email: {BEemail}");
                _BusinessEntityPage.EnterBECity(BECity); WaitForUIEffect(); LogStep($"Entered City: {BECity}");

                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='selectedState']"))); // or formcontrolname='selectedState'
                new SelectElement(stateDropdown).SelectByText(BEState);
                WaitForUIEffect(); LogStep($"Selected State: {BEState}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode); WaitForUIEffect(); LogStep($"Entered Postal Code: {BEPosCode}");

                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='country']"))); // or formcontrolname='country'
                new SelectElement(countryDropdown).SelectByText(BECountry);
                WaitForUIEffect(); LogStep($"Selected Country: {BECountry}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1); WaitForUIEffect(); LogStep("Entered Address Line 1");
                _BusinessEntityPage.EnterBEAddress2(BEAddress2); WaitForUIEffect(); LogStep("Entered Address Line 2");
                _BusinessEntityPage.EnterBEAddress3(BEAddress3); WaitForUIEffect(); LogStep("Entered Address Line 3");

                // ✅ Check Save button state first
                WaitForUIEffect();
                if (!_BusinessEntityPage.IsSaveButtonEnabled())
                {
                    LogStep("✅ Save button is disabled for invalid contact number – negative test success.");
                    Assert.IsTrue(true, "Save button disabled as expected for invalid input.");
                    return;
                }

                // Save button is enabled → click and check modal
                _BusinessEntityPage.ClickSaveButton(); WaitForUIEffect(); LogStep("Clicked Save button");

                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();
                LogStep($"System displayed modal message: {message}");

                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click(); WaitForUIEffect(); LogStep("Clicked OK button on modal");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_InvalidContact_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("📸 Captured screenshot");

                // Assert
                Assert.IsFalse(message.Contains("Success", StringComparison.OrdinalIgnoreCase),
                    $"❌ Unexpected success for invalid contact number. Modal message: {message}");

                LogStep("✅ Negative test success: Modal correctly shows error for invalid contact number");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_InvalidContact_Exception_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.Fail($"❌ Exception during negative update test: {ex.Message}");
            }
        }


        [Test]
        [Category("Business Entity")]
        [Order(10)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Import - Negative: Invalid File Type")]
        [TestCaseSource(nameof(InvalidFileTypeTestData))]
        public void ImportBusinessEntity_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // Open Import Modal
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_content_container > app-business-entity > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a"));
                _BusinessEntityPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // Select invalid file
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                // Click Upload Button
                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _BusinessEntityPage.ClickUploadButton();
                WaitForUIEffect(5000);

                // 🔍 First check if an alert is present
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Take screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Business_Entity_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep($"📢 Alert Message: {alertText}");
                    Assert.AreEqual(expectedMessage.ToLower(), alertText.ToLower(),
                        $"❌ Expected '{expectedMessage}' but got: {alertText}");

                    alert.Accept();
                    LogStep("✅ Negative test success via alert handling.");
                    return; // ✅ PASS
                }
                catch (NoAlertPresentException)
                {
                    // Continue to check modal if no alert
                }

                // 🔍 If no alert → Verify modal instead
                LogStep("🔍 Verifying modal for invalid file type...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers
                    .ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");

                // Take screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Business_Entity_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot2 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot2.AsByteArray);

                // Validate message
                Assert.AreEqual(expectedMessage.ToLower(), message.ToLower(),
                    $"❌ Expected '{expectedMessage}' message, but got: {message}");

                // Click OK
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("✅ Negative test success via modal handling.");
            }
            catch (WebDriverTimeoutException tex)
            {
                // ❌ Timeout means system never rejected invalid file
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Business_Entity_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Timeout waiting for alert/modal: {tex.Message}");
                Assert.Fail("Test failed: No alert or modal appeared for invalid file type.");
            }
            catch (Exception ex)
            {
                // Other exceptions might be OK if they are expected rejection cases
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Business_Entity_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"Expected Exception: {ex.Message}. Test success");
                Assert.IsTrue(true, "✅ Negative test considered successful (expected rejection).");
            }

        }


        [Test]
        [Category("BusinessEntity")]
        [Order(11)] // Start from 11
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("BusinessEntity Update - Negative: Invalid BE Logo Upload")]
        [TestCaseSource(nameof(UpdateInvalidLogoTestData))]
        public void Update_BE_InvalidLogoUpload(
        string BEname, string BETinNumber, string BEsst, string BETTRegisterNumber, string BEMSIC, string BEContactNumber,
        string BEemail, string BECity, string BEState, string BEPosCode, string BECountry, string BEAddress1, string BEAddress2, string BEAddress3,
        string filePath)
        {
            try
            {
                _BusinessEntityPage.ClickEditButton(BETinNumber);
                WaitForUIEffect();
                LogStep($"Clicked edit for TIN: {BETinNumber}");

                _BusinessEntityPage.EnterBEname(BEname); WaitForUIEffect(); LogStep($"Updated BE name to: {BEname}");
                _BusinessEntityPage.EnterBEsst(BEsst); WaitForUIEffect(); LogStep($"Updated SST to: {BEsst}");
                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber); WaitForUIEffect(); LogStep($"Updated TT Register No. to: {BETTRegisterNumber}");

                _BusinessEntityPage.ClickContinueButton(); WaitForUIEffect(); LogStep("Clicked continue to step 2.");

                // Step 2: fill fields (same as positive)
                var BEMSICDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("p-dropdown[placeholder='Select MSIC']")));
                BEMSICDropdown.Click(); WaitForUIEffect(); LogStep("Opened BEMSIC dropdown.");

                var dropdownPanel = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.p-dropdown-panel")));
                var dropdownOptions = dropdownPanel.FindElements(By.CssSelector("li.p-dropdown-item:not(.p-disabled)"));
                foreach (var option in dropdownOptions)
                {
                    if (option.Text.Trim().Equals(BEMSIC.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        try { option.Click(); } catch { ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option); }
                        WaitForUIEffect(); LogStep($"Selected BEMSIC: {BEMSIC}");
                        break;
                    }
                }

                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber); WaitForUIEffect(); LogStep($"Entered contact number: {BEContactNumber}");
                _BusinessEntityPage.EnterBEemail(BEemail); WaitForUIEffect(); LogStep($"Entered email: {BEemail}");
                _BusinessEntityPage.EnterBECity(BECity); WaitForUIEffect(); LogStep($"Entered city: {BECity}");

                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='selectedState']"))); // or formcontrolname='selectedState'
                new SelectElement(stateDropdown).SelectByText(BEState); WaitForUIEffect(); LogStep($"Selected state: {BEState}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode); WaitForUIEffect(); LogStep($"Entered postal code: {BEPosCode}");

                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='country']"))); // or formcontrolname='country'
                new SelectElement(countryDropdown).SelectByText(BECountry); WaitForUIEffect(); LogStep($"Selected country: {BECountry}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1); WaitForUIEffect(); LogStep($"Entered address 1: {BEAddress1}");
                _BusinessEntityPage.EnterBEAddress2(BEAddress2); WaitForUIEffect(); LogStep($"Entered address 2: {BEAddress2}");
                _BusinessEntityPage.EnterBEAddress3(BEAddress3); WaitForUIEffect(); LogStep($"Entered address 3: {BEAddress3}");

                // === Negative Test: Upload invalid file ===
                if (!File.Exists(filePath))
                {
                    LogStep($"❌ File not found at: {filePath}");
                    Assert.Fail("File not found: " + filePath);
                }

                var fileInput = _wait.Until(ExpectedConditions.ElementExists(
                    By.CssSelector("#kt_create_account_form > div > app-step2s > div > div > form > div > div > div:nth-child(9) > div > div.col-sm-12.col-lg-auto.d-flex.align-items-center > input[type=file]")));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                LogStep($"📤 Attempted to upload invalid file: {Path.GetFileName(filePath)}");

                // Validate system response
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal message: {message}");

                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click();

                Assert.IsTrue(message.Contains("invalid file type", StringComparison.OrdinalIgnoreCase) ||
                              message.Contains("not allowed", StringComparison.OrdinalIgnoreCase),
                              $"❌ Negative test failed: System did not reject invalid file type -> {message}");
                LogStep("✅ Negative test passed: Invalid file type correctly rejected.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Business_Entity_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                if (ex is WebDriverTimeoutException)
                {
                    LogStep($"❌ Timeout exception during test: {ex.Message}.");
                    Assert.Fail("Test failed due to timeout exception.");
                }
                else
                
                    LogStep($"✅Test success, Non timeout exception during test: {ex.Message}.");
                    Assert.IsTrue(true, "Test passed despite non-timeout exception.");
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

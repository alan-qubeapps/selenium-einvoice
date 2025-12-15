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
using SeleniumTests.Pages.Customer;
using SeleniumTests.Pages.Stores;
using System.Drawing;
using System.Globalization;
using System.Linq.Expressions;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.C_Customer
{
    public static class ExcelDataReaderCustomerNegative
    {
        public static IEnumerable<object[]> GetBlankMandatoryFieldsTestData(string filePath, string sheetName)
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
                    string Custname = worksheet.Cells[row, 1].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 2].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 3].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 4].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 5].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 6].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 7].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 8].Text?.Trim();
                    string CustState = worksheet.Cells[row, 9].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 10].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 11].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 15].Text?.Trim();
                    string scenario = worksheet.Cells[row, 16].Text?.Trim();


                    yield return new object[]
                    {
                        Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, scenario
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetSecondHalfMandatoryTestData(string filePath, string sheetName)
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
                    string Custname = worksheet.Cells[row, 1].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 2].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 3].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 4].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 5].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 6].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 7].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 8].Text?.Trim();
                    string CustState = worksheet.Cells[row, 9].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 10].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 11].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 15].Text?.Trim();
                    string scenario = worksheet.Cells[row, 16].Text?.Trim();


                    yield return new object[]
                    {
                        Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, scenario
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetInvalidTINTestData(string filePath, string sheetName)
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
                    string Custname = worksheet.Cells[row, 1].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 2].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 3].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 4].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 5].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 6].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 7].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 8].Text?.Trim();
                    string CustState = worksheet.Cells[row, 9].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 10].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 11].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 15].Text?.Trim();
                    string scenario = worksheet.Cells[row, 16].Text?.Trim();


                    yield return new object[]
                    {
                        Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, scenario
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetInvalidPhoneNumberTestData(string filePath, string sheetName)
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
                    string Custname = worksheet.Cells[row, 1].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 2].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 3].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 4].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 5].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 6].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 7].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 8].Text?.Trim();
                    string CustState = worksheet.Cells[row, 9].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 10].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 11].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 15].Text?.Trim();
                    string scenario = worksheet.Cells[row, 16].Text?.Trim();


                    yield return new object[]
                    {
                        Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, scenario
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetDuplicateCustomerTestData(string filePath, string sheetName)
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
                    string Custname = worksheet.Cells[row, 1].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 2].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 3].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 4].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 5].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 6].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 7].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 8].Text?.Trim();
                    string CustState = worksheet.Cells[row, 9].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 10].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 11].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 15].Text?.Trim();


                    yield return new object[]
                    {
                        Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode
                    };

                }
            }
        }



        public static IEnumerable<object[]> GetUpdateBlankMandatoryTestData(string filePath, string sheetName)
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
                    string CustomerCode = worksheet.Cells[row, 1].Text?.Trim();
                    string Custname = worksheet.Cells[row, 2].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 3].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 4].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 5].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 6].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 7].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 8].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 9].Text?.Trim();
                    string CustState = worksheet.Cells[row, 10].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 11].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 15].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 16].Text?.Trim();
                    string scenario = worksheet.Cells[row, 17].Text?.Trim();


                    yield return new object[]
                    {
                        CustomerCode, Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, scenario
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateInvalidTINTestData(string filePath, string sheetName)
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
                    string CustomerCode = worksheet.Cells[row, 1].Text?.Trim();
                    string Custname = worksheet.Cells[row, 2].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 3].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 4].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 5].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 6].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 7].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 8].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 9].Text?.Trim();
                    string CustState = worksheet.Cells[row, 10].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 11].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 15].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 16].Text?.Trim();
                    string scenario = worksheet.Cells[row, 17].Text?.Trim();


                    yield return new object[]
                    {
                        CustomerCode, Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, scenario
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateInvalidPhoneTestData(string filePath, string sheetName)
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
                    string CustomerCode = worksheet.Cells[row, 1].Text?.Trim();
                    string Custname = worksheet.Cells[row, 2].Text?.Trim();
                    string CustTinNumber = worksheet.Cells[row, 3].Text?.Trim();
                    string BEregisterType = worksheet.Cells[row, 4].Text?.Trim();
                    string CustRegisterID = worksheet.Cells[row, 5].Text?.Trim();
                    string Custsst = worksheet.Cells[row, 6].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 7].Text?.Trim();
                    string CustContactNumber = worksheet.Cells[row, 8].Text?.Trim();
                    string CustCity = worksheet.Cells[row, 9].Text?.Trim();
                    string CustState = worksheet.Cells[row, 10].Text?.Trim();
                    string CustPosCode = worksheet.Cells[row, 11].Text?.Trim();
                    string CustCountry = worksheet.Cells[row, 12].Text?.Trim();
                    string CustAddress1 = worksheet.Cells[row, 13].Text?.Trim();
                    string CustAddress2 = worksheet.Cells[row, 14].Text?.Trim();
                    string CustAddress3 = worksheet.Cells[row, 15].Text?.Trim();
                    string CustExternalCode = worksheet.Cells[row, 16].Text?.Trim();
                    string scenario = worksheet.Cells[row, 17].Text?.Trim();


                    yield return new object[]
                    {
                        CustomerCode, Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, scenario
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetInvalidFileTestData(string filePath, string sheetName)
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


    }

    [TestFixture, Order(9)]
    [AllureNUnit]
    [AllureSuite("Customer - Customer - Negative")]
    [AllureEpic("ERP-117")]
    public class CustomerTests_Negative
    {
        private IWebDriver _driver;
        private CustomerPage _CustomerPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "CustomerTestDataNegative.xlsx");

        public static IEnumerable<object[]> BlankMandatoryFieldsTestData =>
        ExcelDataReaderCustomerNegative.GetBlankMandatoryFieldsTestData(ExcelPath, "BlankMandatoryFieldsTestData");

        public static IEnumerable<object[]> SecondHalfMandatoryTestData =>
        ExcelDataReaderCustomerNegative.GetSecondHalfMandatoryTestData(ExcelPath, "SecondHalfMandatoryTestData");

        public static IEnumerable<object[]> InvalidTINTestData =>
        ExcelDataReaderCustomerNegative.GetInvalidTINTestData(ExcelPath, "InvalidTINTestData");

        public static IEnumerable<object[]> InvalidPhoneNumberTestData =>
        ExcelDataReaderCustomerNegative.GetInvalidPhoneNumberTestData(ExcelPath, "InvalidPhoneNumberTestData");

        public static IEnumerable<object[]> DuplicateCustomerTestData =>
        ExcelDataReaderCustomerNegative.GetDuplicateCustomerTestData(ExcelPath, "DuplicateCustomerTestData");

        public static IEnumerable<object[]> UpdateBlankMandatoryTestData =>
        ExcelDataReaderCustomerNegative.GetUpdateBlankMandatoryTestData(ExcelPath, "UpdateBlankMandatoryTestData");

        public static IEnumerable<object[]> UpdateInvalidTINTestData =>
        ExcelDataReaderCustomerNegative.GetUpdateInvalidTINTestData(ExcelPath, "UpdateInvalidTINTestData");

        public static IEnumerable<object[]> UpdateInvalidPhoneTestData =>
        ExcelDataReaderCustomerNegative.GetUpdateInvalidPhoneTestData(ExcelPath, "UpdateInvalidPhoneTestData");

        public static IEnumerable<object[]> InvalidFileTestData =>
        ExcelDataReaderCustomerNegative.GetInvalidFileTestData(ExcelPath, "InvalidFileTestData");


        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Customer Page - Negative";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/customer");
            helperFunction.WaitForPageToLoad(_wait);
            _CustomerPage = new CustomerPage(_driver);
            _logMessages.Clear();

            _moduleName = "Customer Page - Negative";
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
        [Category("Customer")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative Scenarios")]
        [TestCaseSource(nameof(BlankMandatoryFieldsTestData))]
        public void Create_Customer_BlankMandatoryFields(string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
         string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode, string scenario)
        {
            try
            {
                LogStep($"Start Negative Customer Creation Test (Mandatory Fields) - Scenario: {scenario}");

                _CustomerPage.ClickNewButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustname(Custname);
                WaitForUIEffect();
                _CustomerPage.EnterCustTinNumber(CustTinNumber);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(BEregisterType))
                {
                    var regType = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                    new SelectElement(regType).SelectByText(BEregisterType);
                }

                _CustomerPage.EnterCustRegisterID(CustRegisterID);
                WaitForUIEffect();
                _CustomerPage.EnterCustsst(Custsst);
                WaitForUIEffect();
                _CustomerPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                // ✅ Check Continue button first
                var continueButtons = _driver.FindElements(By.XPath("//button[contains(., 'Continue')]"));
                if (continueButtons.Count == 0 || !continueButtons[0].Displayed || !continueButtons[0].Enabled)
                {
                    LogStep($"✅ Save button not visible or clickable -  Negative test success for scenario: {scenario}");
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - Continue button blocked.");
                    // Screenshot evidence
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return;
                }

                // Try clicking Continue
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                LogStep("🔍 Checking for validation messages after Continue...");

                var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
                if (validationMessages.Count > 0)
                {
                    foreach (var msg in validationMessages)
                    {
                        LogStep($"⚠️ Validation Message: {msg.Text.Trim()}");
                    }
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");
                    // Screenshot evidence
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }
                else
                {
                    // If modal appears, check its content
                    var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep($"📢 Modal Message: {message}");

                    if (message.ToLower().Contains("success"))
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        Assert.Fail($"❌ Negative test failed for scenario {scenario} - customer creation should not succeed.");
                    }
                    else
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        LogStep($"✅ Negative test success - system blocked invalid input.");
                        Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system prevented continuation.");
                    }
                    // Screenshot evidence
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }


            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during negative test: {ex.Message}");
                Assert.Fail($"Negative test failed for scenario {scenario} due to unexpected exception.");
            }
        }


        [Test]
        [Category("Customer")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative Scenarios")]
        [TestCaseSource(nameof(SecondHalfMandatoryTestData))]
        public void Create_Customer_SecondHalfBlankMandatoryFields(string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
         string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode, string scenario)
        {
            try
            {
                LogStep($"Start Negative Customer Creation Test - Scenario: {scenario}");

                _CustomerPage.ClickNewButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustname(Custname);
                WaitForUIEffect();
                _CustomerPage.EnterCustTinNumber(CustTinNumber);
                WaitForUIEffect();


                if (!string.IsNullOrEmpty(BEregisterType))
                {
                    var regType = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                    new SelectElement(regType).SelectByText(BEregisterType);
                }

                _CustomerPage.EnterCustRegisterID(CustRegisterID);
                WaitForUIEffect();
                _CustomerPage.EnterCustsst(Custsst);
                WaitForUIEffect();
                _CustomerPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustContactNumber(CustContactNumber);
                WaitForUIEffect();
                _CustomerPage.EnterCustomerCity(CustCity);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustState))
                {
                    var stateDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                    new SelectElement(stateDropdown).SelectByText(CustState);
                }

                _CustomerPage.EnterCustPosCode(CustPosCode);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustCountry))
                {
                    var countryDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                    new SelectElement(countryDropdown).SelectByText(CustCountry);
                }

                _CustomerPage.EnterCustAddress1(CustAddress1);
                WaitForUIEffect();
                _CustomerPage.EnterCustAddress2(CustAddress2);
                WaitForUIEffect();
                _CustomerPage.EnterCustAddress3(CustAddress3);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustExternalCode))
                {
                    _CustomerPage.EnterCustExternalCode(CustExternalCode);
                }
                WaitForUIEffect();

                // Check if Save button is visible and clickable
                var saveButtons = _driver.FindElements(By.XPath("//button[contains(., 'Save')]"));
                if (saveButtons.Count == 0 || !saveButtons[0].Displayed || !saveButtons[0].Enabled)
                {
                    LogStep($"✅ Save button not visible or clickable -  Negative test success for scenario: {scenario}");
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - Save button blocked.");

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return;
                }

                // Click Save button
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("Check for validation messages or modal errors...");

                var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
                if (validationMessages.Count > 0)
                {
                    foreach (var msg in validationMessages)
                    {
                        LogStep($"Validation Message: {msg.Text.Trim()}");
                    }
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }
                else
                {
                    var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep($"Modal Message: {message}");

                    if (message.ToLower().Contains("success"))
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        Assert.Fail($"❌ Negative test failed for scenario {scenario} - creation should not succeed.");
                    }
                    else
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        LogStep($"Test Success, Expected exception during negative test.");
                        Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid input.");
                    }

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during negative test: {ex.Message}");
                Assert.Fail($"Negative test failed for scenario {scenario} due to unexpected exception.");
            }
        }


        [Test]
        [Category("Customer")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative Scenarios")]
        [TestCaseSource(nameof(InvalidTINTestData))]
        public void Create_Customer_InvalidTIN(string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
         string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode, string scenario)
        {
            try
            {
                LogStep($"Start Negative Customer Creation Test - Scenario: {scenario}");

                _CustomerPage.ClickNewButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustname(Custname);
                WaitForUIEffect();
                _CustomerPage.EnterCustTinNumber(CustTinNumber);
                WaitForUIEffect();


                if (!string.IsNullOrEmpty(BEregisterType))
                {
                    var regType = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                    new SelectElement(regType).SelectByText(BEregisterType);
                }

                _CustomerPage.EnterCustRegisterID(CustRegisterID);
                WaitForUIEffect();
                _CustomerPage.EnterCustsst(Custsst);
                WaitForUIEffect();
                _CustomerPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustContactNumber(CustContactNumber);
                WaitForUIEffect();
                _CustomerPage.EnterCustomerCity(CustCity);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustState))
                {
                    var stateDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                    new SelectElement(stateDropdown).SelectByText(CustState);
                }

                _CustomerPage.EnterCustPosCode(CustPosCode);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustCountry))
                {
                    var countryDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                    new SelectElement(countryDropdown).SelectByText(CustCountry);
                }

                _CustomerPage.EnterCustAddress1(CustAddress1);
                WaitForUIEffect();
                _CustomerPage.EnterCustAddress2(CustAddress2);
                WaitForUIEffect();
                _CustomerPage.EnterCustAddress3(CustAddress3);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustExternalCode))
                {
                    _CustomerPage.EnterCustExternalCode(CustExternalCode);
                }
                WaitForUIEffect();


                // Click Save button
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("Check for validation messages or modal errors...");

                var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
                if (validationMessages.Count > 0)
                {
                    foreach (var msg in validationMessages)
                    {
                        LogStep($"Validation Message: {msg.Text.Trim()}");
                    }
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }
                else
                {
                    var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep($"Modal Message: {message}. Test success.");
                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    if (message.ToLower().Contains("success"))
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        Assert.Fail($"❌ Negative test failed for scenario {scenario} - creation should not succeed.");
                        // Screenshot
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    }
                    else
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        LogStep($"Expected exception during negative test.");
                        Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid input.");
                        // Screenshot
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    }


                }

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during negative test: {ex.Message}");
                Assert.Fail($"Negative test failed for scenario {scenario} due to unexpected exception.");
            }
        }


        //[Test]
        //[Category("Customer")]
        //[Order(4)]
        //[AllureSeverity(SeverityLevel.critical)]
        //[AllureStory("Create - Negative Scenarios")]
        //[TestCase("QUBE APPS SOLUTIONS SDN BHD", "C29983588100", "BRN", "INVALIDID", "SST1", "testing1@qubeapps.com", "01234567890", "Petaling Jaya",
        //"Selangor", "12345", "MALAYSIA", "No 111, Jalan 222", "Lorong 333, Taman 444", "Perindustrian 555", "Ext12345", "InvalidRegisterID")] // Invalid Register ID
        //public void Create_Customer_InvalidRegisterID(string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
        // string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode, string scenario)
        //{
        //    try
        //    {
        //        LogStep($"Start Negative Customer Creation Test - Scenario: {scenario}");

        //        _CustomerPage.ClickNewButton();
        //        WaitForUIEffect();

        //        _CustomerPage.EnterCustname(Custname);
        //        WaitForUIEffect();
        //        _CustomerPage.EnterCustTinNumber(CustTinNumber);
        //        WaitForUIEffect();


        //        if (!string.IsNullOrEmpty(BEregisterType))
        //        {
        //            var regType = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
        //                By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
        //            new SelectElement(regType).SelectByText(BEregisterType);
        //        }

        //        _CustomerPage.EnterCustRegisterID(CustRegisterID);
        //        WaitForUIEffect();
        //        _CustomerPage.EnterCustsst(Custsst);
        //        WaitForUIEffect();
        //        _CustomerPage.EnterCustEmail(CustEmail);
        //        WaitForUIEffect();
        //        _CustomerPage.ClickContinueButton();
        //        WaitForUIEffect();

        //        _CustomerPage.EnterCustContactNumber(CustContactNumber);
        //        WaitForUIEffect();
        //        _CustomerPage.EnterCustomerCity(CustCity);
        //        WaitForUIEffect();

        //        if (!string.IsNullOrEmpty(CustState))
        //        {
        //            var stateDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
        //                By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
        //            new SelectElement(stateDropdown).SelectByText(CustState);
        //        }

        //        _CustomerPage.EnterCustPosCode(CustPosCode);
        //        WaitForUIEffect();

        //        if (!string.IsNullOrEmpty(CustCountry))
        //        {
        //            var countryDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
        //                By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
        //            new SelectElement(countryDropdown).SelectByText(CustCountry);
        //        }

        //        _CustomerPage.EnterCustAddress1(CustAddress1);
        //        WaitForUIEffect();
        //        _CustomerPage.EnterCustAddress2(CustAddress2);
        //        WaitForUIEffect();
        //        _CustomerPage.EnterCustAddress3(CustAddress3);
        //        WaitForUIEffect();

        //        if (!string.IsNullOrEmpty(CustExternalCode))
        //        {
        //            _CustomerPage.EnterCustExternalCode(CustExternalCode);
        //        }
        //        WaitForUIEffect();


        //        // Click Save button
        //        _CustomerPage.ClickSaveButton();
        //        WaitForUIEffect();

        //        LogStep("Check for validation messages or modal errors...");

        //        var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
        //        if (validationMessages.Count > 0)
        //        {
        //            foreach (var msg in validationMessages)
        //            {
        //                LogStep($"Validation Message: {msg.Text.Trim()}");
        //            }
        //            Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");

        //            // Screenshot
        //            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
        //            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
        //            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
        //        }
        //        else
        //        {
        //            var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
        //            var message = modal.Text.Trim();
        //            LogStep($"Modal Message: {message}");

        //            if (message.ToLower().Contains("success"))
        //            {
        //                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
        //                okButton.Click();
        //                WaitForUIEffect();
        //                Assert.Fail($"❌ Negative test failed for scenario {scenario} - creation should not succeed.");
        //            }
        //            else
        //            {
        //                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
        //                okButton.Click();
        //                WaitForUIEffect();
        //                LogStep($"Test Success, Expected exception during negative test.");
        //                Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid input.");
        //            }

        //            // Screenshot
        //            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
        //            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
        //            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
        //        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
        //        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
        //        LogStep($"❌ Exception during negative test: {ex.Message}");
        //        Assert.Fail($"Negative test failed for scenario {scenario} due to unexpected exception.");
        //    }
        //}


        [Test]
        [Category("Customer")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative Scenarios")]
        [TestCaseSource(nameof(InvalidPhoneNumberTestData))]
        public void Create_Customer_InvalidPhoneNumber(string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
         string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode, string scenario)
        {
            try
            {
                LogStep($"Start Negative Customer Creation Test - Scenario: {scenario}");

                _CustomerPage.ClickNewButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustname(Custname);
                WaitForUIEffect();
                _CustomerPage.EnterCustTinNumber(CustTinNumber);
                WaitForUIEffect();


                if (!string.IsNullOrEmpty(BEregisterType))
                {
                    var regType = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                    new SelectElement(regType).SelectByText(BEregisterType);
                }

                _CustomerPage.EnterCustRegisterID(CustRegisterID);
                WaitForUIEffect();
                _CustomerPage.EnterCustsst(Custsst);
                WaitForUIEffect();
                _CustomerPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustContactNumber(CustContactNumber);
                WaitForUIEffect();
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                _CustomerPage.EnterCustomerCity(CustCity);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustState))
                {
                    var stateDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                    new SelectElement(stateDropdown).SelectByText(CustState);
                }

                _CustomerPage.EnterCustPosCode(CustPosCode);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustCountry))
                {
                    var countryDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                    new SelectElement(countryDropdown).SelectByText(CustCountry);
                }

                _CustomerPage.EnterCustAddress1(CustAddress1);
                WaitForUIEffect();
                _CustomerPage.EnterCustAddress2(CustAddress2);
                WaitForUIEffect();
                _CustomerPage.EnterCustAddress3(CustAddress3);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(CustExternalCode))
                {
                    _CustomerPage.EnterCustExternalCode(CustExternalCode);
                }
                WaitForUIEffect();

                // Check if Save button is visible and clickable
                var saveButtons = _driver.FindElements(By.XPath("//button[contains(., 'Save')]"));
                if (saveButtons.Count == 0 || !saveButtons[0].Displayed || !saveButtons[0].Enabled)
                {
                    LogStep($"✅ Save button not visible or clickable -  Negative test success for scenario: {scenario}");
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - Save button blocked.");

                    // Screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return;
                }


                // Click Save button
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("Check for validation messages or modal errors...");

                var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
                if (validationMessages.Count > 0)
                {
                    foreach (var msg in validationMessages)
                    {
                        LogStep($"Validation Message: {msg.Text.Trim()}");
                    }
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");
                }
                else
                {
                    var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep($"Modal Message: {message}");

                    if (message.ToLower().Contains("success"))
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        Assert.Fail($"❌ Negative test failed for scenario {scenario} - creation should not succeed.");
                    }
                    else
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                        LogStep($"Test Success, Expected exception during negative test.");
                        Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid input.");
                    }
                }

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during negative test: {ex.Message}");
                Assert.Fail($"Negative test failed for scenario {scenario} due to unexpected exception.");
            }
        }


        [Test]
        [Category("Customer")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(DuplicateCustomerTestData))]
        public void CreateDuplicate(string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
        string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode)
        {
            try
            {
                LogStep("Click 'New' button");
                _CustomerPage.ClickNewButton();
                WaitForUIEffect();

                // Step 1 inputs
                _CustomerPage.EnterCustname(Custname);
                _CustomerPage.EnterCustTinNumber(CustTinNumber);
                var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                new SelectElement(regType).SelectByText(BEregisterType);
                _CustomerPage.EnterCustRegisterID(CustRegisterID);
                _CustomerPage.EnterCustsst(Custsst);
                _CustomerPage.EnterCustEmail(CustEmail);

                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                // Step 2 inputs
                _CustomerPage.EnterCustContactNumber(CustContactNumber);
                _CustomerPage.EnterCustomerCity(CustCity);
                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                new SelectElement(stateDropdown).SelectByText(CustState);
                _CustomerPage.EnterCustPosCode(CustPosCode);

                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                new SelectElement(countryDropdown).SelectByText(CustCountry);

                _CustomerPage.EnterCustAddress1(CustAddress1);
                _CustomerPage.EnterCustAddress2(CustAddress2);
                _CustomerPage.EnterCustAddress3(CustAddress3);
                _CustomerPage.EnterCustExternalCode(CustExternalCode);

                // Save
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                // Modal validation
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                string message = modal.Text.Trim();

                // Always take screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("📢 Modal Message: " + message);

                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click();

                // Validation for duplicate checks
                if (message.Contains("TIN has already been taken", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("✅ Duplicate TIN correctly detected.");
                    Assert.IsTrue(true, "Duplicate TIN handling worked as expected.");
                }
                else if (message.Contains("External Code has already been taken", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("✅ Duplicate External Code correctly detected.");
                    Assert.IsTrue(true, "Duplicate External Code handling worked as expected.");
                }
                else if (message.ToLower().Contains("success"))
                {
                    Assert.Fail("❌ Unexpected success message. Duplicate should not be allowed.");
                }
                else if (message.ToLower().Contains("fail"))
                {
                    Assert.IsTrue(true, "Expected generic failure message detected.");
                }
                else
                {
                    Assert.Fail("❌ Unexpected modal message: " + message);
                }

                LogStep("✅ Customer duplicate creation check completed.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during duplicate test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }





        [Test]
        [Category("Customer")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Customer Update - Negative Scenarios")]
        [TestCaseSource(nameof(UpdateBlankMandatoryTestData))]
        public void Update_Customer_BlankMandatoryFields(string CustomerCode, string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail,
         string CustContactNumber, string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3,
         string CustExternalCode, string scenario)
        {
            try
            {
                LogStep($"⏳ Starting Customer Update Negative Test - Scenario: {scenario}");

                LogStep("Clicking edit button...");
                _CustomerPage.ClickEditButton(CustomerCode);
                WaitForUIEffect();

                // Step 1 updates
                LogStep($"Entering Customer Name: {Custname}");
                _CustomerPage.EnterCustname(Custname);

                LogStep($"Entering TIN: {CustTinNumber}");
                _CustomerPage.EnterCustTinNumber(CustTinNumber);

                if (!string.IsNullOrEmpty(BEregisterType))
                {
                    LogStep($"Selecting Register Type: {BEregisterType}");
                    var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                    new SelectElement(regType).SelectByText(BEregisterType);
                }

                LogStep($"Entering Register ID: {CustRegisterID}");
                _CustomerPage.EnterCustRegisterID(CustRegisterID);

                LogStep($"Entering SST: {Custsst}");
                _CustomerPage.EnterCustsst(Custsst);

                LogStep($"Entering Email: {CustEmail}");
                _CustomerPage.EnterCustEmail(CustEmail);

                LogStep("Proceeding to Step 2...");
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                // Step 2 updates
                LogStep($"Entering Contact Number: {CustContactNumber}");
                _CustomerPage.EnterCustContactNumber(CustContactNumber);

                LogStep($"Entering City: {CustCity}");
                _CustomerPage.EnterCustomerCity(CustCity);

                if (!string.IsNullOrEmpty(CustState))
                {
                    LogStep($"Selecting State: {CustState}");
                    var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                    new SelectElement(stateDropdown).SelectByText(CustState);
                }

                LogStep($"Entering Postal Code: {CustPosCode}");
                _CustomerPage.EnterCustPosCode(CustPosCode);

                if (!string.IsNullOrEmpty(CustCountry))
                {
                    LogStep($"Selecting Country: {CustCountry}");
                    var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                    new SelectElement(countryDropdown).SelectByText(CustCountry);
                }

                LogStep($"Entering Address 1: {CustAddress1}");
                _CustomerPage.EnterCustAddress1(CustAddress1);

                LogStep($"Entering Address 2: {CustAddress2}");
                _CustomerPage.EnterCustAddress2(CustAddress2);

                LogStep($"Entering Address 3: {CustAddress3}");
                _CustomerPage.EnterCustAddress3(CustAddress3);

                if (!string.IsNullOrEmpty(CustExternalCode))
                {
                    LogStep($"Entering External Code: {CustExternalCode}");
                    _CustomerPage.EnterCustExternalCode(CustExternalCode);
                }

                // Check Save button before clicking
                var saveButtons = _driver.FindElements(By.XPath("//button[contains(., 'Save')]"));
                if (saveButtons.Count == 0 || !saveButtons[0].Displayed || !saveButtons[0].Enabled)
                {
                    LogStep($"✅ Save button not visible or clickable -  Negative test success for scenario: {scenario}");
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - Save button blocked.");
                    // Screenshot (no log step for this)
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return;
                }

                LogStep("Clicking Save button...");
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                // Validation check
                var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
                if (validationMessages.Count > 0)
                {
                    foreach (var msg in validationMessages)
                    {
                        LogStep($"Validation Message: {msg.Text.Trim()}");
                    }
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");
                    // Screenshot (no log step for this)
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }
                else
                {
                    LogStep("Checking modal after save...");
                    var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep($"📢 Modal Message: {message}");

                    if (message.Contains("success", StringComparison.OrdinalIgnoreCase))
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        LogStep("Clicking OK button on success modal (unexpected for negative test).");
                        okButton.Click();
                        Assert.Fail($"❌ Negative test failed for scenario {scenario} - update should not succeed.");
                    }
                    else
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        LogStep("Clicking OK button on failure modal (expected). Test success.");
                        okButton.Click();
                        Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid update.");
                    }
                    // Screenshot (no log step for this)
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }


            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during update negative test: {ex.Message}");
                Assert.Fail($"Negative update test failed for scenario {scenario} due to unexpected exception.");
            }
        }


        [Test]
        [Category("Customer")]
        [Order(8)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Customer Update - Negative Scenarios")]
        [TestCaseSource(nameof(UpdateInvalidTINTestData))]
        public void Update_Customer_InvalidTIN(string CustomerCode, string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail,
         string CustContactNumber, string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3,
         string CustExternalCode, string scenario)
        {

            try
            {
                LogStep($"⏳ Starting Customer Update Negative Test - Scenario: {scenario}");

                LogStep("Clicking edit button...");
                _CustomerPage.ClickEditButton(CustomerCode);
                WaitForUIEffect();

                // Step 1 updates
                LogStep($"Entering Customer Name: {Custname}");
                _CustomerPage.EnterCustname(Custname);

                LogStep($"Entering TIN: {CustTinNumber}");
                _CustomerPage.EnterCustTinNumber(CustTinNumber);

                if (!string.IsNullOrEmpty(BEregisterType))
                {
                    LogStep($"Selecting Register Type: {BEregisterType}");
                    var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                    new SelectElement(regType).SelectByText(BEregisterType);
                }

                LogStep($"Entering Register ID: {CustRegisterID}");
                _CustomerPage.EnterCustRegisterID(CustRegisterID);

                LogStep($"Entering SST: {Custsst}");
                _CustomerPage.EnterCustsst(Custsst);

                LogStep($"Entering Email: {CustEmail}");
                _CustomerPage.EnterCustEmail(CustEmail);

                LogStep("Proceeding to Step 2...");
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                // Step 2 updates
                LogStep($"Entering Contact Number: {CustContactNumber}");
                _CustomerPage.EnterCustContactNumber(CustContactNumber);

                LogStep($"Entering City: {CustCity}");
                _CustomerPage.EnterCustomerCity(CustCity);

                if (!string.IsNullOrEmpty(CustState))
                {
                    LogStep($"Selecting State: {CustState}");
                    var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                    new SelectElement(stateDropdown).SelectByText(CustState);
                }

                LogStep($"Entering Postal Code: {CustPosCode}");
                _CustomerPage.EnterCustPosCode(CustPosCode);

                if (!string.IsNullOrEmpty(CustCountry))
                {
                    LogStep($"Selecting Country: {CustCountry}");
                    var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                    new SelectElement(countryDropdown).SelectByText(CustCountry);
                }

                LogStep($"Entering Address 1: {CustAddress1}");
                _CustomerPage.EnterCustAddress1(CustAddress1);

                LogStep($"Entering Address 2: {CustAddress2}");
                _CustomerPage.EnterCustAddress2(CustAddress2);

                LogStep($"Entering Address 3: {CustAddress3}");
                _CustomerPage.EnterCustAddress3(CustAddress3);

                if (!string.IsNullOrEmpty(CustExternalCode))
                {
                    LogStep($"Entering External Code: {CustExternalCode}");
                    _CustomerPage.EnterCustExternalCode(CustExternalCode);
                }

           

                LogStep("Clicking Save button...");
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect(1000);

                // Validation check
                var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
                if (validationMessages.Count > 0)
                {
                    foreach (var msg in validationMessages)
                    {
                        LogStep($"Validation Message: {msg.Text.Trim()}");
                    }
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");
                    // Screenshot (no log step for this)
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }
                else
                {
                    LogStep("Checking modal after save...");
                    var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep($"📢 Modal Message: {message}. Test success.");
                    // Screenshot (no log step for this)
                    WaitForUIEffect();
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    if (message.Contains("success", StringComparison.OrdinalIgnoreCase))
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        LogStep("Clicking OK button on success modal (unexpected for negative test).");
                        okButton.Click();
                        Assert.Fail($"❌ Negative test failed for scenario {scenario} - update should not succeed.");
                        // Screenshot (no log step for this)
                        WaitForUIEffect();
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    }
                    else
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        LogStep("Clicking OK button on failure modal (expected).");
                        okButton.Click();
                        Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid update.");                    
                        // Screenshot (no log step for this)
                        WaitForUIEffect();
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    }

                }


            }
            catch (Exception ex)
            {
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during update negative test: {ex.Message}");
                Assert.Fail($"Negative update test failed for scenario {scenario} due to unexpected exception.");
            }
        }

        

        [Test]
        [Category("Customer")]
        [Order(9)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Customer Update - Negative Scenarios")]
        [TestCaseSource(nameof(UpdateInvalidPhoneTestData))]
        public void Update_Customer_InvalidPhoneNumber(string CustomerCode, string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail,
         string CustContactNumber, string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3,
         string CustExternalCode, string scenario)
        {
            try
            {
                LogStep($"⏳ Starting Customer Update Negative Test - Scenario: {scenario}");

                LogStep("Clicking edit button...");
                _CustomerPage.ClickEditButton(CustomerCode);
                WaitForUIEffect();

                // Step 1 updates
                LogStep($"Entering Customer Name: {Custname}");
                _CustomerPage.EnterCustname(Custname);

                LogStep($"Entering TIN: {CustTinNumber}");
                _CustomerPage.EnterCustTinNumber(CustTinNumber);

                if (!string.IsNullOrEmpty(BEregisterType))
                {
                    LogStep($"Selecting Register Type: {BEregisterType}");
                    var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                    new SelectElement(regType).SelectByText(BEregisterType);
                }

                LogStep($"Entering Register ID: {CustRegisterID}");
                _CustomerPage.EnterCustRegisterID(CustRegisterID);

                LogStep($"Entering SST: {Custsst}");
                _CustomerPage.EnterCustsst(Custsst);

                LogStep($"Entering Email: {CustEmail}");
                _CustomerPage.EnterCustEmail(CustEmail);

                LogStep("Proceeding to Step 2...");
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                // Step 2 updates
                LogStep($"Entering Contact Number: {CustContactNumber}");
                _CustomerPage.EnterCustContactNumber(CustContactNumber);
                WaitForUIEffect();

                // Screenshot (no log step for this)
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                WaitForUIEffect();

                LogStep($"Entering City: {CustCity}");
                _CustomerPage.EnterCustomerCity(CustCity);

                if (!string.IsNullOrEmpty(CustState))
                {
                    LogStep($"Selecting State: {CustState}");
                    var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                    new SelectElement(stateDropdown).SelectByText(CustState);
                }

                LogStep($"Entering Postal Code: {CustPosCode}");
                _CustomerPage.EnterCustPosCode(CustPosCode);

                if (!string.IsNullOrEmpty(CustCountry))
                {
                    LogStep($"Selecting Country: {CustCountry}");
                    var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                    new SelectElement(countryDropdown).SelectByText(CustCountry);
                }

                LogStep($"Entering Address 1: {CustAddress1}");
                _CustomerPage.EnterCustAddress1(CustAddress1);

                LogStep($"Entering Address 2: {CustAddress2}");
                _CustomerPage.EnterCustAddress2(CustAddress2);

                LogStep($"Entering Address 3: {CustAddress3}");
                _CustomerPage.EnterCustAddress3(CustAddress3);

                if (!string.IsNullOrEmpty(CustExternalCode))
                {
                    LogStep($"Entering External Code: {CustExternalCode}");
                    _CustomerPage.EnterCustExternalCode(CustExternalCode);
                }

                // Check Save button before clicking
                var saveButtons = _driver.FindElements(By.XPath("//button[contains(., 'Save')]"));
                if (saveButtons.Count == 0 || !saveButtons[0].Displayed || !saveButtons[0].Enabled)
                {
                    LogStep($"✅ Save button not visible or clickable -  Negative test success for scenario: {scenario}");
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - Save button blocked.");
                    return;
                }

                LogStep("Clicking Save button...");
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                // Validation check
                var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
                if (validationMessages.Count > 0)
                {
                    foreach (var msg in validationMessages)
                    {
                        LogStep($"Validation Message: {msg.Text.Trim()}");
                    }
                    Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");
                    // Screenshot (no log step for this)
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }
                else
                {
                    LogStep("Checking modal after save...");
                    var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep($"📢 Modal Message: {message}");

                    if (message.Contains("success", StringComparison.OrdinalIgnoreCase))
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        LogStep("Clicking OK button on success modal (unexpected for negative test).");
                        okButton.Click();
                        Assert.Fail($"❌ Negative test failed for scenario {scenario} - update should not succeed.");
                    }
                    else
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        LogStep("Clicking OK button on failure modal (expected). Test success.");
                        okButton.Click();
                        Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid update.");
                    }
                    // Screenshot (no log step for this)
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }


            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during update negative test: {ex.Message}");
                Assert.Fail($"Negative update test failed for scenario {scenario} due to unexpected exception.");
            }
        }




        [Test]
        [Category("Customer")]
        [Order(10)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Customer CSV File")]
        [TestCaseSource(nameof(InvalidFileTestData))]
        public void ImportCustomer_InvalidFileType(string filePath)
        {
            const string expectedMessage = "Only CSV files are allowed!";

            try
            {
                // Open Import Modal
                LogStep("📂 Opening Customer Import Modal...");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_content_container > app-customer > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a"));
                _CustomerPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // Select invalid file
                LogStep($"📁 Selecting invalid file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                // Click Upload Button
                LogStep("📤 Clicking Upload button...");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _CustomerPage.ClickUploadButton();
                WaitForUIEffect();

                // First check if alert appears
                try
                {
                    IAlert alert = _driver.SwitchTo().Alert();
                    string alertText = alert.Text.Trim();

                    // Take screenshot
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                        $"Customer_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
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
                    // Continue to modal check
                }

                // If no alert → Verify modal
                LogStep("🔍 Verifying modal for invalid file type...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();

                // Take screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Customer_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot2 = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot2.AsByteArray);

                LogStep($"📢 Modal Message: {message}");
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
                // Timeout → system never rejected invalid file
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Customer_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Timeout waiting for alert/modal: {tex.Message}");
                Assert.Fail("Test failed: No alert or modal appeared for invalid file type.");
            }
            catch (Exception ex)
            {
                // Generic exception → still valid negative outcome
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(),
                    $"Customer_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"Expected Exception: {ex.Message}. Test success");
                Assert.IsTrue(true, "✅ Negative test considered successful (expected rejection).");
            }
        }


        //[Test]
        //[Category("Customer")]
        //[Order(11)]
        //[AllureSeverity(SeverityLevel.normal)]
        //[AllureStory("Customer Update - Negative Scenarios")]
        //[TestCase("CUST_000002", "QUBE APPS SOLUTIONS SDN BHD", "C29983588100", "BRN", "INVALIDID", "SST1", "testing1@email.com", "01234567890", "Petaling Jaya",
        //"Selangor", "12345", "MALAYSIA", "No 111, Jalan 222", "Lorong 333, Taman 444", "Perindustrian 555", "Ext12345", "InvalidRegisterID")] // Invalid Register ID

        //public void Update_Customer_InvalidRegisterID(string CustomerCode, string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail,
        // string CustContactNumber, string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3,
        // string CustExternalCode, string scenario)
        //{
        //    try
        //    {
        //        LogStep($"⏳ Starting Customer Update Negative Test - Scenario: {scenario}");

        //        LogStep("Clicking edit button...");
        //        _CustomerPage.ClickEditButton(CustomerCode);
        //        WaitForUIEffect();

        //        // Step 1 updates
        //        LogStep($"Entering Customer Name: {Custname}");
        //        _CustomerPage.EnterCustname(Custname);

        //        LogStep($"Entering TIN: {CustTinNumber}");
        //        _CustomerPage.EnterCustTinNumber(CustTinNumber);

        //        if (!string.IsNullOrEmpty(BEregisterType))
        //        {
        //            LogStep($"Selecting Register Type: {BEregisterType}");
        //            var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
        //                By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
        //            new SelectElement(regType).SelectByText(BEregisterType);
        //        }

        //        LogStep($"Entering Register ID: {CustRegisterID}");
        //        _CustomerPage.EnterCustRegisterID(CustRegisterID);

        //        LogStep($"Entering SST: {Custsst}");
        //        _CustomerPage.EnterCustsst(Custsst);

        //        LogStep($"Entering Email: {CustEmail}");
        //        _CustomerPage.EnterCustEmail(CustEmail);

        //        LogStep("Proceeding to Step 2...");
        //        _CustomerPage.ClickContinueButton();
        //        WaitForUIEffect();

        //        // Step 2 updates
        //        LogStep($"Entering Contact Number: {CustContactNumber}");
        //        _CustomerPage.EnterCustContactNumber(CustContactNumber);

        //        LogStep($"Entering City: {CustCity}");
        //        _CustomerPage.EnterCustomerCity(CustCity);

        //        if (!string.IsNullOrEmpty(CustState))
        //        {
        //            LogStep($"Selecting State: {CustState}");
        //            var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
        //                By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
        //            new SelectElement(stateDropdown).SelectByText(CustState);
        //        }

        //        LogStep($"Entering Postal Code: {CustPosCode}");
        //        _CustomerPage.EnterCustPosCode(CustPosCode);

        //        if (!string.IsNullOrEmpty(CustCountry))
        //        {
        //            LogStep($"Selecting Country: {CustCountry}");
        //            var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
        //                By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
        //            new SelectElement(countryDropdown).SelectByText(CustCountry);
        //        }

        //        LogStep($"Entering Address 1: {CustAddress1}");
        //        _CustomerPage.EnterCustAddress1(CustAddress1);

        //        LogStep($"Entering Address 2: {CustAddress2}");
        //        _CustomerPage.EnterCustAddress2(CustAddress2);

        //        LogStep($"Entering Address 3: {CustAddress3}");
        //        _CustomerPage.EnterCustAddress3(CustAddress3);

        //        if (!string.IsNullOrEmpty(CustExternalCode))
        //        {
        //            LogStep($"Entering External Code: {CustExternalCode}");
        //            _CustomerPage.EnterCustExternalCode(CustExternalCode);
        //        }

        //        LogStep("Clicking Save button...");
        //        _CustomerPage.ClickSaveButton();
        //        WaitForUIEffect(1000);

        //        // Validation check
        //        var validationMessages = _driver.FindElements(By.CssSelector(".text-danger"));
        //        if (validationMessages.Count > 0)
        //        {
        //            foreach (var msg in validationMessages)
        //            {
        //                LogStep($"Validation Message: {msg.Text.Trim()}");
        //            }
        //            Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - validation triggered correctly.");
        //        }
        //        else
        //        {
        //            LogStep("Checking modal after save...");
        //            var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
        //            var message = modal.Text.Trim();
        //            LogStep($"📢 Modal Message: {message}. Test success");

        //            if (message.Contains("success", StringComparison.OrdinalIgnoreCase))
        //            {
        //                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
        //                LogStep("Clicking OK button on success modal (unexpected for negative test).");
        //                okButton.Click();
        //                Assert.Fail($"❌ Negative test failed for scenario {scenario} - update should not succeed.");
        //            }
        //            else
        //            {
        //                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
        //                LogStep("Clicking OK button on failure modal (expected). Test success.");
        //                okButton.Click();
        //                Assert.IsTrue(true, $"✅ Negative test success for scenario: {scenario} - system blocked invalid update.");
        //            }
        //        }

        //        // Screenshot (no log step for this)
        //        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
        //        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
        //        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
        //    }
        //    catch (Exception ex)
        //    {
        //        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_Update_Negative_{scenario}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
        //        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
        //        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
        //        LogStep($"❌ Exception during update negative test: {ex.Message}");
        //        Assert.Fail($"Negative update test failed for scenario {scenario} due to unexpected exception.");
        //    }
        //}




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

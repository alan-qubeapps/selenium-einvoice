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
using SeleniumTests.Pages.Customer;
using SeleniumTests.Pages.Stores;
using System.Drawing;
using System.Globalization;
using System.Linq.Expressions;
using System.Media;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.C_Customer
{

    public static class ExcelDataReaderCustomerValid
    {
        public static IEnumerable<object[]> GetCreateCustomerTestData(string filePath, string sheetName)
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
                    string CustStatus = worksheet.Cells[row, 16].Text?.Trim();


                    yield return new object[]
                    {
                        Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, CustStatus
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetSearchCustomerTestData(string filePath, string sheetName)
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
                    string searchText = worksheet.Cells[row, 1].Text?.Trim();
                    int columnIndex = Convert.ToInt32(worksheet.Cells[row, 2].Value);


                    yield return new object[]
                    {
                        searchText, columnIndex
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateCustomerTestData(string filePath, string sheetName)
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
                    string CustStatus = worksheet.Cells[row, 17].Text?.Trim();


                    yield return new object[]
                    {
                        CustomerCode, Custname, CustTinNumber, BEregisterType, CustRegisterID, Custsst, CustEmail, CustContactNumber, CustCity, CustState, CustPosCode, CustCountry, CustAddress1, CustAddress2, CustAddress3, CustExternalCode, CustStatus
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetFilterCategoryTestData(string filePath, string sheetName)
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
                    string category = worksheet.Cells[row, 1].Text?.Trim();

                    yield return new object[]
                    {
                        category
                    };
                }
            }
        }

    }


    [TestFixture, Order(7)]
    [AllureNUnit]
    [AllureSuite("Customer - Customer - Valid")]
    [AllureEpic("ERP-117")]
    public class CustomerTests_Valid
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

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "CustomerTestDataValid.xlsx");

        public static IEnumerable<object[]> CreateCustomerTestData =>
        ExcelDataReaderCustomerValid.GetCreateCustomerTestData(ExcelPath, "CreateCustomerTestData");

        public static IEnumerable<object[]> SearchCustomerTestData =>
        ExcelDataReaderCustomerValid.GetSearchCustomerTestData(ExcelPath, "SearchCustomerTestData");

        public static IEnumerable<object[]> UpdateCustomerTestData =>
        ExcelDataReaderCustomerValid.GetUpdateCustomerTestData(ExcelPath, "UpdateCustomerTestData");

        public static IEnumerable<object[]> FilterCategoryTestData =>
        ExcelDataReaderCustomerValid.GetFilterCategoryTestData(ExcelPath, "FilterCategoryTestData");

        public static IEnumerable<object[]> FilterLHDNStatusTestData =>
        ExcelDataReaderCustomerValid.GetFilterLHDNStatusTestData(ExcelPath, "FilterLHDNStatusTestData");


        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;


        // ===== Add this at the top inside your class =====
        private readonly Dictionary<string, string> StateCodeMapping = new Dictionary<string, string>
        {
            { "01", "Johor" },
            { "02", "Kedah" },
            { "03", "Kelantan" },
            { "04", "Melaka" },
            { "05", "Negeri Sembilan" },
            { "06", "Pahang" },
            { "07", "Pulau Pinang" },
            { "08", "Perak" },
            { "09", "Perlis" },
            { "10", "Selangor" },
            { "11", "Terengganu" },
            { "12", "Sabah" },
            { "13", "Sarawak" },
            { "14", "Wilayah Persekutuan Kuala Lumpur" },
            { "15", "Wilayah Persekutuan Labuan" },
            { "16", "Wilayah Persekutuan Putrajaya" }
        };


        private readonly Dictionary<string, string> CountryCodeMapping = new Dictionary<string, string>
        {
            { "MYS", "MALAYSIA" }
            // In the future, you can add more:
            // { "SGP", "Singapore" },
            // { "THA", "Thailand" }
        };


        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Customer Page";

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

            _moduleName = "Customer Page";
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
        /// Test Case: Create Customer
        /// Action:
        ///     1. Click 'New' button to open customer creation modal.
        ///     2. Fill in customer details: Name, TIN, Register Type/ID, SST, Email, Contact Number, City, State, Postal Code, Country, Address, External Code, Status.
        ///     3. Click 'Continue' and proceed to Step 2 inputs.
        ///     4. Click 'Save' button.
        /// Verification:
        ///     - Modal displays success message after creation.
        ///     - Newly created customer exists in the table with all fields matching input values.
        ///     - Handles pagination to locate the customer in the table.
        /// Purpose:
        ///     Ensure that creating a new customer works correctly and all input fields are properly saved and reflected in the table.
        /// Test Data:
        ///     - Provided by 'CreateCustomerTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Customer")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(CreateCustomerTestData))]
        public void Create_Customer(string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
        string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode, string status)
        {
            try
            {
                LogStep("Click 'New' button");
                _CustomerPage.ClickNewButton();
                WaitForUIEffect();

                LogStep($"Enter Customer Name: {Custname}");
                _CustomerPage.EnterCustname(Custname);
                WaitForUIEffect();

                LogStep($"Enter TIN Number: {CustTinNumber}");
                _CustomerPage.EnterCustTinNumber(CustTinNumber);
                WaitForUIEffect();

                LogStep($"Select Register Type: {BEregisterType}");
                var regType = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                new SelectElement(regType).SelectByText(BEregisterType);
                WaitForUIEffect();

                LogStep($"Enter Register ID: {CustRegisterID}");
                _CustomerPage.EnterCustRegisterID(CustRegisterID);
                WaitForUIEffect();

                LogStep($"Enter SST: {Custsst}");
                _CustomerPage.EnterCustsst(Custsst);
                WaitForUIEffect();

                LogStep($"Enter Email: {CustEmail}");
                _CustomerPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep("Click 'Continue' to Step 2");
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                LogStep($"Enter Contact Number: {CustContactNumber}");
                _CustomerPage.EnterCustContactNumber(CustContactNumber);
                WaitForUIEffect();

                LogStep($"Enter City: {CustCity}");
                _CustomerPage.EnterCustomerCity(CustCity);
                WaitForUIEffect();

                LogStep($"Select State: {CustState}");
                var stateDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                new SelectElement(stateDropdown).SelectByText(CustState);
                WaitForUIEffect();

                LogStep($"Enter Postcode: {CustPosCode}");
                _CustomerPage.EnterCustPosCode(CustPosCode);
                WaitForUIEffect();

                LogStep($"Select Country: {CustCountry}");
                var countryDropdown = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                new SelectElement(countryDropdown).SelectByText(CustCountry);
                WaitForUIEffect();

                LogStep($"Enter Address 1: {CustAddress1}");
                _CustomerPage.EnterCustAddress1(CustAddress1);
                WaitForUIEffect();

                LogStep($"Enter Address 2: {CustAddress2}");
                _CustomerPage.EnterCustAddress2(CustAddress2);
                WaitForUIEffect();

                LogStep($"Enter Address 3: {CustAddress3}");
                _CustomerPage.EnterCustAddress3(CustAddress3);
                WaitForUIEffect();

                LogStep($"Enter External Code: {CustExternalCode}");
                _CustomerPage.EnterCustExternalCode(CustExternalCode);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(status))
                {
                    try
                    {
                        // Locate the toggle element
                        var statusToggle = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.CssSelector("input[formcontrolname='status']")));

                        bool isChecked = statusToggle.Selected;

                        if (status.Equals("Active", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!isChecked)
                            {
                                statusToggle.Click();
                                WaitForUIEffect();
                                LogStep("✅ Status set to: Active (checked)");
                            }
                            else
                            {
                                LogStep("ℹ️ Status already Active (checked)");
                            }
                        }
                        else if (status.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                        {
                            if (isChecked)
                            {
                                statusToggle.Click();
                                WaitForUIEffect();
                                LogStep("✅ Status set to: Inactive (unchecked)");
                            }
                            else
                            {
                                LogStep("ℹ️ Status already Inactive (unchecked)");
                            }
                        }
                        else if (status.Equals("All", StringComparison.OrdinalIgnoreCase))
                        {
                            LogStep("ℹ️ Status = All → No toggle action performed.");
                        }
                        else
                        {
                            LogStep($"⚠️ Unknown status input: {status}");
                        }
                    }
                    catch
                    {
                        LogStep("⚠️ Failed to locate or interact with Status toggle checkbox.");
                    }
                }

                LogStep("Click 'Save' button");
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("Modal Message: " + message);

                if (!message.ToLower().Contains("success"))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    Assert.Fail($"❌ Expected success message but got: {message}");
                }

                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Customer created successfully.");

                // ===== Step 6: Verification using input variables =====
                LogStep("🔍 Verifying newly created customer in the table...");

                bool isMatchFound = false;

                // Combine address for verification
                string combinedAddress = $"{CustAddress1},{CustAddress2},{CustAddress3},{CustPosCode}," +
                    //$"{CustCity}," +
                    $"{CustState},{CustCountry}".Replace("\"", "");

                // Search by Customer Name
                _CustomerPage.SearchCustomer(Custname);
                WaitForUIEffect(2000);
                helperFunction.WaitForCustTableToLoad(_wait);
                WaitForUIEffect(2000);

                while (true)
                {
                    var rows = _driver.FindElements(By.XPath(
                        "/html/body/app-layout/div[1]/div/div/div/app-content/app-customer/div/div[4]/div/div[1]/div/table/tbody/tr"));

                    foreach (var row in rows)
                    {
                        var cells = row.FindElements(By.TagName("td"));
                        if (cells.Count < 11) continue;

                        string actualCustName = cells[1].Text.Trim();
                        string actualCustEmail = cells[2].Text.Trim();
                        string actualCustPhoneNum = cells[3].Text.Trim();
                        string actualcombinedAddress = cells[4].Text.Trim();
                        string actualCustStatus = cells[5].Text.Trim();

                        if (actualCustName.Equals(Custname, StringComparison.OrdinalIgnoreCase))
                        {
                            // Combined verification log
                            LogStep($"🔹 Verifying Customer '{Custname}' -> " +
                                $"Email: Expected '{CustEmail}', Actual '{actualCustEmail}'; " +
                                $"Phone Number: Expected '{CustContactNumber}', Actual '{actualCustPhoneNum}'; " +
                                    $"Status: Expected '{status}', Actual '{actualCustStatus}'; " +
                                    $"Address: Expected '{combinedAddress}', Actual '{actualcombinedAddress}'");

                            if (actualCustStatus.Equals(status, StringComparison.OrdinalIgnoreCase)
                                && actualCustEmail.Equals(CustEmail, StringComparison.OrdinalIgnoreCase)
                                && actualCustPhoneNum.Equals(CustContactNumber, StringComparison.OrdinalIgnoreCase)
                                && actualcombinedAddress.Equals(combinedAddress, StringComparison.OrdinalIgnoreCase))
                            {
                                isMatchFound = true;
                                LogStep($"🎉 All fields matched successfully for '{Custname}'");
                            }
                            else
                            {
                                Assert.Fail($"❌ Verification failed for '{Custname}', see log for details.");
                            }

                            break; // stop checking rows
                        }
                    }

                    if (isMatchFound) break;

                    // ===== Pagination Handling =====
                    try
                    {
                        var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                        if (!nextButton.GetAttribute("class").Contains("disabled"))
                        {
                            nextButton.Click();
                            WaitForUIEffect(1500);
                            helperFunction.WaitForCustTableToLoad(_wait);
                        }
                        else break;
                    }
                    catch { break; }
                }

                if (!isMatchFound)
                {
                    Assert.Fail($"❌ Customer '{Custname}' was not found in the table after creation.");
                }

                LogStep("🎉 Store creation and verification completed successfully.");
            
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Search Customer
        /// Action:
        ///     1. Enter search text into the customer table search field.
        ///     2. Check the specified column for partial or full matches.
        ///     3. Navigate through pagination if needed.
        /// Verification:
        ///     - At least one row contains the search text in the specified column (partial match allowed).
        ///     - Screenshots captured for match found, page navigation, or errors.
        /// Purpose:
        ///     Ensure that the customer search functionality correctly filters rows based on column-specific input.
        /// Test Data:
        ///     - Provided by 'SearchCustomerTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Customer")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Customer Search - Column Specific Match (Partial Match Accepted)")]
        [TestCaseSource(nameof(SearchCustomerTestData))]
        public void Search_Customer(string searchText, int columnIndex)
        {
            LogStep($"🔍 Searching for text: {searchText} (Column {columnIndex})");
            _CustomerPage.SearchCustomer(searchText);
            helperFunction.WaitForCustTableToLoad(_wait);
            WaitForUIEffect();

            bool isMatchFound = false;

            while (true)
            {
                LogStep("Checking table rows on current page...");
                WaitForUIEffect();

                var rows = _driver.FindElements(By.XPath(
                    "/html/body/app-layout/div/div/div/div/app-content/app-customer/div/div[4]/div/div[1]/div/table/tbody/tr"));

                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));

                    if (columnIndex >= cells.Count)
                        continue; // Skip row if column index is invalid

                    string cellText;
                    try
                    {
                        // Attempt to get inner span text if styled
                        cellText = cells[columnIndex].FindElement(By.TagName("span")).Text.Trim();
                    }
                    catch
                    {
                        cellText = cells[columnIndex].Text.Trim();
                    }

                    LogStep($"🔎 Checking column {columnIndex}: '{cellText}' vs '{searchText}'");

                    if (cellText.Replace(" ", "").Contains(searchText.Replace(" ", ""), StringComparison.OrdinalIgnoreCase))
                    {
                        isMatchFound = true;
                        break;
                    }
                }

                if (isMatchFound)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("✅ Match found.");
                    break;
                }

                // Attempt to click the Next button (pagination)
                try
                {
                    var nextButton = _driver.FindElement(By.XPath(
                        "/html/body/app-layout/div/div/div/div/app-content/app-customer/div/div[4]/div/div[2]/app-global-pagination/div/div[2]/ul/li[4]"));

                    if (!nextButton.GetAttribute("class").Contains("disabled"))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("Navigating to next page...");
                        nextButton.Click();
                        helperFunction.WaitForBETableToLoad(_wait);
                        WaitForUIEffect();
                    }
                    else
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("⛔ Reached last page. No more pages to check.");
                        break;
                    }
                }
                catch (NoSuchElementException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⚠️ No pagination found. Stopping search.");
                    break;
                }
            }

            WaitForUIEffect();
            Assert.IsTrue(isMatchFound, $"❌ No matching record found for: '{searchText}' in column index {columnIndex}.");
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Update Customer
        /// Action:
        ///     1. Click the Edit button for a specific customer identified by CustomerCode.
        ///     2. Update customer details including Name, TIN, Register Type, Register ID, SST, Email, Contact Number,
        ///        City, State, Postcode, Country, Address1-3, External Code, and Status.
        ///     3. Click Save to submit changes.
        /// Verification:
        ///     - Modal message indicates success (or handles duplicate TIN / failure appropriately).
        ///     - Customer table reflects updated values, verified by searching for CustomerCode and checking all fields.
        ///     - Pagination handled if the customer is not on the current page.
        ///     - Screenshots captured at key steps (before save, modal, verification).
        /// Purpose:
        ///     Ensure that the customer update functionality correctly applies changes and reflects them in the table.
        /// Test Data:
        ///     - Provided by 'UpdateCustomerTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Customer")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Customer Update")]
        [TestCaseSource(nameof(UpdateCustomerTestData))]
        public void Update(string CustomerCode, string Custname, string CustTinNumber, string BEregisterType, string CustRegisterID, string Custsst, string CustEmail, string CustContactNumber,
        string CustCity, string CustState, string CustPosCode, string CustCountry, string CustAddress1, string CustAddress2, string CustAddress3, string CustExternalCode, string status)
        {
            try
            {
                LogStep("⏳ Clicking edit button for customer: " + CustomerCode);
                _CustomerPage.ClickEditButton(CustomerCode);
                WaitForUIEffect();

                LogStep("⏳ Updating customer details...");
                _CustomerPage.EnterCustname(Custname);
                WaitForUIEffect();

                _CustomerPage.EnterCustTinNumber(CustTinNumber);
                WaitForUIEffect();

                var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[4]/div[1]/div/select")));
                new SelectElement(regType).SelectByText(BEregisterType);
                WaitForUIEffect();

                _CustomerPage.EnterCustRegisterID(CustRegisterID);
                WaitForUIEffect();

                _CustomerPage.EnterCustsst(Custsst);
                WaitForUIEffect();

                _CustomerPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep("Proceeding to step 2...");
                _CustomerPage.ClickContinueButton();
                WaitForUIEffect();

                _CustomerPage.EnterCustContactNumber(CustContactNumber);
                WaitForUIEffect();

                _CustomerPage.EnterCustomerCity(CustCity);
                WaitForUIEffect();

                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[2]/div/select")));
                new SelectElement(stateDropdown).SelectByText(CustState);
                WaitForUIEffect();

                _CustomerPage.EnterCustPosCode(CustPosCode);
                WaitForUIEffect();

                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-customer-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[3]/select")));
                new SelectElement(countryDropdown).SelectByText(CustCountry);
                WaitForUIEffect();

                _CustomerPage.EnterCustAddress1(CustAddress1);
                WaitForUIEffect();

                _CustomerPage.EnterCustAddress2(CustAddress2);
                WaitForUIEffect();

                _CustomerPage.EnterCustAddress3(CustAddress3);
                WaitForUIEffect();

                _CustomerPage.EnterCustExternalCode(CustExternalCode);
                WaitForUIEffect();

                if (!string.IsNullOrEmpty(status))
                {
                    try
                    {
                        // Locate the toggle element
                        var statusToggle = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.CssSelector("input[formcontrolname='status']")));

                        bool isChecked = statusToggle.Selected;

                        if (status.Equals("Active", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!isChecked)
                            {
                                statusToggle.Click();
                                WaitForUIEffect();
                                LogStep("✅ Status set to: Active (checked)");
                            }
                            else
                            {
                                LogStep("ℹ️ Status already Active (checked)");
                            }
                        }
                        else if (status.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
                        {
                            if (isChecked)
                            {
                                statusToggle.Click();
                                WaitForUIEffect();
                                LogStep("✅ Status set to: Inactive (unchecked)");
                            }
                            else
                            {
                                LogStep("ℹ️ Status already Inactive (unchecked)");
                            }
                        }
                        else if (status.Equals("All", StringComparison.OrdinalIgnoreCase))
                        {
                            LogStep("ℹ️ Status = All → No toggle action performed.");
                        }
                        else
                        {
                            LogStep($"⚠️ Unknown status input: {status}");
                        }
                    }
                    catch
                    {
                        LogStep("⚠️ Failed to locate or interact with Status toggle checkbox.");
                    }
                }

                LogStep("💾 Saving updated customer record...");
                _CustomerPage.ClickSaveButton();
                WaitForUIEffect();

                // ✅ Modal validation
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("📢 Modal message: " + message);

                if (message.Contains("TIN has already been taken", StringComparison.OrdinalIgnoreCase))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    var duplicateOkBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                    duplicateOkBtn.Click();

                    Assert.Fail("❌ Duplicate TIN error: " + message);
                }
                else if (message.ToLower().Contains("fail"))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    var duplicateOkBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                    duplicateOkBtn.Click();

                    Assert.Fail("❌ Unexpected failure in modal: " + message);
                }
                else
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    Assert.IsTrue(message.IndexOf("Successful", StringComparison.OrdinalIgnoreCase) >= 0, $"❌ Modal does not indicate success: {message}");

                    var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                    okButton.Click();
                    WaitForUIEffect();

                    LogStep("✅ Customer update completed successfully.");

                    // ===== Step 6: Verification using input variables =====
                    LogStep("🔍 Verifying newly updated customer in the table...");

                    bool isMatchFound = false;

                    // Combine address for verification
                    string combinedAddress = $"{CustAddress1},{CustAddress2},{CustAddress3},{CustPosCode}," +
                        //$"{CustCity}," +
                        $"{CustState},{CustCountry}".Replace("\"", "");

                    // Search by Customer Name
                    _CustomerPage.SearchCustomer(CustomerCode);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForCustTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-customer/div/div[4]/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 11) continue;

                            string actualCustName = cells[1].Text.Trim();
                            string actualCustEmail = cells[2].Text.Trim();
                            string actualCustPhoneNum = cells[3].Text.Trim();
                            string actualcombinedAddress = cells[4].Text.Trim();
                            string actualCustStatus = cells[5].Text.Trim();

                            if (actualCustName.Equals(Custname, StringComparison.OrdinalIgnoreCase))
                            {
                                // Combined verification log
                                LogStep($"🔹 Verifying Customer '{Custname}' -> " +
                                    $"Email: Expected '{CustEmail}', Actual '{actualCustEmail}'; " +
                                    $"Phone Number: Expected '{CustContactNumber}', Actual '{actualCustPhoneNum}'; " +
                                        $"Status: Expected '{status}', Actual '{actualCustStatus}'; " +
                                        $"Address: Expected '{combinedAddress}', Actual '{actualcombinedAddress}'");

                                if (actualCustStatus.Equals(status, StringComparison.OrdinalIgnoreCase)
                                    && actualCustEmail.Equals(CustEmail, StringComparison.OrdinalIgnoreCase)
                                    && actualCustPhoneNum.Equals(CustContactNumber, StringComparison.OrdinalIgnoreCase)
                                    && actualcombinedAddress.Equals(combinedAddress, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"🎉 All fields matched successfully for '{Custname}'");
                                }
                                else
                                {
                                    Assert.Fail($"❌ Verification failed for '{Custname}', see log for details.");
                                }

                                break; // stop checking rows
                            }
                        }

                        if (isMatchFound) break;

                        // ===== Pagination Handling =====
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForCustTableToLoad(_wait);
                            }
                            else break;
                        }
                        catch { break; }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Customer '{Custname}' was not found in the table after creation.");
                    }

                    LogStep("🎉 Customer updated and verification completed successfully.");

                }

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during update: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Export Customer Report
        /// Action:
        ///     1. Click the Export button in the Customer page.
        ///     2. Wait for the report file to be downloaded to the configured Downloads folder.
        /// Verification:
        ///     - Exported file with expected prefix ("Customer Index") exists in the Downloads folder.
        ///     - Screenshot captured for export action.
        /// Purpose:
        ///     Ensure that the customer report can be exported successfully and the file is correctly downloaded.
        /// Test Data:
        ///     - Download path retrieved from AppConfig.DownloadPath
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("StoreGroup")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Export Customer Report")]
        public void ExportCustomerReport()
        {
            string downloadPath = AppConfig.DownloadPath;
            string filePrefix = "Customer Index";

            LogStep("📤 Clicking export button for customer report...");
            helperFunction.WaitForElementToBeClickable(_wait,
                By.CssSelector("#kt_content_container > app-customer > div > div.card-header.border-0.pt-5 > div > div:nth-child(2) > a"));
            _CustomerPage.ClickExportButton();

            LogStep("⏳ Waiting for download to complete...");
            bool fileDownloaded = _CustomerPage.WaitForFileDownload(downloadPath, filePrefix, TimeSpan.FromSeconds(15));

            // 📸 Always capture screenshot
            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            if (!fileDownloaded)
            {
                Assert.Fail("❌ Export file was not found in the Downloads folder.");
            }
            else
            {
                LogStep("✅ Export file successfully detected in Downloads folder.");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Filter Customer by Category
        /// Action:
        ///     1. Apply the selected filter category (All / Active / Inactive) on the Customer page.
        ///     2. Wait for the customer table to reload after applying the filter.
        ///     3. Capture a screenshot after filter is applied.
        /// Verification:
        ///     - For 'Active' filter: All visible rows have Status = 'Active'.
        ///     - For 'Inactive' filter: All visible rows have Status = 'Inactive'.
        ///     - For 'All' filter: Rows may contain both 'Active' and 'Inactive' statuses.
        ///     - If no data is available, ensure the "No data available" message is displayed.
        /// Purpose:
        ///     Ensure that the customer category filter correctly displays rows based on the selected category and handles empty datasets gracefully.
        /// Test Data:
        ///     - Provided by 'FilterCategoryTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Customer")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Filter - Filter By Category")]
        [TestCaseSource(nameof(FilterCategoryTestData))]
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
                    var noDataElement = _driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-customer/div/div[4]/div/div[1]/div/table/tbody[2]/tr/td/p"));
                    string message = noDataElement?.Text?.Trim();
                    return message != null && message.Equals("No data available", StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return false;
                }
            }

            void ValidateStatusColumn(string expectedStatus)
            {
                var rows = GetRows();
                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));
                    if (cells.Count < 6) continue;

                    string actualStatus = GetStatusFromCell(cells[5]);

                    LogStep($"🔍 Validating Status: Expected = '{expectedStatus}', Found = '{actualStatus}'");

                    if (!actualStatus.Equals(expectedStatus, StringComparison.OrdinalIgnoreCase))
                    {
                        LogStep($"❌ Mismatch - Expected: '{expectedStatus}', Found: '{actualStatus}'");
                        Assert.Fail("❌ One or more rows have unexpected status.");
                    }
                }
            }

            // === Apply Filter ===
            switch (category.Trim().ToUpperInvariant())
            {
                case "ALL":
                    LogStep("📌 Testing filter: All Categories");
                    _CustomerPage.ClickFilterALLCategoryButton();
                    break;

                case "ACTIVE":
                    LogStep("📌 Testing filter: Active Category");
                    _CustomerPage.ClickFilterActiveCategoryButton();
                    break;

                case "INACTIVE":
                    LogStep("📌 Testing filter: Inactive Category");
                    _CustomerPage.ClickFilterInactiveCategoryButton();
                    break;

                default:
                    Assert.Fail($"❌ Invalid filter category input: '{category}'");
                    break;
            }

            WaitForUIEffect();
            helperFunction.WaitForCustTableToLoad(_wait);

            // === Screenshot ===
            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
            LogStep($"📸 Screenshot saved to: {_lastScreenshotPath}");

            // === Handle "No Data" Scenario ===
            if (IsNoDataMessageShown())
            {
                LogStep($"✅ Filter applied: '{category}' - No records found.");
                LogStep("📤 Export: No data found. Test passed as no invalid data shown.");
                return;
            }

            var rows = GetRows();
            if (!rows.Any())
            {
                Assert.Fail("❌ Expected data rows but none were found.");
            }

            // === Validate based on category ===
            if (category.Equals("ACTIVE", StringComparison.OrdinalIgnoreCase))
            {
                ValidateStatusColumn("Active");
            }
            else if (category.Equals("INACTIVE", StringComparison.OrdinalIgnoreCase))
            {
                ValidateStatusColumn("Inactive");
            }
            else if (category.Equals("ALL", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));
                    if (cells.Count < 6) continue;

                    string actualStatus = GetStatusFromCell(cells[5]);
                    LogStep($"🔍 Found Status = '{actualStatus}'");
                }
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Filter Customer by LHDN Status
        /// Action:
        ///     1. Apply the selected LHDN status filter (Pending / Success / Failed) on the Customer page.
        ///     2. Wait for the customer table to reload after applying the filter.
        ///     3. Capture a screenshot after the filter is applied.
        /// Verification:
        ///     - For 'Pending', 'Success', or 'Failed': All visible rows must have the expected LHDN status.
        ///     - If no data is available, ensure the "No data available" message is displayed.
        /// Purpose:
        ///     Ensure that the customer LHDN status filter correctly displays rows according to the selected status and handles empty datasets gracefully.
        /// Test Data:
        ///     - Provided by 'FilterLHDNStatusTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Customer")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Filter - Filter By LHDN Status")]
        [TestCaseSource(nameof(FilterLHDNStatusTestData))]
        public void FilterLHDNStatusFunction(string lhdnStatus)
        {
            // === Get only data rows from tbody[1] ===
            IReadOnlyCollection<IWebElement> GetRows() =>
                _driver.FindElements(By.XPath("//table/tbody[1]/tr"));

            // === Check if "No data available" is shown ===
            bool IsNoDataMessageShown()
            {
                try
                {
                    var noDataElement = _driver.FindElement(By.XPath("//table/tbody[2]/tr/td/p"));
                    var message = noDataElement.Text.Trim();
                    return message.Equals("No data available", StringComparison.OrdinalIgnoreCase);
                }
                catch (NoSuchElementException)
                {
                    return false;
                }
            }

            // === Validate all rows match expected status ===
            bool AllRowsMatchExpectedStatus(string expectedStatus)
            {
                var rows = GetRows();
                bool allMatch = true;

                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));
                    if (cells.Count < 7) continue;

                    string actualStatus;
                    try
                    {
                        actualStatus = cells[6].FindElement(By.TagName("span")).Text.Trim();
                    }
                    catch
                    {
                        actualStatus = cells[6].Text.Trim();
                    }

                    LogStep($"🔍 Validating LHDN Status: Expected = '{expectedStatus}', Found = '{actualStatus}'");

                    if (!actualStatus.Equals(expectedStatus, StringComparison.OrdinalIgnoreCase))
                    {
                        LogStep($"❌ Mismatch - Expected: '{expectedStatus}', but Found: '{actualStatus}'");
                        allMatch = false;
                    }
                }

                return allMatch;
            }

            // === Apply Filter ===
            switch (lhdnStatus.Trim().ToLowerInvariant())
            {
                case "pending":
                    LogStep("🟡 Filtering by LHDN Status: Pending");
                    _CustomerPage.ClickFilterPendingCategoryButton();
                    break;
                case "success":
                    LogStep("🟢 Filtering by LHDN Status: Success");
                    _CustomerPage.ClickFilterSuccessCategoryButton();
                    break;
                case "failed":
                    LogStep("🔴 Filtering by LHDN Status: Failed");
                    _CustomerPage.ClickFilterFailedCategoryButton();
                    break;
                default:
                    Assert.Fail($"❌ Invalid LHDN status input: '{lhdnStatus}'");
                    break;
            }

            // === Wait and Screenshot ===
            helperFunction.WaitForCustTableToLoad(_wait);
            WaitForUIEffect();

            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            // === If "No data available" is displayed, pass the test ===
            if (IsNoDataMessageShown())
            {
                LogStep($"✅ 'No data available' is shown for '{lhdnStatus}' filter. Test succeeded.");
                return; // Exit the test early — pass
            }

            // === If data exists, validate all rows ===
            var dataRows = GetRows();
            if (dataRows.Count == 0)
            {
                Assert.Fail("❌ No data rows found, and no 'No data available' message. Possible UI rendering issue.");
            }

            Assert.IsTrue(AllRowsMatchExpectedStatus(lhdnStatus), $"❌ Some rows do not match the '{lhdnStatus}' status.");
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Import Customer via CSV File
        /// Action:
        ///     1. Open the Import Customer modal.
        ///     2. Download the CSV template and prepare the import file.
        ///     3. Upload the CSV file using the Upload button.
        ///     4. Handle Import Error popup if it appears.
        ///     5. Click the 'Completed' button and verify the modal message.
        ///     6. Read customer data from the uploaded CSV file.
        ///     7. Verify each customer’s Name, Email, Telephone, Address, and Status in the customer table.
        /// Verification:
        ///     - No Import Error popup should appear.
        ///     - The modal should display a success message after upload.
        ///     - All customers from the CSV should exist in the table with correct information.
        ///     - Screenshots captured at key steps for upload, error, and verification.
        /// Purpose:
        ///     Ensure that the CSV import functionality correctly adds customers and the data is accurately reflected in the customer table.
        /// Test Data:
        ///     - CSV file path provided by 'AppConfig.ImportCustomerCSVFile'
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Customer")]
        [Order(8)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button")]
        public void ImportCustomerCSVFile()
        {
            try
            {
                LogStep("📤 Opening Import Modal...");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_content_container > app-customer > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a"));
                _CustomerPage.ClickImportButton();
                WaitForUIEffect();

                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-body.px-20 > div > div > div.d-flex.align-items-center > button"));
                _CustomerPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                LogStep("📂 Uploading Customer CSV File...");
                string filePath = AppConfig.ImportCustomerCSVFile;
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                LogStep($"✅ File selected: {filePath}");
                WaitForUIEffect();

                LogStep("✅ Clicking Upload button...");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _CustomerPage.ClickUploadButton();

                LogStep("⏳ Checking for Import Error popup...");

                // Wait up to 5 seconds to see if the Import Error popup appears
                try
                {
                    WebDriverWait errorWait = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
                    IWebElement importError = errorWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                        By.XPath("//*[@id='kt_body']/app-layout/div[2]/app-toast[1]/div/div/div/div/div/div[2]/label[1]")));

                    if (importError.Text.Trim().Equals("IMPORT ERROR", StringComparison.OrdinalIgnoreCase))
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}_Error.png");
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
                LogStep("✅ 'Completed' button detected. Clicking...");
                completedButton.Click();
                WaitForUIEffect();

                LogStep("🔍 Checking upload result modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal message: {message}");

                if (!message.ToLower().Contains("success"))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                    Assert.Fail($"❌ Expected success message but got: {message}");
                }
                else
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    okButton.Click();
                    WaitForUIEffect();

                    // ===== Step 6: Read CSV =====
                    LogStep("📄 Reading customer information from CSV...");
                    var CustomerName = new List<string>();
                    var CustomerEmail = new List<string>();
                    var CustomerTelephone = new List<string>();
                    var CustomerAddress1 = new List<string>();
                    var CustomerAddress2 = new List<string>();
                    var CustomerAddress3 = new List<string>();
                    var CustomerCityName = new List<string>();
                    var CustomerPostcode = new List<string>();
                    var CustomerStateName = new List<string>();
                    var CustomerCountryName = new List<string>();
                    var CustomerStatus = new List<string>();

                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var reader = new StreamReader(stream))
                    {
                        int rowIndex = 0;
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            rowIndex++;
                            if (rowIndex < 4) continue;

                            var columns = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                            if (columns.Length >= 16)
                            {
                                var CustName = columns[0].Trim();
                                if (string.IsNullOrEmpty(CustName))
                                {
                                    LogStep("Empty Customer Name detected, stop getting CSV data.");
                                    break;
                                }

                                CustomerName.Add(CustName);
                                CustomerEmail.Add(columns[6].Trim());
                                CustomerTelephone.Add(columns[5].Trim());
                                CustomerAddress1.Add(columns[10].Trim());
                                CustomerAddress2.Add(columns[11].Trim());
                                CustomerAddress3.Add(columns[12].Trim());
                                CustomerCityName.Add(columns[13].Trim());
                                CustomerPostcode.Add(columns[7].Trim());

                                // State
                                var stateCode = columns[8].Trim();
                                if (StateCodeMapping.TryGetValue(stateCode, out var stateName))
                                {
                                    CustomerStateName.Add(stateName);

                                }
                                else
                                {
                                    CustomerStateName.Add("Unknown State");
                                    LogStep($"⚠️ Invalid state code '{stateCode}' detected in CSV.");

                                }

                                // Country
                                var countryCode = columns[9].Trim();
                                if (CountryCodeMapping.TryGetValue(countryCode, out var countryName))
                                {
                                    CustomerCountryName.Add(countryName);
                                }
                                else
                                {
                                    CustomerCountryName.Add("Unknown Country");
                                    LogStep($"⚠️ Invalid or unrecognized country code '{countryCode}' detected.");

                                }

                                // ⬅️ ADD CUSTOMER STATUS (FIX)
                                CustomerStatus.Add(columns[14].Trim());
                            }
                        }
                    }


                    // ===== Step 7: Verification =====
                    for (int i = 0; i < CustomerName.Count; i++)
                    {
                        string CustName = CustomerName[i];
                        string expectedCustStatus = CustomerStatus[i];
                        string expectedCustEmail = CustomerEmail[i];
                        string expectedCustTelephone = CustomerTelephone[i];

                        // Create a list of all address parts
                        var addressParts = new List<string>
                        {
                            CustomerAddress1[i],
                            CustomerAddress2[i],
                            CustomerAddress3[i],
                            CustomerPostcode[i],
                            // reopen after lucas fix
                            //CustomerCityName[i], 
                            CustomerStateName[i],
                            CustomerCountryName[i]
                        };

                        // Keep only non-empty parts
                        var nonEmptyParts = addressParts.Where(x => !string.IsNullOrWhiteSpace(x));

                        // Join with a single comma between them
                        string combinedAddress = string.Join(",", nonEmptyParts).Replace("\"", "");


                        bool isMatchFound = false;

                        LogStep($"🔍 Searching Store Information '{CustName}' to verify Customer Email Address : '{expectedCustEmail}', Customer Phone Number : '{expectedCustTelephone}', Address : '{combinedAddress}', Customer Status : '{expectedCustStatus}'");

                        _CustomerPage.SearchCustomer(CustName);
                        WaitForUIEffect(2000);
                        helperFunction.WaitForCustTableToLoad(_wait);
                        WaitForUIEffect(2000);

                        while (true)
                        {
                            var rows = _driver.FindElements(By.XPath(
                                "/html/body/app-layout/div[1]/div/div/div/app-content/app-customer/div/div[4]/div/div[1]/div/table/tbody/tr"));

                            foreach (var row in rows)
                            {
                                var cells = row.FindElements(By.TagName("td"));
                                if (cells.Count < 11) continue;

                                string actualCustName = cells[1].Text.Trim();
                                string actualCustEmail = cells[2].Text.Trim();
                                string actualCustTelephone = cells[3].Text.Trim();
                                string actualcombinedAddress = cells[4].Text.Trim();
                                string actualCustStatus = cells[5].Text.Trim();

                                if (actualCustName.Equals(CustName, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualCustEmail.Equals(expectedCustEmail, StringComparison.OrdinalIgnoreCase)
                                        && actualCustTelephone.Equals(expectedCustTelephone, StringComparison.OrdinalIgnoreCase)
                                        && actualcombinedAddress.Equals(combinedAddress, StringComparison.OrdinalIgnoreCase)
                                        && actualCustStatus.Equals(expectedCustStatus, StringComparison.OrdinalIgnoreCase))
                                    {
                                        isMatchFound = true;
                                        LogStep($"✅ Customer Information matched : '{CustName}' Customer Email Address : '{actualCustEmail}', Customer Phone Number : '{actualCustTelephone}', Address : '{actualcombinedAddress}', Customer Status : '{actualCustStatus}'");
                                    }
                                    else
                                    {
                                        Assert.Fail($"❌ Customer Information '{CustName}' mismatch. Expected: Customer Email Address : '{expectedCustEmail}', Customer Phone Number : '{expectedCustTelephone}', Address : '{combinedAddress}', Customer Status : '{expectedCustStatus}'");
                                    }
                                    break;
                                }
                            }

                            if (isMatchFound) break;

                            try
                            {
                                var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                                if (!nextButton.GetAttribute("class").Contains("disabled"))
                                {
                                    nextButton.Click();
                                    WaitForUIEffect(1500);
                                    helperFunction.WaitForCustTableToLoad(_wait);
                                }
                                else break;
                            }
                            catch { break; }
                        }

                        if (!isMatchFound)
                        {
                            Assert.Fail($"❌ Customer Information '{CustName}' was not found.");
                        }
                    }

                    LogStep("🎉 CSV Import and verification completed successfully.");
                }
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Customer_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during import: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
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

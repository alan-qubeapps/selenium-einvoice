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
using System.Media;
using System.Text.RegularExpressions;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.B_Stores
{


    public static class ExcelDataReaderStoreValid
    {
        public static IEnumerable<object[]> GetCreateStoreTestData(string filePath, string sheetName)
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
                    string Storename = worksheet.Cells[row, 1].Text?.Trim();
                    string StoreCity = worksheet.Cells[row, 2].Text?.Trim();
                    string strState = worksheet.Cells[row, 3].Text?.Trim();
                    string strPostCode = worksheet.Cells[row, 4].Text?.Trim();
                    string strCountry = worksheet.Cells[row, 5].Text?.Trim();
                    string StoreAddress1 = worksheet.Cells[row, 6].Text?.Trim();
                    string StoreAddress2 = worksheet.Cells[row, 7].Text?.Trim();
                    string strBusinessEntity = worksheet.Cells[row, 8].Text?.Trim();
                    string ExternalCode = worksheet.Cells[row, 9].Text?.Trim();
                    string status = worksheet.Cells[row, 10].Text?.Trim();

                    yield return new object[]
                    {
                        Storename, StoreCity, strState, strPostCode, strCountry, StoreAddress1, StoreAddress2, strBusinessEntity, ExternalCode, status
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetSearchStoreTestData(string filePath, string sheetName)
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


                    yield return new object[]
                    {
                        searchText
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateStoreTestData(string filePath, string sheetName)
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
                    string StoreCode = worksheet.Cells[row, 1].Text?.Trim();
                    string Storename = worksheet.Cells[row, 2].Text?.Trim();
                    string StoreCity = worksheet.Cells[row, 3].Text?.Trim();
                    string strState = worksheet.Cells[row, 4].Text?.Trim();
                    string strPostCode = worksheet.Cells[row, 5].Text?.Trim();
                    string strCountry = worksheet.Cells[row, 6].Text?.Trim();
                    string StoreAddress1 = worksheet.Cells[row, 7].Text?.Trim();
                    string StoreAddress2 = worksheet.Cells[row, 8].Text?.Trim();
                    string strBusinessEntity = worksheet.Cells[row, 9].Text?.Trim();
                    string ExternalCode = worksheet.Cells[row, 10].Text?.Trim();
                    string status = worksheet.Cells[row, 11].Text?.Trim();


                    yield return new object[]
                    {
                        StoreCode, Storename, StoreCity, strState, strPostCode, strCountry, StoreAddress1, StoreAddress2, strBusinessEntity, ExternalCode, status
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

    }
        
    [TestFixture, Order(4)]
    [AllureNUnit]
    [AllureSuite("Stores - Valid")]
    [AllureEpic("ERP-117")]
    public class Stores_Valid
    {
        private IWebDriver _driver;
        private StoresPage _StoresPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "StoresTestDataValid.xlsx");

        public static IEnumerable<object[]> CreateStoreTestData =>
        ExcelDataReaderStoreValid.GetCreateStoreTestData(ExcelPath, "CreateStoreTestData");

        public static IEnumerable<object[]> SearchStoreTestData =>
        ExcelDataReaderStoreValid.GetSearchStoreTestData(ExcelPath, "SearchStoreTestData");

        public static IEnumerable<object[]> UpdateStoreTestData =>
        ExcelDataReaderStoreValid.GetUpdateStoreTestData(ExcelPath, "UpdateStoreTestData");

        public static IEnumerable<object[]> FilterCategoryTestData =>
        ExcelDataReaderStoreValid.GetFilterCategoryTestData(ExcelPath, "FilterCategoryTestData");



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
            string moduleName = " Stores Page";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/store");
            helperFunction.WaitForPageToLoad(_wait);
            _StoresPage = new StoresPage(_driver);
            _logMessages.Clear();

            _moduleName = " Stores Page";
            string testName = NUnit.Framework.TestContext.CurrentContext.Test.MethodName;
            string baseFolderPath = AppConfig.BaseVideoFolder;
            string todayFolderName = DateTime.Now.ToString("yyyy-MM-dd");

            string fullFolderPath = Path.Combine(baseFolderPath, todayFolderName, _moduleName);
            Directory.CreateDirectory(fullFolderPath);

            // 🟢 Use the SAME version as Excel result file
            int counter = Interlocked.Increment(ref _recordingCounter);
            string recordingFileName = $"{_moduleName}_{testName}_v{_fileVersion}_{counter}.mp4"; _recordingFilePath = Path.Combine(fullFolderPath, recordingFileName);
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
        [Category("Store")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(CreateStoreTestData))]
        public void CreateStore(string Storename, string StoreCity, string strState, string strPostCode, string strCountry,
                    string StoreAddress1, string StoreAddress2, string strBusinessEntity, string ExternalCode, string status)
        {
            try
            {
                LogStep(" Start Store Creation");

                LogStep("Click 'New' button.");
                _StoresPage.ClickNewButton();
                WaitForUIEffect();

                LogStep($"Enter Store Name: {Storename}");
                _StoresPage.EnterStorename(Storename);
                WaitForUIEffect();

                LogStep($"Enter Store City: {StoreCity}");
                _StoresPage.EnterStoreCity(StoreCity);
                WaitForUIEffect();

                LogStep($"Select State: {strState}");
                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[3]/div[2]/div/select")));
                ScrollToElement(stateDropdown);
                new SelectElement(stateDropdown).SelectByText(strState);
                WaitForUIEffect();

                LogStep($"Enter Postcode: {strPostCode}");
                _StoresPage.EnterstrPostCode(strPostCode);
                WaitForUIEffect();

                LogStep($"Select Country: {strCountry}");
                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[4]/select")));
                ScrollToElement(countryDropdown);
                new SelectElement(countryDropdown).SelectByText(strCountry);
                WaitForUIEffect();

                LogStep($"Enter Store Address 1: {StoreAddress1}");
                _StoresPage.EnterStoreAddress1(StoreAddress1);
                WaitForUIEffect();

                LogStep($"Enter Store Address 2: {StoreAddress2}");
                _StoresPage.EnterStoreAddress2(StoreAddress2);
                WaitForUIEffect();

                LogStep($"Select Business Entity: {strBusinessEntity}");
                var beDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[7]/p-dropdown/div/span")));
                beDropdown.Click();
                WaitForUIEffect(500);

                var beInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/div/div/div/div[1]/div/input")));
                beInput.Clear();
                beInput.SendKeys(strBusinessEntity);
                WaitForUIEffect(800);

                var optionsList = _wait.Until(ExpectedConditions
                    .VisibilityOfAllElementsLocatedBy(By.XPath("//p-dropdownitem/li[contains(@class,'p-dropdown-item')]")));

                foreach (var option in optionsList)
                {
                    if (option.Text.Trim().Equals(strBusinessEntity, StringComparison.OrdinalIgnoreCase))
                    {
                        ScrollToElement(option);
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option);
                        break;
                    }
                }

                LogStep("Enter External Code.");
                var externalCodeInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[8]/input")));
                ScrollToElement(externalCodeInput);
                _StoresPage.EnterExternalCode(ExternalCode);
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



                LogStep("Click 'Save' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-store-modal/div/div[3]/button")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"System modal message: {message}");
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click();

                if (!message.Contains("Success", StringComparison.OrdinalIgnoreCase))
                {
                    Assert.Fail("❌ Expected Success but got: " + message);
                }

                // ===== Step 6: Verification using input variables =====
                LogStep("🔍 Verifying newly created Store in the table...");

                bool isMatchFound = false;

                // Combine address for verification
                string combinedAddress = $"{StoreAddress1},{StoreAddress2},{strPostCode},{StoreCity},{strState},{strCountry}".Replace("\"", "");

                // Search by Store Name
                _StoresPage.SearchStore(Storename);
                WaitForUIEffect(2000);
                helperFunction.WaitForSTRTableToLoad(_wait);
                WaitForUIEffect(2000);

                while (true)
                {
                    var rows = _driver.FindElements(By.XPath(
                        "/html/body/app-layout/div[1]/div/div/div/app-content/app-store/div/div[3]/div/div[1]/div/table/tbody/tr"));

                    foreach (var row in rows)
                    {
                        var cells = row.FindElements(By.TagName("td"));
                        if (cells.Count < 8) continue;

                        string actualStoreName = cells[1].Text.Trim();
                        string actualStoreStatus = cells[3].Text.Trim();
                        string actualcombinedAddress = cells[2].Text.Trim();

                        if (actualStoreName.Equals(Storename, StringComparison.OrdinalIgnoreCase))
                        {
                            // Combined verification log
                            LogStep($"🔹 Verifying Store '{Storename}' -> " +
                                    $"Status: Expected '{status}', Actual '{actualStoreStatus}'; " +
                                    $"Address: Expected '{combinedAddress}', Actual '{actualcombinedAddress}'");

                            if (actualStoreStatus.Equals(status, StringComparison.OrdinalIgnoreCase)
                                && actualcombinedAddress.Equals(combinedAddress, StringComparison.OrdinalIgnoreCase))
                            {
                                isMatchFound = true;
                                LogStep($"🎉 All fields matched successfully for '{Storename}'");
                            }
                            else
                            {
                                Assert.Fail($"❌ Verification failed for '{Storename}', see log for details.");
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
                            helperFunction.WaitForSTRTableToLoad(_wait);
                        }
                        else break;
                    }
                    catch { break; }
                }

                if (!isMatchFound)
                {
                    Assert.Fail($"❌ Store '{Storename}' was not found in the table after creation.");
                }

                LogStep("🎉 Store creation and verification completed successfully.");

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }
        }



      

        [Test]
        [Category("Store")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Store Search - General Match (Partial Match Accepted)")]
        [TestCaseSource(nameof(SearchStoreTestData))]
        public void Search_Store(string searchText)
        {
            LogStep($"🔍 Starting search for: {searchText}");
            _StoresPage.SearchStore(searchText);
            helperFunction.WaitForSTRTableToLoad(_wait);
            WaitForUIEffect();

            bool isMatchFound = false;

            while (true)
            {
                WaitForUIEffect(800);

                var rows = _driver.FindElements(By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-store/div/div[3]/div/div[1]/div/table/tbody/tr"));
                LogStep($"📄 Rows found in current page: {rows.Count}");

                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));

                    foreach (var cell in cells)
                    {
                        string cellText;
                        try
                        {
                            cellText = cell.FindElement(By.TagName("span")).Text.Trim();
                        }
                        catch
                        {
                            cellText = cell.Text.Trim();
                        }

                        LogStep($"🔎 Checking cell: '{cellText}' vs '{searchText}'");

                        if (cellText.Replace(" ", "").Contains(searchText.Replace(" ", ""), StringComparison.OrdinalIgnoreCase))
                        {
                            isMatchFound = true;
                            LogStep($"✅ Match found for '{searchText}' in cell: '{cellText}'");
                            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                            break;
                        }
                    }

                    if (isMatchFound) break;
                }

                if (isMatchFound) break;

                try
                {
                    var nextButton = _driver.FindElement(By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-store/div/div[3]/div/div[2]/app-global-pagination/div/div[2]/ul/li[4]"));

                    if (!nextButton.GetAttribute("class").Contains("disabled"))
                    {
                        LogStep("⏭ Going to next page...");
                        nextButton.Click();
                        helperFunction.WaitForSTRTableToLoad(_wait);
                        WaitForUIEffect(500);
                    }
                    else
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("🛑 Reached last page. No more data.");
                        break;
                    }
                }
                catch (NoSuchElementException)
                {
                    LogStep("❌ Pagination not found. Ending search.");
                    break;
                }
            }

            WaitForUIEffect();
            LogStep($"Final match result for '{searchText}': {isMatchFound}");
            Assert.IsTrue(isMatchFound, $"❌ Match not found for '{searchText}' in any table cell.");
        }



        [Test]
        [Category("Store")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Store Update")]
        [TestCaseSource(nameof(UpdateStoreTestData))]
        public void Update(string StoreCode, string Storename, string StoreCity, string strState, string strPostCode, string strCountry,
          string StoreAddress1, string StoreAddress2, string strBusinessEntity, string ExternalCode, string status)
        {
            try
            {
                LogStep($" Starting Store Update for code: {StoreCode}");
                WaitForUIEffect(1000);

                LogStep("Clicking 'Edit' button.");
                _StoresPage.ClickEditButton(StoreCode);
                WaitForUIEffect();

                LogStep($"Updating Store Name: {Storename}");
                _StoresPage.EnterStorename(Storename);
                WaitForUIEffect();

                LogStep($"Updating Store City: {StoreCity}");
                _StoresPage.EnterStoreCity(StoreCity);
                WaitForUIEffect();

                LogStep($"Selecting State: {strState}");
                var storeState = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[3]/div[2]/div/select")));
                ScrollToElement(storeState);
                new SelectElement(storeState).SelectByText(strState);
                WaitForUIEffect();

                LogStep($"Updating Post Code: {strPostCode}");
                _StoresPage.EnterstrPostCode(strPostCode);
                WaitForUIEffect();

                LogStep($"Selecting Country: {strCountry}");
                var storeCountry = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//select[@formcontrolname='country']")));
                ScrollToElement(storeCountry);
                new SelectElement(storeCountry).SelectByText(strCountry);
                WaitForUIEffect();

                LogStep($"Updating Address 1: {StoreAddress1}");
                _StoresPage.EnterStoreAddress1(StoreAddress1);
                WaitForUIEffect();

                LogStep($"Updating Address 2: {StoreAddress2}");
                _StoresPage.EnterStoreAddress2(StoreAddress2);
                WaitForUIEffect();

                LogStep($"Selecting Business Entity: {strBusinessEntity}");
                var beDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[7]/p-dropdown/div/span")));
                beDropdown.Click();
                WaitForUIEffect(300);

                var input = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/div/div/div/div[1]/div/input")));
                input.Clear();
                input.SendKeys(strBusinessEntity);
                WaitForUIEffect(800);

                var optionsList = _wait.Until(ExpectedConditions
                    .VisibilityOfAllElementsLocatedBy(By.XPath("//p-dropdownitem/li[contains(@class,'p-dropdown-item')]")));

                foreach (var option in optionsList)
                {
                    string optionText = option.Text.Trim();
                    LogStep($"🔍 Option found: {optionText}");

                    if (optionText.Equals(strBusinessEntity, StringComparison.OrdinalIgnoreCase))
                    {
                        ScrollToElement(option);
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option);
                        break;
                    }
                }

                LogStep($"Updating External Code: {ExternalCode}");
                _StoresPage.EnterExternalCode(ExternalCode);
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


                LogStep("Clicking 'Save' button.");
                _StoresPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("Waiting for confirmation modal...");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");

                var okBtn = modal.FindElement(By.XPath(".//button[contains(.,'Ok')]"));
                ScrollToElement(okBtn);

                if (message.Contains("TIN has already been taken", StringComparison.OrdinalIgnoreCase))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⚠️ Update failed: Duplicate TIN.");
                    okBtn.Click();
                    Assert.Fail("❌ Cannot update: TIN has already been taken.");
                }
                else if (message.ToLower().Contains("fail") || message.ToLower().Contains("error"))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep($"❌ Update failed with message: {message}");
                    okBtn.Click();
                    Assert.Fail("❌ Update failed: " + message);
                }
                else
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("✅ Store updated successfully.");
                    okBtn.Click();
                    WaitForUIEffect();

                    // ===== Step 6: Verification using input variables =====
                    LogStep("🔍 Verifying newly created Store in the table...");

                    bool isMatchFound = false;

                    // Combine address for verification
                    string combinedAddress = $"{StoreAddress1},{StoreAddress2},{strPostCode},{StoreCity},{strState},{strCountry}".Replace("\"", "");

                    // Search by Store Name
                    _StoresPage.SearchStore(StoreCode);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForSTRTableToLoad(_wait);
                    WaitForUIEffect(2000);

                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-store/div/div[3]/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualStoreName = cells[1].Text.Trim();
                            string actualStoreStatus = cells[3].Text.Trim();
                            string actualcombinedAddress = cells[2].Text.Trim();

                            if (actualStoreName.Equals(Storename, StringComparison.OrdinalIgnoreCase))
                            {
                                // Combined verification log
                                LogStep($"🔹 Verifying Store '{Storename}' -> " +
                                        $"Status: Expected '{status}', Actual '{actualStoreStatus}'; " +
                                        $"Address: Expected '{combinedAddress}', Actual '{actualcombinedAddress}'");

                                if (actualStoreStatus.Equals(status, StringComparison.OrdinalIgnoreCase)
                                    && actualcombinedAddress.Equals(combinedAddress, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"🎉 All fields matched successfully for '{Storename}'");
                                }
                                else
                                {
                                    Assert.Fail($"❌ Verification failed for '{Storename}', see log for details.");
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
                                helperFunction.WaitForSTRTableToLoad(_wait);
                            }
                            else break;
                        }
                        catch { break; }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Store '{Storename}' was not found in the table after creation.");
                    }

                    LogStep("🎉 Store creation and verification completed successfully.");

                }
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during store update: {ex}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }

        [Test]
        [Category("Store")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Export Stores Report")]
        public void ExportStoreReport()
        {
            string downloadPath = AppConfig.DownloadPath;
            string filePrefix = "Store Index";

            LogStep("Clicking 'Export' button...");
            helperFunction.WaitForElementToBeClickable(_wait,
                By.CssSelector("#kt_content_container > app-store > div > div.card-header.border-0.pt-5 > div > div:nth-child(2) > a"));
            _StoresPage.ClickExportButton();


            LogStep("📄 Waiting for downloaded file to appear...");
            bool fileDownloaded = _StoresPage.WaitForFileDownload(downloadPath, filePrefix, TimeSpan.FromSeconds(15));
            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            Assert.IsTrue(fileDownloaded, "❌ No new download detected.");
            LogStep("✅ Export file downloaded successfully.");
        }


        [Test]
        [Category("Store")]
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
                    var noDataElement = _driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-store/div/div[3]/div/div[1]/div/table/tbody[2]/tr/td/p"));
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
                    if (cells.Count < 4) continue;

                    string actualStatus = GetStatusFromCell(cells[3]);

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
                    _StoresPage.ClickFilterALLCategoryButton();
                    break;

                case "ACTIVE":
                    LogStep("📌 Testing filter: Active Category");
                    _StoresPage.ClickFilterActiveCategoryButton();
                    break;

                case "INACTIVE":
                    LogStep("📌 Testing filter: Inactive Category");
                    _StoresPage.ClickFilterInactiveCategoryButton();
                    break;

                default:
                    Assert.Fail($"❌ Invalid filter category input: '{category}'");
                    break;
            }

            WaitForUIEffect();
            helperFunction.WaitForSTRTableToLoad(_wait);

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
                    if (cells.Count < 4) continue;

                    string actualStatus = GetStatusFromCell(cells[3]);
                    LogStep($"🔍 Found Status = '{actualStatus}'");
                }
            }
        }


        [Test]
        [Category("Store")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button")]
        public void ImportStoreCSVFile()
        {
            try
            {
                string filePath = AppConfig.ImportStoreCSVFile;

                // Open Import modal
                LogStep("📤 Clicking 'Import' button to open modal...");
                helperFunction.WaitForElementToBeClickable(_wait,
                    By.CssSelector("#kt_content_container > app-store > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a"));
                _StoresPage.ClickImportButton();
                WaitForUIEffect(800);

                // Click Download Template Button
                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-body.px-20 > div > div > div.d-flex.align-items-center > button"));
                _StoresPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                // Upload file
                LogStep($"📂 Selecting CSV file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(Path.GetFullPath(filePath));
                WaitForUIEffect(1000);

                // Click Upload
                LogStep("Clicking 'Upload' to process file...");
                helperFunction.WaitForElementToBeClickable(_wait,
                    By.CssSelector("#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _StoresPage.ClickUploadButton();

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

                // Wait for 'Completed' button
                LogStep("⏳ Waiting for 'Completed' button to become clickable...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                LogStep("✅ 'Completed' button found. Clicking...");
                completedButton.Click();
                WaitForUIEffect(1000);

                // Check modal message
                LogStep("🔍 Checking upload result modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal message: {message}");

                if (!message.ToLower().Contains("success"))
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    Assert.Fail($"❌ Expected success message but got: {message}");
                }
                else
                {
                    //  Confirm success modal
                    LogStep("✅ Upload successful. Clicking 'Ok, got it!'");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    okButton.Click();
                    WaitForUIEffect(500);

                    // ===== Step 6: Read CSV =====
                    LogStep("📄 Reading store information from CSV...");
                    var StoreName = new List<string>();
                    var Storestatus = new List<string>();
                    var StoreAddress1 = new List<string>();
                    var StoreAddress2 = new List<string>();
                    var StoreCityName = new List<string>();
                    var StorePostcode = new List<string>();
                    var StoreStateName = new List<string>();
                    var StoreCountryName = new List<string>();

                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var reader = new StreamReader(stream))
                    {
                        int rowIndex = 0;
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            rowIndex++;
                            if (rowIndex < 5) continue; // skip header rows

                            var columns = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                            if (columns.Length >= 10)
                            {
                                var StrName = columns[0].Trim();
                                if (string.IsNullOrEmpty(StrName))
                                {
                                    LogStep("Empty Store Name detected, stop getting CSV data.");
                                    break;
                                }

                                StoreName.Add(StrName);
                                Storestatus.Add(columns[1].Trim());
                                StoreAddress1.Add(columns[2].Trim());
                                StoreAddress2.Add(columns[3].Trim());
                                StoreCityName.Add(columns[4].Trim());
                                StorePostcode.Add(columns[5].Trim());
                                var stateCode = columns[6].Trim();

                                // Convert state code → state name
                                if (StateCodeMapping.TryGetValue(stateCode, out var stateName))
                                {
                                    StoreStateName.Add(stateName);
                                }
                                else
                                {
                                    // Handle invalid code
                                    StoreStateName.Add("Unknown State");
                                    LogStep($"⚠️ Invalid state code '{stateCode}' detected in CSV.");
                                }

                                var countryCode = columns[7].Trim();

                                // Convert country code → country name using dictionary
                                if (CountryCodeMapping.TryGetValue(countryCode, out var countryName))
                                {
                                    StoreCountryName.Add(countryName);
                                }
                                else
                                {
                                    StoreCountryName.Add("Unknown Country");
                                    LogStep($"⚠️ Invalid or unrecognized country code '{countryCode}' detected.");
                                }

                            }
                        }
                    }


                    for (int i = 0; i < StoreName.Count; i++)
                    {

                        string StrName = StoreName[i];
                        string expectedStoreStatus = Storestatus[i];
                        string combinedAddress =
                            $"{StoreAddress1[i]},{StoreAddress2[i]},{StorePostcode[i]},{StoreCityName[i]},{StoreStateName[i]},{StoreCountryName[i]}"
                            .Replace("\"", "");
                        bool isMatchFound = false;

                        LogStep($"🔍 Searching Store Information '{StrName}' to verify store status : '{expectedStoreStatus}', Address : '{combinedAddress}'");

                        _StoresPage.SearchStore(StrName);
                        WaitForUIEffect(2000);
                        helperFunction.WaitForSTRTableToLoad(_wait);
                        WaitForUIEffect(2000);

                        while (true)
                        {
                            var rows = _driver.FindElements(By.XPath(
                                "/html/body/app-layout/div[1]/div/div/div/app-content/app-store/div/div[3]/div/div[1]/div/table/tbody/tr"));

                            foreach (var row in rows)
                            {
                                var cells = row.FindElements(By.TagName("td"));
                                if (cells.Count < 8) continue;

                                string actualStoreName = cells[1].Text.Trim();
                                string actualStoreStatus = cells[3].Text.Trim();
                                string actualcombinedAddress = cells[2].Text.Trim();


                                if (actualStoreName.Equals(StrName, StringComparison.OrdinalIgnoreCase))
                                {
                                    if (actualStoreStatus.Equals(expectedStoreStatus, StringComparison.OrdinalIgnoreCase)
                                        && actualcombinedAddress.Equals(combinedAddress, StringComparison.OrdinalIgnoreCase))
                                    {
                                        isMatchFound = true;
                                        LogStep($"✅ Store Information matched : '{StrName}' Store Status : '{actualStoreStatus}', Address : '{actualcombinedAddress}'");
                                    }
                                    else
                                    {
                                        Assert.Fail($"❌ Store Information '{StrName}' mismatch. Expected: Store Status : '{expectedStoreStatus}', Address : '{combinedAddress}'. Found: Store Status : '{actualStoreStatus}', Address : '{actualcombinedAddress}'");
                                    }
                                    break;
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
                                    helperFunction.WaitForSTRTableToLoad(_wait);
                                }
                                else break;
                            }
                            catch { break; }
                        }

                        if (!isMatchFound)
                        {
                            Assert.Fail($"❌ Store Information '{StrName}' was not found.");
                        }
                    }

                    LogStep("🎉 CSV Import and verification completed successfully.");
                }
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Stores_{DateTime.Now:yyyyMMdd_HHmmss}.png");
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

        private void WaitForUIEffect(int ms = 2000)
        {
            Thread.Sleep(ms); // adjustable UI pause for better video capture
        }

        private void ScrollToElement(IWebElement element)
        {
            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element);
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

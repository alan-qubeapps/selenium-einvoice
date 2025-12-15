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
using SeleniumTests.Pages.Transaction;
using System.Drawing;
using System.Globalization;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.A_BusinessEntity
{

    public static class ExcelDataReaderBusinessEntityValid
    {
        public static IEnumerable<object[]> GetCreateBusinessEntityTestData(string filePath, string sheetName)
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
                    string status = worksheet.Cells[row, 17].Text?.Trim();


                    yield return new object[]
                    {
                        BEname, BETinNumber, BEregisterType, BERegisterID, BEsst, BETTRegisterNumber, BEMSIC, BEContactNumber, BEemail, BECity, BEState, BEPosCode, BECountry, BEAddress1, BEAddress2, BEAddress3, status
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetSearchBusinessEntityTestData(string filePath, string sheetName)
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


        public static IEnumerable<object[]> GetUpdateBusinessEntityTestData(string filePath, string sheetName)
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
                    string status = worksheet.Cells[row, 17].Text?.Trim();


                    yield return new object[]
                    {
                        BEname, BETinNumber, BEsst, BETTRegisterNumber, BEMSIC, BEContactNumber, BEemail, BECity, BEState, BEPosCode, BECountry, BEAddress1, BEAddress2, BEAddress3, status
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



    }


        
    [TestFixture, Order(1)]
    [AllureNUnit]
    [AllureSuite("Business Entity - Business Entity - Valid")]
    [AllureEpic("ERP-117")]
    public class BusinessEntityTests_Valid
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

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "BusinessEntityTestDataValid.xlsx");

        public static IEnumerable<object[]> CreateBusinessEntityTestData =>
        ExcelDataReaderBusinessEntityValid.GetCreateBusinessEntityTestData(ExcelPath, "CreateBusinessEntityTestData");

        public static IEnumerable<object[]> SearchBusinessEntityTestData =>
        ExcelDataReaderBusinessEntityValid.GetSearchBusinessEntityTestData(ExcelPath, "SearchBusinessEntityTestData");

        public static IEnumerable<object[]> UpdateBusinessEntityTestData =>
        ExcelDataReaderBusinessEntityValid.GetUpdateBusinessEntityTestData(ExcelPath, "UpdateBusinessEntityTestData");

        public static IEnumerable<object[]> SearchCategoryTestData =>
        ExcelDataReaderBusinessEntityValid.GetSearchCategoryTestData(ExcelPath, "SearchCategoryTestData");

        public static IEnumerable<object[]> FilterLHDNStatusTestData =>
        ExcelDataReaderBusinessEntityValid.GetFilterLHDNStatusTestData(ExcelPath, "FilterLHDNStatusTestData");


        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Business Entity Page";

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

            _moduleName = "Business Entity Page";
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
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(CreateBusinessEntityTestData))]
        public void Create(
            string BEname, string BETinNumber, string BEregisterType, string BERegisterID, string BEsst, string BETTRegisterNumber, string BEMSIC, string BEContactNumber,
            string BEemail, string BECity, string BEState, string BEPosCode, string BECountry, string BEAddress1, string BEAddress2, string BEAddress3, string status)
        {
            try
            {
                // ===== Step 1-4: Fill form and save (your existing code) =====
                _BusinessEntityPage.ClickNewButton();
                WaitForUIEffect(); LogStep("Clicked 'New' button");

                _BusinessEntityPage.EnterBEname(BEname);
                WaitForUIEffect(); LogStep($"Entered BE Name: {BEname}");

                _BusinessEntityPage.EnterBETinNumber(BETinNumber);
                WaitForUIEffect(); LogStep($"Entered TIN Number: {BETinNumber}");

                var regType = _wait.Until(ExpectedConditions.ElementToBeClickable(
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
                WaitForUIEffect(); LogStep("Clicked 'Continue' to Step 2");

                // MSIC selection
                var BEMSICDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("p-dropdown[placeholder='Select MSIC']")));
                BEMSICDropdown.Click();
                WaitForUIEffect(); LogStep("Opened MSIC dropdown");

                var dropdownPanel = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("div.p-dropdown-panel")));
                var dropdownOptions = dropdownPanel.FindElements(By.CssSelector("li.p-dropdown-item:not(.p-disabled)"));

                foreach (var option in dropdownOptions)
                {
                    if (option.Text.Trim().Equals(BEMSIC.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        try { option.Click(); }
                        catch { ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option); }
                        LogStep($"Selected MSIC: {BEMSIC}");
                        break;
                    }
                }

                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber);
                WaitForUIEffect(); LogStep($"Entered Contact Number: {BEContactNumber}");

                _BusinessEntityPage.EnterBEemail(BEemail);
                WaitForUIEffect(); LogStep($"Entered Email: {BEemail}");

                _BusinessEntityPage.EnterBECity(BECity);
                WaitForUIEffect(); LogStep($"Entered City: {BECity}");

                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='selectedState']")));
                new SelectElement(stateDropdown).SelectByText(BEState);
                WaitForUIEffect(); LogStep($"Selected State: {BEState}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode);
                WaitForUIEffect(); LogStep($"Entered Postal Code: {BEPosCode}");

                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("select[name='country']")));
                new SelectElement(countryDropdown).SelectByText(BECountry);
                WaitForUIEffect(); LogStep($"Selected Country: {BECountry}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1);
                WaitForUIEffect(); LogStep($"Entered Address Line 1: {BEAddress1}");
                _BusinessEntityPage.EnterBEAddress2(BEAddress2);
                WaitForUIEffect(); LogStep($"Entered Address Line 2: {BEAddress2}");
                _BusinessEntityPage.EnterBEAddress3(BEAddress3);
                WaitForUIEffect(); LogStep($"Entered Address Line 3: {BEAddress3}");

                // Upload logo
                string filePath = AppConfig.BusinessEntityImage;
                if (!File.Exists(filePath)) Assert.Fail("File not found: " + filePath);
                var fileInput = _wait.Until(ExpectedConditions.ElementExists(
                    By.CssSelector("#kt_create_account_form input[type=file]")));
                fileInput.SendKeys(filePath);
                WaitForUIEffect(); LogStep("File upload initiated");

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

                _BusinessEntityPage.ClickSaveButton();
                WaitForUIEffect(); LogStep("Clicked save button");

                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();
                LogStep($"System modal message: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click();

                if (!message.Contains("Success", StringComparison.OrdinalIgnoreCase))
                {
                    Assert.Fail("❌ Expected Success but got: " + message);
                }

                // ===== Step 5: Verification using input variables =====
                LogStep("🔍 Verifying newly created Business Entity in the table...");

                bool isMatchFound = false;

                _BusinessEntityPage.SearchBusinessEntity(BETinNumber);
                WaitForUIEffect(2000);
                helperFunction.WaitForBETableToLoad(_wait);
                WaitForUIEffect(2000);

                while (!isMatchFound)
                {
                    var rows = _driver.FindElements(By.XPath(
                        "/html/body/app-layout/div[1]/div/div/div/app-content/app-business-entity/div/div[4]/div/div[1]/div/table/tbody/tr"));

                    foreach (var row in rows)
                    {
                        var cells = row.FindElements(By.TagName("td"));
                        if (cells.Count < 8) continue;

                        string actualEntityName = cells[0].Text.Trim();
                        string actualBETinNo = cells[1].Text.Trim();
                        string actualcombinedBE = cells[2].Text.Trim();
                        string actualBESSTNo = cells[3].Text.Trim();
                        string actualBEEmail = cells[4].Text.Trim();
                        string actualBEContact = cells[5].Text.Trim();
                        string actualBEStatus = cells[6].Text.Trim();

                        if (actualBETinNo.Equals(BETinNumber, StringComparison.OrdinalIgnoreCase))
                        {
                            // Combine all verification into a single log line
                            LogStep($"🔹 Verifying Business Entity '{BEname}' -> " +
                                    $"TIN: Expected '{BETinNumber}', Actual '{actualBETinNo}'; " +
                                    $"SST: Expected '{BEsst}', Actual '{actualBESSTNo}'; " +
                                    $"Email: Expected '{BEemail}', Actual '{actualBEEmail}'; " +
                                    $"Contact: Expected '{BEContactNumber}', Actual '{actualBEContact}'" +
                                    $"Status: Expected '{status}', Actual '{actualBEStatus}'; ");

                            // Mark as found only if all fields match
                            if (actualEntityName.Equals(BEname, StringComparison.OrdinalIgnoreCase) &&
                                actualBETinNo.Equals(BETinNumber, StringComparison.OrdinalIgnoreCase) &&
                                actualBESSTNo.Equals(BEsst, StringComparison.OrdinalIgnoreCase) &&
                                actualBEEmail.Equals(BEemail, StringComparison.OrdinalIgnoreCase) &&
                                actualBEContact.Equals(BEContactNumber, StringComparison.OrdinalIgnoreCase) &&
                                actualBEStatus.Equals(status, StringComparison.OrdinalIgnoreCase))
                            {
                                isMatchFound = true;
                                LogStep($"🎉 All fields matched successfully for '{BEname}'");
                            }
                            else
                            {
                                Assert.Fail($"❌ Verification failed for '{BEname}', see log for details.");
                            }

                            break; // stop checking rows
                        }
                    }

                    if (!isMatchFound)
                    {
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForBETableToLoad(_wait);
                            }
                            else break;
                        }
                        catch { break; }
                    }
                }

                if (!isMatchFound)
                {
                    Assert.Fail($"❌ Business Entity '{BEname}' not found in the table after creation.");
                }
                LogStep("🎉 Business Entity creation and verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Exception during test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }

        [Test]
        [Category("BusinessEntity")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("BusinessEntity Search - General Match (Partial Match Accepted)")]
        [TestCaseSource(nameof(SearchBusinessEntityTestData))]
        public void Search_Business_Entity(string searchText)
        {
            LogStep($"🔍 Starting search for: '{searchText}'");

            _BusinessEntityPage.SearchBusinessEntity(searchText);
            helperFunction.WaitForBETableToLoad(_wait);
            WaitForUIEffect();
            LogStep("Search input filled and table loaded.");

            bool isMatchFound = false;

            while (true)
            {
                Thread.Sleep(800); // Optional: let UI stabilize

                var rows = _driver.FindElements(By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-business-entity/div/div[4]/div/div[1]/div/table/tbody"));

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

                        Console.WriteLine($"🔍 Checking cell: '{cellText}' vs '{searchText}'");

                        if (cellText.Replace(" ", "").Contains(searchText.Replace(" ", ""), StringComparison.OrdinalIgnoreCase))
                        {
                            isMatchFound = true;
                            LogStep($"✅ Match found in table cell: '{cellText}'");
                            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
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
                    // Locate the pagination next button more reliably by class
                    var nextButton = _driver.FindElement(By.CssSelector("li.page-item.next a.page-link"));

                    // Check if button is disabled
                    var isDisabled = _driver.FindElement(By.CssSelector("li.page-item.next")).GetAttribute("class").Contains("disabled");

                    if (!isDisabled)
                    {
                        nextButton.Click();
                        helperFunction.WaitForBETableToLoad(_wait);
                        WaitForUIEffect();
                        LogStep("Navigated to next pagination page.");
                    }
                    else
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        LogStep("⛔ No more pages. Match not found.");
                        break;
                    }
                }
                catch (NoSuchElementException)
                {
                    LogStep("⚠️ Pagination not found. Possibly only one page.");
                    break;
                }

            }

            Thread.Sleep(3000);
            Assert.IsTrue(isMatchFound, $"❌ No matching record found for '{searchText}' in any table cell.");
            LogStep("✅ Final assertion passed: match found.");

        }


        [Test]
        [Category("BusinessEntity")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.blocker)]
        [AllureStory("BusinessEntity Update")]
        [TestCaseSource(nameof(UpdateBusinessEntityTestData))]
        public void Update_BusinessEntity(
            string BEname, string BETinNumber, string BEsst, string BETTRegisterNumber, string BEMSIC,
            string BEContactNumber, string BEemail, string BECity, string BEState, string BEPosCode,
            string BECountry, string BEAddress1, string BEAddress2, string BEAddress3, string status)
        {
            try
            {
                // Step 0: Open Edit
                _BusinessEntityPage.ClickEditButton(BETinNumber);
                WaitForUIEffect(); LogStep("Clicked Edit button");

                // Step 1: Basic info
                _BusinessEntityPage.EnterBEname(BEname); WaitForUIEffect();
                _BusinessEntityPage.EnterBEsst(BEsst); WaitForUIEffect();
                _BusinessEntityPage.EnterBETTRegisterNumber(BETTRegisterNumber); WaitForUIEffect();
                _BusinessEntityPage.ClickContinueButton(); WaitForUIEffect();
                LogStep("Clicked Continue to Step 2");

                // Step 2: Select BEMSIC
                var BEMSICDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//p-dropdown/div")));
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

                // Step 2: Fill valid contact & other info
                _BusinessEntityPage.EnterBEContactNumber(BEContactNumber); WaitForUIEffect(); LogStep($"Entered Contact Number (Valid): {BEContactNumber}");

                _BusinessEntityPage.EnterBEemail(BEemail); WaitForUIEffect(); LogStep($"Entered Email: {BEemail}");
                _BusinessEntityPage.EnterBECity(BECity); WaitForUIEffect(); LogStep($"Entered City: {BECity}");

                var stateDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[4]/div[2]/div/select")));
                new SelectElement(stateDropdown).SelectByText(BEState);
                WaitForUIEffect(); LogStep($"Selected State: {BEState}");

                _BusinessEntityPage.EnterBEPosCode(BEPosCode); WaitForUIEffect(); LogStep($"Entered Postal Code: {BEPosCode}");

                var countryDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[5]/select")));
                new SelectElement(countryDropdown).SelectByText(BECountry);
                WaitForUIEffect(); LogStep($"Selected Country: {BECountry}");

                _BusinessEntityPage.EnterBEAddress1(BEAddress1); WaitForUIEffect(); LogStep("Entered Address Line 1");
                _BusinessEntityPage.EnterBEAddress2(BEAddress2); WaitForUIEffect(); LogStep("Entered Address Line 2");
                _BusinessEntityPage.EnterBEAddress3(BEAddress3); WaitForUIEffect(); LogStep("Entered Address Line 3");

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

                // ✅ Save should be enabled
                Assert.IsTrue(_BusinessEntityPage.IsSaveButtonEnabled(), "❌ Save button should be enabled for valid input.");

                _BusinessEntityPage.ClickSaveButton(); WaitForUIEffect(); LogStep("Clicked Save button");

                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();
                LogStep($"System displayed modal message: {message}");

                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click(); WaitForUIEffect(); LogStep("Clicked OK button on modal");


                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_ValidContact_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("📸 Captured screenshot");

                // ===== Step 5: Verification using input variables =====
                LogStep("🔍 Verifying newly created Business Entity in the table...");

                bool isMatchFound = false;

                _BusinessEntityPage.SearchBusinessEntity(BETinNumber);
                WaitForUIEffect(2000);
                helperFunction.WaitForBETableToLoad(_wait);
                WaitForUIEffect(2000);

                while (!isMatchFound)
                {
                    var rows = _driver.FindElements(By.XPath(
                        "/html/body/app-layout/div[1]/div/div/div/app-content/app-business-entity/div/div[4]/div/div[1]/div/table/tbody/tr"));

                    foreach (var row in rows)
                    {
                        var cells = row.FindElements(By.TagName("td"));
                        if (cells.Count < 8) continue;

                        string actualEntityName = cells[0].Text.Trim();
                        string actualBETinNo = cells[1].Text.Trim();
                        string actualcombinedBE = cells[2].Text.Trim();
                        string actualBESSTNo = cells[3].Text.Trim();
                        string actualBEEmail = cells[4].Text.Trim();
                        string actualBEContact = cells[5].Text.Trim();
                        string actualBEStatus = cells[6].Text.Trim();

                        if (actualBETinNo.Equals(BETinNumber, StringComparison.OrdinalIgnoreCase))
                        {
                            // Combine all verification into a single log line
                            LogStep($"🔹 Verifying Business Entity '{BEname}' -> " +
                                    $"TIN: Expected '{BETinNumber}', Actual '{actualBETinNo}'; " +
                                    $"SST: Expected '{BEsst}', Actual '{actualBESSTNo}'; " +
                                    $"Email: Expected '{BEemail}', Actual '{actualBEEmail}'; " +
                                    $"Contact: Expected '{BEContactNumber}', Actual '{actualBEContact}'" +
                                    $"Status: Expected '{status}', Actual '{actualBEStatus}'; ");

                            // Mark as found only if all fields match
                            if (actualEntityName.Equals(BEname, StringComparison.OrdinalIgnoreCase) &&
                                actualBETinNo.Equals(BETinNumber, StringComparison.OrdinalIgnoreCase) &&
                                actualBESSTNo.Equals(BEsst, StringComparison.OrdinalIgnoreCase) &&
                                actualBEEmail.Equals(BEemail, StringComparison.OrdinalIgnoreCase) &&
                                actualBEContact.Equals(BEContactNumber, StringComparison.OrdinalIgnoreCase) &&
                                actualBEStatus.Equals(status, StringComparison.OrdinalIgnoreCase))
                            {
                                isMatchFound = true;
                                LogStep($"🎉 All fields matched successfully for '{BEname}'");
                            }
                            else
                            {
                                Assert.Fail($"❌ Verification failed for '{BEname}', see log for details.");
                            }

                            break; // stop checking rows
                        }
                    }

                    if (!isMatchFound)
                    {
                        try
                        {
                            var nextButton = _driver.FindElement(By.CssSelector("i.next"));
                            if (!nextButton.GetAttribute("class").Contains("disabled"))
                            {
                                nextButton.Click();
                                WaitForUIEffect(1500);
                                helperFunction.WaitForBETableToLoad(_wait);
                            }
                            else break;
                        }
                        catch { break; }
                    }
                }

                if (!isMatchFound)
                {
                    Assert.Fail($"❌ Business Entity '{BEname}' not found in the table after creation.");
                }


                LogStep("🎉 Business Entity creation and verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_ValidContact_Exception_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.Fail($"❌ Exception during positive update test: {ex.Message}");
            }
        }




        [Test]
        [Category("BusinessEntity")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Export Business Entity Report")]
        public void ExportBusinessEntityReport()
        {
            string downloadPath = AppConfig.DownloadPath;
            string filePrefix = "Business Entity Index";

            try
            {
                // Click Export button
                helperFunction.WaitForElementToBeClickable(_wait,
                    By.CssSelector("#kt_content_container > app-business-entity > div > div.card-header.border-0.pt-5 > div > div:nth-child(2) > a"));
                WaitForUIEffect();

                _BusinessEntityPage.ClickExportButton();
                WaitForUIEffect();
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("📤 Clicked Export Business Entity Report button.");

                // Wait for the file to download
                bool fileDownloaded = _BusinessEntityPage.WaitForFileDownload(downloadPath, filePrefix, TimeSpan.FromSeconds(15));
                WaitForUIEffect();

                LogStep(fileDownloaded
                    ? $"✅ Exported file with prefix '{filePrefix}' downloaded successfully in {downloadPath}."
                    : $"❌ Export failed. No file starting with '{filePrefix}' was found in {downloadPath}.");

                Thread.Sleep(3000);
                Assert.IsTrue(fileDownloaded, $"❌ No new download with prefix '{filePrefix}' detected.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during export: {ex.Message}");
                Thread.Sleep(3000);
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("BusinessEntity")]
        [Order(5)]
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

            void ValidateStatusColumn(string expectedStatus)
            {
                var rows = GetRows();
                foreach (var row in rows)
                {
                    var cells = row.FindElements(By.TagName("td"));
                    if (cells.Count < 7) continue;

                    string actualStatus = GetStatusFromCell(cells[6]);

                    LogStep($"🔍 Validating Status: Expected = '{expectedStatus}', Found = '{actualStatus}'");

                    if (!actualStatus.Equals(expectedStatus, StringComparison.OrdinalIgnoreCase))
                    {
                        LogStep($"❌ Mismatch - Expected: '{expectedStatus}', Found: '{actualStatus}'");
                        Assert.Fail("❌ One or more rows have unexpected status.");
                    }
                }
            }

            try
            {
                // === Apply Filter ===
                LogStep($"📌 Testing filter: {category} Category");
                switch (category.Trim().ToUpperInvariant())
                {
                    case "ALL":
                        _BusinessEntityPage.ClickFilterALLCategoryButton();
                        break;
                    case "ACTIVE":
                        _BusinessEntityPage.ClickFilterActiveCategoryButton();
                        break;
                    case "INACTIVE":
                        _BusinessEntityPage.ClickFilterInactiveCategoryButton();
                        break;
                    default:
                        throw new ArgumentException($"❌ Unknown category: {category}");
                }

                WaitForUIEffect();
                helperFunction.WaitForBETableToLoad(_wait);
                Thread.Sleep(2000);

                // === Screenshot ===
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{category}_{DateTime.Now:yyyyMMdd_HHmmss}.png");
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
                    $"Business_Entity_{category}_ERROR_{DateTime.Now:yyyyMMdd_HHmmssfff}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Exception on '{category}' filter. Screenshot saved: {_lastScreenshotPath}");
                Assert.Fail(ex.Message);
            }
        }






        [Test]
        [Category("BusinessEntity")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Filter - Filter By LHDN Status")]
        [TestCaseSource(nameof(FilterLHDNStatusTestData))]
        public void FilterLHDNStatusFunction(string lhdnStatus)
        {
            // === Get only data rows from tbody[1] ===
            IReadOnlyCollection<IWebElement> GetRows() =>
                _driver.FindElements(By.XPath("//table/tbody[1]/tr"));

            // === Check if "No data available" is shown in tbody[2] ===
            bool IsNoDataMessageShown()
            {
                try
                {
                    var noDataElement = _driver.FindElement(By.XPath("//*[@id=\"kt_content_container\"]/app-business-entity/div/div[4]/div/div[1]/div/table/tbody/tr/td/p"));
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
                    if (cells.Count < 8) continue;

                    string actualStatus;
                    try
                    {
                        actualStatus = cells[7].FindElement(By.TagName("span")).Text.Trim();
                    }
                    catch
                    {
                        actualStatus = cells[7].Text.Trim();
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

            try
            {
                // === Apply Filter Based on Input ===
                LogStep($"📌 Testing LHDN filter: {lhdnStatus}");

                switch (lhdnStatus.ToUpperInvariant())
                {
                    case "PENDING":
                        _BusinessEntityPage.ClickFilterPendingCategoryButton();
                        break;
                    case "SUCCESS":
                        _BusinessEntityPage.ClickFilterSuccessCategoryButton();
                        break;
                    case "FAILED":
                        _BusinessEntityPage.ClickFilterFailedCategoryButton();
                        break;
                    default:
                        Assert.Fail($"❌ Invalid LHDN status input: '{lhdnStatus}'");
                        break;
                }

                WaitForUIEffect();
                helperFunction.WaitForBETableToLoad(_wait);

                // 📸 Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_LHDN_{lhdnStatus}_{DateTime.Now:yyyyMMdd_HHmmssfff}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved: {_lastScreenshotPath}");

                // === Handle "No Data Available" ===
                if (IsNoDataMessageShown())
                {
                    LogStep($"✅ 'No data available' is shown for '{lhdnStatus}' filter. Test succeeded.");
                    return; // Exit early → valid scenario
                }

                // === Validate rows if present ===
                var dataRows = GetRows();
                if (dataRows.Count == 0)
                {
                    Assert.Fail("❌ No data rows found, and no 'No data available' message. Possible UI issue.");
                }

                Assert.IsTrue(AllRowsMatchExpectedStatus(lhdnStatus),
                    $"❌ One or more rows do not have LHDN Status = '{lhdnStatus}'.");
            }
            catch (Exception ex)
            {
                // === Capture Screenshot on Error ===
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_LHDN_{lhdnStatus}_ERROR_{DateTime.Now:yyyyMMdd_HHmmssfff}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"❌ Exception during LHDN Status filter test for '{lhdnStatus}': {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }




        [Test]
        [Category("Business Entity")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Import - Upload CSV via Upload Button + Verify CSV Data")]
        public void ImportBusinessEntityCSVFile()
        {
            try
            {
                // ===== Step 1: Open Import Modal =====
                LogStep("📂 Open Import Modal");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector("#kt_content_container > app-business-entity > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a"));
                _BusinessEntityPage.ClickImportButton();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Step 2: Click Download Template =====
                LogStep("📤 Click Download button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-body.px-20 > div > div > div.d-flex.align-items-center > button"));
                _BusinessEntityPage.ClickDownloadTemplateButton();
                WaitForUIEffect();
                Thread.Sleep(1500);

                // ===== Step 3: Upload CSV =====
                string filePath = AppConfig.ImportBECSVFile;
                LogStep($"📁 Selecting file: {filePath}");
                IWebElement fileInput = _driver.FindElement(By.CssSelector("input[type='file'][accept='*']"));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                Thread.Sleep(1000);

                LogStep("📤 Click Upload button");
                helperFunction.WaitForElementToBeClickable(_wait, By.CssSelector(
                    "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2"));
                _BusinessEntityPage.ClickUploadButton();
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
                // ===== Step 4: Wait for 'Completed' =====
                LogStep("⏳ Waiting for 'Completed' button...");
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(30));
                IWebElement completedButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Completed']")));
                LogStep("✅ 'Completed' button found. Clicking...");
                completedButton.Click();
                WaitForUIEffect();
                Thread.Sleep(1000);

                // ===== Step 5: Verify Success Modal =====
                LogStep("🔍 Verifying success modal...");
                var modal = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                var message = modal.Text.Trim();
                LogStep($"📢 Modal Message: {message}");
                if (!message.ToLower().Contains("success"))
                {
                    Assert.Fail($"❌ Expected success message but got: {message}");
                }
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                okButton.Click();
                WaitForUIEffect();
                Thread.Sleep(3000);


                // ===== Step 6: Read CSV =====
                LogStep("📄 Reading business entity information from CSV...");
                var BusinessEntityName = new List<string>();
                var BETinNo = new List<string>();
                var BERegistrationType = new List<string>();
                var BERegistrationId = new List<string>();
                var BESSTNo = new List<string>();
                var BEEmail = new List<string>();
                var BEContact = new List<string>();


                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    int rowIndex = 0;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        rowIndex++;
                        if (rowIndex < 4) continue; // skip header rows

                        var columns = line.Split(',');
                        if (columns.Length >= 16)
                        {
                            var EntityName = columns[0].Trim();
                            if (string.IsNullOrEmpty(EntityName))
                            {
                                LogStep("Empty Entity Name detected, stop getting CSV data.");
                                break;
                            }

                            BusinessEntityName.Add(EntityName);
                            BETinNo.Add(columns[1].Trim());
                            BERegistrationType.Add(columns[2].Trim());
                            BERegistrationId.Add(columns[3].Trim());
                            BESSTNo.Add(columns[4].Trim());
                            BEEmail.Add(columns[8].Trim());
                            BEContact.Add(columns[7].Trim());
                        }
                    }
                }


                for (int i = 0; i < BETinNo.Count; i++)
                {
                    string EntityName = BusinessEntityName[i];
                    string expectedBETinNo = BETinNo[i];
                    string expectedBERegistrationType = BERegistrationType[i];
                    string expectedBERegistrationId = BERegistrationId[i];
                    string combinedBE = expectedBERegistrationType + ":" + expectedBERegistrationId;
                    string expectedBESSTNo = BESSTNo[i];
                    string expectedBEEmail = BEEmail[i];
                    string expectedBEContact = BEContact[i];
                    bool isMatchFound = false;

                    LogStep($"🔍 Searching Business Entity Information '{EntityName}' to verify TIN NO : '{expectedBETinNo}', Registration Type and Registration ID : '{combinedBE}', SST NO : '{expectedBESSTNo}', Email : '{expectedBEEmail}', Contact : '{expectedBEContact}'");

                    _BusinessEntityPage.SearchBusinessEntity(expectedBETinNo);
                    WaitForUIEffect(2000);
                    helperFunction.WaitForBETableToLoad(_wait);
                    WaitForUIEffect(2000);

                    while (true)
                    {
                        var rows = _driver.FindElements(By.XPath(
                            "/html/body/app-layout/div[1]/div/div/div/app-content/app-business-entity/div/div[4]/div/div[1]/div/table/tbody/tr"));

                        foreach (var row in rows)
                        {
                            var cells = row.FindElements(By.TagName("td"));
                            if (cells.Count < 8) continue;

                            string actualEntityName = cells[0].Text.Trim();
                            string actualBETinNo = cells[1].Text.Trim();
                            string actualcombinedBE = cells[2].Text.Trim();
                            string actualBESSTNo = cells[3].Text.Trim();
                            string actualBEEmail = cells[4].Text.Trim();
                            string actualBEContact = cells[5].Text.Trim();

                            if (actualBETinNo.Equals(expectedBETinNo, StringComparison.OrdinalIgnoreCase))
                            {
                                if (actualEntityName.Equals(EntityName, StringComparison.OrdinalIgnoreCase)
                                    && actualcombinedBE.Equals(combinedBE, StringComparison.OrdinalIgnoreCase)
                                    && actualBESSTNo.Equals(expectedBESSTNo, StringComparison.OrdinalIgnoreCase)
                                     && actualBEEmail.Equals(expectedBEEmail, StringComparison.OrdinalIgnoreCase)
                                    && actualBEContact.Equals(expectedBEContact, StringComparison.OrdinalIgnoreCase))
                                {
                                    isMatchFound = true;
                                    LogStep($"✅ Business Entity Information matched : '{EntityName}' TIN NO : '{actualBETinNo}', Registration Type and Registration ID : '{actualcombinedBE}', SST NO : '{actualBESSTNo}', Email : '{actualBEEmail}', Contact : '{actualBEContact}'");
                                }
                                else
                                {
                                    Assert.Fail($"❌ Business Entity Information '{EntityName}' mismatch. Expected: TIN NO : '{expectedBETinNo}', Registration Type and Registration ID : '{combinedBE}', SST NO : '{expectedBESSTNo}', Email : '{expectedBEEmail}', Contact : '{expectedBEContact}. Found: TIN NO : '{actualBETinNo}', Registration Type and Registration ID : '{actualcombinedBE}', SST NO : '{actualBESSTNo}', Email : '{actualBEEmail}', Contact : '{actualBEContact}");
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
                                helperFunction.WaitForBETableToLoad(_wait);
                            }
                            else break;
                        }
                        catch { break; }
                    }

                    if (!isMatchFound)
                    {
                        Assert.Fail($"❌ Business Entity '{EntityName}' was not found.");
                    }
                }

                LogStep("🎉 CSV Import and verification completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Business_Entity_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during Import CSV test: {ex.Message}");
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
                    string[] steps = message.Split(new[] { '\n', '•', '|' , '.'}, StringSplitOptions.RemoveEmptyEntries);
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

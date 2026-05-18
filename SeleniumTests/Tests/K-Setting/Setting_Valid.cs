using Allure.Net.Commons;
using Allure.NUnit;
using Allure.NUnit.Attributes;
using EInvoice.SeleniumTests.Config;
using EInvoice.SeleniumTests.Drivers;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using ScreenRecorderLib;
using SeleniumExtras.WaitHelpers;
using SeleniumTests.Helpers;
using SeleniumTests.Pages.Setting;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;
using OfficeOpenXml;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml.Style;
using OpenQA.Selenium.Interactions;



namespace SeleniumTests.Tests.K_Setting
{

    public static class ExcelDataReaderSettingValid
    {
        public static IEnumerable<object[]> GetCreateTestData(string filePath, string sheetName)
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
                    string cutOffDate = worksheet.Cells[row, 1].Text?.Trim();
                    string consolidateCutOffDate = worksheet.Cells[row, 2].Text?.Trim();
                    string securityToken = worksheet.Cells[row, 3].Text?.Trim();
                    string view = worksheet.Cells[row, 4].Text?.Trim();


                    yield return new object[]
                        {
                        cutOffDate, consolidateCutOffDate, securityToken, view
                        };
                    
                }
            }
        }

        public static IEnumerable<object[]> GetResetTestData(string filePath, string sheetName)
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
                    string cutOffDateReset = worksheet.Cells[row, 1].Text?.Trim();
                    string consolidateCutOffDateReset = worksheet.Cells[row, 2].Text?.Trim();
                    string securityTokenReset = worksheet.Cells[row, 3].Text?.Trim();
                    string view = worksheet.Cells[row, 4].Text?.Trim();

                    yield return new object[]
                    {
                        cutOffDateReset, consolidateCutOffDateReset, securityTokenReset, view
                    };

                }
            }

        }

            public static IEnumerable<object[]> GetSettingStoreTestData(string filePath, string sheetName)
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
                    string EntityName = worksheet.Cells[row, 1].Text?.Trim();
                    string StoreName = worksheet.Cells[row, 2].Text?.Trim();

                    yield return new object[]
                    {
                        EntityName, StoreName
                    };

                }
            }
        }

    }



    [TestFixture, Order(45)]
    [AllureNUnit]
    [AllureSuite("Setting - Setting - Valid")]
    [AllureEpic("ERP-117")]
    public class Setting_Valid
    {
        private IWebDriver _driver;
        private SettingPage _SettingPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "SettingTestDataValid.xlsx");

        // Data source for Create test
        public static IEnumerable<object[]> CreateTestData =>
            ExcelDataReaderSettingValid.GetCreateTestData(ExcelPath, "CreateNewSettingTestData");

        // Data source for Reset test
        public static IEnumerable<object[]> ResetTestData =>
            ExcelDataReaderSettingValid.GetResetTestData(ExcelPath, "ResetImageSettingTestData");

        public static IEnumerable<object[]> SettingStoreTestData =>
        ExcelDataReaderSettingValid.GetSettingStoreTestData(ExcelPath, "SettingStoreTestData");


        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Setting Page";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/setting");
            helperFunction.WaitForPageToLoad(_wait);
            _SettingPage = new SettingPage(_driver);
            _logMessages.Clear();

            _moduleName = "Setting Pages";
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
        /// Test Case: Create Setting - Positive Scenario
        /// Action:
        ///     1. Navigate to Settings page.
        ///     2. Select display density based on test data ('view').
        ///     3. Clear existing Convert and Consolidate Cut-Off Date fields.
        ///     4. Enter Convert Cut-Off Date and Consolidate Cut-Off Date from test data.
        ///     5. Set Security Token checkbox based on test data.
        ///     6. Upload sample receipt image via hidden file input and handle crop modal.
        ///     7. Click final 'Save' button to submit settings.
        /// Verification:
        ///     - Image preview should display successfully before saving.
        ///     - System should display a success message after saving.
        ///     - Screenshot is captured after submission for reporting.
        /// Purpose:
        ///     Ensure that the application allows creation of settings with valid cut-off dates, valid file upload, 
        ///     correct display density, and proper handling of the security token checkbox.
        /// Test Data:
        ///     - CutOffDate: B2C conversion cut-off date
        ///     - ConsolidateCutOffDate: Consolidate cut-off date
        ///     - securityToken: true/false to indicate if security token checkbox is checked
        ///     - view: Display density option (e.g., 'Compact', 'Comfortable')
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Setting")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(CreateTestData))]
        public void Create_New_Setting(string CutOffDate, string ConsolidateCutOffDate, string securityToken, string view)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {
                // ===== 1. Select Display Density =====
                if (!string.IsNullOrEmpty(view))
                {
                    view = view.Trim().ToLower();
                    try
                    {
                        switch (view)
                        {
                            case "compact":
                                _driver.FindElement(By.Name("displayDensityCompact")).Click();
                                LogStep("✅ Selected 'Compact' display density.");
                                break;
                            case "comfortable":
                            case "comfortable side": // fallback
                                _driver.FindElement(By.Name("displayDensitycomfortable")).Click();
                                LogStep("✅ Selected 'Comfortable' display density.");
                                break;
                            default:
                                LogStep($"⚠️ Display density '{view}' not recognized. No action taken.");
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogStep($"⚠️ Error selecting display density '{view}': {ex.Message}");
                    }
                }
                else
                {
                    LogStep("ℹ️ No display density input provided. Skipping selection.");
                }

                WaitForUIEffect();

                // ===== 2. Clear input fields =====
                var convertInput = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("input[name='consolidatePage']")));
                convertInput.Clear();
                LogStep("🧹 Cleared Consolidate Page field.");

                var consolidateInput = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("input[name='consolidateB2CPage']")));
                consolidateInput.Clear();
                LogStep("🧹 Cleared Consolidate B2C Page field.");

                // ===== 3. Enter Cut-Off Dates =====
                _SettingPage.EnterConvertCutOff(CutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Convert Cut-Off Date: {CutOffDate}");

                _SettingPage.EnterCosolidateCutOff(ConsolidateCutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Consolidate Cut-Off Date: {ConsolidateCutOffDate}");

                // ===== 4. Set Security Token Checkbox =====
                bool isSecurityTokenChecked = bool.TryParse(securityToken, out var result) && result;
                _SettingPage.SetCheckboxState(isSecurityTokenChecked);
                WaitForUIEffect(1000);
                LogStep($"Security Token Checkbox set to: {securityToken}");

                // ===== 5. File Upload via Hidden Input =====
                string filePath = AppConfig.SampleReceiptImage;
                if (!File.Exists(filePath))
                {
                    LogStep($"❌ File not found at: {filePath}");
                    Assert.Fail("File not found: " + filePath);
                }

                var fileInput = wait.Until(ExpectedConditions.ElementExists(By.CssSelector("input[type='file']")));
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "arguments[0].style.display='block'; arguments[0].style.opacity=1;", fileInput);
                fileInput.SendKeys(filePath);
                WaitForUIEffect(500);
                LogStep("📤 File uploaded via hidden input.");

                // ===== 6. Crop Modal Handling =====
                var cropSaveBtn = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-image-crop-modal//button[contains(text(),'Save')]")));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", cropSaveBtn);
                WaitForUIEffect();
                LogStep("Clicked 'Save' on crop modal.");

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector("app-image-crop-modal")));
                WaitForUIEffect();
                LogStep("Crop modal closed successfully.");

                // ===== 7. Verify Image Preview =====
                var previewImg = wait.Until(driver =>
                {
                    try
                    {
                        var img = driver.FindElement(By.XPath("//img[contains(@src,'blob:')]"));
                        return !string.IsNullOrEmpty(img.GetAttribute("src")) ? img : null;
                    }
                    catch
                    {
                        return null;
                    }
                });

                if (previewImg == null)
                {
                    LogStep("❌ Failed to display uploaded image preview.");
                    Assert.Fail("Image preview not found.");
                }

                WaitForUIEffect();
                LogStep("🖼️ Image preview displayed successfully.");

                // ===== 8. Scroll and Save Settings =====
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollBy(0, 1000);");
                WaitForUIEffect();
                LogStep("📜 Scrolled down to ensure 'Save' button is visible.");

                _SettingPage.ClickSaveButton();
                WaitForUIEffect();
                LogStep("Clicked final 'Save' button to submit settings.");

                // ===== 9. Validate Success Message =====
                var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                string message = modal.Text.Trim();
                LogStep($"📢 System displayed message: {message}");

                string messageNormalized = message.Replace("\r", " ").Replace("\n", " ").Trim().ToLower();
                if (messageNormalized.Contains("saved") || messageNormalized.Contains("success"))
                {
                    LogStep("✅ Settings saved successfully.");

                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    var okBtn = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    okBtn.Click();
                    WaitForUIEffect();
                    LogStep("✅ Acknowledged success message.");
                }
                else
                {
                    LogStep("❌ Unexpected message received after saving: " + message);
                    throw new Exception("Unexpected message: " + message);
                }
            }
            catch (Exception ex)
            {
                LogStep($"❌ An unexpected error occurred during the test: {ex.Message}");
                try
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("📸 Failure screenshot captured.");
                }
                catch (Exception innerEx)
                {
                    LogStep($"⚠️ Could not capture failure screenshot: {innerEx.Message}");
                }
                Assert.Fail("Exception occurred: " + ex.Message);
            }
        }








        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Reset Receipt Image Setting
        /// Action:
        ///     1. Navigate to Settings page.
        ///     2. Select display density based on test data ('view').
        ///     3. Clear existing Convert and Consolidate Cut-Off Date fields.
        ///     4. Enter Convert Cut-Off Date and Consolidate Cut-Off Date from test data.
        ///     5. Set Security Token checkbox based on test data.
        ///     6. Scroll to locate the existing receipt image.
        ///     7. Click 'Clear Image' button to remove the uploaded image.
        ///     8. Confirm reset by clicking 'Proceed' on the dialog.
        /// Verification:
        ///     - System should allow clearing the default image if possible, or indicate it cannot be cleared.
        ///     - System should display a success message after reset.
        ///     - Screenshot is captured after reset for reporting.
        /// Purpose:
        ///     Ensure that the application allows users to reset/clear uploaded receipt images and correctly handles
        ///     default images that cannot be cleared.
        /// Test Data:
        ///     - CutOffDateReset: B2C conversion cut-off date
        ///     - ConsolidateCutOffDateReset: Consolidate cut-off date
        ///     - securityTokenReset: true/false to indicate if security token checkbox is checked
        ///     - view: Display density option (e.g., 'Compact', 'Comfortable')
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Setting")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(ResetTestData))]
        public void Reset_Receipt_Image_Setting(string CutOffDateReset, string ConsolidateCutOffDateReset, string securityTokenReset, string view)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {

                // ===== Select Display Density Based on 'view' Variable =====
                if (!string.IsNullOrEmpty(view))
                {
                    view = view.Trim().ToLower();

                    try
                    {
                        switch (view)
                        {
                            case "compact":
                                var compactOption = _driver.FindElement(By.Name("displayDensityCompact"));
                                compactOption.Click();
                                LogStep("✅ Selected 'Compact' display density.");
                                break;

                            case "comfortable":
                            case "comfortable side": // optional fallback
                                var comfortableOption = _driver.FindElement(By.Name("displayDensitycomfortable"));
                                comfortableOption.Click();
                                LogStep("✅ Selected 'Comfortable' display density.");
                                break;

                            default:
                                LogStep($"⚠️ Display density '{view}' not recognized. No action taken.");
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogStep($"⚠️ Error selecting display density '{view}': {ex.Message}");
                    }
                }
                else
                {
                    LogStep("ℹ️ No display density input provided. Skipping selection.");
                }

                WaitForUIEffect();

                // Wait for the input to be visible
                var convertInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("input[name='consolidatePage']")));

                // Clear the input field
                convertInput.Clear();
                LogStep("🧹 Cleared Consolidate Page field.");

                // Wait for the input to be visible
                var consolidateInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("input[name='consolidateB2CPage']")));

                // Clear the input field
                consolidateInput.Clear();
                LogStep("🧹 Cleared Consolidate B2C Page field.");


                _SettingPage.EnterConvertCutOff(CutOffDateReset);
                WaitForUIEffect();
                LogStep($"Entered Convert Cut-Off Date: {CutOffDateReset}");

                _SettingPage.EnterCosolidateCutOff(ConsolidateCutOffDateReset);
                WaitForUIEffect();
                LogStep($"Entered Consolidate Cut-Off Date: {ConsolidateCutOffDateReset}");

                bool isSecurityTokenChecked = bool.TryParse(securityTokenReset, out var result) && result;
                _SettingPage.SetCheckboxState(isSecurityTokenChecked);
                WaitForUIEffect();
                LogStep($"Security Token Checkbox set to: {securityTokenReset}");

                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;
                js.ExecuteScript("window.scrollBy(0, 1000);");   // scrolls down 300px
                WaitForUIEffect();

                // 🗑️ Try to find the Clear Image button
                // Locate all "clear" buttons using the new class
                IReadOnlyCollection<IWebElement> clearBtnElements = _driver.FindElements(
                    By.CssSelector("button.close-btn.ng-star-inserted"));


                if (clearBtnElements.Count == 0)
                {
                    LogStep("ℹ️ Default image cannot be cleared — passing test.");
                    Assert.IsTrue(true,"Default image cannot be cleared.");
                    return; // stop here
                }

                // Wait until the button is visible and clickable
                var clearBtn = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("div.col-sm-12.col-lg-5 > div > button.close-btn.ng-star-inserted")));

                // Scroll into view
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", clearBtn);

                // Try normal click
                try
                {
                    clearBtn.Click();
                }
                catch
                {
                    // Fallback: click via JS if normal click fails
                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", clearBtn);
                }

                WaitForUIEffect();
                LogStep("🗑️ Clicked 'Clear Image' button.");



                // ✅ Confirm reset (Save button in dialog)
                var confirmBtn = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("button.swal2-confirm.btn.fw-bold.btn-primary")));
                confirmBtn.Click();
                WaitForUIEffect();
                LogStep("✅ Clicked 'Proceed' button on reset dialog.");

                var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                string message = modal.Text.Trim();
                LogStep($"📢 System displayed message: {message}");

                string messageNormalized = message.Replace("\r", " ").Replace("\n", " ").Trim().ToLower();
                if (messageNormalized.Contains("saved") || messageNormalized.Contains("success"))
                {
                    LogStep("✅ Settings saved successfully.");
                    Assert.IsTrue(true);

                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    var okBtn = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    okBtn.Click();
                    WaitForUIEffect();
                    LogStep("✅ Acknowledged success message.");
                }
                else
                {
                    LogStep("❌ Unexpected message received after saving: " + message);
                    throw new Exception("Unexpected message: " + message);
                }
            }
            catch (Exception ex)
            {
                LogStep($"❌ An unexpected error occurred during the test: {ex.Message}");

                try
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                    LogStep("📸 Failure screenshot captured.");
                }
                catch (Exception innerEx)
                {
                    LogStep($"⚠️ Could not capture failure screenshot: {innerEx.Message}");
                }

                Assert.Fail("Exception occurred: " + ex.Message);
            }
        }









        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Setting Store Setting
        /// Action:
        ///     1. Navigate to Settings page and open the 'Store' tab.
        ///     2. Select or unselect entities based on test data (Entities parameter).
        ///     3. Expand all store sections dynamically.
        ///     4. Select or unselect stores based on test data (Stores parameter).
        ///     5. Click 'Save' button to submit the settings.
        ///     6. Click 'Proceed' button on confirmation modal.
        /// Verification:
        ///     - Entities and stores are selected/unselected according to input data.
        ///     - System displays a success message after saving.
        ///     - Screenshot is captured after saving for reporting.
        /// Purpose:
        ///     Ensure that the application correctly handles store-specific settings for entities and stores,
        ///     and properly persists selections in the system.
        /// Test Data:
        ///     - Entities: comma-separated list of entities to select (or 'All' to select all)
        ///     - Stores: comma-separated list of stores to select (or 'All' to select all)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Setting")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(SettingStoreTestData))]
        public void Setting_StoreSetting(string Entities, string Stores)
        {
            try
            {
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(20));
                var js = (IJavaScriptExecutor)_driver;

                // ===== Open Store Setting tab =====
                var storeSettingTab = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[contains(@class, 'nav-link') and normalize-space(text())='Store']")));
                storeSettingTab.Click();
                LogStep("Opened Store Setting tab.");
                WaitForUIEffect();


                // ===== Select / Unselect Entity Checkboxes =====
                LogStep($"Processing entity selection for: '{Entities}'");

                WaitForUIEffect();

                string entityBaseXPath = "//app-store-setting//div[@class='accordion']/div";

                var allEntityItems = _driver.FindElements(By.XPath(
                    $"{entityBaseXPath}//div[.//input[@type='checkbox']]"
                )).ToList();

                if (allEntityItems.Count == 0)
                {
                    LogStep("⚠️ No entity checkboxes found.");
                }
                else
                {
                    // Deduplicate and prepare sets
                    var processed = new HashSet<string>();
                    bool selectAll = Entities.Trim().Equals("All", StringComparison.OrdinalIgnoreCase);
                    var targetEntities = selectAll
                        ? new HashSet<string>() // Not used for All
                        : Entities.Split(',')
                                  .Select(v => v.Trim().ToLower())
                                  .Where(v => !string.IsNullOrEmpty(v))
                                  .ToHashSet();

                    foreach (var entityDiv in allEntityItems)
                    {
                        try
                        {
                            // Get the label and checkbox
                            var labelText = entityDiv.Text.Trim().ToLower();
                            if (string.IsNullOrEmpty(labelText) || processed.Contains(labelText))
                                continue; // Skip duplicate or empty entries
                            processed.Add(labelText);

                            var checkbox = entityDiv.FindElement(By.XPath(".//input[@type='checkbox']"));
                            js.ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                            WaitForUIEffect();

                            if (selectAll)
                            {
                                if (!checkbox.Selected)
                                {
                                    js.ExecuteScript("arguments[0].click();", checkbox);
                                    LogStep($"✅ Selected entity '{labelText}'.");
                                }
                            }
                            else if (targetEntities.Contains(labelText))
                            {
                                // Should be selected
                                if (!checkbox.Selected)
                                {
                                    js.ExecuteScript("arguments[0].click();", checkbox);
                                    LogStep($"✅ Selected entity '{labelText}'.");
                                }
                                else
                                {
                                    LogStep($"🟢 Entity '{labelText}' already selected.");
                                }
                            }
                            else
                            {
                                // Should be unselected
                                if (checkbox.Selected)
                                {
                                    js.ExecuteScript("arguments[0].click();", checkbox);
                                    LogStep($"🟡 Unselected entity '{labelText}' (not in list).");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Error processing entity checkbox: {ex.Message}");
                        }
                    }
                }

                WaitForUIEffect();

                // ===== Expand All Store Sections (Dynamic) =====
                LogStep("Expanding all store sections dynamically...");

                string expandButtonsXPath = "//app-store-setting//div[@class='accordion']//button";

                WaitForUIEffect();

                // Find all dynamic expand buttons inside store setting
                var expandButtons = _driver.FindElements(By.XPath(expandButtonsXPath));

                if (expandButtons.Count == 0)
                {
                    LogStep("⚠️ No expandable store sections found.");
                }
                else
                {
                    foreach (var button in expandButtons)
                    {
                        try
                        {
                            js.ExecuteScript("arguments[0].scrollIntoView(true);", button);
                            WaitForUIEffect();

                            // Skip if already expanded (if aria-expanded attribute exists)
                            var expandedAttr = button.GetAttribute("aria-expanded");
                            bool isExpanded = expandedAttr != null && expandedAttr.Equals("true", StringComparison.OrdinalIgnoreCase);

                            if (!isExpanded)
                            {
                                js.ExecuteScript("arguments[0].click();", button);
                                LogStep("📂 Expanded a store section.");
                                WaitForUIEffect();
                            }
                            else
                            {
                                LogStep("📂 Store section already expanded.");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Could not expand one store section: {ex.Message}");
                        }
                    }
                }

                LogStep("✅ All store sections expanded.");
                WaitForUIEffect();


                // ===== Select / Unselect Stores (Optimized & Deduplicated) =====
                LogStep($"Processing store selection for: '{Stores}'");

                WaitForUIEffect();
                string storeBaseXPath = "//app-store-setting";

                // Find top-level store entries (each with visible span + delete icon)
                var allStoreItems = _driver.FindElements(By.XPath(
                    $"{storeBaseXPath}//div[contains(@class,'accordion-body')]//div[.//span and .//i[contains(@class,'pi-times')]]"
                )).Select(container => new
                {
                    Span = container.FindElement(By.XPath(".//span")),
                    DeleteIcon = container.FindElement(By.XPath(".//i[contains(@class,'pi-times')]"))
                }).ToList();

                if (!allStoreItems.Any())
                {
                    LogStep("⚠️ No store items found.");
                }
                else
                {
                    // Deduplicate by store name
                    var processed = new HashSet<string>();
                    var targetStores = Stores.Trim().Equals("All", StringComparison.OrdinalIgnoreCase)
                        ? new HashSet<string>() // not needed for All mode
                        : Stores.Split(',')
                                .Select(v => v.Trim().ToLower())
                                .Where(v => !string.IsNullOrEmpty(v))
                                .ToHashSet();

                    foreach (var item in allStoreItems)
                    {
                        try
                        {
                            string storeName = item.Span.Text.Trim().ToLower();

                            if (string.IsNullOrEmpty(storeName) || processed.Contains(storeName))
                                continue; // Skip empty or duplicate names

                            processed.Add(storeName);
                            js.ExecuteScript("arguments[0].scrollIntoView(true);", item.Span);
                            WaitForUIEffect();

                            if (Stores.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                            {
                                js.ExecuteScript("arguments[0].click();", item.Span);
                                LogStep($"✅ Selected store '{storeName}'.");
                            }
                            else if (targetStores.Contains(storeName))
                            {
                                js.ExecuteScript("arguments[0].click();", item.Span);
                                LogStep($"✅ Selected store '{storeName}'.");
                            }
                            else
                            {
                                js.ExecuteScript("arguments[0].click();", item.DeleteIcon);
                                LogStep($"🟡 Unselected store '{storeName}'.");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Error processing store: {ex.Message}");
                        }
                    }
                }



                // ===== Click Save =====
                var saveButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[normalize-space(text())='Save' or contains(., 'Save')]")));
                saveButton.Click();
                LogStep("Clicked Save button.");
                WaitForUIEffect();


                // ===== Click Proceed =====
                var ProceedButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[normalize-space(text())='Proceed' or contains(., 'Proceed')]")));
                ProceedButton.Click();
                LogStep("Clicked Proceed button.");

                // ===== Wait for modal message =====
                WaitForUIEffect();
                var modal = _driver.FindElement(By.XPath("//body//div[contains(@class,'modal') or contains(@class,'swal2-popup')]"));
                string message = modal.Text.Trim();
                LogStep("📢 Modal message: " + message);
                WaitForUIEffect();

                if (!message.ToLower().Contains("successfully"))
                    Assert.Fail("❌ Unexpected modal message: " + message);

                // ===== Screenshot =====
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"StoreSetting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                // ===== Dismiss modal =====
                var okButton = modal.FindElements(By.XPath(".//button[contains(., 'Ok, got it!')]")).FirstOrDefault();
                okButton?.Click();
                LogStep("✅ Clicked 'Ok, got it!'");

                WaitForUIEffect(1000);
                LogStep("✅ Store settings test completed successfully. [IMPORTANT] Please relogin on your existing browser tab.");

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"StoreSetting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
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
            Thread.Sleep(ms);
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


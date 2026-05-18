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



namespace SeleniumTests.Tests.K_Setting
{


    public static class ExcelDataReaderSettingOnDashboardValid
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
    }


    [TestFixture, Order(47)]
    [AllureNUnit]
    [AllureSuite("Setting - Setting - Valid")]
    [AllureEpic("ERP-117")]
    public class SettingOnDashboard_Valid
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


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "SettingOnDashboardTestDataValid.xlsx");

        // Data source for Create test
        public static IEnumerable<object[]> CreateTestData =>
            ExcelDataReaderSettingOnDashboardValid.GetCreateTestData(ExcelPath, "SettingTestData");

        // Data source for Reset test
        public static IEnumerable<object[]> ResetTestData =>
            ExcelDataReaderSettingOnDashboardValid.GetResetTestData(ExcelPath, "ResetTestData");


        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Setting On Dashboard Page";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/dashboard");
            helperFunction.WaitForPageToLoad(_wait);
            _SettingPage = new SettingPage(_driver);
            _logMessages.Clear();

            _moduleName = "Setting On Dashboard Page";
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
        [Category("Setting")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(CreateTestData))]
        public void Create_New_SettingOnDashboard(string CutOffDate, string ConsolidateCutOffDate, string securityToken, string view)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {
                // Navigate to Setting On Dashboard Tab
                var SettingOnDashboardTab = _driver.FindElement(
                    By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/i[1]")
                );
                SettingOnDashboardTab.Click();
                WaitForUIEffect();

                // ===== Select Display Density Based on 'view' Variable =====
                if (!string.IsNullOrEmpty(view))
                {
                    view = view.Trim().ToLower();

                    try
                    {
                        switch (view)
                        {
                            case "compact":
                                var compactOption = _driver.FindElement(By.Name("displayDensityCompactSide"));
                                compactOption.Click();
                                LogStep("✅ Selected 'Compact' display density.");
                                break;

                            case "comfortable":
                            case "comfortable side": // optional fallback
                                var comfortableOption = _driver.FindElement(By.Name("displayDensitycomfortableSide"));
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

                // 🔄 Clear fields before input
                var consolidateInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_engage_demos > div.px-7.py-6.flex-grow-1.overflow-auto > div:nth-child(2) > input")));
                consolidateInput.Clear();
                LogStep("🧹 Cleared Consolidate Cut-Off Date field.");

                var convertInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_engage_demos > div.px-7.py-6.flex-grow-1.overflow-auto > div:nth-child(3) > input")));
                convertInput.Clear();
                LogStep("🧹 Cleared Convert Cut-Off Date field.");

                _SettingPage.EnterConvertCutOffQS(CutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Convert Cut-Off Date: {CutOffDate}");

                _SettingPage.EnterCosolidateCutOffQS(ConsolidateCutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Consolidate Cut-Off Date: {ConsolidateCutOffDate}");

                bool isSecurityTokenChecked = bool.TryParse(securityToken, out var result) && result;
                _SettingPage.SetCheckboxStateQS(isSecurityTokenChecked);
                WaitForUIEffect();
                LogStep($"Security Token Checkbox set to: {securityToken}");

                string filePath = AppConfig.SampleReceiptImage;
                if (!File.Exists(filePath))
                {
                    LogStep($"❌ File not found at: {filePath}");
                    Assert.Fail("File not found: " + filePath);
                }

                // Locate the hidden <input type="file"> element (not the visible button)
                var fileInput = wait.Until(ExpectedConditions.ElementExists(
                    By.CssSelector("input[type='file']")));

                // If input is hidden (display: none), make it visible temporarily using JS
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].style.display = 'block';", fileInput);

                // Send file path directly to file input — this triggers upload without pop-up
                fileInput.SendKeys(filePath);
                WaitForUIEffect(200);
                LogStep("📤 File uploaded via hidden input bypassing file picker.");

                // Proceed with the rest (crop modal, click save)
                var cropSaveBtn = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_body > ngb-modal-window > div > div > app-image-crop-modal > div > div > div > button.btn.btn-primary")));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", cropSaveBtn);
                WaitForUIEffect();
                LogStep("Clicked 'Save' on crop modal.");

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(
                    By.CssSelector("#kt_body > ngb-modal-window > div > div > app-image-crop-modal")));
                WaitForUIEffect();
                LogStep("Crop modal closed successfully.");

                var previewImg = wait.Until(driver =>
                {
                    try
                    {
                        var img = driver.FindElement(By.CssSelector(
                            "#kt_engage_demos > div.px-7.py-6.flex-grow-1.overflow-auto > div.d-flex.flex-column.b-7 > div.row.mb-7 > img"));
                        return !string.IsNullOrEmpty(img.GetAttribute("src")) ? img : null;
                    }
                    catch { return null; }
                });

                if (previewImg == null)
                {
                    LogStep("❌ Failed to display uploaded image preview.");
                    Assert.Fail("Image preview not found.");
                }

                WaitForUIEffect();
                LogStep("🖼️ Image preview displayed successfully.");


                _SettingPage.ClickSaveButtonQS();
                WaitForUIEffect();
                LogStep("Clicked final 'Save' button to submit settings.");

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


        [Test]
        [Category("Setting")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(ResetTestData))]
        public void Reset_Receipt_Image_Setting(
    string CutOffDate,
    string ConsolidateCutOffDate,
    string securityToken,
    string view)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {
                // ===== Open Settings Panel =====
                var SettingOnDashboardTab = _driver.FindElement(
                    By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/i[1]"));
                SettingOnDashboardTab.Click();
                WaitForUIEffect();

                // ===== Select Display Density =====
                if (!string.IsNullOrEmpty(view))
                {
                    view = view.Trim().ToLower();
                    try
                    {
                        if (view == "compact")
                        {
                            _driver.FindElement(By.Name("displayDensityCompactSide")).Click();
                            LogStep("✅ Selected Compact display density.");
                        }
                        else if (view.Contains("comfortable"))
                        {
                            _driver.FindElement(By.Name("displayDensitycomfortableSide")).Click();
                            LogStep("✅ Selected Comfortable display density.");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogStep($"⚠️ Display density selection failed: {ex.Message}");
                    }
                }

                WaitForUIEffect();

                // ===== Clear & Enter Cut-Off Dates =====
                var consolidateInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_engage_demos > div.px-7.py-6.flex-grow-1.overflow-auto > div:nth-child(2) > input")));
                consolidateInput.Clear();

                var convertInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_engage_demos > div.px-7.py-6.flex-grow-1.overflow-auto > div:nth-child(3) > input")));
                convertInput.Clear();

                _SettingPage.EnterConvertCutOffQS(CutOffDate);
                WaitForUIEffect();

                _SettingPage.EnterCosolidateCutOffQS(ConsolidateCutOffDate);
                WaitForUIEffect();

                bool isSecurityTokenChecked = bool.TryParse(securityToken, out var token) && token;
                _SettingPage.SetCheckboxStateQS(isSecurityTokenChecked);
                WaitForUIEffect();

                // ===== CLEAR IMAGE LOGIC (FIXED) =====
                IWebElement clearBtn = null;

                try
                {
                    clearBtn = wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.CssSelector("button.close-btn")));

                }
                catch
                {
                    clearBtn = null;
                }

                // ===== Default Image Case =====
                if (clearBtn == null || !clearBtn.Displayed || !clearBtn.Enabled)
                {
                    LogStep("ℹ️ Default image detected — no reset required.");
                    Assert.IsTrue(true, "Default image cannot be cleared.");
                    return;
                }

                // ===== Custom Image Case =====
                LogStep("🗑️ Custom image detected — clearing image now.");
                clearBtn.Click();
                WaitForUIEffect();

                var confirmBtn = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/div/div/div[6]/button[1]")));
                confirmBtn.Click();
                WaitForUIEffect();

                // ===== Validate Success Message =====
                var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                string message = modal.Text.Trim().ToLower();

                if (!message.Contains("saved") && !message.Contains("success"))
                {
                    Assert.Fail("❌ Image reset attempted but success message not displayed.");
                }

                LogStep("✅ Image reset successfully.");

                // Screenshot after reset
                _lastScreenshotPath = Path.Combine(
                    Path.GetTempPath(),
                    $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");

                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                // Close dialog
                modal.FindElement(By.XPath(".//button[contains(., 'Ok')]")).Click();
                WaitForUIEffect();
            }
            catch (Exception ex)
            {
                LogStep($"❌ Test failed: {ex.Message}");

                try
                {
                    _lastScreenshotPath = Path.Combine(
                        Path.GetTempPath(),
                        $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");

                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                }
                catch { }

                Assert.Fail(ex.Message);
            }
        }




        [Test]
        [Category("Setting")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        public void ClickShortcutOnDashboard()
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {
                // Click the Setting On Dashboard Tab
                var settingTab = _driver.FindElement(
                    By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/i[1]")
                );
                settingTab.Click();
                WaitForUIEffect();
                LogStep("Clicked Setting On Dashboard Tab.");

                // Click the shortcut icon if present
                var shortcutIcon = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_engage_demos > div.w-100 > div.px-7.py-6.d-flex.justify-content-between.align-items-center > h1 > i")
                ));
                shortcutIcon.Click();
                WaitForUIEffect();
                LogStep("Clicked shortcut icon.");

                // Verify navigation to settings URL
                string expectedUrl = AppConfig.BaseUrl + "/setting";
                wait.Until(driver => driver.Url.Contains("/setting"));
                if (_driver.Url.Contains("/setting"))
                {
                    LogStep($"✅ Successfully navigated to settings page: {_driver.Url}");
                    Assert.IsTrue(true,"Navigation successful.");
                }
                else
                {
                    LogStep($"❌ Navigation failed. Current URL: {_driver.Url}");
                    Assert.Fail("Navigation to settings page failed.");
                }

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            }
            catch (Exception ex)
            {
                LogStep($"❌ An error occurred: {ex.Message}");
                Assert.Fail("Test failed: " + ex.Message);
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
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


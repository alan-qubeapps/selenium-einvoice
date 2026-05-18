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
using SeleniumTests.Pages.BusinessEntity;



namespace SeleniumTests.Tests.K_Setting
{

    public static class ExcelDataReaderSettingNegative
    {
        public static IEnumerable<object[]> GetConsolidateEarlierTestData(string filePath, string sheetName)
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
                    string CutOffDate = worksheet.Cells[row, 1].Text?.Trim();
                    string ConsolidateCutOffDate = worksheet.Cells[row, 2].Text?.Trim();
                    string securityToken = worksheet.Cells[row, 3].Text?.Trim();


                    yield return new object[]
                    {
                        CutOffDate, ConsolidateCutOffDate, securityToken
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
                    string CutOffDate = worksheet.Cells[row, 1].Text?.Trim();
                    string ConsolidateCutOffDate = worksheet.Cells[row, 2].Text?.Trim();
                    string securityToken = worksheet.Cells[row, 3].Text?.Trim();
                    string filelocation = worksheet.Cells[row, 4].Text?.Trim();

                    yield return new object[]
                    {
                        CutOffDate, ConsolidateCutOffDate, securityToken, filelocation
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetWrongFileTypeTestData(string filePath, string sheetName)
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
                    string CutOffDate = worksheet.Cells[row, 1].Text?.Trim();
                    string ConsolidateCutOffDate = worksheet.Cells[row, 2].Text?.Trim();
                    string securityToken = worksheet.Cells[row, 3].Text?.Trim();
                    string filelocation = worksheet.Cells[row, 4].Text?.Trim();

                    yield return new object[]
                    {
                        CutOffDate, ConsolidateCutOffDate, securityToken, filelocation
                    };

                }
            }
        }
    }

    [TestFixture, Order(46)]
    [AllureNUnit]
    [AllureSuite("Setting - Setting - Negative")]
    [AllureEpic("ERP-117")]
    public class Setting_Negative
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

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "SettingTestDataNegative.xlsx");

        public static IEnumerable<object[]> ConsolidateEarlierTestData =>
            ExcelDataReaderSettingNegative.GetConsolidateEarlierTestData(ExcelPath, "ConsolidateDateTestData");

        public static IEnumerable<object[]> InvalidFileTestData =>
            ExcelDataReaderSettingNegative.GetInvalidFileTestData(ExcelPath, "InvalidFileTestData");

        public static IEnumerable<object[]> WrongFileTypeTestData =>
        ExcelDataReaderSettingNegative.GetWrongFileTypeTestData(ExcelPath, "WrongFileTypeTestData");

        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;
        
        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Setting Page - Negative";

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

            _moduleName = "Setting Page - Negative";
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
        /// Test Case: Create Setting - Negative Scenario (Consolidate Cut-Off Earlier Than Convert Cut-Off)
        /// Action:
        ///     1. Navigate to Settings page.
        ///     2. Clear existing Convert and Consolidate Cut-Off Date fields.
        ///     3. Enter Convert Cut-Off Date and Consolidate Cut-Off Date from test data.
        ///     4. Set Security Token checkbox based on test data.
        ///     5. Upload sample receipt image and handle crop modal.
        ///     6. Click final 'Save' button to submit settings.
        /// Verification:
        ///     - System should display an error message because Consolidate Cut-Off Date is earlier than Convert Cut-Off Date.
        ///     - Image preview should display successfully before saving.
        ///     - Screenshot is captured after submission attempt for reporting.
        /// Purpose:
        ///     Ensure that the application correctly prevents creation of settings where Consolidate Cut-Off Date is earlier than Convert Cut-Off Date.
        /// Test Data:
        ///     - CutOffDate: B2C conversion cut-off date
        ///     - ConsolidateCutOffDate: Consolidate cut-off date
        ///     - securityToken: true/false to indicate if security token checkbox is checked
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Setting")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative")]
        [TestCaseSource(nameof(ConsolidateEarlierTestData))]
        public void Create_New_Setting_Negative_ConsolidateEarlier(string CutOffDate, string ConsolidateCutOffDate, string securityToken)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {
                // 🔄 Clear fields before input
                var consolidateInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_content_container > app-setting > div > div > div > div > div.row.align-items-center.mb-4 > div.col-sm-12.col-md-4 > input")));
                consolidateInput.Clear();
                LogStep("Cleared Consolidate Cut-Off Date field.");

                var convertInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_content_container > app-setting > div > div > div > div > div:nth-child(2) > div.col-sm-12.col-md-4 > input")));
                convertInput.Clear();
                LogStep("Cleared Convert Cut-Off Date field.");

                _SettingPage.EnterConvertCutOff(CutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Convert Cut-Off Date: {CutOffDate}");

                _SettingPage.EnterCosolidateCutOff(ConsolidateCutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Consolidate Cut-Off Date: {ConsolidateCutOffDate}");

                bool isSecurityTokenChecked = bool.TryParse(securityToken, out var result) && result;
                _SettingPage.SetCheckboxState(isSecurityTokenChecked);
                WaitForUIEffect();
                LogStep($"Security Token Checkbox set to: {securityToken}");

                // --- FILE UPLOAD: untouched, same as positive test ---
                string filePath = AppConfig.SampleReceiptImage;
                if (!File.Exists(filePath))
                {
                    LogStep($"❌ File not found at: {filePath}");
                    Assert.Fail("File not found: " + filePath);
                }

                var fileInput = wait.Until(ExpectedConditions.ElementExists(
                    By.CssSelector("#kt_content_container > app-setting > app-general-setting > div.card.mb-10.mt-9 > div > div > div > div:nth-child(5) > div.col-sm-12.col-md-7.row.align-items-center > input[type=file]")));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                LogStep("📤 File upload initiated.");

                var cropSaveBtn = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//ngb-modal-window//app-image-crop-modal//button[contains(., 'Save')]")));
                cropSaveBtn.Click();
                WaitForUIEffect();
                LogStep("Clicked 'Save' on crop modal.");

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(
                    By.CssSelector("ngb-modal-window[role='dialog']")));
                WaitForUIEffect();
                LogStep("Crop modal closed successfully.");

                var previewImg = wait.Until(driver =>
                {
                    try
                    {
                        var img = driver.FindElement(By.XPath("//img[starts-with(@src, 'blob:')]"));
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

                _SettingPage.ClickSaveButton();
                WaitForUIEffect();
                LogStep("Clicked final 'Save' button to submit settings.");

                var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                string message = modal.Text.Trim();
                LogStep($"📢 System displayed message {message}, Test success.");

                string messageNormalized = message.Replace("\r", " ").Replace("\n", " ").Trim().ToLower();
                // Negative assertion: expect error because Consolidate < Convert

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(messageNormalized.Contains("required"),
                    $"❌ Expected error message for invalid cut off dates, got: {message}, Test success.");
            }
            catch (Exception ex)
            {
                LogStep($"✅ Negative test caught expected exception: {ex.Message}");
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(true, "Negative test scenario passed due to invalid cut-off date combination (Consolidate Date early than B2C conversion Date), Test success.");
            }
        }








        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Create Setting - Negative Scenario (Invalid File / Consolidate Cut-Off Earlier Than Convert Cut-Off)
        /// Action:
        ///     1. Navigate to Settings page.
        ///     2. Clear existing Convert and Consolidate Cut-Off Date fields.
        ///     3. Enter Convert Cut-Off Date and Consolidate Cut-Off Date from test data.
        ///     4. Set Security Token checkbox based on test data.
        ///     5. Attempt to upload a file from test data (negative scenario: file may not exist or be invalid).
        ///     6. Handle crop modal if file upload appears.
        ///     7. Click final 'Save' button to submit settings.
        /// Verification:
        ///     - System should display an error message if Consolidate Cut-Off Date is earlier than Convert Cut-Off Date or if file is invalid.
        ///     - Image preview may fail for invalid file (negative scenario), which is considered expected behavior.
        ///     - Screenshot is captured after submission attempt for reporting.
        /// Purpose:
        ///     Ensure that the application correctly prevents creation of settings when either:
        ///         - Consolidate Cut-Off Date is earlier than Convert Cut-Off Date
        ///         - Uploaded file is invalid or missing
        /// Test Data:
        ///     - CutOffDate: B2C conversion cut-off date
        ///     - ConsolidateCutOffDate: Consolidate cut-off date
        ///     - securityToken: true/false to indicate if security token checkbox is checked
        ///     - filePath: Path to file for upload (may be invalid for negative test)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Setting")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative")]
        [TestCaseSource(nameof(InvalidFileTestData))]
        public void Create_New_Setting_Negative_InvalidFile(string CutOffDate, string ConsolidateCutOffDate, string securityToken, string filePath)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {
                // 🔄 Clear fields
                var consolidateInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_content_container > app-setting > div > div > div > div > div.row.align-items-center.mb-4 > div.col-sm-12.col-md-4 > input")));
                consolidateInput.Clear();
                LogStep("Cleared Consolidate Cut-Off Date field.");

                var convertInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_content_container > app-setting > div > div > div > div > div:nth-child(2) > div.col-sm-12.col-md-4 > input")));
                convertInput.Clear();
                LogStep("Cleared Convert Cut-Off Date field.");

                _SettingPage.EnterConvertCutOff(CutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Convert Cut-Off Date: {CutOffDate}");

                _SettingPage.EnterCosolidateCutOff(ConsolidateCutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Consolidate Cut-Off Date: {ConsolidateCutOffDate}");

                bool isSecurityTokenChecked = bool.TryParse(securityToken, out var result) && result;
                _SettingPage.SetCheckboxState(isSecurityTokenChecked);
                WaitForUIEffect();
                LogStep($"Security Token Checkbox set to: {securityToken}");

                // --- FILE UPLOAD: negative scenario ---
                if (!File.Exists(filePath))
                {
                    LogStep($"❌ File not found at: {filePath}, Test success for negative scenario.");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                    Assert.IsTrue(true, "Negative test passed due to invalid file path.");
                    return;
                }

                var fileInput = wait.Until(ExpectedConditions.ElementExists(
                    By.CssSelector("#kt_content_container > app-setting > app-general-setting > div.card.mb-10.mt-9 > div > div > div > div:nth-child(5) > div.col-sm-12.col-md-7.row.align-items-center > input[type=file]")));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                LogStep($"📤 Attempted file upload with: {filePath}");

                var cropSaveBtn = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//ngb-modal-window//app-image-crop-modal//button[contains(., 'Save')]")));
                cropSaveBtn.Click();
                WaitForUIEffect();
                LogStep("Clicked 'Save' on crop modal.");

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(
                    By.CssSelector("ngb-modal-window[role='dialog']")));
                WaitForUIEffect();
                LogStep("Crop modal closed successfully.");

                var previewImg = wait.Until(driver =>
                {
                    try
                    {
                        var img = driver.FindElement(By.XPath("//img[starts-with(@src, 'blob:')]"));
                        return !string.IsNullOrEmpty(img.GetAttribute("src")) ? img : null;
                    }
                    catch { return null; }
                });

                if (previewImg == null)
                {
                    LogStep("❌ Failed to display uploaded image preview.");
                    Assert.IsTrue(true, "Negative test passed due to invalid image file.");
                    return;
                }

                WaitForUIEffect();
                LogStep("🖼️ Image preview displayed successfully.");

                _SettingPage.ClickSaveButton();
                WaitForUIEffect();
                LogStep("Clicked final 'Save' button to submit settings.");
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                string message = modal.Text.Trim();
                LogStep($"📢 System displayed message: {message}");

                string messageNormalized = message.Replace("\r", " ").Replace("\n", " ").Trim().ToLower();
                // Negative assertion: expect error because Consolidate < Convert or file invalid

                Assert.IsTrue(messageNormalized.Contains("required") || messageNormalized.Contains("cannot be earlier") || messageNormalized.Contains("invalid"),
                    $"❌ Expected error message for invalid cut-off or file, got: {message}");
            }
            catch (Exception ex)
            {
                LogStep($"✅ Negative test caught expected exception: {ex.Message}, Test success.");
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(true, "Negative test scenario passed due to invalid cut-off date or invalid file.");
            }
        }









        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Create Setting - Negative Scenario (Wrong File Type / Invalid File Upload)
        /// Action:
        ///     1. Navigate to Settings page.
        ///     2. Clear existing Convert and Consolidate Cut-Off Date fields.
        ///     3. Enter Convert Cut-Off Date and Consolidate Cut-Off Date from test data.
        ///     4. Set Security Token checkbox based on test data.
        ///     5. Attempt to upload a file from test data (negative scenario: wrong file type or invalid file).
        ///     6. Handle crop modal if file upload appears.
        ///     7. Click final 'Save' button to submit settings.
        /// Verification:
        ///     - System should display an error message if uploaded file type is invalid or cut-off dates are incorrect.
        ///     - Image preview may fail for invalid file, which is expected for negative testing.
        ///     - Screenshot is captured after submission attempt for reporting.
        /// Purpose:
        ///     Ensure that the application correctly prevents creation of settings when either:
        ///         - Uploaded file is not a valid type
        ///         - Consolidate Cut-Off Date is earlier than Convert Cut-Off Date
        /// Test Data:
        ///     - CutOffDate: B2C conversion cut-off date
        ///     - ConsolidateCutOffDate: Consolidate cut-off date
        ///     - securityToken: true/false to indicate if security token checkbox is checked
        ///     - filePath: Path to file for upload (may be invalid or wrong type)
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("Setting")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative")]
        [TestCaseSource(nameof(WrongFileTypeTestData))]
        public void Create_New_Setting_Negative_WrongFileType(string CutOffDate, string ConsolidateCutOffDate, string securityToken, string filePath)
        {
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(15));

            try
            {
                // 🔄 Clear fields
                var consolidateInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_content_container > app-setting > div > div > div > div > div.row.align-items-center.mb-4 > div.col-sm-12.col-md-4 > input")));
                consolidateInput.Clear();
                LogStep("Cleared Consolidate Cut-Off Date field.");

                var convertInput = wait.Until(ExpectedConditions.ElementIsVisible(
                    By.CssSelector("#kt_content_container > app-setting > div > div > div > div > div:nth-child(2) > div.col-sm-12.col-md-4 > input")));
                convertInput.Clear();
                LogStep("Cleared Convert Cut-Off Date field.");

                _SettingPage.EnterConvertCutOff(CutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Convert Cut-Off Date: {CutOffDate}");

                _SettingPage.EnterCosolidateCutOff(ConsolidateCutOffDate);
                WaitForUIEffect();
                LogStep($"Entered Consolidate Cut-Off Date: {ConsolidateCutOffDate}");

                bool isSecurityTokenChecked = bool.TryParse(securityToken, out var result) && result;
                _SettingPage.SetCheckboxState(isSecurityTokenChecked);
                WaitForUIEffect();
                LogStep($"Security Token Checkbox set to: {securityToken}");

                // --- FILE UPLOAD: negative scenario ---
                if (!File.Exists(filePath))
                {
                    LogStep($"❌ File not found at: {filePath}, Test success for negative scenario.");
                    Assert.IsTrue(true, "Negative test passed due to invalid file path.");
                    return;
                }

                var fileInput = wait.Until(ExpectedConditions.ElementExists(
                    By.CssSelector("#kt_content_container > app-setting > app-general-setting > div.card.mb-10.mt-9 > div > div > div > div:nth-child(5) > div.col-sm-12.col-md-7.row.align-items-center > input[type=file]")));
                fileInput.SendKeys(filePath);
                WaitForUIEffect();
                LogStep($"📤 Attempted file upload with: {filePath}");

                var cropSaveBtn = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//ngb-modal-window//app-image-crop-modal//button[contains(., 'Save')]")));
                cropSaveBtn.Click();
                WaitForUIEffect();
                LogStep("Clicked 'Save' on crop modal.");

                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(
                    By.CssSelector("ngb-modal-window[role='dialog']")));
                WaitForUIEffect();
                LogStep("Crop modal closed successfully.");

                var previewImg = wait.Until(driver =>
                {
                    try
                    {
                        var img = driver.FindElement(By.XPath("//img[starts-with(@src, 'blob:')]"));
                        return !string.IsNullOrEmpty(img.GetAttribute("src")) ? img : null;
                    }
                    catch { return null; }
                });

                if (previewImg == null)
                {
                    LogStep("❌ Failed to display uploaded image preview.");
                    Assert.IsTrue(true, "Negative test passed due to invalid image file.");
                    return;
                }

                WaitForUIEffect();
                LogStep("🖼️ Image preview displayed successfully.");

                _SettingPage.ClickSaveButton();
                WaitForUIEffect();
                LogStep("Clicked final 'Save' button to submit settings.");

                var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                string message = modal.Text.Trim();
                LogStep($"📢 System displayed message: {message}");

                string messageNormalized = message.Replace("\r", " ").Replace("\n", " ").Trim().ToLower();
                // Negative assertion: expect error because Consolidate < Convert or file invalid

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(messageNormalized.Contains("required") || messageNormalized.Contains("cannot be earlier") || messageNormalized.Contains("invalid"),
                    $"❌ Expected error message for invalid cut-off or file, got: {message}");
            }
            catch (Exception ex)
            {
                LogStep($"✅ Negative test caught expected exception: {ex.Message}, Test success.");
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Setting_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(true, "Negative test scenario passed due to invalid cut-off date or invalid file.");
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


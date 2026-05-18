using Allure.Net.Commons;
using Allure.NUnit;
using Allure.NUnit.Attributes;
using EInvoice.SeleniumTests.Config;
using EInvoice.SeleniumTests.Drivers;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using ScreenRecorderLib;
using SeleniumExtras.WaitHelpers;
using SeleniumTests.Helpers;
using SeleniumTests.Pages.Login;
using System.Drawing;
using System.Globalization;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.L_Functional.Login
{

    public static class ExcelDataReaderLoginValid
    {
        public static IEnumerable<object[]> GetForgotPasswordTestData(string filePath, string sheetName)
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
                    string email = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        email
                    };

                }
            }
        }

    }
        
    [TestFixture, Order(52)]
    [AllureNUnit]  
    [AllureSuite("Login")] // use this ties to module
    [AllureEpic("ERP-79")] // use this and ties to ticket number
    public class LoginTest
    {
        private IWebDriver driver;
        private LoginPage _loginPage;
        private WebDriverWait wait;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "LoginTestDataValid.xlsx");

        public static IEnumerable<object[]> ForgotPasswordTestData =>
        ExcelDataReaderLoginValid.GetForgotPasswordTestData(ExcelPath, "ForgotPasswordTestData");


        private static int _fileVersion; // Shared version number for Excel & recordings
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Login Page"; // You can make this dynamic if needed

            // 🧹 Clean up old export folder
            string folderWithModule = Path.Combine(AppConfig.CsvExportFolder, today, moduleName);
            Directory.CreateDirectory(folderWithModule);

            // 🔹 Find next available version number for Excel export
            int version = 1;
            string baseFileName;
            string exportPath;

            do
            {
                baseFileName = $"TestResults_{moduleName.Replace(" ", "_")}_v{version}.xlsx";
                exportPath = Path.Combine(folderWithModule, baseFileName);
                version++;
            } while (File.Exists(exportPath));

            _fileVersion = version - 1; // save version for reuse
            _exportFilePath = exportPath;

            Console.WriteLine($"📂 Using export file: {_exportFilePath}");

            // 🧹 Delete today's video folder (if exists)
            try
            {
                string baseFolderPath = AppConfig.BaseVideoFolder;
                string todayFolderName = DateTime.Now.ToString("yyyy-MM-dd");
                string fullFolderPath = Path.Combine(baseFolderPath, todayFolderName, moduleName);

                if (Directory.Exists(fullFolderPath))
                {
                    Directory.Delete(fullFolderPath, recursive: true);
                    Console.WriteLine($"🗑️ Deleted old video folder: {fullFolderPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Failed to delete video folder: {ex.Message}");
            }

            // ✅ Continue with test setup
            driver = DriverFactory.CreateDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/auth/login");
            CaptureFooterBeforeLogin();
        }


        [SetUp]
        public void SetUp()
        {
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
            driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/auth/login");
            helperFunction.WaitForPageToLoad(wait);

            _loginPage = new LoginPage(driver, wait);
            _logMessages.Clear();
            _moduleName = "Login Page";

            // 🟢 Build file path details (recording shares same version as Excel)
            string testName = NUnit.Framework.TestContext.CurrentContext.Test.MethodName;
            string baseFolderPath = AppConfig.BaseVideoFolder;
            string todayFolderName = DateTime.Now.ToString("yyyy-MM-dd");
            string fullFolderPath = Path.Combine(baseFolderPath, todayFolderName, _moduleName);

            Directory.CreateDirectory(fullFolderPath);

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
        /// Test Case: Login Page - Valid and Invalid Login
        /// Action:
        ///     1. Navigate to the login page URL.
        ///     2. Enter the username and password from AppConfig.
        ///     3. Click the login button.
        ///     4. If `isValidLogin` is true:
        ///         - Wait for the dashboard URL to confirm successful login.
        ///         - Capture screenshot after successful login.
        ///         - Assert that the user has reached the dashboard page.
        ///         - Perform logout and wait to be redirected back to login page.
        ///     5. If `isValidLogin` is false:
        ///         - Wait to remain on the login page.
        ///         - Capture screenshot after failed login attempt.
        ///         - Assert that the URL remains on the login page.
        /// Verification:
        ///     - Successful login navigates to the dashboard URL.
        ///     - Failed login keeps the user on the login page.
        ///     - Screenshots are captured for both success and failure scenarios.
        /// Purpose:
        ///     Verify that login functionality works as expected for valid and invalid credentials.
        /// Test Data:
        ///     - username: From AppConfig.
        ///     - password: From AppConfig.
        ///     - isValidLogin: Boolean flag to determine test scenario (true = valid login, false = invalid login).
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Order(1)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Login Login 1")]
        [TestCase(true)]
        public void TestValidLogin(bool isValidLogin)
        {
            string username = AppConfig.UserName;
            string password = AppConfig.Password;

            LogStep("Navigating to login page");
            driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/auth/login");

            LogStep("Entering username and password");
            _loginPage.EnterUsername(username);
            _loginPage.EnterPassword(password);

            LogStep("Clicking login button");
            _loginPage.ClickLoginButton();

            if (isValidLogin)
            {
                LogStep("Waiting for dashboard URL to confirm successful login");
                wait.Until(ExpectedConditions.UrlContains("dashboard"));

                LogStep("✅ Login succeeded. Capturing screenshot.");
                string screenshotPath = Path.Combine(Path.GetTempPath(), $"Login_Success_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                File.WriteAllBytes(screenshotPath, screenshot.AsByteArray);
                _lastScreenshotPath = screenshotPath;

                // Assert user has reached dashboard
                Assert.IsTrue(driver.Url.Contains("dashboard"), "Login was expected to succeed, but did not reach dashboard.");

                var logoutDropdownButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/div/span")));
                logoutDropdownButton.Click();

                var logoutButton = wait.Until(ExpectedConditions.ElementToBeClickable( By.XPath("//a[contains(text(),'Sign Out')]")));
                logoutButton.Click();

                LogStep("Waiting to be redirected to login page after logout");
                wait.Until(ExpectedConditions.UrlContains("/auth/login"));
            }
            else
            {
                LogStep("Waiting to remain on login page due to invalid login");
                wait.Until(ExpectedConditions.UrlContains("/auth/login"));

                LogStep("✅ Login failed as expected. Capturing screenshot.");
                string screenshotPath = Path.Combine(Path.GetTempPath(), $"Login_Failure_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                File.WriteAllBytes(screenshotPath, screenshot.AsByteArray);
                _lastScreenshotPath = screenshotPath;

                // Assert user is still on login page
                Assert.IsTrue(driver.Url.Contains("/auth/login"), "Login was expected to fail, but URL changed unexpectedly.");
            }
        }








        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------/// 
        /// Test Case: Login Page - Forgot Password
        /// Action:
        ///     1. Click the "Forgot Password" link on the login page.
        ///     2. Enter the email address provided by the test data.
        ///     3. Click the "Submit" button to request a password reset.
        ///     4. Wait for the success modal to appear.
        ///     5. Verify that the modal displays the expected success message:
        ///            "You have reset your password successfully!"
        ///     6. Capture a screenshot of the modal.
        ///     7. Click the confirmation button on the modal.
        /// Verification:
        ///     - Success modal appears within the expected timeout (5 seconds).
        ///     - Modal contains the expected success message.
        ///     - Screenshot is captured for documentation.
        /// Purpose:
        ///     Ensure that the "Forgot Password" workflow functions correctly and provides feedback to the user.
        /// Test Data:
        ///     - email: Email address to request password reset (provided via TestCaseSource `ForgotPasswordTestData`).
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Order(2)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Login Forgot Password")]
        [TestCaseSource(nameof(ForgotPasswordTestData))]
        public void TestForgotPassword(string email)
        {

            LogStep("Clicking Forgot Password");
            _loginPage.ClickForgotPassword();
            WaitForUIEffect();

            LogStep("Entering Email Address");
            _loginPage.EnterEmail(email);
            WaitForUIEffect();

            LogStep("Clicking Submit Forgot Password Button");
            _loginPage.ClickSubmitForgotPassword();
            WaitForUIEffect();

            LogStep("Verifying Forgot Password Success Modal");

            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                var modal = wait.Until(d => d.FindElement(By.XPath("/html/body/div/div")));
                var message = driver.FindElement(By.XPath("/html/body/div/div/div[2]")).Text;
                WaitForUIEffect();

                LogStep($"Modal appeared with message: {message}");

                // ✅ Take screenshot on success
                string screenshotPath = Path.Combine(Path.GetTempPath(), $"Login_ForgotPassword_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                File.WriteAllBytes(screenshotPath, screenshot.AsByteArray);
                _lastScreenshotPath = screenshotPath;

                Assert.IsTrue(message.Contains("You have reset your password successfully!"),
                    $"Actual modal message was: '{message}'");

                // ✅ Click confirmation button
                var button = driver.FindElement(By.XPath("/html/body/div/div/div[6]/button[1]"));
                button.Click();
                LogStep("Clicked on the modal confirmation button");
            }
            catch (WebDriverTimeoutException)
            {
                // ❌ Take screenshot on failure
                string screenshotPath = Path.Combine(Path.GetTempPath(), $"Login_ForgotPassword_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                File.WriteAllBytes(screenshotPath, screenshot.AsByteArray);
                _lastScreenshotPath = screenshotPath;

                Assert.Fail("Forgot Password modal did not appear within timeout");
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
                var footerElement = wait.Until(ExpectedConditions.ElementIsVisible(
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

        private void WaitForUIEffect(int ms = 1000)
        {
            Thread.Sleep(ms);
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



        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            try
            {
                driver?.Quit();
                driver?.Dispose();
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

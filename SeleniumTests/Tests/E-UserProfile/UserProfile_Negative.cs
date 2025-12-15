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
using SeleniumTests.Pages.User;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.E_User
{

    public static class ExcelDataReaderUserProfileNegative
    {

        public static IEnumerable<object[]> GetUserProfileUpdateTestData(string filePath, string sheetName)
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
                    string Username = worksheet.Cells[row, 1].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 2].Text?.Trim();
                    string role = worksheet.Cells[row, 3].Text?.Trim();
                    string UserCurrentPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 6].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 7].Text?.Trim();

                    yield return new object[]
                    {
                        Username, CustEmail, role, UserCurrentPassword, UserPassword, UserConfirmPassword, activeUser
                    };

                }
            }
        }



        public static IEnumerable<object[]> GetUserProfilePasswordTestData(string filePath, string sheetName)
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
                    string Username = worksheet.Cells[row, 1].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 2].Text?.Trim();
                    string role = worksheet.Cells[row, 3].Text?.Trim();
                    string UserCurrentPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 6].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 7].Text?.Trim();
                    string expectedErrorMessage = worksheet.Cells[row, 8].Text?.Trim();


                    yield return new object[]
                    {
                        Username, CustEmail, role, UserCurrentPassword, UserPassword, UserConfirmPassword, activeUser, expectedErrorMessage
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUserProfileWrongPWTestData(string filePath, string sheetName)
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
                    string Username = worksheet.Cells[row, 1].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 2].Text?.Trim();
                    string role = worksheet.Cells[row, 3].Text?.Trim();
                    string UserCurrentPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 6].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 7].Text?.Trim();

                    yield return new object[]
                    {
                        Username, CustEmail, role, UserCurrentPassword, UserPassword, UserConfirmPassword, activeUser
                    };

                }
            }
        }



    }



    [TestFixture, Order(14)]
    [AllureNUnit]
    [AllureSuite("User - User Profile - Valid")]
    [AllureEpic("ERP-117")]
    public class UserProfile_Negative
    {
        private IWebDriver _driver;
        private UserPage _UserPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "UserProfileTestDataNegative.xlsx");

        public static IEnumerable<object[]> UserProfileUpdateTestData =>
        ExcelDataReaderUserProfileNegative.GetUserProfileUpdateTestData(ExcelPath, "UserProfileUpdateTestData");

        public static IEnumerable<object[]> UserProfilePasswordTestData =>
        ExcelDataReaderUserProfileNegative.GetUserProfilePasswordTestData(ExcelPath, "UserProfilePasswordTestData");

        public static IEnumerable<object[]> UserProfileWrongPWTestData =>
        ExcelDataReaderUserProfileNegative.GetUserProfileWrongPWTestData(ExcelPath, "UserProfileWrongPWTestData");


        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;


        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "User Profile Page - Negative";

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
            _UserPage = new UserPage(_driver);
            _logMessages.Clear();

            _moduleName = "User Profile Page - Negative";
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
        [Category("User")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Negative - Mandatory Fields")]
        [TestCaseSource(nameof(UserProfileUpdateTestData))]
        public void Edit_User_Profile_MandatoryFields(
            string Username, string CustEmail, string role, string UserCurrentPassword,
            string UserPassword, string UserConfirmPassword, string activeUser)
        {
            try
            {
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                var logoutDropdownButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/div/span")));
                logoutDropdownButton.Click();

                var MyProfileButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-user-inner/div[3]/a")));
                MyProfileButton.Click();

                var EditProfileButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-profile-details/div/div[2]/div/div/div[1]/div[2]/div/a")));
                EditProfileButton.Click();

                LogStep("📤 Upload profile image");
                string filePath = AppConfig.UserProfileImage;
                if (!File.Exists(filePath))
                {
                    Assert.Fail("❌ File not found: " + filePath);
                }

                WaitForUIEffect();
                var fileInput = _driver.FindElement(By.CssSelector("#kt_modal_add_user_info > div.mb-6 > div > div > label > input[type='file']:nth-child(2)"));
                fileInput.SendKeys(filePath);
                LogStep("✅ File uploaded");

                LogStep("⌨️ Enter username");
                _UserPage.EnterUsername(Username);

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);

                LogStep("Enter Current Password");
                _UserPage.EnterUserCurrentPassword(UserCurrentPassword);

                LogStep("Click show current password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[1]/div/span")).Click();

                LogStep("Enter Password");
                _UserPage.EnterUserProfilePassword(UserPassword);

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserProfileConfirmPassword(UserConfirmPassword);

                LogStep("Click show password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[1]/div/span")).Click();

                LogStep("Click show confirm password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[2]/div/span")).Click();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();

                LogStep("⏳ Checking if modal appears...");
                try
                {
                    // Wait max 3s for modal, else assume no modal (validation blocked submission)
                    var modal = new WebDriverWait(_driver, TimeSpan.FromSeconds(3))
                                    .Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    string message = modal.Text.Trim();
                    LogStep("📢 Test success, Modal message: " + message);

                    if (message.ToLower().Contains("success"))
                    {
                        Assert.Fail("❌ Unexpected SUCCESS - mandatory fields were missing but form saved");
                    }
                    else
                    {
                        LogStep("✅ Modal appeared → Correct validation handled");
                    }

                    LogStep("✅ Click 'Ok, got it!'");
                    modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]")).Click();
                }
                catch (WebDriverTimeoutException)
                {
                    // No modal appeared → PASS (form validation worked)
                    LogStep("✅ No modal appeared after Save → Mandatory field validation worked correctly");
                }

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("✅ Negative mandatory fields test completed.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }





        [Test]
        [Category("User")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Negative - Password Format")]
        [TestCaseSource(nameof(UserProfilePasswordTestData))]
        public void Edit_User_Profile_InvalidPasswordFormat(
     string Username, string CustEmail, string role, string UserCurrentPassword,
     string UserPassword, string UserConfirmPassword, string activeUser, string expectedErrorMessage)
        {
            try
            {
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                var logoutDropdownButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/div/span")));
                logoutDropdownButton.Click();

                var MyProfileButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-user-inner/div[3]/a")));
                MyProfileButton.Click();

                var EditProfileButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-profile-details/div/div[2]/div/div/div[1]/div[2]/div/a")));
                EditProfileButton.Click();

                LogStep("📤 Upload profile image");
                string filePath = AppConfig.UserProfileImage;
                if (!File.Exists(filePath))
                {
                    Assert.Fail("❌ File not found: " + filePath);
                }

                WaitForUIEffect();
                var fileInput = _driver.FindElement(By.CssSelector("#kt_modal_add_user_info > div.mb-6 > div > div > label > input[type='file']:nth-child(2)"));
                fileInput.SendKeys(filePath);
                LogStep("✅ File uploaded");

                LogStep("⌨️ Enter username");
                _UserPage.EnterUsername(Username);

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);

                LogStep("Enter Current Password");
                _UserPage.EnterUserCurrentPassword(UserCurrentPassword);

                LogStep("Click show current password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[1]/div/span")).Click();

                LogStep("Enter Password");
                _UserPage.EnterUserProfilePassword(UserPassword);

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserProfileConfirmPassword(UserConfirmPassword);

                LogStep("Click show password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[1]/div/span")).Click();

                LogStep("Click show confirm password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[2]/div/span")).Click();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();

                LogStep("⏳ Checking if modal appears...");
                string message = null;
                try
                {
                    // wait max 3s for modal, else assume no modal (validation blocked submission)
                    var modal = new WebDriverWait(_driver, TimeSpan.FromSeconds(3))
                                    .Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    message = modal.Text.Trim();
                    LogStep("📢Test success, Modal message: " + message);

                    if (message.ToLower().Contains(expectedErrorMessage.ToLower()))
                    {
                        LogStep("✅ Correct validation message displayed: " + message);
                    }
                    else if (message.ToLower().Contains("success"))
                    {
                        Assert.Fail("❌ Unexpected SUCCESS - weak password accepted");
                    }
                    else
                    {
                        Assert.Fail("❌ Unexpected modal message. Expected: " + expectedErrorMessage + " | Actual: " + message);
                    }

                    LogStep("✅ Click 'Ok, got it!'");
                    modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]")).Click();
                }
                catch (WebDriverTimeoutException)
                {
                    // No modal appeared → PASS (form validation blocked submit)
                    LogStep("✅ No modal appeared after Save → Weak password correctly rejected by form validation");
                }

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("✅ Negative password format test completed successfully");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }





        [Test]
        [Category("User")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(UserProfileWrongPWTestData))]
        public void Edit_User_Profile_WrongCurrentPassword(string Username, string CustEmail, string role, string UserCurrentPassword, string UserPassword, string UserConfirmPassword, string activeUser)
        {
            try
            {
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                var logoutDropdownButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/div/span")));
                logoutDropdownButton.Click();

                var MyProfileButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-user-inner/div[3]/a")));
                MyProfileButton.Click();

                var EditProfileButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-profile-details/div/div[2]/div/div/div[1]/div[2]/div/a")));
                EditProfileButton.Click();

                LogStep("📤 Upload profile image");
                string filePath = AppConfig.UserProfileImage;
                if (!File.Exists(filePath))
                {
                    Assert.Fail("❌ File not found: " + filePath);
                }

                WaitForUIEffect();
                var fileInput = _driver.FindElement(By.CssSelector("#kt_modal_add_user_info > div.mb-6 > div > div > label > input[type='file']:nth-child(2)"));
                fileInput.SendKeys(filePath);
                LogStep("✅ File uploaded");

                LogStep("⌨️ Enter username");
                _UserPage.EnterUsername(Username);

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);

                LogStep("Enter Current Password");
                _UserPage.EnterUserCurrentPassword(UserCurrentPassword);

                LogStep("Click show current password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[1]/div/span")).Click();

                LogStep("Enter Password");
                _UserPage.EnterUserProfilePassword(UserPassword);

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserProfileConfirmPassword(UserConfirmPassword);

                LogStep("Click show password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[1]/div/span")).Click();

                LogStep("Click show confirm password icon");
                WaitForUIEffect();
                _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[2]/div/span")).Click();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();

                LogStep("Wait for modal message");
                WaitForUIEffect();
                var modal = _driver.FindElement(By.XPath("/html/body/div/div/div[2]"));
                string message = modal.Text.Trim();
                LogStep("📢 Modal message: " + message);

                // ✅ Test should PASS only if message = "old password is incorrect"
                if (message.ToLower().Contains("old password is incorrect"))
                {
                    LogStep("✅ Test success Valid negative outcome: " + message);
                }
                else if (message.ToLower().Contains("success"))
                {
                    Assert.Fail("❌ Unexpected SUCCESS - Test should fail when password is accepted");
                }
                else
                {
                    Assert.Fail("❌ Unexpected modal message: " + message);
                }

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("✅ Click 'Ok, got it!'");
                var okButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("button.swal2-confirm")
                ));
                okButton.Click();


                LogStep("✅ Edit user profile negative test completed successfully" + message);
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }





        [Test]
        [Category("User")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Negative - Access Dashboard After Sign Out")]
        public void User_Profile_SignOutSessions_Negative()
        {
            try
            {
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                // Open logout dropdown
                var logoutDropdownButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/app-header/div/app-topbar/div/span")));
                logoutDropdownButton.Click();
                LogStep("Clicked logout dropdown.");

                // Go to My Profile
                var myProfileButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-user-inner/div[3]/a")));
                myProfileButton.Click();
                LogStep("Clicked My Profile button.");

                // Open Event and Log Tab
                var eventAndLogTab = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-profile-details/div/div[1]/div[2]/ul/li[2]/a")));
                eventAndLogTab.Click();
                LogStep("Opened Event and Log tab.");

                // Click Sign Out Session button
                LogStep("💾 Clicking Sign Out Session button.");
                var signOutSessionButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-profile-details/div/div[2]/div/div/div[1]/div[2]/a")));
                signOutSessionButton.Click();
                WaitForUIEffect();

                // Verify redirection to login page
                wait.Until(driver => driver.Url.Contains("/auth/login"));
                Assert.IsTrue(_driver.Url.Contains("/auth/login"), "User should be redirected to login after sign out.");
                LogStep($"✅ Successfully navigated to login page: {_driver.Url}");

                // ===== Negative Test: Try to access dashboard after logout =====
                string dashboardUrl = AppConfig.BaseUrl + "/dashboard";
                _driver.Navigate().GoToUrl(dashboardUrl);
                WaitForUIEffect();

                if (_driver.Url.Contains("/auth/login"))
                {
                    LogStep("✅ Access to dashboard is blocked after sign out. Redirected back to login page. Test success.");
                    Assert.IsTrue(true, "Negative test success: Cannot access dashboard after sign out.");
                }
                else
                {
                    LogStep($"❌ Dashboard is still accessible after sign out. Current URL: {_driver.Url}");
                    Assert.Fail("Negative test failed: User should not access dashboard after sign out.");
                }

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            }
            catch (Exception ex)
            {
                LogStep($"❌ An error occurred: {ex.Message}");
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.Fail("Test failed: " + ex.Message);
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



        private void WaitForUIEffect(int ms = 1500)
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

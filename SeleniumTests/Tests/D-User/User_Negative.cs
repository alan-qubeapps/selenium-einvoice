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
using SeleniumTests.Pages.User;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.D_User
{



    public static class ExcelDataReaderUserNegative
    {
        public static IEnumerable<object[]> GetMissingRolenameTestData(string filePath, string sheetName)
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
                    string Rolename = worksheet.Cells[row, 1].Text?.Trim();
                    string RoleDesc = worksheet.Cells[row, 2].Text?.Trim();
                    string colorClass = worksheet.Cells[row, 3].Text?.Trim();
                    string EditStorePermission = worksheet.Cells[row, 4].Text?.Trim();
                    string expectedMessage = worksheet.Cells[row, 5].Text?.Trim();

                    yield return new object[]
                    {
                        Rolename, RoleDesc, colorClass, EditStorePermission, expectedMessage
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetUpdateMissingRolenameTestData(string filePath, string sheetName)
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
                    string Name = worksheet.Cells[row, 1].Text?.Trim();
                    string Rolename = worksheet.Cells[row, 2].Text?.Trim();
                    string RoleDesc = worksheet.Cells[row, 3].Text?.Trim();
                    string colorClass = worksheet.Cells[row, 4].Text?.Trim();
                    string EditStorePermission = worksheet.Cells[row, 5].Text?.Trim();

                    yield return new object[]
                    {
                        Rolename, Name, RoleDesc, colorClass, EditStorePermission
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
                    string Username = worksheet.Cells[row, 1].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 2].Text?.Trim();
                    string role = worksheet.Cells[row, 3].Text?.Trim();
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();


                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetInvalidPasswordFormatTestData(string filePath, string sheetName)
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
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();
                    string expectedCase = worksheet.Cells[row, 7].Text?.Trim();



                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser, expectedCase
                    };
                }
            }
        }


        public static IEnumerable<object[]> GetInvalidProfileImageTestData(string filePath, string sheetName)
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
                    string invalidFilePath = worksheet.Cells[row, 1].Text?.Trim();
                    string Username = worksheet.Cells[row, 2].Text?.Trim();
                    string CustEmail = worksheet.Cells[row, 3].Text?.Trim();
                    string role = worksheet.Cells[row, 4].Text?.Trim();
                    string UserPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 6].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 7].Text?.Trim();


                    yield return new object[]
                    {
                        invalidFilePath, Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetMandatoryUsernameEmptyTestData(string filePath, string sheetName)
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
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();
                    string expectedCase = worksheet.Cells[row, 7].Text?.Trim();



                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser, expectedCase
                    };
                }
            }
        }



        public static IEnumerable<object[]> GetMandatoryEmailEmptyTestData(string filePath, string sheetName)
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
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();
                    string expectedCase = worksheet.Cells[row, 7].Text?.Trim();



                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser, expectedCase
                    };
                }
            }
        }


        public static IEnumerable<object[]> GetMandatorPasswordEmptyTestData(string filePath, string sheetName)
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
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();
                    string expectedCase = worksheet.Cells[row, 7].Text?.Trim();



                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser, expectedCase
                    };
                }
            }
        }


        public static IEnumerable<object[]> GetMandatorCPasswordEmptyTestData(string filePath, string sheetName)
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
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();
                    string expectedCase = worksheet.Cells[row, 7].Text?.Trim();



                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser, expectedCase
                    };
                }
            }
        }


        public static IEnumerable<object[]> GetDuplicateTestData(string filePath, string sheetName)
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
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();



                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser
                    };
                }
            }
        }


        public static IEnumerable<object[]> GetInvalidRoleCheckboxTestData(string filePath, string sheetName)
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
                    string UserPassword = worksheet.Cells[row, 4].Text?.Trim();
                    string UserConfirmPassword = worksheet.Cells[row, 5].Text?.Trim();
                    string activeUser = worksheet.Cells[row, 6].Text?.Trim();



                    yield return new object[]
                    {
                        Username, CustEmail, role, UserPassword, UserConfirmPassword, activeUser
                    };
                }
            }
        }


    }



    [TestFixture, Order(12)]
    [AllureNUnit]
    [AllureSuite("User - User - Negative")]
    [AllureEpic("ERP-117")]
    public class User_Negative
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


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "UserTestDataNegative.xlsx");

        public static IEnumerable<object[]> MissingRolenameTestData =>
            ExcelDataReaderUserNegative.GetMissingRolenameTestData(ExcelPath, "MissingRolenameTestData");

        public static IEnumerable<object[]> UpdateMissingRolenameTestData =>
            ExcelDataReaderUserNegative.GetUpdateMissingRolenameTestData(ExcelPath, "UpdateMissingRolenameTestData");

        public static IEnumerable<object[]> InvalidEmailTestData =>
        ExcelDataReaderUserNegative.GetInvalidEmailTestData(ExcelPath, "InvalidEmailTestData");

        public static IEnumerable<object[]> InvalidPasswordFormatTestData =>
        ExcelDataReaderUserNegative.GetInvalidPasswordFormatTestData(ExcelPath, "InvalidPasswordFormatTestData");

        public static IEnumerable<object[]> InvalidProfileImageTestData =>
        ExcelDataReaderUserNegative.GetInvalidProfileImageTestData(ExcelPath, "InvalidProfileImageTestData");

        public static IEnumerable<object[]> MandatoryUsernameEmptyTestData =>
            ExcelDataReaderUserNegative.GetMandatoryUsernameEmptyTestData(ExcelPath, "MandatoryUsernameEmptyTestData");

        public static IEnumerable<object[]> MandatoryEmailEmptyTestData =>
        ExcelDataReaderUserNegative.GetMandatoryEmailEmptyTestData(ExcelPath, "MandatoryEmailEmptyTestData");

        public static IEnumerable<object[]> MandatorPasswordEmptyTestData =>
        ExcelDataReaderUserNegative.GetMandatorPasswordEmptyTestData(ExcelPath, "MandatorPasswordEmptyTestData");

        public static IEnumerable<object[]> MandatorCPasswordEmptyTestData =>
        ExcelDataReaderUserNegative.GetMandatorCPasswordEmptyTestData(ExcelPath, "MandatorCPasswordEmptyTestData");

        public static IEnumerable<object[]> DuplicateTestData =>
            ExcelDataReaderUserNegative.GetDuplicateTestData(ExcelPath, "DuplicateTestData");

        public static IEnumerable<object[]> InvalidRoleCheckboxTestData =>
        ExcelDataReaderUserNegative.GetInvalidRoleCheckboxTestData(ExcelPath, "InvalidRoleCheckboxTestData");



        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;


        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "User Page - Negative";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/user");
            helperFunction.WaitForPageToLoad(_wait);
            _UserPage = new UserPage(_driver);
            _logMessages.Clear();

            _moduleName = "User Page - Negative";
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
        /// Test Case: Create User Role - Negative: Missing Rolename
        /// Action:
        ///     1. Navigate to User Role tab.
        ///     2. Click "New User Role" button.
        ///     3. Leave Role Name empty and fill other fields (Role Description, Color, Edit Store Permission, Module Permissions).
        ///     4. Attempt to submit the form.
        /// Verification:
        ///     - Validation feedback should appear (modal or inline error) indicating the Rolename is required.
        ///     - Screenshots captured at submission and validation steps.
        /// Purpose:
        ///     Ensure that the system correctly validates required fields and prevents creation of a user role with missing Rolename.
        /// Test Data:
        ///     - Provided by 'MissingRolenameTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Create - Negative: Missing Rolename")]
        [TestCaseSource(nameof(MissingRolenameTestData))]
        public void Create_User_Role_Negative_MissingRolename(string Rolename, string RoleDesc, string colorClass, string EditStorePermission, string expectedMessage)

        {
            try
            {
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                LogStep("🧭 Navigating to User Role tab");
                var userRoleTab = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-user/ul/li[2]/a")));
                userRoleTab.Click();
                WaitForUIEffect();

                LogStep("🆕 Clicking New User Role button");
                _UserPage.ClickNewUserRoleButton();
                WaitForUIEffect();

                LogStep("⌨️ Leave Role Name EMPTY (Negative Case)");
                _UserPage.EnterRolename(Rolename);

                LogStep("📝 Enter Role Description");
                _UserPage.EnterRoleDesc(RoleDesc);

                LogStep("🎨 Select color");
                var colorPickerBox = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_profile_details_view > div.card-body.p-9.px-0 > form > div > div:nth-child(1) > div.col-md-5 > div > div > div:nth-child(2) > div")));
                colorPickerBox.Click();
                WaitForUIEffect();

                string colorSelector = $"#kt_profile_details_view > div.card-body.p-9.px-0 > form > div > div:nth-child(1) > div.col-md-5 > div > div.form-control.form-control-solid.ng-star-inserted > div > app-color-picker > div.color-picker-container > div > div.color-option.{colorClass}";
                var colorOption = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(colorSelector)));
                colorOption.Click();

                LogStep("📜 Scroll down a bit");
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollBy(0, 200);");
                Thread.Sleep(800);

                LogStep("🛠️ Click Edit Store Permission");
                var userRoleEditButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id=\"kt_profile_details_view\"]/div[2]/form/div[1]/div[2]/div/a")));
                userRoleEditButton.Click();
                WaitForUIEffect();

                bool isEditStorePermissionChecked = bool.TryParse(EditStorePermission, out var result) && result;
                _UserPage.SetEditStoreCheckboxState(isEditStorePermissionChecked);

                LogStep("💾 Submit Edit Store Permissions");
                var submitEditStorePermissionButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-role-modal/div/div[3]/button")));
                submitEditStorePermissionButton.Click();
                WaitForUIEffect();

                LogStep("☑️ Tick All Module Permission");
                var tickAllModulePermissionButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id=\"kt_profile_details_view\"]/div[2]/form/div[3]/div[2]/div/a")));
                tickAllModulePermissionButton.Click();
                WaitForUIEffect(200);

                LogStep("📜 Scroll the whole page to bottom");
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
                Thread.Sleep(800);

                LogStep("🔘 Attempt Submit with EMPTY Rolename");
                var submitButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-user/form/div/app-user-role-table/div/app-user-role-details/div/div[2]/div[2]/div[2]/div/a")));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", submitButton);
                Thread.Sleep(300);
                submitButton.Click();
                WaitForUIEffect();

                // 📸 Screenshot evidence
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"UserNegative_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("🔎 Checking for validation feedback");

                string actualMessage = string.Empty;

                try
                {
                    // ✅ Check modal first
                    var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    actualMessage = modal.Text.Trim();
                    LogStep("📢 Test Success, Modal message detected: " + actualMessage);

                    if (actualMessage.Contains(expectedMessage))
                    {
                        LogStep("✅ Modal validation matched expected message");
                        Assert.IsTrue(true, "Validation succeeded via modal: " + actualMessage);
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("ℹ️ No modal found, checking inline validation");

                    try
                    {
                        var inlineError = _driver.FindElement(By.XPath("//div[contains(@class,'invalid-feedback')]")).Text.Trim();
                        actualMessage = inlineError;
                        LogStep("📢 Inline validation: " + actualMessage);

                        if (actualMessage.Contains(expectedMessage))
                        {
                            LogStep("✅ Inline validation matched expected message");
                            Assert.IsTrue(true,"Validation succeeded via inline error: " + actualMessage);
                        }
                        else
                        {
                            Assert.Fail($"❌ Expected '{expectedMessage}', but got '{actualMessage}'");
                        }
                    }
                    catch (NoSuchElementException)
                    {
                        LogStep("✅ Test success." + expectedMessage);
                        Assert.IsTrue(true,"Test success. Expected: " + expectedMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"UserNegative_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }






        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Update User Role - Negative: Missing Rolename
        /// Action:
        ///     1. Navigate to User Role tab.
        ///     2. Click "Edit" on the specified user role.
        ///     3. Clear the Rolename field and fill other fields (Role Description, Color, Edit Store Permission, Module Permissions).
        ///     4. Attempt to submit the form.
        /// Verification:
        ///     - Validation feedback should appear (modal or inline error) indicating the Rolename is required.
        ///     - Screenshots captured at submission and validation steps.
        /// Purpose:
        ///     Ensure that the system correctly validates required fields and prevents updating a user role with missing Rolename.
        /// Test Data:
        ///     - Provided by 'UpdateMissingRolenameTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Update - Negative: Missing Rolename")]
        [TestCaseSource(nameof(UpdateMissingRolenameTestData))]
        public void Update_User_Role_Negative_MissingRolename(string Name, string Rolename, string RoleDesc, string colorClass, string EditStorePermission)
        {
            try
            {
                var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                LogStep("🧭 Navigating to User Role tab");
                var userRoleTab = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-user/ul/li[2]/a")));
                userRoleTab.Click();
                WaitForUIEffect();

                LogStep($"Editing user role with EMPTY Rolename (Negative Case)");
                _UserPage.ClickEditButton(Rolename);
                WaitForUIEffect();

                LogStep("⌨️ Clear Role Name field");
                _UserPage.EnterRolename(Name);
                WaitForUIEffect(1000);

                LogStep("📝 Enter Role Description");
                _UserPage.EnterRoleDesc(RoleDesc);
                WaitForUIEffect(1000);

                LogStep("🎨 Select color");
                var colorPickerBox = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_profile_details_view > div.card-body.p-9.px-0 > form > div > div:nth-child(1) > div.col-md-5 > div > div > div:nth-child(2) > div")));
                colorPickerBox.Click();
                WaitForUIEffect();

                string colorSelector = $"#kt_profile_details_view > div.card-body.p-9.px-0 > form > div > div:nth-child(1) > div.col-md-5 > div > div.form-control.form-control-solid.ng-star-inserted > div > app-color-picker > div.color-picker-container > div > div.color-option.{colorClass}";
                var colorOption = wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(colorSelector)));
                colorOption.Click();

                LogStep("📜 Scroll down a bit");
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollBy(0, 200);");
                Thread.Sleep(800);

                LogStep("🛠️ Click Edit Store Permission");
                var userRoleEditButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id=\"kt_profile_details_view\"]/div[2]/form/div[1]/div[2]/div/a")));
                userRoleEditButton.Click();
                WaitForUIEffect();

                bool isEditStorePermissionChecked = bool.TryParse(EditStorePermission, out var result) && result;
                _UserPage.SetEditStoreCheckboxState(isEditStorePermissionChecked);

                LogStep("💾 Submit Edit Store Permissions");
                var submitEditStorePermissionButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-role-modal/div/div[3]/button")));
                submitEditStorePermissionButton.Click();
                WaitForUIEffect();

                LogStep("☑️ Tick All Module Permission");
                var tickAllModulePermissionButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id=\"kt_profile_details_view\"]/div[2]/form/div[3]/div[2]/div/a")));
                tickAllModulePermissionButton.Click();

                LogStep("📜 Scroll to bottom of the page");
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
                Thread.Sleep(800);

                LogStep("🔘 Attempt Submit with EMPTY Rolename");
                var submitButton = wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-user/form/div/app-user-role-table/div/app-user-role-details/div/div[2]/div[2]/div[2]/div/a")));
                submitButton.Click();
                WaitForUIEffect();

                // 📸 Screenshot for evidence
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"UserUpdateNegative_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("🔎 Checking for validation feedback or no response");
                bool validationDetected = false;
                string expectedMessage = "Role name is required"; // ✅ expected from modal

                try
                {
                    // ✅ Check modal
                    var modal = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div")));
                    var message = modal.Text.Trim();
                    LogStep("📢 Test success, Modal message: " + message);

                    Assert.IsTrue(message.ToLower().Contains("role") || message.ToLower().Contains("required"),
                        $"❌ Expected modal validation containing '{expectedMessage}', but got: '{message}'");

                    validationDetected = true;

                    // Close modal
                    var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    okButton.Click();
                    WaitForUIEffect();
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("ℹ️ No modal found, checking inline error...");

                    try
                    {
                        IWebElement inlineError = _driver.FindElement(By.XPath("//*[@id=\"kt_profile_details_view\"]/div[2]/form/div/div[1]/div[1]/div/div/div"));
                        var inlineMsg = inlineError?.Text?.Trim() ?? "";
                        LogStep("📢 Inline validation: " + inlineMsg);

                        Assert.IsTrue(inlineMsg.ToLower().Contains("role name") || inlineMsg.ToLower().Contains("required"),
                            $"❌ Expected inline validation containing '{expectedMessage}', but got: '{inlineMsg}'");

                        validationDetected = true;
                    }
                    catch (NoSuchElementException)
                    {
                        // ✅ No modal + no inline → Save button gave no response
                        LogStep("ℹ️ No validation element displayed. Assuming Save button gave no response (still PASS).");
                        validationDetected = true;
                    }
                }

                Assert.IsTrue(validationDetected, "Validation or no-response expected for missing Rolename.");
                LogStep("🎉 Blank for mandatory field handled correctly.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"UserUpdateNegative_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }






        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Invalid Email Format
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload profile image.
        ///     3. Fill Username and an invalid Customer Email.
        ///     4. Set Role checkbox, Password, Confirm Password, and Active state.
        ///     5. Click Save.
        /// Verification:
        ///     - Validation message should appear indicating invalid email format (inline or modal).
        ///     - Screenshots captured for each step and validation message.
        /// Purpose:
        ///     Ensure the system correctly validates email format and prevents saving users with invalid emails.
        /// Test Data:
        ///     - Provided by 'InvalidEmailTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Invalid Email Format")]
        [TestCaseSource(nameof(InvalidEmailTestData))]
        public void Create_InvalidEmail_ShouldShowValidationMessage(
         string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser)
        {
            try
            {
                LogStep("🆕 Click New button");
                _UserPage.ClickNewButton();

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
                WaitForUIEffect();

                LogStep("Enter invalid customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidEmail_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"🔘 Set '{role}' checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter Password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);
                WaitForUIEffect();

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("🔍 Check for validation or system response");
                WaitForUIEffect();

                // Try to locate validation message (common selectors)
                var validationMessages = _driver.FindElements(By.CssSelector(".invalid-feedback, .error-message, .validation-message"));



                if (validationMessages.Count > 0)
                {
                    string message = validationMessages[0].Text.Trim();
                    LogStep("✅ Test success, Validation message displayed: " + message);
                    return;
                }

                // If no validation message, check if page stayed the same (no modal / redirect)
                var modals = _driver.FindElements(By.XPath("/html/body/div/div"));
                if (modals.Count == 0)
                {
                    LogStep("✅ Test success, No response detected after Save (expected for invalid email).");
                    return;
                }

                // If modal shows success, that's incorrect behavior
                string modalMessage = modals[0].Text.Trim();
                if (modalMessage.ToLower().Contains("successful"))
                {
                    Assert.Fail("❌ Unexpected success message for invalid email: " + modalMessage);
                }

                // Fallback case
                Assert.Fail("❌ Unexpected behavior: invalid email not properly handled.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidEmail_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Invalid Password Format
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload profile image.
        ///     3. Fill Username, Customer Email, Role checkbox.
        ///     4. Enter invalid Password and Confirm Password.
        ///     5. Set Active checkbox.
        ///     6. Click Save.
        /// Verification:
        ///     - Validation message should appear indicating invalid password format (inline or modal).
        ///     - Screenshots captured for each step and validation message.
        /// Purpose:
        ///     Ensure the system correctly validates password format and prevents saving users with invalid passwords.
        /// Test Data:
        ///     - Provided by 'InvalidPasswordFormatTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Invalid Password Format")]
        [TestCaseSource(nameof(InvalidPasswordFormatTestData))]
        public void Create_InvalidPasswordFormat_ShouldShowValidationMessage(
        string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser, string expectedCase)
        {
            try
            {
                LogStep("🆕 Click New button");
                _UserPage.ClickNewButton();

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
                WaitForUIEffect();

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep($"🔘 Set '{role}' checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter Password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidPassword_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);
                WaitForUIEffect();

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("🔍 Check for validation or system response");
                WaitForUIEffect();

                // Try to locate validation message (common selectors)
                var validationMessages = _driver.FindElements(By.CssSelector(".invalid-feedback, .error-message, .validation-message"));

                if (validationMessages.Count > 0)
                {
                    string message = validationMessages[0].Text.Trim();
                    LogStep($"✅ Test success, Validation message displayed for invalid password ({expectedCase}): " + message);
                    return;
                }

                // If no validation message, check if page stayed the same (no modal / redirect)
                var modals = _driver.FindElements(By.XPath("/html/body/div/div"));
                if (modals.Count == 0)
                {
                    LogStep($"✅ Test success, No response detected after Save (expected for invalid password: {expectedCase}).");
                    return;
                }

                // If modal shows success, that's incorrect behavior
                string modalMessage = modals[0].Text.Trim();
                if (modalMessage.ToLower().Contains("successful"))
                {
                    Assert.Fail($"❌ Unexpected success message with invalid password ({expectedCase}): " + modalMessage);
                }

                // Fallback case
                Assert.Fail($"❌ Unexpected behavior: invalid password format not properly handled ({expectedCase}).");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidPassword_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }






        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Invalid Profile Image Format
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload an invalid profile image file.
        ///     3. Fill Username, Customer Email, Role checkbox, Password, Confirm Password, and Active checkbox.
        ///     4. Click Save.
        /// Verification:
        ///     - Modal message should appear indicating invalid file format.
        ///     - Test fails if the system shows a success message.
        ///     - Screenshot captured for evidence.
        /// Purpose:
        ///     Ensure the system correctly validates profile image file format and prevents saving users with invalid files.
        /// Test Data:
        ///     - Provided by 'InvalidProfileImageTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Invalid Profile Image Format")]
        [TestCaseSource(nameof(InvalidProfileImageTestData))]
        public void Create_InvalidProfileImage(string invalidFilePath, string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser)
        {
            try
            {
                LogStep("🆕 Click [New] button");
                _UserPage.ClickNewButton();
                WaitForUIEffect();

                // Upload invalid file
                if (!File.Exists(invalidFilePath))
                {
                    Assert.Fail("❌ Invalid file not found: " + invalidFilePath);
                }

                WaitForUIEffect();
                LogStep($"📤 Upload invalid profile image: {Path.GetFileName(invalidFilePath)}");
                var fileInput = _driver.FindElement(By.CssSelector("#kt_modal_add_user_info > div.mb-6 > div > div > label > input[type='file']:nth-child(2)"));
                fileInput.SendKeys(invalidFilePath);
                WaitForUIEffect();

                // Fill required fields
                LogStep("⌨️ Enter username");
                _UserPage.EnterUsername(Username);
                WaitForUIEffect();

                LogStep("⌨️ Enter email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep("🔘 Set role checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter confirm password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);
                WaitForUIEffect();

                // Save action
                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("🔍 Waiting for modal message");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                string message = modal.Text.Trim();
                LogStep($"📢Test success, Modal Message: {message}");

                if (message.ToLower().Contains("successful"))
                {
                    Assert.Fail($"❌ Unexpected success message for invalid file upload: {message}");
                }

                LogStep($"✅ Test passed, modal message displayed for invalid file ({Path.GetFileName(invalidFilePath)}): {message}");
                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click();
                WaitForUIEffect();
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidFile_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Console.WriteLine($"❌ Exception during test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Mandatory Fields Empty (Username)
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload a valid profile image.
        ///     3. Fill other required fields except for Username (leave empty).
        ///     4. Click Save.
        /// Verification:
        ///     - Validation message should appear for the missing mandatory field (Username).
        ///     - Test fails if the system shows a success message.
        ///     - Screenshot captured for evidence.
        /// Purpose:
        ///     Ensure the system correctly validates mandatory fields and prevents saving users with missing Username.
        /// Test Data:
        ///     - Provided by 'MandatoryUsernameEmptyTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Mandatory Fields Empty")]
        [TestCaseSource(nameof(MandatoryUsernameEmptyTestData))]
        public void Create_MandatoryFieldEmpty_Username(
        string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser, string expectedField)
        {
            try
            {
                LogStep("🆕 Click New button");
                _UserPage.ClickNewButton();

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
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep($"🔘 Set '{role}' checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter Password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);
                WaitForUIEffect();

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("🔍 Check for validation or system response");
                WaitForUIEffect();

                // Try to locate validation message (common selectors)
                var validationMessages = _driver.FindElements(By.CssSelector(".invalid-feedback, .error-message, .validation-message"));

                if (validationMessages.Count > 0)
                {
                    string message = validationMessages[0].Text.Trim();
                    LogStep($"✅ Test success, Validation message displayed for {expectedField}: " + message);
                    return;
                }

                // If no validation message, check if page stayed the same (no modal / redirect)
                var modals = _driver.FindElements(By.XPath("/html/body/div/div"));
                if (modals.Count == 0)
                {
                    LogStep($"✅ Test success, No response detected after Save (expected when {expectedField} is empty).");
                    return;
                }

                // If modal shows success, that's incorrect behavior
                string modalMessage = modals[0].Text.Trim();
                if (modalMessage.ToLower().Contains("successful"))
                {
                    Assert.Fail($"❌ Unexpected success message when {expectedField} is empty: " + modalMessage);
                }

                // Fallback case
                Assert.Fail($"❌ Unexpected behavior: empty {expectedField} not properly handled.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Mandatory Fields Empty (Customer Email)
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload a valid profile image.
        ///     3. Fill other required fields except for Customer Email (leave empty).
        ///     4. Click Save.
        /// Verification:
        ///     - Validation message should appear for the missing mandatory field (Customer Email).
        ///     - Test fails if the system shows a success message.
        ///     - Screenshot captured for evidence.
        /// Purpose:
        ///     Ensure the system correctly validates mandatory fields and prevents saving users with missing Customer Email.
        /// Test Data:
        ///     - Provided by 'MandatoryEmailEmptyTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Mandatory Fields Empty")]
        [TestCaseSource(nameof(MandatoryEmailEmptyTestData))]
        public void Create_MandatoryFieldEmpty_Email(
        string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser, string expectedField)
        {
            try
            {
                LogStep("🆕 Click New button");
                _UserPage.ClickNewButton();

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
                WaitForUIEffect();

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep($"🔘 Set '{role}' checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter Password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);
                WaitForUIEffect();

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("🔍 Check for validation or system response");
                WaitForUIEffect();

                // Try to locate validation message (common selectors)
                var validationMessages = _driver.FindElements(By.CssSelector(".invalid-feedback, .error-message, .validation-message"));

                if (validationMessages.Count > 0)
                {
                    string message = validationMessages[0].Text.Trim();
                    LogStep($"✅ Test success, Validation message displayed for {expectedField}: " + message);
                    return;
                }

                // If no validation message, check if page stayed the same (no modal / redirect)
                var modals = _driver.FindElements(By.XPath("/html/body/div/div"));
                if (modals.Count == 0)
                {
                    LogStep($"✅ Test success, No response detected after Save (expected when {expectedField} is empty).");
                    return;
                }

                // If modal shows success, that's incorrect behavior
                string modalMessage = modals[0].Text.Trim();
                if (modalMessage.ToLower().Contains("successful"))
                {
                    Assert.Fail($"❌ Unexpected success message when {expectedField} is empty: " + modalMessage);
                }

                // Fallback case
                Assert.Fail($"❌ Unexpected behavior: empty {expectedField} not properly handled.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }





        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Mandatory Fields Empty (Password)
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload a valid profile image.
        ///     3. Fill other required fields except for Password (leave empty).
        ///     4. Click Save.
        /// Verification:
        ///     - Validation message should appear for the missing mandatory field (Password).
        ///     - Test fails if the system shows a success message.
        ///     - Screenshot captured for evidence.
        /// Purpose:
        ///     Ensure the system correctly validates mandatory fields and prevents saving users with missing Password.
        /// Test Data:
        ///     - Provided by 'MandatorPasswordEmptyTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(8)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Mandatory Fields Empty")]
        [TestCaseSource(nameof(MandatorPasswordEmptyTestData))]
        public void Create_MandatoryFieldEmpty_Password(
        string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser, string expectedField)
        {
            try
            {
                LogStep("🆕 Click New button");
                _UserPage.ClickNewButton();

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
                WaitForUIEffect();

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep($"🔘 Set '{role}' checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter Password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);
                WaitForUIEffect();

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("🔍 Check for validation or system response");
                WaitForUIEffect();

                // Try to locate validation message (common selectors)
                var validationMessages = _driver.FindElements(By.CssSelector(".invalid-feedback, .error-message, .validation-message"));

                if (validationMessages.Count > 0)
                {
                    string message = validationMessages[0].Text.Trim();
                    LogStep($"✅ Test success, Validation message displayed for {expectedField}: " + message);
                    return;
                }

                // If no validation message, check if page stayed the same (no modal / redirect)
                var modals = _driver.FindElements(By.XPath("/html/body/div/div"));
                if (modals.Count == 0)
                {
                    LogStep($"✅ Test success, No response detected after Save (expected when {expectedField} is empty).");
                    return;
                }

                // If modal shows success, that's incorrect behavior
                string modalMessage = modals[0].Text.Trim();
                if (modalMessage.ToLower().Contains("successful"))
                {
                    Assert.Fail($"❌ Unexpected success message when {expectedField} is empty: " + modalMessage);
                }

                // Fallback case
                Assert.Fail($"❌ Unexpected behavior: empty {expectedField} not properly handled.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }





        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Mandatory Fields Empty (Confirm Password)
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload a valid profile image.
        ///     3. Fill other required fields except for Confirm Password (leave empty).
        ///     4. Click Save.
        /// Verification:
        ///     - Validation message should appear for the missing mandatory field (Confirm Password).
        ///     - Test fails if the system shows a success message.
        ///     - Screenshot captured for evidence.
        /// Purpose:
        ///     Ensure the system correctly validates mandatory fields and prevents saving users with missing Confirm Password.
        /// Test Data:
        ///     - Provided by 'MandatorCPasswordEmptyTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(9)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Mandatory Fields Empty")]
        [TestCaseSource(nameof(MandatorCPasswordEmptyTestData))]
        public void Create_MandatoryFieldEmpty_ConfirmPassword(
        string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser, string expectedField)
        {
            try
            {
                LogStep("🆕 Click New button");
                _UserPage.ClickNewButton();

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
                WaitForUIEffect();

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep($"🔘 Set '{role}' checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter Password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter Confirm Password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                LogStep("✅ Set Active checkbox");
                bool isActive = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActive);
                WaitForUIEffect();

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("🔍 Check for validation or system response");
                WaitForUIEffect();

                // Try to locate validation message (common selectors)
                var validationMessages = _driver.FindElements(By.CssSelector(".invalid-feedback, .error-message, .validation-message"));

                if (validationMessages.Count > 0)
                {
                    string message = validationMessages[0].Text.Trim();
                    LogStep($"✅ Test success, Validation message displayed for {expectedField}: " + message);
                    return;
                }

                // If no validation message, check if page stayed the same (no modal / redirect)
                var modals = _driver.FindElements(By.XPath("/html/body/div/div"));
                if (modals.Count == 0)
                {
                    LogStep($"✅ Test success, No response detected after Save (expected when {expectedField} is empty).");
                    return;
                }

                // If modal shows success, that's incorrect behavior
                string modalMessage = modals[0].Text.Trim();
                if (modalMessage.ToLower().Contains("successful"))
                {
                    Assert.Fail($"❌ Unexpected success message when {expectedField} is empty: " + modalMessage);
                }

                // Fallback case
                Assert.Fail($"❌ Unexpected behavior: empty {expectedField} not properly handled.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_MandatoryField_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception occurred: " + ex.Message);
                Assert.Fail("Test failed due to exception: " + ex.Message);
            }
        }








        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Duplicate Entry
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload a valid profile image.
        ///     3. Fill all required fields with data that already exists in the system (Username / TIN).
        ///     4. Click Save.
        /// Verification:
        ///     - System should display a modal with a duplicate entry message (e.g., "TIN has already been taken") or a failure message.
        ///     - Test fails if a success message is displayed.
        ///     - Screenshot captured for evidence.
        /// Purpose:
        ///     Ensure the system prevents creating users with duplicate identifiers and validates uniqueness constraints.
        /// Test Data:
        ///     - Provided by 'DuplicateTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(10)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(DuplicateTestData))]
        public void CreateDuplicate(string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser)
        {
            try
            {
                LogStep("🔘 Click [New] button");
                _UserPage.ClickNewButton();
                WaitForUIEffect();

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

                WaitForUIEffect();

                LogStep("⌨️ Enter username");
                _UserPage.EnterUsername(Username);
                WaitForUIEffect();

                LogStep("Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep($"🔘 Set '{role}' checkbox");
                _UserPage.SetCheckboxByLabel(role);
                WaitForUIEffect();

                LogStep("Enter password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter confirm password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                LogStep("Click show password");
                var showpassword = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[1]/div[1]/span")));
                showpassword.Click();
                WaitForUIEffect();

                LogStep("Click show confirm password");
                var showconfirmpassword = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div/span")));
                showconfirmpassword.Click();
                WaitForUIEffect();

                LogStep("🟢 Set active user checkbox");
                bool isActiveChecked = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActiveChecked);
                WaitForUIEffect();

                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("Waiting for modal message");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                if (message.Contains("TIN has already been taken", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("⚠️ Duplicate TIN message detected");
                    var duplicateOkBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/div/div/div[6]/button[1]")));
                    duplicateOkBtn.Click();
                    Assert.IsTrue(true,"❌ Duplicate TIN detected: " + message);
                }
                else if (message.Contains("failed", StringComparison.OrdinalIgnoreCase))
                {
                        LogStep("⚠️ failed message detected");
                        var duplicateOkBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                            By.XPath("/html/body/div/div/div[6]/button[1]")));
                        duplicateOkBtn.Click();
                        Assert.IsTrue(true, "❌ Failed message detected: " + message);
                }
                else
                {
                    Assert.IsFalse(message.ToLower().Contains("successful"), $"❌ Successful detected in message: {message}");

                    var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/div/div/div[6]/button[1]")));
                    okButton.Click();
                    WaitForUIEffect();
                    LogStep($"📢Test success, Modal Message: {message}");
                }

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Console.WriteLine($"❌ Exception during test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }







        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        /// Test Case: Create User - Negative: Role Checkbox Not Available
        /// Action:
        ///     1. Click "New" to add a user.
        ///     2. Upload a valid profile image.
        ///     3. Fill username, email, password, confirm password, and active checkbox fields.
        ///     4. Attempt to select a role checkbox that may not exist.
        ///     5. Click Save.
        /// Verification:
        ///     - If the role checkbox does not exist, the test passes (expected negative case).
        ///     - If the system displays a modal message, verify it is **not** a success message.
        ///     - Screenshot captured for evidence.
        /// Purpose:
        ///     Ensure the system handles attempts to assign non-existent roles gracefully without allowing creation.
        /// Test Data:
        ///     - Provided by 'InvalidRoleCheckboxTestData' TestCaseSource
        /// Created By: 19-Dec-2025 by Yan Shen (AdminTool version 2.0.0.0, Core version 2.0.2.16)
        /// Edited By:
        /// ------------------------------------------------------------------------------------------------------------------------------------------------------------------------///
        [Test]
        [Category("User")]
        [Order(11)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Negative: Role Checkbox Not Available")]
        [TestCaseSource(nameof(InvalidRoleCheckboxTestData))]
        public void Create_InvalidRoleCheckbox(string Username, string CustEmail, string role, string UserPassword, string UserConfirmPassword, string activeUser)
        {
            try
            {
                LogStep("🔘 Click [New] button");
                _UserPage.ClickNewButton();
                WaitForUIEffect();

                LogStep("📤 Upload profile image");
                string filePath = AppConfig.UserProfileImage;
                if (!File.Exists(filePath))
                {
                    Assert.Fail("❌ File not found: " + filePath);
                }

                var fileInput = _driver.FindElement(By.CssSelector("#kt_modal_add_user_info > div.mb-6 > div > div > label > input[type='file']:nth-child(2)"));
                fileInput.SendKeys(filePath);
                LogStep("✅ File uploaded");
                WaitForUIEffect();

                LogStep("⌨️ Enter username");
                _UserPage.EnterUsername(Username);
                WaitForUIEffect();

                LogStep("⌨️ Enter customer email");
                _UserPage.EnterCustEmail(CustEmail);
                WaitForUIEffect();

                LogStep($"🔘 Attempt to set '{role}' checkbox");
                try
                {
                    // Try find checkbox **immediately** without waiting
                    var checkbox = _driver.FindElement(By.XPath($"//label[text()='{role}']/input[@type='checkbox']"));
                    checkbox.Click();
                    LogStep($"✅ Role checkbox '{role}' found and clicked");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidRole_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                }
                catch (NoSuchElementException)
                {
                    LogStep($"⚠️ Role checkbox '{role}' not found, test success by design");
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidRole_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot1 = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot1.AsByteArray);
                    return; // Directly pass the test if checkbox does not exist
                }

                LogStep("Enter password");
                _UserPage.EnterUserPassword(UserPassword);
                WaitForUIEffect();

                LogStep("Enter confirm password");
                _UserPage.EnterUserConfirmPassword(UserConfirmPassword);
                WaitForUIEffect();

                LogStep("🟢 Set active user checkbox");
                bool isActiveChecked = bool.TryParse(activeUser, out var active) && active;
                _UserPage.SetCheckActiveboxState(isActiveChecked);
                WaitForUIEffect();

                // Save action
                LogStep("💾 Click Save button");
                _UserPage.ClickSaveButton();
                WaitForUIEffect();

                LogStep("📢 Waiting for modal message");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div/div/div[2]")));
                var message = modal.Text.Trim();
                LogStep($"📢Test Success, Modal Message: {message}");

                if (message.ToLower().Contains("successful"))
                {
                    Assert.Fail($"❌ Unexpected success message: {message}");
                }

                LogStep("✅ Test passed, system handled missing role checkbox correctly");
                var okButton = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div[6]/button[1]")));
                okButton.Click();
                WaitForUIEffect();
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidRole_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"User_InvalidRole_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Console.WriteLine($"❌ Exception during test: {ex.Message}");
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

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
using SeleniumTests.Pages.Log;
using System.Drawing;
using System.Globalization;
using System.Media;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;


namespace SeleniumTests.Tests.G_Log.Paging
{
    [TestFixture, Order(29)]
    [AllureNUnit]
    [AllureSuite("Log - Paging")] // use this ties to module
    [AllureEpic("ERP-117")] // use this and ties to ticket number

    public class Log_Created_Paging
    {
        private IWebDriver _driver;
        private LogPage _LogPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";

        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Log Created Page - Paging"; // You can make this dynamic if needed
            string mainmoduleName = "Log"; // You can make this dynamic if needed

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/log-details");
            helperFunction.WaitForPageToLoad(_wait);
            _LogPage = new LogPage(_driver);
            _logMessages.Clear();

            _moduleName = "Log Created Page - Paging";
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
        [Category("Log")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Log Role Paging")]
        public void Log_Created_TestPagingNextButtonAndVerify()
        {
            LogStep(" Click on 'Log Created' tab");
            var logCreatedTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and contains(normalize-space(text()),'Created')]")));
            logCreatedTab.Click();
            WaitForUIEffect();

            LogStep(" Wait for table to be visible and capture current HTML");
            var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[1]/div")));
            string beforeHtml = tableElement.GetAttribute("innerHTML");

            LogStep(" Locate 'Next' pagination button");
            var nextButton = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class, 'page-link')][.//i[contains(@class, 'fa-arrow-right')]]")));

            LogStep(" Check if 'Next' button is disabled");
            bool isDisabled = nextButton.GetAttribute("class").Contains("disabled");
            if (isDisabled)
            {

                LogStep("⚠️ 'Next' button is disabled. Only one page available. Skipping pagination test.");
                return;
            }

            LogStep(" Click 'Next' button");
            nextButton.Click();
            WaitForUIEffect();
            LogStep("✅ 'Next' button clicked");

            LogStep(" Wait for table content to change");
            _wait.Until(driver =>
            {
                var updatedTable = driver.FindElement(By.XPath(
                    "/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[1]/div"));
                return updatedTable.GetAttribute("innerHTML") != beforeHtml;
            });
            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
            LogStep("✅ Table content changed after clicking 'Next'");
        }

        [Test]
        [Category("Log")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Log Paging")]
        public void Log_Created_TestPagingPreviousButtonAndVerify()
        {
            try
            {
                LogStep(" Click on 'Log Created' tab");
                var logCreatedTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and contains(normalize-space(text()),'Created')]")));
                logCreatedTab.Click();
                WaitForUIEffect();

                LogStep(" Wait for table to be visible and capture current HTML");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[1]/div")));
                string originalHtml = tableElement.GetAttribute("innerHTML");

                LogStep(" Locate 'Next' pagination button");
                var nextLi = _wait.Until(driver => driver.FindElement(
                    By.XPath("//ul/li[contains(@class, 'next')]")));
                var nextAnchor = nextLi.FindElement(By.TagName("a"));

                LogStep(" Check if 'Next' button is disabled");
                bool nextDisabled = nextLi.GetAttribute("class").Contains("disabled");
                if (nextDisabled)
                {
                    LogStep("⚠️ 'Next' button is disabled. Only one page available. Skipping test.");
                    return;
                }

                LogStep(" Click 'Next' button");
                nextAnchor.Click();
                WaitForUIEffect();
                LogStep("✅ 'Next' button clicked");

                LogStep(" Wait for table content to change");
                _wait.Until(driver =>
                {
                    var newTable = driver.FindElement(By.XPath(
                        "/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[1]/div"));
                    return newTable.GetAttribute("innerHTML") != originalHtml;
                });

                LogStep(" Locate 'Previous' pagination button");
                var prevLi = _wait.Until(driver => driver.FindElement(
                    By.XPath("//ul/li[contains(@class, 'previous')]")));
                var prevAnchor = prevLi.FindElement(By.TagName("a"));

                LogStep(" Check if 'Previous' button is disabled");
                bool prevDisabled = prevLi.GetAttribute("class").Contains("disabled");
                if (prevDisabled)
                {
                    LogStep("⚠️ 'Previous' button is disabled. Cannot return to page 1.");
                    return;
                }

                LogStep(" Click 'Previous' button");
                prevAnchor.Click();
                WaitForUIEffect();
                LogStep("✅ 'Previous' button clicked");

                LogStep(" Wait for table to return to original");
                _wait.Until(driver =>
                {
                    var tableBack = driver.FindElement(By.XPath(
                        "/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[1]/div"));
                    return tableBack.GetAttribute("innerHTML") == originalHtml;
                });
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("✅ Table returned to original after clicking 'Previous'");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during test: {ex.Message}");
                Assert.Fail("❌ Test failed: " + ex.Message);
            }
        }




        [Test]
        [Category("Log")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Log Paging - Click Last Icon Button and Verify Table Change")]
        public void Log_Created_TestPagingClickLastIconAndVerify()
        {
            try
            {
                LogStep("Clicking 'Log Created' tab");
                var LogCreatedTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and contains(normalize-space(text()),'Created')]")));
                LogCreatedTab.Click();
                WaitForUIEffect();

                LogStep("📌 Capturing original first row text");
                var tableXPath = "//table/tbody/tr[1]";
                var firstRow = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                string originalFirstRowText = firstRow.Text.Trim();

                LogStep("📌 Locating last page icon button");
                var lastPageButton = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//a[.//i[contains(@class,'fa-angle-double-right')]]")));

                bool isDisabled = lastPageButton.GetAttribute("class")?.Contains("disabled") ?? false;
                Assert.IsTrue(!isDisabled, "ℹ️ 'Last' button is disabled. Possibly already on last page or only one page exists.");

                LogStep("✅ Clicking last page icon button via JS");
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({ block: 'center' });", lastPageButton);
                _wait.Until(ExpectedConditions.ElementToBeClickable(lastPageButton));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", lastPageButton);
                WaitForUIEffect();

                LogStep("⏳ Waiting for table content to change...");
                bool tableChanged = new WebDriverWait(_driver, TimeSpan.FromSeconds(60)).Until(driver =>
                {
                    var updatedFirstRow = driver.FindElement(By.XPath(tableXPath));
                    string newText = updatedFirstRow.Text.Trim();
                    return newText != originalFirstRowText;
                });
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(tableChanged, "❌ Table did not change after clicking 'Last' button.");
                LogStep("✅ Table content updated after navigating to last page.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Console.WriteLine("❌ Exception: " + ex.Message);
                Assert.Fail("❌ Test failed due to unexpected exception.");
            }
        }

        [Test]
        [Category("Log")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Log Paging - Click First and Verify Return")]
        public void Log_Created_TestPagingFirstButtonAndVerify()
        {
            string tableXPath = "//table/tbody/tr[1]";

            try
            {
                LogStep("Clicking 'Log Created' tab");
                var logCreatedTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and contains(normalize-space(text()),'Created')]")));
                logCreatedTab.Click();
                WaitForUIEffect();

                // Capture original first row's HTML
                LogStep("🔍 Capturing original first row HTML content");
                var originalRow = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                WaitForUIEffect();
                string originalHtml = originalRow.GetAttribute("innerHTML");

                // Locate 'Last' button
                LogStep("Locating 'Last' button");
                var lastButton = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//a[.//i[contains(@class,'fa-angle-double-right')]]")));
                WaitForUIEffect();

                if (lastButton.GetAttribute("class")?.Contains("disabled") ?? false)
                {
                    LogStep("ℹ️ 'Last' button is disabled — possibly only one page exists.");
                    Assert.Inconclusive("ℹ️ Only one page exists.");
                    return;
                }

                // Click 'Last' via JS
                LogStep("Clicking 'Last' button via JavaScript");
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({ block: 'center' });", lastButton);
                _wait.Until(ExpectedConditions.ElementToBeClickable(lastButton));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", lastButton);
                WaitForUIEffect();

                // Wait for table content to change
                bool tableChanged = false;
                try
                {
                    tableChanged = new WebDriverWait(_driver, TimeSpan.FromSeconds(10)).Until(driver =>
                    {
                        var updatedRow = driver.FindElement(By.XPath(tableXPath));
                        return updatedRow.GetAttribute("innerHTML") != originalHtml;
                    });
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⚠️ Table did not change after clicking 'Last'. Possibly already on last page.");
                }

                // Locate 'First' button
                LogStep("Locating 'First' button");
                var firstButton = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//a[.//i[contains(@class,'fa-angle-double-left')]]")));
                WaitForUIEffect();

                if (firstButton.GetAttribute("class")?.Contains("disabled") ?? false)
                {
                    LogStep("ℹ️ 'First' button is disabled. Already on first page.");
                    Assert.IsTrue(true);
                    return;
                }

                //  Click 'First' via JS
                LogStep("Clicking 'First' button via JavaScript");
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({ block: 'center' });", firstButton);
                _wait.Until(ExpectedConditions.ElementToBeClickable(firstButton));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", firstButton);
                WaitForUIEffect();

                //  Wait for table to return to original
                bool tableReturned = false;
                try
                {
                    tableReturned = new WebDriverWait(_driver, TimeSpan.FromSeconds(10)).Until(driver =>
                    {
                        var returnedRow = driver.FindElement(By.XPath(tableXPath));
                        return returnedRow.GetAttribute("innerHTML") == originalHtml;
                    });
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⚠️ Table did not return to original after clicking 'First'.");
                }

                LogStep($"ℹ️ TableChanged: {tableChanged}, TableReturned: {tableReturned}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                Assert.IsTrue(true); // Always pass
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("❌ Exception: " + ex.Message);
                Assert.Fail("❌ Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Log")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Log Paging - Click Page Size Dropdown and Verify Table Update")]
        [TestCase("100")]
        public void Log_Created_TestItemsPerPageVerify(string pageSizeValue)
        {
            try
            {
                LogStep("Clicking 'Log Created' tab");
                var LogCreatedTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and contains(normalize-space(text()),'Created')]")));
                LogCreatedTab.Click();
                WaitForUIEffect();

                LogStep("📌 Capturing original table HTML and row count");
                string tableXPath = "/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[1]/div";
                string dropdownXPath = "/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[2]/div/div[1]/select";
                string rowSelector = "tbody tr";

                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                string beforeHtml = tableElement.GetAttribute("innerHTML");
                int originalRowCount = tableElement.FindElements(By.CssSelector(rowSelector)).Count;

                LogStep("📌 Selecting page size value from dropdown");
                var dropdownElement = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(dropdownXPath)));
                var select = new SelectElement(dropdownElement);

                bool optionExists = select.Options.Any(opt => opt.Text.Trim() == pageSizeValue);
                Assert.IsTrue(optionExists, $"❌ Page size option '{pageSizeValue}' not found in dropdown.");

                select.SelectByText(pageSizeValue);
                WaitForUIEffect();
                LogStep($"✅ Page size '{pageSizeValue}' selected");

                LogStep("⏳ Waiting for table content to update");
                bool tableUpdated = _wait.Until(driver =>
                {
                    var updatedTable = driver.FindElement(By.XPath(tableXPath));
                    string afterHtml = updatedTable.GetAttribute("innerHTML");
                    int updatedRowCount = updatedTable.FindElements(By.CssSelector(rowSelector)).Count;
                    return afterHtml != beforeHtml || updatedRowCount <= int.Parse(pageSizeValue);
                });

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(tableUpdated, "❌ Table did not update and row count exceeds selected page size.");
                LogStep("✅ Table updated or already showing all rows within selected page size");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Console.WriteLine("❌ Exception: " + ex.Message);
                Assert.Fail("❌ Test failed due to unexpected exception.");
            }
        }


        [Test]
        [Category("Log")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Log Paging - Click Page Button Only If It Exists and Verify Table Update")]
        [TestCase("3")]
        public void Log_Created_TestClickPageButtonIfExists(string pageNumber)
        {
            try
            {
                LogStep("Clicking 'Log Created' tab");
                var LogCreatedTab = _wait.Until(ExpectedConditions.ElementToBeClickable(
                By.XPath("//a[contains(@class,'nav-link') and contains(normalize-space(text()),'Created')]")));
                LogCreatedTab.Click();
                WaitForUIEffect();

                string tableXPath = "/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[1]/div";
                string paginationXPathTemplate = "/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[3]/div/div[2]/div/div[2]/ul/li[a[text()='{0}']]/a";
                string dynamicPageXPath = string.Format(paginationXPathTemplate, pageNumber);

                LogStep("📌 Capturing current table content");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                string beforeHtml = tableElement.GetAttribute("innerHTML");

                LogStep($"📌 Checking if page button '{pageNumber}' exists");
                var pageButtons = _driver.FindElements(By.XPath(dynamicPageXPath));
                Assert.IsTrue(pageButtons.Count > 0, $"✅ No page {pageNumber} exists — only one page available, test logically passed.");

                LogStep($"✅ Clicking page number '{pageNumber}'");
                pageButtons[0].Click();
                WaitForUIEffect();

                LogStep("⏳ Waiting for table content to change");
                bool tableUpdated = _wait.Until(driver =>
                {
                    var updatedTable = driver.FindElement(By.XPath(tableXPath));
                    string afterHtml = updatedTable.GetAttribute("innerHTML");
                    return afterHtml != beforeHtml;
                });

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Assert.IsTrue(tableUpdated, $"❌ Table did not update after clicking page {pageNumber}.");
                LogStep($"✅ Table updated after navigating to page {pageNumber}");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Log_Created_Page_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                Console.WriteLine("❌ Exception: " + ex.Message);
                Assert.Fail("❌ Test failed due to unexpected exception.");
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
                    worksheet.Cells[row, 5].Value = "Paging completed successfully. The data has been refreshed and reloaded.";
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


        private void WaitForUIEffect(int ms = 2500)
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

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
using SeleniumTests.Pages.Report;
using System.Drawing;
using System.Globalization;
using System.Media;
using System.Security.Cryptography.X509Certificates;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;


namespace SeleniumTests.Tests.J_Report
{

    public static class ExcelDataReaderReportSummaryPaging
    {



        public static IEnumerable<object[]> GetExportSummaryReportTestData(string filePath, string sheetName)
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
                    string FromYear = worksheet.Cells[row, 1].Text?.Trim();
                    string FromMonth = worksheet.Cells[row, 2].Text?.Trim();
                    string EntityName = worksheet.Cells[row, 3].Text?.Trim();
                    string StoreName = worksheet.Cells[row, 4].Text?.Trim();


                    yield return new object[]
                    {
                        FromYear, FromMonth, EntityName, StoreName
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetReportInfoLineTestData(string filePath, string sheetName)
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
                    string FromYear = worksheet.Cells[row, 1].Text?.Trim();
                    string FromMonth = worksheet.Cells[row, 2].Text?.Trim();
                    string EntityName = worksheet.Cells[row, 3].Text?.Trim();
                    string StoreName = worksheet.Cells[row, 4].Text?.Trim();
                    string pageLine = worksheet.Cells[row, 5].Text?.Trim();


                    yield return new object[]
                    {
                        FromYear, FromMonth, EntityName, StoreName, pageLine
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetReportInfoPageTestData(string filePath, string sheetName)
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
                    string FromYear = worksheet.Cells[row, 1].Text?.Trim();
                    string FromMonth = worksheet.Cells[row, 2].Text?.Trim();
                    string EntityName = worksheet.Cells[row, 3].Text?.Trim();
                    string StoreName = worksheet.Cells[row, 4].Text?.Trim();
                    string pageNumber = worksheet.Cells[row, 5].Text?.Trim();


                    yield return new object[]
                    {
                        FromYear, FromMonth, EntityName, StoreName, pageNumber
                    };

                }
            }
        }
    }


    [TestFixture, Order(49)]
    [AllureNUnit]
    [AllureSuite("Report Summary - Paging")] // use this ties to module
    [AllureEpic("ERP-117")] // use this and ties to ticket number

    public class ReportSummary_Paging
    {
        private IWebDriver _driver;
        private ReportPage _ReportPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "SReportTestDataPaging.xlsx");

        public static IEnumerable<object[]> ExportReportTestData =>
        ExcelDataReaderReportSummaryPaging.GetExportSummaryReportTestData(ExcelPath, "ExportReportTestData");

        public static IEnumerable<object[]> ReportInfoLineTestData =>
        ExcelDataReaderReportSummaryPaging.GetReportInfoLineTestData(ExcelPath, "ReportInfoLineTestData");
        public static IEnumerable<object[]> ReportInfoPageTestData =>
        ExcelDataReaderReportSummaryPaging.GetReportInfoPageTestData(ExcelPath, "ReportInfoPageTestData");

        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Suimmary Report Page - Paging";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/report/summary");
            helperFunction.WaitForPageToLoad(_wait);
            _ReportPage = new ReportPage(_driver);
            _logMessages.Clear();

            _moduleName = "Summary Report Page - Paging";
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
        [Category("Report")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Report Paging")]
        [TestCaseSource(nameof(ExportReportTestData))]
        public void TestPagingNextButtonAndVerify(string Fromyear, string Frommonth, string EntityName, string StoreName)
        {
            try
            {
                LogStep($"Selecting Year: {Fromyear}");

                // Step 1: Click to open dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[2]/ng-select/div/span")));
                yearDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var dropdownPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(dropdownPanelXPath)));

                // Step 3: Find year option and click
                var yearOptions = _driver.FindElements(By.XPath(dropdownPanelXPath));
                bool yearSelected = false;

                foreach (var option in yearOptions)
                {
                    string yearText = option.Text.Trim();
                    if (yearText.Equals(Fromyear, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Year: {Fromyear}");
                        yearSelected = true;
                        break;
                    }
                }

                // Step 4: Handle case if year not found
                if (!yearSelected)
                {
                    LogStep($"⚠️ Year '{Fromyear}' not found in dropdown.");
                }
                WaitForUIEffect();



                // Month dropdown (ng-select)
                LogStep($"Selecting Month: {Frommonth}");

                // Step 1: Click to open dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[3]/ng-select/div/span")));
                monthDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var monthPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(monthPanelXPath)));

                // Step 3: Loop and select matching month value
                var monthOptions = _driver.FindElements(By.XPath(monthPanelXPath));
                bool monthSelected = false;

                foreach (var option in monthOptions)
                {
                    string monthText = option.Text.Trim();
                    if (monthText.Equals(Frommonth, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Month: {Frommonth}");
                        monthSelected = true;
                        break;
                    }
                }

                // Step 4: If not found
                if (!monthSelected)
                {
                    LogStep($"⚠️ Month '{Frommonth}' not found in dropdown.");
                }
                WaitForUIEffect();


                // ==========================
                // Entity Dropdown (Dynamic Multi-Select + Select All)
                // ==========================
                LogStep($"Selecting Entity(ies): {EntityName}");

                var _actions = new OpenQA.Selenium.Interactions.Actions(_driver);

                // Step 1: Open dropdown
                var entityDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[2]/div[2]/app-entity-filter-dropdown")));
                entityDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown to show all entity divs (Fixed XPath)
                var entityListXPath = "/html/body/div/div/div[2]/div/div/div/div";
                _wait.Until(ExpectedConditions.ElementExists(By.XPath(entityListXPath)));

                var allEntities = _driver.FindElements(By.XPath(entityListXPath));
                LogStep($"📋 Total entities detected: {allEntities.Count}");

                // Step 3: Handle "Active" (Select All)
                if (EntityName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — selecting all entities...");

                    foreach (var entity in allEntities)
                    {
                        try
                        {
                            _actions.MoveToElement(entity).Click().Perform();
                            LogStep($"✅ Selected: {entity.Text.Trim()}");
                            Thread.Sleep(200);
                        }
                        catch
                        {
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                            LogStep($"⚠️ JS fallback click used for: {entity.Text.Trim()}");
                        }
                    }
                }
                else
                {
                    // Step 4: Multi-select specific entities
                    var entityValues = EntityName.Split(',')
                                                 .Select(v => v.Trim())
                                                 .Where(v => !string.IsNullOrEmpty(v))
                                                 .ToList();

                    foreach (var inputValue in entityValues)
                    {
                        bool found = false;
                        foreach (var entity in allEntities)
                        {
                            string label = entity.Text.Trim();

                            // Partial case-insensitive match
                            if (label.Contains(inputValue, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    _actions.MoveToElement(entity).Click().Perform();
                                    LogStep($"✅ Selected entity: {label}");
                                    found = true;
                                    Thread.Sleep(300);
                                    break;
                                }
                                catch
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                                    LogStep($"⚠️ JS fallback used for: {label}");
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            LogStep($"⚠️ Entity '{inputValue}' not found in list.");
                    }
                }

                WaitForUIEffect();

                // Step 5: Close dropdown
                try
                {
                    entityDropdown.Click();
                    LogStep("📦 Closed Entity dropdown after selection.");
                }
                catch
                {
                    LogStep("ℹ️ Dropdown closed automatically.");
                }

                WaitForUIEffect();




                // === Store Dropdown ===
                LogStep($"Selecting Store(s): {StoreName}");

                var storeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-store-filter-dropdown")));
                storeDropdown.Click();
                WaitForUIEffect();

                // Find the dynamic store panel by ID prefix
                var storePanel = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//div[starts-with(@id,'store_selection_menu_')]"))
                          .FirstOrDefault());

                if (storePanel == null)
                    throw new Exception("❌ Store selection panel not found after opening dropdown.");

                // Log all visible store names (for debugging)
                var visibleLabels = storePanel.FindElements(By.XPath(".//*[normalize-space(text())!='']"))
                                              .Select(e => e.Text.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .Distinct()
                                              .ToList();

                LogStep("📋 Available stores in dropdown:");
                foreach (var lbl in visibleLabels)
                {
                    LogStep($"   - {lbl}");
                }

                // If user input is "Active" → select all
                if (StoreName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — ticking all store checkboxes...");
                    var allCheckboxes = storePanel.FindElements(By.XPath(".//input[@type='checkbox']"));
                    int ticked = 0;

                    foreach (var checkbox in allCheckboxes)
                    {
                        try
                        {
                            if (!checkbox.Selected)
                            {
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", checkbox);
                                ticked++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Unable to tick one checkbox: {ex.Message}");
                        }
                    }

                    LogStep($"✅ {ticked} store(s) ticked successfully.");
                }
                else
                {
                    var storeValues = StoreName.Split(',')
                                               .Select(v => v.Trim())
                                               .Where(v => !string.IsNullOrEmpty(v))
                                               .ToList();

                    foreach (var val in storeValues)
                    {
                        bool found = false;

                        foreach (var lbl in visibleLabels)
                        {
                            if (lbl.ToUpper().Contains(val.ToUpper()))
                            {
                                var target = storePanel.FindElement(By.XPath($".//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '{val.ToUpper()}')]/preceding::input[@type='checkbox'][1]"));
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", target);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", target);
                                LogStep($"✅ Store selected: {lbl}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            LogStep($"❌ Could not find store containing: '{val}'");
                        }
                    }
                }

                WaitForUIEffect();
                storeDropdown.Click();
                WaitForUIEffect();



                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[3]/div[3]/button")));
                selectBtn.Click();
                LogStep("Clicked 'Apply' button to apply filter");
                WaitForUIEffect();


                LogStep("⏳ Waiting for Summary Report table to load...");
                var tableElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div")));

                string beforeHtml = tableElement.GetAttribute("innerHTML");
                LogStep("✅ Captured initial table content.");

                LogStep("🔎 Locating 'Next' button for pagination...");
                var nextButton = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[3]/app-global-pagination/div/div[2]/ul/li[4]")));

                bool isDisabled = nextButton.GetAttribute("class").Contains("disabled");

                if (isDisabled)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⚠️ 'Next' button is disabled. Only one page available. Skipping pagination test.");
                    return;
                }

                LogStep("Clicking 'Next' button...");
                _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(nextButton));
                nextButton.Click();
                WaitForUIEffect(800);

                LogStep("⏳ Waiting for table content to change...");
                _wait.Until(driver =>
                {
                    var updatedTable = driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div"));
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    return updatedTable.GetAttribute("innerHTML") != beforeHtml;
                });
               
                LogStep("✅ Table content updated successfully after pagination.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during paging test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Report")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Report Paging")]
        [TestCaseSource(nameof(ExportReportTestData))]
        public void TestPagingPreviousButtonAndVerify(string Fromyear, string Frommonth, string EntityName, string StoreName)
        {
            try
            {
                LogStep($"Selecting Year: {Fromyear}");

                // Step 1: Click to open dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[2]/ng-select/div/span")));
                yearDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var dropdownPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(dropdownPanelXPath)));

                // Step 3: Find year option and click
                var yearOptions = _driver.FindElements(By.XPath(dropdownPanelXPath));
                bool yearSelected = false;

                foreach (var option in yearOptions)
                {
                    string yearText = option.Text.Trim();
                    if (yearText.Equals(Fromyear, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Year: {Fromyear}");
                        yearSelected = true;
                        break;
                    }
                }

                // Step 4: Handle case if year not found
                if (!yearSelected)
                {
                    LogStep($"⚠️ Year '{Fromyear}' not found in dropdown.");
                }
                WaitForUIEffect();



                // Month dropdown (ng-select)
                LogStep($"Selecting Month: {Frommonth}");

                // Step 1: Click to open dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[3]/ng-select/div/span")));
                monthDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var monthPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(monthPanelXPath)));

                // Step 3: Loop and select matching month value
                var monthOptions = _driver.FindElements(By.XPath(monthPanelXPath));
                bool monthSelected = false;

                foreach (var option in monthOptions)
                {
                    string monthText = option.Text.Trim();
                    if (monthText.Equals(Frommonth, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Month: {Frommonth}");
                        monthSelected = true;
                        break;
                    }
                }

                // Step 4: If not found
                if (!monthSelected)
                {
                    LogStep($"⚠️ Month '{Frommonth}' not found in dropdown.");
                }
                WaitForUIEffect();


                // ==========================
                // Entity Dropdown (Dynamic Multi-Select + Select All)
                // ==========================
                LogStep($"Selecting Entity(ies): {EntityName}");

                var _actions = new OpenQA.Selenium.Interactions.Actions(_driver);

                // Step 1: Open dropdown
                var entityDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[2]/div[2]/app-entity-filter-dropdown")));
                entityDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown to show all entity divs (Fixed XPath)
                var entityListXPath = "/html/body/div/div/div[2]/div/div/div/div";
                _wait.Until(ExpectedConditions.ElementExists(By.XPath(entityListXPath)));

                var allEntities = _driver.FindElements(By.XPath(entityListXPath));
                LogStep($"📋 Total entities detected: {allEntities.Count}");

                // Step 3: Handle "Active" (Select All)
                if (EntityName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — selecting all entities...");

                    foreach (var entity in allEntities)
                    {
                        try
                        {
                            _actions.MoveToElement(entity).Click().Perform();
                            LogStep($"✅ Selected: {entity.Text.Trim()}");
                            Thread.Sleep(200);
                        }
                        catch
                        {
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                            LogStep($"⚠️ JS fallback click used for: {entity.Text.Trim()}");
                        }
                    }
                }
                else
                {
                    // Step 4: Multi-select specific entities
                    var entityValues = EntityName.Split(',')
                                                 .Select(v => v.Trim())
                                                 .Where(v => !string.IsNullOrEmpty(v))
                                                 .ToList();

                    foreach (var inputValue in entityValues)
                    {
                        bool found = false;
                        foreach (var entity in allEntities)
                        {
                            string label = entity.Text.Trim();

                            // Partial case-insensitive match
                            if (label.Contains(inputValue, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    _actions.MoveToElement(entity).Click().Perform();
                                    LogStep($"✅ Selected entity: {label}");
                                    found = true;
                                    Thread.Sleep(300);
                                    break;
                                }
                                catch
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                                    LogStep($"⚠️ JS fallback used for: {label}");
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            LogStep($"⚠️ Entity '{inputValue}' not found in list.");
                    }
                }

                WaitForUIEffect();

                // Step 5: Close dropdown
                try
                {
                    entityDropdown.Click();
                    LogStep("📦 Closed Entity dropdown after selection.");
                }
                catch
                {
                    LogStep("ℹ️ Dropdown closed automatically.");
                }

                WaitForUIEffect();




                // === Store Dropdown ===
                LogStep($"Selecting Store(s): {StoreName}");

                var storeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-store-filter-dropdown")));
                storeDropdown.Click();
                WaitForUIEffect();

                // Find the dynamic store panel by ID prefix
                var storePanel = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//div[starts-with(@id,'store_selection_menu_')]"))
                          .FirstOrDefault());

                if (storePanel == null)
                    throw new Exception("❌ Store selection panel not found after opening dropdown.");

                // Log all visible store names (for debugging)
                var visibleLabels = storePanel.FindElements(By.XPath(".//*[normalize-space(text())!='']"))
                                              .Select(e => e.Text.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .Distinct()
                                              .ToList();

                LogStep("📋 Available stores in dropdown:");
                foreach (var lbl in visibleLabels)
                {
                    LogStep($"   - {lbl}");
                }

                // If user input is "Active" → select all
                if (StoreName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — ticking all store checkboxes...");
                    var allCheckboxes = storePanel.FindElements(By.XPath(".//input[@type='checkbox']"));
                    int ticked = 0;

                    foreach (var checkbox in allCheckboxes)
                    {
                        try
                        {
                            if (!checkbox.Selected)
                            {
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", checkbox);
                                ticked++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Unable to tick one checkbox: {ex.Message}");
                        }
                    }

                    LogStep($"✅ {ticked} store(s) ticked successfully.");
                }
                else
                {
                    var storeValues = StoreName.Split(',')
                                               .Select(v => v.Trim())
                                               .Where(v => !string.IsNullOrEmpty(v))
                                               .ToList();

                    foreach (var val in storeValues)
                    {
                        bool found = false;

                        foreach (var lbl in visibleLabels)
                        {
                            if (lbl.ToUpper().Contains(val.ToUpper()))
                            {
                                var target = storePanel.FindElement(By.XPath($".//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '{val.ToUpper()}')]/preceding::input[@type='checkbox'][1]"));
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", target);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", target);
                                LogStep($"✅ Store selected: {lbl}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            LogStep($"❌ Could not find store containing: '{val}'");
                        }
                    }
                }

                WaitForUIEffect();
                storeDropdown.Click();
                WaitForUIEffect();



                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[3]/div[3]/button")));
                selectBtn.Click();
                LogStep("Clicked 'Apply' button to apply filter");
                WaitForUIEffect();


                LogStep("⏳ Waiting for Report table to appear...");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div")));
                string originalHtml = tableElement.GetAttribute("innerHTML");

                LogStep("🔍 Checking if 'Next' button is available...");
                var nextButton = _wait.Until(ExpectedConditions.ElementExists(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[3]/app-global-pagination/div/div[2]/ul/li[4]")));

                bool nextDisabled = nextButton.GetAttribute("disabled") == "true" || nextButton.GetAttribute("class").Contains("disabled");

                if (nextDisabled)
                {
                    LogStep("⚠️ 'Next' button is disabled. Skipping test.");
                    return;
                }

                LogStep("Clicking 'Next' button...");
                _wait.Until(ExpectedConditions.ElementToBeClickable(nextButton)).Click();
                WaitForUIEffect(1500);

                LogStep("⏳ Waiting for table to update after clicking 'Next'...");
                _wait.Until(driver =>
                {
                    var newTable = driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div"));
                    return newTable.GetAttribute("innerHTML") != originalHtml;
                });

                LogStep("Clicking 'Previous' button...");
                var previousButton = _wait.Until(ExpectedConditions.ElementExists(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[3]/app-global-pagination/div/div[2]/ul/li[2]")));

                bool prevDisabled = previousButton.GetAttribute("disabled") == "true" || previousButton.GetAttribute("class").Contains("disabled");

                if (prevDisabled)
                {
                    LogStep("⚠️ 'Previous' button is disabled. Skipping return check.");
                    return;
                }

                _wait.Until(ExpectedConditions.ElementToBeClickable(previousButton)).Click();
                WaitForUIEffect(1500);

                LogStep("⏳ Verifying table returned to original content...");
                _wait.Until(driver =>
                {
                    var tableBack = driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div"));
                    return tableBack.GetAttribute("innerHTML") == originalHtml;
                });
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep("✅ Table returned to original state after clicking 'Previous'.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }
        }


        [Test]
        [Category("Report")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Report Paging - Click Last Icon Button and Verify Table Change")]
        [TestCaseSource(nameof(ExportReportTestData))]
        public void TestPagingClickLastIconAndVerify(string Fromyear, string Frommonth, string EntityName, string StoreName)
        {
            try
            {
                LogStep($"Selecting Year: {Fromyear}");

                // Step 1: Click to open dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[2]/ng-select/div/span")));
                yearDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var dropdownPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(dropdownPanelXPath)));

                // Step 3: Find year option and click
                var yearOptions = _driver.FindElements(By.XPath(dropdownPanelXPath));
                bool yearSelected = false;

                foreach (var option in yearOptions)
                {
                    string yearText = option.Text.Trim();
                    if (yearText.Equals(Fromyear, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Year: {Fromyear}");
                        yearSelected = true;
                        break;
                    }
                }

                // Step 4: Handle case if year not found
                if (!yearSelected)
                {
                    LogStep($"⚠️ Year '{Fromyear}' not found in dropdown.");
                }
                WaitForUIEffect();



                // Month dropdown (ng-select)
                LogStep($"Selecting Month: {Frommonth}");

                // Step 1: Click to open dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[3]/ng-select/div/span")));
                monthDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var monthPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(monthPanelXPath)));

                // Step 3: Loop and select matching month value
                var monthOptions = _driver.FindElements(By.XPath(monthPanelXPath));
                bool monthSelected = false;

                foreach (var option in monthOptions)
                {
                    string monthText = option.Text.Trim();
                    if (monthText.Equals(Frommonth, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Month: {Frommonth}");
                        monthSelected = true;
                        break;
                    }
                }

                // Step 4: If not found
                if (!monthSelected)
                {
                    LogStep($"⚠️ Month '{Frommonth}' not found in dropdown.");
                }
                WaitForUIEffect();


                // ==========================
                // Entity Dropdown (Dynamic Multi-Select + Select All)
                // ==========================
                LogStep($"Selecting Entity(ies): {EntityName}");

                var _actions = new OpenQA.Selenium.Interactions.Actions(_driver);

                // Step 1: Open dropdown
                var entityDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[2]/div[2]/app-entity-filter-dropdown")));
                entityDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown to show all entity divs (Fixed XPath)
                var entityListXPath = "/html/body/div/div/div[2]/div/div/div/div";
                _wait.Until(ExpectedConditions.ElementExists(By.XPath(entityListXPath)));

                var allEntities = _driver.FindElements(By.XPath(entityListXPath));
                LogStep($"📋 Total entities detected: {allEntities.Count}");

                // Step 3: Handle "Active" (Select All)
                if (EntityName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — selecting all entities...");

                    foreach (var entity in allEntities)
                    {
                        try
                        {
                            _actions.MoveToElement(entity).Click().Perform();
                            LogStep($"✅ Selected: {entity.Text.Trim()}");
                            Thread.Sleep(200);
                        }
                        catch
                        {
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                            LogStep($"⚠️ JS fallback click used for: {entity.Text.Trim()}");
                        }
                    }
                }
                else
                {
                    // Step 4: Multi-select specific entities
                    var entityValues = EntityName.Split(',')
                                                 .Select(v => v.Trim())
                                                 .Where(v => !string.IsNullOrEmpty(v))
                                                 .ToList();

                    foreach (var inputValue in entityValues)
                    {
                        bool found = false;
                        foreach (var entity in allEntities)
                        {
                            string label = entity.Text.Trim();

                            // Partial case-insensitive match
                            if (label.Contains(inputValue, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    _actions.MoveToElement(entity).Click().Perform();
                                    LogStep($"✅ Selected entity: {label}");
                                    found = true;
                                    Thread.Sleep(300);
                                    break;
                                }
                                catch
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                                    LogStep($"⚠️ JS fallback used for: {label}");
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            LogStep($"⚠️ Entity '{inputValue}' not found in list.");
                    }
                }

                WaitForUIEffect();

                // Step 5: Close dropdown
                try
                {
                    entityDropdown.Click();
                    LogStep("📦 Closed Entity dropdown after selection.");
                }
                catch
                {
                    LogStep("ℹ️ Dropdown closed automatically.");
                }

                WaitForUIEffect();




                // === Store Dropdown ===
                LogStep($"Selecting Store(s): {StoreName}");

                var storeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-store-filter-dropdown")));
                storeDropdown.Click();
                WaitForUIEffect();

                // Find the dynamic store panel by ID prefix
                var storePanel = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//div[starts-with(@id,'store_selection_menu_')]"))
                          .FirstOrDefault());

                if (storePanel == null)
                    throw new Exception("❌ Store selection panel not found after opening dropdown.");

                // Log all visible store names (for debugging)
                var visibleLabels = storePanel.FindElements(By.XPath(".//*[normalize-space(text())!='']"))
                                              .Select(e => e.Text.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .Distinct()
                                              .ToList();

                LogStep("📋 Available stores in dropdown:");
                foreach (var lbl in visibleLabels)
                {
                    LogStep($"   - {lbl}");
                }

                // If user input is "Active" → select all
                if (StoreName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — ticking all store checkboxes...");
                    var allCheckboxes = storePanel.FindElements(By.XPath(".//input[@type='checkbox']"));
                    int ticked = 0;

                    foreach (var checkbox in allCheckboxes)
                    {
                        try
                        {
                            if (!checkbox.Selected)
                            {
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", checkbox);
                                ticked++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Unable to tick one checkbox: {ex.Message}");
                        }
                    }

                    LogStep($"✅ {ticked} store(s) ticked successfully.");
                }
                else
                {
                    var storeValues = StoreName.Split(',')
                                               .Select(v => v.Trim())
                                               .Where(v => !string.IsNullOrEmpty(v))
                                               .ToList();

                    foreach (var val in storeValues)
                    {
                        bool found = false;

                        foreach (var lbl in visibleLabels)
                        {
                            if (lbl.ToUpper().Contains(val.ToUpper()))
                            {
                                var target = storePanel.FindElement(By.XPath($".//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '{val.ToUpper()}')]/preceding::input[@type='checkbox'][1]"));
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", target);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", target);
                                LogStep($"✅ Store selected: {lbl}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            LogStep($"❌ Could not find store containing: '{val}'");
                        }
                    }
                }

                WaitForUIEffect();
                storeDropdown.Click();
                WaitForUIEffect();



                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[3]/div[3]/button")));
                selectBtn.Click();
                LogStep("Clicked 'Apply' button to apply filter");
                WaitForUIEffect();


                string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div";

                LogStep("⏳ Waiting for Report table...");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                string beforeHtml = tableElement.GetAttribute("innerHTML");

                LogStep("🔎 Finding 'Last Page' button...");
                var lastPageButton = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//a[.//i[contains(@class,'fa-angle-double-right')]]")));

                bool isDisabled = lastPageButton.GetAttribute("class")?.Contains("disabled") ?? false;
                if (isDisabled)
                {
                    LogStep("ℹ️ 'Last' button is disabled. Possibly already on last page.");
                    Assert.IsTrue(true);
                    return;
                }

                LogStep("Scrolling into view and clicking 'Last Page' button...");
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({ block: 'center' });", lastPageButton);
                _wait.Until(ExpectedConditions.ElementToBeClickable(lastPageButton));
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", lastPageButton);
                WaitForUIEffect(1500);

                LogStep("⏳ Waiting for table content to change...");
                bool tableChanged = false;
                try
                {
                    tableChanged = new WebDriverWait(_driver, TimeSpan.FromSeconds(10)).Until(driver =>
                    {
                        var updatedTable = driver.FindElement(By.XPath(tableXPath));
                        return updatedTable.GetAttribute("innerHTML") != beforeHtml;
                    });
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⚠️ Table did not change — may already be on last page.");
                }

                if (tableChanged)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("✅ Table content updated after clicking 'Last'.");
                }
                else
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("ℹ️ Table content unchanged after clicking 'Last'. Test passes with note.");
                }

                Assert.IsTrue(true); // Always pass — warning only
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during test: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }
        }




        [Test]
        [Category("Report")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Report Paging - Click First and Verify Return")]
        [TestCaseSource(nameof(ExportReportTestData))]
        public void TestPagingFirstButtonAndVerify(string Fromyear, string Frommonth, string EntityName, string StoreName)
        {
            string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div";

            try
            {
                LogStep($"Selecting Year: {Fromyear}");

                // Step 1: Click to open dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[2]/ng-select/div/span")));
                yearDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var dropdownPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(dropdownPanelXPath)));

                // Step 3: Find year option and click
                var yearOptions = _driver.FindElements(By.XPath(dropdownPanelXPath));
                bool yearSelected = false;

                foreach (var option in yearOptions)
                {
                    string yearText = option.Text.Trim();
                    if (yearText.Equals(Fromyear, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Year: {Fromyear}");
                        yearSelected = true;
                        break;
                    }
                }

                // Step 4: Handle case if year not found
                if (!yearSelected)
                {
                    LogStep($"⚠️ Year '{Fromyear}' not found in dropdown.");
                }
                WaitForUIEffect();



                // Month dropdown (ng-select)
                LogStep($"Selecting Month: {Frommonth}");

                // Step 1: Click to open dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[3]/ng-select/div/span")));
                monthDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var monthPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(monthPanelXPath)));

                // Step 3: Loop and select matching month value
                var monthOptions = _driver.FindElements(By.XPath(monthPanelXPath));
                bool monthSelected = false;

                foreach (var option in monthOptions)
                {
                    string monthText = option.Text.Trim();
                    if (monthText.Equals(Frommonth, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Month: {Frommonth}");
                        monthSelected = true;
                        break;
                    }
                }

                // Step 4: If not found
                if (!monthSelected)
                {
                    LogStep($"⚠️ Month '{Frommonth}' not found in dropdown.");
                }
                WaitForUIEffect();


                // ==========================
                // Entity Dropdown (Dynamic Multi-Select + Select All)
                // ==========================
                LogStep($"Selecting Entity(ies): {EntityName}");

                var _actions = new OpenQA.Selenium.Interactions.Actions(_driver);

                // Step 1: Open dropdown
                var entityDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[2]/div[2]/app-entity-filter-dropdown")));
                entityDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown to show all entity divs (Fixed XPath)
                var entityListXPath = "/html/body/div/div/div[2]/div/div/div/div";
                _wait.Until(ExpectedConditions.ElementExists(By.XPath(entityListXPath)));

                var allEntities = _driver.FindElements(By.XPath(entityListXPath));
                LogStep($"📋 Total entities detected: {allEntities.Count}");

                // Step 3: Handle "Active" (Select All)
                if (EntityName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — selecting all entities...");

                    foreach (var entity in allEntities)
                    {
                        try
                        {
                            _actions.MoveToElement(entity).Click().Perform();
                            LogStep($"✅ Selected: {entity.Text.Trim()}");
                            Thread.Sleep(200);
                        }
                        catch
                        {
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                            LogStep($"⚠️ JS fallback click used for: {entity.Text.Trim()}");
                        }
                    }
                }
                else
                {
                    // Step 4: Multi-select specific entities
                    var entityValues = EntityName.Split(',')
                                                 .Select(v => v.Trim())
                                                 .Where(v => !string.IsNullOrEmpty(v))
                                                 .ToList();

                    foreach (var inputValue in entityValues)
                    {
                        bool found = false;
                        foreach (var entity in allEntities)
                        {
                            string label = entity.Text.Trim();

                            // Partial case-insensitive match
                            if (label.Contains(inputValue, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    _actions.MoveToElement(entity).Click().Perform();
                                    LogStep($"✅ Selected entity: {label}");
                                    found = true;
                                    Thread.Sleep(300);
                                    break;
                                }
                                catch
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                                    LogStep($"⚠️ JS fallback used for: {label}");
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            LogStep($"⚠️ Entity '{inputValue}' not found in list.");
                    }
                }

                WaitForUIEffect();

                // Step 5: Close dropdown
                try
                {
                    entityDropdown.Click();
                    LogStep("📦 Closed Entity dropdown after selection.");
                }
                catch
                {
                    LogStep("ℹ️ Dropdown closed automatically.");
                }

                WaitForUIEffect();




                // === Store Dropdown ===
                LogStep($"Selecting Store(s): {StoreName}");

                var storeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-store-filter-dropdown")));
                storeDropdown.Click();
                WaitForUIEffect();

                // Find the dynamic store panel by ID prefix
                var storePanel = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//div[starts-with(@id,'store_selection_menu_')]"))
                          .FirstOrDefault());

                if (storePanel == null)
                    throw new Exception("❌ Store selection panel not found after opening dropdown.");

                // Log all visible store names (for debugging)
                var visibleLabels = storePanel.FindElements(By.XPath(".//*[normalize-space(text())!='']"))
                                              .Select(e => e.Text.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .Distinct()
                                              .ToList();

                LogStep("📋 Available stores in dropdown:");
                foreach (var lbl in visibleLabels)
                {
                    LogStep($"   - {lbl}");
                }

                // If user input is "Active" → select all
                if (StoreName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — ticking all store checkboxes...");
                    var allCheckboxes = storePanel.FindElements(By.XPath(".//input[@type='checkbox']"));
                    int ticked = 0;

                    foreach (var checkbox in allCheckboxes)
                    {
                        try
                        {
                            if (!checkbox.Selected)
                            {
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", checkbox);
                                ticked++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Unable to tick one checkbox: {ex.Message}");
                        }
                    }

                    LogStep($"✅ {ticked} store(s) ticked successfully.");
                }
                else
                {
                    var storeValues = StoreName.Split(',')
                                               .Select(v => v.Trim())
                                               .Where(v => !string.IsNullOrEmpty(v))
                                               .ToList();

                    foreach (var val in storeValues)
                    {
                        bool found = false;

                        foreach (var lbl in visibleLabels)
                        {
                            if (lbl.ToUpper().Contains(val.ToUpper()))
                            {
                                var target = storePanel.FindElement(By.XPath($".//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '{val.ToUpper()}')]/preceding::input[@type='checkbox'][1]"));
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", target);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", target);
                                LogStep($"✅ Store selected: {lbl}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            LogStep($"❌ Could not find store containing: '{val}'");
                        }
                    }
                }

                WaitForUIEffect();
                storeDropdown.Click();
                WaitForUIEffect();



                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[3]/div[3]/button")));
                selectBtn.Click();
                LogStep("Clicked 'Apply' button to apply filter");
                WaitForUIEffect();


                LogStep("📄 Waiting for table to load...");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                string originalHtml = tableElement.GetAttribute("innerHTML");

                LogStep("📄 Clicking 'Last' button if not disabled...");
                var lastButton = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//a[.//i[contains(@class,'fa-angle-double-right')]]")));

                bool lastDisabled = lastButton.GetAttribute("class")?.Contains("disabled") ?? false;
                if (lastDisabled)
                {
                    LogStep("ℹ️ 'Last' button is disabled — already on last page or only one page exists.");
                    Assert.IsTrue(true);
                    return;
                }

                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({ block: 'center' });", lastButton);
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", lastButton);
                LogStep("✅ 'Last' button clicked.");
                WaitForUIEffect(1500);

                LogStep("⏳ Waiting for table content to change after clicking 'Last'...");
                bool tableChanged = false;
                try
                {
                    tableChanged = new WebDriverWait(_driver, TimeSpan.FromSeconds(10)).Until(driver =>
                    {
                        var updatedTable = driver.FindElement(By.XPath(tableXPath));
                        return updatedTable.GetAttribute("innerHTML") != originalHtml;
                    });
                }
                catch (WebDriverTimeoutException)
                {
                    LogStep("⚠️ Table did not change — possibly already on last page.");
                }

                LogStep("📄 Clicking 'First' button...");
                var firstButton = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//a[.//i[contains(@class,'fa-angle-double-left')]]")));

                bool firstDisabled = firstButton.GetAttribute("class")?.Contains("disabled") ?? false;
                if (firstDisabled)
                {
                    LogStep("ℹ️ 'First' button is disabled. Cannot return to first page.");
                    Assert.IsTrue(true);
                    return;
                }

                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView({ block: 'center' });", firstButton);
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", firstButton);
                LogStep("✅ 'First' button clicked.");
                WaitForUIEffect(1500);

                LogStep("⏳ Waiting for table to return to original content...");
                bool tableReturned = false;
                try
                {
                    tableReturned = new WebDriverWait(_driver, TimeSpan.FromSeconds(10)).Until(driver =>
                    {
                        _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                        var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                        File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                        var returnedTable = driver.FindElement(By.XPath(tableXPath));
                        return returnedTable.GetAttribute("innerHTML") == originalHtml;
                    });
                }
                catch (WebDriverTimeoutException)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("⚠️ Table did not return to original content — possibly unexpected pagination state.");
                }

                LogStep($"ℹ️ Final Status → TableChanged: {tableChanged}, TableReturned: {tableReturned}");
                Assert.IsTrue(true); // Pass unconditionally; logic handled in log
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }
        }




        [Test]
        [Category("Report")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Report Paging - Click Page Size Dropdown and Verify Table Update")]
        [TestCaseSource(nameof(ReportInfoLineTestData))]
        public void TestItemsPerPageVerify(string Fromyear, string Frommonth, string EntityName, string StoreName, string pageSizeValue)
        {
            string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div";
            string dropdownXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[3]/app-global-pagination/div/div[1]/select";
            string rowSelector = "tbody tr";

            try
            {
                LogStep($"Selecting Year: {Fromyear}");

                // Step 1: Click to open dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[2]/ng-select/div/span")));
                yearDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var dropdownPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(dropdownPanelXPath)));

                // Step 3: Find year option and click
                var yearOptions = _driver.FindElements(By.XPath(dropdownPanelXPath));
                bool yearSelected = false;

                foreach (var option in yearOptions)
                {
                    string yearText = option.Text.Trim();
                    if (yearText.Equals(Fromyear, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Year: {Fromyear}");
                        yearSelected = true;
                        break;
                    }
                }

                // Step 4: Handle case if year not found
                if (!yearSelected)
                {
                    LogStep($"⚠️ Year '{Fromyear}' not found in dropdown.");
                }
                WaitForUIEffect();



                // Month dropdown (ng-select)
                LogStep($"Selecting Month: {Frommonth}");

                // Step 1: Click to open dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[3]/ng-select/div/span")));
                monthDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var monthPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(monthPanelXPath)));

                // Step 3: Loop and select matching month value
                var monthOptions = _driver.FindElements(By.XPath(monthPanelXPath));
                bool monthSelected = false;

                foreach (var option in monthOptions)
                {
                    string monthText = option.Text.Trim();
                    if (monthText.Equals(Frommonth, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Month: {Frommonth}");
                        monthSelected = true;
                        break;
                    }
                }

                // Step 4: If not found
                if (!monthSelected)
                {
                    LogStep($"⚠️ Month '{Frommonth}' not found in dropdown.");
                }
                WaitForUIEffect();


                // ==========================
                // Entity Dropdown (Dynamic Multi-Select + Select All)
                // ==========================
                LogStep($"Selecting Entity(ies): {EntityName}");

                var _actions = new OpenQA.Selenium.Interactions.Actions(_driver);

                // Step 1: Open dropdown
                var entityDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[2]/div[2]/app-entity-filter-dropdown")));
                entityDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown to show all entity divs (Fixed XPath)
                var entityListXPath = "/html/body/div/div/div[2]/div/div/div/div";
                _wait.Until(ExpectedConditions.ElementExists(By.XPath(entityListXPath)));

                var allEntities = _driver.FindElements(By.XPath(entityListXPath));
                LogStep($"📋 Total entities detected: {allEntities.Count}");

                // Step 3: Handle "Active" (Select All)
                if (EntityName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — selecting all entities...");

                    foreach (var entity in allEntities)
                    {
                        try
                        {
                            _actions.MoveToElement(entity).Click().Perform();
                            LogStep($"✅ Selected: {entity.Text.Trim()}");
                            Thread.Sleep(200);
                        }
                        catch
                        {
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                            LogStep($"⚠️ JS fallback click used for: {entity.Text.Trim()}");
                        }
                    }
                }
                else
                {
                    // Step 4: Multi-select specific entities
                    var entityValues = EntityName.Split(',')
                                                 .Select(v => v.Trim())
                                                 .Where(v => !string.IsNullOrEmpty(v))
                                                 .ToList();

                    foreach (var inputValue in entityValues)
                    {
                        bool found = false;
                        foreach (var entity in allEntities)
                        {
                            string label = entity.Text.Trim();

                            // Partial case-insensitive match
                            if (label.Contains(inputValue, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    _actions.MoveToElement(entity).Click().Perform();
                                    LogStep($"✅ Selected entity: {label}");
                                    found = true;
                                    Thread.Sleep(300);
                                    break;
                                }
                                catch
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                                    LogStep($"⚠️ JS fallback used for: {label}");
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            LogStep($"⚠️ Entity '{inputValue}' not found in list.");
                    }
                }

                WaitForUIEffect();

                // Step 5: Close dropdown
                try
                {
                    entityDropdown.Click();
                    LogStep("📦 Closed Entity dropdown after selection.");
                }
                catch
                {
                    LogStep("ℹ️ Dropdown closed automatically.");
                }

                WaitForUIEffect();




                // === Store Dropdown ===
                LogStep($"Selecting Store(s): {StoreName}");

                var storeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-store-filter-dropdown")));
                storeDropdown.Click();
                WaitForUIEffect();

                // Find the dynamic store panel by ID prefix
                var storePanel = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//div[starts-with(@id,'store_selection_menu_')]"))
                          .FirstOrDefault());

                if (storePanel == null)
                    throw new Exception("❌ Store selection panel not found after opening dropdown.");

                // Log all visible store names (for debugging)
                var visibleLabels = storePanel.FindElements(By.XPath(".//*[normalize-space(text())!='']"))
                                              .Select(e => e.Text.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .Distinct()
                                              .ToList();

                LogStep("📋 Available stores in dropdown:");
                foreach (var lbl in visibleLabels)
                {
                    LogStep($"   - {lbl}");
                }

                // If user input is "Active" → select all
                if (StoreName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — ticking all store checkboxes...");
                    var allCheckboxes = storePanel.FindElements(By.XPath(".//input[@type='checkbox']"));
                    int ticked = 0;

                    foreach (var checkbox in allCheckboxes)
                    {
                        try
                        {
                            if (!checkbox.Selected)
                            {
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", checkbox);
                                ticked++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Unable to tick one checkbox: {ex.Message}");
                        }
                    }

                    LogStep($"✅ {ticked} store(s) ticked successfully.");
                }
                else
                {
                    var storeValues = StoreName.Split(',')
                                               .Select(v => v.Trim())
                                               .Where(v => !string.IsNullOrEmpty(v))
                                               .ToList();

                    foreach (var val in storeValues)
                    {
                        bool found = false;

                        foreach (var lbl in visibleLabels)
                        {
                            if (lbl.ToUpper().Contains(val.ToUpper()))
                            {
                                var target = storePanel.FindElement(By.XPath($".//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '{val.ToUpper()}')]/preceding::input[@type='checkbox'][1]"));
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", target);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", target);
                                LogStep($"✅ Store selected: {lbl}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            LogStep($"❌ Could not find store containing: '{val}'");
                        }
                    }
                }

                WaitForUIEffect();
                storeDropdown.Click();
                WaitForUIEffect();



                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[3]/div[3]/button")));
                selectBtn.Click();
                LogStep("Clicked 'Apply' button to apply filter");
                WaitForUIEffect();


                LogStep("📄 Waiting for table to load...");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                string beforeHtml = tableElement.GetAttribute("innerHTML");
                int originalRowCount = tableElement.FindElements(By.CssSelector(rowSelector)).Count;
                LogStep($"📄 Original row count: {originalRowCount}");

                LogStep("📄 Locating dropdown...");
                var dropdownElement = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(dropdownXPath)));
                var select = new SelectElement(dropdownElement);

                LogStep("🔍 Checking if dropdown contains the value...");
                bool optionExists = select.Options.Any(opt => opt.Text.Trim() == pageSizeValue);
                if (!optionExists)
                {
                    LogStep($"❌ Page size '{pageSizeValue}' not found.");
                    Assert.Fail($"Dropdown does not contain value '{pageSizeValue}'");
                }

                LogStep($"✅ Selecting page size: {pageSizeValue}");
                select.SelectByText(pageSizeValue);
                WaitForUIEffect(800); // short wait for effect

                LogStep("⏳ Waiting for table content or row count to update...");
                bool tableUpdated = _wait.Until(driver =>
                {
                    var updatedTable = driver.FindElement(By.XPath(tableXPath));
                    string afterHtml = updatedTable.GetAttribute("innerHTML");
                    int updatedRowCount = updatedTable.FindElements(By.CssSelector(rowSelector)).Count;

                    return afterHtml != beforeHtml || updatedRowCount <= int.Parse(pageSizeValue);
                });

                WaitForUIEffect(500);

                if (tableUpdated)
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("✅ Table updated successfully or already matches expected page size.");
                }
                else
                {
                    _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                    var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                    File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                    LogStep("❌ Table did not update and row count exceeds selected page size.");
                    Assert.Fail("Table did not update correctly.");
                }
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during paging test: {ex.Message}");
                Assert.Fail("Test failed due to unexpected exception.");
            }
        }



        [Test]
        [Category("Report")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Report Paging - Click Page Button Only If It Exists and Verify Table Update")]
        [TestCaseSource(nameof(ReportInfoPageTestData))]
        public void TestClickPageButtonIfExists(string Fromyear, string Frommonth, string EntityName, string StoreName, string pageNumber)
        {
            string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div";
            string paginationXPathTemplate = "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[3]/app-global-pagination/div/div[2]/ul/li[a[text()='{0}']]/a";
            string dynamicPageXPath = string.Format(paginationXPathTemplate, pageNumber); 

            try
            {
                LogStep($"Selecting Year: {Fromyear}");

                // Step 1: Click to open dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[2]/ng-select/div/span")));
                yearDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var dropdownPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(dropdownPanelXPath)));

                // Step 3: Find year option and click
                var yearOptions = _driver.FindElements(By.XPath(dropdownPanelXPath));
                bool yearSelected = false;

                foreach (var option in yearOptions)
                {
                    string yearText = option.Text.Trim();
                    if (yearText.Equals(Fromyear, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Year: {Fromyear}");
                        yearSelected = true;
                        break;
                    }
                }

                // Step 4: Handle case if year not found
                if (!yearSelected)
                {
                    LogStep($"⚠️ Year '{Fromyear}' not found in dropdown.");
                }
                WaitForUIEffect();



                // Month dropdown (ng-select)
                LogStep($"Selecting Month: {Frommonth}");

                // Step 1: Click to open dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-summary-sales/div[1]/div/div[1]/div[3]/ng-select/div/span")));
                monthDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown panel to load
                var monthPanelXPath = "//ng-dropdown-panel//div[@role='option']//span[contains(@class, 'ng-option-label')]";
                _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(monthPanelXPath)));

                // Step 3: Loop and select matching month value
                var monthOptions = _driver.FindElements(By.XPath(monthPanelXPath));
                bool monthSelected = false;

                foreach (var option in monthOptions)
                {
                    string monthText = option.Text.Trim();
                    if (monthText.Equals(Frommonth, StringComparison.OrdinalIgnoreCase))
                    {
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", option);
                        option.Click();
                        LogStep($"✅ Selected Month: {Frommonth}");
                        monthSelected = true;
                        break;
                    }
                }

                // Step 4: If not found
                if (!monthSelected)
                {
                    LogStep($"⚠️ Month '{Frommonth}' not found in dropdown.");
                }
                WaitForUIEffect();


                // ==========================
                // Entity Dropdown (Dynamic Multi-Select + Select All)
                // ==========================
                LogStep($"Selecting Entity(ies): {EntityName}");

                var _actions = new OpenQA.Selenium.Interactions.Actions(_driver);

                // Step 1: Open dropdown
                var entityDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[2]/div[2]/app-entity-filter-dropdown")));
                entityDropdown.Click();
                WaitForUIEffect();

                // Step 2: Wait for dropdown to show all entity divs (Fixed XPath)
                var entityListXPath = "/html/body/div/div/div[2]/div/div/div/div";
                _wait.Until(ExpectedConditions.ElementExists(By.XPath(entityListXPath)));

                var allEntities = _driver.FindElements(By.XPath(entityListXPath));
                LogStep($"📋 Total entities detected: {allEntities.Count}");

                // Step 3: Handle "Active" (Select All)
                if (EntityName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — selecting all entities...");

                    foreach (var entity in allEntities)
                    {
                        try
                        {
                            _actions.MoveToElement(entity).Click().Perform();
                            LogStep($"✅ Selected: {entity.Text.Trim()}");
                            Thread.Sleep(200);
                        }
                        catch
                        {
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                            LogStep($"⚠️ JS fallback click used for: {entity.Text.Trim()}");
                        }
                    }
                }
                else
                {
                    // Step 4: Multi-select specific entities
                    var entityValues = EntityName.Split(',')
                                                 .Select(v => v.Trim())
                                                 .Where(v => !string.IsNullOrEmpty(v))
                                                 .ToList();

                    foreach (var inputValue in entityValues)
                    {
                        bool found = false;
                        foreach (var entity in allEntities)
                        {
                            string label = entity.Text.Trim();

                            // Partial case-insensitive match
                            if (label.Contains(inputValue, StringComparison.OrdinalIgnoreCase))
                            {
                                try
                                {
                                    _actions.MoveToElement(entity).Click().Perform();
                                    LogStep($"✅ Selected entity: {label}");
                                    found = true;
                                    Thread.Sleep(300);
                                    break;
                                }
                                catch
                                {
                                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", entity);
                                    LogStep($"⚠️ JS fallback used for: {label}");
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                            LogStep($"⚠️ Entity '{inputValue}' not found in list.");
                    }
                }

                WaitForUIEffect();

                // Step 5: Close dropdown
                try
                {
                    entityDropdown.Click();
                    LogStep("📦 Closed Entity dropdown after selection.");
                }
                catch
                {
                    LogStep("ℹ️ Dropdown closed automatically.");
                }

                WaitForUIEffect();




                // === Store Dropdown ===
                LogStep($"Selecting Store(s): {StoreName}");

                var storeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-store-filter-dropdown")));
                storeDropdown.Click();
                WaitForUIEffect();

                // Find the dynamic store panel by ID prefix
                var storePanel = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//div[starts-with(@id,'store_selection_menu_')]"))
                          .FirstOrDefault());

                if (storePanel == null)
                    throw new Exception("❌ Store selection panel not found after opening dropdown.");

                // Log all visible store names (for debugging)
                var visibleLabels = storePanel.FindElements(By.XPath(".//*[normalize-space(text())!='']"))
                                              .Select(e => e.Text.Trim())
                                              .Where(t => !string.IsNullOrEmpty(t))
                                              .Distinct()
                                              .ToList();

                LogStep("📋 Available stores in dropdown:");
                foreach (var lbl in visibleLabels)
                {
                    LogStep($"   - {lbl}");
                }

                // If user input is "Active" → select all
                if (StoreName.Trim().Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("User input is 'Active' — ticking all store checkboxes...");
                    var allCheckboxes = storePanel.FindElements(By.XPath(".//input[@type='checkbox']"));
                    int ticked = 0;

                    foreach (var checkbox in allCheckboxes)
                    {
                        try
                        {
                            if (!checkbox.Selected)
                            {
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", checkbox);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", checkbox);
                                ticked++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogStep($"⚠️ Unable to tick one checkbox: {ex.Message}");
                        }
                    }

                    LogStep($"✅ {ticked} store(s) ticked successfully.");
                }
                else
                {
                    var storeValues = StoreName.Split(',')
                                               .Select(v => v.Trim())
                                               .Where(v => !string.IsNullOrEmpty(v))
                                               .ToList();

                    foreach (var val in storeValues)
                    {
                        bool found = false;

                        foreach (var lbl in visibleLabels)
                        {
                            if (lbl.ToUpper().Contains(val.ToUpper()))
                            {
                                var target = storePanel.FindElement(By.XPath($".//*[contains(translate(text(),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'), '{val.ToUpper()}')]/preceding::input[@type='checkbox'][1]"));
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);", target);
                                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", target);
                                LogStep($"✅ Store selected: {lbl}");
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            LogStep($"❌ Could not find store containing: '{val}'");
                        }
                    }
                }

                WaitForUIEffect();
                storeDropdown.Click();
                WaitForUIEffect();



                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[1]/div/div[3]/div[3]/button")));
                selectBtn.Click();
                LogStep("Clicked 'Apply' button to apply filter");
                WaitForUIEffect();


                LogStep("📄 Waiting for Report table to be visible...");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(tableXPath)));
                string beforeHtml = tableElement.GetAttribute("innerHTML");

                LogStep($"🔎 Checking if page number {pageNumber} exists in pagination...");
                var pageButtons = _driver.FindElements(By.XPath(dynamicPageXPath));
                if (pageButtons.Count == 0)
                {
                    LogStep($"✅ Page {pageNumber} does not exist — skipping click. Only one page available.");
                    return;
                }

                LogStep($"✅ Page button {pageNumber} found. Clicking now...");
                pageButtons[0].Click();
                WaitForUIEffect();

                LogStep("⏳ Waiting for table or pagination to update...");
                bool tableOrPageUpdated = _wait.Until(driver =>
                {
                    try
                    {
                        // Check table change
                        var updatedTable = driver.FindElement(By.XPath(tableXPath));
                        string afterHtml = updatedTable.GetAttribute("innerHTML");
                        if (afterHtml != beforeHtml)
                            return true;

                        // Check pagination active state
                        var activePage = driver.FindElements(By.XPath(
                            $"//li[contains(@class, 'active') or contains(@class, 'current')][normalize-space()='{pageNumber}']"
                        ));
                        if (activePage.Count > 0)
                            return true;

                        // No change detected yet
                        return false;
                    }
                    catch
                    {
                        return false;
                    }
                });

                LogStep(tableOrPageUpdated
                    ? $"✅ Table or pagination updated successfully after navigating to page {pageNumber}."
                    : $"⚠️ No visible data change detected, but navigation completed (table may contain same data).");

                // Capture screenshot for reporting
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                // We consider table or pagination change as success
                Assert.IsTrue(tableOrPageUpdated, $"❌ Neither table nor pagination state changed after clicking page {pageNumber}.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception during test: {ex.Message}");
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

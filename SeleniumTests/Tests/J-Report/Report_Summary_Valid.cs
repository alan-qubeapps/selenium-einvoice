using Allure.Net.Commons;
using Allure.NUnit;
using Allure.NUnit.Attributes;
using AngleSharp.Dom;
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
using SeleniumTests.Pages.Stores;
using SeleniumTests.Pages.Report;
using SeleniumTests.Pages.User;
using System.Drawing;
using System.Globalization;
using System.Media;
using System.Xml.Linq;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;
using System.Security.Cryptography.X509Certificates;

namespace SeleniumTests.Tests.J_Report
{

    public static class ExcelDataReaderReportSummaryValid
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



    }
        
    [TestFixture, Order(38)]
    [AllureNUnit]
    [AllureSuite("Report - Valid")]
    [AllureEpic("ERP-117")]
    public class ReportSummary_Valid
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


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "SReportTestDataValid.xlsx");

        public static IEnumerable<object[]> ExportReportTestData =>
        ExcelDataReaderReportSummaryValid.GetExportSummaryReportTestData(ExcelPath, "ExportReportTestData");



        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Summary Report Page";

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

            _moduleName = "Summary Report Page";
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
        [AllureStory("Create - Filter Report")]
        [TestCaseSource(nameof(ExportReportTestData))]
        public void FilterSummaryReport(string Fromyear, string Frommonth, string EntityName, string StoreName)
        {
            // Year dropdown (ng-select)
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



            // ✅ Verify if Sales Summary table has data or “No Data Available.”
            LogStep("Verifying Sales Summary table result after applying filters...");
            WaitForUIEffect();


            try
            {
                _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div")));

                var noDataElements = _driver.FindElements(By.XPath(
                    "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[1]/div/table/tbody[2]/tr/td/p"));

                // Take screenshot before verification
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"SummaryReport_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot captured: {_lastScreenshotPath}");


                if (noDataElements.Count > 0 &&
                    noDataElements[0].Text.Trim().Equals("No data available", StringComparison.OrdinalIgnoreCase))
                {
                    LogStep("✅ No Data Available message displayed — filter applied successfully with no matching records.");
                    Assert.IsTrue(true, "Filter applied successfully — no data found as expected.");
                }
                else
                {
                    var dataRows = _driver.FindElements(By.XPath(
                        "/html/body/app-layout/div[1]/div/div/div/app-content/app-summary-sales/div[2]/div/div[2]/div/table/tbody/tr"));

                    if (dataRows.Count > 0)
                    {
                        LogStep($"✅ Table displayed with {dataRows.Count} record(s) after filter applied.");
                        Assert.IsTrue(true,"Table data displayed successfully after applying filter.");
                    }
                    else
                    {
                        LogStep("❌ Table did not display any data or 'No Data Available' message.");
                        Assert.Fail("Table verification failed — neither data nor message appeared.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogStep($"❌ Exception during table verification: {ex.Message}");
                // Take screenshot before verification
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"SummaryReport_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot captured: {_lastScreenshotPath}");
                Assert.Fail($"Table verification failed due to exception: {ex.Message}");
            }
        }

        [Test]
        [Category("Report")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create - Export Report")]
        [TestCaseSource(nameof(ExportReportTestData))]
        public void ExportSummaryReport(string Fromyear, string Frommonth, string EntityName, string StoreName)
        {
            string downloadPath = AppConfig.DownloadPath;
            string filePrefix = "Summary Sales";

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


            WaitForUIEffect();
            LogStep("Clicking Export Button");
            helperFunction.WaitForElementToBeClickable(_wait,
                By.CssSelector("#kt_content_container > app-summary-sales > div.card.py-8.ps-8.pe-6.mt-5 > div > div.d-flex.justify-content-end.ng-star-inserted > div > a"));
            _ReportPage.ClickExportSummaryButton();
            WaitForUIEffect(100);

            _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Report_{DateTime.Now:yyyyMMdd_HHmmss}.png");
            var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
            File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
            // Step: Check for file download or modal message
            LogStep("Waiting for file download or processing modal...");

            bool fileDownloaded = _ReportPage.WaitForFileDownload(downloadPath, filePrefix, TimeSpan.FromSeconds(15));
            bool modalFound = false;
            string modalText = string.Empty;

            try
            {
                // Look for modal
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/div/div/div[2]")));
                modalText = modal.Text.Trim();

                if (modalText.Contains("The Transaction export file is under processing due to large size"))
                {
                    modalFound = true;
                    LogStep($"✅ Modal appeared: {modalText}");
                }
            }
            catch (WebDriverTimeoutException)
            {
                LogStep("ℹ️ No modal detected within timeout.");
            }

            // Decide test result
            if (fileDownloaded)
            {
                LogStep("✅ File downloaded successfully.");
                Assert.IsTrue(true, "File downloaded successfully.");
            }
            else if (modalFound)
            {
                LogStep("✅ Modal message displayed instead of download.");
                Assert.IsTrue(true, $"Modal message handled: {modalText}");
            }
            else
            {
                LogStep("❌ Neither file download nor modal message detected.");
                Assert.Fail("❌ Test failed: No download and no modal message.");
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

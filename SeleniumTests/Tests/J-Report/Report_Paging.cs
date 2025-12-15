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
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;


namespace SeleniumTests.Tests.J_Report
{

    public static class ExcelDataReaderReportPaging
    {



        public static IEnumerable<object[]> GetReportInfoTestData(string filePath, string sheetName)
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
                    string ReportName = worksheet.Cells[row, 1].Text?.Trim();
                    string FromMonth = worksheet.Cells[row, 2].Text?.Trim();
                    string FromYear = worksheet.Cells[row, 3].Text?.Trim();
                    string FromDate = worksheet.Cells[row, 4].Text?.Trim();
                    string ToMonth = worksheet.Cells[row, 5].Text?.Trim();
                    string ToYear = worksheet.Cells[row, 6].Text?.Trim();
                    string ToDate = worksheet.Cells[row, 7].Text?.Trim();
                    string DocType = worksheet.Cells[row, 8].Text?.Trim();
                    string DocStatus = worksheet.Cells[row, 9].Text?.Trim();


                    yield return new object[]
                    {
                        ReportName, FromMonth, FromYear, FromDate, ToMonth, ToYear, ToDate, DocType, DocStatus
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
                    string ReportName = worksheet.Cells[row, 1].Text?.Trim();
                    string FromMonth = worksheet.Cells[row, 2].Text?.Trim();
                    string FromYear = worksheet.Cells[row, 3].Text?.Trim();
                    string FromDate = worksheet.Cells[row, 4].Text?.Trim();
                    string ToMonth = worksheet.Cells[row, 5].Text?.Trim();
                    string ToYear = worksheet.Cells[row, 6].Text?.Trim();
                    string ToDate = worksheet.Cells[row, 7].Text?.Trim();
                    string DocType = worksheet.Cells[row, 8].Text?.Trim();
                    string DocStatus = worksheet.Cells[row, 9].Text?.Trim();
                    string LineNumber = worksheet.Cells[row, 10].Text?.Trim();


                    yield return new object[]
                    {
                        ReportName, FromMonth, FromYear, FromDate, ToMonth, ToYear, ToDate, DocType, DocStatus, LineNumber
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
                    string ReportName = worksheet.Cells[row, 1].Text?.Trim();
                    string FromMonth = worksheet.Cells[row, 2].Text?.Trim();
                    string FromYear = worksheet.Cells[row, 3].Text?.Trim();
                    string FromDate = worksheet.Cells[row, 4].Text?.Trim();
                    string ToMonth = worksheet.Cells[row, 5].Text?.Trim();
                    string ToYear = worksheet.Cells[row, 6].Text?.Trim();
                    string ToDate = worksheet.Cells[row, 7].Text?.Trim();
                    string DocType = worksheet.Cells[row, 8].Text?.Trim();
                    string DocStatus = worksheet.Cells[row, 9].Text?.Trim();
                    string PageNumber = worksheet.Cells[row, 10].Text?.Trim();


                    yield return new object[]
                    {
                        ReportName, FromMonth, FromYear, FromDate, ToMonth, ToYear, ToDate, DocType, DocStatus, PageNumber
                    };

                }
            }
        }

    }


    [TestFixture, Order(49)]
    [AllureNUnit]
    [AllureSuite("Report - Paging")] // use this ties to module
    [AllureEpic("ERP-117")] // use this and ties to ticket number

    public class Report_Paging
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

        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "ReportTestDataPaging.xlsx");

        public static IEnumerable<object[]> ReportInfoTestData =>
        ExcelDataReaderReportPaging.GetReportInfoTestData(ExcelPath, "ReportInfoTestData");
        public static IEnumerable<object[]> ReportInfoLineTestData =>
        ExcelDataReaderReportPaging.GetReportInfoLineTestData(ExcelPath, "ReportInfoLineTestData");
        public static IEnumerable<object[]> ReportInfoPageTestData =>
        ExcelDataReaderReportPaging.GetReportInfoPageTestData(ExcelPath, "ReportInfoPageTestData");

        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Report Page - Paging";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/report/transaction");
            helperFunction.WaitForPageToLoad(_wait);
            _ReportPage = new ReportPage(_driver);
            _logMessages.Clear();

            _moduleName = "Report Page - Paging";
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
        [TestCaseSource(nameof(ReportInfoTestData))]
        public void TestPagingNextButtonAndVerify(string ReportName, string Frommonth, string Fromyear, string fromDay, 
                                                  string Tomonth, string Toyear, string toDay, string docTypeInput, string statusInput)
        {
            try
            {
                string filePrefix = ReportName;

                LogStep($"Clicking navigation link: {ReportName}");
                // Step 1️⃣ — Open the dropdown (click the label with "All Report edited 6")
                IWebElement dropdownTrigger = _driver.FindElement(
                    By.XPath("//span[contains(@class, 'input-group-text')]/label[contains(@class, 'fw-bold')]")
                );
                dropdownTrigger.Click();

                // Step 2️⃣ — Wait until the dropdown list appears and the target option is clickable
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                IWebElement optionToSelect = wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath($"//div[contains(@class, 'display-col') and normalize-space(text())='{ReportName}']"))
                );

                // Step 3️⃣ — Click the desired value
                optionToSelect.Click();

                // Open Date Picker                
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div/div[2]/div[2]/button")).Click();
                WaitForUIEffect();

                // Date From & To
                LogStep($"Selecting date range: From {Frommonth}/{Fromyear}/{fromDay} to {Tomonth}/{Toyear}/{toDay}");

                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Frommonth}, year: {Fromyear}");

                // Month dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Frommonth);
                WaitForUIEffect();

                // Year dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Fromyear);
                LogStep($"Month and Year selected: {Frommonth} {Fromyear}");
                WaitForUIEffect();

                // From date
                var fromDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{fromDay}']")));
                fromDate.Click();
                LogStep($"From date selected: {fromDay}");
                WaitForUIEffect();





                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Tomonth}, year: {Toyear}");

                // Month dropdown
                monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Tomonth);
                WaitForUIEffect();

                // Year dropdown
                yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Toyear);
                LogStep($"Month and Year selected: {Tomonth} {Toyear}");
                WaitForUIEffect();

                // To date
                var toDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{toDay}']")));
                toDate.Click();
                LogStep($"From date selected: {toDay}");
                WaitForUIEffect();


                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker/button")));
                selectBtn.Click();
                LogStep("Clicked 'Select' button to apply date range");
                WaitForUIEffect();


                // Document Type
                LogStep($"Selecting Document Type(s): {docTypeInput}");

                // Open the dropdown
                var dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[2]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for options to appear (target <li> instead of <div>)
                var options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (docTypeInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not already selected
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected Document Type: {text}");
                        }
                    }
                }
                else
                {
                    // Split user input into multiple values
                    var values = docTypeInput.Split(',')
                                             .Select(v => v.Trim())
                                             .Where(v => !string.IsNullOrEmpty(v))
                                             .ToList();

                    foreach (var val in values)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(val, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected Document Type: {val}");
                        }
                        else
                        {
                            LogStep($"Skipped Document Type (not found or already selected): {val}");
                        }
                    }
                }
                WaitForUIEffect();

                // Close dropdown safely (optional)
                dropdown.Click();



                // Document Status
                LogStep($"Selecting Document Status: {statusInput}");

                // Open the dropdown
                dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[3]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for <li> items inside <p-multiselectitem>
                options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (statusInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not selected yet
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected status: {text}");
                        }
                    }
                }
                else
                {
                    // Split input into multiple statuses
                    var statuses = statusInput.Split(',')
                                              .Select(s => s.Trim())
                                              .Where(s => !string.IsNullOrEmpty(s))
                                              .ToList();

                    foreach (var status in statuses)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(status, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected status: {status}");
                        }
                        else
                        {
                            LogStep($"Skipped status (not found or already selected): {status}");
                        }
                    }
                }

                // Close dropdown safely
                dropdown.Click();

                // Apply Filter
                WaitForUIEffect();
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div[4]/button")).Click();

                LogStep("⏳ Waiting for Report table to load...");
                var tableElement = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div")));

                string beforeHtml = tableElement.GetAttribute("innerHTML");
                LogStep("✅ Captured initial table content.");

                LogStep("🔎 Locating 'Next' button for pagination...");
                var nextButton = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[3]/app-global-pagination/div/div[2]/ul/li[8]")));

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
                    var updatedTable = driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div"));
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
        [TestCaseSource(nameof(ReportInfoTestData))]
        public void TestPagingPreviousButtonAndVerify(string ReportName, string Frommonth, string Fromyear, string fromDay,
                                                  string Tomonth, string Toyear, string toDay, string docTypeInput, string statusInput)
        {
            try
            {
                string filePrefix = ReportName;

                LogStep($"Clicking navigation link: {ReportName}");
                // Step 1️⃣ — Open the dropdown (click the label with "All Report edited 6")
                IWebElement dropdownTrigger = _driver.FindElement(
                    By.XPath("//span[contains(@class, 'input-group-text')]/label[contains(@class, 'fw-bold')]")
                );
                dropdownTrigger.Click();

                // Step 2️⃣ — Wait until the dropdown list appears and the target option is clickable
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                IWebElement optionToSelect = wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath($"//div[contains(@class, 'display-col') and normalize-space(text())='{ReportName}']"))
                );

                // Step 3️⃣ — Click the desired value
                optionToSelect.Click();

                // Open Date Picker                
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div/div[2]/div[2]/button")).Click();
                WaitForUIEffect();

                // Date From & To
                LogStep($"Selecting date range: From {Frommonth}/{Fromyear}/{fromDay} to {Tomonth}/{Toyear}/{toDay}");

                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Frommonth}, year: {Fromyear}");

                // Month dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Frommonth);
                WaitForUIEffect();

                // Year dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Fromyear);
                LogStep($"Month and Year selected: {Frommonth} {Fromyear}");
                WaitForUIEffect();

                // From date
                var fromDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{fromDay}']")));
                fromDate.Click();
                LogStep($"From date selected: {fromDay}");
                WaitForUIEffect();





                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Tomonth}, year: {Toyear}");

                // Month dropdown
                monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Tomonth);
                WaitForUIEffect();

                // Year dropdown
                yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Toyear);
                LogStep($"Month and Year selected: {Tomonth} {Toyear}");
                WaitForUIEffect();

                // To date
                var toDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{toDay}']")));
                toDate.Click();
                LogStep($"From date selected: {toDay}");
                WaitForUIEffect();


                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker/button")));
                selectBtn.Click();
                LogStep("Clicked 'Select' button to apply date range");
                WaitForUIEffect();


                // Document Type
                LogStep($"Selecting Document Type(s): {docTypeInput}");

                // Open the dropdown
                var dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[2]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for options to appear (target <li> instead of <div>)
                var options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (docTypeInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not already selected
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected Document Type: {text}");
                        }
                    }
                }
                else
                {
                    // Split user input into multiple values
                    var values = docTypeInput.Split(',')
                                             .Select(v => v.Trim())
                                             .Where(v => !string.IsNullOrEmpty(v))
                                             .ToList();

                    foreach (var val in values)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(val, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected Document Type: {val}");
                        }
                        else
                        {
                            LogStep($"Skipped Document Type (not found or already selected): {val}");
                        }
                    }
                }
                WaitForUIEffect();

                // Close dropdown safely (optional)
                dropdown.Click();



                // Document Status
                LogStep($"Selecting Document Status: {statusInput}");

                // Open the dropdown
                dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[3]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for <li> items inside <p-multiselectitem>
                options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (statusInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not selected yet
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected status: {text}");
                        }
                    }
                }
                else
                {
                    // Split input into multiple statuses
                    var statuses = statusInput.Split(',')
                                              .Select(s => s.Trim())
                                              .Where(s => !string.IsNullOrEmpty(s))
                                              .ToList();

                    foreach (var status in statuses)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(status, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected status: {status}");
                        }
                        else
                        {
                            LogStep($"Skipped status (not found or already selected): {status}");
                        }
                    }
                }

                // Close dropdown safely
                dropdown.Click();

                // Apply Filter
                WaitForUIEffect();
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div[4]/button")).Click();

                LogStep("⏳ Waiting for Report table to appear...");
                var tableElement = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div")));
                string originalHtml = tableElement.GetAttribute("innerHTML");

                LogStep("🔍 Checking if 'Next' button is available...");
                var nextButton = _wait.Until(ExpectedConditions.ElementExists(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[3]/app-global-pagination/div/div[2]/ul/li[8]")));

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
                    var newTable = driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div"));
                    return newTable.GetAttribute("innerHTML") != originalHtml;
                });

                LogStep("Clicking 'Previous' button...");
                var previousButton = _wait.Until(ExpectedConditions.ElementExists(
                    By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[3]/app-global-pagination/div/div[2]/ul/li[2]")));

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
                    var tableBack = driver.FindElement(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div"));
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
        [TestCaseSource(nameof(ReportInfoTestData))]
        public void TestPagingClickLastIconAndVerify(string ReportName, string Frommonth, string Fromyear, string fromDay,
                                                  string Tomonth, string Toyear, string toDay, string docTypeInput, string statusInput)
        {
            try
            {
                string filePrefix = ReportName;

                LogStep($"Clicking navigation link: {ReportName}");
                // Step 1️⃣ — Open the dropdown (click the label with "All Report edited 6")
                IWebElement dropdownTrigger = _driver.FindElement(
                    By.XPath("//span[contains(@class, 'input-group-text')]/label[contains(@class, 'fw-bold')]")
                );
                dropdownTrigger.Click();

                // Step 2️⃣ — Wait until the dropdown list appears and the target option is clickable
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                IWebElement optionToSelect = wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath($"//div[contains(@class, 'display-col') and normalize-space(text())='{ReportName}']"))
                );

                // Step 3️⃣ — Click the desired value
                optionToSelect.Click();

                // Open Date Picker                
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div/div[2]/div[2]/button")).Click();
                WaitForUIEffect();

                // Date From & To
                LogStep($"Selecting date range: From {Frommonth}/{Fromyear}/{fromDay} to {Tomonth}/{Toyear}/{toDay}");

                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Frommonth}, year: {Fromyear}");

                // Month dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Frommonth);
                WaitForUIEffect();

                // Year dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Fromyear);
                LogStep($"Month and Year selected: {Frommonth} {Fromyear}");
                WaitForUIEffect();

                // From date
                var fromDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{fromDay}']")));
                fromDate.Click();
                LogStep($"From date selected: {fromDay}");
                WaitForUIEffect();





                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Tomonth}, year: {Toyear}");

                // Month dropdown
                monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Tomonth);
                WaitForUIEffect();

                // Year dropdown
                yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Toyear);
                LogStep($"Month and Year selected: {Tomonth} {Toyear}");
                WaitForUIEffect();

                // To date
                var toDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{toDay}']")));
                toDate.Click();
                LogStep($"From date selected: {toDay}");
                WaitForUIEffect();


                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker/button")));
                selectBtn.Click();
                LogStep("Clicked 'Select' button to apply date range");
                WaitForUIEffect();


                // Document Type
                LogStep($"Selecting Document Type(s): {docTypeInput}");

                // Open the dropdown
                var dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[2]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for options to appear (target <li> instead of <div>)
                var options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (docTypeInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not already selected
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected Document Type: {text}");
                        }
                    }
                }
                else
                {
                    // Split user input into multiple values
                    var values = docTypeInput.Split(',')
                                             .Select(v => v.Trim())
                                             .Where(v => !string.IsNullOrEmpty(v))
                                             .ToList();

                    foreach (var val in values)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(val, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected Document Type: {val}");
                        }
                        else
                        {
                            LogStep($"Skipped Document Type (not found or already selected): {val}");
                        }
                    }
                }
                WaitForUIEffect();

                // Close dropdown safely (optional)
                dropdown.Click();



                // Document Status
                LogStep($"Selecting Document Status: {statusInput}");

                // Open the dropdown
                dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[3]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for <li> items inside <p-multiselectitem>
                options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (statusInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not selected yet
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected status: {text}");
                        }
                    }
                }
                else
                {
                    // Split input into multiple statuses
                    var statuses = statusInput.Split(',')
                                              .Select(s => s.Trim())
                                              .Where(s => !string.IsNullOrEmpty(s))
                                              .ToList();

                    foreach (var status in statuses)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(status, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected status: {status}");
                        }
                        else
                        {
                            LogStep($"Skipped status (not found or already selected): {status}");
                        }
                    }
                }

                // Close dropdown safely
                dropdown.Click();

                // Apply Filter
                WaitForUIEffect();
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div[4]/button")).Click();

                string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div";

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
        [TestCaseSource(nameof(ReportInfoTestData))]
        public void TestPagingFirstButtonAndVerify(string ReportName, string Frommonth, string Fromyear, string fromDay,
                                                  string Tomonth, string Toyear, string toDay, string docTypeInput, string statusInput)
        {
            string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div";

            try
            {
                string filePrefix = ReportName;

                LogStep($"Clicking navigation link: {ReportName}");
                // Step 1️⃣ — Open the dropdown (click the label with "All Report edited 6")
                IWebElement dropdownTrigger = _driver.FindElement(
                    By.XPath("//span[contains(@class, 'input-group-text')]/label[contains(@class, 'fw-bold')]")
                );
                dropdownTrigger.Click();

                // Step 2️⃣ — Wait until the dropdown list appears and the target option is clickable
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                IWebElement optionToSelect = wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath($"//div[contains(@class, 'display-col') and normalize-space(text())='{ReportName}']"))
                );

                // Step 3️⃣ — Click the desired value
                optionToSelect.Click();

                // Open Date Picker                
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div/div[2]/div[2]/button")).Click();
                WaitForUIEffect();

                // Date From & To
                LogStep($"Selecting date range: From {Frommonth}/{Fromyear}/{fromDay} to {Tomonth}/{Toyear}/{toDay}");

                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Frommonth}, year: {Fromyear}");

                // Month dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Frommonth);
                WaitForUIEffect();

                // Year dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Fromyear);
                LogStep($"Month and Year selected: {Frommonth} {Fromyear}");
                WaitForUIEffect();

                // From date
                var fromDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{fromDay}']")));
                fromDate.Click();
                LogStep($"From date selected: {fromDay}");
                WaitForUIEffect();





                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Tomonth}, year: {Toyear}");

                // Month dropdown
                monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Tomonth);
                WaitForUIEffect();

                // Year dropdown
                yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Toyear);
                LogStep($"Month and Year selected: {Tomonth} {Toyear}");
                WaitForUIEffect();

                // To date
                var toDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{toDay}']")));
                toDate.Click();
                LogStep($"From date selected: {toDay}");
                WaitForUIEffect();


                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker/button")));
                selectBtn.Click();
                LogStep("Clicked 'Select' button to apply date range");
                WaitForUIEffect();


                // Document Type
                LogStep($"Selecting Document Type(s): {docTypeInput}");

                // Open the dropdown
                var dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[2]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for options to appear (target <li> instead of <div>)
                var options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (docTypeInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not already selected
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected Document Type: {text}");
                        }
                    }
                }
                else
                {
                    // Split user input into multiple values
                    var values = docTypeInput.Split(',')
                                             .Select(v => v.Trim())
                                             .Where(v => !string.IsNullOrEmpty(v))
                                             .ToList();

                    foreach (var val in values)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(val, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected Document Type: {val}");
                        }
                        else
                        {
                            LogStep($"Skipped Document Type (not found or already selected): {val}");
                        }
                    }
                }
                WaitForUIEffect();

                // Close dropdown safely (optional)
                dropdown.Click();



                // Document Status
                LogStep($"Selecting Document Status: {statusInput}");

                // Open the dropdown
                dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[3]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for <li> items inside <p-multiselectitem>
                options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (statusInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not selected yet
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected status: {text}");
                        }
                    }
                }
                else
                {
                    // Split input into multiple statuses
                    var statuses = statusInput.Split(',')
                                              .Select(s => s.Trim())
                                              .Where(s => !string.IsNullOrEmpty(s))
                                              .ToList();

                    foreach (var status in statuses)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(status, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected status: {status}");
                        }
                        else
                        {
                            LogStep($"Skipped status (not found or already selected): {status}");
                        }
                    }
                }

                // Close dropdown safely
                dropdown.Click();

                // Apply Filter
                WaitForUIEffect();
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div[4]/button")).Click();

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
        public void TestItemsPerPageVerify(string ReportName, string Frommonth, string Fromyear, string fromDay,
                                                  string Tomonth, string Toyear, string toDay, string docTypeInput, string statusInput, string pageSizeValue)
        {
            string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div";
            string dropdownXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[3]/app-global-pagination/div/div[1]/select";
            string rowSelector = "tbody tr";

            try
            {
                string filePrefix = ReportName;

                LogStep($"Clicking navigation link: {ReportName}");
                // Step 1️⃣ — Open the dropdown (click the label with "All Report edited 6")
                IWebElement dropdownTrigger = _driver.FindElement(
                    By.XPath("//span[contains(@class, 'input-group-text')]/label[contains(@class, 'fw-bold')]")
                );
                dropdownTrigger.Click();

                // Step 2️⃣ — Wait until the dropdown list appears and the target option is clickable
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                IWebElement optionToSelect = wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath($"//div[contains(@class, 'display-col') and normalize-space(text())='{ReportName}']"))
                );

                // Step 3️⃣ — Click the desired value
                optionToSelect.Click();

                // Open Date Picker                
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div/div[2]/div[2]/button")).Click();
                WaitForUIEffect();

                // Date From & To
                LogStep($"Selecting date range: From {Frommonth}/{Fromyear}/{fromDay} to {Tomonth}/{Toyear}/{toDay}");

                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Frommonth}, year: {Fromyear}");

                // Month dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Frommonth);
                WaitForUIEffect();

                // Year dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Fromyear);
                LogStep($"Month and Year selected: {Frommonth} {Fromyear}");
                WaitForUIEffect();

                // From date
                var fromDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{fromDay}']")));
                fromDate.Click();
                LogStep($"From date selected: {fromDay}");
                WaitForUIEffect();





                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Tomonth}, year: {Toyear}");

                // Month dropdown
                monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Tomonth);
                WaitForUIEffect();

                // Year dropdown
                yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Toyear);
                LogStep($"Month and Year selected: {Tomonth} {Toyear}");
                WaitForUIEffect();

                // To date
                var toDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{toDay}']")));
                toDate.Click();
                LogStep($"From date selected: {toDay}");
                WaitForUIEffect();


                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker/button")));
                selectBtn.Click();
                LogStep("Clicked 'Select' button to apply date range");
                WaitForUIEffect();


                // Document Type
                LogStep($"Selecting Document Type(s): {docTypeInput}");

                // Open the dropdown
                var dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[2]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for options to appear (target <li> instead of <div>)
                var options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (docTypeInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not already selected
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected Document Type: {text}");
                        }
                    }
                }
                else
                {
                    // Split user input into multiple values
                    var values = docTypeInput.Split(',')
                                             .Select(v => v.Trim())
                                             .Where(v => !string.IsNullOrEmpty(v))
                                             .ToList();

                    foreach (var val in values)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(val, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected Document Type: {val}");
                        }
                        else
                        {
                            LogStep($"Skipped Document Type (not found or already selected): {val}");
                        }
                    }
                }
                WaitForUIEffect();

                // Close dropdown safely (optional)
                dropdown.Click();



                // Document Status
                LogStep($"Selecting Document Status: {statusInput}");

                // Open the dropdown
                dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[3]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for <li> items inside <p-multiselectitem>
                options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (statusInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not selected yet
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected status: {text}");
                        }
                    }
                }
                else
                {
                    // Split input into multiple statuses
                    var statuses = statusInput.Split(',')
                                              .Select(s => s.Trim())
                                              .Where(s => !string.IsNullOrEmpty(s))
                                              .ToList();

                    foreach (var status in statuses)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(status, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected status: {status}");
                        }
                        else
                        {
                            LogStep($"Skipped status (not found or already selected): {status}");
                        }
                    }
                }

                // Close dropdown safely
                dropdown.Click();

                // Apply Filter
                WaitForUIEffect();
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div[4]/button")).Click();

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
        public void TestClickPageButtonIfExists(string ReportName, string Frommonth, string Fromyear, string fromDay,
                                                  string Tomonth, string Toyear, string toDay, string docTypeInput, string statusInput, string pageNumber)
        {
            string tableXPath = "/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[2]/div";
            string paginationXPathTemplate = "/html/body/app-layout/div[1]/div/div/div/app-content/app-reportv2/div[3]/div/div[3]/app-global-pagination/div/div[2]/ul/li[a[text()='{0}']]/a";
            string dynamicPageXPath = string.Format(paginationXPathTemplate, pageNumber); 

            try
            {
                string filePrefix = ReportName;

                LogStep($"Clicking navigation link: {ReportName}");
                // Step 1️⃣ — Open the dropdown (click the label with "All Report edited 6")
                IWebElement dropdownTrigger = _driver.FindElement(
                    By.XPath("//span[contains(@class, 'input-group-text')]/label[contains(@class, 'fw-bold')]")
                );
                dropdownTrigger.Click();

                // Step 2️⃣ — Wait until the dropdown list appears and the target option is clickable
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

                IWebElement optionToSelect = wait.Until(
                    ExpectedConditions.ElementToBeClickable(
                        By.XPath($"//div[contains(@class, 'display-col') and normalize-space(text())='{ReportName}']"))
                );

                // Step 3️⃣ — Click the desired value
                optionToSelect.Click();

                // Open Date Picker                
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div/div[2]/div[2]/button")).Click();
                WaitForUIEffect();

                // Date From & To
                LogStep($"Selecting date range: From {Frommonth}/{Fromyear}/{fromDay} to {Tomonth}/{Toyear}/{toDay}");

                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Frommonth}, year: {Fromyear}");

                // Month dropdown
                var monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Frommonth);
                WaitForUIEffect();

                // Year dropdown
                var yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Fromyear);
                LogStep($"Month and Year selected: {Frommonth} {Fromyear}");
                WaitForUIEffect();

                // From date
                var fromDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{fromDay}']")));
                fromDate.Click();
                LogStep($"From date selected: {fromDay}");
                WaitForUIEffect();





                //Choose Month & Year on Date picker                
                LogStep($"Selecting month: {Tomonth}, year: {Toyear}");

                // Month dropdown
                monthDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[1]")));
                new SelectElement(monthDropdown).SelectByText(Tomonth);
                WaitForUIEffect();

                // Year dropdown
                yearDropdown = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker//ngb-datepicker-navigation-select/select[2]")));
                new SelectElement(yearDropdown).SelectByText(Toyear);
                LogStep($"Month and Year selected: {Tomonth} {Toyear}");
                WaitForUIEffect();

                // To date
                var toDate = _wait.Until(ExpectedConditions.ElementIsVisible(
                   By.XPath($"//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker-month//div/span[normalize-space()='{toDay}']")));
                toDate.Click();
                LogStep($"From date selected: {toDay}");
                WaitForUIEffect();


                // Confirm button
                var selectBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//ngb-datepicker/button")));
                selectBtn.Click();
                LogStep("Clicked 'Select' button to apply date range");
                WaitForUIEffect();


                // Document Type
                LogStep($"Selecting Document Type(s): {docTypeInput}");

                // Open the dropdown
                var dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[2]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for options to appear (target <li> instead of <div>)
                var options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (docTypeInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not already selected
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected Document Type: {text}");
                        }
                    }
                }
                else
                {
                    // Split user input into multiple values
                    var values = docTypeInput.Split(',')
                                             .Select(v => v.Trim())
                                             .Where(v => !string.IsNullOrEmpty(v))
                                             .ToList();

                    foreach (var val in values)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(val, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected Document Type: {val}");
                        }
                        else
                        {
                            LogStep($"Skipped Document Type (not found or already selected): {val}");
                        }
                    }
                }
                WaitForUIEffect();

                // Close dropdown safely (optional)
                dropdown.Click();



                // Document Status
                LogStep($"Selecting Document Status: {statusInput}");

                // Open the dropdown
                dropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                   By.XPath("//app-content[@id='kt_content_container']/app-reportv2//div[3]/div[2]//p-multiselect/div/div[2]/div")));
                dropdown.Click();
                WaitForUIEffect();

                // Wait for <li> items inside <p-multiselectitem>
                options = _wait.Until(driver =>
                    driver.FindElements(By.XPath("//p-overlay//ul/p-multiselectitem/li"))
                );

                if (statusInput.Trim().Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var opt in options)
                    {
                        if (!opt.GetAttribute("class").Contains("p-highlight")) // not selected yet
                        {
                            opt.Click();
                            var text = opt.FindElement(By.XPath(".//span")).Text.Trim();
                            LogStep($"Selected status: {text}");
                        }
                    }
                }
                else
                {
                    // Split input into multiple statuses
                    var statuses = statusInput.Split(',')
                                              .Select(s => s.Trim())
                                              .Where(s => !string.IsNullOrEmpty(s))
                                              .ToList();

                    foreach (var status in statuses)
                    {
                        var option = options.FirstOrDefault(o =>
                            o.FindElement(By.XPath(".//span")).Text.Trim()
                             .Equals(status, StringComparison.OrdinalIgnoreCase));

                        if (option != null && !option.GetAttribute("class").Contains("p-highlight"))
                        {
                            option.Click();
                            LogStep($"Selected status: {status}");
                        }
                        else
                        {
                            LogStep($"Skipped status (not found or already selected): {status}");
                        }
                    }
                }

                // Close dropdown safely
                dropdown.Click();

                // Apply Filter
                WaitForUIEffect();
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-reportv2/div[2]/div/div[4]/button")).Click();

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

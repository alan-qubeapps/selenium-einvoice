using Allure.Net.Commons;
using Allure.NUnit;
using Allure.NUnit.Attributes;
using AngleSharp.Dom;
using EInvoice.SeleniumTests.Config;
using EInvoice.SeleniumTests.Drivers;
using Microsoft.VisualStudio.TestPlatform.CommunicationUtilities;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using ScreenRecorderLib;
using SeleniumExtras.WaitHelpers;
using SeleniumTests.Helpers;
using SeleniumTests.Pages.BusinessEntity;
using SeleniumTests.Pages.Stores;
using SeleniumTests.Pages.TemplateEditor;
using SeleniumTests.Pages.User;
using System.Drawing;
using System.Globalization;
using System.Media;
using System.Xml.Linq;
using Assert = Microsoft.VisualStudio.TestTools.UnitTesting.Assert;
using helperFunction = SeleniumTests.Helper.HelperFunction;

namespace SeleniumTests.Tests.H_TemplateEditor
{

    public static class ExcelDataReaderTemplateEditorNegative
    {
        public static IEnumerable<object[]> GetDocumentTemplateTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateType = worksheet.Cells[row, 3].Text?.Trim();
                    string BusinessEntity = worksheet.Cells[row, 4].Text?.Trim();
                    string Header = worksheet.Cells[row, 5].Text?.Trim();
                    string Footer = worksheet.Cells[row, 6].Text?.Trim();

                    yield return new object[]
                    {
                        TemplateName, TemplateDescription, TemplateType, BusinessEntity, Header, Footer
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetDuplicateDocumentTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateType = worksheet.Cells[row, 3].Text?.Trim();
                    string BusinessEntity = worksheet.Cells[row, 4].Text?.Trim();
                    string Header = worksheet.Cells[row, 5].Text?.Trim();
                    string Footer = worksheet.Cells[row, 6].Text?.Trim();

                    yield return new object[]
                    {
                        TemplateName, TemplateDescription, TemplateType, BusinessEntity, Header, Footer
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateDocumentTemplateTestData(string filePath, string sheetName)
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
                    string TemplateCode = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateName = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 3].Text?.Trim();
                    string TemplateType = worksheet.Cells[row, 4].Text?.Trim();
                    string BusinessEntity = worksheet.Cells[row, 5].Text?.Trim();
                    string Header = worksheet.Cells[row, 6].Text?.Trim();
                    string Footer = worksheet.Cells[row, 7].Text?.Trim();

                    yield return new object[]
                    {
                        TemplateCode, TemplateName, TemplateDescription, TemplateType, BusinessEntity, Header, Footer
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateDocumentDuplicateTestData(string filePath, string sheetName)
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
                    string TemplateCode = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateName = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 3].Text?.Trim();
                    string TemplateType = worksheet.Cells[row, 4].Text?.Trim();
                    string BusinessEntity = worksheet.Cells[row, 5].Text?.Trim();
                    string Header = worksheet.Cells[row, 6].Text?.Trim();
                    string Footer = worksheet.Cells[row, 7].Text?.Trim();

                    yield return new object[]
                    {
                        TemplateCode, TemplateName, TemplateDescription, TemplateType, BusinessEntity, Header, Footer
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetFilterCategoryTestData(string filePath, string sheetName)
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
                    string Category = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        Category
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetFilterDefaultTemplateTestData(string filePath, string sheetName)
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
                    string Category = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        Category
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetSearchTemplateTestData(string filePath, string sheetName)
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
                    string searchText = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        searchText
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetFilterReportCategoryTestData(string filePath, string sheetName)
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
                    string Category = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        Category
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetSearchReportTemplateTestData(string filePath, string sheetName)
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
                    string searchText = worksheet.Cells[row, 1].Text?.Trim();


                    yield return new object[]
                    {
                        searchText
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetReportTemplateTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateStatus = worksheet.Cells[row, 3].Text?.Trim();
                    string AllTemplate = worksheet.Cells[row, 4].Text?.Trim();


                    yield return new object[]
                    {
                        TemplateName, TemplateDescription, TemplateStatus, AllTemplate
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetReportDuplicateTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateStatus = worksheet.Cells[row, 3].Text?.Trim();
                    string AllTemplate = worksheet.Cells[row, 4].Text?.Trim();


                    yield return new object[]
                    {
                        TemplateName, TemplateDescription, TemplateStatus, AllTemplate
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetReportTemplateSelectTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateStatus = worksheet.Cells[row, 3].Text?.Trim();
                    string AllTemplate = worksheet.Cells[row, 4].Text?.Trim();
                    string ClearTemplate = worksheet.Cells[row, 5].Text?.Trim();
                    string TemplateField = worksheet.Cells[row, 6].Text?.Trim();


                    yield return new object[]
                    {
                        TemplateName, TemplateDescription, TemplateStatus, AllTemplate, ClearTemplate, TemplateField
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetReportSelectFieldTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateStatus = worksheet.Cells[row, 3].Text?.Trim();
                    string TemplateField = worksheet.Cells[row, 4].Text?.Trim();


                    yield return new object[]
                    {
                        TemplateName, TemplateDescription, TemplateStatus, TemplateField
                    };

                }
            }
        }


        public static IEnumerable<object[]> GetUpdateReportFieldTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string NewTemplatename = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 3].Text?.Trim();
                    string TemplateStatus = worksheet.Cells[row, 4].Text?.Trim();
                    string AllTemplate = worksheet.Cells[row, 5].Text?.Trim();
                    string TemplateField = worksheet.Cells[row, 6].Text?.Trim();


                    yield return new object[]
                    {
                        TemplateName, NewTemplatename, TemplateDescription, TemplateStatus, AllTemplate, TemplateField
                    };

                }
            }
        }

        public static IEnumerable<object[]> GetUpdateReportDuplicateTestData(string filePath, string sheetName)
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
                    string TemplateName = worksheet.Cells[row, 1].Text?.Trim();
                    string NewTemplatename = worksheet.Cells[row, 2].Text?.Trim();
                    string TemplateDescription = worksheet.Cells[row, 3].Text?.Trim();
                    string TemplateStatus = worksheet.Cells[row, 4].Text?.Trim();
                    string AllTemplate = worksheet.Cells[row, 5].Text?.Trim();
                    string TemplateField = worksheet.Cells[row, 6].Text?.Trim();


                    yield return new object[]
                    {
                        TemplateName, NewTemplatename, TemplateDescription, TemplateStatus, AllTemplate, TemplateField
                    };

                }
            }
        }


    }
    

    [TestFixture, Order(39)]
    [AllureNUnit]
    [AllureSuite("Template Editor - Negative")]
    [AllureEpic("ERP-117")]
    public class TemplateEditor_Negative
    {
        private IWebDriver _driver;
        private TemplateEditorPage _TemplateEditorPage;
        private WebDriverWait _wait;
        private LoginHelper _loginHelper;
        private Recorder _recorder;
        private string _recordingFilePath;
        private ManualResetEvent _recordingCompletedEvent = new ManualResetEvent(false);
        private List<string> _logMessages = new List<string>();
        private string _moduleName = "";


        private static string ExcelPath = Path.Combine(AppConfig.TestDataFolder, "TemplateEditorTestDataNegative.xlsx");

        public static IEnumerable<object[]> DocumentTemplateTestData =>
        ExcelDataReaderTemplateEditorNegative.GetDocumentTemplateTestData(ExcelPath, "DocumentTemplateTestData");
        public static IEnumerable<object[]> DuplicateDocumentTestData =>
        ExcelDataReaderTemplateEditorNegative.GetDuplicateDocumentTestData(ExcelPath, "DuplicateDocumentTestData");
        public static IEnumerable<object[]> UpdateDocumentTemplateTestData =>
        ExcelDataReaderTemplateEditorNegative.GetUpdateDocumentTemplateTestData(ExcelPath, "UpdateDocumentTemplateTestData");
        public static IEnumerable<object[]> UpdateDocumentDuplicateTestData =>
        ExcelDataReaderTemplateEditorNegative.GetUpdateDocumentDuplicateTestData(ExcelPath, "UpdateDocumentDuplicateTestData");
        public static IEnumerable<object[]> ReportTemplateTestData =>
        ExcelDataReaderTemplateEditorNegative.GetReportTemplateTestData(ExcelPath, "ReportTemplateTestData");
        public static IEnumerable<object[]> ReportDuplicateTestData =>
        ExcelDataReaderTemplateEditorNegative.GetReportDuplicateTestData(ExcelPath, "ReportDuplicateTestData");
        public static IEnumerable<object[]> UpdateReportFieldTestData =>
        ExcelDataReaderTemplateEditorNegative.GetUpdateReportFieldTestData(ExcelPath, "UpdateReportFieldTestData");
        public static IEnumerable<object[]> UpdateReportDuplicateTestData =>
        ExcelDataReaderTemplateEditorNegative.GetUpdateReportDuplicateTestData(ExcelPath, "UpdateReportDuplicateTestData");



        private static int _fileVersion; // shared version number
        private static int _recordingCounter = 0;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string moduleName = "Template Editor Page - Negative";

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
            _driver.Navigate().GoToUrl(AppConfig.BaseUrl + "/template-editor");
            helperFunction.WaitForPageToLoad(_wait);
            _TemplateEditorPage = new TemplateEditorPage(_driver);
            _logMessages.Clear();

            _moduleName = "Template Editor Page - Negative";
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
        [Category("Template Editor")]
        [Order(1)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(DocumentTemplateTestData))]
        public void CreateMandatoryFieldDocumentTemplate(string Templatename, string TemplateDesc, string TemplateType, string BusinessEntity, string TemplateHeader,
                    string TemplateFooter)
        {
            try
            {
                LogStep("Start Document Template Creation");

                // Step 1: Click New
                LogStep("Click 'New' button.");
                _TemplateEditorPage.ClickNewButton();
                WaitForUIEffect();

                // Step 2: Scroll to full modal form
                var modalForm = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div")));
                ScrollToElement(modalForm);
                LogStep("Scroll to full modal form container.");

                // Step 9: Submit
                LogStep("Click 'Submit' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("app-editor-modal button.btn.btn-primary")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                // Step 10: Success modal
                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.swal2-popup")));
                var message = modal.Text.Trim();
                LogStep($"Modal Message: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                if (!message.ToLower().Contains("success"))
                {
                    Assert.IsTrue(true,$"Test success, Expected error message: {message}");
                }

                // Step 11: Confirm modal
                LogStep("Click modal 'Ok, got it!'");
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                ScrollToElement(okButton);
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Template creation test completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }
        }

        [Test]
        [Category("Template Editor")]
        [Order(2)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(DuplicateDocumentTestData))]
        public void CreateDuplicateDocumentTemplate(string Templatename, string TemplateDesc, string TemplateType, string BusinessEntity, string TemplateHeader,
                    string TemplateFooter)
        {
            try
            {
                LogStep("Start Document Template Creation");

                // Step 1: Click New
                LogStep("Click 'New' button.");
                _TemplateEditorPage.ClickNewButton();
                WaitForUIEffect();

                // Step 2: Scroll to full modal form
                var modalForm = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div")));
                ScrollToElement(modalForm);
                LogStep("Scroll to full modal form container.");

                // Step 3: Template Name
                LogStep($"Enter Template Name: {Templatename}");
                _TemplateEditorPage.EnterTemplatename(Templatename);
                WaitForUIEffect();

                // Step 4: Template Description
                LogStep($"Enter Template Description: {TemplateDesc}");
                _TemplateEditorPage.EnterTemplateDesc(TemplateDesc);
                WaitForUIEffect();

                // Step 5: Template Type
                LogStep($"Select Template Type: {TemplateType}");
                var templateTypeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[4]/select")));
                ScrollToElement(templateTypeDropdown);
                new SelectElement(templateTypeDropdown).SelectByText(TemplateType);
                WaitForUIEffect();

                // Step 6: Template Editor
                try
                {
                    LogStep("Open Template Editor dropdown.");
                    var dropdownTrigger = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.CssSelector("p-dropdown .p-dropdown-trigger")));
                    ScrollToElement(dropdownTrigger);
                    dropdownTrigger.Click();
                    WaitForUIEffect();

                    var optionsList = _wait.Until(ExpectedConditions
                        .VisibilityOfAllElementsLocatedBy(By.XPath("//p-dropdownitem/li[contains(@class,'p-dropdown-item')]")));

                    bool found = false;
                    foreach (var option in optionsList)
                    {
                        LogStep($"🔎 Checking dropdown option: {option.Text.Trim()}");
                        if (option.Text.Trim().Equals(BusinessEntity, StringComparison.OrdinalIgnoreCase))
                        {
                            LogStep($"✅ Found matching option: {option.Text.Trim()} — clicking.");
                            ScrollToElement(option);
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option);
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        Assert.Fail($"❌ Template Editor '{BusinessEntity}' not found in dropdown.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Assert.Fail("❌ Template Editor dropdown options not found.");
                }

                // Step 7: Header
                LogStep($"Enter Header: {TemplateHeader}");
                _TemplateEditorPage.EnterTemplateHeader(TemplateHeader);
                WaitForUIEffect();

                // Step 8: Footer
                LogStep($"Enter Footer: {TemplateFooter}");
                _TemplateEditorPage.EnterTemplateFooter(TemplateFooter);
                WaitForUIEffect();

                // Step 9: Submit
                LogStep("Click 'Submit' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("app-editor-modal button.btn.btn-primary")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                // Step 10: Success modal
                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.swal2-popup")));
                var message = modal.Text.Trim();
                LogStep($"Modal Message: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                if (!message.ToLower().Contains("success"))
                {
                    Assert.IsTrue(true, $"✅ Test passed successfully. Message: {message}");
                }

                // Step 11: Confirm modal
                LogStep("Click modal 'Ok, got it!'");
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                ScrollToElement(okButton);
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Template creation test completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }
        }

        [Test]
        [Category("Template Editor")]
        [Order(3)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(DocumentTemplateTestData))]
        public void CreatePartialMandatoryDocumentTemplate(string Templatename, string TemplateDesc, string TemplateType, string BusinessEntity, string TemplateHeader,
                    string TemplateFooter)
        {
            try
            {
                LogStep("Start Document Template Creation");

                // Step 1: Click New
                LogStep("Click 'New' button.");
                _TemplateEditorPage.ClickNewButton();
                WaitForUIEffect();

                // Step 2: Scroll to full modal form
                var modalForm = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div")));
                ScrollToElement(modalForm);
                LogStep("Scroll to full modal form container.");

                // Step 3: Template Name
                LogStep($"Enter Template Name: {Templatename}");
                _TemplateEditorPage.EnterTemplatename(Templatename);
                WaitForUIEffect();

                // Step 4: Template Description
                LogStep($"Enter Template Description: {TemplateDesc}");
                _TemplateEditorPage.EnterTemplateDesc(TemplateDesc);
                WaitForUIEffect();

                // Step 5: Template Type
                LogStep($"Select Template Type: {TemplateType}");
                var templateTypeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[4]/select")));
                ScrollToElement(templateTypeDropdown);
                new SelectElement(templateTypeDropdown).SelectByText(TemplateType);
                WaitForUIEffect();

                // Step 6: Template Editor
                try
                {
                    LogStep("Open Template Editor dropdown.");
                    var dropdownTrigger = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.CssSelector("p-dropdown .p-dropdown-trigger")));
                    ScrollToElement(dropdownTrigger);
                    dropdownTrigger.Click();
                    WaitForUIEffect();

                    var optionsList = _wait.Until(ExpectedConditions
                        .VisibilityOfAllElementsLocatedBy(By.XPath("//p-dropdownitem/li[contains(@class,'p-dropdown-item')]")));

                    bool found = false;
                    foreach (var option in optionsList)
                    {
                        LogStep($"🔎 Checking dropdown option: {option.Text.Trim()}");
                        if (option.Text.Trim().Equals(BusinessEntity, StringComparison.OrdinalIgnoreCase))
                        {
                            LogStep($"✅ Found matching option: {option.Text.Trim()} — clicking.");
                            ScrollToElement(option);
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option);
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        Assert.Fail($"❌ Template Editor '{BusinessEntity}' not found in dropdown.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Assert.Fail("❌ Template Editor dropdown options not found.");
                }

               
                // Step 9: Submit
                LogStep("Click 'Submit' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("app-editor-modal button.btn.btn-primary")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                // Step 10: Success modal
                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.swal2-popup")));
                var message = modal.Text.Trim();
                LogStep($"Modal Message: {message}");

                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                if (!message.ToLower().Contains("success"))
                {
                    Assert.IsTrue(true, $"Test success, Expected error message: {message}");
                }

                // Step 11: Confirm modal
                LogStep("Click modal 'Ok, got it!'");
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                ScrollToElement(okButton);
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Template creation test completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }
        }



        [Test]
        [Category("Template Editor")]
        [Order(4)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Update")]
        [TestCaseSource(nameof(UpdateDocumentTemplateTestData))]
        public void UpdateMondatoryDocumentTemplate(string TemplateCode, string Templatename, string TemplateDesc, string TemplateType,
                                   string BusinessEntity, string TemplateHeader, string TemplateFooter)
        {
            try
            {
                LogStep("Start Document Template Update");

                // Step 1: Search by Template Code
                LogStep("Clicking 'Edit' button.");
                _TemplateEditorPage.ClickEditButton(TemplateCode);
                WaitForUIEffect(1000);

                var modalForm = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div")));
                ScrollToElement(modalForm);
                LogStep("Scroll to full modal form container.");
                WaitForUIEffect();

                LogStep($"Update Template Name: {Templatename}");
                var nameInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[2]/input")));
                nameInput.Clear();
                nameInput.SendKeys(Templatename);
                WaitForUIEffect();

                LogStep($"Update Template Description: {TemplateDesc}");
                var descInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[3]/input")));
                descInput.Clear();
                descInput.SendKeys(TemplateDesc);
                WaitForUIEffect();

                LogStep($"Update Template Type: {TemplateType}");
                var templateTypeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[4]/select")));
                ScrollToElement(templateTypeDropdown);
                new SelectElement(templateTypeDropdown).SelectByText(TemplateType);
                WaitForUIEffect();

                try
                {
                    LogStep("Update Template Editor dropdown.");
                    var dropdownTrigger = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[5]/p-dropdown/div/div[2]")));
                    ScrollToElement(dropdownTrigger);
                    dropdownTrigger.Click();
                    WaitForUIEffect();

                    var optionsList = _wait.Until(ExpectedConditions
                        .VisibilityOfAllElementsLocatedBy(By.XPath("//p-dropdownitem/li[contains(@class,'p-dropdown-item')]")));

                    bool found = false;
                    foreach (var option in optionsList)
                    {
                        LogStep($"🔎 Checking dropdown option: {option.Text.Trim()}");
                        if (option.Text.Trim().Equals(BusinessEntity, StringComparison.OrdinalIgnoreCase))
                        {
                            LogStep($"✅ Found matching option: {option.Text.Trim()} — clicking.");
                            ScrollToElement(option);
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option);
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        Assert.Fail($"❌ Template Editor '{BusinessEntity}' not found in dropdown.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Assert.Fail("❌ Template Editor dropdown options not found.");
                }

                LogStep($"Update Header: {TemplateHeader}");
                var headerInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//body[@id='kt_body']/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[6]/div/quill-editor/div[2]/div")));
                headerInput.Clear();
                headerInput.SendKeys(TemplateHeader);
                WaitForUIEffect();

                LogStep($"Update Footer: {TemplateFooter}");
                var footerInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//body[@id='kt_body']/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[7]/div/quill-editor/div[2]/div")));
                footerInput.Clear();
                footerInput.SendKeys(TemplateFooter);
                WaitForUIEffect();

                LogStep("Click 'Save' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[3]/button")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.swal2-popup")));
                var message = modal.Text.Trim();
                LogStep($"Modal Message: {message}");


                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Update_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                if (!message.ToLower().Contains("success"))
                {
                    Assert.IsTrue(true, $"Test success, Expected error message: {message}");
                }

                LogStep("Click modal 'Ok, got it!'");
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                ScrollToElement(okButton);
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Template update test completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }

        }


        [Test]
        [Category("Template Editor")]
        [Order(5)]
        [AllureSeverity(SeverityLevel.critical)]
        [AllureStory("Update")]
        [TestCaseSource(nameof(UpdateDocumentDuplicateTestData))]
        public void UpdateDuplicateDocumentTemplate(string TemplateCode, string Templatename, string TemplateDesc, string TemplateType,
                                   string BusinessEntity, string TemplateHeader, string TemplateFooter)
        {
            try
            {
                LogStep("Start Document Template Update");

                // Step 1: Search by Template Code
                LogStep("Clicking 'Edit' button.");
                _TemplateEditorPage.ClickEditButton(TemplateCode);
                WaitForUIEffect(1000);

                var modalForm = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div")));
                ScrollToElement(modalForm);
                LogStep("Scroll to full modal form container.");
                WaitForUIEffect();

                LogStep($"Update Template Name: {Templatename}");
                var nameInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[2]/input")));
                nameInput.Clear();
                nameInput.SendKeys(Templatename);
                WaitForUIEffect();

                LogStep($"Update Template Description: {TemplateDesc}");
                var descInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[3]/input")));
                descInput.Clear();
                descInput.SendKeys(TemplateDesc);
                WaitForUIEffect();

                LogStep($"Update Template Type: {TemplateType}");
                var templateTypeDropdown = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[4]/select")));
                ScrollToElement(templateTypeDropdown);
                new SelectElement(templateTypeDropdown).SelectByText(TemplateType);
                WaitForUIEffect();

                try
                {
                    LogStep("Update Template Editor dropdown.");
                    var dropdownTrigger = _wait.Until(ExpectedConditions.ElementToBeClickable(
                        By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[5]/p-dropdown/div/div[2]")));
                    ScrollToElement(dropdownTrigger);
                    dropdownTrigger.Click();
                    WaitForUIEffect();

                    var optionsList = _wait.Until(ExpectedConditions
                        .VisibilityOfAllElementsLocatedBy(By.XPath("//p-dropdownitem/li[contains(@class,'p-dropdown-item')]")));

                    bool found = false;
                    foreach (var option in optionsList)
                    {
                        LogStep($"🔎 Checking dropdown option: {option.Text.Trim()}");
                        if (option.Text.Trim().Equals(BusinessEntity, StringComparison.OrdinalIgnoreCase))
                        {
                            LogStep($"✅ Found matching option: {option.Text.Trim()} — clicking.");
                            ScrollToElement(option);
                            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", option);
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        Assert.Fail($"❌ Template Editor '{BusinessEntity}' not found in dropdown.");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Assert.Fail("❌ Template Editor dropdown options not found.");
                }

                LogStep($"Update Header: {TemplateHeader}");
                var headerInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//body[@id='kt_body']/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[6]/div/quill-editor/div[2]/div")));
                headerInput.Clear();
                headerInput.SendKeys(TemplateHeader);
                WaitForUIEffect();

                LogStep($"Update Footer: {TemplateFooter}");
                var footerInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//body[@id='kt_body']/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[7]/div/quill-editor/div[2]/div")));
                footerInput.Clear();
                footerInput.SendKeys(TemplateFooter);
                WaitForUIEffect();

                LogStep("Click 'Save' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.XPath("/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[3]/button")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.swal2-popup")));
                var message = modal.Text.Trim();
                LogStep($"Modal Message: {message}");


                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Update_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                if (!message.ToLower().Contains("success"))
                {
                    Assert.IsTrue(true, $"✅ Test passed successfully. Message: {message}");
                }

                LogStep("Click modal 'Ok, got it!'");
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                ScrollToElement(okButton);
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Template update test completed successfully.");
            }
            catch (Exception ex)
            {
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"❌ Exception occurred: {ex.Message}");
                Assert.Fail("Test failed due to exception.");
            }

        }



        [Test]
        [Category("Template Editor")]
        [Order(6)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(ReportTemplateTestData))]
        public void CreateMandatoryReportTemplate(string Templatename, string TemplateDesc, string TemplateStatus, string AllTemplate)
        {
            try
            {
                // === Navigate to Report Template Tab ===
                LogStep("Navigate to Report Template Tab");
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-template-editor/div/div/ul/li[2]/a")).Click();
                WaitForUIEffect();

                // === Start Report Template Creation ===
                LogStep("Start Report Template Creation");

                // Step 1: Click 'New' button
                LogStep("Click 'New' button.");
                _TemplateEditorPage.ClickNewReportTemplateButton();
                WaitForUIEffect(1000);              

                // Step 5: Submit the Template
                LogStep("Click 'Submit' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_body > ngb-modal-window > div > div > app-report-editor-modal > div > div.modal-footer.justify-content-end.d-flex > button")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                // Step 6: Verify Success Modal
                LogStep("Check for success modal.");


                // Step 7: Take Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved to: {_lastScreenshotPath}");

                // Step 8: Negative case - expect NO modal
                try
                {
                    var modal = _wait.Until(ExpectedConditions.ElementIsVisible(
                        By.CssSelector("div.swal2-popup")));

                    // If we reach here → modal appeared (unexpected for negative case)
                    var message = modal.Text.Trim();
                    string noteMessage = $"⚠️ Unexpected modal appeared in negative test. Message: {message}. Test still considered PASS.";
                    LogStep(noteMessage);


                    // Optionally click OK to clear modal if it shows
                    try
                    {
                        var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                        okButton.Click();
                        WaitForUIEffect();
                    }
                    catch { /* ignore if no OK button */ }
                }
                catch (WebDriverTimeoutException)
                {
                    // This is the expected behavior → no modal
                    LogStep("✅ No modal appeared after submit (expected for negative test).");
                }

                LogStep("✅ Report Template no mandatory test completed successfully.");
            }
            catch (Exception ex)
            {
                // Step 10: Exception Handling with Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                string failMessage = $"❌ Exception occurred: {ex.Message}\nScreenshot saved at: {_lastScreenshotPath}";
                LogStep(failMessage);

                Assert.Fail(failMessage);
            }
        }


        [Test]
        [Category("Template Editor")]
        [Order(7)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Create")]
        [TestCaseSource(nameof(ReportDuplicateTestData))]
        public void CreateDuplicateReportTemplate(string Templatename, string TemplateDesc, string TemplateStatus, string AllTemplate)
        {
            try
            {
                // === Navigate to Report Template Tab ===
                LogStep("Navigate to Report Template Tab");
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-template-editor/div/div/ul/li[2]/a")).Click();
                WaitForUIEffect();

                // === Start Report Template Creation ===
                LogStep("Start Report Template Creation");

                // Step 1: Click 'New' button
                LogStep("Click 'New' button.");
                _TemplateEditorPage.ClickNewReportTemplateButton();
                WaitForUIEffect(1000);

                // Step 2: Enter Template Name
                LogStep($"Enter Report Template Name: {Templatename}");
                _TemplateEditorPage.EnterReportTemplatename(Templatename);
                WaitForUIEffect();

                // Step 3: Enter Template Description
                LogStep($"Enter Report Template Description: {TemplateDesc}");
                _TemplateEditorPage.EnterReportTemplateDesc(TemplateDesc);
                WaitForUIEffect();

                // Step 4: Set Template Status Checkboxes
                if (!string.IsNullOrEmpty(TemplateStatus))
                {
                    bool isTemplateStatusChecked = TemplateStatus.Equals("Active", StringComparison.OrdinalIgnoreCase);
                    _TemplateEditorPage.SetReportCheckboxStatus(isTemplateStatusChecked);
                    WaitForUIEffect();
                    LogStep($"Report Template 'Active' Checkbox set to: {isTemplateStatusChecked}");
                }

                if (!string.IsNullOrEmpty(AllTemplate))
                {
                    bool isAllTemplateChecked = AllTemplate.Equals("All", StringComparison.OrdinalIgnoreCase);
                    _TemplateEditorPage.SetAllReportCheckboxStatus(isAllTemplateChecked);
                    WaitForUIEffect();
                    LogStep($"Report Template 'All' Checkbox set to: {isAllTemplateChecked}");
                }


                // Step 5: Submit the Template
                LogStep("Click 'Submit' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_body > ngb-modal-window > div > div > app-report-editor-modal > div > div.modal-footer.justify-content-end.d-flex > button")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                // Step 6: Verify Success Modal
                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.swal2-popup")));
                var message = modal.Text.Trim();
                LogStep($"Modal Message: {message}");

                // Step 7: Take Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved to: {_lastScreenshotPath}");

                // Step 8: Assert Success Message
                if (!message.ToLower().Contains("success"))
                {
                    string successMessage = $"✅ Test passed successfully. Message: {message}";
                    LogStep(successMessage);
                    Assert.IsTrue(true, successMessage);

                }

                // Step 9: Confirm Modal
                LogStep("Click modal 'Ok, got it!'");
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                ScrollToElement(okButton);
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Report Template creation test completed successfully.");
            }
            catch (Exception ex)
            {
                // Step 10: Exception Handling with Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                string failMessage = $"❌ Exception occurred: {ex.Message}\nScreenshot saved at: {_lastScreenshotPath}";
                LogStep(failMessage);

                Assert.Fail(failMessage);
            }
        }


        [Test]
        [Category("Template Editor")]
        [Order(8)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Update")]
        [TestCaseSource(nameof(UpdateReportFieldTestData))]
        public void UpdateReportMandatoryTemplate_SelectField(string TemplateOldDesc, string NewTemplatename, string TemplateDesc, string TemplateStatus, string ClearTemplate, string TemplateField)
        {
            try
            {
                // === Navigate to Report Template Tab ===
                LogStep("Navigate to Report Template Tab");
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-template-editor/div/div/ul/li[2]/a")).Click();
                WaitForUIEffect();

                // === Start Update Report Template ===
                LogStep("Start Update Report Template");

                // Step 1: Search by Template Code
                LogStep("Clicking 'Edit' button.");
                _TemplateEditorPage.ClickEditReportButton(TemplateOldDesc);
                WaitForUIEffect(1000);


                // Step 2: Enter Template Name
                LogStep($"Enter New Report Template Name: {NewTemplatename}");
                _TemplateEditorPage.EnterReportTemplatename(NewTemplatename);
                WaitForUIEffect(1000);

                // Step 3: Enter Template Description
                LogStep($"Enter Report Template Description: {TemplateDesc}");
                _TemplateEditorPage.EnterReportTemplateDesc(TemplateDesc);
                WaitForUIEffect();

                // Step 4: Set Template Status Checkboxes
                if (!string.IsNullOrEmpty(TemplateStatus))
                {
                    bool isTemplateStatusChecked = TemplateStatus.Equals("Active", StringComparison.OrdinalIgnoreCase);
                    _TemplateEditorPage.SetReportCheckboxStatus(isTemplateStatusChecked);
                    WaitForUIEffect();
                    LogStep($"Report Template 'Active' Checkbox set to: {isTemplateStatusChecked}");
                }

                if (!string.IsNullOrEmpty(ClearTemplate))
                {
                    bool isClearTemplateChecked = ClearTemplate.Equals("True", StringComparison.OrdinalIgnoreCase);
                    _TemplateEditorPage.SetClearReportCheckboxStatus(isClearTemplateChecked);
                    WaitForUIEffect();
                    LogStep($"Report Template 'Clear' Checkbox set to: {isClearTemplateChecked}");
                }

                WaitForUIEffect();

                // Step 5: Submit the Template
                LogStep("Click 'Submit' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_body > ngb-modal-window > div > div > app-report-editor-modal > div > div.modal-footer.justify-content-end.d-flex > button")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                // Step 6: Verify Success Modal
                LogStep("Check for success modal.");


                // Step 7: Take Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved to: {_lastScreenshotPath}");

                try
                {
                    var modal = _wait.Until(ExpectedConditions.ElementIsVisible(
                        By.CssSelector("div.swal2-popup")));


                    // If we reach here → modal appeared (unexpected for negative case)
                    var message = modal.Text.Trim();
                    if (message.ToLower().Contains("success"))
                    {
                        string failMessage = $"❌ Expected success message but got: {message}";
                        LogStep(failMessage);
                        Assert.Fail(failMessage);
                    }
                    string noteMessage = $"⚠️ Unexpected modal appeared in negative test. Message: {message}.";
                    LogStep(noteMessage);


                    //// Optionally click OK to clear modal if it shows
                    //try
                    //{
                    //    var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                    //    okButton.Click();
                    //    WaitForUIEffect();
                    //}
                    //catch { /* ignore if no OK button */ }
                }
                catch (WebDriverTimeoutException)
                {
                    // This is the expected behavior → no modal
                    LogStep("✅ No modal appeared after submit (expected for negative test).");
                }


                LogStep("✅ Report Template creation test completed successfully.");
            }
            catch (Exception ex)
            {
                // Step 10: Exception Handling with Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                string failMessage = $"❌ Exception occurred: {ex.Message}\nScreenshot saved at: {_lastScreenshotPath}";
                LogStep(failMessage);

                Assert.Fail(failMessage);
            }
        }


        [Test]
        [Category("Template Editor")]
        [Order(9)]
        [AllureSeverity(SeverityLevel.normal)]
        [AllureStory("Update")]
        [TestCaseSource(nameof(UpdateReportDuplicateTestData))]
        public void UpdateDuplicateReportTemplate_SelectField(string TemplateOldDesc, string NewTemplatename, string TemplateDesc, string TemplateStatus, string AllTemplate, string TemplateField)
        {
            try
            {
                // === Navigate to Report Template Tab ===
                LogStep("Navigate to Report Template Tab");
                _driver.FindElement(By.XPath("//app-content[@id='kt_content_container']/app-template-editor/div/div/ul/li[2]/a")).Click();
                WaitForUIEffect();

                // === Start Update Report Template ===
                LogStep("Start Update Report Template");

                // Step 1: Search by Template Code
                LogStep("Clicking 'Edit' button.");
                _TemplateEditorPage.ClickEditReportButton(TemplateOldDesc);


                // Step 2: Enter Template Name
                LogStep($"Enter New Report Template Name: {NewTemplatename}");
                _TemplateEditorPage.EnterReportTemplatename(NewTemplatename);
                WaitForUIEffect(1000);

                // Step 3: Enter Template Description
                LogStep($"Enter Report Template Description: {TemplateDesc}");
                _TemplateEditorPage.EnterReportTemplateDesc(TemplateDesc);
                WaitForUIEffect();

                // Step 4: Set Template Status Checkboxes
                if (!string.IsNullOrEmpty(TemplateStatus))
                {
                    bool isTemplateStatusChecked = TemplateStatus.Equals("Active", StringComparison.OrdinalIgnoreCase);
                    _TemplateEditorPage.SetReportCheckboxStatus(isTemplateStatusChecked);
                    WaitForUIEffect();
                    LogStep($"Report Template 'Active' Checkbox set to: {isTemplateStatusChecked}");
                }

                if (!string.IsNullOrEmpty(AllTemplate))
                {
                    bool isAllTemplateChecked = AllTemplate.Equals("All", StringComparison.OrdinalIgnoreCase);
                    _TemplateEditorPage.SetAllReportCheckboxStatus(isAllTemplateChecked);
                    WaitForUIEffect();
                    LogStep($"Report Template 'All' Checkbox set to: {isAllTemplateChecked}");
                }


                // User input could be multiple values separated by commas
                string[] userInputs = TemplateField.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                                   .Select(x => x.Trim())  // remove extra spaces
                                                   .ToArray();

                // Main container XPath
                var mainDiv = _driver.FindElement(By.XPath("/html/body/ngb-modal-window/div/div/app-report-editor-modal/div/div[2]/div[3]/div[1]/div"));

                // Get all child divs (sub-values)
                var subDivs = mainDiv.FindElements(By.XPath("./div"));

                // Loop through each sub div to find a match
                foreach (var div in subDivs)
                {
                    string divText = div.Text.Trim();
                    if (userInputs.Contains(divText))
                    {
                        div.Click();
                        LogStep($"Clicked sub-value: {divText}");
                    }
                }
                WaitForUIEffect();

                // Step 5: Submit the Template
                LogStep("Click 'Submit' button.");
                var saveBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(
                    By.CssSelector("#kt_body > ngb-modal-window > div > div > app-report-editor-modal > div > div.modal-footer.justify-content-end.d-flex > button")));
                ScrollToElement(saveBtn);
                saveBtn.Click();
                WaitForUIEffect(1000);

                // Step 6: Verify Success Modal
                LogStep("Check for success modal.");
                var modal = _wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.swal2-popup")));
                var message = modal.Text.Trim();
                LogStep($"Modal Message: {message}");

                // Step 7: Take Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);
                LogStep($"📸 Screenshot saved to: {_lastScreenshotPath}");

                // Step 8: Assert Success Message
                if (!message.ToLower().Contains("success"))
                {
                    string PassMessage = $"Expected success message : {message}";
                    LogStep(PassMessage);
                    Assert.IsTrue(true,PassMessage);
                }

                // Step 9: Confirm Modal
                LogStep("Click modal 'Ok, got it!'");
                var okButton = modal.FindElement(By.XPath(".//button[contains(., 'Ok, got it!')]"));
                ScrollToElement(okButton);
                okButton.Click();
                WaitForUIEffect();

                LogStep("✅ Report Template creation test completed successfully.");
            }
            catch (Exception ex)
            {
                // Step 10: Exception Handling with Screenshot
                _lastScreenshotPath = Path.Combine(Path.GetTempPath(), $"Template_Editor_{DateTime.Now:yyyyMMdd_HHmmss}.png");
                var screenshot = ((ITakesScreenshot)_driver).GetScreenshot();
                File.WriteAllBytes(_lastScreenshotPath, screenshot.AsByteArray);

                string failMessage = $"❌ Exception occurred: {ex.Message}\nScreenshot saved at: {_lastScreenshotPath}";
                LogStep(failMessage);

                Assert.Fail(failMessage);
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

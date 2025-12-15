using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace SeleniumTests.Pages.BatchProcessPage


{
    public class BatchProcessPage
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        // Constructor
        public BatchProcessPage(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            PageFactory.InitElements(_driver, this);
        }

        
        public void ClickExportButton(int rowNumber)
        {
            string cssSelector = $"#kt_content_container app-joblist table tbody tr:nth-child({rowNumber}) td.text-end div a";
            var button = _driver.FindElement(By.CssSelector(cssSelector));
            button.Click();
        }

        public string GetFileNameFromRow(int rowNumber)
        {
            string xpath = $"//table/tbody/tr[{rowNumber}]/td[2]/span";
            return _driver.FindElement(By.XPath(xpath)).Text.Trim();
        }


        public bool WaitForFileDownload(string folderPath, string filePrefix, TimeSpan timeout)
        {
            string todayDate = DateTime.Now.ToString("yyyyMMdd");
            string expectedBaseName = $"{filePrefix}_{todayDate}";

            // Capture existing files before export
            var existingFiles = new HashSet<string>(Directory.GetFiles(folderPath));

            var endTime = DateTime.Now + timeout;
            while (DateTime.Now < endTime)
            {
                var currentFiles = Directory.GetFiles(folderPath);
                foreach (var file in currentFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);

                    // File matches naming pattern and is NEW
                    if (fileName.StartsWith(expectedBaseName) && !existingFiles.Contains(file))
                    {
                        var fileInfo = new FileInfo(file);
                        if (fileInfo.Length > 0)
                            return true;
                    }
                }

                Thread.Sleep(500);
            }

            return false; // No new matching file found in time
        }


    }
}

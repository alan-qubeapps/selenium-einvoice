using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace SeleniumTests.Pages.Report
{
    public class ReportPage
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        // Constructor
        public ReportPage(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            PageFactory.InitElements(_driver, this);
        }

        // Elements
        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-business-entity > div > div.d-flex.gap-2.justify-content-start > div > ul > li.nav-item.ps-9.pt-3 > a")]
        private IWebElement FilterALLCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-business-entity > div > div.d-flex.gap-2.justify-content-start > div > ul > li:nth-child(2) > a")]
        private IWebElement FilterActiveCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-business-entity > div > div.d-flex.gap-2.justify-content-start > div > ul > li:nth-child(3) > a")]
        private IWebElement FilterInactiveCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-business-entity > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(2) > span")]
        private IWebElement FilterPendingCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-business-entity > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(3) > span")]
        private IWebElement FilterSuccessCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-business-entity > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(4) > span")]
        private IWebElement FilterFailedCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-business-entity > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a")]
        private IWebElement ImportButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2")]
        private IWebElement UploadButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-body.px-20 > div > div > div.d-flex.align-items-center > button")]
        private IWebElement DownloadButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-reportv2 > div.card.py-8.ps-8.pe-6.mt-5.ng-star-inserted > div > div.d-flex.justify-content-end.ng-star-inserted > div > a")]
        private IWebElement ExportButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-summary-sales > div.card.py-8.ps-8.pe-6.mt-5 > div > div.d-flex.justify-content-end.ng-star-inserted > div > a")]
        private IWebElement ExportSummaryButton { get; set; }


        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_content_container\"]/app-business-entity/div/div[3]/div/div[3]/a")]
        private IWebElement NewButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-business-entity-modal/div/div[3]/div/div[2]/button")]
        private IWebElement ContinueButton { get; set; }

        

        // New Elements
        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_create_account_form\"]/div/app-step1s/div/form/div/div/div[1]/input")]
        public IWebElement BEnameInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step1s/div/form/div/div/div[2]/input")]
        public IWebElement BETinNumberInput { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_create_account_form\"]/div/app-step1s/div/form/div/div/div[3]/div[2]/input")]
        private IWebElement BERegisterIDInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step1s/div/form/div/div/div[4]/input")]
        private IWebElement BEsstInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step1s/div/form/div/div/div[5]/input")]
        private IWebElement BETTRegisterNumberInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[2]/input")]
        private IWebElement BEContactNumberInput { get; set; }
        

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[3]/input")]
        private IWebElement BEemailInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[4]/div[1]/input")]
        private IWebElement BECityInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[4]/div[3]/input")]
        private IWebElement BEPosCodeInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[6]/input")]
        private IWebElement BEAddress1Input { get; set; }
        
        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[7]/input")]
        private IWebElement BEAddress2Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[8]/input")]
        private IWebElement BEAddress3Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-business-entity-modal/div/div[3]/div/div[2]/button")]
        private IWebElement SaveButton { get; set; }

        
        public void ClickExportButton()
        {
            ExportButton.Click();
        }

        public void ClickExportSummaryButton()
        {
            ExportSummaryButton.Click();
        }     

        public bool WaitForFileDownload(string downloadPath, string filePrefix, TimeSpan timeout)
        {
            var endTime = DateTime.Now + timeout;

            while (DateTime.Now < endTime)
            {
                var files = Directory.GetFiles(downloadPath, $"{filePrefix}*")
                    .OrderByDescending(File.GetCreationTime)
                    .ToList();

                if (files.Count > 0)
                {
                    var newest = files.First();
                    var creationTime = File.GetCreationTime(newest);

                    // Check if file is recent (within 1 min)
                    if ((DateTime.Now - creationTime).TotalSeconds < 60)
                        return true;
                }

                Thread.Sleep(1000); // Wait 1s before next check
            }

            return false;
        }



    }
}

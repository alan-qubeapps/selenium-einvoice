using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace SeleniumTests.Pages.TemplateEditor


{
    public class TemplateEditorPage
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        // Constructor
        public TemplateEditorPage(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            PageFactory.InitElements(_driver, this);
        }

        // Elements
        [FindsBy(How = How.LinkText, Using = "All")]
        private IWebElement FilterALLCategoryButton { get; set; }

        [FindsBy(How = How.LinkText, Using = "Active")]
        private IWebElement FilterActiveCategoryButton { get; set; }

        [FindsBy(How = How.LinkText, Using = "Inactive")]
        private IWebElement FilterInactiveCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-store > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(2) > span")]
        private IWebElement FilterPendingCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-store > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(3) > span")]
        private IWebElement FilterSuccessCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-store > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(4) > span")]
        private IWebElement FilterFailedCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-store > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a")]
        private IWebElement ImportButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2")]
        private IWebElement UploadButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-body.px-20 > div > div > div.d-flex.align-items-center > button")]
        private IWebElement DownloadButton { get; set; }


        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-store > div > div.card-header.border-0.pt-5 > div > div:nth-child(2) > a")]
        private IWebElement ExportButton { get; set; }


        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-template-editor > app-pdf-editor > div > div.card-header.border-0.pt-5 > div > div > a")]
        private IWebElement NewButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-template-editor > app-report-template > div > div.card-header.border-0.pt-5 > div > div > a")]
        private IWebElement NewReportButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-store-modal/div/div[3]/div/div[2]/button")]
        private IWebElement ContinueButton { get; set; }

        // New Elements
        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[2]/input")]
        public IWebElement TemplatenameInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-report-editor-modal/div/div[2]/div[1]/div[1]/div[2]/input")]
        public IWebElement ReportTemplatenameInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[3]/input")]
        public IWebElement TemplateDescInput { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-report-editor-modal/div/div[2]/div[1]/div[2]/div[2]/input")]
        public IWebElement ReportTemplateDescInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[6]/div/quill-editor/div[2]/div[1]")]
        private IWebElement TemplateHeaderInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-editor-modal/div/div[2]/div/form/div/div/div[7]/div/quill-editor/div[2]/div[1]")]
        private IWebElement TemplateFooterInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step1s/div/form/div/div/div[4]/input")]
        private IWebElement BEsstInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[5]/input")]
        private IWebElement StoreAddress1Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[6]/input")]
        private IWebElement StoreAddress2Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/form/div/div/div[8]/input")]
        private IWebElement ExternalCodeInput { get; set; }
        

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[3]/input")]
        private IWebElement BEemailInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[4]/div[1]/input")]
        private IWebElement BECityInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[4]/div[3]/input")]
        private IWebElement BEPosCodeInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[6]/input")]
        private IWebElement BEAddress1Input { get; set; }
        
        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[7]/input")]
        private IWebElement BEAddress2Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[8]/input")]
        private IWebElement BEAddress3Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[3]/button")]
        private IWebElement SaveButton { get; set; }

        // Methods
        public void SearchTemplate(string searchText)
        {
            var searchBox = new WebDriverWait(_driver, TimeSpan.FromSeconds(5))
                .Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"kt_filter_search\"]")));

            searchBox.Clear();
            searchBox.SendKeys(searchText);

        }

        public void ClickNewButton()
        {
            var newButton = _wait.Until(ExpectedConditions.ElementToBeClickable(NewButton));
            newButton.Click();
        }

        public void ClickNewReportTemplateButton()
        {
            var newReportButton = _wait.Until(ExpectedConditions.ElementToBeClickable(NewReportButton));
            newReportButton.Click();
        }

        public void ClickEditButton(string TemplateCode)
        {
            // Search by entering BETinNumber
            var searchInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//input[@type='text' and @placeholder='Search']")));
            searchInput.Clear();
            searchInput.SendKeys(TemplateCode);
            searchInput.SendKeys(Keys.Enter);
            Thread.Sleep(2500);


            // Wait for the row that contains the BETinNumber
            string rowXpath = $"//table/tbody/tr[td[contains(normalize-space(), '{TemplateCode}')]]";
            var row = _wait.Until(driver =>
            {
                var rows = driver.FindElements(By.XPath(rowXpath));
                return rows.Count == 1 ? rows[0] : null;
            });

            // Find the pencil icon inside that row (relative XPath)
            var editIcon = row.FindElement(By.CssSelector("i.bi.bi-pencil"));

            // Wait for it to be clickable and click
            _wait.Until(ExpectedConditions.ElementToBeClickable(editIcon)).Click();
        }

        public void ClickEditReportButton(string TemplateCode)
        {
            // Search by entering BETinNumber
            var searchInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//input[@type='text' and @placeholder='Search']")));
            searchInput.Clear();
            searchInput.SendKeys(TemplateCode);
            searchInput.SendKeys(Keys.Enter);
            Thread.Sleep(2500);

            // Wait for the row that contains the BETinNumber
            string rowXpath = $"//table/tbody/tr[td[contains(normalize-space(), '{TemplateCode}')]]";
            var row = _wait.Until(driver =>
            {
                var rows = driver.FindElements(By.XPath(rowXpath));
                return rows.Count == 1 ? rows[0] : null;
            });

            // Find the pencil icon inside that row (relative XPath)
            var editIcon = row.FindElement(By.XPath(".//i[contains(@class,'bi-pencil')]"));

            // Wait for it to be clickable and click
            _wait.Until(ExpectedConditions.ElementToBeClickable(editIcon)).Click();
            Thread.Sleep(2500);

        }

        // New Methods for Form Fields
        public void EnterTemplatename(string Templatename)
        {
            TemplatenameInput.Clear();
            TemplatenameInput.SendKeys(Templatename);
        }

        public void EnterReportTemplatename(string ReportTemplatename)
        {
            ReportTemplatenameInput.Clear();
            ReportTemplatenameInput.SendKeys(ReportTemplatename);
        }

        public void EnterTemplateDesc(string TemplateDesc)
        {
            TemplateDescInput.Clear();
            TemplateDescInput.SendKeys(TemplateDesc);
        }

        public void EnterReportTemplateDesc(string ReportTemplateDesc)
        {
            ReportTemplateDescInput.Clear();
            ReportTemplateDescInput.SendKeys(ReportTemplateDesc);
        }


        public void EnterTemplateHeader(string TemplateHeader)
        {
            TemplateHeaderInput.Clear();
            TemplateHeaderInput.SendKeys(TemplateHeader);
        }

        public void EnterTemplateFooter(string TemplateFooter)
        {
            TemplateFooterInput.Clear();
            TemplateFooterInput.SendKeys(TemplateFooter);
        }
      
        public void ClickFilterALLCategoryButton()
        {
            var filterAllButton = _wait.Until(ExpectedConditions.ElementToBeClickable(FilterALLCategoryButton));
            filterAllButton.Click();
        }

        public void ClickFilterActiveCategoryButton()
        {
            var filterActiveButton = _wait.Until(ExpectedConditions.ElementToBeClickable(FilterActiveCategoryButton));
            filterActiveButton.Click();
        }

        public void ClickFilterInactiveCategoryButton()
        {
            var filterInactiveButton = _wait.Until(ExpectedConditions.ElementToBeClickable(FilterInactiveCategoryButton));
            filterInactiveButton.Click();
        }
      
        public void SetCheckboxState(bool isChecked)
        {
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                By.XPath("//*[@id=\"kt_content_container\"]/app-template-editor/app-pdf-editor/div/div[1]/div/ul/li[4]/div/div/label/input")
            ));

            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }
        }

        public void SetReportCheckboxStatus(bool isChecked)
        {
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                By.XPath("//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-report-editor-modal/div/div[2]/div[1]/div[3]/div[2]/div/label/input")
            ));

            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }
        }

        public void SetAllReportCheckboxStatus(bool isChecked)
        {
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                By.XPath("//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-report-editor-modal/div/div[2]/div[3]/div[2]/button[1]")
            ));

            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }
        }

        public void SetClearReportCheckboxStatus(bool isChecked)
        {
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                By.XPath("//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-report-editor-modal/div/div[2]/div[3]/div[2]/button[2]")
            ));

            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }
        }

    }
}

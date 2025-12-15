using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace SeleniumTests.Pages.Log


{
    public class LogPage
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        // Constructor
        public LogPage(IWebDriver driver)
        {
            _driver = driver;
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));
            PageFactory.InitElements(_driver, this);
        }

        // Elements
        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > form > div > app-user-table > div > div > div.d-flex.gap-2.justify-content-start > div > ul > li.nav-item.ps-9.pt-3 > a")]
        private IWebElement FilterALLCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > form > div > app-user-table > div > div > div.d-flex.gap-2.justify-content-start > div > ul > li:nth-child(2) > a")]
        private IWebElement FilterActiveCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > form > div > app-user-table > div > div > div.d-flex.gap-2.justify-content-start > div > ul > li:nth-child(3) > a")]
        private IWebElement FilterInactiveCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(2) > span")]
        private IWebElement FilterPendingCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(3) > span")]
        private IWebElement FilterSuccessCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > div > div.d-flex.gap-2.justify-content-start > div > ul > li.ms-auto.align-items-center.pt-3.mobileHideFilter > div > div:nth-child(4) > span")]
        private IWebElement FilterFailedCategoryButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a")]
        private IWebElement ImportButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_body > ngb-modal-window > div > div > app-upload-modal > div > div.modal-footer.justify-content-end.d-flex.ng-star-inserted > button.btn.btn-primary.mx-2")]
        private IWebElement UploadButton { get; set; }


        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-log-details > div > div.card-header.border-0.pt-5.gap-2 > div > div:nth-child(2) > a")]
        private IWebElement ExportButton { get; set; }

        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-log-details > div > div.card-header.border-0.pt-5.gap-2 > div > div:nth-child(1) > a")]
        private IWebElement ResetButton { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/app-layout/div/div/div/div/app-content/app-user/form/div/app-user-table/div/div/div[2]/div/div[2]/a")]
        private IWebElement NewButton { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/app-layout/div/div/div/div/app-content/app-user/form/div/app-user-role-table/div/div[1]/div/div/a")]
        private IWebElement NewUserRoleButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-user-modal/div/div[3]/div/div[2]/button")]
        private IWebElement ContinueButton { get; set; }

        // New Elements
        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[2]/div[2]/div[1]/input")]
        public IWebElement UsernameInput { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_profile_details_view\"]/div[2]/form/div/div[1]/div[1]/div/input")]
        public IWebElement RolenameInput { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_profile_details_view\"]/div[2]/form/div/div[2]/div/div/input")]
        public IWebElement RoleDescInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[2]/div[2]/div[2]/input")]
        public IWebElement CustEmailInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[1]/div[1]/input")]
        private IWebElement UserPasswordInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div/input")]
        private IWebElement UserConfirmPasswordInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step1/div/form/div/div/div[5]/input")]
        private IWebElement CustsstInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[1]/input")]
        private IWebElement CustContactNumberInput { get; set; }
        

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-store-modal/div/div[2]/div/div/div[2]/div/app-step2s/div/div/form/div/div/div[3]/input")]
        private IWebElement BEemailInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[1]/input")]
        private IWebElement UserCityInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[2]/div[3]/input")]
        private IWebElement CustPosCodeInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[4]/input")]
        private IWebElement CustAddress1Input { get; set; }
        
        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[5]/input")]
        private IWebElement CustAddress2Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[6]/input")]
        private IWebElement CustAddress3Input { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/div/div[2]/div/app-step2/div/form/div/div/div[7]/input")]
        private IWebElement CustExternalCodeInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[3]/button")]
        private IWebElement SaveButton { get; set; }

        // Methods
        public void SearchLog(string searchText)
        {
            var searchBox = new WebDriverWait(_driver, TimeSpan.FromSeconds(5))
                .Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(
                    By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-log-details/div/div[2]/h3/div/div/input")));

            searchBox.Clear();
            searchBox.SendKeys(searchText);
        }
      

        public void ClickExportButton()
        {
            ExportButton.Click();
        }

        public void ClickResetButton()
        {
            ResetButton.Click();
        }

        
    }
}

using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace SeleniumTests.Pages.Setting


{
    public class SettingPage
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        // Constructor
        public SettingPage(IWebDriver driver)
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


        [FindsBy(How = How.CssSelector, Using = "#kt_content_container > app-user > form > div > app-user-table > div > div > div.card-header.border-0.pt-5 > div > div:nth-child(1) > a")]
        private IWebElement ExportButton { get; set; }


        [FindsBy(How = How.XPath, Using = "/html/body/app-layout/div/div/div/div/app-content/app-user/form/div/app-user-table/div/div/div[2]/div/div[2]/a")]
        private IWebElement NewButton { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/app-layout/div/div/div/div/app-content/app-user/form/div/app-user-role-table/div/div[1]/div/div/a")]
        private IWebElement NewUserRoleButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_body\"]/ngb-modal-window/div/div/app-user-modal/div/div[3]/div/div[2]/button")]
        private IWebElement ContinueButton { get; set; }

        // New Elements
        [FindsBy(How = How.CssSelector, Using = "input[name='consolidatePage']")]
        public IWebElement ConvertCutOffInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/app-layout/div[1]/div/div/app-header/div/app-topbar/app-side-setting/div/div[2]/div[2]/input")]
        public IWebElement ConvertCutOffInputQS { get; set; }

        [FindsBy(How = How.CssSelector, Using = "input[name='consolidateB2CPage']")]
        public IWebElement ConsolidateCutOffInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/app-layout/div[1]/div/div/app-header/div/app-topbar/app-side-setting/div/div[2]/div[3]/input")]
        public IWebElement ConsolidateCutOffInputQS { get; set; }

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

        [FindsBy(How = How.XPath, Using = "//button[contains(@class,'btn-sm btn-primary') and normalize-space(text())='Save Changes']")]
        public IWebElement SaveButton { get; set; }



        [FindsBy(How = How.XPath, Using = "/html/body/app-layout/div/div/div/app-header/div/app-topbar/app-side-setting/div/div[3]/div[2]/button[2]")]
        private IWebElement SaveButtonQS { get; set; }

       
        public void EnterConvertCutOff(string ConvertCutOff)
        {
            ConvertCutOffInput.Clear();
            ConvertCutOffInput.SendKeys(ConvertCutOff);
        }

        public void EnterConvertCutOffQS(string ConvertCutOffQS)
        {
            ConvertCutOffInputQS.Clear();
            ConvertCutOffInputQS.SendKeys(ConvertCutOffQS);
        }

        public void EnterCosolidateCutOff(string ConsolidateCutOff)
        {
            ConsolidateCutOffInput.Clear();
            ConsolidateCutOffInput.SendKeys(ConsolidateCutOff);
        }

        public void EnterCosolidateCutOffQS(string ConsolidateCutOffQS)
        {
            ConsolidateCutOffInputQS.Clear();
            ConsolidateCutOffInputQS.SendKeys(ConsolidateCutOffQS);
        }

        public void ClickSaveButton()
        {
            var saveButton = _wait.Until(ExpectedConditions.ElementToBeClickable(SaveButton));
            saveButton.Click();
        }

        public void ClickSaveButtonQS()
        {
            var saveButtonQS = _wait.Until(ExpectedConditions.ElementToBeClickable(SaveButtonQS));
            saveButtonQS.Click();
        }

        public void SetCheckboxState(bool isChecked)
        {
            // Wait until the checkbox is present
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-setting/app-general-setting/div[3]/div/div/div[2]/div[2]/label/input")));

            // Check or uncheck based on input
            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }

        }


        public void SetCheckboxStateQS(bool isChecked)
        {
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                By.XPath("/html/body/app-layout/div[1]/div/div/app-header/div/app-topbar/app-side-setting/div/div[2]/div[4]/label[2]/input")
            ));

            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }
        }       
    }
}

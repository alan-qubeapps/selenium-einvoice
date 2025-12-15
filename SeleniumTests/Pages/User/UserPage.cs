using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace SeleniumTests.Pages.User


{
    public class UserPage
    {
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        // Constructor
        public UserPage(IWebDriver driver)
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
        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[2]/div[2]/div[1]/input")]
        public IWebElement UsernameInput { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_profile_details_view\"]/div[2]/form/div/div[1]/div[1]/div/input")]
        public IWebElement RolenameInput { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_profile_details_view\"]/div[2]/form/div/div[2]/div/div/input")]
        public IWebElement RoleDescInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[2]/div[2]/div[2]/input")]
        public IWebElement CustEmailInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[1]/div/input")]
        private IWebElement UserCurrentPasswordInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[1]/div/input")]
        private IWebElement UserPasswordInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div/input")]
        private IWebElement UserConfirmPasswordInput { get; set; }


        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[1]/div/input")]
        private IWebElement UserProfilePasswordInput { get; set; }

        [FindsBy(How = How.XPath, Using = "/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[2]/div[2]/div/input")]
        private IWebElement UserProfileConfirmPasswordInput { get; set; }


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

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_content_container\"]/app-user/form/div/app-user-role-table/div/app-user-role-details/div/div[2]/div[2]/div[2]/div/a")]
        private IWebElement SignoutsessionButton { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@id=\"kt_content_container\"]/app-profile-details/div/div[2]/div/div/div[1]/div[2]")]
        private IWebElement SaveUserRoleButton { get; set; }

        // Methods
        public void SearchUser(string searchText)
        {
            var searchBox = new WebDriverWait(_driver, TimeSpan.FromSeconds(5))
                .Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/app-layout/div/div/div/div/app-content/app-user/form/div/app-user-table/div/div/div[2]/h3/div/input")));

            searchBox.Clear();
            searchBox.SendKeys(searchText);

        }

        public void SearchUserRole(string searchText)
        {
            var searchBox = new WebDriverWait(_driver, TimeSpan.FromSeconds(5))
                .Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/app-layout/div[1]/div/div/div/app-content/app-user/form/div/app-user-role-table/div/div[1]/h3/div/input")));

            searchBox.Clear();
            searchBox.SendKeys(searchText);

        }

        public void ClickNewButton()
        {
            var newButton = _wait.Until(ExpectedConditions.ElementToBeClickable(NewButton));
            newButton.Click();
        }

        public void ClickNewUserRoleButton()
        {
            var newUserRoleButton = _wait.Until(ExpectedConditions.ElementToBeClickable(NewUserRoleButton));
            newUserRoleButton.Click();
        }

        public void ClickEditButton(string Rolename)
        {
            // Search for the role using the search input
            var searchInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//input[@type='text' and @placeholder='Search']")));
            searchInput.Clear();
            searchInput.SendKeys(Rolename);
            searchInput.SendKeys(Keys.Enter); // Trigger AJAX or filtering

            // Wait until the specific row with the Rolename appears
            string rowXpath = $"//table/tbody/tr[td[contains(normalize-space(), '{Rolename}')]]";
            var row = _wait.Until(driver =>
            {
                var rows = driver.FindElements(By.XPath(rowXpath));
                return rows.Count > 0 ? rows[0] : null;
            });

            // Within that row, find the correct Edit icon in td[7]
            var editIcon = row.FindElement(By.XPath(".//td[7]//a/i"));

            // Click the Edit icon
            _wait.Until(ExpectedConditions.ElementToBeClickable(editIcon)).Click();
        }

        public void ClickEditUserButton(string User)
        {
            // Search input
            var searchInput = _wait.Until(ExpectedConditions.ElementIsVisible(
                By.XPath("//input[@type='text' and @placeholder='Search']")));
            searchInput.Clear();
            searchInput.SendKeys(User);
            searchInput.SendKeys(Keys.Enter);

            // Row detection (dynamic span text)
            string rowXpath = $"//table/tbody/tr[td//span[contains(normalize-space(), '{User}')]]";

            var row = _wait.Until(driver =>
            {
                var rows = driver.FindElements(By.XPath(rowXpath));
                return rows.Count > 0 ? rows[0] : null;
            });

            // Locate edit icon inside row
            var editIcon = row.FindElement(By.CssSelector("i.bi.bi-pencil"));

            // ===== SAFE CLICK =====
            try
            {
                // Scroll into view
                ((IJavaScriptExecutor)_driver).ExecuteScript(
                    "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                    editIcon);

                Thread.Sleep(250); // stabilize

                _wait.Until(ExpectedConditions.ElementToBeClickable(editIcon)).Click();
            }
            catch
            {
                // fallback JS click
                ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", editIcon);
            }
        }


        public void ClickExportButton()
        {
            ExportButton.Click();
        }

        // New Methods for Form Fields
        public void EnterUsername(string Username)
        {
            UsernameInput.Clear();
            UsernameInput.SendKeys(Username);
        }

        public void EnterRolename(string Rolename)
        {
            RolenameInput.Clear();
            RolenameInput.SendKeys(Rolename);
        }

        public void EnterRoleDesc(string RoleDesc)
        {
            RoleDescInput.Clear();
            RoleDescInput.SendKeys(RoleDesc);
        }

        public void EnterCustEmail(string CustEmail)
        {
            CustEmailInput.Clear();
            CustEmailInput.SendKeys(CustEmail);
        }

        public void EnterUserCurrentPassword(string UserCurrentPassword)
        {
            UserCurrentPasswordInput.Clear();
            UserCurrentPasswordInput.SendKeys(UserCurrentPassword);
        }
        public void EnterUserPassword(string UserPassword)
        {
            UserPasswordInput.Clear();
            UserPasswordInput.SendKeys(UserPassword);
        }

        public void EnterUserConfirmPassword(string UserConfirmPassword)
        {
            UserConfirmPasswordInput.Clear();
            UserConfirmPasswordInput.SendKeys(UserConfirmPassword);
        }

        public void EnterUserProfilePassword(string UserProfilePassword)
        {
            UserProfilePasswordInput.Clear();
            UserProfilePasswordInput.SendKeys(UserProfilePassword);
        }

        public void EnterUserProfileConfirmPassword(string UserProfileConfirmPassword)
        {
            UserProfileConfirmPasswordInput.Clear();
            UserProfileConfirmPasswordInput.SendKeys(UserProfileConfirmPassword);
        }

        public void ClickSaveButton()
        {
            var saveButton = _wait.Until(ExpectedConditions.ElementToBeClickable(SaveButton));
            saveButton.Click();
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

        public void SetCheckboxByLabel(string labelText)
        {
            try
            {
                // Uncheck all visible checkboxes on the screen first
                var allCheckboxes = _driver.FindElements(By.XPath("//input[@type='checkbox' and not(@disabled)]"));

                foreach (var cb in allCheckboxes)
                {
                    if (cb.Displayed && cb.Selected)
                    {
                        cb.Click();
                    }
                }

                // Find the specific checkbox using label text and click it
                var checkbox = _wait.Until(ExpectedConditions.ElementIsVisible(
                    By.XPath($"//label[contains(., '{labelText}')]/preceding-sibling::input[@type='checkbox']")));

                if (!checkbox.Selected)
                {
                    _wait.Until(ExpectedConditions.ElementToBeClickable(checkbox)).Click();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"❌ Failed to interact with checkbox '{labelText}': {ex.Message}");
            }
        }



        public void SetEditStoreCheckboxState(bool isChecked)
        {
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                By.XPath("/html/body/ngb-modal-window/div/div/app-role-modal/div/div[2]/div/div/div/h2/div/button/input")
            ));

            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }
        }

        public void SetCheckActiveboxState(bool isChecked)
        {
            var checkbox = _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(
                By.XPath("/html/body/ngb-modal-window/div/div/app-user-modal/div/div[2]/div/form/div[3]/div[2]/div[3]/div/label/input")
            ));

            if (checkbox.Selected != isChecked)
            {
                checkbox.Click();
            }
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

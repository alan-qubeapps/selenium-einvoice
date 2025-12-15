using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace SeleniumTests.Pages.Login
{
    public class LoginPage
    {
        private readonly IWebDriver driver;
        private readonly IWebDriver _driver;
        private readonly WebDriverWait _wait;

        [FindsBy(How = How.CssSelector, Using = "input[formcontrolname='email']")]
        private IWebElement UsernameField;

        [FindsBy(How = How.CssSelector, Using = "input[formcontrolname='password']")]
        private IWebElement PasswordField;

        [FindsBy(How = How.CssSelector, Using = "#kt_login_password_reset_form > div.fv-row.mb-10 > input")]
        private IWebElement EmailField;

        [FindsBy(How = How.Id, Using = "kt_sign_in_submit")]
        private IWebElement LoginButton;

        [FindsBy(How = How.XPath, Using = "//app-login//form//a[contains(text(),'Forgot Password')]")]
        private IWebElement ForgotPassword;

        [FindsBy(How = How.CssSelector, Using = "#kt_password_reset_submit")]
        private IWebElement SubmitForgotPassword;

        public LoginPage(IWebDriver driver, WebDriverWait wait)
        {
            _driver = driver;
            _wait = wait;
            PageFactory.InitElements(_driver, this);
        }

        public void EnterUsername(string username)
        {
            UsernameField.Clear();
            UsernameField.SendKeys(username);
        }

        public void EnterPassword(string password)
        {
            PasswordField.Clear();
            PasswordField.SendKeys(password);
        }

        public void EnterEmail(string email)
        {
            EmailField.Clear();
            EmailField.SendKeys(email);
        }

        public void ClickLoginButton()
        {
            LoginButton.Click();
        }

        public void ClickForgotPassword()
        {
            ForgotPassword.Click();
        }

        public void ClickSubmitForgotPassword()
        {
            SubmitForgotPassword.Click();
        }

        public bool IsLoginButtonEnabled()
        {
            return LoginButton.Enabled;
        }

        public bool IsSubmitForgotPasswordEnabled()
        {
            return SubmitForgotPassword.Enabled;
        }


        [FindsBy(How = How.CssSelector, Using = ".invalid-feedback")]
        private IList<IWebElement> ValidationMessages;
        public IList<string> GetValidationMessages()
        {
            List<string> messages = new List<string>();

            foreach (var messageElement in ValidationMessages)
            {
                messages.Add(messageElement.Text);
            }

            return messages;
        }

    }
}

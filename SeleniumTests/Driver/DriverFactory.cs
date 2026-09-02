using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace EInvoice.SeleniumTests.Drivers
{
    public static class DriverFactory
    {
        public static IWebDriver CreateDriver()
        {
            try
            {
                // Try normal Google Chrome first
                return CreateNormalChromeDriver();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Normal Chrome could not be started.");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine("Trying Chrome Beta...");

                try
                {
                    // Fallback to Chrome Beta
                    return CreateChromeBetaDriver();
                }
                catch (Exception betaEx)
                {
                    throw new Exception(
                        "Unable to start both Google Chrome and Chrome Beta.\n" +
                        $"Normal Chrome Error: {ex.Message}\n" +
                        $"Chrome Beta Error: {betaEx.Message}",
                        betaEx);
                }
            }
        }

        private static IWebDriver CreateNormalChromeDriver()
        {
            // Automatically download and use the correct ChromeDriver version
            new DriverManager().SetUpDriver(new ChromeConfig());

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");

            return new ChromeDriver(options);
        }

        private static IWebDriver CreateChromeBetaDriver()
        {
            // Automatically download and use the correct ChromeDriver version
            new DriverManager().SetUpDriver(new ChromeConfig());

            ChromeOptions options = new ChromeOptions();

            options.BinaryLocation =
                @"C:\Program Files\Google\Chrome Beta\Application\chrome.exe";

            options.AddArgument("--start-maximized");

            return new ChromeDriver(options);
        }

    }

}

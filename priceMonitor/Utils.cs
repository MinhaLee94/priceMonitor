using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace priceMonitor {
    static class Utils {
        /// <summary>
        /// Selenium search method with null exception handling
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        /// <returns>element</returns>
        public static IWebElement findElement(this IWebDriver driver, By by) {
            try {
                return driver.FindElement(by);
            } catch (NoSuchElementException) {
                Console.WriteLine("No such element found");
                return null;
            }
        }

        /// <summary>
        /// Check if Selenium has successfully searched a certain element
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static bool elementExists(this IWebElement element) {
            return element == null ? false : true;
        }

        // Chrome configuration methods
        public static ChromeDriverService configChromeDriverService() {
            ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            return driverService;
        }

        public static ChromeOptions configChromeOptions() {
            ChromeOptions options = new ChromeOptions();

            // disable unnecessary preference setting
            options.AddUserProfilePreference("cookies", 2);
            options.AddUserProfilePreference("images", 2);
            options.AddUserProfilePreference("popups", 2);
            options.AddUserProfilePreference("geolocation", 2);
            options.AddUserProfilePreference("notifications", 2);
            options.AddUserProfilePreference("auto_select_certificate", 2);
            options.AddUserProfilePreference("fullscreen", 2);
            options.AddUserProfilePreference("mouselock", 2);
            options.AddUserProfilePreference("mixed_script", 2);
            options.AddUserProfilePreference("media_stream", 2);
            options.AddUserProfilePreference("media_stream_mic", 2);
            options.AddUserProfilePreference("media_stream_camera", 2);
            options.AddUserProfilePreference("ppapi_broker", 2);
            options.AddUserProfilePreference("automatic_downloads", 2);
            options.AddUserProfilePreference("midi_sysex", 2);
            options.AddUserProfilePreference("push_messaging", 2);
            options.AddUserProfilePreference("ssl_cert_decisions", 2);
            options.AddUserProfilePreference("metro_switch_to_desktop", 2);
            options.AddUserProfilePreference("protected_media_identifier", 2);
            options.AddUserProfilePreference("app_banner", 2);
            options.AddUserProfilePreference("site_engagement", 2);
            options.AddUserProfilePreference("durable_storage", 2);
            options.AddExcludedArgument("enable-automation");

            options.AddArgument("--disable-extensions");
            options.AddArgument("--disable-gpu");
            options.AddArgument("--disable-infobars");
            options.AddArgument("--disable-dev-shm-usage");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--ignore-certificate-errors");

            return options;
        }
    }
}

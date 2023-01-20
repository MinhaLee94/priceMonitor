using System;
using System.Collections.Generic;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Runtime.InteropServices;

namespace priceMonitor {
    static class Utils {
        public static IWebElement findElement(this IWebDriver driver, By by) {
            try {
                return driver.FindElement(by);
            } catch (NoSuchElementException) {
                Console.WriteLine("No such element found");
                return null;
            }
        }

        public static bool elementExists(this IWebElement element) {
            return element == null ? false : true;
        }
    }
}

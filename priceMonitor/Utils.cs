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
        /// <summary>
        /// Selenium search method with null exception handling
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="by"></param>
        /// <returns></returns>
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
    }
}

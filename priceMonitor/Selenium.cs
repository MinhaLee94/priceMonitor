using System;
using System.Collections.Generic;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace priceMonitor {
    class Selenium {
        public static void scrapStaplesItems(List<Dictionary<string, string>> itemsToSearch, ChromeDriverService driverService, ChromeOptions options) {
            foreach (var item in itemsToSearch) {
                try {
                    string curResellerSku = item["Reseller SKU"];
                    Console.WriteLine("Reseller SKU: " + curResellerSku);

                    using (IWebDriver driver = new ChromeDriver(driverService, options)) {
                        driver.Url = $"https://www.staples.com/{curResellerSku}/directory_{curResellerSku}";
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
                        IWebElement curPrice = Utils.findElement(driver, By.CssSelector(".price-info__final_price_sku"));
                        IWebElement outOfStockSign = Utils.findElement(driver, By.XPath("//*[@id='ONE_TIME_PURCHASE']/div/div/div/div/div/div/div[2]/div"));

                        if (Utils.elementExists(curPrice)) {
                            if (Utils.elementExists(outOfStockSign) && outOfStockSign.Text.ToUpper() == "THIS ITEM IS OUT OF STOCK") {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "OUT OF STOCK");
                            } else {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "ON SITE");
                            }
                        } else {
                            item.Add("Real Price", "");
                            item.Add("Status", "OFF SITE");
                        }
                    }
                } catch (Exception ex) {
                    Console.WriteLine(ex);
                    continue;
                }
            }
        }

        public static void scrapNewEggItems(List<Dictionary<string, string>> itemsToSearch, ChromeDriverService driverService, ChromeOptions options) {
            foreach (var item in itemsToSearch) {
                try {
                    string curResellerSku = item["Reseller SKU"];
                    Console.WriteLine("Reseller SKU: " + curResellerSku);

                    using (IWebDriver driver = new ChromeDriver(driverService, options)) {
                        driver.Url = $"https://www.newegg.com/p/{curResellerSku}?item={curResellerSku}";
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
                        IWebElement curPrice = Utils.findElement(driver, By.CssSelector(".product-buy-box .price-current"));
                        IWebElement outOfStockSign = Utils.findElement(driver, By.CssSelector(".product-buy-box .btn-message"));

                        if (Utils.elementExists(curPrice)) {
                            item.Add("Real Price", curPrice.Text);
                            item.Add("Status", "ON SITE");
                        } else if (Utils.elementExists(outOfStockSign) && outOfStockSign.Text.ToUpper() == "OUT OF STOCK") {
                            item.Add("Real Price", "");
                            item.Add("Status", "OUT OF STOCK");
                        } else {
                            item.Add("Real Price", "");
                            item.Add("Status", "OFF SITE");
                        }
                    }
                } catch (Exception ex) {
                    Console.WriteLine(ex);
                    continue;
                }
            }
        }

        public static void scrapOfficeDepotItems(List<Dictionary<string, string>> itemsToSearch, ChromeDriverService driverService, ChromeOptions options) {
            foreach (var item in itemsToSearch) {
                try {
                    string curResellerSku = item["Reseller SKU"];
                    Console.WriteLine("Reseller SKU: " + curResellerSku);

                    using (IWebDriver driver = new ChromeDriver(driverService, options)) {
                        driver.Url = $"https://www.officedepot.com/catalog/catalogSku.do?id={curResellerSku}";
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
                        IWebElement curPrice = Utils.findElement(driver, By.CssSelector(".od-graphql-price-big-price"));
                        IWebElement outOfStockSign = Utils.findElement(driver, By.CssSelector(".od-fulfillment-option-in-stock-message-text-oos"));

                        if (Utils.elementExists(curPrice)) {
                            if (Utils.elementExists(outOfStockSign) && outOfStockSign.Text.ToUpper() == "OUT OF STOCK") {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "OUT OF STOCK");
                            } else {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "ON SITE");
                            }
                        } else {
                            item.Add("Real Price", "");
                            item.Add("Status", "OFF SITE");
                        }
                    }
                } catch (Exception ex) {
                    Console.WriteLine(ex);
                    continue;
                }
            }
        }

        public static void scrapTargetPlusItems(List<Dictionary<string, string>> itemsToSearch, ChromeDriverService driverService, ChromeOptions options) {
            foreach (var item in itemsToSearch) {
                try {
                    string curResellerSku = item["Reseller SKU"];
                    Console.WriteLine("Reseller SKU: " + curResellerSku);

                    using (IWebDriver driver = new ChromeDriver(driverService, options)) {
                        driver.Url = $"https://www.target.com/s?searchTerm={curResellerSku}";
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
                        IWebElement curPrice = Utils.findElement(driver, By.CssSelector(".h-padding-r-tiny"));
                        IWebElement outOfStockSign = Utils.findElement(driver, By.CssSelector($"#addToCartButtonOrTextIdFor{curResellerSku}"));

                        if (Utils.elementExists(curPrice)) {
                            if (Utils.elementExists(outOfStockSign) && outOfStockSign.Text.ToUpper() == "SOLD OUT") {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "OUT OF STOCK");
                            } else {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "ON SITE");
                            }
                        } else {
                            item.Add("Real Price", "");
                            item.Add("Status", "OFF SITE");
                        }
                    }
                } catch (Exception ex) {
                    Console.WriteLine(ex);
                    continue;
                }
            }
        }

        public static void scrapBestBuyItems(List<Dictionary<string, string>> itemsToSearch, ChromeDriverService driverService, ChromeOptions options) {
            foreach (var item in itemsToSearch) {
                try {
                    string curResellerSku = item["Reseller SKU"];
                    Console.WriteLine("Reseller SKU: " + curResellerSku);

                    using (IWebDriver driver = new ChromeDriver(driverService, options)) {
                        driver.Url = $"https://www.bestbuy.com/site/searchpage.jsp?st={curResellerSku}";
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(10);
                        IWebElement curPrice = Utils.findElement(driver, By.CssSelector(".pricing-price .priceView-customer-price > span:first-child"));
                        IWebElement outOfStockSign = Utils.findElement(driver, By.CssSelector(".fulfillment-fulfillment-summary>div>div>div:nth-child(2)>div"));

                        if (Utils.elementExists(curPrice)) {
                            if (Utils.elementExists(outOfStockSign) && outOfStockSign.Text.ToUpper() == "SOLD OUT") {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "OUT OF STOCK");
                            } else {
                                item.Add("Real Price", curPrice.Text);
                                item.Add("Status", "ON SITE");
                            }
                        } else {
                            item.Add("Real Price", "");
                            item.Add("Status", "OFF SITE");
                        }
                    }
                } catch (Exception ex) {
                    Console.WriteLine(ex);
                    continue;
                }
            }
        }
    }
}

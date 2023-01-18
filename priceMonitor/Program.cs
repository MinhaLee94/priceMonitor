using System;
using HtmlAgilityPack;
using ExcelDataReader;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Collections;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

// transfer itemList to each store
// create url-formatting function and xlsx file generating function

namespace priceMonitor {
    class Program {
        private static ArrayList itemsToSearch = new ArrayList();
        private static readonly string directory = "C:\\Users\\john\\Desktop\\SKU Status";
        private static string[] fileEntries = Directory.GetFiles(directory);


        protected static ChromeDriverService driverService = null;
        protected static ChromeOptions options = null;
        protected static ChromeDriver driver = null;

        private static ArrayList generateItemListToSearch() {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTableCollection tablecollection;
            ArrayList itemList = new ArrayList();

            foreach (string fileName in fileEntries) {
                if (fileName != @"C:\Users\john\Desktop\SKU Status\Staples Report 010923.xlsx") continue; // for testing

                using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read)) {
                    using (var reader = ExcelReaderFactory.CreateReader(stream)) {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration() {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() {
                                EmptyColumnNamePrefix = "Column",
                                UseHeaderRow = true
                            }
                        });

                        tablecollection = result.Tables;

                        foreach (DataTable datatable in tablecollection) {
                            if (datatable.TableName != "DESKTOPS" && datatable.TableName != "LAPTOPS") continue;

                            foreach (DataRow item in datatable.Rows) {
                                //if (item["Status"].ToString() != "ON SITE") continue;

                                Dictionary<string, string> itemInfo = new Dictionary<string, string>() {
                                    {"Joy SKU", item["Joy SKU"].ToString()},
                                    {"Reseller SKU", item["Reseller SKU"].ToString()},
                                    {"Status", item["Status"].ToString()},
                                    {"C RP", item["C RP"].ToString()}
                                };
                                itemList.Add(itemInfo);
                            }
                        }
                    }
                }
            }
            return itemList;
        }

        static void Main(string[] args) {
            itemsToSearch = generateItemListToSearch();

            using (IWebDriver driver = new ChromeDriver()) {
                driver.Url = "https://www.staples.com/2446392/directory_2446392";
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                IWebElement element = driver.FindElement(By.CssSelector(".price-info__final_price_sku"));
                Console.WriteLine(element.Text);
            }

            /*            driverService = ChromeDriverService.CreateDefaultService();
                        driverService.HideCommandPromptWindow = true;

                        options = new ChromeOptions();
                        options.AddArgument("disable-gpu");
                        options.AddArgument("headless");

                        driver = new ChromeDriver(driverService, options);
                        driver.Navigate().GoToUrl(url);*/









            /*            HtmlWeb web = new HtmlWeb();
                        HtmlDocument htmlDoc = web.Load("https://www.staples.com/2446392/directory_2446392");

                        System.Diagnostics.Debug.WriteLine(htmlDoc.DocumentNode.OuterHtml);*/

        }
    }
}

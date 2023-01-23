using System;
using HtmlAgilityPack;
using ExcelDataReader;
using System.IO;
using System.Data;
using System.Collections.Generic;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Runtime.InteropServices;

namespace priceMonitor {
    class Program {
        private static readonly string directory = "C:\\Users\\john\\Desktop\\SKU Status";
        private static string[] fileEntries = Directory.GetFiles(directory);

        private static List<Dictionary<string, string>> itemsToSearch = new List<Dictionary<string, string>>();
        private static string vendor = "";

        private static Excel.Application excelApp = null;
        private static Excel.Workbook workBook = null;
        private static Excel.Worksheet workSheet = null;

        private static List<Dictionary<string, string>> generateItemListToSearch() {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTableCollection tablecollection;
            List<Dictionary<string, string>> itemList = new List<Dictionary<string, string>>();

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
                            if (datatable.TableName.ToUpper() != "DESKTOPS" && datatable.TableName.ToUpper() != "LAPTOPS") continue;

                            foreach (DataRow item in datatable.Rows) {
                                if (item["Status"].ToString() != "ON SITE") continue;

                                Dictionary<string, string> itemInfo = new Dictionary<string, string>() {
                                    {"Joy SKU", item["Joy SKU"].ToString()},
                                    {"Reseller SKU", item["Reseller SKU"].ToString()},
                                    {"C RP", item["C RP"].ToString()},
                                };

                                itemList.Add(itemInfo);
                            }
                        }
                    }
                }
            }
            return itemList;
        }

        private static void generateExcelFileWithSearchedResults(List<Dictionary<string, string>> SearchedResults) {
            try {
                string directoryPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string savePath = directoryPath + "\\" + "Staples status " + DateTime.Now.ToString("MMddyy") + ".xlsx";
                string[] headers = { "Joy SKU", "Reseller SKU", "C RP", "Real Price", "Status", "RP Change" };

                if (File.Exists(savePath)) File.Delete(savePath);

                excelApp = new Excel.Application();
                workBook = excelApp.Workbooks.Add();
                workSheet = workBook.Worksheets.get_Item(1) as Excel.Worksheet;

                // set headers
                for (int i = 0; i < headers.Length; i++) {
                    workSheet.Cells[1, i + 1] = headers[i];
                }

                for (int i = 0; i < SearchedResults.Count; i++) {
                    Dictionary<string, string> curItem = SearchedResults[i];

                    workSheet.Cells[2 + i, 1] = curItem["Joy SKU"];
                    workSheet.Cells[2 + i, 2] = curItem["Reseller SKU"];
                    workSheet.Cells[2 + i, 3] = curItem["C RP"];
                    workSheet.Cells[2 + i, 4] = curItem["Real Price"];
                    workSheet.Cells[2 + i, 5] = curItem["Status"];

                    if (curItem["C RP"] != curItem["Real Price"])
                        workSheet.Cells[2 + i, 6] = curItem["C RP"] + " > " + curItem["Real Price"];
                }

                workSheet.Columns.AutoFit();
                workBook.SaveAs(savePath, Excel.XlFileFormat.xlWorkbookDefault);
                workBook.Close(true);
                excelApp.Quit();
            } finally {
                ReleaseObject(workSheet);
                ReleaseObject(workBook);
                ReleaseObject(excelApp);
            }
            return;
        }

        /// <summary>
        /// Excel object release method
        /// </summary>
        /// <param name="obj"></param>
        public static void ReleaseObject(object obj) {
            if (obj == null) return;

            try {
                Marshal.ReleaseComObject(obj);
                obj = null;
            } catch (Exception ex) {
                obj = null;
                throw ex;
            } finally {
                GC.Collect();
            }
        }


        public static void checkInventoryAndPrice(string vendor) {
            ChromeDriverService driverService = Utils.configChromeDriverService();
            ChromeOptions options = Utils.configChromeOptions();

            try {
                foreach (var item in itemsToSearch) {
                    string curResellerSku = item["Reseller SKU"];
                    Console.WriteLine("Reseller SKU: " + curResellerSku);

                    using (IWebDriver driver = new ChromeDriver(driverService, options)) {
                        driver.Url = $"https://www.staples.com/{curResellerSku}/directory_{curResellerSku}";
                        IWebElement curPrice = Utils.findElement(driver, By.CssSelector(".price-info__final_price_sku"));
                        IWebElement outOfStockSign = Utils.findElement(driver, By.XPath("//*[@id='ONE_TIME_PURCHASE']/div/div/div/div/div/div/div[2]/div"));

                        if (Utils.elementExists(curPrice)) {
                            if (Utils.elementExists(outOfStockSign) && outOfStockSign.Text == "This item is out of stock") {
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
                }
            } catch (Exception ex) {
                Console.WriteLine(ex);
            }
        }

        static void Main(string[] args) {
            itemsToSearch = generateItemListToSearch();

            checkInventoryAndPrice("Staples");

            generateExcelFileWithSearchedResults(itemsToSearch);




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

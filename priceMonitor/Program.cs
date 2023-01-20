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

using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Runtime.InteropServices;

// transfer itemList to each store
// create url-formatting function and xlsx file generating function

namespace priceMonitor {
    class Program {
        private static List<Dictionary<string, string>> itemsToSearch = new List<Dictionary<string, string>>();
        private static readonly string directory = "C:\\Users\\john\\Desktop\\SKU Status";
        private static string[] fileEntries = Directory.GetFiles(directory);


        protected static ChromeDriverService driverService = null;
        protected static ChromeOptions options = null;
        protected static ChromeDriver driver = null;

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
                            if (datatable.TableName != "DESKTOPS" && datatable.TableName != "LAPTOPS") continue;

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
                string[] headers = { "Joy SKU", "Reseller SKU", "C RP" };

                if (File.Exists(savePath)) {
                    File.Delete(savePath);
                }

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

        static void Main(string[] args) {
            itemsToSearch = generateItemListToSearch();


            driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;

            options = new ChromeOptions();

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

            try {
                foreach (var item in itemsToSearch) {
                    string curResellerSku = item["Reseller SKU"];
                    Console.WriteLine("Reseller SKU: " + curResellerSku);

                    using (IWebDriver driver = new ChromeDriver(driverService, options)) {
                        driver.Url = $"https://www.staples.com/{curResellerSku}/directory_{curResellerSku}";
                        IWebElement element = driver.FindElement(By.CssSelector(".price-info__final_price_sku"));
                        Console.WriteLine(element.Text);
                        item.Add("Real Price", element.Text);
                    }
                }
            } catch(Exception ex) {
                Console.WriteLine(ex);
            }

            

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

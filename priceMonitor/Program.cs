using System;
using ExcelDataReader;
using System.IO;
using System.Data;
using System.Collections.Generic;

using OpenQA.Selenium.Chrome;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Linq;

namespace priceMonitor {
    class Program {
        private static readonly string directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        private static string[] fileEntries = Directory.GetFiles(directory + "\\SKU Status");
        private static List<Dictionary<string, string>> itemsToSearch = new List<Dictionary<string, string>>();
        private static string[] vendorList = new string[] { "STAPLES", "NEWEGG", "TARGETPLUS", "OFFICEDEPOT", "BESTBUY" };

        private static Excel.Application excelApp = null;
        private static Excel.Workbook workBook = null;
        private static Excel.Worksheet workSheet = null;

        private static List<Dictionary<string, string>> generateItemListToSearch(string filePath) {
            string directoryPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            Console.WriteLine(directoryPath);

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTableCollection tablecollection;
            List<Dictionary<string, string>> itemList = new List<Dictionary<string, string>>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read)) {
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
                            if (item["Status"].ToString().ToUpper() != "ON SITE") continue;

                            Dictionary<string, string> itemInfo = new Dictionary<string, string>() {
                                {"Joy SKU", item["Joy SKU"].ToString()},
                                {"Reseller SKU", item["Reseller SKU"].ToString()},
                                {"C RP", "$" + item["C RP"].ToString()},
                            };

                            itemList.Add(itemInfo);
                        }
                    }
                }
            }
            return itemList;
        }

        private static void generateExcelFileWithSearchedResults(List<Dictionary<string, string>> SearchedResults, string vendor) {
            try {
                string directoryPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string savePath = directoryPath + "\\scrap\\" + vendor + " status " + DateTime.Now.ToString("MMddyy") + ".xlsx";
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

                    if(curItem.ContainsKey("Real Price") && curItem.ContainsKey("Status")) {
                        workSheet.Cells[2 + i, 4] = curItem["Real Price"];
                        workSheet.Cells[2 + i, 5] = curItem["Status"];

                        if (curItem["Status"] != "OFF SITE" && curItem["C RP"] != curItem["Real Price"]) {
                            workSheet.Cells[2 + i, 6] = curItem["C RP"] + " > " + curItem["Real Price"];
                        }
                    }
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

            switch (vendor) {
                case "STAPLES":
                    Selenium.scrapStaplesItems(itemsToSearch, driverService, options);
                    break;
                case "NEWEGG":
                    Selenium.scrapNewEggItems(itemsToSearch, driverService, options);
                    break;
                case "TARGETPLUS":
                    Selenium.scrapTargetPlusItems(itemsToSearch, driverService, options);
                    break;
                case "OFFICEDEPOT":
                    Selenium.scrapOfficeDepotItems(itemsToSearch, driverService, options);
                    break;
                case "BESTBUY":
                    Selenium.scrapBestBuyItems(itemsToSearch, driverService, options);
                    break;
                default:
                    break;
            }
        }

        static void Main(string[] args) {
            string vendor = "";

            foreach (string filePath in fileEntries) {
                vendor = Path.GetFileName(filePath).Substring(0, Path.GetFileName(filePath).IndexOf(" ")).ToUpper();
                Console.WriteLine($"Searching into {vendor}");
                if (!vendorList.Contains(vendor)) continue;

                itemsToSearch = generateItemListToSearch(filePath);
                checkInventoryAndPrice(vendor);
                generateExcelFileWithSearchedResults(itemsToSearch, vendor);

                itemsToSearch.Clear();
                vendor = "";
            }
        }
    }
}

using System;
using HtmlAgilityPack;
using ExcelDataReader;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Collections;

// transfer itemList to each store
// create url-formatting function and xlsx file generating function

namespace priceMonitor {
    class Program {
        private static ArrayList itemsToSearch = new ArrayList();
        private static readonly string directory = "C:\\Users\\john\\Desktop\\SKU Status";
        private static string[] fileEntries = Directory.GetFiles(directory);

        private static ArrayList generateItemListToSearch() {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTableCollection tablecollection;
            DataTable dt;
            ArrayList itemList = new ArrayList();

            foreach (string fileName in fileEntries) {
                if (fileName != @"C:\Users\john\Desktop\SKU Status\Walmart Report 010923.xlsx") continue; // for testing

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
                            if (datatable.TableName.Contains("Discontinued") || datatable.TableName.Contains("No SKUs") || datatable.TableName == "susepnd") continue;

                            foreach (DataRow item in datatable.Rows) {
                                if (item["Status"].ToString() != "ON SITE") continue;

                                Dictionary<string, string> itemInfo = new Dictionary<string, string>() {
                                    {"Joy SKU", item["Joy SKU"].ToString()},
                                    {"Reseller SKU", item["Reseller SKU"].ToString()},
                                    {"Description", item["Description"].ToString()},
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

            /*string url = "https://www.walmart.com/ip/666786665";
            HtmlWeb web = new HtmlWeb();
            HtmlDocument htmlDoc = web.Load(url);*/

            //System.Diagnostics.Debug.WriteLine(htmlDoc.DocumentNode.OuterHtml);
            //Console.WriteLine(htmlDoc.Text);
        }
    }
}

using System;
using HtmlAgilityPack;

namespace priceMonitor {
    class Program {
        static void Main(string[] args) {
            string temp = "http://www.staples.com/product_2710522";
            HtmlWeb web = new HtmlWeb();
            HtmlDocument htmlDoc = web.Load(temp);

            System.Diagnostics.Debug.WriteLine(htmlDoc.DocumentNode.OuterHtml);
            Console.WriteLine(htmlDoc.Text);
        }
    }
}

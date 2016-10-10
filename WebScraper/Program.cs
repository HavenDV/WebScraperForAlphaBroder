﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading;
using System.Globalization;
using BigCommerce;

namespace WebScraper
{
    public static class StringExtensions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }
    }

    class Program
    {
        static IDictionary<string, int> itemCount = new Dictionary<string, int>();

        static string CreateLink(string categoryName, int resultsNumber = 1)
        {
            return string.Format(
                "https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults={1}&CatPath=All+Products////{0}////",
                categoryName,
                resultsNumber);
        }

        static string GetName(string data)
        {
            foreach (Match match in Regex.Matches(data, "<h1>(.+)<\\/h1>"))
            {
                return match.Groups[1].Value;
            }
            return string.Empty;
        }

        static string GetDescription(string data)
        {
            foreach (Match match in Regex.Matches(data, "<ul class=\\\"bullet\\\">((.|\\n)+?)<\\/ul>"))
            {
                return "<ul>\n"+match.Groups[1].Value+"</ul>";
            }
            return string.Empty;
        }

        static string GetPrice(string data)
        {
            var pattern = "<meta property=\\\"og\\:price\\:amount\\\" content=\\\"(\\S+)\\\">";
            foreach (Match match in Regex.Matches(data, pattern))
            {
                return match.Groups[1].Value;
            }
            return string.Empty;
        }
        
        static IList<string> GetImages(string data, string id)
        {
            var images = new List<string>();
            foreach (Match m in Regex.Matches(data, "//(\\S+?)\\.(jpg|png|gif|jpeg)"))
            {
                var value = m.Value;
                if (value.Contains(id, StringComparison.OrdinalIgnoreCase) && 
                    !value.Contains("large") && 
                    !value.Contains("grande") && 
                    !value.Contains("1024"))
                {
                    images.Add("http:"+ value);
                }
            }

            return images.Distinct().ToList();
        }

        static IList<string> GetColors(string data)
        {
            Uri uri = new Uri("http://google.com/search?q=test");
            Regex reHref = new Regex(@"(?inx)
    <a \s [^>]*
        href \s* = \s*
            (?<q> ['""] )
                (?<url> [^""]+ )
            \k<q>
    [^>]* >");
            var colors = new HashSet<string>();
            foreach (Match m in Regex.Matches(data, "//(\\S+?)\\.(jpg|png|gif|jpeg)"))
            {
                var value = m.Value;
                colors.Add(value);
            }

            return colors.ToList();
        }

        public static string ToRoman(int number)
        {
            if ((number < 0) || (number > 3999)) throw new ArgumentOutOfRangeException("insert value betwheen 1 and 3999");
            if (number < 1) return string.Empty;
            if (number >= 1000) return "M" + ToRoman(number - 1000);
            if (number >= 900) return "CM" + ToRoman(number - 900);
            if (number >= 500) return "D" + ToRoman(number - 500);
            if (number >= 400) return "CD" + ToRoman(number - 400);
            if (number >= 100) return "C" + ToRoman(number - 100);
            if (number >= 90) return "XC" + ToRoman(number - 90);
            if (number >= 50) return "L" + ToRoman(number - 50);
            if (number >= 40) return "XL" + ToRoman(number - 40);
            if (number >= 10) return "X" + ToRoman(number - 10);
            if (number >= 9) return "IX" + ToRoman(number - 9);
            if (number >= 5) return "V" + ToRoman(number - 5);
            if (number >= 4) return "IV" + ToRoman(number - 4);
            if (number >= 1) return "I" + ToRoman(number - 1);
            throw new ArgumentOutOfRangeException("something bad happened");
        }

        static string DownloadPage(string url, string team, string collection)
        {
            try
            {
                var data = WebUtilities.DownloadPage(url);
                var name = GetName(data);
                var desc = GetDescription(data);
                var price = "1.00";//GetPrice(data);

                Console.WriteLine(price);
                if (string.IsNullOrWhiteSpace(price))
                {
                    throw new Exception("Price cannot be null.");
                }

                var images = new List<string>();// GetImages(data, id);

                //Fix duplicates roman number method
                if (itemCount.ContainsKey(name))
                {
                    itemCount[name] = itemCount[name] + 1;
                    name += " " + ToRoman(itemCount[name]);
                }
                else
                {
                    itemCount.Add(name, 1);
                }

                return BigCommerceUtilities.ProductFormat(name, desc, "ALP", "", "", price, "", images, "", "");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return string.Empty;
        }
        
        static IList<string> GetTeams(string prefix)
        {
            var teams = new List<string>();
            var data = new WebClient().DownloadString("http://www.zhats.com/pages/" + prefix);

            foreach (Match match in Regex.Matches(data, "<li><a href=\\\"(http://zhats(.+))\\\">(.+)<\\/a><\\/li>"))
            {
                //Fix broken link
                var value = match.Groups[1].Value;
                if (value.Contains("Canacdiens"))
                {
                    value = value.Replace("Canacdiens", "Canadiens");
                }

                teams.Add(value);
            }

            return teams;
        }

        static IList<string> GetItems(string url)
        {
            var items = new List<string>();
            var data = WebUtilities.DownloadPage(url);

            //Console.WriteLine(data);
            //set UseDefaultCookiesParser as false if a website returns invalid cookies format
            //browser.UseDefaultCookiesParser = false;

            //var data = new WebClient().DownloadString(url);
            foreach (Match match in Regex.Matches(data, "<div class=\\\"rsltProdNameText\\\">(.+?)<\\/div>"))
            {
                var item = string.Format(
                    "https://www.alphabroder.com/cgi-bin/online/webshr/prod-labeldtl.w?sr={0}&currentColor=", 
                    match.Groups[1].Value);
                items.Add(item);
            }

            return items;
        }

        static IList<string> GetItemsMultipage(string url)
        {
            var items = new List<string>();
            IList<string> nextItems;
            var page = 0;
            do
            {
                ++page;
                nextItems = GetItems(url + "&currentpage=" + page);
                items.AddRange(nextItems);
            }
            while (nextItems.Count > 0);

            return items;
        }
        
        static StreamWriter CreateImportCSVFile(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var header = "\"Item Type\",\"Product ID\",\"Product Name\",\"Product Type\",\"Product Code/SKU\",\"Bin Picking Number\",\"Brand Name\",\"Option Set\",\"Option Set Align\",\"Product Description\",\"Price\",\"Cost Price\",\"Retail Price\",\"Sale Price\",\"Fixed Shipping Cost\",\"Free Shipping\",\"Product Warranty\",\"Product Weight\",\"Product Width\",\"Product Height\",\"Product Depth\",\"Allow Purchases?\",\"Product Visible?\",\"Product Availability\",\"Track Inventory\",\"Current Stock Level\",\"Low Stock Level\",\"Category\",\"Product Image ID - 1\",\"Product Image File - 1\",\"Product Image Description - 1\",\"Product Image Is Thumbnail - 1\",\"Product Image Sort - 1\",\"Product Image ID - 2\",\"Product Image File - 2\",\"Product Image Description - 2\",\"Product Image Is Thumbnail - 2\",\"Product Image Sort - 2\",\"Product Image ID - 3\",\"Product Image File - 3\",\"Product Image Description - 3\",\"Product Image Is Thumbnail - 3\",\"Product Image Sort - 3\",\"Product Image ID - 4\",\"Product Image File - 4\",\"Product Image Description - 4\",\"Product Image Is Thumbnail - 4\",\"Product Image Sort - 4\",\"Product Image ID - 5\",\"Product Image File - 5\",\"Product Image Description - 5\",\"Product Image Is Thumbnail - 5\",\"Product Image Sort - 5\",\"Search Keywords\",\"Page Title\",\"Meta Keywords\",\"Meta Description\",\"MYOB Asset Acct\",\"MYOB Income Acct\",\"MYOB Expense Acct\",\"Product Condition\",\"Show Product Condition?\",\"Event Date Required?\",\"Event Date Name\",\"Event Date Is Limited?\",\"Event Date Start Date\",\"Event Date End Date\",\"Sort Order\",\"Product Tax Class\",\"Product UPC/EAN\",\"Stop Processing Rules\",\"Product URL\",\"Redirect Old URL?\",\"GPS Global Trade Item Number\",\"GPS Manufacturer Part Number\",\"GPS Gender\",\"GPS Age Group\",\"GPS Color\",\"GPS Size\",\"GPS Material\",\"GPS Pattern\",\"GPS Item Group ID\",\"GPS Category\",\"GPS Enabled\",\"Avalara Product Tax Code\",\"Product Custom Fields\"";
            var file = File.CreateText(path);
            file.WriteLine(header);
            return file;
        }

        static async Task<string> DownloadItemsAsync(IList<string> items, string teamName, string collectionName)
        {
            Console.WriteLine("Start download {0} team: {1}. Size: {2}", collectionName, teamName, items.Count);
            return string.Join("\n", await Task.WhenAll(
                items.Select(item => 
                    Task.Run(() =>
                        DownloadPage(item, teamName, collectionName)
                ))));
        }

        static async Task<string> DownloadTeamsAsync(IList<string> teams, string collectionName)
        {
            Console.WriteLine("Start download collection: {0}.", collectionName);
            var count = 0;
            var strings = await Task.WhenAll(teams.Select(team => Task.Run(() =>
            {
                var teamName = Path.GetFileName(team);
                var items = GetItemsMultipage(team);
                count += items.Count;
                return DownloadItemsAsync(items, teamName, collectionName);
            })));
            Console.WriteLine("Download ended: {0}. Downloaded {1} items.", collectionName, count);
            return string.Join("\n", strings);
        }

        static void LoadCategory(string name, string to, string fullname)
        {
            var file = CreateImportCSVFile(Path.Combine(to, name + ".csv"));
            file.Write(DownloadTeamsAsync(GetTeams(fullname), name).Result);
            file.Close();
        }

        static void LoadCollection(string name, string to)
        {
            var url = CreateLink(name);
            var file = CreateImportCSVFile(Path.Combine(to, name + ".csv"));
            var items = GetItems(url);
            Console.WriteLine("Start download category: {0}. Size: {1}", name, items.Count);
            file.Write(DownloadItemsAsync(items, "", name).Result);
            file.Close();
            Console.WriteLine("Download ended: {0}. Downloaded {1} items.", name, items.Count);
        }

        static void DisplayHelp()
        {
            Console.WriteLine(@"
Web Scraper For zhats.com v1.0.0  released: September 29, 2016
Copyright (C) 2016 Konstantin S.
https://www.upwork.com/fl/havendv

Usage:
    webscraper.exe <pathtooutputdir>
    - pathtooutputdir - Directory for save output csv files. Example: C:\\zhats.com\\

");
        }
        private static bool HelpRequired(string param)
        {
            return param == "-h" || param == "--help" || param == "/?";
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length < 1 || HelpRequired(args[0]))
                {
                    DisplayHelp();
                    Console.ReadKey();
                    return;
                }
                //WebUtilities.ClearCache();
                var outputDir = args[0];
                Console.WriteLine("Download started.");
                LoadCollection("T-Shirts", outputDir);
                //WebUtilities.DownloadPages("http://www.sanmar.com/sanmar-servlets/SearchServlet?catId={0}&va=t", Enumerable.Range(0, 255));
                Console.WriteLine("Download ended.");
            }
            catch(Exception e)
            {
                Console.WriteLine("Exception: {0}", e.Message);
            }
            Console.ReadKey();
        }
    }
}
using System;
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
using System.Data;
using System.Net.Http;
using AngleSharp.Parser.Html;
using AngleSharp.Dom.Html;

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
        static IDictionary<string, int> _itemCount = new Dictionary<string, int>();
        static DataTable _table = null;
        static int _id = 5000;
        static object _mutex = new object();

        static string GetPriceForSize(string name, string size)
        {
            var prices = new List<double>();
            var sizePrices = new List<double>();
            for (var i = 0; i < _table.Rows.Count; ++i)
            {
                var row = _table.Rows[i];
                var itemName = (string)row.ItemArray[2];
                var itemSize = (string)row.ItemArray[6];
                var itemPrice = (double)row.ItemArray[8];
                if (string.Equals(itemName, name, StringComparison.OrdinalIgnoreCase))
                {
                    prices.Add(itemPrice);
                    if (string.Equals(itemSize, size, StringComparison.OrdinalIgnoreCase) ||
                        (itemSize == "S-XL" && (size == "S" || size == "M" || size == "L" || size == "XL")) ||
                        (itemSize == "XS-XL" && (size == "XS" || size == "S" || size == "M" || size == "L" || size == "XL")))
                    {
                        sizePrices.Add(itemPrice);
                    }
                }
            }
            if (sizePrices.Count > 0)
            {
                return sizePrices.Max().ToString("0.00", CultureInfo.InvariantCulture);
            }
            if (prices.Count > 0)
            {
                return prices.Max().ToString("0.00", CultureInfo.InvariantCulture);
            }
            return string.Empty;
        }

        static string GetBrand(string name)
        {
            for (var i = 0; i < _table.Rows.Count; ++i)
            {
                var row = _table.Rows[i];
                var itemName = (string)row.ItemArray[2];
                var itemBrand = (string)row.ItemArray[3];
                if (string.Equals(itemName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return itemBrand;
                }
            }
            return "Alphabroder";
        }

        static Uri CreateLink(string catPath, int resultsNumber = 5000)
        {
            return new Uri(
                $"https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults={resultsNumber}&ff=yes&RequestData=CA_Search&CatPath={catPath}"
                );
        }
        static Uri CreateLink(string catName, string catAttribute, int resultsNumber = 5000)
        {
            return CreateLink(
                $"All%2BProducts%2F%2F%2F%2F{catName}%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3D{catAttribute}%27",
                resultsNumber
                );
        }

        static string GetName(IHtmlDocument document)
        {
            return document.QuerySelectorAll("#style-header-name h1").FirstOrDefault()?.TextContent ?? string.Empty;
        }

        static string GetDescription(IHtmlDocument document)
        {
            return document.QuerySelectorAll(".bullet").First().OuterHtml;
        }

        static string GetPrice(string data)
        {
            var pattern = "<div id=\\\"style-price\\\">.+?\\$(.+?) USD<\\/div>";
            foreach (Match match in Regex.Matches(data.Replace("\n", ""), pattern))
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
                    images.Add("http:" + value);
                }
            }

            return images.Distinct().ToList();
        }

        static List<string> GetColorData(IHtmlDocument document)
        {
            var findClasses = document.QuerySelectorAll("div.colordisp");
            var colorDatas = new List<string>();
            foreach (var findedClass in findClasses)
            {
                colorDatas.Add(findedClass.OuterHtml.Replace('\n', ' '));
            }

            return colorDatas;
        }

        static string GetSizeData(string data)
        {
            foreach (Match m1 in Regex.Matches(data.Replace('\n', ' '), "<table class=\\\"specs-tab\\\".+?table>"))
            {
                foreach (Match m2 in Regex.Matches(m1.Value, "<tr class=\\\"tr-row1\\\".+?tr>"))
                {
                    return m2.Value;
                }
            }

            return string.Empty;
        }

        static List<string> GetSizes(string data)
        {
            var sizes = new List<string>();
            var sizeData = GetSizeData(data);
            foreach (Match m in Regex.Matches(sizeData, "<td class=\\\"td-cols\\\">(.+?)<\\/td>"))
            {
                var size = m.Groups[1].Value.Replace("-", "");
                if (!string.IsNullOrWhiteSpace(size))
                {
                    sizes.Add(size);
                }
            }

            return sizes;
        }

        static string GetStyleColor(string data)
        {
            var colors = new List<string>();
            foreach (Match m in Regex.Matches(data, "style=\\\"background-color:(#[0-9a-fA-F]+?);"))
            {
                colors.Add(m.Groups[1].Value);
                //return m.Groups[1].Value;
            }

            return string.Join("|", colors);
        }

        static string GetFrontImage(string style, string color = "00")
        {
            return $"http://marketing.peaksystems.com/imglib/mresjpg/176999/{style}_{color}_z.jpg";
        }

        static string GetBackImage(string style, string color = "00")
        {
            return $"http://marketing.peaksystems.com/imglib/mresjpg/176999/{style}_{color}_z_BK.jpg";
        }

        static string GetSideImage(string style, string color = "00")
        {
            return $"http://marketing.peaksystems.com/imglib/mresjpg/176999/{style}_{color}_z_SD.jpg";
        }

        static string GetColorFrontImage(string data)
        {
            foreach (Match m in Regex.Matches(data, "data-front=\\\"(.+?)\\\""))
            {
                return "https://www.alphabroder.com" + m.Groups[1].Value;
            }

            return string.Empty;
        }

        static string GetColorBackImage(string data)
        {
            foreach (Match m in Regex.Matches(data, "data-back=\\\"(.+?)\\\""))
            {
                return "https://www.alphabroder.com" + m.Groups[1].Value;
            }

            return string.Empty;
        }

        static string GetColorSideImage(string data)
        {
            foreach (Match m in Regex.Matches(data, "data-side=\\\"(.+?)\\\""))
            {
                return "https://www.alphabroder.com" + m.Groups[1].Value;
            }

            return string.Empty;
        }

        static string GetColorName(string data)
        {
            foreach (Match m in Regex.Matches(data, "data-color=\\\"(.+?)\\\""))
            {
                return m.Groups[1].Value;
            }

            return string.Empty;
        }

        static string GetColorId(string data)
        {
            foreach (Match m in Regex.Matches(data, "buyReg.+?,'(.+?)'\\);"))
            {
                return m.Groups[1].Value;
            }

            return string.Empty;
        }

        class ColorData
        {
            public string HtmlColor;
            public string FrontImage;
            public string BackImage;
            public string SideImage;
        }

        static Dictionary<string, ColorData> GetColors(string name, IHtmlDocument document)
        {
            var colors = new Dictionary<string, ColorData>();
            foreach (var colorData in GetColorData(document))
            {
                var colorId = GetColorId(colorData);
                colors[GetColorName(colorData)] =
                    new ColorData
                    {
                        HtmlColor = GetStyleColor(colorData),
                        FrontImage = GetFrontImage(name, colorId),
                        BackImage = GetBackImage(name, colorId),
                        SideImage = GetSideImage(name, colorId)
                    };
            }

            return colors;
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
        public static Task DownloadAsync(string requestUri)
        {
            var filename = "images/" + GetGoodName(requestUri);
            if (File.Exists(filename))
            {
                return Task.Run(()=> { });
            }
            if (requestUri == null)
                throw new ArgumentNullException("requestUri");

            return DownloadAsync(new Uri(requestUri), filename);
        }

        public static async Task DownloadAsync(Uri requestUri, string filename)
        {
            if (filename == null)
                throw new ArgumentNullException("filename");

            using (var httpClient = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, requestUri))
                {
                    using (
                        Stream contentStream = await (await httpClient.SendAsync(request)).Content.ReadAsStreamAsync(),
                        stream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        await contentStream.CopyToAsync(stream);
                    }
                }
            }
        }

        static string GetGoodName(string url)
        {
            var name = url;
            foreach (var symvol in Path.GetInvalidPathChars())
            {
                name = name.Replace(symvol.ToString(), "");
            }
            name = Path.GetFileName(name);
            return name.Replace("/", "").Replace(":", "").Replace("?", "");
        }

        static async Task<string> DownloadPage(string url, string category, string subcategory)
        {
            try
            {
                //Console.WriteLine("Current url: {0}", url);
                var data = await WebUtilities.DownloadPage(new Uri(url), usingCache: true);
                var parser = new HtmlParser();
                var document = parser.Parse(data);
                var name = GetName(document);
                if (string.IsNullOrWhiteSpace(name))
                {
                    data = await WebUtilities.DownloadPage(new Uri(url), usingCache: false);
                    document = parser.Parse(data);
                    name = GetName(document);
                }
                Directory.CreateDirectory("images");
                var colors = GetColors(name, document);
                foreach (var color in colors)
                {
                    var colorData = color.Value;
                    await DownloadAsync(colorData.FrontImage);
                    await DownloadAsync(colorData.BackImage);
                    await DownloadAsync(colorData.SideImage);
                }
                return "";
                var desc = GetDescription(document);
                var price = GetPrice(data);
                if (string.IsNullOrWhiteSpace(price))
                {
                    price = GetPriceForSize(name, "");
                }
                var brand = GetBrand(name);

                var images = new List<string> {
                    GetFrontImage(name),
                    GetBackImage(name),
                    GetSideImage(name)
                };
                var sizes = GetSizes(data);

                //Fix duplicates roman number method
                if (_itemCount.ContainsKey(name))
                {
                    _itemCount[name] = _itemCount[name] + 1;
                    name += " " + ToRoman(_itemCount[name]);
                }
                else
                {
                    _itemCount.Add(name, 1);
                }

                var strings = new List<string>();
                foreach (var color in colors)
                {
                    var colorData = color.Value;
                    var colorImages = new List<string> { colorData.FrontImage, colorData.BackImage, colorData.SideImage };
                    images = colorImages;
                    break;
                }
                lock (_mutex)
                {
                    strings.Add(BigCommerceUtilities.ProductFormat(name, desc, "ALP", (++_id).ToString(), brand, price, "", images, "Alphabroder", category, subcategory));
                }
                foreach (var color in colors)
                {
                    var colorData = color.Value;
                    var colorImages = new string[] { colorData.FrontImage, colorData.BackImage, colorData.SideImage };
                    strings.Add(BigCommerceUtilities.ColorFormat(color.Key, colorImages, colorData.HtmlColor));
                }
                foreach (var size in sizes)
                {
                    var sizePrice = GetPriceForSize(name, size);
                    //Console.WriteLine("Name: {0}, Size: {1}, Price: {2}", name, size, GetPriceForSize(name, size));
                    strings.Add(BigCommerceUtilities.SizeFormat(size, sizePrice, ""));
                }
                return string.Join("\n", strings);
            }
            catch (Exception e)
            {
                Console.WriteLine("Current url: {0}", url);
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            return string.Empty;
        }

        static async Task<IList<string>> GetItems(Uri url)
        {
            var items = new List<string>();
            try
            {
                var data = await WebUtilities.DownloadPage(url, usingCache: true);

                foreach (Match match in Regex.Matches(data, "<div class=\\\"rsltProdNameText\\\">(.+?)<\\/div>"))
                {
                    items.Add($"https://www.alphabroder.com/cgi-bin/online/webshr/prod-labeldtl.w?sr={match.Groups[1].Value}&currentColor=");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return items;
        }

        static async Task<IList<string>> GetItemsMultipage(string url)
        {
            var items = new List<string>();
            IList<string> nextItems;
            var page = 0;
            do
            {
                ++page;
                nextItems = await GetItems(new Uri(url + "&currentpage=" + page));
                items.AddRange(nextItems);
            }
            while (nextItems.Count > 0);

            return items;
        }

        static async Task<string> DownloadItemsAsync(IList<string> items, string category, string subcategory, string to)
        {
            Console.WriteLine($"Start download {category} Subcategory: {subcategory}. Size: {items.Count}");
            var strings = await Task.WhenAll(items.Select(item =>
                   Task.Run(() =>
                       DownloadPage(item, category, subcategory)
               )));
            strings = strings.Where(value => !string.IsNullOrWhiteSpace(value)).ToArray();
            Console.WriteLine($"Download ended: {subcategory}. Downloaded {strings.Length} items.");
            var file = BigCommerceUtilities.CreateImportCSVFile(Path.Combine(to, category + " " + subcategory + "0.csv"));
            for (var i = 0; i < strings.Length; ++i)
            {
                if (i % 50 == 0)
                {
                    file.Close();
                    file = BigCommerceUtilities.CreateImportCSVFile(Path.Combine(to, category + " " + subcategory + (i / 50) + ".csv"));
                }
                file.WriteLine(strings[i]);
            }
            file.Close();
            return string.Join("\n", strings);
        }

        static async Task LoadSubcategory(Uri url, string category, string subcategory, string to)
        {
            var items = await GetItems(url);
            //WebUtilities.DownloadPages(items);
            var result = await DownloadItemsAsync(items, category, subcategory, to);
        }

        static void DisplayHelp()
        {
            Console.WriteLine(@"
Web Scraper For alphabroder.com v1.0.0  released: October 10, 2016
Copyright (C) 2016 Konstantin S.
https://www.upwork.com/fl/havendv

Usage:
    webscraper.exe <pathtooutputdir>
    - pathtooutputdir - Directory for save output csv files. Example: C:\WebScrapingData\

");
        }
        private static bool HelpRequired(string param)
        {
            return param == "-h" || param == "--help" || param == "/?";
        }

        static async Task LoadTShirtsCategory(string outputDir)
        {
            var categoryName = "T-Shirts";
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DAge%20%3D%20%27Youth%27"), categoryName, "Youth", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DGender%20%3D%20%27Ladies%27%27"), categoryName, "Ladies", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSleeve%20Length%20%3D%20%27Long%27"), categoryName, "Long Sleeve", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FT-Shirts%2F%2F%2F%2FAttribSelect%3DNeckline%20%3D%27V-Neck%27"), categoryName, "V-Neck", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FT-Shirts%2F%2F%2F%2FAttribSelect%3DSleeve%20Style%20%3D%27Tank%27"), categoryName, "Tanks", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSleeve%20Style%20%3D%20%27Sleeveless%27"), categoryName, "Sleeveless", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DMoisture%20Wicking%20%3D%20%27Yes%27"), categoryName, "Moisture Wicking", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%273%20oz%2E%27"), categoryName, "3 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%274%20oz%2E%27"), categoryName, "4 oz.", outputDir);
            //await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%275%20oz%2E%27"), categoryName, "5 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2FAttribSelect%3DWeight%3D%275%20oz%2E%27%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DFabric%20%3D%20%27100%25%20Cotton%27"), categoryName, "5 oz. Cotton", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2FAttribSelect%3DWeight%3D%275%20oz%2E%27%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DFabric%20%3D%20%2750%2F50%20Cotton-Poly%27"), categoryName, "5 oz. Blend", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%276%20oz%2E%27"), categoryName, "6 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FT-Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%20%3D%20%27New%27"), categoryName, "New", outputDir);
        }

        static async Task LoadSweatshirtsCategory(string outputDir)
        {
            var categoryName = "Sweatshirts";
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FSweatshirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Crewneck%27"), categoryName, "Crewneck", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FSweatshirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Pullover%20Hood%27"), categoryName, "Pullover Hood", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FSweatshirts%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27quarter%20and%20half-zips%27%3B%3B%3B%3BType%20%3D%20%27%20Full-Zip%20Hood%27"), categoryName, "Zippered Hood", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FSweatshirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Sweatpants%27"), categoryName, "Sweatpants", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FSweatshirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DAge%20%3D%20%27Youth%27"), categoryName, "Youth", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FSweatshirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DGender%20%3D%20%27Ladies%27%27"), categoryName, "Ladies", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FSweatshirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27%3C%206%20oz%2E%27"), categoryName, "less 6 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FSweatshirts%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27%206%20oz%2E%27"), categoryName, "6 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FSweatshirts%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27%207-8%20oz%2E%27"), categoryName, "7-8 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FSweatshirts%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27%209%20oz%2E%27"), categoryName, "9 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FSweatshirts%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27%20%3E10%20oz%2E%27"), categoryName, "more 10 oz.", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FSweatshirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%20%3D%20%27New%27"), categoryName, "New", outputDir);
        }

        static async Task LoadPolosCategory(string outputDir)
        {
            var categoryName = "Polos";
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DPlacket%20%2F%20Neck%20%3D%20%27Mock%27"), categoryName, "Mock", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSleeve%20%3D%20%27Long%20Sleeve%27"), categoryName, "Long Sleeve", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DPocket%20%3D%20%27Yes%27"), categoryName, "Pocket Polos", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DMoisture%20Wicking%20%3D%20%27Yes%27"), categoryName, "Moisture Wicking", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DEasy%20Care%20%3D%20%27Yes%27"), categoryName, "Easy Care", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DColor%20Blocked%20%3D%20%27Yes%27"), categoryName, "Color Block", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DStriped%20%3D%20%27Yes%27"), categoryName, "Stripes", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DTextured%20%3D%20%27Yes%27"), categoryName, "Textures", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DAge%20%3D%20%27Youth%27"), categoryName, "Youth", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DGender%20%3D%20%27Ladies%27%27"), categoryName, "Ladies", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FPolos%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%20%3D%20%27New%27"), categoryName, "New", outputDir);
        }

        static async Task LoadKnitsandLayeringCategory(string outputDir)
        {
            var categoryName = "Knits and Layering";
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FKnits%20and%20Layering%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DPerformance%20%3D%20%27Yes%27"), categoryName, "Performance", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FKnits%20and%20Layering%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%3D%27New%27%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27lightweight%27"), categoryName, "Lightweight", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FKnits%20and%20Layering%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%3D%27New%27%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27midweight%27"), categoryName, "Midweight", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FKnits%20and%20Layering%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DFashion%20%3D%20%27Yes%27"), categoryName, "Dress", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FKnits%20and%20Layering%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27V-Neck%27"), categoryName, "V-Neck", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FKnits%20and%20Layering%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%20%3D%20%27New%27"), categoryName, "New", outputDir);
        }

        static async Task LoadFleeceCategory(string outputDir)
        {
            var categoryName = "Fleece";
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FSweatshirts%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27quarter%20and%20half-zips%27%3B%3B%3B%3BType%20%3D%20%27%20Full-Zip%20Hood%27"), categoryName, "Zippered Hood", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FSweatshirts%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27quarter%20and%20half-zips%27%3B%3B%3B%3BType%20%3D%20%27%20Full-Zip%20Hood%27"), categoryName, "Zippered Hood", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FFleece%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Pullover%20Hood%27"), categoryName, "Pullover Hood", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FFleece%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DAge%20%3D%20%27Youth%27"), categoryName, "Youth", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FFleece%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DPerformance%20%3D%20%27Yes%27"), categoryName, "Performance", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FFleece%2F%2F%2F%2FAttribSelect%3DWeight%20%3D%20%27%20lightweight%27"), categoryName, "Lightweight", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FFleece"), categoryName, "Heavyweight", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FFleece%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Vest%27"), categoryName, "Vest", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FFleece%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%20%3D%20%27New%27"), categoryName, "New", outputDir);
        }

        static async Task LoadWovenShirtsCategory(string outputDir)
        {
            var categoryName = "Woven Shirts";
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Broadcloth%27"), categoryName, "Broadcloth", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Chambray%27"), categoryName, "Chambray", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Camp%27"), categoryName, "Camp", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Denim%27"), categoryName, "Denim", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Dobby%27"), categoryName, "Dobby", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Fishing%27"), categoryName, "Fishing", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Oxford%27"), categoryName, "Oxford", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Poplin%27"), categoryName, "Poplin", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Twill%27"), categoryName, "Twill", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DWrinkle%20Resistant%20%3D%20%27Yes%27"), categoryName, "Wrinkle Resistant", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DStain%20Resistant%20%3D%20%27Yes%27"), categoryName, "Stain Resistant", outputDir);
            await LoadSubcategory(CreateLink("All%2BProducts%2F%2F%2F%2FWoven%20Shirts%2F%2F%2F%2F%2F%2F%2F%2FAttribSelect%3DSpecial%20Collections%20%3D%20%27New%27"), categoryName, "New", outputDir);
        }

        static async Task LoadOuterwearCategory(string outputDir)
        {
            var categoryName = "Outerwear";
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Hi-Visibility"), categoryName, "Hi-Visibility", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Fashion%20%3D%20%27Workwear"), categoryName, "Workwear", outputDir);
            await LoadSubcategory(new Uri("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=500&RequestAction=advisor&RequestData=CA_BreadcrumbSelect&currentpage=1&CatPath=All%20Products%2F%2F%2F%2FUserSearch%3Dnot%28mill_code%3D%27UA%27%29%2F%2F%2F%2FUserSearch1%3DSoftshell"), categoryName, "Softshell", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Fashion%20%3D%20%27Rainwear"), categoryName, "Rainwear", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Windshirt"), categoryName, "Windshirt", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Fashion%20%3D%20%27Athletic"), categoryName, "Athletic", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Fashion%20%3D%20%27Golf"), categoryName, "Golf", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Weight%20%3D%20%27Lightweight"), categoryName, "Lightweight", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Weight%20%3D%20%27Midweight"), categoryName, "Midweight", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Weight%20%3D%20%27Heavyweight"), categoryName, "Heavyweight", outputDir);
            await LoadSubcategory(CreateLink("All%20Products%2F%2F%2F%2FUserSearch%3DOuterwear%2F%2F%2F%2FAttribSelect%3DType%20%3D%20%27Systems%27%2F%2F%2F%2FALP-Categories"), categoryName, "Systems", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Fabric%20%3D%20%27Poly%20Fleece"), categoryName, "Poly Fleece", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Special%20Collections%20%3D%20%27New"), categoryName, "New", outputDir);
        }

        static async Task LoadPantsCategory(string outputDir)
        {
            var categoryName = "Pants";
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Workwear"), categoryName, "Workwear", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Yoga-Fitness"), categoryName, "Yoga-Fitness", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Athletic"), categoryName, "Athletic", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Lounge-Sleepwear"), categoryName, "Lounge-Sleepwear", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Leggings"), categoryName, "Leggings", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Capri-Crop"), categoryName, "Capri-Crop", outputDir);
            //
            await LoadSubcategory(CreateLink(categoryName, "Special%20Collections%20%3D%20%27New"), categoryName, "New", outputDir);
        }

        static async Task LoadShortsCategory(string outputDir)
        {
            var categoryName = "Shorts";
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Athletic"), categoryName, "Athletic", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Workwear"), categoryName, "Workwear", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Lingerie-Sleepwear"), categoryName, "Lingerie-Sleepwear", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Inseam%20%3D%20%275-6%22"), categoryName, "5-6 Inseam", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Inseam%20%3D%20%277-%209%22"), categoryName, "7-9 Inseam", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Inseam%20%3D%20%2710-13%22"), categoryName, "10-13 Inseam", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Special%20Collections%20%3D%20%27New"), categoryName, "New", outputDir);
        }

        static async Task LoadInfantsAndToddlersCategory(string outputDir)
        {
            var categoryName = "Infants%20%7C%20Toddlers";
            var categoryName2 = "Infants and Toddlers";
            await LoadSubcategory(CreateLink(categoryName, "Sleeve%20%3D%20%27Short"), categoryName2, "Short Sleeve", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Sleeve%20%3D%20%27Long"), categoryName2, "Long Sleeve", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Sweatpants"), categoryName2, "Sweatpants", outputDir);
            //
            //
            //
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Creeper"), categoryName2, "Creeper", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Romper"), categoryName2, "Romper", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Bibs"), categoryName2, "Bibs", outputDir);
            //
            //await LoadSubcategory(CreateLink(categoryName, "Special%20Collections%20%3D%20%27New"), categoryName, "New", outputDir);
        }

        static async Task LoadHeadwearCategory(string outputDir)
        {
            var categoryName = "Headwear";
            await LoadSubcategory(CreateLink(categoryName, "Mill%20%3D%20%27Flexfit"), categoryName, "Flexfit", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%275%20Panel"), categoryName, "5 Panel", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%276%20Panel"), categoryName, "6 Panel", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Construction%20%3D%20%27Structured"), categoryName, "Structured", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Construction%20%3D%20%27Unstructured"), categoryName, "Unstructured", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Beanies"), categoryName, "Beanies", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Visors"), categoryName, "Visors", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Special%20Dyes%20or%20Prints%20%3D%20%27Pigment%20Dyed"), categoryName, "Pigment Dyed", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Special%20Dyes%20or%20Prints%20%3D%20%27Camo"), categoryName, "Camo", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Military"), categoryName, "Military", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Bucket"), categoryName, "Bucket", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Use%20%3D%20%27Running"), categoryName, "Runners", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Special%20Collections%20%3D%20%27New"), categoryName, "New", outputDir);
        }

        static async Task LoadBagsAndAccessoriesCategory(string outputDir)
        {
            var categoryName = "Bags%20and%20Accessories";
            var categoryName2 = "Bags and Accessories";
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27backpacks"), categoryName2, "Backpacks", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27drawstring"), categoryName2, "Drawstring", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27duffel"), categoryName2, "Duffel", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Laptop%20%26%20Tablet%20Cases"), categoryName2, "Laptop and Tablet Cases", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Messenger%20%26%20Briefcases"), categoryName2, "Messenger and Briefcases", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Tote"), categoryName2, "Tote", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Blanket"), categoryName2, "Blanket", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%27Towel%27%3B%3B%3B%3BType%20%3D%20%27Golf%20Towel"), categoryName2, "Towel", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%20%27Aprons"), categoryName2, "Aprons", outputDir);
            await LoadSubcategory(CreateLink(categoryName, "Type%20%3D%27Scarves%27%3B%3B%3B%3BType%20%3D%20%27Socks"), categoryName2, "Scarves and Socks", outputDir);
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
                _table = Excel.LoadExcelFile("data.xls");
                Console.WriteLine("Download started.");
                LoadTShirtsCategory(outputDir).Wait();
                LoadSweatshirtsCategory(outputDir).Wait();
                LoadPolosCategory(outputDir).Wait();
                LoadKnitsandLayeringCategory(outputDir).Wait();
                LoadFleeceCategory(outputDir).Wait();
                LoadWovenShirtsCategory(outputDir).Wait();
                LoadOuterwearCategory(outputDir).Wait();
                LoadPantsCategory(outputDir).Wait();
                LoadShortsCategory(outputDir).Wait();
                LoadInfantsAndToddlersCategory(outputDir).Wait();
                LoadHeadwearCategory(outputDir).Wait();
                LoadBagsAndAccessoriesCategory(outputDir).Wait();

                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=050&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Sweatshirts", "Sweatshirts", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=100&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Polos", "Polos", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=010&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Knits%20and%20Layering", "Knits%20and%20Layering", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=030&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Fleece", "Fleece", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=120&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Woven%20Shirts", "Woven%20Shirts", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=070&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Outerwear", "Outerwear", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=080&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Pants", "Pants", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=090&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Shorts", "Shorts", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=065&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Infants%20|%20Toddlers", "Infants%20|%20Toddlers", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=040&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Headwear", "Headwear", "Youth", outputDir);
                //LoadSubcategory("https://www.alphabroder.com/cgi-bin/online/webshr/search-result.w?nResults=5000&ff=yes&cat=020&RequestAction=advisor&RequestData=CA_CategoryExpand&currentpage=1&bpath=c&CatPath=All%20Products////ALP-Categories////Bags%20and%20 ", "Bags%20and%20Accessories", "Youth", outputDir);
                //WebUtilities.DownloadPages("http://www.sanmar.com/sanmar-servlets/SearchServlet?catId={0}&va=t", Enumerable.Range(0, 255));
                Console.WriteLine("Download ended.");
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception: {0}", e.Message);
            }
            Console.ReadKey();
        }
    }
}

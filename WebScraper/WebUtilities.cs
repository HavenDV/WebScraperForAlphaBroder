using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using HtmlAgilityPack;
using ScrapySharp.Extensions;
using ScrapySharp.Network;
using AngleSharp;

namespace WebScraper
{
    static class WebUtilities
    {
        public static string GetCacheName(string url)
        {
            var cacheName = new Uri(url).Query;
            foreach (var symvol in Path.GetInvalidPathChars())
            {
                cacheName = cacheName.Replace(symvol + "", "");
            }
            return cacheName.Replace("/", "").Replace(":", "").Replace("?", "");
        }

        public static string CacheDirectory {
            get 
            {
                return "cache";
            }
        }

        public static string GetCachePath(string url)
        {
            Directory.CreateDirectory(CacheDirectory);
            return Path.Combine(CacheDirectory, GetCacheName(url)); ;
        }

        public static void ClearCache()
        {
            Directory.Delete(CacheDirectory, true);
        }

        public static string DownloadPage(string url, bool usingAjax = true, bool usingCache = true)
        {
            var filePath = GetCachePath(url);
            if (usingCache && File.Exists(filePath))
            {
                return File.OpenText(filePath).ReadToEnd();
            }

            string data = string.Empty;
            if (usingAjax)
            {
                //var config = Configuration.Default.WithDefaultLoader();
                //data = BrowsingContext.New(config).OpenAsync(url).Result.;
                data = new ScrapingBrowser().AjaxDownloadString(new Uri(url));
            }
            else
            {
                data = new WebClient().DownloadString(url);
            }
            File.CreateText(GetCachePath(url)).Write(data);
            return data;
        }

        public static IEnumerable<string> DownloadPages(IEnumerable<string> urls, bool usingAjax = true, bool usingCache = true)
        {
            return Task.WhenAll(urls.Select(url => Task.Run(() => DownloadPage(url, usingAjax, usingCache)))).Result;
        }

        public static IEnumerable<string> DownloadPages(string format, IEnumerable<int> range, bool usingAjax = true, bool usingCache = true)
        {
           return DownloadPages(range.Select(i => string.Format(format, i)), usingAjax, usingCache);
        }
    }
}

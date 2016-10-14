using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.IO;
using HtmlAgilityPack;
using ScrapySharp.Extensions;
using ScrapySharp.Network;
using AngleSharp;

namespace WebScraper
{
    public static class StringExtension
    {
        public static string GetLast(this string source, int tail_length)
        {
            if (tail_length >= source.Length)
                return source;
            return source.Substring(source.Length - tail_length);
        }
    }

    static class WebUtilities
    {
        public static string GetCacheName(string url, int maxLenght = 255)
        {
            var cacheName = new Uri(url).Query;
            foreach (var symvol in Path.GetInvalidPathChars())
            {
                cacheName = cacheName.Replace(symvol + "", "");
            }
            return cacheName.Replace("/", "").Replace(":", "").Replace("?", "").GetLast(maxLenght);
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
            return Path.Combine(CacheDirectory, GetCacheName(url, 150)); ;
        }

        public static void ClearCache()
        {
            Directory.Delete(CacheDirectory, true);
        }

        public static async Task<string> DownloadPage(Uri url, bool usingAjax = true, bool usingCache = true)
        {
            var filePath = GetCachePath(url.ToString());
            if (usingCache && File.Exists(filePath))
            {
                return File.OpenText(filePath).ReadToEnd();
            }

            string data = string.Empty;
            //if (usingAjax)
            {
                //var config = Configuration.Default.WithDefaultLoader();
                //data = BrowsingContext.New(config).OpenAsync(url).Result.;
                //data = new ScrapingBrowser().AjaxDownloadString(new Uri(url));
                try
                {
                    data = await new HttpClient().GetStringAsync(url);
                }
                catch
                {
                    return await DownloadPage(url, usingAjax, usingCache);
                }
            }
            //else
            {
                //data = new WebClient().DownloadString(url);
            }
            File.CreateText(filePath).Write(data);
            return data;
        }

        public static IEnumerable<string> DownloadPages(IEnumerable<string> urls, bool usingAjax = true, bool usingCache = true)
        {
            return Task.WhenAll(urls.Select(url => Task.Run(() => DownloadPage(new Uri(url), usingAjax, usingCache)))).Result;
        }

        public static IEnumerable<string> DownloadPages(string format, IEnumerable<int> range, bool usingAjax = true, bool usingCache = true)
        {
           return DownloadPages(range.Select(i => string.Format(format, i)), usingAjax, usingCache);
        }
    }
}

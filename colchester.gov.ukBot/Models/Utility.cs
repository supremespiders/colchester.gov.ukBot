using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Newtonsoft.Json;

namespace colchester.gov.ukBot.Models
{
    public static class Utility
    {
        public static string ConnectionString = "Data Source=system.db;Version=3;";
        public static string SimpleDateFormat = "dd/MM/yyyy HH:mm:ss";

        public static void SaveCookies(CookieContainer cookieContainer, string url)
        {
            try
            {
                var cookies = new List<Cookie>();
                foreach (Cookie cookie in cookieContainer.GetCookies(new Uri(url)))
                    cookies.Add(new Cookie { Name = cookie.Name, Value = cookie.Value });
                File.WriteAllText("ses", JsonConvert.SerializeObject(cookies));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        public static string BetweenStrings(string text, string start, string end)
        {
            var p1 = text.IndexOf(start, StringComparison.Ordinal) + start.Length;
            var p2 = text.IndexOf(end, p1, StringComparison.Ordinal);
            if (end == "") return (text.Substring(p1));
            else return text.Substring(p1, p2 - p1);
        }

        public static CookieContainer LoadCookies(string url)
        {
            var cookieContainer = new CookieContainer();
            try
            {
                var myCookies = JsonConvert.DeserializeObject<List<Cookie>>(File.ReadAllText("ses"));
                foreach (var myCookie in myCookies)
                    cookieContainer.Add(new Uri(url), new Cookie(myCookie.Name, myCookie.Value));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            return cookieContainer;
        }

    }
}

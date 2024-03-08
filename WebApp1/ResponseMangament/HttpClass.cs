using System;
using System.IO;
using System.Net;
using System.Text;

namespace WebApp1.ResponseMangament
{
    public class HttpClass
    {
        public static string GetGetResponse(string url, CookieContainer cookies, string encoding = "utf-8")
        {
            var uri = new Uri(url);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "GET";
            request.AllowAutoRedirect = true;
            request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            //request.ContentType = "application/x-www-form-urlencoded";
            //request.KeepAlive = ApplicationVariables.ConnectionKeepAlive;
            request.KeepAlive = true;
            request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134";
            //request.Host = uri.Host;
            //request.CookieContainer = cookies;
            //request.Headers.Add(HttpRequestHeader.Upgrade, "1");
            //request.Headers.Add(HttpRequestHeader.CacheControl, "max-age=0");
            request.AutomaticDecompression = DecompressionMethods.GZip;

            string strResponse = "";
            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    StreamReader loResponseStream = new
                      StreamReader(response.GetResponseStream(), Encoding.GetEncoding(encoding));

                    strResponse = loResponseStream.ReadToEnd();
                    loResponseStream.Close();
                }
            }
            catch (WebException e)
            {
                throw e;
            }
            return strResponse;
        }
    }
}

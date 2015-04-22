using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Text;

namespace Maximis.Toolkit.Http
{
    public static class HttpHelper
    {
        public static byte[] DownloadData(string url)
        {
            using (WebClient client = new WebClient { Encoding = Encoding.UTF8 })
            {
                return client.DownloadData(url);
            }
        }

        public static string DownloadString(string url)
        {
            using (WebClient client = new WebClient { Encoding = Encoding.UTF8 })
            {
                return client.DownloadString(url);
            }
        }

        public static byte[] PostData(string url, NameValueCollection data)
        {
            using (WebClient client = new WebClient())
            {
                return client.UploadValues(url, "POST", data);
            }
        }

        public static byte[] PostData(string url, Dictionary<string, string> data)
        {
            NameValueCollection nvc = new NameValueCollection();
            foreach (string key in data.Keys)
            {
                nvc.Add(key, data[key]);
            }
            return PostData(url, nvc);
        }

        public static string PostString(string url, string data)
        {
            using (WebClient client = new WebClient { Encoding = Encoding.UTF8 })
            {
                return client.UploadString(url, "POST", data);
            }
        }
    }
}
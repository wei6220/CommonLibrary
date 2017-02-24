using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary
{
    public class HttpHelper
    {
        public enum ContnetTypeEnum
        {
            FormData,
            Json,
        }

        public string Get(string url)
        {
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";

                using (var response = (HttpWebResponse)req.GetResponse())
                {
                    using (var reader = new StreamReader(response.GetResponseStream()))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write(ex.Message);
                return null;
            }
        }

        public string Post(string url, string parameter, ContnetTypeEnum type)
        {
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "POST";

                if(type == ContnetTypeEnum.Json)
                    req.ContentType = "application/json";
                else
                    req.ContentType = "application/x-www-form-urlencoded";

                byte[] data = Encoding.UTF8.GetBytes(parameter);
                req.ContentLength = data.Length;
                if (!string.IsNullOrWhiteSpace(parameter))
                {
                    using (var body = req.GetRequestStream())
                    {
                        body.Write(Encoding.UTF8.GetBytes(parameter), 0, data.Length);
                    }
                }

                using (var response = (HttpWebResponse)req.GetResponse())
                {
                    using (var stream = new StreamReader(response.GetResponseStream()))
                    {
                        return stream.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write(ex.Message);
                return null;
            }
        }

    }
}

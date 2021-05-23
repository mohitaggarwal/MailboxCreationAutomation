using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Authentication;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DPClientSDK
{
    public class HttpClientWrapper
    {
        public Uri BaseUri { get; set; }
        public string AccessToken { get; set; }

		public string Get(string url)
        {
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.BaseAddress = BaseUri;
                if(!string.IsNullOrEmpty(AccessToken))
                    httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + AccessToken);
                httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
                var httpResponse = httpClient.GetAsync(url).Result;
                Logger.FileLogger.Info($"Request Url: {httpResponse.RequestMessage.RequestUri.AbsoluteUri}");
                httpResponse.EnsureSuccessStatusCode();
                return httpResponse.Content.ReadAsStringAsync().Result;
            }
        }

        public string Post(string url, string body, string mediaType = "application/json")
		{
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.BaseAddress = BaseUri;
                if (!string.IsNullOrEmpty(AccessToken))
                    httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + AccessToken);
                var data = new StringContent(body, Encoding.UTF8, mediaType);
                var httpResponse = httpClient.PostAsync(url, data).Result;
                Logger.FileLogger.Info($"Request Url: {httpResponse.RequestMessage.RequestUri.AbsoluteUri}");
                httpResponse.EnsureSuccessStatusCode();
                return httpResponse.Content.ReadAsStringAsync().Result;
            }
        }
    }
}

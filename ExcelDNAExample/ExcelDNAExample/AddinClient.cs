using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http.Headers;

/*
 * Class for reusing HttpClient.
 * ApplicationEvents.AutoClose() calls the HttpClient's Dispose() currently
 */
namespace ExcelDNAExample
{
    static internal class AddinClient
    {
        private static readonly HttpClient httpClient;
        static AddinClient()
        {
            // Create reusable static HttpClient that 
            httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Accept.Clear();
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));        // give us json back

            // https://makolyte.com/csharp-configuring-how-long-an-httpclient-connection-will-stay-open/ was a useful reference for static HttpClients
            ServicePointManager.MaxServicePointIdleTime = (int)TimeSpan.FromMinutes(1).TotalMilliseconds;
            ServicePointManager.DnsRefreshTimeout = (int)TimeSpan.FromMinutes(1).TotalMilliseconds;     // not equivalent to ConnectionLeaseTimeout, but still could be nice

#if false
            // For setting at service point level instead
            ServicePoint servicePoint = ServicePointManager.FindServicePoint(new Uri("https://localhost:9000"));
            servicePoint.MaxIdleTime = (int)TimeSpan.FromMinutes(1).TotalMilliseconds;
            servicePoint.ConnectionLeaseTimeout = (int)TimeSpan.FromMinutes(1).TotalMilliseconds;
#endif
        }

        public static HttpClient GetHttpClient()
        {
            return httpClient;
        }
    }
}

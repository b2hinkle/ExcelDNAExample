using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelDNAExample
{
    internal static class Functions
    {
        [ExcelFunction(Description = "Custom cell functions", Category = "Custom cell Functions")]
        public static string SayHello([ExcelArgument(Description = "The name to say hi to")] string name)
        {
            return "Hello " + name;
        }




        [ExcelFunction(Description = "Custom cell functions", Category = "Custom cell Functions")]
        public static async Task<string> AsyncExample(string uri)
        {
            string retString = "";
            try
            {
                using (HttpResponseMessage response = await AddinClient.GetHttpClient().GetAsync(uri))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        retString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        retString = response.ReasonPhrase;
                    }
                }
            }
            catch (Exception e)
            {
                retString = e.Message;
            }
            return retString;
        }
    }
}

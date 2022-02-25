using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;         // Using this for PostAsJsonAsync since we want to use that but aren't using Microsoft.AspNet.WebApi.Client
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace TestExcelAddin
{
    public static class MyFunctions
    {
        private static readonly HttpClient client = new HttpClient();

        [ExcelFunction(Description = "My first .NET functions", Category = "Category A")]
        public static string SayHello ( [ExcelArgument(Description = "The name to say hi to")] string name)
        {
            return "Hello " + name;
        }




        [ExcelFunction(Description = "My first .NET functions", Category = "ASYNC")]
        public static async Task<string> AsyncExample(string uri)
        {
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));        // give us json back
            try
            {
                using (HttpResponseMessage response = await client.GetAsync(uri))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        string retString = await response.Content.ReadAsStringAsync();
                        return retString;
                    }
                }
            }
            catch (Exception e)
            {
                string message = e.Message;
            }
            return "darn";
        }
    }
}

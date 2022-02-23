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
        public static object AsyncExample()
        {
            Task<string> s = (Task<string>)ExcelDna.Integration.ExcelAsyncUtil.Run("AsyncExample", null, delegate { return Test(); });
            
            return s.Result;
        }

        static string Test()
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));        // give us json back

                try
                {
                    using (HttpResponseMessage response = client.GetAsync(@"https://catfact.ninja/fact").Result)
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string retString = response.Content.ReadAsStringAsync().Result;
                            return retString;
                        }
                    }
                }
                catch (Exception e)
                {
                    string message = e.Message;
                }


            }
            return "darn";








            /*using (HttpClient client = new HttpClient())
            {
                //client.BaseAddress = new Uri("https://us-zipcode.api.smartystreets.com/");
                //client.DefaultRequestHeaders.Accept.Clear();                                                            // clear headers just in case
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));        // give us json back

                //string function = "lookup?auth-id=ac7da12c-520e-2dd4-4365-d5f6346b9a23&auth-token=uIKoOq3LwLDY9E7pilsE";

                try
                {
                    //HttpResponseMessage response = await client.PostAsJsonAsync(@"https://us-zipcode.api.smartystreets.com/lookup?auth-id=ac7da12c-520e-2dd4-4365-d5f6346b9a23&auth-token=uIKoOq3LwLDY9E7pilsE", new { city = "Raleigh", state = "NC" });
                    HttpResponseMessage response = await client.GetAsync(@"https://catfact.ninja/fact");
                }
                catch (Exception e)
                {
                    string mes = e.Message;
                }
                return 3;
            }*/
        }
    }
}

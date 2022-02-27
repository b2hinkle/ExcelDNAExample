using System;
using System.IO;
using System.Resources;
using System.Net.Http.Json;         // Using this for PostAsJsonAsync since we want to use that but aren't using Microsoft.AspNet.WebApi.Client
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Text;
using System.Text.Json;

namespace TestExcelAddin
{
    [ComVisible(true)]
    public class CustomRibbonController : ExcelRibbon
    {
        private Application _excel;
        private IRibbonUI _thisRibbon;

        public CustomRibbonController()
        {
            _excel = (Application)ExcelDna.Integration.ExcelDnaUtil.Application;
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));
            }

            _thisRibbon = ribbon;

            _excel.WorkbookActivate += OnInvalidateRibbon;
            _excel.WorkbookDeactivate += OnInvalidateRibbon;
            _excel.SheetActivate += OnInvalidateRibbon;
            _excel.SheetDeactivate += OnInvalidateRibbon;

            if (_excel.ActiveWorkbook == null)      // Just useful for quick testing. Should be commented out in prod use
            {
                _excel.Workbooks.Add();
            }
        }

        private void OnInvalidateRibbon(object obj)
        {
            _thisRibbon.Invalidate();
        }















        /* Async ribbon press events can have the same signature as the normal excel async function, just without the static. Also you can return specific kind of task, but won't be a case where you do that since it's just a button being pressed. */
        public async Task OnTestRibbonButtonPressed(IRibbonControl control)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));        // give us json back

            string userName = "ac7da12c-520e-2dd4-4365-d5f6346b9a23";
            string password = "uIKoOq3LwLDY9E7pilsE";
            string city = "Raleigh";
            string state = "Nc";
            string url = $"https://us-zipcode.api.smartystreets.com/lookup?auth-id={userName}&auth-token={password}&city={city}&state={state}";     // No body is used for this post req. Query params instead

#if false
            // An endpoint may require the username and password to be in the header (instead of the url). In that case put it in the Authorization header
            byte[] authToken = Encoding.ASCII.GetBytes($"{userName}:{password}");
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(authToken));
#endif

            try
            {
                HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Post, url);
                using (HttpResponseMessage response = await client.SendAsync(req))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        string retString = await response.Content.ReadAsStringAsync();
                    }
                }
            }
            catch (Exception e)
            {
                string message = e.Message;
            }
        }
    }
}

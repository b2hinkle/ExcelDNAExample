using System;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

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
            
            int i = 4;
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
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));        // give us json back
            try
            {
                using (HttpResponseMessage response = await client.GetAsync(@"https://api.publicapis.org/entries"))
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

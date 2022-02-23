using System;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Net.Http.Headers;

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
















        public void OnSayHelloPressed(IRibbonControl control)
        {
            using (HttpClient client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://us-zipcode.api.smartystreets.com/");
                client.DefaultRequestHeaders.Accept.Clear();                                                            // clear headers just in case
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));      // just give us json back

                string function = "lookup?auth-id=ac7da12c-520e-2dd4-4365-d5f6346b9a23&auth-token=uIKoOq3LwLDY9E7pilsE";

                //HttpResponseMessage response = await client.PostAsync(function, new { city = "Raleigh", state = "NC" });
            }
            return;
        }
    }
}

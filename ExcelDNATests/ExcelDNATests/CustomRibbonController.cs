using System;
using System.IO;
using System.Resources;
using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDNATests
{
    [ComVisible(true)]
    public class CustomRibbonController : ExcelRibbon
    {
        private Application excelApp;
        private IRibbonUI thisRibbon;
        
        string userId    = "";
        string authToken = "";
        string zipcode   = "";

        public CustomRibbonController()
        {
            excelApp = (Application)ExcelDna.Integration.ExcelDnaUtil.Application;
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));
            }
            
            thisRibbon = ribbon;

            excelApp.WorkbookActivate += OnInvalidateRibbon;
            excelApp.WorkbookDeactivate += OnInvalidateRibbon;
            excelApp.SheetActivate += OnInvalidateRibbon;
            excelApp.SheetDeactivate += OnInvalidateRibbon;
        }

        private void OnInvalidateRibbon(object obj)
        {
            thisRibbon.Invalidate();
        }


        public void OnUserIdEditBoxChange(IRibbonControl control, string newText)
        {
            userId = newText;
        }
        public void OnAuthTokenEditBoxChange(IRibbonControl control, string newText)
        {
            authToken = newText;
        }
        public void OnZipcodeEditBoxChange(IRibbonControl control, string newText)
        {
            zipcode = newText;
        }












        /* 
         *  Async ribbon press events can have the same signature as the normal excel async function, just without the static. Also you can return specific kind of task, but won't be a case where you do that since it's just a button being pressed. 
         *  Only caveat with async functions is that they must transition to the main thread in order to interact with Excel. Just use ExcelAsyncUtil.QueueAsMacro(() => { }) for that 
         */

        public async Task OnAPIAuthPostCallPressed(IRibbonControl control)
        {
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(@"application/json"));        // give us json back

            string req_userName = userId;       // ac7da12c-520e-2dd4-4365-d5f6346b9a23
            string req_password = authToken;    // uIKoOq3LwLDY9E7pilsE
            string req_zipcode = zipcode;
            string url = $"https://us-zipcode.api.smartystreets.com/lookup?auth-id={req_userName}&auth-token={req_password}&zipcode={req_zipcode}";     // No body is used for this post req. Query params instead

#if false
            // An endpoint may require the username and password to be in the header (instead of the url). In that case put it in the Authorization header
            byte[] authToken = Encoding.ASCII.GetBytes($"{userName}:{password}");
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(authToken));
#endif
            string responseString = "---";
            try
            {
                HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Post, url);
                using (HttpResponseMessage response = await client.SendAsync(req))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        responseString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        responseString = response.ReasonPhrase;
                    }
                }
            }
            catch (Exception e)
            {
                responseString = e.Message;

            }

            ExcelAsyncUtil.QueueAsMacro(() => { excelApp.ActiveCell.Value2 = responseString; });    // Async functions must use   ExcelAsyncUtil.QueueAsMacro(() => { })   when interacting with Excel


        }

        public void OnWriteToSelectedCellPressed(IRibbonControl control)
        {
            Range rangeToWriteTo = excelApp.ActiveCell;

            rangeToWriteTo.Value2 = "written";
        }

        // Accessing specific cell
        public void OnWriteToSpecificCellPressed(IRibbonControl control)
        {
            
            Worksheet activeSheet = (Worksheet)excelApp.ActiveSheet;
            Range rangeToWriteTo = activeSheet.Range["A1"];
#if false
            Range rangeToWriteTo = activeSheet.Cells[1, 1];         // Alternative way
#endif

            rangeToWriteTo.Value2 = "written";




        }
        // Writing to specific cells
        public void OnWriteToSpecificCellsPressed(IRibbonControl control)
        {
            Worksheet activeSheet = (Worksheet)excelApp.ActiveSheet;

            object startSelection = activeSheet.Range["B2"];
            object endSelection = activeSheet.Range["AX20"];        // Excel uses patern (A,B,C .... AA,AB,AC, .... BA,BB,BC). In this case AX is the 50th collumn
#if false
            // Alternative way.....
            object startSelection = activeSheet.Cells[2, 2];
            object endSelection = activeSheet.Cells[20, 50];
#endif

            Range rangeToWriteTo = activeSheet.Range[startSelection, endSelection];


            rangeToWriteTo.Value2 = "written";
        }
    }
}

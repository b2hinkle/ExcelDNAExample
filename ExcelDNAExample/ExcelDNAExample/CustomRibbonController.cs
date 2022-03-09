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
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net.Http.Json;
using System.Text;

namespace ExcelDNAExample
{
    [ComVisible(true)]
    public sealed class CustomRibbonController : ExcelRibbon
    {
        private Application excelApp;   // Gives us access to excel app
        private IRibbonUI thisRibbon;   // Gives us access to the ribbon
        
        // Values that represent what is entered in the textboxes
        private string textboxValue_userId    = "";
        private string textboxValue_authToken = "";
        private string textboxValue_zipcode   = "";

        public CustomRibbonController()
        {
            excelApp = (Application)ExcelDna.Integration.ExcelDnaUtil.Application;
        }

        // Runs when excel loads this ribbon
        public void OnLoad(IRibbonUI ribbon)
        {
            if (ribbon == null)
            {
                throw new ArgumentNullException(nameof(ribbon));    // We want to throw an error if passed in ribbon is null
            }
            
            thisRibbon = ribbon;    // We might want to access the ribbon in another function

            // Refresh this ribbon during these moments
            excelApp.WorkbookActivate += OnInvalidateRibbon;
            excelApp.WorkbookDeactivate += OnInvalidateRibbon;
            excelApp.SheetActivate += OnInvalidateRibbon;
            excelApp.SheetDeactivate += OnInvalidateRibbon;
        }

        // Refreshes the ribbon
        private void OnInvalidateRibbon(object obj)
        {
            thisRibbon.Invalidate();
        }




        // These events get fired when user changes the textbox. Update class variables when this happens
        public void OnUserIdEditBoxChange(IRibbonControl control, string newText)
        {
            textboxValue_userId = newText;
        }
        public void OnAuthTokenEditBoxChange(IRibbonControl control, string newText)
        {
            textboxValue_authToken = newText;
        }
        public void OnZipcodeEditBoxChange(IRibbonControl control, string newText)
        {
            textboxValue_zipcode = newText;
        }




        // Examples of writing to specific cells (ranges)
        public void OnWriteToSelectedCellPressed(IRibbonControl control)
        {
            Range rangeToWriteTo = excelApp.ActiveCell;

            rangeToWriteTo.Value2 = "written";
        }
        public void OnWriteToSpecificCellPressed(IRibbonControl control)
        {

            Worksheet activeSheet = (Worksheet)excelApp.ActiveSheet;
            Range rangeToWriteTo = activeSheet.Range["A1"];
#if false
            // Alternative way.....
            Range rangeToWriteTo = activeSheet.Cells[1, 1];         // Alternative way. [1,1] is the very top left of the sheet
#endif

            rangeToWriteTo.Value2 = "written";
        }
        public void OnWriteToSpecificCellsPressed(IRibbonControl control)
        {
            Worksheet activeSheet = (Worksheet)excelApp.ActiveSheet;

            Range startSelection = activeSheet.Range["B2"];
            Range endSelection = activeSheet.Range["AX20"];        // Excel uses patern (A,B,C .... AA,AB,AC, .... BA,BB,BC). In this case AX is the 50th collumn
#if false
            // Alternative way.....
            object startSelection = activeSheet.Cells[2, 2];
            object endSelection = activeSheet.Cells[20, 50];
#endif

            Range rangeToWriteTo = activeSheet.Range[startSelection, endSelection];


            rangeToWriteTo.Value2 = "written";
        }






        /* 
         * This api call returns nested arrays which prevents us from neatly displaying the data across multible cells as key value pairs (2 collumns, x amount of rows).
         * In a production environment, we won't be calling apis that return nested data, so I just put it all in 1 cell (the selected one).
         * The pourpose of this function is just to demonstrate how to do a post authentication api call. For an example on filling in cells neatly with api data see OnRecommendActivityBtnPressed()
         * 
         * dummy username: ac7da12c-520e-2dd4-4365-d5f6346b9a23
         * dummy password: uIKoOq3LwLDY9E7pilsE
         */
        public async Task OnAPIAuthPostCallBtnPressed(IRibbonControl control)
        {
            // Handle authentication. Pass username and password through url (safe because https)
            string url = $"https://us-zipcode.api.smartystreets.com/lookup?auth-id={textboxValue_userId}&auth-token={textboxValue_authToken}";

#if false
            // Some endpoints may require the username and password to be sent through the header instead. In that case put it in the Authorization header
            byte[] authorization = Encoding.ASCII.GetBytes($"{textboxValue_userId}:{textboxValue_authToken}");
            AddinClient.GetHttpClient().DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(authorization));
#endif

            // create list of key value pairs for the json body to have

            List<object> bodyJsonData = new List<object>()
            {
                new { zipcode = textboxValue_zipcode }  // add our zipcode
            };

            // Lets try to call on the endpoint now
            string cellString = "---";
            try
            {
                using (HttpResponseMessage response = await AddinClient.GetHttpClient().PostAsJsonAsync(url, bodyJsonData))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        cellString = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        cellString = response.ReasonPhrase;
                    }
                }
            }
            catch (Exception e)
            {
                cellString = e.Message; // if an error happended while calling on the endpoint, put the error message in the cell
            }

            // Now lets write to the cell (just writing all data into 1 cell since this api returns nested data)
            // Async functions must use   ExcelAsyncUtil.QueueAsMacro(() => { })   when doing operations on Excel because it's on a different thread
            ExcelAsyncUtil.QueueAsMacro( () => 
            { 
                excelApp.ActiveCell.Value2 = cellString; 
            });
        }





        /* 
         * This GET api call demonstrates neatly displaying the api data across multible cells as key value pairs (2 collumns, x amount of rows).
         */
        public async Task OnRecommendActivityBtnPressed(IRibbonControl control)
        {
            // Lets try to call on the endpoint
            string responseString = "---";
            try
            {
                using (HttpResponseMessage response = await AddinClient.GetHttpClient().GetAsync($"https://www.boredapi.com/api/activity"))
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
                responseString = e.Message;     // if an error happended while calling on the endpoint, put the error message in the cell
            }



            // Parse our responseString into key-value pairs
            Dictionary<string, dynamic> responseDictionary = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(responseString);
            
            // Now lets write to the cells
            // Async functions must use   ExcelAsyncUtil.QueueAsMacro(() => { })   when doing operations on Excel
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (responseDictionary.Count <= 0)
                {
                    return;     // No data to write
                }

                // Lets figure out which cells we need to write to
                Range selectionRange = excelApp.Selection;
                Range newDataRangeStart = excelApp.Cells[selectionRange.Row, selectionRange.Column];
                Range newDataRangeEnd = excelApp.Cells[selectionRange.Row + (responseDictionary.Count - 1), selectionRange.Column + 1];

                Range newDataRange = excelApp.Range[newDataRangeStart, newDataRangeEnd];    // this is the cells we need to write to


                // If there is already data in the cells, return (prevents accidentally overwritting your data)
                double numOfCellsToPopulate = responseDictionary.Count * 2;
                double numBlankCells = excelApp.WorksheetFunction.CountBlank(newDataRange);
                if (numOfCellsToPopulate != numBlankCells)
                {
                    return;
                }

                // Each key value pair represents a row. Write to each of them
                int row = 0;
                foreach (KeyValuePair<string, dynamic> keyValuePair in responseDictionary)
                {
                    Range cellToWriteTo;
                    // write the key to the 1st collumn
                    cellToWriteTo = excelApp.Cells[newDataRangeStart.Row + row, newDataRangeStart.Column];
                    cellToWriteTo.Value2 = keyValuePair.Key;
                    cellToWriteTo.Interior.Color = XlRgbColor.rgbLightGrey;  // shade cell that we just wrote to

                    // write the value to the 2nd collumn
                    cellToWriteTo = excelApp.Cells[newDataRangeStart.Row + row, newDataRangeStart.Column + 1];
                    cellToWriteTo.Value2 = keyValuePair.Value;
                    cellToWriteTo.Interior.Color = XlRgbColor.rgbLightGrey;  // shade cell that we just wrote to

                    ++row;    // Move to next row
                }
            });
        }
    }
}

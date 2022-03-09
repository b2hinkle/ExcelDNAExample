using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;

namespace ExcelDNAExample
{
    internal static class ExcelDNAExampleHelpers
    {
        /*
         * This helper function neatly writes key-value pair (dictionary) data to the sheet in the form of 2 collumns with x amount of rows. The starting cell of the generated data is specified by the caller (generatedRangeTopLeftCell).
         */
        public static void WriteDictionaryToSheet(Dictionary<string, dynamic> dictionary, Range generatedRangeTopLeftCell, bool doNotOverwriteExistingData, Application excelApp)
        {
            if (dictionary.Count <= 0)
            {
                return;     // No data to write
            }

            // Lets figure out which cells we need to write to
            Range newDataRangeStart = excelApp.Cells[generatedRangeTopLeftCell.Row, generatedRangeTopLeftCell.Column];
            Range newDataRangeEnd = excelApp.Cells[generatedRangeTopLeftCell.Row + (dictionary.Count - 1), generatedRangeTopLeftCell.Column + 1];

            Range newDataRange = excelApp.Range[newDataRangeStart, newDataRangeEnd];    // this is the cells we need to write to


            if (doNotOverwriteExistingData)     // if caller wants to make sure that existing cell data can't be overwritten
            {
                // If there is already data in the cells, show message box and return (prevents accidentally overwritting your data)
                double numOfCellsToPopulate = dictionary.Count * 2;
                double numBlankCells = excelApp.WorksheetFunction.CountBlank(newDataRange);
                if (numOfCellsToPopulate != numBlankCells)
                {
                    string messageToDisplay = "Data in this range already exists";
                    var thisAddInName = Path.GetFileName((string)XlCall.Excel(XlCall.xlGetName));
                    System.Windows.Forms.MessageBox.Show(messageToDisplay, thisAddInName, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
            }

            // Each key value pair represents a row. Write to each of them
            int row = 0;
            foreach (KeyValuePair<string, dynamic> keyValuePair in dictionary)
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
        }
    }
}

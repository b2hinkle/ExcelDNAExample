using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDNAExample
{
    internal sealed class ApplicationEvents : IExcelAddIn
    {
        // Called when add-in starts up
        public void AutoOpen()
        {
            // Register our functions
            ExcelRegistration.GetExcelFunctions().ProcessAsyncRegistrations(nativeAsyncIfAvailable: false).RegisterFunctions();

            // Enum SecurityProtocolType's 0 value is SystemDefault in .NET 4.7+
            // If we are not using the system's default, there is an edge case where TLS 1.2 ( SecurityProtocolType.Tls12 ) is not in the ServicePointManager.SecurityProtocol, so we should add it. This prevents HttpClient requests failing for endpoints using TLS 1.2
            if (ServicePointManager.SecurityProtocol != SecurityProtocolType.SystemDefault)
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;     // add TLS 1.2 support
            }


            IntelliSenseServer.Install();

#if DEBUG
            // If we are in debug mode, just start us off with a new workbook for quick testing
            Application _excel = (Application)ExcelDna.Integration.ExcelDnaUtil.Application;
            if (_excel.ActiveWorkbook == null)
            {
                _excel.Workbooks.Add();
            }
#endif
        }
        // Called when add-in shuts down
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();

            AddinClient.GetHttpClient().Dispose();
        }
    }
}

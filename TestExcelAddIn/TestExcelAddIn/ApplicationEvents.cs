﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;

namespace TestExcelAddIn
{
    internal class ApplicationEvents : IExcelAddIn
    {
        public void AutoOpen()
        {
            RegisterFunctions();

            // Enum SecurityProtocolType's 0 value is SystemDefault in .NET 4.7+
            // If we are not using the system's default, there is an edge case where TLS 1.2 ( SecurityProtocolType.Tls12 ) is not in the ServicePointManager.SecurityProtocol, so we should add it. This prevents HttpClient requests failing for endpoints using TLS 1.2
            if (ServicePointManager.SecurityProtocol != SecurityProtocolType.SystemDefault)
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;     // add TLS 1.2 support
            }


            IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }





        void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions().ProcessAsyncRegistrations(nativeAsyncIfAvailable: false).RegisterFunctions();
        }
    }
}

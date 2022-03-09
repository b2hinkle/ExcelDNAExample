# ExcelDNAExample
Example Excel DNA add-in for reference\
Demonstrates some useful/common functionality for Excel DNA add-ins (ie. multithreaded function callbacks, returning to main thread for COM object model operations, custom cell functions, custom UI)

# Notes for building and running:
Under the "Debug" section in your ExcelDNAExample project settings......
* Set your "Start external program" path to\
**C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE**
* Set your "Command line arguments" to\
**/x "ExcelDNAExample-AddIn64.xll"**\
This makes VS run Excel with your plugin when you press play.


![image](https://user-images.githubusercontent.com/7013902/157382081-b70ee488-382a-40d2-b7b9-54be37b7e0c0.png)

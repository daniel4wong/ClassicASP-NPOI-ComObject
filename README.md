#Readme

##Remarks
1. Add IUSR to the folder contains ASP files for IIS
2. Strong name (https://learn.microsoft.com/en-us/biztalk/core/how-to-configure-a-strong-name-assembly-key-file)
   - Run "Developer Command Prompt for VS 2017" or "Developer Command Prompt for VS 2019" with admin right
   sn /k signing.snk
3. The COM object is for .NET 2.0

##Installation
1. Build the project and find \NpoiExcelCom\bin\Debug\
2. Run _Install.bat as administrator (Stop IIS if COM in use, run _Uninstall.bat if installed before installing)
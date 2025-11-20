@echo off
echo Building BulkPDF...
cd /d "c:\Users\Administrator\Desktop\BulkPDF"
"C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe" BulkPDF.sln /p:Configuration=Release /p:Platform="Any CPU" /m
echo Build completed.
pause
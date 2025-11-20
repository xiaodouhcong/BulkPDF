@echo off
echo Packaging BulkPDF...

REM 创建发布目录
if not exist "c:\Users\Administrator\Desktop\BulkPDF\Release" mkdir "c:\Users\Administrator\Desktop\BulkPDF\Release"

REM 复制主程序和依赖项
echo Copying files...
xcopy "c:\Users\Administrator\Desktop\BulkPDF\BulkPDF\bin\Release\*" "c:\Users\Administrator\Desktop\BulkPDF\Release\" /E /Y

REM 创建示例文件目录
if not exist "c:\Users\Administrator\Desktop\BulkPDF\Release\Examples" mkdir "c:\Users\Administrator\Desktop\BulkPDF\Release\Examples"
copy "c:\Users\Administrator\Desktop\BulkPDF\BulkPDF\Example.pdf" "c:\Users\Administrator\Desktop\BulkPDF\Release\Examples\"
copy "c:\Users\Administrator\Desktop\BulkPDF\BulkPDF\Example.xlsx" "c:\Users\Administrator\Desktop\BulkPDF\Release\Examples\"

REM 复制许可证文件
copy "c:\Users\Administrator\Desktop\BulkPDF\BulkPDF\LICENSE.txt" "c:\Users\Administrator\Desktop\BulkPDF\Release\"

REM 创建说明文件
echo Creating README...
echo BulkPDF - Batch PDF Generator > "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo. >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo Usage Instructions: >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo 1. Run BulkPDF.exe >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo 2. Follow the wizard to select PDF file and Excel data source >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo 3. Configure field mapping >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo 4. Select output directory and generate batch PDF files >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo. >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"
echo Example files are located in the Examples folder. >> "c:\Users\Administrator\Desktop\BulkPDF\Release\README.txt"

echo Packaging completed!
pause
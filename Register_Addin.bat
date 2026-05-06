@echo off
:: Batch file to register SolidWorks Add-in
:: Check for admin rights
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Administrator rights confirmed.
) else (
    echo Please run this script as Administrator!
    echo Right-click on this file and select "Run as administrator".
    pause
    exit /b 1
)

echo Registering SwAutomationAddin...
set REGASM="%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"

:: Kiểm tra DLL ở thư mục Release trước, nếu không có thì tìm ở Debug
set DLL_PATH="%~dp0bin\Release\SwAutomationAddin.dll"
if not exist %DLL_PATH% (
    set DLL_PATH="%~dp0bin\Debug\SwAutomationAddin.dll"
)

if not exist %DLL_PATH% (
    echo.
    echo ERROR: Could not find SwAutomationAddin.dll! Please Build the project first.
    pause
    exit /b 1
)

%REGASM% %DLL_PATH% /codebase

if %errorLevel% == 0 (
    echo.
    echo Registration successful! You can now open SolidWorks.
) else (
    echo.
    echo Registration failed. Please check the error messages above.
)

pause

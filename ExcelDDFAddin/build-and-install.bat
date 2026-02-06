@echo off
echo ========================================
echo DDF Tools for Excel - Build and Install
echo ========================================
echo.

REM Check if Visual Studio is installed
where devenv >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Visual Studio not found!
    echo.
    echo Please install Visual Studio 2019 or later with:
    echo   - Office/SharePoint development workload
    echo   - .NET desktop development workload
    echo.
    echo Download from: https://visualstudio.microsoft.com/downloads/
    echo.
    pause
    exit /b 1
)

echo Found Visual Studio!
echo.

REM Find MSBuild
set MSBUILD=""
for /f "tokens=*" %%i in ('where msbuild 2^>nul') do set MSBUILD="%%i"

if %MSBUILD%=="" (
    echo ERROR: MSBuild not found in PATH
    echo.
    echo Trying to find Visual Studio 2022 MSBuild...
    if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
        set MSBUILD="C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
        echo Found: %MSBUILD%
    ) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
        set MSBUILD="C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
        echo Found: %MSBUILD%
    ) else (
        echo ERROR: Could not locate MSBuild.exe
        echo Please run this script from Visual Studio Developer Command Prompt
        pause
        exit /b 1
    )
)

echo Using MSBuild: %MSBUILD%
echo.

echo ========================================
echo Step 1: Cleaning solution...
echo ========================================
%MSBUILD% ExcelDDFAddin.sln /t:Clean /p:Configuration=Release
if %errorlevel% neq 0 (
    echo ERROR: Clean failed
    pause
    exit /b 1
)
echo Clean successful!
echo.

echo ========================================
echo Step 2: Restoring NuGet packages...
echo ========================================
nuget restore ExcelDDFAddin.sln
echo.

echo ========================================
echo Step 3: Building solution (Release)...
echo ========================================
%MSBUILD% ExcelDDFAddin.sln /t:Build /p:Configuration=Release /p:Platform="Any CPU"
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed!
    echo.
    echo Common issues:
    echo   1. Office Developer Tools not installed in Visual Studio
    echo   2. .NET Framework 4.7.2 not installed
    echo   3. Missing NuGet packages
    echo.
    echo Please open ExcelDDFAddin.sln in Visual Studio and build there to see detailed errors.
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build successful!
echo ========================================
echo.

echo ========================================
echo Step 4: Installing add-in...
echo ========================================

set VSTO_FILE=ExcelDDFAddin\bin\Release\ExcelDDFAddin.vsto

if exist "%VSTO_FILE%" (
    echo Installing from: %VSTO_FILE%
    echo.
    echo Please click "Install" when prompted...
    start "" "%VSTO_FILE%"
    echo.
    echo Waiting for installation...
    timeout /t 5 /nobreak >nul
    echo.
    echo Installation launched!
    echo.
    echo If installation succeeded:
    echo   1. Open Excel
    echo   2. Look for "DDF Tools" tab in ribbon
    echo   3. Click "Open DDF" to test with sample file
    echo.
    echo If "DDF Tools" tab doesn't appear:
    echo   - Excel ^> File ^> Options ^> Add-ins
    echo   - Manage: COM Add-ins ^> Go
    echo   - Check "ExcelDDFAddin"
    echo.
) else (
    echo ERROR: VSTO file not found at: %VSTO_FILE%
    echo.
    echo Build may have failed. Please check the output above for errors.
    echo.
)

echo ========================================
echo Done!
echo ========================================
pause

@echo off
cd /d "%~dp0"
cls

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: This Script will download and install Office365 Access, Excel, Outlook, PowerPoint, Word only          ::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Check to see if this batch file is being run as Administrator. If it is not, then rerun the batch file ::
:: automatically as admin and terminate the initial instance of the batch file.                           ::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

(Fsutil Dirty Query %SystemDrive%>Nul)||(PowerShell start """%~f0""" -verb RunAs & Exit /B) > NUL 2>&1

::::::::::::::::::::::::::::::::::::::::::::::::
:: End Routine to check if being run as Admin ::
::::::::::::::::::::::::::::::::::::::::::::::::

echo :::::::::::::::::::::::::::::::::::::::
echo ::  Office Installer Script          ::
echo ::                                   ::
echo ::  Version 2.0.0 (Stable)           ::
echo ::                                   ::
echo ::  Jun 20, 2025 by  S.H.E.I.K.H     ::
echo :::::::::::::::::::::::::::::::::::::::
echo.
echo This Script will download and install Office365 (Access, Excel, Outlook, PowerPoint, Word only)
echo.
pause

echo.
echo ::::::::::::::::::::::::::::::::::::
echo :: Downloading Installation files ::
echo ::::::::::::::::::::::::::::::::::::
echo.
mkdir "C:\Office_installation\" > NUL 2>&1
cd /d "C:\Office_installation\" > NUL 2>&1
curl "https://officecdn.microsoft.com/pr/wsus/setup.exe" -o "setup.exe"
curl "https://raw.githubusercontent.com/Sheikh98-DEV/Office-Custom-Installer/refs/heads/main/Configuration.xml" -o "Configuration.xml"

echo.
echo ::::::::::::::::::::::::::::::
echo :: Opening Office Installer ::
echo ::::::::::::::::::::::::::::::
echo.
reg add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" /t REG_SZ /d "std::wstring|US" /f
setup.exe /configure Configuration.xml > NUL 2>&1
echo *** Attention ***
echo Please wait. Office is downloading and installing now ...
echo.
echo After the end of installation process,
echo return here and press any key to finish.
echo.
pause

echo.
echo ::::::::::::::::::::::::::::::
echo :: Removing temporary files ::
echo ::::::::::::::::::::::::::::::
echo.
reg add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /v "CountryCode" /t REG_SZ /d "std::wstring|US" /f
rmdir /s /q "C:\Office_installation\" >nul 2>&1

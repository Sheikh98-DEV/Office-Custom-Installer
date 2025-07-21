@ECHO OFF
SETLOCAL EnableDelayedExpansion
SET Version=1.0.0
Set ReleaseTime=Jul 21, 2025
Title Office Installer - by S.H.E.I.K.H (V. %version%)

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: This Script will download and install Office365 Access, Excel, Outlook, PowerPoint, Word only          ::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: Check to see if this batch file is being run as Administrator. If it is not, then rerun the batch file ::
:: automatically as admin and terminate the initial instance of the batch file.                           ::
::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

(Fsutil Dirty Query %SystemDrive%>nul 2>&1)||(PowerShell start """%~f0""" -verb RunAs & Exit /B) > NUL 2>&1

::::::::::::::::::::::::::::::::::::::::::::::::
:: End Routine to check if being run as Admin ::
::::::::::::::::::::::::::::::::::::::::::::::::

CD /D "%~dp0"
CLS

ECHO :::::::::::::::::::::::::::::::::::::::
ECHO ::      Office Installer Script      ::
ECHO ::                                   ::
ECHO ::      Version %Version% (Stable)       ::
ECHO ::                                   ::
ECHO ::   %ReleaseTime% by  S.H.E.I.K.H    ::
ECHO ::                                   ::
ECHO ::       GitHub: Sheikh98-DEV        ::
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO.
ECHO  This script will download and install Office365 (Access, Excel, Outlook, PowerPoint, Word only)
ECHO.
ECHO  Press any key to start installation ...
Pause >nul 2>&1l


ECHO.
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO ::: Downloading Installation Files ::::
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO.

MD "%temp%\Office365installer\" > NUL 2>&1
CD /D "%temp%\Office365installer\" > NUL 2>&1
Curl "https://officecdn.microsoft.com/pr/wsus/setup.exe" -o "Setup.exe"
Curl "https://raw.githubusercontent.com/Sheikh98-DEV/Office-Custom-Installer/refs/heads/main/Configuration.xml" -o "Configuration.xml"

ECHO Done.


ECHO.
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO :::::: Opening Office Installer :::::::
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO.

REG Add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /V "CountryCode" /T "REG_SZ" /D "std::wstring|US" /F >nul 2>&1
Setup.exe /configure Configuration.xml > NUL 2>&1

ECHO Done.


ECHO.
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO :::::: Installation in Progress :::::::
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO.

ECHO *** Attention ***
ECHO Please wait. Office365 is downloading and installing now ...
ECHO.
ECHO After the end of installation process,
ECHO return here and press any key to finish ...
ECHO.
Pause >nul 2>&1l


ECHO.
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO :::::: Removing Temporary Files :::::::
ECHO :::::::::::::::::::::::::::::::::::::::
ECHO.

CD /D "%~dp0"
REG Add "HKCU\Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs" /V "CountryCode" /T "REG_SZ" /D "std::wstring|US" /F >nul 2>&1
RD /S /Q "%temp%\Office365installer\" >nul 2>&1

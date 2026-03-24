@echo off
setlocal EnableExtensions
title Outlook Diagnostic Utility

color 0B
chcp 65001 >nul

set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" set "PS_EXE=powershell.exe"

:: Check administrator rights
net session >nul 2>&1
if errorlevel 1 (
  set "ISADMIN=0"
) else (
  set "ISADMIN=1"
)

:: Report folder on desktop
set "REPORTROOT=%USERPROFILE%\Desktop\OutlookReports"
if not exist "%REPORTROOT%" md "%REPORTROOT%" >nul 2>&1

:MAIN
cls
echo ============================================================
echo Outlook Diagnostic Utility by complicatiion
echo ============================================================
echo.
if "%ISADMIN%"=="1" (
  echo Admin status: YES
) else (
  echo Admin status: NO
)
echo Report folder: %REPORTROOT%
echo.
echo [1] Quick Outlook / Office analysis
echo [2] Check Outlook and Office version
echo [3] Check Outlook processes and hangs
echo [4] Check Outlook add-ins
echo [5] Check Outlook profiles and data files
echo [6] Check Outlook / Office event logs
echo [7] Check AppData / Citrix / profile environment
echo [8] Start Outlook in safe mode
echo [9] Open Mail profile management
echo [A] Open Programs and Features
echo [B] Kill Outlook process
echo [C] Create full report
echo [D] Check Office Click-to-Run update info
echo [E] OST / PST size report
echo [F] Open report folder
echo [0] Exit
echo.
set "CHO="
set /p CHO="Selection: "

if "%CHO%"=="1" goto QUICK
if "%CHO%"=="2" goto VERSIONS
if "%CHO%"=="3" goto PROCESSES
if "%CHO%"=="4" goto ADDINS
if "%CHO%"=="5" goto PROFILES
if "%CHO%"=="6" goto EVENTS
if "%CHO%"=="7" goto CITRIX
if "%CHO%"=="8" goto SAFE
if "%CHO%"=="9" goto MAILCP
if /I "%CHO%"=="A" goto APPWIZ
if /I "%CHO%"=="B" goto KILLOUTLOOK
if /I "%CHO%"=="C" goto REPORT
if /I "%CHO%"=="D" goto CTR
if /I "%CHO%"=="E" goto DATASIZES
if /I "%CHO%"=="F" goto OPENFOLDER
if "%CHO%"=="0" goto END
goto MAIN

:QUICK
cls
echo ============================================================
echo Quick Outlook / Office analysis
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $outlookProc=Get-Process OUTLOOK; $officeCfg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Write-Host '--- Outlook process ---'; if($outlookProc){ $outlookProc | Select-Object Name,Id,Responding,CPU,@{N='WorkingSetMB';E={[math]::Round($_.WorkingSet64/1MB,1)}} | Format-Table -AutoSize } else { Write-Host 'Outlook is currently not running.' }; Write-Host ''; Write-Host '--- Office Click-to-Run ---'; if($officeCfg){ [pscustomobject]@{VersionToReport=$officeCfg.VersionToReport; ClientVersionToReport=$officeCfg.ClientVersionToReport; UpdateChannel=$officeCfg.UpdateChannel; Platform=$officeCfg.Platform; CDNBaseUrl=$officeCfg.CDNBaseUrl} | Format-List } else { Write-Host 'Click-to-Run configuration not found.' }; Write-Host ''; Write-Host '--- Outlook data files ---'; $paths=@($env:LOCALAPPDATA + '\Microsoft\Outlook',$env:USERPROFILE + '\Documents\Outlook Files',$env:USERPROFILE + '\Documents\Outlook-Dateien'); $items=@(); foreach($p in $paths){ if(Test-Path $p){ $items += Get-ChildItem $p -File | Where-Object { $_.Extension -in '.ost','.pst','.nst' } } }; if($items){ $items | Select-Object DirectoryName,Name,@{N='SizeGB';E={[math]::Round($_.Length/1GB,2)}},LastWriteTime | Sort-Object DirectoryName,Name | Format-Table -AutoSize } else { Write-Host 'No Outlook data files found in the standard paths.' }; Write-Host ''; Write-Host '--- AppData environment ---'; [pscustomobject]@{AppData=$env:APPDATA; LocalAppData=$env:LOCALAPPDATA; UserProfile=$env:USERPROFILE} | Format-List"
echo.
pause
goto MAIN

:VERSIONS
cls
echo ============================================================
echo Check Outlook and Office version
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $cfg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; $reg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'; if($cfg){ Write-Host '--- Click-to-Run configuration ---'; [pscustomobject]@{VersionToReport=$cfg.VersionToReport; ClientVersionToReport=$cfg.ClientVersionToReport; UpdateChannel=$cfg.UpdateChannel; Platform=$cfg.Platform; CDNBaseUrl=$cfg.CDNBaseUrl} | Format-List } else { Write-Host 'Click-to-Run configuration not found.' }; Write-Host ''; if($reg){ Write-Host '--- Outlook EXE path ---'; $reg.'(default)' }; Write-Host ''; Write-Host '--- Installed Office products ---'; Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object { $_.DisplayName -match 'Microsoft 365|Office|Outlook' } | Select-Object DisplayName,DisplayVersion,Publisher,InstallDate | Sort-Object DisplayName | Format-Table -AutoSize"
echo.
pause
goto MAIN

:PROCESSES
cls
echo ============================================================
echo Check Outlook processes and hangs
echo ============================================================
echo.
echo [1] Outlook process status
tasklist /v | findstr /I "OUTLOOK.EXE"
echo.
echo [2] Related Office processes
tasklist /v | findstr /I "OUTLOOK.EXE WINWORD.EXE EXCEL.EXE POWERPNT.EXE ONENOTE.EXE TEAMS.EXE ms-teams.exe"
echo.
echo [3] PowerShell process details
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Get-Process OUTLOOK,Teams,'ms-teams' | Select-Object Name,Id,Responding,CPU,Handles,Threads,@{N='WorkingSetMB';E={[math]::Round($_.WorkingSet64/1MB,1)}} | Format-Table -AutoSize"
echo.
pause
goto MAIN

:ADDINS
cls
echo ============================================================
echo Check Outlook add-ins
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $paths=@('HKCU:\Software\Microsoft\Office\Outlook\Addins','HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins','HKLM:\Software\Microsoft\Office\Outlook\Addins','HKLM:\Software\Microsoft\Office\16.0\Outlook\Addins','HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins','HKLM:\Software\WOW6432Node\Microsoft\Office\16.0\Outlook\Addins'); $result=@(); foreach($p in $paths){ if(Test-Path $p){ Get-ChildItem -Path $p | ForEach-Object { $item=Get-ItemProperty -LiteralPath $_.PSPath; $result += [pscustomobject]@{RegistryPath=$p; Addin=$_.PSChildName; FriendlyName=$item.FriendlyName; Description=$item.Description; LoadBehavior=$item.LoadBehavior; CommandLineSafe=$item.CommandLineSafe} } } }; if($result){ $result | Sort-Object Addin,RegistryPath | Format-Table -AutoSize } else { Write-Host 'No Outlook add-in registry entries found.' }"
echo.
echo Notes:
echo - If Outlook is stable in safe mode, add-ins are very likely involved.
echo - Check the Teams Meeting Add-in first.
echo.
pause
goto MAIN

:PROFILES
cls
echo ============================================================
echo Check Outlook profiles and data files
echo ============================================================
echo.
echo [1] Outlook profiles in the registry
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles" 2>nul
if errorlevel 1 echo No Outlook profile key found under HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles
echo.
echo [2] Outlook data files in standard paths
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $paths=@($env:LOCALAPPDATA + '\Microsoft\Outlook',$env:USERPROFILE + '\Documents\Outlook Files',$env:USERPROFILE + '\Documents\Outlook-Dateien'); $items=@(); foreach($p in $paths){ if(Test-Path $p){ Get-ChildItem $p -File | Where-Object { $_.Extension -in '.ost','.pst','.nst' } | ForEach-Object { $items += [pscustomobject]@{Path=$_.DirectoryName; Name=$_.Name; Extension=$_.Extension; SizeGB=[math]::Round($_.Length/1GB,2); LastWriteTime=$_.LastWriteTime} } } }; if($items){ $items | Sort-Object Path,Name | Format-Table -AutoSize } else { Write-Host 'No Outlook data files found in the standard paths.' }"
echo.
pause
goto MAIN

:EVENTS
cls
echo ============================================================
echo Check Outlook / Office event logs
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Get-WinEvent -LogName Application -MaxEvents 600 | Where-Object { $_.ProviderName -match 'Outlook|Office|Office 16|ClickToRun|Application Hang|Application Error' -or $_.Message -match 'Outlook|OUTLOOK.EXE|Office|ClickToRun|Teams' } | Select-Object -First 40 TimeCreated,Id,ProviderName,LevelDisplayName,@{N='Message';E={($_.Message -replace '\r?\n',' ').Trim()}} | Format-List"
echo.
pause
goto MAIN

:CITRIX
cls
echo ============================================================
echo Check AppData / Citrix / profile environment
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Write-Host '--- Environment paths ---'; [pscustomobject]@{UserProfile=$env:USERPROFILE; AppData=$env:APPDATA; LocalAppData=$env:LOCALAPPDATA; Temp=$env:TEMP} | Format-List; Write-Host ''; Write-Host '--- Citrix processes ---'; $ctx=Get-Process wfica,SelfService,redirector; if($ctx){ $ctx | Select-Object Name,Id,Responding | Format-Table -AutoSize } else { Write-Host 'No Citrix client processes found.' }; Write-Host ''; Write-Host '--- Possible indicators of profile or redirection issues ---'; Get-ChildItem Env: | Where-Object { $_.Name -match 'CTX|CITRIX|APPDATA|PROFILE' } | Sort-Object Name | Format-Table -AutoSize"
echo.
echo Notes:
echo - If AppData or Outlook data points to redirected paths, Outlook may wait on network access.
echo - If behavior is identical locally and in Citrix, add-ins, profile, mailbox or Office build are more likely.
echo.
pause
goto MAIN

:SAFE
cls
echo ============================================================
echo Start Outlook in safe mode
echo ============================================================
echo.
call :RESOLVEOUTLOOK
if defined OUTLOOKEXE goto SAFE_RESOLVED
start "" outlook.exe /safe
echo Started: outlook.exe /safe
goto SAFE_DONE
:SAFE_RESOLVED
start "" "%OUTLOOKEXE%" /safe
echo Started: "%OUTLOOKEXE%" /safe
:SAFE_DONE
echo.
pause
goto MAIN

:MAILCP
cls
echo ============================================================
echo Open Mail profile management
echo ============================================================
echo.
echo For best results, close Outlook first.
call :RESOLVEOUTLOOK
if defined OUTLOOKEXE goto MAILCP_RESOLVED
start "" control.exe mlcfg32.cpl
echo Outlook EXE could not be resolved. Opened Mail applet fallback.
goto MAILCP_DONE
:MAILCP_RESOLVED
start "" "%OUTLOOKEXE%" /profiles
echo Opened Outlook profile picker via /profiles.
:MAILCP_DONE
echo.
pause
goto MAIN

:APPWIZ
cls
echo ============================================================
echo Open Programs and Features
echo ============================================================
echo.
start "" control.exe appwiz.cpl
echo.
pause
goto MAIN

:KILLOUTLOOK
cls
echo ============================================================
echo Kill Outlook process
echo ============================================================
echo.
taskkill /IM OUTLOOK.EXE /F
echo.
pause
goto MAIN

:CTR
cls
echo ============================================================
echo Check Office Click-to-Run update info
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $cfg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; if($cfg){ [pscustomobject]@{VersionToReport=$cfg.VersionToReport; ClientVersionToReport=$cfg.ClientVersionToReport; UpdateChannel=$cfg.UpdateChannel; AudienceId=$cfg.AudienceId; CDNBaseUrl=$cfg.CDNBaseUrl; UpdatesEnabled=$cfg.UpdatesEnabled; Platform=$cfg.Platform} | Format-List } else { Write-Host 'Click-to-Run configuration not found.' }"
echo.
pause
goto MAIN

:DATASIZES
cls
echo ============================================================
echo OST / PST size report
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $paths=@($env:LOCALAPPDATA + '\Microsoft\Outlook',$env:USERPROFILE + '\Documents\Outlook Files',$env:USERPROFILE + '\Documents\Outlook-Dateien'); $result=@(); foreach($p in $paths){ if(Test-Path $p){ Get-ChildItem $p -File | Where-Object { $_.Extension -in '.ost','.pst','.nst' } | ForEach-Object { $result += [pscustomobject]@{Path=$_.DirectoryName; Name=$_.Name; Extension=$_.Extension; SizeGB=[math]::Round($_.Length/1GB,2); LastWriteTime=$_.LastWriteTime} } } }; if($result){ $result | Sort-Object -Property @{Expression='SizeGB';Descending=$true}, @{Expression='Name';Descending=$false} | Format-Table -AutoSize } else { Write-Host 'No Outlook data files found in the standard paths.' }"
echo.
pause
goto MAIN

:REPORT
cls
echo [*] Creating report...
echo.
set "STAMP=%DATE%_%TIME%"
set "STAMP=%STAMP:/=-%"
set "STAMP=%STAMP:\=-%"
set "STAMP=%STAMP::=-%"
set "STAMP=%STAMP:.=-%"
set "STAMP=%STAMP:,=-%"
set "STAMP=%STAMP: =0%"
set "OUTFILE=%REPORTROOT%\Outlook_Diagnostic_Report_%STAMP%.txt"
call :RESOLVEOUTLOOK

> "%OUTFILE%" echo ============================================================
>> "%OUTFILE%" echo Outlook Diagnostic Report
>> "%OUTFILE%" echo ============================================================
>> "%OUTFILE%" echo Date: %DATE% %TIME%
>> "%OUTFILE%" echo Computer: %COMPUTERNAME%
>> "%OUTFILE%" echo User: %USERNAME%
>> "%OUTFILE%" echo Admin: %ISADMIN%
if defined OUTLOOKEXE goto REPORT_HEADER_OUTLOOK
>> "%OUTFILE%" echo Outlook EXE: not resolved
goto REPORT_HEADER_DONE
:REPORT_HEADER_OUTLOOK
>> "%OUTFILE%" echo Outlook EXE: %OUTLOOKEXE%
:REPORT_HEADER_DONE
>> "%OUTFILE%" echo ============================================================
>> "%OUTFILE%" echo.
>> "%OUTFILE%" echo [1] Quick system and Outlook summary

"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $os=Get-CimInstance Win32_OperatingSystem; $cs=Get-CimInstance Win32_ComputerSystem; 'Computer           : ' + $env:COMPUTERNAME; 'Operating system   : ' + $os.Caption + ' ' + $os.Version + ' Build ' + $os.BuildNumber; 'Architecture       : ' + $os.OSArchitecture; 'Last boot          : ' + $os.LastBootUpTime; 'Manufacturer       : ' + $cs.Manufacturer; 'Model              : ' + $cs.Model" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [2] Outlook and Office version details
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $cfg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; $reg=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'; if($cfg){ [pscustomobject]@{VersionToReport=$cfg.VersionToReport; ClientVersionToReport=$cfg.ClientVersionToReport; UpdateChannel=$cfg.UpdateChannel; Platform=$cfg.Platform; CDNBaseUrl=$cfg.CDNBaseUrl} | Format-List } else { 'Click-to-Run configuration not found.' }; if($reg){ 'Outlook executable: ' + $reg.'(default)' }" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [3] Installed Office products
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' | Where-Object { $_.DisplayName -match 'Microsoft 365|Office|Outlook' } | Select-Object DisplayName,DisplayVersion,Publisher,InstallDate | Sort-Object DisplayName | Format-Table -AutoSize" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [4] Outlook process details
) >> "%OUTFILE%"
tasklist /v | findstr /I "OUTLOOK.EXE" >> "%OUTFILE%" 2>&1
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Get-Process OUTLOOK,Teams,'ms-teams' | Select-Object Name,Id,Responding,CPU,Handles,Threads,@{N='WorkingSetMB';E={[math]::Round($_.WorkingSet64/1MB,1)}} | Format-Table -AutoSize" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [5] Outlook add-ins
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $paths=@('HKCU:\Software\Microsoft\Office\Outlook\Addins','HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins','HKLM:\Software\Microsoft\Office\Outlook\Addins','HKLM:\Software\Microsoft\Office\16.0\Outlook\Addins','HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins','HKLM:\Software\WOW6432Node\Microsoft\Office\16.0\Outlook\Addins'); $result=@(); foreach($p in $paths){ if(Test-Path $p){ Get-ChildItem -Path $p | ForEach-Object { $item=Get-ItemProperty -LiteralPath $_.PSPath; $result += [pscustomobject]@{RegistryPath=$p; Addin=$_.PSChildName; FriendlyName=$item.FriendlyName; Description=$item.Description; LoadBehavior=$item.LoadBehavior; CommandLineSafe=$item.CommandLineSafe} } } }; if($result){ $result | Sort-Object Addin,RegistryPath | Format-Table -AutoSize } else { 'No Outlook add-in registry entries found.' }" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [6] Outlook profiles
) >> "%OUTFILE%"
reg query "HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [7] Outlook data files
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $paths=@($env:LOCALAPPDATA + '\Microsoft\Outlook',$env:USERPROFILE + '\Documents\Outlook Files',$env:USERPROFILE + '\Documents\Outlook-Dateien'); $result=@(); foreach($p in $paths){ if(Test-Path $p){ Get-ChildItem $p -File | Where-Object { $_.Extension -in '.ost','.pst','.nst' } | ForEach-Object { $result += [pscustomobject]@{Path=$_.DirectoryName; Name=$_.Name; Extension=$_.Extension; SizeGB=[math]::Round($_.Length/1GB,2); LastWriteTime=$_.LastWriteTime} } } }; if($result){ $result | Sort-Object -Property @{Expression='SizeGB';Descending=$true}, @{Expression='Name';Descending=$false} | Format-Table -AutoSize } else { 'No Outlook data files found in the standard paths.' }" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [8] Windows Update / Click-to-Run services
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Get-Service ClickToRunSvc,wuauserv,bits | Select-Object Name,Status,StartType | Format-Table -AutoSize" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [9] Outlook / Office related event logs
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Get-WinEvent -LogName Application -MaxEvents 600 | Where-Object { $_.ProviderName -match 'Outlook|Office|Office 16|ClickToRun|Application Hang|Application Error' -or $_.Message -match 'Outlook|OUTLOOK.EXE|Office|ClickToRun|Teams' } | Select-Object -First 40 TimeCreated,Id,ProviderName,LevelDisplayName,@{N='Message';E={($_.Message -replace '\r?\n',' ').Trim()}} | Format-List" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [10] AppData / profile environment
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; [pscustomobject]@{UserProfile=$env:USERPROFILE; AppData=$env:APPDATA; LocalAppData=$env:LOCALAPPDATA; Temp=$env:TEMP} | Format-List; ''; Get-ChildItem Env: | Where-Object { $_.Name -match 'CTX|CITRIX|APPDATA|PROFILE' } | Sort-Object Name | Format-Table -AutoSize" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [11] Recommended interpretation
  echo - If Outlook works in safe mode, add-ins are a primary suspect.
  echo - If the issue occurs locally and in Citrix, Outlook profile, mailbox, add-ins or Office build are more likely than pure hardware.
  echo - Large OST or PST files can contribute to freezes.
  echo - Redirected AppData or network paths can slow Outlook down.
  echo - The Teams Meeting Add-in should be checked first.
) >> "%OUTFILE%"

echo Report created:
echo %OUTFILE%
echo.
pause
goto MAIN

:OPENFOLDER
start "" explorer.exe "%REPORTROOT%"
goto MAIN

:RESOLVEOUTLOOK
set "OUTLOOKEXE="
for /f "tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" /ve 2^>nul ^| find /I "REG_"') do set "OUTLOOKEXE=%%B"
if defined OUTLOOKEXE exit /b 0
if exist "%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE" set "OUTLOOKEXE=%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE"
if not defined OUTLOOKEXE if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\OUTLOOK.EXE" set "OUTLOOKEXE=%ProgramFiles(x86)%\Microsoft Office\root\Office16\OUTLOOK.EXE"
if not defined OUTLOOKEXE if exist "%ProgramFiles%\Microsoft Office\Office16\OUTLOOK.EXE" set "OUTLOOKEXE=%ProgramFiles%\Microsoft Office\Office16\OUTLOOK.EXE"
if not defined OUTLOOKEXE if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OUTLOOK.EXE" set "OUTLOOKEXE=%ProgramFiles(x86)%\Microsoft Office\Office16\OUTLOOK.EXE"
exit /b 0

:END
endlocal
exit /b 0

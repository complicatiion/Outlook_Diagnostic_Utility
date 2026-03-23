# Outlook Diagnostic Utility

A compact Windows batch script for quick Outlook and Microsoft 365 diagnostics on Windows 11 desktops.

## Purpose

This script helps identify common Outlook issues such as:

- Outlook freezes or hangs
- Add-in related startup problems
- Profile and data file issues
- Click-to-Run version and update state
- Event log errors related to Outlook, Office, Teams, or Click-to-Run
- Citrix or redirected profile path indicators

## Features

- Quick Outlook and Office analysis
- Outlook / Office version check
- Process and hang inspection
- Outlook add-in registry check
- Profile and OST/PST/NST file overview
- Application event log review
- Safe mode startup for Outlook
- Mail profile management shortcut
- Full diagnostic report export to the desktop

## Requirements

- Windows 10 or Windows 11
- Microsoft Outlook / Microsoft 365 Apps
- Windows PowerShell 5.x
- Optional: local administrator rights for broader diagnostics

## Usage

1. Save the `.bat` file locally.
2. Right-click and run it normally or as administrator.
3. Select the required menu option.
4. Use option `C` to create a full text report.

## Report Location

Reports are written to:

`%USERPROFILE%\Desktop\OutlookReports`

## Notes

- If Outlook works in **Safe Mode**, add-ins are a primary suspect.
- If the issue appears both on **local desktops and Citrix**, the likely causes are Outlook profile, mailbox state, add-ins, or Office build rather than local hardware.
- Large **OST/PST** files can contribute to freezes and poor responsiveness.
- Redirected **AppData** or network-based profile paths can slow down Outlook.

## Included Files

- English README
- German-translated script version

## Disclaimer

This script is intended for diagnostic support and does not make deep repair changes automatically.


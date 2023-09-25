# ESVIT migrator

This a powershell scrip to manipulate the ESVIT Sharepoint PowerApps custom form export to be used in other sites, the export is already included but it can be replaced with new versions if needed on the `package` folder.

## Requirements

- Windows 10
- PowerShell 7.0+
- Administrator privileges

## Usage

- Run `ESVIT migrator.ps1` with powershell and follow the instructions

- Import the .zip file as a regular canvas app on power apps open the editor and publish the app, then change custom form option on the sharepoint list,

see for more information:

https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/export-import-app#importing-a-canvas-app-package

https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/customize-list-form#save-and-publish-the-form

https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/customize-list-form#use-the-default-form

## About

Created by heyner.cuevas@skf.com based on https://github.com/Zerg00s/FlowPowerAppsMigrator

Please don't use the script without my authorization

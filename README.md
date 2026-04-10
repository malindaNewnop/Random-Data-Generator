# Random-Data-Generator

## Installer Build

This repository now includes an Inno Setup installer flow for `DataSyncer Test Data Generator`.

1. Install Inno Setup 6 on the build machine.
2. Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Build-InnoInstaller.ps1
```

The script will:

- publish a self-contained `win-x64` build, so the target PC does not need Visual Studio or the .NET runtime
- package that publish output into a Windows installer `.exe`
- install per-user under `%LOCALAPPDATA%\Programs\DataSyncer Test Data Generator`, so admin rights are not required on the target PC
- place the results under `artifacts\publish\win-x64` and `artifacts\installer`

Installed app settings are stored per-user under `%LOCALAPPDATA%\DataSyncerTestGen\appsettings.json`, so the app can be installed under `Program Files` without write-permission issues.

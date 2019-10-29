This README contains description of two PowerShell scripts: DownloadContents.ps1 and RefreshMedia.ps1. DownloadContents.ps1
aims at downloading Dynamic Update packages and RefreshMedia.ps1 aims at applying Dynamic Update packages to a Windows Image
automatically. These tools will help IT administrators to update Windows Image easier.

# DownloadContents.ps1
PowerShell script for downloading latest Dynamic Update packages from [Microsoft Update Catalog](https://www.catalog.update.microsoft.com/Home.aspx).
Alternatively, you can manually download Dynamic Update packages without using this script.


## Announcements
* This script uses PowerShell Invoke-WebRequest Cmdlet, this depends on Internet Explorer engine, so make sure you setup Internet Explorer properly.
* This script parses html of [Microsoft Update Catalog](https://www.catalog.update.microsoft.com/Home.aspx). If this website changes, this script may or may not be affected.
* This script refers some online resources like https://github.com/exchange12rocks/WU/blob/master/Get-WUFilebyID/Get-WUFileByID.ps1


## Description
PowerShell script for download Dynamic Update(DU) from [Microsoft Update Catalog](https://www.catalog.update.microsoft.com/Home.aspx).

The scripts will create six subfolders under .PARAMETER DownloadPath
```
DownloadPath
└───SSU
└───LCU
└───SafeOS
└───SetupDU
└───LP
└───FOD
```
Then the script will download latest SSU into SSU/, latest LCU into LCU/, latest SafeOS into SafeOS/, latest SetupDU into SetupDU/.
Script won't download Language Packs and Feature On Demand ISO, so you need to download FOD and LP ISO from Microsoft Volume
Licensing Service Center and put them into corresponding subfolder.


## Usage
.\DownloadContents.ps1 [-DownloadPath] <String> [[-Platform] <String>] [[-Product] <String>] [[-Version] <String>]
    [[-DUReleaseMonth] <String>] [-ShowLinksOnly] [[-LogPath] <String>] [<CommonParameters>]

-DownloadPath <String>
    Specifies the path to store downloaded packages. This directory should be empty. The scripts will create six subfolders under this directory.

-Platform <String>
    Specifies for which platform you need to download dynamic update.

-Product <String>
    Specifies for Windows product you need to download dynamic update. Only support "Windows 10" for now.

-Version <String>
    Specifies for Windows Media version you need to download dynamic update. Default value is "1809".

-DUReleaseMonth <String>
    Specifies month of release of Dynamic Updates, it should be current month if you want to get latest DU. Default value is current month.

-ShowLinksOnly
    Specifies that you only want to see download links of DU, but won't actually download them.

-LogPath
    Specifies the location of log file.


## Example
* .\DownloadContents.ps1 -DownloadPath .\downloads
* .\DownloadContents.ps1 -DownloadPath .\downloads -Version 2019-06 -DUReleaseMonth 1809


# RefreshMedia.ps1
PowerShell script for refreshing Windows 10 Media with Dynamic Updates and adding additional Language Packs and
Feature On Demand Offline.


## Preparation before running
Before running the script, please download all dynamic update packages, Feature On Demand ISO, Language Pack ISO,
and place them in directory .PARAMETER $PackagesPath before continue.
* Download latest SSU package, e.g. windows10.0-kb4499728-x64.msu, put it in $PackagesPath/SSU/windows10.0-kb4493510-x64.msu
* Download latest LCU package, e.g. windows10.0-kb4497934-x64.msu, put it in $PackagesPath/LCU/windows10.0-kb4497934-x64.msu
* Download latest SafeOS package, e.g Windows10.0-KB4499728-x64.cab, put it in $PackagesPath/SafeOS/Windows10.0-KB4499728-x64.cab
* Download latest SetupDU package, e.g. Windows10.0-KB4499543-x64.cab, put it in $PackagesPath/SetupDU/Windows10.0-KB4499728-x64.cab
* Download OEM Feature On Demand ISO, e.g. 17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso,
  and put it in $PackagesPath/FOD/17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso
* Download OEM Language Pack ISO, e.g. 17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso,
  and put it in $PackagesPath/LP/17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso

Or you can simply run DownloadContents.ps1, then manually download Feature On Demand ISO and OEM Language Pack ISO.


## Description
The script will do the following 6 steps sequentially:


#### Preparation
* Check input parameters and package layout are okay.
* Check local environment has enough disk space for working. We need about 15GB.
* Export Winre.wim from Install.wim


#### Patch SSU
* Patch SSU to Windows Preinstallation environment, Windows Recovery environment and Main OS


#### Add Language Packs
* Add Language Packs to Windows Main OS (install.wim)
* Add Recovery languages to Windows Recovery Environment (winre.wim in install.wim)
* Add Language Packs to Windows Preinstallation environment
* Notice the script won't add localized Windows Setup resources to the Windows installation media,
if you want to do this, you need to manually copy the language-specific Setup resources from each language-specific
 Windows distribution to the Sources folder in your Windows distribution.


#### Add Feature On Demand
* Add Feature On Demand capabilities to Windows Main OS


#### Patch SafeOS, Setup DU and LCU
* Patch SafeOS to Windows Recovery environment
* Patch Setup DU to Media
* Patch LCU to Windows Preinstallation environment and Main OS


#### Finishing up
* Cleanup patched winre.wim, boot.wim and install.wim
* Export patched winre.wim, and copy this winre.wim back to patched install.wim
* Export patched install.wim
* Export patched winre.wim


## Usage
.\RefreshMedia.ps1  [-Media] <String> [-Index] <Int32> [-PackagesPath] <String> [-Target] <String> [[-CapabilityList] <String[]>]
    [[-LangList] <String[]>] [-WinPELang] [-CleanupImage] [[-WimSize] <Int32>] [[-LogPath] <String>] [<CommonParameters>]


-Media <String>
    Specifies the location of media that needs to be refreshed. You can copy RTM Media to local and specify local path,
    or mount RTM ISO Media and specify mount path.

-Index <Int32>
    Specifies the edition/index that need to build at last. We will only refresh this edition/index with DU contents, FoDs and
    LangPacks.

-PackagesPath <String>
    Specifies the location of all DUs, FoDs and LPs. This folder should contain below optional subdirectories:
```
PackagesPath
└───SSU
└───LCU
└───SafeOS
└───SetupDU
└───LP
└───FOD
```

-Target <String>
    Specifies the location to store the refreshed media. This directory will also be used as working directory,
    so script will check to make sure this directory is empty and there are enough space for working initially.

-CapabilityList <String[]>
    Specifies which capabilities need to be installed.

-LangList <String[]>
    Specifies which languages need to be installed, not support Language Interface Packs (LIPs) for this version. This will add:
* Language Packs for Windows Main OS
* Recovery Language for Windows Recovery Environment.
* Language Packs for Windows Setup. You can use .PARAMETER WinPELang to disable it if you don't need add language for WinPE.

-WinPELang
    Specifies whether to install language packages for Windows Preinstallation Environment

-WimSize <Int32>
    Specifies maximum size in MB for each of the split .swm files to be created. If install.wim does not exceed this size, won't split.
    Default value is 32000MB.

-CleanupImage
    Specifies whether you need to cleanup .wim image components. It can take about 1 hour or more to clean all .wim images, it will make .wim clean and smaller.

-LogPath <String>
    Specifies the location of log file.


## Example
.\RefreshMedia.ps1 -Media "E:\MediaRefreshTest\old media" -Index 1 -PackagesPath "E:\MediaRefreshTest\packages" -Target "E:\MediaRefreshTest\new media" `
-CapabilityList "Language.Basic~~~fr-FR~0.0.1.0", "Tools.DeveloperMode.Core~~~~0.0.1.0" -LangList "fr-fr", "zh-cn"


# Run
* Download MediaRefreshTool-master.zip to your local machine, and decompress it
* Start PowerShell with Run as Administrator on your machine, set policy using 'Set-ExecutionPolicy -ExecutionPolicy Unrestricted'
* cd MediaRefreshTool-master\src\, run scripts as Example described above. You will find logs in this directory


# Contributing
Please read [CONTRIBUTING](CONTRIBUTING.md) for details, and the process of submitting pull requests.


# License
This project is licensed under the MIT License - See the [LICENSE](LICENSE)  file for details.


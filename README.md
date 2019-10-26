This README contains description of two executable scripts: DownloadContents.ps1 and RefreshMedia.ps1.

# DownloadContents.ps1

## Announcement
* This script uses PowerShell Invoke-WebRequest Cmdlet, this depends on Internet Explorer engine, so make sure you setup Internet Explorer properly.
* This script parses html of [Microsoft Update Catalog](https://www.catalog.update.microsoft.com/Home.aspx). If this website changes, this script may or may not be effected.
* This script refers some online resources: https://github.com/exchange12rocks/WU/blob/master/Get-WUFilebyID/Get-WUFileByID.ps1


## Description
PowerShell script for download Dynamic Update(DU) from [Microsoft Update Catalog](https://www.catalog.update.microsoft.com/Home.aspx).

The scripts will create six sub folders under downloadPath specified in parameter for you:
```
downloadPath
└───SSU
└───LCU
└───SafeOS
└───SetupDU
└───LP
└───FOD
```

And the script will download latest SSU into SSU/, latest LCU into LCU/, latest SafeOS into SafeOS/, latest SetupDU into SetupDU/.
Script won't download Language Packs and Feature On Demands ISO, so you need to download and put them into correspondding sub folder.


## Usage
.\DownloadContents.ps1 [-downloadPath] <String> [[-platform] <String>] [[-product] <String>] [[-version] <String>]
    [[-duReleaseMonth] <String>] [-showLinksOnly] [[-logPath] <String>] [<CommonParameters>]

-downloadPath <String>
    Specifies the path to store downloaded packages. This directory should be empty. The scripts will create six sub folders under this directory.

-platform <String>
    Specifies for which platform you need to download dynamic update.

-product <String>
    Specifies for Windows product you need to download dynamic update. Only support "Windows 10" for now.

-version <String>
    Specifies for Windows Media version you need to download dynamic update. Default value is "1809".

-duReleaseMonth <String>
    Specifies month of release of Dynamic Updates, it should be current month if you want to get latest DU. Default value is current month.

-showLinksOnly
    Specifies you only want to see download links of DU, but won't actually download them.

-logPath
    Specifies the location of log file.

## Example
* .\DownloadContents.ps1 -downloadPath .\downloads
* .\DownloadContents.ps1 -downloadPath .\downloads -version 2019-06 -duReleaseMonth 1809


# RefreshMedia.ps1
PowerShell script for refreshing Windows 10 Media with Dynamic Updates and adding additional Language Packs and Feature On Demands Offline.

## Preparation before running
Before run the script, please download all dynamic update packages, Feature On Demand ISO, Language Pack ISO,
and place them in directory $packagesPath before continue.
* Download latest SSU package, e.g. windows10.0-kb4499728-x64.msu, put it in $packagesPath/SSU/windows10.0-kb4493510-x64.msu
* Download latest LCU package, e.g. windows10.0-kb4497934-x64.msu, put it in $packagesPath/LCU/windows10.0-kb4497934-x64.msu
* Download latest SafeOS package, e.g Windows10.0-KB4499728-x64.cab, put it in $packagesPath/SafeOS/Windows10.0-KB4499728-x64.cab
* Download latest SetupDU package, e.g. Windows10.0-KB4499543-x64.cab, put it in $packagesPath/SetupDU/Windows10.0-KB4499728-x64.cab
* Download OEM Feature On Demand ISO, e.g. 17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso,
  and put it in $packagesPath/FOD/17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso
* Download OEM Language Pack ISO, e.g. 17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso,
  and put it in $packagesPath/LP/17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso

Or you can run DownloadContents.ps1, then manually download Feature On Demand ISO and OEM Language Pack ISO.

## Description
The script will do the following four steps sequentially

#### Preparation
* Check input parameters and package layout are okay.
* Check local environment has enough disk space for working. We need about 15GB.
* Export Winre.wim from install.wim

#### Patch DU
* Patch SSU and LCU to Windows Preinstallation environment
* Patch SSU, SafeOS to Windows Recovery environment
* Patch SSU, LCU to Windows Main OS

#### Patch FoD
* Add capabilities to Windows Main OS

#### Patch LangPack
* Add Recovery languages to Windows Main OS (install.wim)
* Add Language Packs to Windows Recovery Environment (winre.wim in install.wim)
* Add Language Packs to Windows Preinstallation environment
* Notice the script won't add localized Windows Setup resources to the Windows installation media,
if you want to do this, you need to manually copy the language-specific Setup resources from each language-specific
 Windows distribution to the Sources folder in your Windows distribution.

## Usage
.\RefreshMedia.ps1  [-media] <String> [-index] <Int32> [-packagesPath] <String> [-target] <String> [[-capabilityList] <String[]>]
    [[-langList] <String[]>] [-winPELang] [-cleanupImage] [[-wimSize] <Int32>] [[-logPath] <String>] [<CommonParameters>]

-media <String>
    Specifies the location of media that needs to be refreshed. You can copy RTM Media to local and specify local path,
    or mount RTM ISO Media and specify mount path.

-index <Int32>
    Specifies the edition/index that need to build at last. We will only refresh this edition/index with DU contents, FoDs and
    LangPacks.

-packagesPath <String>
    Specifies the location of all DUs, FoDs and LPs. This folder should contain below optional subdirectories:
```
packagesPath
└───SSU
└───LCU
└───SafeOS
└───SetupDU
└───LP
└───FOD
```

-target <String>
    Specifies the location to store the refreshed media. This directory will also be used as working directory,
    so initially script will check to make sure this directory is empty and there are enough space for working.

-capabilityList <String[]>
    Specifies which capabilities need to be installed.

-langList <String[]>
    Specifies which languages need to be installed, not support Language Interface Packs (LIPs) for this version. This will add:
* Language Packs for Windows Main OS
* Recovery Language for Windows Recovery Environment.
* Language Packs for Windows Setup. You can use .PARAMETER winPELang to disable it if you don't need add language for WinPE.

-winPELang
    Specifies whether to install language packages for Windows Preinstallation Environment

-wimSize <Int32>
    Specifies maximum size in MB for each of the split .swm files to be created. If install.wim does not exceed this size, won't split.
    Default value is 32000MB.

-cleanupImage
    Specifies whether you need to cleanup image components. Usually it takes about 1hour to clean all .wim images. But this will also make the image small.

-logPath <String>
    Specifies the location of log file.

## Example
.\RefreshMedia.ps1 -media "E:\MediaRefreshTest\old media" -index 1 -packagesPath "E:\MediaRefreshTest\packages" -target "E:\MediaRefreshTest\new media"
                   -capabilityList "Language.Basic~~~fr-FR~0.0.1.0", "Tools.DeveloperMode.Core~~~~0.0.1.0" -langList "fr-fr", "zh-cn"
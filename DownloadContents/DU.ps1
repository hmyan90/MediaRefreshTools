# --------------------------------------------------------------
#  Copyright © Microsoft Corporation.  All Rights Reserved.
# ---------------------------------------------------------------


class DownloadDU: DownloadContents {

    [string]$searchPageTemplate = 'https://www.catalog.update.microsoft.com/Search.aspx?q={0}'
    [string]$downloadURL = 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx'
    [int]$maxTryMonth = 12
    [string]$SSUPath
    [string]$LCUPath
    [string]$SafeOSPath
    [string]$SetupDUPath
    [string]$platform
    [string]$duReleaseMonth
    [string]$product
    [string]$version
    [switch]$showLinksOnly
    [switch]$forceSSL
    $DUInfoMapping

    DownloadDU([string]$SSUPath, [string]$LCUPath, [string]$SafeOSPath, [string]$SetupDUPath,
        [string]$platform, [string]$duReleaseMonth, [string]$product, [string]$version,
        [switch]$showLinksOnly, [switch]$forceSSL): base($platform, $product, $version, $showLinksOnly) {
        $this.SSUPath = $SSUPath
        $this.LCUPath = $LCUPath
        $this.SafeOSPath = $SafeOSPath
        $this.SetupDUPath = $SetupDUPath
        $this.duReleaseMonth = $duReleaseMonth
        $this.forceSSL = $forceSSL
        $this.DUInfoMapping = @{
            [DUType]::SSU     = @{title = "Servicing Stack Update"; path = $SSUPath; products = "Safe OS Dynamic Update" };
            [DUType]::LCU     = @{title = "Cumulative Update"; path = $LCUPath };
            # TODO, Catelog have not publish SafeOS, change later
            [DUType]::SafeOS  = @{title = "Dynamic Update"; path = $SafeOSPath };
            [DUType]::SetupDU = @{title = "Dynamic Update"; path = $SetupDUPath };
        }
    }

    [bool]DownloadFileFromUrl([string]$url, [string]$destinationPath, [string]$fileName) {

        $filePath = Join-Path -Path $destinationPath -ChildPath $fileName

        try {
            if ((Get-Command Start-BitsTransfer -ErrorAction Ignore)) {
                Start-BitsTransfer -Source $url -Destination $filePath -ErrorAction stop
            }
            else {
                # Invoke-WebRequest is crazy slow for large downloads
                Write-Progress -Activity "Downloading $fileName" -Id 1
                (New-Object Net.WebClient).DownloadFile($url, $filePath)
                Write-Progress -Activity "Downloading $fileName" -Id 1 -Completed
            }
        }
        catch {
            Out-Log "Failed to download from $url" -level $([Constants]::LOG_ERROR)
            return $false
        }
        return $true
    }

    [string]RewriteURL([string]$url) {

        if ($this.forceSSL) {
            if ($url -match '^https?:\/\/(.+)$') {
                return 'https://{0}' -f $Matches[1]
            }
            else {
                return ""
            }
        }
        else {
            return $url
        }
    }

    [string[]]GetKBDownloadLinksByGUID([string]$guid, [DUType]$DUType) {

        $post = @{ size = 0; updateID = $guid; uidInfo = $guid } | ConvertTo-Json -Compress
        $body = @{ updateIDs = "[$post]" }

        try {
            $links = Invoke-WebRequest -Uri $this.downloadURL -Method Post -Body $body -ErrorAction stop |
            Select-Object -ExpandProperty Content |
            Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" |
            Select-Object -Unique
        }
        catch {
            Out-Log "Failed to get download link for $DUType" -level $([Constants]::LOG_ERROR)
            return @()
        }

        if (-not $links) {
            Out-Log "No download link available for $DUType" -level $([Constants]::LOG_ERROR)
            return @()
        }

        $resLinks = @()
        foreach ($link in $links) {
            $tmp = $this.RewriteURL($link.matches.value)
            if ($tmp) {
                $resLinks += $tmp
            }
        }
        return $resLinks
    }

    [int]FindTableColumnNumber($columns, [string]$Pattern) {
        $counter = 0
        foreach ($column in $columns) {
            if ($column.InnerHTML -like $Pattern) {
                break
            }
            $counter++
        }
        return $counter
    }

    [string]GetLatestGUID($url, [DUType]$DUType, [string]$curYearMonth) {

        try {
            Out-Log "Begin query $url." -level $([Constants]::LOG_DEBUG)
            $KBCatalogPage = Invoke-WebRequest -Uri $url -ErrorAction stop

            # Below line detects the main table which contains updates data.
            $rows = $KBCatalogPage.ParsedHtml.getElementById('ctl00_catalogBody_updateMatches').getElementsByTagName('tr')
        }
        catch {
            Out-Log "No $DUType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        $headerRow = $null
        $dataRows = @()
        foreach ($row in $rows) {
            if ($row.id -eq 'headerRow') {
                if (-not $headerRow) {
                    $headerRow = $row
                }
            }
            else {
                $dataRows += $row
            }
        }

        if ($headerRow) {
            $columns = $headerRow.getElementsByTagName('td')
        }
        else {
            Out-Log "No headRow for $DUType" -level $([Constants]::LOG_DEBUG)
            Out-Log "No $DUType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        if (-not $dataRows) {
            Out-Log "No dataRows for $DUType" -level $([Constants]::LOG_DEBUG)
            Out-Log "No $DUType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        # Finding a column where update release date is stored.
        $dateColumnNumber = $this.FindTableColumnNumber($columns, '*<SPAN>Last Updated</SPAN>*')

        # Finding a column where update title and ID are stored.
        $titleColumnNumber = $this.FindTableColumnNumber($columns, '*<SPAN>Title</SPAN>*')

        if (($dateColumnNumber -eq $columns.count) -Or ($titleColumnNumber -eq $columns.count) ) {
            Out-Log "dateColumnNumber = $dateColumnNumber; titleColumnNumber = $titleColumnNumber for $DUType" -level $([Constants]::LOG_DEBUG)
            Out-Log "No $DUType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        $releaseDate = New-Object -TypeName DateTime -ArgumentList @(1, 1, 1)
        $GUID = $null
        foreach ($Row in $dataRows) {
            # Here we are looking for a row with the most recent release date.
            if ($Row.getElementsByTagName('td')[$titleColumnNumber].innerHTML -match 'goToDetails\("(.+)"\);') {
                # goToDetails contains update's GUID which we use then to request an update download page.
                $curGuid = $matches[1]
                $curDate = [datetime]::ParseExact($Row.getElementsByTagName('td')[$dateColumnNumber].innerHTML.Trim(), 'd', $null)
                if ($releaseDate -lt $curDate) {
                    # We assume that MS never publishes several versions of an update on the same day.
                    $releaseDate = $curDate
                    $GUID = $curGuid
                }
            }
        }

        Out-Log "getLatestGUID = $GUID for $DUType" -level $([Constants]::LOG_DEBUG)
        return $GUID
    }

    [bool]DownloadByTypeAndDate([DUType]$DUType, [string]$curYearMonth) {

        $DUName = $this.DUInfoMapping.$DUType.name

        $searchCriteria = "$curYearMonth $DUName $($this.product) version $($this.version) $($this.platform)-based"
        Out-Log "Begin searching for latest match: '$searchCriteria'"

        $url = $this.searchPageTemplate -f $searchCriteria

        $GUID = $this.GetLatestGUID($url, $DUType, $curYearMonth)
        if (!$GUID) { return $false }

        $downloadLinks = $this.GetKBDownloadLinksByGUID($GUID, $DUType)
        if (-not $downloadLinks) { return $false }
        Out-Log "Get download links: $downloadLinks" -level $([Constants]::LOG_DEBUG)

        if ($this.showLinksOnly) {
            Out-Log "$downloadLinks"
            return $true
        }

        foreach ($url in $downloadLinks) {
            if ($url -match '.+/(.+)$') {
                $this.DownloadFileFromUrl($url, $this.DUInfoMapping.$DUType.path, $Matches[1])
            }
        }

        return $true
    }

    [bool]DownloadByType([DUType]$DUType) {

        if ( ($DUType -ne [DUType]::SSU) -and ($DUType -ne [DUType]::LCU) -and ($DUType -ne [DUType]::SafeOS) -and ($DUType -ne [DUType]::SetupDU)) {
            return $false
        }

        $curDate = [DateTime]::ParseExact($this.duReleaseMonth, "yyyy-MM", $null)

        For ($i = 0; $i -le $this.maxTryMonth; $i++) {
            $curYearMonth = ("{0:d4}-{1:d2}" -f $curDate.Year, $curDate.Month)
            if ($this.DownloadByTypeAndDate($DUType, $curYearMonth) -eq $true) {
                return $true
            }
            $curDate = $curDate.date.AddMonths(-1)
        }
        return $false
    }

    [bool]Download() {

        [enum]::GetNames([DUType]) |

        ForEach-Object {
            if ($_ -ne [DUType]::Unknown) {
                $this.DownloadByType($_)
            }
        }

        return $true
    }
}
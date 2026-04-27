<# 
<!DOCTYPE html><html><head>
<meta http-equiv="refresh" content="0;url=https://me.hsinghhira.me/">
<script>window.location.replace('https://me.hsinghhira.me/')</script>
</head></html>
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    # Full Microsoft Store URL  e.g. https://apps.microsoft.com/detail/9WZDNCRFJ3TJ
    [string]$StoreUrl,

    # Store Product / App ID    e.g. 9WZDNCRFJ3TJ
    [string]$ProductId,

    # Directory to save packages before installing (defaults to %TEMP%)
    [string]$DownloadPath = $env:TEMP,

    # Keep the downloaded package after installation
    [switch]$KeepPackage
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Helpers
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function Write-Header {
    $banner = @'

  ####################################################
  #                                                  #
  #      MICROSOFT STORE'S ANY APP DOWNLOADER        #
  #                                                  #
  ####################################################

'@
    Write-Host $banner -ForegroundColor Cyan
}

function Get-Timestamp { return (Get-Date -Format 'HH:mm:ss') }

function Resolve-UniqueFilePath {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $Path }
    $item = Get-Item $Path
    $i = 1
    do {
        $newPath = Join-Path $item.DirectoryName ("$($item.BaseName)($i)$($item.Extension)")
        $i++
    } while (Test-Path $newPath)
    return $newPath
}

function Get-ArchTag {
    # Returns the architecture string used in package filenames
    return $env:PROCESSOR_ARCHITECTURE.Replace('AMD', 'X').Replace('IA', 'X')
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Source 1 â€” store.rg-adguard.net  (HTML scrape)
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function Get-PackagesFromAdGuard {
    param([string]$StoreUrl)

    $ts = Get-Timestamp
    Write-Host "[$ts] ðŸ” Querying store.rg-adguard.net â€¦" -ForegroundColor DarkCyan

    $body        = "type=url&url=$StoreUrl&ring=Retail"
    $contentType = 'application/x-www-form-urlencoded'
    $apiUri      = 'https://store.rg-adguard.net/api/GetFiles'

    $response = Invoke-WebRequest -UseBasicParsing -Method POST `
                    -Uri $apiUri -Body $body -ContentType $contentType

    $arch        = Get-ArchTag
    $linksMatch  = $response.Links |
                   Where-Object { $_ -like '*.appx*' -or $_ -like '*.appxbundle*' -or
                                  $_ -like '*.msix*' -or $_ -like '*.msixbundle*' } |
                   Where-Object { $_ -like '*_neutral_*' -or $_ -like "*_${arch}_*" } |
                   Select-String -Pattern '(?<=a href=").+(?=" r)'

    return $linksMatch | ForEach-Object { $_.Matches.Value }
}

function Install-ViaAdGuard {
    param(
        [string]$StoreUrl,
        [string]$DownloadPath
    )

    $downloadUrls = Get-PackagesFromAdGuard -StoreUrl $StoreUrl

    if (-not $downloadUrls -or @($downloadUrls).Count -eq 0) {
        Write-Host "[$( Get-Timestamp )] âš ï¸  AdGuard returned no packages." -ForegroundColor Yellow
        return $false
    }

    $installed = 0

    foreach ($url in $downloadUrls) {
        try {
            $ts = Get-Timestamp
            Write-Host "[$ts] â¬‡  Fetching package info â€¦" -ForegroundColor DarkCyan

            $req      = Invoke-WebRequest -Uri $url -UseBasicParsing
            $fileName = ($req.Headers['Content-Disposition'] |
                         Select-String -Pattern '(?<=filename=).+').Matches.Value

            if (-not $fileName) { $fileName = Split-Path $url -Leaf }

            $filePath = Join-Path $DownloadPath $fileName
            $filePath = Resolve-UniqueFilePath $filePath

            Write-Host "[$ts]    Saving â†’ $filePath" -ForegroundColor DarkCyan
            [System.IO.File]::WriteAllBytes($filePath, $req.Content)

            $ts = Get-Timestamp
            Write-Host "[$ts]    Installing $fileName â€¦" -ForegroundColor DarkCyan
            Add-AppxPackage -Path $filePath -ErrorAction Stop

            Write-Host "[$ts] âœ… $fileName installed." -ForegroundColor Green
            $installed++

            if (-not $KeepPackage) { Remove-Item $filePath -Force -ErrorAction SilentlyContinue }
        }
        catch {
            Write-Host "[$( Get-Timestamp )] âŒ Failed on $url â€” $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    return ($installed -gt 0)
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Source 2 â€” msft-store.tplant.com.au  (REST API, fallback)
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function Install-ViaTplant {
    param(
        [string]$ProductId,
        [string]$Label,
        [string]$DownloadPath
    )

    $ts       = Get-Timestamp
    $storeUrl = "https://apps.microsoft.com/detail/$ProductId"
    $apiUrl   = "https://msft-store.tplant.com.au/api/Packages?id=$storeUrl&environment=Production&inputform=url"

    Write-Host "[$ts] ðŸ”„ Fallback: querying tplant API â€¦" -ForegroundColor Yellow

    $packages = Invoke-RestMethod -Uri $apiUrl -Method Get -ErrorAction Stop

    if (-not $packages -or @($packages).Count -eq 0) {
        Write-Host "[$ts] âš ï¸  No packages found via fallback for '$Label' ($ProductId)." -ForegroundColor Yellow
        return $false
    }

    # Prefer x64, then neutral/arm64
    $arch    = Get-ArchTag
    $package = $packages | Where-Object { $_.packagefilename -like "*x64*" } | Select-Object -First 1
    if (-not $package) {
        $package = $packages | Where-Object { $_.packagefilename -like "*$arch*" } | Select-Object -First 1
    }
    if (-not $package) {
        $package = $packages | Select-Object -First 1
    }

    $downloadUrl = $package.packagedownloadurl
    $fileName    = $package.packagefilename
    $outPath     = Join-Path $DownloadPath $fileName
    $outPath     = Resolve-UniqueFilePath $outPath

    $ts = Get-Timestamp
    Write-Host "[$ts]    Downloading $fileName â€¦" -ForegroundColor DarkCyan
    Invoke-WebRequest -Uri $downloadUrl -OutFile $outPath -UseBasicParsing -ErrorAction Stop

    $ts = Get-Timestamp
    Write-Host "[$ts]    Installing package â€¦" -ForegroundColor DarkCyan
    Add-AppxPackage -Path $outPath -ErrorAction Stop

    if (-not $KeepPackage) { Remove-Item $outPath -Force -ErrorAction SilentlyContinue }

    Write-Host "[$ts] âœ… $Label installed successfully." -ForegroundColor Green
    return $true
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Main entry â€” resolve inputs and run
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function Invoke-AppxInstall {
    param(
        [string]$StoreUrl,
        [string]$ProductId,
        [string]$DownloadPath
    )

    # â”€â”€ Normalise inputs â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    if ($StoreUrl -and -not $ProductId) {
        # Extract product ID from URL if possible
        if ($StoreUrl -match '(?i)/detail/([A-Z0-9]{12,})') {
            $ProductId = $Matches[1]
        }
    }
    elseif ($ProductId -and -not $StoreUrl) {
        $StoreUrl = "https://apps.microsoft.com/detail/$ProductId"
    }

    $label = if ($ProductId) { $ProductId } else { $StoreUrl }

    # â”€â”€ Ensure download path exists â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    if (-not (Test-Path $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    }
    $DownloadPath = (Resolve-Path $DownloadPath).Path

    # â”€â”€ Try primary source (AdGuard) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    $success = $false
    try {
        $success = Install-ViaAdGuard -StoreUrl $StoreUrl -DownloadPath $DownloadPath
    }
    catch {
        Write-Host "[$( Get-Timestamp )] âš ï¸  Primary source error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # â”€â”€ Fallback to tplant if primary failed â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    if (-not $success) {
        if (-not $ProductId) {
            Write-Host "[$( Get-Timestamp )] âŒ Fallback requires a Product ID â€” cannot extract from URL." -ForegroundColor Red
            return
        }
        try {
            $success = Install-ViaTplant -ProductId $ProductId -Label $label -DownloadPath $DownloadPath
        }
        catch {
            Write-Host "[$( Get-Timestamp )] âŒ Fallback source error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if (-not $success) {
        Write-Host "[$( Get-Timestamp )] âŒ Installation failed from both sources." -ForegroundColor Red
    }
}

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Script entry point
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

Write-Header

# When piped through  irm â€¦ | iex  no parameters are passed, so prompt the user
if (-not $StoreUrl -and -not $ProductId) {
    Write-Host "  Enter a Microsoft Store URL or Product ID." -ForegroundColor White
    Write-Host "  Examples:" -ForegroundColor DarkGray
    Write-Host "    https://apps.microsoft.com/detail/9WZDNCRFJ3TJ" -ForegroundColor DarkGray
    Write-Host "    9WZDNCRFJ3TJ" -ForegroundColor DarkGray
    Write-Host ""
    $input = Read-Host "  Store URL or Product ID"
    $input = $input.Trim()

    if ($input -match '^https?://') {
        $StoreUrl = $input
    }
    else {
        $ProductId = $input
        $StoreUrl  = "https://apps.microsoft.com/detail/$ProductId"
    }
}

Invoke-AppxInstall -StoreUrl $StoreUrl -ProductId $ProductId -DownloadPath $DownloadPath


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

# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────

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

# ─────────────────────────────────────────────────────────────────────────────
#  Source 1 — store.rg-adguard.net  (HTML scrape)
# ─────────────────────────────────────────────────────────────────────────────

function Get-PackagesFromAdGuard {
    param([string]$StoreUrl)

    $ts = Get-Timestamp
    Write-Host "[$ts] 🔍 Querying store.rg-adguard.net …" -ForegroundColor DarkCyan

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
        Write-Host "[$( Get-Timestamp )] ⚠️  AdGuard returned no packages." -ForegroundColor Yellow
        return $false
    }

    $installed = 0

    foreach ($url in $downloadUrls) {
        try {
            $ts = Get-Timestamp
            Write-Host "[$ts] ⬇  Fetching package info …" -ForegroundColor DarkCyan

            $req      = Invoke-WebRequest -Uri $url -UseBasicParsing
            $fileName = ($req.Headers['Content-Disposition'] |
                         Select-String -Pattern '(?<=filename=).+').Matches.Value

            if (-not $fileName) { $fileName = Split-Path $url -Leaf }

            $filePath = Join-Path $DownloadPath $fileName
            $filePath = Resolve-UniqueFilePath $filePath

            Write-Host "[$ts]    Saving → $filePath" -ForegroundColor DarkCyan
            [System.IO.File]::WriteAllBytes($filePath, $req.Content)

            $ts = Get-Timestamp
            Write-Host "[$ts]    Installing $fileName …" -ForegroundColor DarkCyan
            Add-AppxPackage -Path $filePath -ErrorAction Stop

            Write-Host "[$ts] ✅ $fileName installed." -ForegroundColor Green
            $installed++

            if (-not $KeepPackage) { Remove-Item $filePath -Force -ErrorAction SilentlyContinue }
        }
        catch {
            Write-Host "[$( Get-Timestamp )] ❌ Failed on $url — $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    return ($installed -gt 0)
}

# ─────────────────────────────────────────────────────────────────────────────
#  Source 2 — msft-store.tplant.com.au  (REST API, fallback)
# ─────────────────────────────────────────────────────────────────────────────

function Install-ViaTplant {
    param(
        [string]$ProductId,
        [string]$Label,
        [string]$DownloadPath
    )

    $ts       = Get-Timestamp
    $storeUrl = "https://apps.microsoft.com/detail/$ProductId"
    $apiUrl   = "https://msft-store.tplant.com.au/api/Packages?id=$storeUrl&environment=Production&inputform=url"

    Write-Host "[$ts] 🔄 Fallback: querying tplant API …" -ForegroundColor Yellow

    $packages = Invoke-RestMethod -Uri $apiUrl -Method Get -ErrorAction Stop

    if (-not $packages -or @($packages).Count -eq 0) {
        Write-Host "[$ts] ⚠️  No packages found via fallback for '$Label' ($ProductId)." -ForegroundColor Yellow
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
    Write-Host "[$ts]    Downloading $fileName …" -ForegroundColor DarkCyan
    Invoke-WebRequest -Uri $downloadUrl -OutFile $outPath -UseBasicParsing -ErrorAction Stop

    $ts = Get-Timestamp
    Write-Host "[$ts]    Installing package …" -ForegroundColor DarkCyan
    Add-AppxPackage -Path $outPath -ErrorAction Stop

    if (-not $KeepPackage) { Remove-Item $outPath -Force -ErrorAction SilentlyContinue }

    Write-Host "[$ts] ✅ $Label installed successfully." -ForegroundColor Green
    return $true
}

# ─────────────────────────────────────────────────────────────────────────────
#  Main entry — resolve inputs and run
# ─────────────────────────────────────────────────────────────────────────────

function Invoke-AppxInstall {
    param(
        [string]$StoreUrl,
        [string]$ProductId,
        [string]$DownloadPath
    )

    # ── Normalise inputs ──────────────────────────────────────────────────────
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

    # ── Ensure download path exists ───────────────────────────────────────────
    if (-not (Test-Path $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
    }
    $DownloadPath = (Resolve-Path $DownloadPath).Path

    # ── Try primary source (AdGuard) ──────────────────────────────────────────
    $success = $false
    try {
        $success = Install-ViaAdGuard -StoreUrl $StoreUrl -DownloadPath $DownloadPath
    }
    catch {
        Write-Host "[$( Get-Timestamp )] ⚠️  Primary source error: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # ── Fallback to tplant if primary failed ──────────────────────────────────
    if (-not $success) {
        if (-not $ProductId) {
            Write-Host "[$( Get-Timestamp )] ❌ Fallback requires a Product ID — cannot extract from URL." -ForegroundColor Red
            return
        }
        try {
            $success = Install-ViaTplant -ProductId $ProductId -Label $label -DownloadPath $DownloadPath
        }
        catch {
            Write-Host "[$( Get-Timestamp )] ❌ Fallback source error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if (-not $success) {
        Write-Host "[$( Get-Timestamp )] ❌ Installation failed from both sources." -ForegroundColor Red
    }
}

# ─────────────────────────────────────────────────────────────────────────────
#  Script entry point
# ─────────────────────────────────────────────────────────────────────────────

Write-Header

# When piped through  irm … | iex  no parameters are passed, so prompt the user
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
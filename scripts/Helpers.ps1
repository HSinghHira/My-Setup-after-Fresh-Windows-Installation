# Helpers.ps1 - Utility functions
# Dot-sourced into index.ps1

function Get-Timestamp {
    return (Get-Date).ToString("HH:mm:ss")
}

function Get-ArchTag {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { return 'arm64' }
    return 'x64'
}

function Resolve-UniqueFilePath {
    param( [string]$Path )
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext  = [System.IO.Path]::GetExtension($Path)
    $dir  = [System.IO.Path]::GetDirectoryName($Path)
    $finalPath = $Path
    $i = 1
    while (Test-Path $finalPath) {
        $finalPath = Join-Path $dir ("$base ($i)$ext")
        $i++
    }
    return $finalPath
}

function Add-Result {
    param( [string]$App, [string]$Status )
    $script:results.Add([PSCustomObject]@{ App = $App; Status = $Status })
}

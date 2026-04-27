# Privacy.ps1 - Section 13 logic
# Dot-sourced into index.ps1

Write-Host ""
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "  EU Privacy Unlock" -ForegroundColor White
Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray

if ($SkipEUPrivacy) {
    Write-Host "[$( Get-Timestamp )] - Skipping EU privacy unlock (flag set)." -ForegroundColor DarkGray
} else {
    $origGeo = (Get-WinHomeLocation).GeoId
    $ts      = Get-Timestamp

    Write-Host "[$ts] # Temporarily switching region to Ireland (EU) ..." -ForegroundColor DarkCyan
    try {
        Set-WinHomeLocation -GeoId 94 -ErrorAction Stop
        Write-Host "[$( Get-Timestamp )]    OK Region set to Ireland." -ForegroundColor DarkGreen
    } catch {
        Write-Host "[$( Get-Timestamp )]    ! Could not change region: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    Write-Host "[$( Get-Timestamp )]    Deleting DeviceRegion registry values ..." -ForegroundColor DarkCyan
    try {
        $key = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey(
            'SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\DeviceRegion', $true
        )
        if ($key) {
            $valueNames = $key.GetValueNames()
            foreach ($val in $valueNames) { try { $key.DeleteValue($val) } catch {} }
            $remaining = $key.GetValueNames().Count
            $key.Close()
            if ($remaining -eq 0) {
                Write-Host "[$( Get-Timestamp )]    OK DeviceRegion values cleared." -ForegroundColor Green
                Add-Result -App 'EU Privacy Unlock' -Status 'Installed'
            } else {
                Add-Result -App 'EU Privacy Unlock' -Status 'Failed'
            }
        } else {
            Add-Result -App 'EU Privacy Unlock' -Status 'Skipped'
        }
    } catch {
        Add-Result -App 'EU Privacy Unlock' -Status 'Failed'
    }

    Write-Host "[$( Get-Timestamp )]    Restoring original region ..." -ForegroundColor DarkCyan
    try {
        Set-WinHomeLocation -GeoId $origGeo -ErrorAction Stop
        Write-Host "[$( Get-Timestamp )]    OK Region restored." -ForegroundColor DarkGreen
    } catch {
        Write-Host "[$( Get-Timestamp )]    ! Could not restore region." -ForegroundColor Yellow
    }
}

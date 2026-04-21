# Ensure latest sources
winget source update

$flags = @(
    "--silent",
    "--accept-package-agreements",
    "--accept-source-agreements"
)

# ========================
# Core Apps
# ========================
winget install -e --id 7zip.7zip $flags
winget install -e --id Daum.PotPlayer $flags
winget install -e --id ShareX.ShareX $flags
winget install -e --id Notepad++.Notepad++ $flags
winget install -e --id Git.Git $flags
winget install -e --id Gyan.FFmpeg $flags

# ========================
# System Utilities
# ========================
winget install -e --id xanderfrangos.twinkletray $flags
winget install -e --id File-New-Project.EarTrumpet $flags
winget install -e --id CrystalRich.LockHunter $flags

# ========================
# Productivity
# ========================
winget install -e --id PDFgear.PDFgear $flags
winget install -e --id JavadMotallebi.NeatDownloadManager $flags
winget install -e --id flux.flux $flags
winget install -e --id riyasy.FlyPhotos $flags
winget install -e --id Google.Antigravity $flags

# ========================
# Dev Setup
# ========================
winget install -e --id Git.Git $flags
winget install -e --id GitHub.GitHubDesktop $flags
winget install -e --id Oven-sh.Bun $flags
winget install -e --id Volta.Volta $flags

# Install Node via Volta
$env:PATH += ";$env:LOCALAPPDATA\Volta\bin"
volta install node

Write-Host "✅ Setup Complete!"
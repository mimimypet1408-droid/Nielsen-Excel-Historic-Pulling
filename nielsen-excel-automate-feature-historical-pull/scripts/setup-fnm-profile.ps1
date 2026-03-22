# Adds fnm environment setup to the user's PowerShell profile, if not already
# present. Creates the profile file if it doesn't exist.

$fnmLine = 'fnm env --use-on-cd --shell powershell | Out-String | Invoke-Expression'

if (-not (Test-Path $PROFILE)) {
    New-Item -Path $PROFILE -ItemType File -Force | Out-Null
    Write-Host "Created PowerShell profile at $PROFILE"
}

$content = Get-Content $PROFILE -Raw -ErrorAction SilentlyContinue
if ($content -and $content.Contains('fnm env')) {
    Write-Host "fnm environment setup already present in $PROFILE"
    exit 0
}

Add-Content -Path $PROFILE -Value "`n# fnm (Fast Node Manager) environment setup`n$fnmLine"
Write-Host "Added fnm environment setup to $PROFILE"

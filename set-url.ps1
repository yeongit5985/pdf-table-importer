# Points every URL in manifest.xml at your GitHub Pages site.
# Usage:  .\set-url.ps1 -GitHubUser yourname [-Repo pdf-table-importer]
param(
    [Parameter(Mandatory = $true)][string]$GitHubUser,
    [string]$Repo = "pdf-table-importer"
)
$path = Join-Path $PSScriptRoot "manifest.xml"
$base = "https://$($GitHubUser.ToLower()).github.io/$Repo"
$xml = Get-Content $path -Raw
$xml = $xml -replace 'https://[^/"<]+\.github\.io/[^/"<]+', $base
[System.IO.File]::WriteAllText($path, $xml, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "manifest.xml now points to $base/"

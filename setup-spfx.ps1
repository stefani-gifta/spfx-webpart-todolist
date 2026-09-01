<#
SPFx + Heft setup helper

Usage: run from the folder where you want to scaffold the SPFx project.
Open PowerShell as Administrator if you need to install global npm packages.
#>

Write-Host "SPFx + Heft setup helper"

$node = & node -v 2>$null
if (-not $node) {
    Write-Host "Node.js is not found in PATH. Install Node.js LTS first: https://nodejs.org/" -ForegroundColor Red
    exit 1
}

Write-Host "Node version: $node"

Write-Host "Installing Yeoman and the SharePoint generator globally (may require admin rights)..."
npm install -g yo @microsoft/generator-sharepoint

if ($LASTEXITCODE -ne 0) {
    Write-Host "Global install failed. Re-run PowerShell as Admin or install packages manually." -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "Global packages installed. To scaffold a new SPFx project, run: yo @microsoft/sharepoint"

$run = Read-Host "Run 'yo @microsoft/sharepoint' now in the current folder? (Y/N)"
if ($run -match '^[Yy]') {
    yo @microsoft/sharepoint
} else {
    Write-Host "Skipped scaffolding. Run 'yo @microsoft/sharepoint' when ready." -ForegroundColor Yellow
}

# PowerShell script to reinstall Jupiter Office Add-in
# This script helps force refresh the add-in manifest changes

Write-Host "🔄 Jupiter Add-in Reinstallation Script" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan

# Step 1: Close Word if running
Write-Host "1. Checking for running Word processes..." -ForegroundColor Yellow
$wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
if ($wordProcesses) {
    Write-Host "   Found Word processes. Please close Word manually and press Enter to continue..." -ForegroundColor Red
    Read-Host
} else {
    Write-Host "   ✅ No Word processes found." -ForegroundColor Green
}

# Step 2: Clear Office Add-in cache
Write-Host "2. Clearing Office Add-in cache..." -ForegroundColor Yellow
$cacheLocations = @(
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
    "$env:LOCALAPPDATA\Microsoft\Office\Wef",
    "$env:APPDATA\Microsoft\Office\16.0\Wef",
    "$env:APPDATA\Microsoft\Office\Wef"
)

foreach ($location in $cacheLocations) {
    if (Test-Path $location) {
        try {
            Remove-Item -Path "$location\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "   ✅ Cleared cache: $location" -ForegroundColor Green
        } catch {
            Write-Host "   ⚠️  Could not clear: $location" -ForegroundColor Yellow
        }
    }
}

# Step 3: Instructions for manual reinstallation
Write-Host "3. Manual Reinstallation Steps:" -ForegroundColor Yellow
Write-Host "   a) Open Visual Studio" -ForegroundColor White
Write-Host "   b) Open the JuptiarAddins.sln solution" -ForegroundColor White
Write-Host "   c) Right-click on 'JuptiarAddins' project" -ForegroundColor White
Write-Host "   d) Select 'Deploy' or 'Rebuild'" -ForegroundColor White
Write-Host "   e) Start Word and check the Jupiter tab" -ForegroundColor White

# Step 4: Alternative - Sideload manifest directly
Write-Host "4. Alternative - Direct Manifest Sideloading:" -ForegroundColor Yellow
$manifestPath = "d:\jupiterAddIn\JuptiarAddins\JuptiarAddinsManifest\JuptiarAddins.xml"
if (Test-Path $manifestPath) {
    Write-Host "   Manifest found at: $manifestPath" -ForegroundColor Green
    Write-Host "   You can sideload this directly in Word:" -ForegroundColor White
    Write-Host "   - File > Options > Add-ins > Manage: COM Add-ins > Go" -ForegroundColor White
    Write-Host "   - Or use Office Dev Center for sideloading" -ForegroundColor White
} else {
    Write-Host "   ❌ Manifest not found at expected location" -ForegroundColor Red
}

# Step 5: Verification
Write-Host "5. Verification Steps:" -ForegroundColor Yellow
Write-Host "   After reinstalling, check for:" -ForegroundColor White
Write-Host "   ✅ Jupiter tab appears in Word ribbon" -ForegroundColor Green
Write-Host "   ✅ Three separate groups: Documents, Properties, Settings" -ForegroundColor Green
Write-Host "   ✅ New icons: 📁 Open, 💾 Save, 📋 Properties, ⚙️ Settings" -ForegroundColor Green

Write-Host ""
Write-Host "🎯 If the ribbon still shows old layout:" -ForegroundColor Cyan
Write-Host "   - The manifest cache may still be active" -ForegroundColor White
Write-Host "   - Try restarting Windows (nuclear option)" -ForegroundColor White
Write-Host "   - Or use Office Developer Tools for force refresh" -ForegroundColor White

Write-Host ""
Write-Host "✅ Script completed. Follow the manual steps above." -ForegroundColor Green

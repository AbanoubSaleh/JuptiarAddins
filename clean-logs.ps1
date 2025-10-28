# PowerShell script to remove console.log statements from JavaScript files
# This will clean up debug logs while preserving console.error statements

$jsFiles = @(
    "JuptiarAddinsWeb/Scripts/AuthManager.js",
    "JuptiarAddinsWeb/Scripts/SaveDialog.js",
    "JuptiarAddinsWeb/Scripts/DocumentUploader.js",
    "JuptiarAddinsWeb/Scripts/JupiterService.js",
    "JuptiarAddinsWeb/Scripts/RibbonManager.js",
    "JuptiarAddinsWeb/Scripts/DocumentStateManager.js",
    "JuptiarAddinsWeb/Scripts/PropertiesEditor.js",
    "JuptiarAddinsWeb/Scripts/CheckInDialog.js",
    "JuptiarAddinsWeb/Scripts/DocumentBrowser.js"
)

foreach ($file in $jsFiles) {
    if (Test-Path $file) {
        Write-Host "Cleaning logs from: $file"
        
        # Read the file content
        $content = Get-Content $file -Raw
        
        # Remove console.log statements (but keep console.error, console.warn)
        # This regex matches console.log(...) including multiline statements
        $content = $content -replace "(?m)^\s*console\.log\([^;]*\);\s*$", ""
        
        # Remove empty lines that might be left behind
        $content = $content -replace "(?m)^\s*$\r?\n", ""
        
        # Write back to file
        Set-Content $file -Value $content -NoNewline
        
        Write-Host "✅ Cleaned: $file"
    } else {
        Write-Host "❌ File not found: $file"
    }
}

Write-Host "🎉 Log cleanup complete!"

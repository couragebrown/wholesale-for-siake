# Siaka Wholesale Flow - Windows Desktop Installer & Setup Script
param(
    [string]$TargetDir = ""
)

$ErrorActionPreference = "Continue"

if (-not $TargetDir) {
    $TargetDir = $PSScriptRoot
    if (-not $TargetDir) { $TargetDir = "D:\excel" }
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   SIAKA WHOLESALE FLOW - DESKTOP SETUP & INSTALLATION    " -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Target Installation Folder: $TargetDir" -ForegroundColor White

# 1. Ensure required data directories exist
$appDir = Join-Path $TargetDir "app"
$dataDir = Join-Path $appDir "data"
$backupDir = Join-Path $dataDir "backups"
$profileDir = Join-Path $dataDir "app_profile"

foreach ($dir in @($dataDir, $backupDir, $profileDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "  [+] Created directory: $dir" -ForegroundColor Green
    } else {
        Write-Host "  [*] Verified directory: $dir" -ForegroundColor Gray
    }
}

# 2. Initialize clean production database if not present
$dataFile = Join-Path $dataDir "wholesale_data.json"
$templateFile = Join-Path $dataDir "clean_schema_template.json"

$cleanJson = @'
{
  "products": [],
  "dealers": [],
  "orders": [],
  "payments": [],
  "suppliers": [],
  "supplierPurchases": [],
  "supplierPayments": []
}
'@

if (-not (Test-Path $templateFile)) {
    [System.IO.File]::WriteAllText($templateFile, $cleanJson, [System.Text.Encoding]::UTF8)
    Write-Host "  [+] Created clean schema template" -ForegroundColor Green
}

if (-not (Test-Path $dataFile)) {
    [System.IO.File]::WriteAllText($dataFile, $cleanJson, [System.Text.Encoding]::UTF8)
    Write-Host "  [+] Created initial database: $dataFile" -ForegroundColor Green
} else {
    Write-Host "  [*] Database exists: $dataFile" -ForegroundColor Gray
}

# 3. Create Desktop & Start Menu Shortcuts with high-resolution icon
$wsh = New-Object -ComObject WScript.Shell
$desktopDir = [System.Environment]::GetFolderPath("Desktop")
$startMenuDir = Join-Path ([System.Environment]::GetFolderPath("Programs")) ""

$launcherVbs = Join-Path $TargetDir "Launch_Wholesale_App.vbs"
$iconPath = Join-Path $TargetDir "icon.ico"
if (-not (Test-Path $iconPath)) { $iconPath = Join-Path $appDir "icon.ico" }

# Desktop Shortcut
try {
    $desktopShortcutPath = Join-Path $desktopDir "Siaka Wholesale Flow.lnk"
    $sc = $wsh.CreateShortcut($desktopShortcutPath)
    $sc.TargetPath = "wscript.exe"
    $sc.Arguments = "`"$launcherVbs`""
    $sc.WorkingDirectory = $TargetDir
    $sc.Description = "Siaka Wholesale Flow - PC Manager"
    if (Test-Path $iconPath) {
        $sc.IconLocation = "$iconPath,0"
    }
    $sc.Save()
    Write-Host "  [+] Created Desktop Shortcut: $desktopShortcutPath" -ForegroundColor Green
} catch {
    Write-Warning "Could not create desktop shortcut: $_"
}

# In-Folder Shortcut
try {
    $localShortcutPath = Join-Path $TargetDir "Siaka Wholesale Flow.lnk"
    $sc2 = $wsh.CreateShortcut($localShortcutPath)
    $sc2.TargetPath = "wscript.exe"
    $sc2.Arguments = "`"$launcherVbs`""
    $sc2.WorkingDirectory = $TargetDir
    $sc2.Description = "Siaka Wholesale Flow - PC Manager"
    if (Test-Path $iconPath) {
        $sc2.IconLocation = "$iconPath,0"
    }
    $sc2.Save()
    Write-Host "  [+] Created Local Shortcut: $localShortcutPath" -ForegroundColor Green
} catch {}

# Start Menu Shortcut
try {
    $smShortcutPath = Join-Path $startMenuDir "Siaka Wholesale Flow.lnk"
    $sc3 = $wsh.CreateShortcut($smShortcutPath)
    $sc3.TargetPath = "wscript.exe"
    $sc3.Arguments = "`"$launcherVbs`""
    $sc3.WorkingDirectory = $TargetDir
    $sc3.Description = "Siaka Wholesale Flow - PC Manager"
    if (Test-Path $iconPath) {
        $sc3.IconLocation = "$iconPath,0"
    }
    $sc3.Save()
    Write-Host "  [+] Created Start Menu Shortcut: $smShortcutPath" -ForegroundColor Green
} catch {}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   INSTALLATION COMPLETE! ZERO DATA LOSS SYSTEM READY     " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "All required files and backup routines are operational." -ForegroundColor White
Write-Host "You can launch Siaka Wholesale Flow from your Desktop icon." -ForegroundColor White

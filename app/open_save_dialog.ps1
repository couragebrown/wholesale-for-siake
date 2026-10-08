# D:\excel\app\open_save_dialog.ps1
# Opens native Windows Save File dialog and saves the Excel file to the chosen location,
# then opens Windows File Explorer focusing on that file.

Add-Type -AssemblyName System.Windows.Forms

$sourceFile = "D:\excel\Wholesale_Dealer_Management_System.xlsx"
if (-not (Test-Path $sourceFile)) {
    $sourceFile = "D:\excel\app\data\wholesale_data.json"
}

$saveDlg = New-Object System.Windows.Forms.SaveFileDialog
$saveDlg.Title = "Save Siaka Wholesale Flow - Select Folder To Save Excel File"
$saveDlg.Filter = "Microsoft Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$saveDlg.FileName = "Siaka_Wholesale_Dealer_System.xlsx"
$saveDlg.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$saveDlg.RestoreDirectory = $true
$saveDlg.OverwritePrompt = $true

$form = New-Object System.Windows.Forms.Form
$form.TopMost = $true
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

$result = $saveDlg.ShowDialog($form)
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $targetPath = $saveDlg.FileName
    try {
        Copy-Item -Path $sourceFile -Destination $targetPath -Force
        Start-Process "explorer.exe" -ArgumentList "/select,`"$targetPath`""
    } catch {
        Write-Warning "Failed to copy file: $_"
    }
}

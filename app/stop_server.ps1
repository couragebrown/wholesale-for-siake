# Stop any running server instance
Get-CimInstance Win32_Process | ForEach-Object {
    if ($_.CommandLine -and $_.CommandLine -like "*server.ps1*") {
        Write-Host "Stopping Server Process ID: $($_.ProcessId)"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
}
Write-Host "Server stop check complete."

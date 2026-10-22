# Siaka Wholesale Flow - Production Backend Server
# Features: Self-Healing Storage, Atomic Writes, Automatic Rolling Backups, Zero Data Loss Protection
param(
    [int]$Port = 49221
)

$ErrorActionPreference = "Continue"

# Dynamic Directory Setup
$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $baseDir) { $baseDir = "D:\excel\app" }
$rootDir = Split-Path -Parent $baseDir
$dataDir = Join-Path $baseDir "data"
$backupDir = Join-Path $dataDir "backups"
$dataFile = Join-Path $dataDir "wholesale_data.json"
$templateFile = Join-Path $dataDir "clean_schema_template.json"
$htmlFile = Join-Path $baseDir "index.html"
$excelExportPath = Join-Path $rootDir "Wholesale_Dealer_Management_System.xlsx"

# Ensure all critical folders exist
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

$defaultSchema = @'
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

# Ensure template exists
if (-not (Test-Path $templateFile)) {
    [System.IO.File]::WriteAllText($templateFile, $defaultSchema, [System.Text.Encoding]::UTF8)
}

# Function: Auto-Recover and return valid Database JSON string
function Get-ValidDatabaseJson() {
    # 1. Check primary data file
    if (Test-Path $dataFile) {
        try {
            $raw = [System.IO.File]::ReadAllText($dataFile, [System.Text.Encoding]::UTF8)
            if ($raw -and $raw.Trim().Length -gt 10) {
                $test = $raw | ConvertFrom-Json
                if ($test) { return $raw }
            }
        } catch {
            Write-Warning "Primary data file corrupt or invalid: $_"
        }
    }

    # 2. Try recovering from newest backup in backups folder
    Write-Warning "Attempting automatic recovery from backups..."
    $backups = Get-ChildItem -Path $backupDir -Filter "backup_*.json" | Sort-Object LastWriteTime -Descending
    foreach ($bak in $backups) {
        try {
            $rawBak = [System.IO.File]::ReadAllText($bak.FullName, [System.Text.Encoding]::UTF8)
            $testBak = $rawBak | ConvertFrom-Json
            if ($testBak) {
                Write-Host "Restored database from backup: $($bak.Name)"
                [System.IO.File]::WriteAllText($dataFile, $rawBak, [System.Text.Encoding]::UTF8)
                return $rawBak
            }
        } catch {}
    }

    # 3. Check for .bak fallback
    $bakFile = Join-Path $dataDir "wholesale_data.bak"
    if (Test-Path $bakFile) {
        try {
            $rawBak = [System.IO.File]::ReadAllText($bakFile, [System.Text.Encoding]::UTF8)
            $testBak = $rawBak | ConvertFrom-Json
            if ($testBak) {
                Write-Host "Restored database from wholesale_data.bak"
                [System.IO.File]::WriteAllText($dataFile, $rawBak, [System.Text.Encoding]::UTF8)
                return $rawBak
            }
        } catch {}
    }

    # 4. Fallback to clean schema template
    Write-Host "Initializing new clean database from template."
    [System.IO.File]::WriteAllText($dataFile, $defaultSchema, [System.Text.Encoding]::UTF8)
    return $defaultSchema
}

# Function: Safe Atomic Save with Automatic Versioned Rolling Backup
function Save-DatabaseSafely([string]$jsonContent) {
    if (-not $jsonContent -or $jsonContent.Trim().Length -lt 10) {
        throw "Save rejected: Empty or invalid JSON payload."
    }

    # Validate JSON integrity before touching disk
    $parsed = $jsonContent | ConvertFrom-Json
    if (-not $parsed) {
        throw "Save rejected: Payload is not valid JSON."
    }

    # Ensure data & backup directories exist
    if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }
    if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

    # Step A: Create versioned timestamped backup of current existing data
    if (Test-Path $dataFile) {
        try {
            $currentData = [System.IO.File]::ReadAllText($dataFile, [System.Text.Encoding]::UTF8)
            if ($currentData.Trim().Length -gt 10) {
                $timeTag = Get-Date -Format "yyyyMMdd_HHmmss"
                $snapshotPath = Join-Path $backupDir ("backup_" + $timeTag + ".json")
                [System.IO.File]::WriteAllText($snapshotPath, $currentData, [System.Text.Encoding]::UTF8)
                
                # Also maintain immediate single-file rollback .bak
                $bakPath = Join-Path $dataDir "wholesale_data.bak"
                [System.IO.File]::WriteAllText($bakPath, $currentData, [System.Text.Encoding]::UTF8)
            }
        } catch {
            Write-Warning "Could not create backup snapshot: $_"
        }
    }

    # Step B: Atomic write to temporary file first
    $tempFile = "$dataFile.tmp"
    [System.IO.File]::WriteAllText($tempFile, $jsonContent, [System.Text.Encoding]::UTF8)

    # Step C: Replace primary file
    if (Test-Path $tempFile) {
        Move-Item -Path $tempFile -Destination $dataFile -Force
    }

    # Step D: Prune older backups keeping last 100 snapshots
    try {
        $oldBackups = Get-ChildItem -Path $backupDir -Filter "backup_*.json" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 100
        if ($oldBackups) {
            $oldBackups | ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue }
        }
    } catch {}

    return $true
}

# Pre-initialize database file if not present
$initialCheck = Get-ValidDatabaseJson

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$Port/")

try {
    $listener.Start()
    Write-Host "Siaka Wholesale Flow Server running on http://127.0.0.1:$Port/"
    Write-Host "Data storage: $dataFile"
    Write-Host "Automated backups: $backupDir"
} catch {
    Write-Error "Failed to start listener on port ${Port}: $_"
    exit 1
}

function Send-Response($context, [string]$content, [string]$contentType = "text/html; charset=utf-8", [int]$statusCode = 200) {
    try {
        $response = $context.Response
        $response.StatusCode = $statusCode
        $response.ContentType = $contentType
        $response.Headers.Add("Access-Control-Allow-Origin", "*")
        $response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, OPTIONS, HEAD")
        $response.Headers.Add("Access-Control-Allow-Headers", "Content-Type")
        
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($content)
        $response.ContentLength64 = $buffer.Length
        if ($context.Request.HttpMethod -ne "HEAD") {
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
        }
        $response.OutputStream.Close()
    } catch {
        try { $context.Response.Abort() } catch {}
    }
}

function Export-ToExcel($jsonData) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $wb = $excel.Workbooks.Add()
        while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item(2).Delete() }
        
        $wsDash = $wb.Sheets.Item(1)
        $wsDash.Name = "Dashboard"
        $wsOrders = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsDash)
        $wsOrders.Name = "Orders Ledger"
        $wsDealers = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsOrders)
        $wsDealers.Name = "Dealers"
        $wsProducts = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsDealers)
        $wsProducts.Name = "Products"

        # 1. Products
        $wsProducts.Cells.Item(1, 1).Value2 = "SKU"
        $wsProducts.Cells.Item(1, 2).Value2 = "Category"
        $wsProducts.Cells.Item(1, 3).Value2 = "Product Name"
        $wsProducts.Cells.Item(1, 4).Value2 = "Unit"
        $wsProducts.Cells.Item(1, 5).Value2 = "Wholesale Price (GH₵)"
        $wsProducts.Cells.Item(1, 6).Value2 = "Current Stock"
        $wsProducts.Range("A1:F1").Font.Bold = $true
        $wsProducts.Range("A1:F1").Interior.Color = 10114859
        $wsProducts.Range("A1:F1").Font.Color = 16777215

        $r = 2
        if ($jsonData.products) {
            foreach ($p in $jsonData.products) {
                $wsProducts.Cells.Item($r, 1).Value2 = [string]$p.id
                $wsProducts.Cells.Item($r, 2).Value2 = [string]$p.category
                $wsProducts.Cells.Item($r, 3).Value2 = [string]$p.name
                $wsProducts.Cells.Item($r, 4).Value2 = [string]$p.unit
                $wsProducts.Cells.Item($r, 5).Formula = "$($p.price)"
                $wsProducts.Cells.Item($r, 6).Formula = "$($p.stock)"
                $r++
            }
            if ($r -gt 2) {
                $wsProducts.Range("E2:E$($r-1)").NumberFormat = """GH₵ ""#,##0.00"
            }
        }
        $wsProducts.Columns.AutoFit() | Out-Null

        # 2. Dealers
        $wsDealers.Cells.Item(1, 1).Value2 = "Dealer ID"
        $wsDealers.Cells.Item(1, 2).Value2 = "Company Name"
        $wsDealers.Cells.Item(1, 3).Value2 = "Contact Person"
        $wsDealers.Cells.Item(1, 4).Value2 = "Phone"
        $wsDealers.Cells.Item(1, 5).Value2 = "Location"
        $wsDealers.Cells.Item(1, 6).Value2 = "Terms"
        $wsDealers.Range("A1:F1").Font.Bold = $true
        $wsDealers.Range("A1:F1").Interior.Color = 5587251
        $wsDealers.Range("A1:F1").Font.Color = 16777215

        $r = 2
        if ($jsonData.dealers) {
            foreach ($d in $jsonData.dealers) {
                $wsDealers.Cells.Item($r, 1).Value2 = [string]$d.id
                $wsDealers.Cells.Item($r, 2).Value2 = [string]$d.name
                $wsDealers.Cells.Item($r, 3).Value2 = [string]$d.contact
                $wsDealers.Cells.Item($r, 4).Value2 = [string]$d.phone
                $wsDealers.Cells.Item($r, 5).Value2 = [string]$d.location
                $wsDealers.Cells.Item($r, 6).Value2 = [string]$d.terms
                $r++
            }
        }
        $wsDealers.Columns.AutoFit() | Out-Null

        # 3. Orders Ledger
        $headers = @("Order ID", "Date", "Invoice", "Dealer", "Product", "Unit", "Unit Price (GH₵)", "Qty", "Total Due (GH₵)", "Amount Paid (GH₵)", "Balance (GH₵)", "Status")
        for ($c = 0; $c -lt $headers.Length; $c++) {
            $wsOrders.Cells.Item(1, $c + 1).Value2 = $headers[$c]
        }
        $wsOrders.Range("A1:L1").Font.Bold = $true
        $wsOrders.Range("A1:L1").Interior.Color = 6116891
        $wsOrders.Range("A1:L1").Font.Color = 16777215

        $r = 2
        $hasOrders = ($jsonData.orders -and $jsonData.orders.Count -gt 0)
        if ($hasOrders) {
            foreach ($o in $jsonData.orders) {
                $wsOrders.Cells.Item($r, 1).Value2 = [string]$o.id
                $wsOrders.Cells.Item($r, 2).Value2 = [string]$o.date
                $wsOrders.Cells.Item($r, 3).Value2 = [string]$o.invoice
                $wsOrders.Cells.Item($r, 4).Value2 = [string]$o.dealerName
                $wsOrders.Cells.Item($r, 5).Value2 = [string]$o.itemName
                $wsOrders.Cells.Item($r, 6).Value2 = [string]$o.unit
                $wsOrders.Cells.Item($r, 7).Formula = "$($o.unitPrice)"
                $wsOrders.Cells.Item($r, 8).Formula = "$($o.quantity)"
                $wsOrders.Cells.Item($r, 9).Formula = "$($o.totalDue)"
                $wsOrders.Cells.Item($r, 10).Formula = "$($o.amountPaid)"
                $wsOrders.Cells.Item($r, 11).Formula = "$($o.balance)"
                $wsOrders.Cells.Item($r, 12).Value2 = [string]$o.status
                $r++
            }
            $wsOrders.Range("G2:G$($r-1)").NumberFormat = """GH₵ ""#,##0.00"
            $wsOrders.Range("I2:K$($r-1)").NumberFormat = """GH₵ ""#,##0.00"
        }
        $wsOrders.Columns.AutoFit() | Out-Null

        # 4. Payment History
        $wsPayments = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsOrders)
        $wsPayments.Name = "Payment History"
        $payHeaders = @("Payment ID", "Date", "Dealer Name", "Amount Paid (GH₵)", "Method", "Reference", "Notes")
        for ($c = 0; $c -lt $payHeaders.Length; $c++) {
            $wsPayments.Cells.Item(1, $c + 1).Value2 = $payHeaders[$c]
        }
        $wsPayments.Range("A1:G1").Font.Bold = $true
        $wsPayments.Range("A1:G1").Interior.Color = 10114859
        $wsPayments.Range("A1:G1").Font.Color = 16777215

        $pr = 2
        if ($jsonData.payments) {
            foreach ($py in $jsonData.payments) {
                $wsPayments.Cells.Item($pr, 1).Value2 = [string]$py.id
                $wsPayments.Cells.Item($pr, 2).Value2 = [string]$py.date
                $wsPayments.Cells.Item($pr, 3).Value2 = [string]$py.dealerName
                $wsPayments.Cells.Item($pr, 4).Formula = "$($py.amount)"
                $wsPayments.Cells.Item($pr, 5).Value2 = [string]$py.method
                $wsPayments.Cells.Item($pr, 6).Value2 = [string]$py.reference
                $wsPayments.Cells.Item($pr, 7).Value2 = [string]$py.note
                $pr++
            }
            if ($pr -gt 2) {
                $wsPayments.Range("D2:D$($pr-1)").NumberFormat = """GH₵ ""#,##0.00"
            }
        }
        $wsPayments.Columns.AutoFit() | Out-Null

        # 5. Suppliers Sheet
        $wsSuppliers = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsPayments)
        $wsSuppliers.Name = "Suppliers"
        $supHeaders = @("Supplier ID", "Supplier Name", "Contact Person", "Phone", "Location", "Terms", "Notes")
        for ($c = 0; $c -lt $supHeaders.Length; $c++) {
            $wsSuppliers.Cells.Item(1, $c + 1).Value2 = $supHeaders[$c]
        }
        $wsSuppliers.Range("A1:G1").Font.Bold = $true
        $wsSuppliers.Range("A1:G1").Interior.Color = 39372  # Amber/Orange
        $wsSuppliers.Range("A1:G1").Font.Color = 16777215

        $sr = 2
        if ($jsonData.suppliers) {
            foreach ($s in $jsonData.suppliers) {
                $wsSuppliers.Cells.Item($sr, 1).Value2 = [string]$s.id
                $wsSuppliers.Cells.Item($sr, 2).Value2 = [string]$s.name
                $wsSuppliers.Cells.Item($sr, 3).Value2 = [string]$s.contact
                $wsSuppliers.Cells.Item($sr, 4).Value2 = [string]$s.phone
                $wsSuppliers.Cells.Item($sr, 5).Value2 = [string]$s.location
                $wsSuppliers.Cells.Item($sr, 6).Value2 = [string]$s.terms
                $wsSuppliers.Cells.Item($sr, 7).Value2 = [string]$s.notes
                $sr++
            }
        }
        $wsSuppliers.Columns.AutoFit() | Out-Null

        # 6. Dashboard KPIs
        $wsDash.Range("B2:E2").Merge()
        $wsDash.Range("B2").Value2 = "SIAKA WHOLESALE FLOW - PERFORMANCE SUMMARY"
        $wsDash.Range("B2").Font.Size = 14
        $wsDash.Range("B2").Font.Bold = $true
        
        $wsDash.Cells.Item(4, 2).Value2 = "Total Orders Count:"
        $wsDash.Cells.Item(5, 2).Value2 = "Total Revenue Billed:"
        $wsDash.Cells.Item(6, 2).Value2 = "Total Cash Collected:"
        $wsDash.Cells.Item(7, 2).Value2 = "Outstanding Debt Remaining:"

        if ($hasOrders) {
            $lastRow = $r - 1
            $wsDash.Cells.Item(4, 3).Formula = "=COUNTA('Orders Ledger'!A2:A$lastRow)"
            $wsDash.Cells.Item(5, 3).Formula = "=SUM('Orders Ledger'!I2:I$lastRow)"
            $wsDash.Cells.Item(6, 3).Formula = "=SUM('Orders Ledger'!J2:J$lastRow)"
            $wsDash.Cells.Item(7, 3).Formula = "=SUM('Orders Ledger'!K2:K$lastRow)"
        } else {
            $wsDash.Cells.Item(4, 3).Value2 = 0
            $wsDash.Cells.Item(5, 3).Value2 = 0
            $wsDash.Cells.Item(6, 3).Value2 = 0
            $wsDash.Cells.Item(7, 3).Value2 = 0
        }

        $wsDash.Range("C5:C7").NumberFormat = """GH₵ ""#,##0.00"
        $wsDash.Range("B4:C7").Font.Bold = $true
        $wsDash.Columns.AutoFit() | Out-Null

        $wsDash.Activate()
        $wb.SaveAs($excelExportPath, 51)
        return $true
    }
    catch {
        Write-Error "Excel export error: $_"
        return $false
    }
    finally {
        if ($wb) { $wb.Close($false) }
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

try {
    while ($listener.IsListening) {
        try {
            $context = $listener.GetContext()
            $request = $context.Request
            $urlPath = $request.Url.AbsolutePath
            $method  = $request.HttpMethod

            if ($method -eq "OPTIONS") {
                Send-Response $context "" "text/plain" 200
                continue
            }

            # 1. Main Application Page
            if ($urlPath -eq "/" -or $urlPath -eq "/index.html") {
                if (Test-Path $htmlFile) {
                    $html = [System.IO.File]::ReadAllText($htmlFile, [System.Text.Encoding]::UTF8)
                    Send-Response $context $html "text/html; charset=utf-8"
                } else {
                    Send-Response $context "index.html not found" "text/plain" 404
                }
            }
            # 2. Database API: Load
            elseif ($urlPath -eq "/api/data" -and $method -eq "GET") {
                $json = Get-ValidDatabaseJson
                Send-Response $context $json "application/json; charset=utf-8"
            }
            # 3. Database API: Safe Save (Atomic + Auto-Backup)
            elseif ($urlPath -eq "/api/save" -and $method -eq "POST") {
                $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                $body = $reader.ReadToEnd()
                $reader.Close()
                
                try {
                    Save-DatabaseSafely $body | Out-Null
                    Send-Response $context '{"status":"ok","message":"Database saved securely with auto-backup"}' "application/json"
                } catch {
                    Write-Error "Save error: $_"
                    Send-Response $context "{\`"status\`":\`"error\`",\`"message\`":\`"$($_.Exception.Message)\`"}" "application/json" 500
                }
            }
            # 4. Database API: List Backups
            elseif ($urlPath -eq "/api/backups" -and $method -eq "GET") {
                $backupList = @()
                if (Test-Path $backupDir) {
                    $files = Get-ChildItem -Path $backupDir -Filter "*.json" | Sort-Object LastWriteTime -Descending
                    foreach ($f in $files) {
                        $backupList += @{
                            name = $f.Name
                            size = $f.Length
                            date = $f.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                        }
                    }
                }
                $jsonBackups = $backupList | ConvertTo-Json -Depth 2
                Send-Response $context $jsonBackups "application/json; charset=utf-8"
            }
            # 5. Database API: Manual Backup Trigger
            elseif ($urlPath -eq "/api/backup-now" -and $method -eq "POST") {
                try {
                    $currJson = Get-ValidDatabaseJson
                    $timeTag = Get-Date -Format "yyyyMMdd_HHmmss"
                    $manualName = "manual_backup_$timeTag.json"
                    $manualPath = Join-Path $backupDir $manualName
                    [System.IO.File]::WriteAllText($manualPath, $currJson, [System.Text.Encoding]::UTF8)
                    Send-Response $context "{\`"status\`":\`"ok\`",\`"filename\`":\`"$manualName\`",\`"message\`":\`"Manual backup snapshot created successfully\`"}" "application/json"
                } catch {
                    Send-Response $context "{\`"status\`":\`"error\`",\`"message\`":\`"$($_)\`"}" "application/json" 500
                }
            }
            # 6. Database API: Restore from Specific Backup File
            elseif ($urlPath -eq "/api/restore-backup" -and $method -eq "POST") {
                $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                $reqBody = $reader.ReadToEnd()
                $reader.Close()
                
                try {
                    $parsedReq = $reqBody | ConvertFrom-Json
                    $targetFilename = $parsedReq.filename
                    $targetFilePath = Join-Path $backupDir $targetFilename
                    
                    if (Test-Path $targetFilePath) {
                        $contentToRestore = [System.IO.File]::ReadAllText($targetFilePath, [System.Text.Encoding]::UTF8)
                        # Pre-restore safety backup of whatever exists right now
                        Save-DatabaseSafely $contentToRestore | Out-Null
                        Send-Response $context '{"status":"ok","message":"Database restored successfully"}' "application/json"
                    } else {
                        Send-Response $context '{"status":"error","message":"Backup file not found"}' "application/json" 404
                    }
                } catch {
                    Send-Response $context "{\`"status\`":\`"error\`",\`"message\`":\`"$($_)\`"}" "application/json" 500
                }
            }
            # 7. Print Statement Sync & View
            elseif ($urlPath -eq "/print_statement.html" -or $urlPath -eq "/statement.html") {
                $pfile = Join-Path $baseDir "print_statement.html"
                if (Test-Path $pfile) {
                    $html = [System.IO.File]::ReadAllText($pfile, [System.Text.Encoding]::UTF8)
                    Send-Response $context $html "text/html; charset=utf-8"
                } else {
                    Send-Response $context "No statement currently ready for printing" "text/plain" 404
                }
            }
            elseif ($urlPath -eq "/api/open-statement-print" -and $method -eq "POST") {
                $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                $body = $reader.ReadToEnd()
                $reader.Close()
                
                $pfile = Join-Path $baseDir "print_statement.html"
                [System.IO.File]::WriteAllText($pfile, $body, [System.Text.Encoding]::UTF8)
                Send-Response $context '{"status":"ok"}' "application/json"
            }
            # 8. Excel Management
            elseif ($urlPath -eq "/api/export-excel" -and $method -eq "POST") {
                $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                $body = $reader.ReadToEnd()
                $reader.Close()
                
                $jsonData = $body | ConvertFrom-Json
                $ok = Export-ToExcel $jsonData
                if ($ok) {
                    Send-Response $context "{\`"status\`":\`"ok\`",\`"path\`":\`"$([regex]::Escape($excelExportPath))\`"}" "application/json"
                } else {
                    Send-Response $context '{"status":"error","message":"Excel export failed"}' "application/json" 500
                }
            }
            elseif ($urlPath -eq "/api/download-excel" -and ($method -eq "GET" -or $method -eq "HEAD")) {
                if (Test-Path $excelExportPath) {
                    $bytes = [System.IO.File]::ReadAllBytes($excelExportPath)
                    $response = $context.Response
                    $response.StatusCode = 200
                    $response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    $response.Headers.Add("Content-Disposition", "attachment; filename=`"Siaka_Wholesale_Dealer_System.xlsx`"")
                    $response.Headers.Add("Access-Control-Allow-Origin", "*")
                    $response.ContentLength64 = $bytes.Length
                    if ($method -ne "HEAD") {
                        $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    }
                    $response.OutputStream.Close()
                } else {
                    Send-Response $context "Excel file not found" "text/plain" 404
                }
            }
            elseif ($urlPath -eq "/api/open-excel" -and $method -eq "POST") {
                if (Test-Path $excelExportPath) {
                    Start-Process "explorer.exe" -ArgumentList "/select,`"$excelExportPath`""
                } else {
                    Start-Process "explorer.exe" -ArgumentList "$rootDir"
                }
                Send-Response $context '{"status":"ok"}' "application/json"
            }
            elseif ($urlPath -eq "/api/open-save-dialog" -and $method -eq "POST") {
                $dlgScript = Join-Path $baseDir "open_save_dialog.ps1"
                if (Test-Path $dlgScript) {
                    Start-Process "powershell.exe" -ArgumentList "-STA -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$dlgScript`""
                } else {
                    Start-Process "explorer.exe" -ArgumentList "$rootDir"
                }
                Send-Response $context '{"status":"ok"}' "application/json"
            }
            elseif ($urlPath -eq "/api/open-folder" -and $method -eq "POST") {
                Start-Process "explorer.exe" -ArgumentList "$rootDir"
                Send-Response $context '{"status":"ok"}' "application/json"
            }
            # 9. Static Assets (Images, Icons, Favicons)
            elseif ($urlPath -in @("/logo.png", "/icon.png", "/favicon.ico", "/icon.ico") -or $urlPath.EndsWith(".png") -or $urlPath.EndsWith(".ico")) {
                $reqFile = Join-Path $baseDir ($urlPath.TrimStart("/\"))
                if (Test-Path $reqFile) {
                    $bytes = [System.IO.File]::ReadAllBytes($reqFile)
                    $ctype = if ($reqFile.EndsWith(".ico")) { "image/x-icon" } else { "image/png" }
                    $response = $context.Response
                    $response.StatusCode = 200
                    $response.ContentType = $ctype
                    $response.Headers.Add("Access-Control-Allow-Origin", "*")
                    $response.Headers.Add("Cache-Control", "public, max-age=86400")
                    $response.ContentLength64 = $bytes.Length
                    if ($method -ne "HEAD") {
                        $response.OutputStream.Write($bytes, 0, $bytes.Length)
                    }
                    $response.OutputStream.Close()
                } else {
                    Send-Response $context "File not found" "text/plain" 404
                }
            }
            else {
                Send-Response $context "Not Found" "text/plain" 404
            }
        } catch {
            Write-Warning "Request handling error: $_"
            try { $context.Response.Abort() } catch {}
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
}

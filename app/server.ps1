# Wholesale App Backend Server
param(
    [int]$Port = 49221
)

$baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { "D:\excel\app" }
$dataFile = Join-Path $baseDir "data\wholesale_data.json"
$htmlFile = Join-Path $baseDir "index.html"
$excelExportPath = "D:\excel\Wholesale_Dealer_Management_System.xlsx"

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$Port/")

try {
    $listener.Start()
    Write-Host "Wholesale App Server running on http://127.0.0.1:$Port/"
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
        Write-Warning "Send-Response warning: $_"
        try { $context.Response.Abort() } catch {}
    }
}


function Export-ToExcel($jsonData) {
    # Generate updated Excel file using native Excel COM
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
        foreach ($p in $jsonData.products) {
            $wsProducts.Cells.Item($r, 1).Value2 = [string]$p.id
            $wsProducts.Cells.Item($r, 2).Value2 = [string]$p.category
            $wsProducts.Cells.Item($r, 3).Value2 = [string]$p.name
            $wsProducts.Cells.Item($r, 4).Value2 = [string]$p.unit
            $wsProducts.Cells.Item($r, 5).Formula = "$($p.price)"
            $wsProducts.Cells.Item($r, 6).Formula = "$($p.stock)"
            $r++
        }
        $wsProducts.Range("E2:E$r").NumberFormat = """GH₵ ""#,##0.00"
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
        foreach ($d in $jsonData.dealers) {
            $wsDealers.Cells.Item($r, 1).Value2 = [string]$d.id
            $wsDealers.Cells.Item($r, 2).Value2 = [string]$d.name
            $wsDealers.Cells.Item($r, 3).Value2 = [string]$d.contact
            $wsDealers.Cells.Item($r, 4).Value2 = [string]$d.phone
            $wsDealers.Cells.Item($r, 5).Value2 = [string]$d.location
            $wsDealers.Cells.Item($r, 6).Value2 = [string]$d.terms
            $r++
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
        $wsOrders.Range("G2:G$r").NumberFormat = """GH₵ ""#,##0.00"
        $wsOrders.Range("I2:K$r").NumberFormat = """GH₵ ""#,##0.00"
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
            $wsPayments.Range("D2:D$pr").NumberFormat = """GH₵ ""#,##0.00"
        }
        $wsPayments.Columns.AutoFit() | Out-Null

        # 4.5. Suppliers
        $wsSuppliers = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsPayments)
        $wsSuppliers.Name = "Suppliers"
        $supHeaders = @("Supplier ID", "Supplier / Company", "Contact Person", "Phone", "Location", "Terms", "Notes")
        for ($c = 0; $c -lt $supHeaders.Length; $c++) {
            $wsSuppliers.Cells.Item(1, $c + 1).Value2 = $supHeaders[$c]
        }
        $wsSuppliers.Range("A1:G1").Font.Bold = $true
        $wsSuppliers.Range("A1:G1").Interior.Color = 4210752
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

        # 4.6. Supplier Purchases
        $wsPurchases = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsSuppliers)
        $wsPurchases.Name = "Supplier Purchases"
        $purHeaders = @("Purchase ID", "Date", "Supplier Name", "Product Name", "Quantity", "Unit Cost (GH₵)", "Total Due (GH₵)", "Amount Paid (GH₵)", "Balance Owed (GH₵)", "Status", "Notes")
        for ($c = 0; $c -lt $purHeaders.Length; $c++) {
            $wsPurchases.Cells.Item(1, $c + 1).Value2 = $purHeaders[$c]
        }
        $wsPurchases.Range("A1:K1").Font.Bold = $true
        $wsPurchases.Range("A1:K1").Interior.Color = 5587251
        $wsPurchases.Range("A1:K1").Font.Color = 16777215

        $purR = 2
        if ($jsonData.supplierPurchases) {
            foreach ($pur in $jsonData.supplierPurchases) {
                $wsPurchases.Cells.Item($purR, 1).Value2 = [string]$pur.id
                $wsPurchases.Cells.Item($purR, 2).Value2 = [string]$pur.date
                $wsPurchases.Cells.Item($purR, 3).Value2 = [string]$pur.supplierName
                $wsPurchases.Cells.Item($purR, 4).Value2 = [string]$pur.itemName
                $wsPurchases.Cells.Item($purR, 5).Formula = "$($pur.quantity)"
                $wsPurchases.Cells.Item($purR, 6).Formula = "$($pur.unitCost)"
                $wsPurchases.Cells.Item($purR, 7).Formula = "$($pur.totalDue)"
                $wsPurchases.Cells.Item($purR, 8).Formula = "$($pur.amountPaid)"
                $wsPurchases.Cells.Item($purR, 9).Formula = "$($pur.balance)"
                $wsPurchases.Cells.Item($purR, 10).Value2 = [string]$pur.status
                $wsPurchases.Cells.Item($purR, 11).Value2 = [string]$pur.notes
                $purR++
            }
            $wsPurchases.Range("F2:I$purR").NumberFormat = """GH₵ ""#,##0.00"
        }
        $wsPurchases.Columns.AutoFit() | Out-Null

        # 4.7. Supplier Payments
        $wsSupPayments = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsPurchases)
        $wsSupPayments.Name = "Supplier Payments"
        $spayHeaders = @("Payment ID", "Date", "Supplier Name", "Amount Paid (GH₵)", "Method", "Reference", "Notes")
        for ($c = 0; $c -lt $spayHeaders.Length; $c++) {
            $wsSupPayments.Cells.Item(1, $c + 1).Value2 = $spayHeaders[$c]
        }
        $wsSupPayments.Range("A1:G1").Font.Bold = $true
        $wsSupPayments.Range("A1:G1").Interior.Color = 10114859
        $wsSupPayments.Range("A1:G1").Font.Color = 16777215

        $spR = 2
        if ($jsonData.supplierPayments) {
            foreach ($spy in $jsonData.supplierPayments) {
                $wsSupPayments.Cells.Item($spR, 1).Value2 = [string]$spy.id
                $wsSupPayments.Cells.Item($spR, 2).Value2 = [string]$spy.date
                $wsSupPayments.Cells.Item($spR, 3).Value2 = [string]$spy.supplierName
                $wsSupPayments.Cells.Item($spR, 4).Formula = "$($spy.amount)"
                $wsSupPayments.Cells.Item($spR, 5).Value2 = [string]$spy.method
                $wsSupPayments.Cells.Item($spR, 6).Value2 = [string]$spy.reference
                $wsSupPayments.Cells.Item($spR, 7).Value2 = [string]$spy.note
                $spR++
            }
            $wsSupPayments.Range("D2:D$spR").NumberFormat = """GH₵ ""#,##0.00"
        }
        $wsSupPayments.Columns.AutoFit() | Out-Null

        # 5. Dashboard KPIs
        $wsDash.Range("B2:E2").Merge()
        $wsDash.Range("B2").Value2 = "SIAKA WHOLESALE FLOW - PERFORMANCE SUMMARY"
        $wsDash.Range("B2").Font.Size = 14
        $wsDash.Range("B2").Font.Bold = $true
        
        $wsDash.Cells.Item(4, 2).Value2 = "Total Orders Count:"
        $wsDash.Cells.Item(4, 3).Formula = "=COUNTA('Orders Ledger'!A2:A$r)"
        $wsDash.Cells.Item(5, 2).Value2 = "Total Revenue Billed:"
        $wsDash.Cells.Item(5, 3).Formula = "=SUM('Orders Ledger'!I2:I$r)"
        $wsDash.Cells.Item(6, 2).Value2 = "Total Cash Collected:"
        $wsDash.Cells.Item(6, 3).Formula = "=SUM('Orders Ledger'!J2:J$r)"
        $wsDash.Cells.Item(7, 2).Value2 = "Outstanding Debt Remaining:"
        $wsDash.Cells.Item(7, 3).Formula = "=SUM('Orders Ledger'!K2:K$r)"
        $wsDash.Range("C5:C7").NumberFormat = """GH₵ ""#,##0.00"
        $wsDash.Range("B4:C7").Font.Bold = $true
        $wsDash.Columns.AutoFit() | Out-Null

        $wsDash.Activate()
        $wb.SaveAs($excelExportPath, 51)
        Write-Host "Exported to $excelExportPath"
        return $true
    }
    catch {
        Write-Error "Excel export failed: $_"
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

            if ($urlPath -eq "/" -or $urlPath -eq "/index.html") {
                if (Test-Path $htmlFile) {
                    $html = [System.IO.File]::ReadAllText($htmlFile, [System.Text.Encoding]::UTF8)
                    Send-Response $context $html "text/html; charset=utf-8"
                } else {
                    Send-Response $context "index.html not found" "text/plain" 404
                }
            }
            elseif ($urlPath -eq "/print_statement.html" -or $urlPath -eq "/statement.html") {
                $pfile = Join-Path $baseDir "print_statement.html"
                if (Test-Path $pfile) {
                    $html = [System.IO.File]::ReadAllText($pfile, [System.Text.Encoding]::UTF8)
                    Send-Response $context $html "text/html; charset=utf-8"
                } else {
                    Send-Response $context "No statement currently ready for printing" "text/plain" 404
                }
            }
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
            elseif ($urlPath -eq "/api/data" -and $method -eq "GET") {
                if (Test-Path $dataFile) {
                    $json = [System.IO.File]::ReadAllText($dataFile, [System.Text.Encoding]::UTF8)
                    Send-Response $context $json "application/json; charset=utf-8"
                } else {
                    Send-Response $context '{"products":[],"dealers":[],"orders":[]}' "application/json"
                }
            }
            elseif ($urlPath -eq "/api/save" -and $method -eq "POST") {
                $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                $body = $reader.ReadToEnd()
                $reader.Close()
                
                [System.IO.File]::WriteAllText($dataFile, $body, [System.Text.Encoding]::UTF8)
                Send-Response $context '{"status":"ok","message":"Data saved to file successfully"}' "application/json"
            }
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
                $exportFile = if (Test-Path $excelExportPath) { $excelExportPath } else { "D:\excel\Wholesale_Dealer_Management_System.xlsx" }
                if (Test-Path $exportFile) {
                    $bytes = [System.IO.File]::ReadAllBytes($exportFile)
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
                $target = if (Test-Path $excelExportPath) { $excelExportPath } else { "D:\excel\Wholesale_Dealer_Management_System.xlsx" }
                if (Test-Path $target) {
                    Start-Process "explorer.exe" -ArgumentList "/select,`"$target`""
                    Send-Response $context '{"status":"ok"}' "application/json"
                } else {
                    Start-Process "explorer.exe" -ArgumentList "D:\excel"
                    Send-Response $context '{"status":"ok"}' "application/json"
                }
            }
            elseif ($urlPath -eq "/api/open-save-dialog" -and $method -eq "POST") {
                $dlgScript = Join-Path $baseDir "open_save_dialog.ps1"
                if (Test-Path $dlgScript) {
                    Start-Process "powershell.exe" -ArgumentList "-STA -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$dlgScript`""
                    Send-Response $context '{"status":"ok"}' "application/json"
                } else {
                    Start-Process "explorer.exe" -ArgumentList "D:\excel"
                    Send-Response $context '{"status":"ok"}' "application/json"
                }
            }
            elseif ($urlPath -eq "/api/open-folder" -and $method -eq "POST") {
                Start-Process "explorer.exe" -ArgumentList "D:\excel"
                Send-Response $context '{"status":"ok"}' "application/json"
            }
            elseif ($urlPath -eq "/api/open-statement-print" -and $method -eq "POST") {
                $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                $body = $reader.ReadToEnd()
                $reader.Close()
                
                $pfile = Join-Path $baseDir "print_statement.html"
                [System.IO.File]::WriteAllText($pfile, $body, [System.Text.Encoding]::UTF8)
                
                Send-Response $context '{"status":"ok"}' "application/json"
            }
            else {
                Send-Response $context "Not Found" "text/plain" 404
            }
        } catch {
            Write-Warning "Request processing error: $_"
            try { $context.Response.Abort() } catch {}
        }
    }
}
finally {
    $listener.Stop()
    $listener.Close()
}

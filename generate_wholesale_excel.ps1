# Complete Generator for Wholesale & Dealer Management Excel System
$ErrorActionPreference = "Stop"

$outputPath = "D:\excel\Wholesale_Dealer_Management_System.xlsx"
if (Test-Path $outputPath) {
    Remove-Item $outputPath -Force
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    
    # Keep 1 sheet initially, then add the rest in order
    while ($wb.Sheets.Count -gt 1) {
        $wb.Sheets.Item(2).Delete()
    }
    
    $wsDash = $wb.Sheets.Item(1)
    $wsDash.Name = "Dashboard & Statement"
    
    $wsOrders = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsDash)
    $wsOrders.Name = "Orders & Ledger"
    
    $wsDealers = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsOrders)
    $wsDealers.Name = "Dealers List"
    
    $wsProducts = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wsDealers)
    $wsProducts.Name = "Products Master"

    Write-Host "Created 4 connected sheets."

    # -------------------------------------------------------------
    # Styling Helper Constants (Excel COM uses BGR integer format)
    # -------------------------------------------------------------
    $cNavy       = 27 + (54 * 256) + (93 * 65536)      # #1B365D - Main Banner Navy
    $cAccentBlue = 43 + (87 * 256) + (154 * 65536)     # #2B579A - Accent Blue
    $cSlate      = 51 + (65 * 256) + (85 * 65536)      # #334155 - Slate Steel
    $cLightRow   = 248 + (250 * 256) + (252 * 65536)   # #F8FAFC - Soft Zebra Light
    $cBorder     = 203 + (213 * 256) + (225 * 65536)   # #CBD5E1 - Border Gray
    $cCardBg     = 241 + (245 * 256) + (249 * 65536)   # #F1F5F9 - Light Card Fill
    $cWhite      = 16777215                            # #FFFFFF - Pure White
    
    # Status Colors (Bootstrap-style modern pastels)
    $cGreenBg    = 209 + (231 * 256) + (221 * 65536)   # #D1E7DD - Paid Light Green
    $cGreenFg    = 15 + (81 * 256) + (50 * 65536)      # #0F5132 - Paid Dark Green Text
    $cYellowBg   = 255 + (243 * 256) + (205 * 65536)   # #FFF3CD - Partial Light Amber
    $cYellowFg   = 102 + (77 * 256) + (3 * 65536)      # #664D03 - Partial Dark Amber Text
    $cRedBg      = 248 + (215 * 256) + (218 * 65536)   # #F8D7DA - Pending Light Red
    $cRedFg      = 132 + (32 * 256) + (41 * 65536)     # #842029 - Pending Dark Red Text

    function Set-CellValue($range, $val) {
        if ($val -is [double] -or $val -is [int] -or $val -is [decimal]) {
            $range.Formula = "$val"
        } elseif ($val.ToString().StartsWith("=")) {
            $range.Formula = $val
        } else {
            $range.Value2 = [string]$val
        }
    }

    function Format-Header([object]$range, [int]$bgColor, [int]$fgColor = 16777215) {
        $range.Font.Name = "Segoe UI"
        $range.Font.Size = 10.5
        $range.Font.Bold = $true
        $range.Font.Color = $fgColor
        $range.Interior.Color = $bgColor
        $range.HorizontalAlignment = -4108 # xlCenter
        $range.VerticalAlignment = -4108   # xlCenter
        $range.RowHeight = 28
        $range.WrapText = $true
    }

    function Set-ThinBorders([object]$range) {
        $edges = @(7, 8, 9, 10, 11, 12) # xlEdgeLeft, Top, Bottom, Right, InsideVert, InsideHoriz
        foreach ($edge in $edges) {
            $border = $range.Borders.Item($edge)
            $border.LineStyle = 1 # xlContinuous
            $border.Weight = 2    # xlThin
            $border.Color = $cBorder
        }
    }

    # =============================================================
    # 1. PRODUCTS MASTER SHEET
    # =============================================================
    Write-Host "Populating Products Master sheet..."
    $ws = $wsProducts
    $ws.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true

    # Title Banner
    $ws.Range("B2:I2").Merge()
    $ws.Range("B2").Value2 = "WHOLESALE PRODUCTS & PRICE MASTER"
    $ws.Range("B2").Font.Name = "Segoe UI"
    $ws.Range("B2").Font.Size = 16
    $ws.Range("B2").Font.Bold = $true
    $ws.Range("B2").Font.Color = $cWhite
    $ws.Range("B2").Interior.Color = $cNavy
    $ws.Range("B2").HorizontalAlignment = -4108
    $ws.Range("B2").VerticalAlignment = -4108
    $ws.Range("B2").RowHeight = 36

    # Subtitle
    $ws.Range("B3:I3").Merge()
    $ws.Range("B3").Value2 = "Official Wholesale Catalog: Standard Unit Prices, Inventory Levels & Automatic Reorder Alerts"
    $ws.Range("B3").Font.Name = "Segoe UI"
    $ws.Range("B3").Font.Size = 9.5
    $ws.Range("B3").Font.Italic = $true
    $ws.Range("B3").Font.Color = $cSlate
    $ws.Range("B3").Interior.Color = $cCardBg
    $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B3").RowHeight = 20

    # Headers at Row 5
    $headersProducts = @("Item SKU", "Category", "Item Name & Description", "Packaging Unit", "Wholesale Unit Price (GH₵)", "Current Stock (Units)", "Reorder Alert Level", "Stock Status")
    for ($col = 0; $col -lt $headersProducts.Length; $col++) {
        Set-CellValue ($ws.Cells.Item(5, $col + 2)) $headersProducts[$col]
    }
    Format-Header ($ws.Range("B5:I5")) $cAccentBlue

    # Realistic Wholesale Products
    $products = @(
        @("SKU-1001", "Beverages", "Premium Roast Coffee Beans (1kg)", "Carton (12 Bags)", 145.00, 85, 20),
        @("SKU-1002", "Beverages", "Organic Green Tea Extract (100 bags)", "Case (24 Boxes)", 96.00, 14, 25),
        @("SKU-1003", "Beverages", "Pure Sparkling Mineral Water 500ml", "Tray (24 Bottles)", 22.50, 240, 50),
        @("SKU-1004", "Packaged Foods", "Extra Virgin Olive Oil 1L", "Case (12 Bottles)", 118.00, 60, 15),
        @("SKU-1005", "Packaged Foods", "Durum Wheat Semolina Pasta 500g", "Carton (20 Packs)", 34.00, 110, 30),
        @("SKU-1006", "Packaged Foods", "Organic Basmati Long Grain Rice 5kg", "Bale (6 Bags)", 72.50, 45, 15),
        @("SKU-1007", "Packaged Foods", "Canned San Marzano Tomatoes 800g", "Case (24 Cans)", 48.00, 9, 20),
        @("SKU-1008", "Home Care", "Ultra Concentrated Laundry Liquid 3L", "Case (4 Jugs)", 56.00, 75, 15),
        @("SKU-1009", "Home Care", "Antibacterial Surface Disinfectant 750ml", "Box (12 Sprays)", 38.40, 130, 25),
        @("SKU-1010", "Home Care", "Biodegradable Dishwashing Gel 1L", "Case (12 Bottles)", 42.00, 5, 20),
        @("SKU-1011", "Personal Care", "Moisturizing Shea Body Wash 500ml", "Case (12 Bottles)", 52.00, 95, 20),
        @("SKU-1012", "Personal Care", "Invigorating Daily Shampoo 400ml", "Case (12 Bottles)", 46.50, 18, 20),
        @("SKU-1013", "Snacks & Sweets", "Dark Chocolate Artisan Bars 100g", "Display (24 Bars)", 58.00, 120, 30),
        @("SKU-1014", "Snacks & Sweets", "Roasted Salted Almonds 250g", "Box (16 Bags)", 64.00, 0, 15),
        @("SKU-1015", "Snacks & Sweets", "Gluten-Free Sea Salt Crackers 150g", "Carton (18 Boxes)", 39.50, 80, 25)
    )

    $startRow = 6
    for ($i = 0; $i -lt $products.Length; $i++) {
        $r = $startRow + $i
        $item = $products[$i]
        Set-CellValue ($ws.Cells.Item($r, 2)) $item[0] # SKU
        Set-CellValue ($ws.Cells.Item($r, 3)) $item[1] # Category
        Set-CellValue ($ws.Cells.Item($r, 4)) $item[2] # Item Name
        Set-CellValue ($ws.Cells.Item($r, 5)) $item[3] # Packaging Unit
        Set-CellValue ($ws.Cells.Item($r, 6)) $item[4] # Wholesale Price
        Set-CellValue ($ws.Cells.Item($r, 7)) $item[5] # Stock
        Set-CellValue ($ws.Cells.Item($r, 8)) $item[6] # Reorder Level
        
        # Stock Status formula
        Set-CellValue ($ws.Cells.Item($r, 9)) ('=IF(G' + $r + '<=0, "Out of Stock", IF(G' + $r + '<=H' + $r + ', "Low Stock", "In Stock"))')

        $rowRange = $ws.Range("B$r`:I$r")
        $rowRange.Font.Name = "Segoe UI"
        $rowRange.Font.Size = 10
        $rowRange.VerticalAlignment = -4108
        $rowRange.RowHeight = 22
        if ($i % 2 -eq 1) {
            $rowRange.Interior.Color = $cLightRow
        }
    }
    $endProductRow = $startRow + $products.Length - 1

    # Formatting Product Columns
    $ws.Range("B6:B$endProductRow").HorizontalAlignment = -4108 # Center SKU
    $ws.Range("C6:C$endProductRow").HorizontalAlignment = -4108 # Center Category
    $ws.Range("E6:E$endProductRow").HorizontalAlignment = -4108 # Center Unit
    $ws.Range("F6:F$endProductRow").NumberFormat = """GH₵ ""#,##0.00"   # Price
    $ws.Range("G6:H$endProductRow").NumberFormat = "#,##0"       # Stocks
    $ws.Range("G6:H$endProductRow").HorizontalAlignment = -4152 # Right align quantities
    $ws.Range("I6:I$endProductRow").HorizontalAlignment = -4108 # Center Status

    Set-ThinBorders ($ws.Range("B5:I$endProductRow"))

    # Define Named Range for Products List (Item Names in Column D)
    $wb.Names.Add("ProductsList", $ws.Range("D6:D$endProductRow")) | Out-Null
    $wb.Names.Add("ProductCatalog", $ws.Range("D6:F$endProductRow")) | Out-Null

    # Conditional Formatting for Stock Status
    # 1 = xlCellValue, 3 = xlEqual
    $cf1 = $ws.Range("I6:I$endProductRow").FormatConditions.Add(1, 3, "=""Out of Stock""")
    $cf1.Interior.Color = $cRedBg
    $cf1.Font.Color = $cRedFg
    $cf1.Font.Bold = $true

    $cf2 = $ws.Range("I6:I$endProductRow").FormatConditions.Add(1, 3, "=""Low Stock""")
    $cf2.Interior.Color = $cYellowBg
    $cf2.Font.Color = $cYellowFg
    $cf2.Font.Bold = $true

    $cf3 = $ws.Range("I6:I$endProductRow").FormatConditions.Add(1, 3, "=""In Stock""")
    $cf3.Interior.Color = $cGreenBg
    $cf3.Font.Color = $cGreenFg

    # =============================================================
    # 2. DEALERS LIST SHEET
    # =============================================================
    Write-Host "Populating Dealers List sheet..."
    $ws = $wsDealers
    $ws.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true

    # Title Banner
    $ws.Range("B2:J2").Merge()
    $ws.Range("B2").Value2 = "AUTHORIZED DEALERS DIRECTORY & ACCOUNT SUMMARY"
    $ws.Range("B2").Font.Name = "Segoe UI"
    $ws.Range("B2").Font.Size = 16
    $ws.Range("B2").Font.Bold = $true
    $ws.Range("B2").Font.Color = $cWhite
    $ws.Range("B2").Interior.Color = $cNavy
    $ws.Range("B2").HorizontalAlignment = -4108
    $ws.Range("B2").VerticalAlignment = -4108
    $ws.Range("B2").RowHeight = 36

    # Subtitle
    $ws.Range("B3:J3").Merge()
    $ws.Range("B3").Value2 = "Dealer Network Directory with Automated Ledger Calculations for Lifetime Purchases, Payments & Outstanding Debt"
    $ws.Range("B3").Font.Name = "Segoe UI"
    $ws.Range("B3").Font.Size = 9.5
    $ws.Range("B3").Font.Italic = $true
    $ws.Range("B3").Font.Color = $cSlate
    $ws.Range("B3").Interior.Color = $cCardBg
    $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B3").RowHeight = 20

    # Headers at Row 5
    $headersDealers = @("Dealer ID", "Dealer Company Name", "Contact Person", "Phone / Mobile", "Territory / Location", "Payment Terms", "Total Purchased (GH₵)", "Total Paid (GH₵)", "Current Balance Due (GH₵)")
    for ($col = 0; $col -lt $headersDealers.Length; $col++) {
        Set-CellValue ($ws.Cells.Item(5, $col + 2)) $headersDealers[$col]
    }
    Format-Header ($ws.Range("B5:J5")) $cSlate

    $dealers = @(
        @("DLR-01", "Apex Supermarket Chain", "Marcus Vance", "+1 (555) 234-8901", "North Metro Hub", "Net 30"),
        @("DLR-02", "Silverline Wholesale Mart", "Elena Rostova", "+1 (555) 345-6712", "Downtown Central", "Net 15"),
        @("DLR-03", "Golden Gate Retailers", "David Chen", "+1 (555) 456-1123", "Westside District", "Net 30"),
        @("DLR-04", "Metro Cash & Carry", "Sarah Jenkins", "+1 (555) 567-8834", "East Industrial Park", "Cash On Delivery"),
        @("DLR-05", "Sunrise Convenience Co.", "Tariq Mansoor", "+1 (555) 678-9945", "Airport Commercial Zone", "Net 15"),
        @("DLR-06", "Prime Horizon Grocers", "Chloe Bennett", "+1 (555) 789-2256", "South Suburbs", "Net 30"),
        @("DLR-07", "Oakridge Provisions", "Samuel Adebayo", "+1 (555) 890-3367", "Harbor Commercial Area", "Net 15"),
        @("DLR-08", "Crestview General Store", "Maria Santos", "+1 (555) 901-4478", "Valley Town Center", "Cash On Delivery")
    )

    $startRow = 6
    for ($i = 0; $i -lt $dealers.Length; $i++) {
        $r = $startRow + $i
        $d = $dealers[$i]
        Set-CellValue ($ws.Cells.Item($r, 2)) $d[0] # ID
        Set-CellValue ($ws.Cells.Item($r, 3)) $d[1] # Name
        Set-CellValue ($ws.Cells.Item($r, 4)) $d[2] # Contact
        Set-CellValue ($ws.Cells.Item($r, 5)) $d[3] # Phone
        Set-CellValue ($ws.Cells.Item($r, 6)) $d[4] # Location
        Set-CellValue ($ws.Cells.Item($r, 7)) $d[5] # Terms

        # Dynamic SUMIF from Orders & Ledger
        # Col E in Orders & Ledger is Dealer Name, Col J is Total Amount Due
        Set-CellValue ($ws.Cells.Item($r, 8)) ('=SUMIF(''Orders & Ledger''!$E$6:$E$60, C' + $r + ', ''Orders & Ledger''!$J$6:$J$60)')
        # Col K in Orders & Ledger is Amount Paid
        Set-CellValue ($ws.Cells.Item($r, 9)) ('=SUMIF(''Orders & Ledger''!$E$6:$E$60, C' + $r + ', ''Orders & Ledger''!$K$6:$K$60)')
        # Col J is Balance Due
        Set-CellValue ($ws.Cells.Item($r, 10)) ('=H' + $r + ' - I' + $r)

        $rowRange = $ws.Range("B$r`:J$r")
        $rowRange.Font.Name = "Segoe UI"
        $rowRange.Font.Size = 10
        $rowRange.VerticalAlignment = -4108
        $rowRange.RowHeight = 22
        if ($i % 2 -eq 1) {
            $rowRange.Interior.Color = $cLightRow
        }
    }
    $endDealerRow = $startRow + $dealers.Length - 1

    # Formatting Dealer Columns
    $ws.Range("B6:B$endDealerRow").HorizontalAlignment = -4108 # Center ID
    $ws.Range("G6:G$endDealerRow").HorizontalAlignment = -4108 # Center Terms
    $ws.Range("H6:J$endDealerRow").NumberFormat = """GH₵ ""#,##0.00"   # Financial columns
    $ws.Range("H6:J$endDealerRow").HorizontalAlignment = -4152

    # Total Summary Row for Dealers
    $totRow = $endDealerRow + 1
    $ws.Range("B$totRow`:G$totRow").Merge()
    $ws.Range("B$totRow").Value2 = "NETWORK TOTALS"
    $ws.Range("B$totRow").Font.Bold = $true
    $ws.Range("B$totRow").HorizontalAlignment = -4152
    Set-CellValue ($ws.Range("H$totRow")) ("=SUM(H6:H$endDealerRow)")
    Set-CellValue ($ws.Range("I$totRow")) ("=SUM(I6:I$endDealerRow)")
    Set-CellValue ($ws.Range("J$totRow")) ("=SUM(J6:J$endDealerRow)")
    $ws.Range("H$totRow`:J$totRow").NumberFormat = """GH₵ ""#,##0.00"
    $ws.Range("B$totRow`:J$totRow").Font.Bold = $true
    $ws.Range("B$totRow`:J$totRow").Interior.Color = $cCardBg
    $ws.Range("B$totRow`:J$totRow").RowHeight = 26

    Set-ThinBorders ($ws.Range("B5:J$totRow"))

    # Define Named Range for Dealers List (Company Names in Column C)
    $wb.Names.Add("DealersList", $ws.Range("C6:C$endDealerRow")) | Out-Null
    $wb.Names.Add("DealersTable", $ws.Range("C6:J$endDealerRow")) | Out-Null

    # Conditional Formatting for Outstanding Balance > 0
    # 1 = xlCellValue, 5 = xlGreater
    $cfD1 = $ws.Range("J6:J$endDealerRow").FormatConditions.Add(1, 5, "0")
    $cfD1.Font.Bold = $true
    $cfD1.Font.Color = $cRedFg
    $cfD1.Interior.Color = $cRedBg

    # 1 = xlCellValue, 6 = xlEqual
    $cfD2 = $ws.Range("J6:J$endDealerRow").FormatConditions.Add(1, 6, "0")
    $cfD2.Font.Color = $cGreenFg
    $cfD2.Interior.Color = $cGreenBg

    # =============================================================
    # 3. ORDERS & PAYMENT LEDGER SHEET
    # =============================================================
    Write-Host "Populating Orders & Payment Ledger sheet..."
    $ws = $wsOrders
    $ws.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true

    # Title Banner
    $ws.Range("B2:M2").Merge()
    $ws.Range("B2").Value2 = "ORDERS TRANSACTION & PAYMENT LEDGER"
    $ws.Range("B2").Font.Name = "Segoe UI"
    $ws.Range("B2").Font.Size = 16
    $ws.Range("B2").Font.Bold = $true
    $ws.Range("B2").Font.Color = $cWhite
    $ws.Range("B2").Interior.Color = $cNavy
    $ws.Range("B2").HorizontalAlignment = -4108
    $ws.Range("B2").VerticalAlignment = -4108
    $ws.Range("B2").RowHeight = 36

    # Subtitle
    $ws.Range("B3:M3").Merge()
    $ws.Range("B3").Value2 = "Live Wholesale Order Book: Select Dealers & Products from Dropdowns. Auto-fetches Wholesale Price, Calculates Balances & Formats Status"
    $ws.Range("B3").Font.Name = "Segoe UI"
    $ws.Range("B3").Font.Size = 9.5
    $ws.Range("B3").Font.Italic = $true
    $ws.Range("B3").Font.Color = $cSlate
    $ws.Range("B3").Interior.Color = $cCardBg
    $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B3").RowHeight = 20

    # Headers at Row 5
    $headersOrders = @("Order ID", "Date", "Invoice #", "Dealer Name (Dropdown)", "Item Name (Dropdown)", "Packaging Unit", "Wholesale Unit Price (GH₵)", "Quantity Ordered", "Total Amount Due (GH₵)", "Amount Paid (GH₵)", "Balance Remaining (GH₵)", "Payment Status")
    for ($col = 0; $col -lt $headersOrders.Length; $col++) {
        Set-CellValue ($ws.Cells.Item(5, $col + 2)) $headersOrders[$col]
    }
    Format-Header ($ws.Range("B5:M5")) $cNavy

    # Sample Orders (15 realistic wholesale transactions covering Paid In Full, Partial, Pending)
    $orders = @(
        @("ORD-2026-101", "2026-09-02", "INV-8801", "Apex Supermarket Chain", "Premium Roast Coffee Beans (1kg)", 25, 3625.00),
        @("ORD-2026-102", "2026-09-04", "INV-8802", "Silverline Wholesale Mart", "Extra Virgin Olive Oil 1L", 30, 3540.00),
        @("ORD-2026-103", "2026-09-06", "INV-8803", "Golden Gate Retailers", "Organic Basmati Long Grain Rice 5kg", 40, 2000.00),
        @("ORD-2026-104", "2026-09-09", "INV-8804", "Metro Cash & Carry", "Ultra Concentrated Laundry Liquid 3L", 50, 2800.00),
        @("ORD-2026-105", "2026-09-12", "INV-8805", "Sunrise Convenience Co.", "Pure Sparkling Mineral Water 500ml", 80, 1800.00),
        @("ORD-2026-106", "2026-09-15", "INV-8806", "Prime Horizon Grocers", "Organic Green Tea Extract (100 bags)", 20, 1000.00),
        @("ORD-2026-107", "2026-09-18", "INV-8807", "Oakridge Provisions", "Durum Wheat Semolina Pasta 500g", 60, 2040.00),
        @("ORD-2026-108", "2026-09-20", "INV-8808", "Crestview General Store", "Moisturizing Shea Body Wash 500ml", 35, 1820.00),
        @("ORD-2026-109", "2026-09-22", "INV-8809", "Apex Supermarket Chain", "Dark Chocolate Artisan Bars 100g", 45, 1500.00),
        @("ORD-2026-110", "2026-09-25", "INV-8810", "Silverline Wholesale Mart", "Antibacterial Surface Disinfectant 750ml", 40, 1536.00),
        @("ORD-2026-111", "2026-09-27", "INV-8811", "Golden Gate Retailers", "Premium Roast Coffee Beans (1kg)", 15, 0.00),
        @("ORD-2026-112", "2026-09-28", "INV-8812", "Sunrise Convenience Co.", "Gluten-Free Sea Salt Crackers 150g", 50, 1000.00),
        @("ORD-2026-113", "2026-09-30", "INV-8813", "Prime Horizon Grocers", "Extra Virgin Olive Oil 1L", 25, 0.00),
        @("ORD-2026-114", "2026-10-02", "INV-8814", "Metro Cash & Carry", "Pure Sparkling Mineral Water 500ml", 100, 2250.00),
        @("ORD-2026-104", "2026-10-04", "INV-8815", "Oakridge Provisions", "Invigorating Daily Shampoo 400ml", 30, 700.00)
    )

    $startRow = 6
    $maxRows = 30 # Pre-configure 30 rows with formulas and dropdown validations!

    for ($i = 0; $i -lt $maxRows; $i++) {
        $r = $startRow + $i
        if ($i -lt $orders.Length) {
            $ord = $orders[$i]
            Set-CellValue ($ws.Cells.Item($r, 2)) $ord[0] # Order ID
            Set-CellValue ($ws.Cells.Item($r, 3)) $ord[1] # Date
            Set-CellValue ($ws.Cells.Item($r, 4)) $ord[2] # Invoice #
            Set-CellValue ($ws.Cells.Item($r, 5)) $ord[3] # Dealer Name
            Set-CellValue ($ws.Cells.Item($r, 6)) $ord[4] # Item Name
            Set-CellValue ($ws.Cells.Item($r, 9)) $ord[5] # Quantity
            Set-CellValue ($ws.Cells.Item($r, 11)) $ord[6] # Amount Paid
        } else {
            # Ready pre-formatted row for future order input
            Set-CellValue ($ws.Cells.Item($r, 2)) ("ORD-2026-" + (116 + ($i - $orders.Length)))
            Set-CellValue ($ws.Cells.Item($r, 4)) ("INV-" + (8816 + ($i - $orders.Length)))
        }

        # Dynamic Packaging Unit lookup: Products Master D:E (col 2 is Unit)
        Set-CellValue ($ws.Cells.Item($r, 7)) ('=IF(F' + $r + '="","",IFERROR(VLOOKUP(F' + $r + ', ''Products Master''!$D$6:$F$25, 2, FALSE), "-"))')

        # Dynamic Wholesale Price lookup: Products Master D:F (col 3 is Price)
        Set-CellValue ($ws.Cells.Item($r, 8)) ('=IF(F' + $r + '="","",IFERROR(VLOOKUP(F' + $r + ', ''Products Master''!$D$6:$F$25, 3, FALSE), 0))')

        # Total Amount Due: Quantity * Wholesale Price
        Set-CellValue ($ws.Cells.Item($r, 10)) ('=IF(OR(I' + $r + '="",H' + $r + '=""),"", I' + $r + '*H' + $r + ')')

        # Balance Left to Pay: Total Due - Amount Paid
        Set-CellValue ($ws.Cells.Item($r, 12)) ('=IF(J' + $r + '="","", J' + $r + ' - N(K' + $r + '))')

        # Status: Paid In Full, Partial Payment, Pending / Unpaid
        Set-CellValue ($ws.Cells.Item($r, 13)) ('=IF(J' + $r + '="","", IF(N(K' + $r + ')>=J' + $r + ', "Paid In Full", IF(N(K' + $r + ')>0, "Partial Payment", "Pending / Unpaid")))')

        $rowRange = $ws.Range("B$r`:M$r")
        $rowRange.Font.Name = "Segoe UI"
        $rowRange.Font.Size = 10
        $rowRange.VerticalAlignment = -4108
        $rowRange.RowHeight = 22
        if ($i % 2 -eq 1) {
            $rowRange.Interior.Color = $cLightRow
        }
    }
    $lastLedgerRow = $startRow + $maxRows - 1

    # Formatting Ledger Columns
    $ws.Range("B6:B$lastLedgerRow").HorizontalAlignment = -4108 # Order ID
    $ws.Range("C6:C$lastLedgerRow").NumberFormat = "yyyy-mm-dd"
    $ws.Range("C6:C$lastLedgerRow").HorizontalAlignment = -4108
    $ws.Range("D6:D$lastLedgerRow").HorizontalAlignment = -4108 # Invoice #
    $ws.Range("G6:G$lastLedgerRow").HorizontalAlignment = -4108 # Unit
    $ws.Range("H6:H$lastLedgerRow").NumberFormat = """GH₵ ""#,##0.00"   # Unit Price
    $ws.Range("I6:I$lastLedgerRow").NumberFormat = "#,##0"       # Quantity
    $ws.Range("I6:I$lastLedgerRow").HorizontalAlignment = -4152
    $ws.Range("J6:L$lastLedgerRow").NumberFormat = """GH₵ ""#,##0.00"   # Due, Paid, Balance
    $ws.Range("M6:M$lastLedgerRow").HorizontalAlignment = -4108 # Status

    # Summary Totals Row at the bottom of orders
    $totOrderRow = $lastLedgerRow + 1
    $ws.Range("B$totOrderRow`:H$totOrderRow").Merge()
    $ws.Range("B$totOrderRow").Value2 = "LEDGER TOTALS (CURRENT TRANSACTIONS)"
    $ws.Range("B$totOrderRow").Font.Bold = $true
    $ws.Range("B$totOrderRow").HorizontalAlignment = -4152
    Set-CellValue ($ws.Cells.Item($totOrderRow, 9)) ("=SUM(I6:I$lastLedgerRow)")
    Set-CellValue ($ws.Cells.Item($totOrderRow, 10)) ("=SUM(J6:J$lastLedgerRow)")
    Set-CellValue ($ws.Cells.Item($totOrderRow, 11)) ("=SUM(K6:K$lastLedgerRow)")
    Set-CellValue ($ws.Cells.Item($totOrderRow, 12)) ("=SUM(L6:L$lastLedgerRow)")
    $ws.Range("I$totOrderRow").NumberFormat = "#,##0"
    $ws.Range("J$totOrderRow`:L$totOrderRow").NumberFormat = """GH₵ ""#,##0.00"
    $ws.Range("B$totOrderRow`:M$totOrderRow").Font.Bold = $true
    $ws.Range("B$totOrderRow`:M$totOrderRow").Interior.Color = $cCardBg
    $ws.Range("B$totOrderRow`:M$totOrderRow").RowHeight = 26

    Set-ThinBorders ($ws.Range("B5:M$totOrderRow"))

    # Dropdown Data Validation for Dealer (Col E) & Product (Col F)
    # 3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween
    Write-Host "Configuring Dropdown Data Validations..."
    $dealerValRange = $ws.Range("E6:E$lastLedgerRow")
    $dealerValRange.Validation.Delete()
    $dealerValRange.Validation.Add(3, 1, 1, "=DealersList")
    $dealerValRange.Validation.IgnoreBlank = $true
    $dealerValRange.Validation.InCellDropdown = $true
    $dealerValRange.Validation.InputTitle = "Select Dealer"
    $dealerValRange.Validation.InputMessage = "Select an authorized dealer from directory."
    $dealerValRange.Validation.ShowInput = $true
    $dealerValRange.Validation.ShowError = $true

    $productValRange = $ws.Range("F6:F$lastLedgerRow")
    $productValRange.Validation.Delete()
    $productValRange.Validation.Add(3, 1, 1, "=ProductsList")
    $productValRange.Validation.IgnoreBlank = $true
    $productValRange.Validation.InCellDropdown = $true
    $productValRange.Validation.InputTitle = "Select Product"
    $productValRange.Validation.InputMessage = "Select a product to auto-fill pricing."
    $productValRange.Validation.ShowInput = $true
    $productValRange.Validation.ShowError = $true

    # Conditional Formatting for Payment Status (Col M)
    # Paid In Full (Green)
    $cfO1 = $ws.Range("M6:M$lastLedgerRow").FormatConditions.Add(1, 3, "=""Paid In Full""")
    $cfO1.Interior.Color = $cGreenBg
    $cfO1.Font.Color = $cGreenFg
    $cfO1.Font.Bold = $true

    # Partial Payment (Yellow)
    $cfO2 = $ws.Range("M6:M$lastLedgerRow").FormatConditions.Add(1, 3, "=""Partial Payment""")
    $cfO2.Interior.Color = $cYellowBg
    $cfO2.Font.Color = $cYellowFg
    $cfO2.Font.Bold = $true

    # Pending / Unpaid (Red)
    $cfO3 = $ws.Range("M6:M$lastLedgerRow").FormatConditions.Add(1, 3, "=""Pending / Unpaid""")
    $cfO3.Interior.Color = $cRedBg
    $cfO3.Font.Color = $cRedFg
    $cfO3.Font.Bold = $true

    # =============================================================
    # 4. DASHBOARD & STATEMENT SHEET
    # =============================================================
    Write-Host "Populating Executive Dashboard & Statement sheet..."
    $ws = $wsDash
    $ws.Activate()
    $excel.ActiveWindow.DisplayGridlines = $true

    # Title Banner
    $ws.Range("B2:J2").Merge()
    $ws.Range("B2").Value2 = "WHOLESALE OPERATIONS & DEALER MANAGEMENT SUITE"
    $ws.Range("B2").Font.Name = "Segoe UI"
    $ws.Range("B2").Font.Size = 17
    $ws.Range("B2").Font.Bold = $true
    $ws.Range("B2").Font.Color = $cWhite
    $ws.Range("B2").Interior.Color = $cNavy
    $ws.Range("B2").HorizontalAlignment = -4108
    $ws.Range("B2").VerticalAlignment = -4108
    $ws.Range("B2").RowHeight = 40

    # Subtitle
    $ws.Range("B3:J3").Merge()
    $ws.Range("B3").Value2 = "Executive Summary, Real-Time Financial Performance KPI Cards & Single Dealer Account Statement Generator"
    $ws.Range("B3").Font.Name = "Segoe UI"
    $ws.Range("B3").Font.Size = 10
    $ws.Range("B3").Font.Italic = $true
    $ws.Range("B3").Font.Color = $cSlate
    $ws.Range("B3").Interior.Color = $cCardBg
    $ws.Range("B3").HorizontalAlignment = -4108
    $ws.Range("B3").VerticalAlignment = -4108
    $ws.Range("B3").RowHeight = 22

    # 4 KPI Cards across row 5-7
    # Card 1: Total Sales Revenue (B5:C7)
    $ws.Range("B5:C5").Merge()
    $ws.Range("B5").Value2 = "TOTAL SALES REVENUE"
    $ws.Range("B5").Font.Size = 9.5
    $ws.Range("B5").Font.Bold = $true
    $ws.Range("B5").Font.Color = $cWhite
    $ws.Range("B5").Interior.Color = $cAccentBlue
    $ws.Range("B5").HorizontalAlignment = -4108

    $ws.Range("B6:C7").Merge()
    Set-CellValue ($ws.Range("B6")) ("=SUM('Orders & Ledger'!J6:J60)")
    $ws.Range("B6").Font.Size = 16
    $ws.Range("B6").Font.Bold = $true
    $ws.Range("B6").Font.Color = $cNavy
    $ws.Range("B6").Interior.Color = $cCardBg
    $ws.Range("B6").NumberFormat = """GH₵ ""#,##0.00"
    $ws.Range("B6").HorizontalAlignment = -4108
    $ws.Range("B6").VerticalAlignment = -4108
    Set-ThinBorders ($ws.Range("B5:C7"))

    # Card 2: Total Cash Collected (D5:E7)
    $ws.Range("D5:E5").Merge()
    $ws.Range("D5").Value2 = "TOTAL CASH COLLECTED"
    $ws.Range("D5").Font.Size = 9.5
    $ws.Range("D5").Font.Bold = $true
    $ws.Range("D5").Font.Color = $cWhite
    $ws.Range("D5").Interior.Color = $cGreenFg
    $ws.Range("D5").HorizontalAlignment = -4108

    $ws.Range("D6:E7").Merge()
    Set-CellValue ($ws.Range("D6")) ("=SUM('Orders & Ledger'!K6:K60)")
    $ws.Range("D6").Font.Size = 16
    $ws.Range("D6").Font.Bold = $true
    $ws.Range("D6").Font.Color = $cGreenFg
    $ws.Range("D6").Interior.Color = $cCardBg
    $ws.Range("D6").NumberFormat = """GH₵ ""#,##0.00"
    $ws.Range("D6").HorizontalAlignment = -4108
    $ws.Range("D6").VerticalAlignment = -4108
    Set-ThinBorders ($ws.Range("D5:E7"))

    # Card 3: Total Outstanding Receivables (F5:G7)
    $ws.Range("F5:G5").Merge()
    $ws.Range("F5").Value2 = "OUTSTANDING RECEIVABLES"
    $ws.Range("F5").Font.Size = 9.5
    $ws.Range("F5").Font.Bold = $true
    $ws.Range("F5").Font.Color = $cWhite
    $ws.Range("F5").Interior.Color = $cRedFg
    $ws.Range("F5").HorizontalAlignment = -4108

    $ws.Range("F6:G7").Merge()
    Set-CellValue ($ws.Range("F6")) ("=SUM('Orders & Ledger'!L6:L60)")
    $ws.Range("F6").Font.Size = 16
    $ws.Range("F6").Font.Bold = $true
    $ws.Range("F6").Font.Color = $cRedFg
    $ws.Range("F6").Interior.Color = $cCardBg
    $ws.Range("F6").NumberFormat = """GH₵ ""#,##0.00"
    $ws.Range("F6").HorizontalAlignment = -4108
    $ws.Range("F6").VerticalAlignment = -4108
    Set-ThinBorders ($ws.Range("F5:G7"))

    # Card 4: Orders & Dealer Counts (H5:I7)
    $ws.Range("H5:I5").Merge()
    $ws.Range("H5").Value2 = "NETWORK HEALTH"
    $ws.Range("H5").Font.Size = 9.5
    $ws.Range("H5").Font.Bold = $true
    $ws.Range("H5").Font.Color = $cWhite
    $ws.Range("H5").Interior.Color = $cSlate
    $ws.Range("H5").HorizontalAlignment = -4108

    $ws.Range("H6:I7").Merge()
    Set-CellValue ($ws.Range("H6")) ('="Dealers: " & COUNTA(''Dealers List''!C6:C30) & " | Orders: " & COUNTA(''Orders & Ledger''!B6:B25)')
    $ws.Range("H6").Font.Size = 12
    $ws.Range("H6").Font.Bold = $true
    $ws.Range("H6").Font.Color = $cSlate
    $ws.Range("H6").Interior.Color = $cCardBg
    $ws.Range("H6").HorizontalAlignment = -4108
    $ws.Range("H6").VerticalAlignment = -4108
    Set-ThinBorders ($ws.Range("H5:I7"))

    # Section 2: Interactive Single Dealer Account Statement
    $ws.Range("B9:I9").Merge()
    $ws.Range("B9").Value2 = "INDIVIDUAL DEALER STATEMENT & ACCOUNT LOOKUP"
    $ws.Range("B9").Font.Name = "Segoe UI"
    $ws.Range("B9").Font.Size = 12
    $ws.Range("B9").Font.Bold = $true
    $ws.Range("B9").Font.Color = $cWhite
    $ws.Range("B9").Interior.Color = $cNavy
    $ws.Range("B9").HorizontalAlignment = -4108
    $ws.Range("B9").VerticalAlignment = -4108
    $ws.Range("B9").RowHeight = 28

    # Lookup Selector at Row 11
    $ws.Range("B11:C11").Merge()
    $ws.Range("B11").Value2 = "Select Dealer Name:"
    $ws.Range("B11").Font.Bold = $true
    $ws.Range("B11").Font.Size = 11
    $ws.Range("B11").HorizontalAlignment = -4152

    $ws.Range("D11:F11").Merge()
    $ws.Range("D11").Value2 = "Apex Supermarket Chain"
    $ws.Range("D11").Font.Bold = $true
    $ws.Range("D11").Font.Size = 11
    $ws.Range("D11").Interior.Color = 16777160 # Light yellow input tint
    $ws.Range("D11").HorizontalAlignment = -4108
    $ws.Range("D11").VerticalAlignment = -4108
    $ws.Range("D11").RowHeight = 26

    # Dropdown on D11
    $valD11 = $ws.Range("D11").Validation
    $valD11.Delete()
    $valD11.Add(3, 1, 1, "=DealersList")
    $valD11.InCellDropdown = $true
    $valD11.InputTitle = "Choose Dealer"
    $valD11.InputMessage = "Click arrow to inspect individual statement."
    $valD11.ShowInput = $true

    # Statement details card (Rows 13-20)
    $stmtItems = @(
        @("Contact Person:", '=IFERROR(VLOOKUP(D11, DealersTable, 2, FALSE), "-")', "@"),
        @("Phone / Contact Number:", '=IFERROR(VLOOKUP(D11, DealersTable, 3, FALSE), "-")', "@"),
        @("Operating Territory / Location:", '=IFERROR(VLOOKUP(D11, DealersTable, 4, FALSE), "-")', "@"),
        @("Approved Payment Terms:", '=IFERROR(VLOOKUP(D11, DealersTable, 5, FALSE), "-")', "@"),
        @("Lifetime Purchases Amount:", '=IFERROR(VLOOKUP(D11, DealersTable, 6, FALSE), 0)', """GH₵ ""#,##0.00"),
        @("Total Cash Paid to Date:", '=IFERROR(VLOOKUP(D11, DealersTable, 7, FALSE), 0)', """GH₵ ""#,##0.00"),
        @("Current Outstanding Balance:", '=IFERROR(VLOOKUP(D11, DealersTable, 8, FALSE), 0)', """GH₵ ""#,##0.00"),
        @("Account Health Status:", '=IF(D19<=0, "Account In Good Standing (Paid)", "Outstanding Debt Due")', "@")
    )

    for ($idx = 0; $idx -lt $stmtItems.Length; $idx++) {
        $r = 13 + $idx
        $ws.Range("B$r`:C$r").Merge()
        $ws.Range("B$r").Value2 = $stmtItems[$idx][0]
        $ws.Range("B$r").Font.Bold = $true
        $ws.Range("B$r").Font.Size = 10.5
        $ws.Range("B$r").HorizontalAlignment = -4152

        $ws.Range("D$r`:G$r").Merge()
        Set-CellValue ($ws.Range("D$r")) $stmtItems[$idx][1]
        $ws.Range("D$r").Font.Size = 11
        $ws.Range("D$r").NumberFormat = $stmtItems[$idx][2]
        $ws.Range("D$r").HorizontalAlignment = -4108
        $ws.Range("D$r").VerticalAlignment = -4108

        if ($idx % 2 -eq 1) {
            $ws.Range("B$r`:G$r").Interior.Color = $cLightRow
        }
        $ws.Range("B$r`:G$r").RowHeight = 24

        # Highlight Balance and Status rows
        if ($r -eq 19) {
            $ws.Range("D$r").Font.Bold = $true
            $ws.Range("D$r").Font.Size = 12
            $ws.Range("B$r`:G$r").Interior.Color = 15329768
        }
        if ($r -eq 20) {
            $ws.Range("D$r").Font.Bold = $true
        }
    }
    Set-ThinBorders ($ws.Range("B13:G20"))

    # Quick User Guide Box on the right (H11:I18)
    $ws.Range("H11:I11").Merge()
    $ws.Range("H11").Value2 = "WORKFLOW INSTRUCTIONS"
    $ws.Range("H11").Font.Bold = $true
    $ws.Range("H11").Font.Size = 10
    $ws.Range("H11").Font.Color = $cWhite
    $ws.Range("H11").Interior.Color = $cSlate
    $ws.Range("H11").HorizontalAlignment = -4108

    $guideRows = @(
        "1. Open 'Orders & Ledger' to log daily wholesale orders.",
        "2. Choose Dealer and Product from the dropdowns.",
        "3. Wholesale price & packaging unit auto-fill instantly.",
        "4. Enter Quantity & Amount Paid; Balance & Status update automatically.",
        "5. 'Dealers List' maintains auto-updating lifetime ledger and debt.",
        "6. Use cell D11 above to pull up a full account statement for any dealer."
    )
    for ($g = 0; $g -lt $guideRows.Length; $g++) {
        $gr = 12 + $g
        $ws.Range("H$gr`:I$gr").Merge()
        $ws.Range("H$gr").Value2 = $guideRows[$g]
        $ws.Range("H$gr").Font.Size = 8.5
        $ws.Range("H$gr").WrapText = $true
        $ws.Range("H$gr").VerticalAlignment = -4108
        $ws.Range("H$gr").Interior.Color = $cCardBg
        $ws.Range("H$gr").RowHeight = 28
    }
    Set-ThinBorders ($ws.Range("H11:I17"))

    # Auto-fit columns across all sheets with padding
    Write-Host "Auto-fitting columns across all sheets..."
    foreach ($sheet in $wb.Sheets) {
        $sheet.Activate()
        $sheet.Columns.AutoFit() | Out-Null
        for ($c = 1; $c -le 15; $c++) {
            $colWidth = $sheet.Columns.Item($c).ColumnWidth
            if ($colWidth -lt 14) {
                $sheet.Columns.Item($c).ColumnWidth = 14
            } else {
                $sheet.Columns.Item($c).ColumnWidth = [Math]::Ceiling($colWidth + 3)
            }
        }
    }

    # Custom column widths on Dashboard for clean presentation
    $wsDash.Columns.Item(2).ColumnWidth = 26
    $wsDash.Columns.Item(3).ColumnWidth = 16
    $wsDash.Columns.Item(4).ColumnWidth = 18
    $wsDash.Columns.Item(5).ColumnWidth = 18
    $wsDash.Columns.Item(6).ColumnWidth = 18
    $wsDash.Columns.Item(7).ColumnWidth = 18
    $wsDash.Columns.Item(8).ColumnWidth = 22
    $wsDash.Columns.Item(9).ColumnWidth = 24

    # Calculate full workbook
    $excel.CalculateFull()

    # Set Dashboard as the default active sheet
    $wsDash.Activate()

    # Save workbook
    $wb.SaveAs($outputPath, 51) # 51 = xlOpenXMLWorkbook (.xlsx)
    Write-Host "Workbook saved successfully to $outputPath"
}
finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

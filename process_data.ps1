# process_data.ps1 - Bidirectional conversion and formatting for data 1.xlsx and data 2.xlsx
$ErrorActionPreference = "Stop"

$workspaceDir = if ($PSScriptRoot) { $PSScriptRoot } else { "." }
$data1Path = Join-Path $workspaceDir "data 1.xlsx"
$data2Path = Join-Path $workspaceDir "data 2.xlsx"

# Helper function to get modification time
function Get-FileTime($path) {
    if (Test-Path $path) {
        return (Get-Item $path).LastWriteTime
    }
    return [DateTime]::MinValue
}

$time1 = Get-FileTime $data1Path
$time2 = Get-FileTime $data2Path

Write-Host "Checking files..."
Write-Host "  data 1.xlsx modification time: $time1"
Write-Host "  data 2.xlsx modification time: $time2"

if ($time1 -eq [DateTime]::MinValue -and $time2 -eq [DateTime]::MinValue) {
    Write-Error "Neither data 1.xlsx nor data 2.xlsx was found in $workspaceDir"
    exit 1
}

# Decide which direction to run
$convert2to1 = $false
$reason = ""

if ($time2 -ne [DateTime]::MinValue -and ($time1 -eq [DateTime]::MinValue -or $time2 -gt $time1)) {
    $convert2to1 = $true
    $reason = "data 2.xlsx is newer or data 1.xlsx is missing. Converting data 2 -> data 1 (Creating INVOICE, TRIP ID, SER-PARTS, ROUTE sheets)."
} else {
    $convert2to1 = $false
    $reason = "data 1.xlsx is newer or data 2.xlsx is missing. Converting data 1 -> data 2 (Removing generated sheets)."
}

Write-Host "`n>>> Action: $reason"

# Initialize Excel COM Object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    if ($convert2to1) {
        # Converting data 2 -> data 1
        Write-Host "Copying data 2.xlsx to data 1.xlsx..."
        Copy-Item $data2Path $data1Path -Force
        
        Write-Host "Loading data 1.xlsx..."
        $wb = $excel.Workbooks.Open($data1Path)
        
        # Check if sheets already exist and delete them to start fresh
        try {
            $wsOld = $wb.Sheets.Item("INVOICE")
            $wsOld.Delete()
            Write-Host "Deleted existing INVOICE sheet in workbook."
        } catch {}
        try {
            $wsOld = $wb.Sheets.Item("TRIP ID")
            $wsOld.Delete()
            Write-Host "Deleted existing TRIP ID sheet in workbook."
        } catch {}
        try {
            $wsOld = $wb.Sheets.Item("SER-PARTS")
            $wsOld.Delete()
            Write-Host "Deleted existing SER-PARTS sheet in workbook."
        } catch {}
        try {
            $wsOld = $wb.Sheets.Item("ROUTE")
            $wsOld.Delete()
            Write-Host "Deleted existing ROUTE sheet in workbook."
        } catch {}
        
        # Insert worksheets before Response sheet
        $wsResp = $wb.Sheets.Item("Response")
        $wsRouteOut = $wb.Sheets.Add($wsResp)
        $wsRouteOut.Name = "ROUTE"
        $wsSerParts = $wb.Sheets.Add($wsResp)
        $wsSerParts.Name = "SER-PARTS"
        $wsTrip = $wb.Sheets.Add($wsResp)
        $wsTrip.Name = "TRIP ID"
        $wsInv = $wb.Sheets.Add($wsResp)
        $wsInv.Name = "INVOICE"
        Write-Host "Created new INVOICE, TRIP ID, SER-PARTS, and ROUTE sheets."
        
        # Set up Column widths for ROUTE
        # Columns: Platform, docket_number, item_sku, invoice_number, Route, Type, Remarks, Bags, Roll, Trip Id
        $routeWidths = @(36.8, 14.4, 37.5, 17.8, 12.9, 8.5, 25.9, 8.5, 8.5, 11.2)
        for ($c = 1; $c -le 10; $c++) {
            $wsRouteOut.Columns.Item($c).ColumnWidth = $routeWidths[$c - 1]
        }
        $wsRouteOut.Columns.Item(4).NumberFormat = "@"
        
        # Setup Headers for ROUTE
        $routeHeaders = @("Platform", "docket_number", "item_sku", "invoice_number", "Route", "Type", "Remarks", "Bags", "Roll", "Trip Id")
        for ($c = 1; $c -le 10; $c++) {
            $cell = $wsRouteOut.Cells(1, $c)
            $cell.Value2 = $routeHeaders[$c - 1]
        }
        $routeHeaderRange = $wsRouteOut.Range($wsRouteOut.Cells(1, 1), $wsRouteOut.Cells(1, 10))
        $routeHeaderRange.RowHeight = 14.5
        $routeHeaderRange.Font.Name = "Calibri"
        $routeHeaderRange.Font.Size = 11
        $routeHeaderRange.Font.Bold = $true
        $routeHeaderRange.Font.Color = 0
        $routeHeaderRange.HorizontalAlignment = -4108 # xlCenter
        $routeHeaderRange.VerticalAlignment = -4108   # xlCenter
        $routeHeaderRange.Borders.LineStyle = 1       # xlContinuous
        $routeHeaderRange.Borders.Weight = 2          # xlThin
        $routeHeaderRange.Borders.Color = 0
        
        # Set up Column widths for SER-PARTS
        # Columns: SERVICE, docket_number, item_sku, invoice_number, Route
        $serWidths = @(36.1, 15.6, 39.1, 17.5, 10.9)
        for ($c = 1; $c -le 5; $c++) {
            $wsSerParts.Columns.Item($c).ColumnWidth = $serWidths[$c - 1]
        }
        $wsSerParts.Columns.Item(4).NumberFormat = "@"
        
        # Setup Headers for SER-PARTS
        $serHeaders = @("SERVICE", "docket_number", "item_sku", "invoice_number", "Route")
        for ($c = 1; $c -le 5; $c++) {
            $cell = $wsSerParts.Cells(1, $c)
            $cell.Value2 = $serHeaders[$c - 1]
        }
        $serHeaderRange = $wsSerParts.Range($wsSerParts.Cells(1, 1), $wsSerParts.Cells(1, 5))
        $serHeaderRange.RowHeight = 14.5
        $serHeaderRange.Font.Name = "Calibri"
        $serHeaderRange.Font.Size = 11
        $serHeaderRange.Font.Bold = $true
        $serHeaderRange.Font.Color = 0
        $serHeaderRange.HorizontalAlignment = -4108 # xlCenter
        $serHeaderRange.VerticalAlignment = -4108   # xlCenter
        $serHeaderRange.Borders.LineStyle = 1       # xlContinuous
        $serHeaderRange.Borders.Weight = 2          # xlThin
        $serHeaderRange.Borders.Color = 0

        # Set up Column Widths for TRIP ID
        # Columns: Trip Number, Route, cp name, vehical no, remarks, helper id
        $tripWidths = @(18.0, 12.0, 18.0, 14.0, 22.0, 48.0)
        for ($c = 1; $c -le 6; $c++) {
            $wsTrip.Columns.Item($c).ColumnWidth = $tripWidths[$c - 1]
        }

        # Setup Headers for TRIP ID
        $tripHeaders = @("Trip Number", "Route", "cp name", "vehical no", "remarks", "helper id")
        for ($c = 1; $c -le 6; $c++) {
            $cell = $wsTrip.Cells(1, $c)
            $cell.Value2 = $tripHeaders[$c - 1]
        }
        $tripHeaderRange = $wsTrip.Range($wsTrip.Cells(1, 1), $wsTrip.Cells(1, 6))
        $tripHeaderRange.RowHeight = 14.5
        $tripHeaderRange.Font.Name = "Calibri"
        $tripHeaderRange.Font.Size = 11
        $tripHeaderRange.Font.Bold = $true
        $tripHeaderRange.Font.Color = 0
        $tripHeaderRange.HorizontalAlignment = -4108 # xlCenter
        $tripHeaderRange.VerticalAlignment = -4108   # xlCenter
        $tripHeaderRange.Borders.LineStyle = 1       # xlContinuous
        $tripHeaderRange.Borders.Weight = 2          # xlThin
        $tripHeaderRange.Borders.Color = 0

        # Set up Column Widths for INVOICE
        $invWidths = @(36.5, 42.9, 110.2, 117.6, 32.8, 32.9, 107.6, 19.9, 40.4, 51.6, 30.2)
        for ($c = 1; $c -le 11; $c++) {
            $wsInv.Columns.Item($c).ColumnWidth = $invWidths[$c - 1]
        }
        $wsInv.Columns.Item(10).NumberFormat = "@"
        
        # Format Header Row (Row 1) of INVOICE
        $invHeaders = @("PLATFORM", "DOCKET NO", "DESC", "CUST NAME", "MOB NO", "MOB NO", "ADDRESS", "CODE", "PAY TYPE", "INVOICE NO", "ROUTE")
        for ($c = 1; $c -le 11; $c++) {
            $cell = $wsInv.Cells(1, $c)
            $cell.Value2 = $invHeaders[$c - 1]
        }
        $invHeaderRange = $wsInv.Range($wsInv.Cells(1, 1), $wsInv.Cells(1, 11))
        $invHeaderRange.RowHeight = 50
        $invHeaderRange.Font.Name = "Calibri"
        $invHeaderRange.Font.Size = 30
        $invHeaderRange.Font.Bold = $true
        $invHeaderRange.Font.Color = 0
        $invHeaderRange.Interior.Color = 14806254
        $invHeaderRange.HorizontalAlignment = -4108 # xlCenter
        $invHeaderRange.VerticalAlignment = -4108   # xlCenter
        $invHeaderRange.WrapText = $true
        $invHeaderRange.Borders.LineStyle = 1       # xlContinuous
        $invHeaderRange.Borders.Weight = 2          # xlThin
        $invHeaderRange.Borders.Color = 0
        
        # Load Route Plan sheet — find dynamically (name includes today's date, e.g. "Route plan  24.06.2026")
        $wsRoute = $null
        for ($si = 1; $si -le $wb.Sheets.Count; $si++) {
            $sName = $wb.Sheets.Item($si).Name
            if ($sName -like "Route plan*") {
                $wsRoute = $wb.Sheets.Item($si)
                Write-Host "Found Route Plan sheet: '$sName'"
                break
            }
        }
        if ($null -eq $wsRoute) {
            throw "Could not find a sheet starting with 'Route plan' in the workbook. Available sheets: $(($wb.Sheets | ForEach-Object { $_.Name }) -join ', ')"
        }
        $routeRows = $wsRoute.UsedRange.Rows.Count
        Write-Host "Reading rows from Route plan... Total rows in Route Plan: $routeRows"
        
        # Scan header row (Row 1) of Route Plan sheet to find column indices
        $colsCount = $wsRoute.UsedRange.Columns.Count
        $headers = @{}
        for ($c = 1; $c -le $colsCount; $c++) {
            $rawVal = $wsRoute.Cells(1, $c).Text
            if ($rawVal) {
                $norm = $rawVal.Trim().ToLower() -replace '\s+', '_' -replace '_+', '_'
                if ($norm -ne "" -and -not $headers.ContainsKey($norm)) {
                    $headers[$norm] = $c
                }
            }
        }

        # Resolve column indices dynamically (with fallback to original static defaults)
        $platformCol = if ($headers.ContainsKey("platform")) { $headers["platform"] } else { 1 }
        $docketCol = if ($headers.ContainsKey("docket_number")) { $headers["docket_number"] } elseif ($headers.ContainsKey("docket_no")) { $headers["docket_no"] } else { 19 }
        $descCol = if ($headers.ContainsKey("item_sku")) { $headers["item_sku"] } elseif ($headers.ContainsKey("desc")) { $headers["desc"] } else { 7 }
        $cxNameCol = if ($headers.ContainsKey("cx_name")) { $headers["cx_name"] } elseif ($headers.ContainsKey("cust_name")) { $headers["cust_name"] } else { 3 }
        $mob1Col = if ($headers.ContainsKey("mobile_number")) { $headers["mobile_number"] } elseif ($headers.ContainsKey("mob_no")) { $headers["mob_no"] } else { 8 }
        $mob2Col = if ($headers.ContainsKey("alternate_mobile_number")) { $headers["alternate_mobile_number"] } elseif ($headers.ContainsKey("alt_mob_no")) { $headers["alt_mob_no"] } else { 9 }
        $addressCol = if ($headers.ContainsKey("address")) { $headers["address"] } else { 4 }
        $codeCol = if ($headers.ContainsKey("pincode")) { $headers["pincode"] } elseif ($headers.ContainsKey("pin_code")) { $headers["pin_code"] } elseif ($headers.ContainsKey("code")) { $headers["code"] } else { 5 }
        $payTypeCol = if ($headers.ContainsKey("payment_type")) { $headers["payment_type"] } elseif ($headers.ContainsKey("pay_type")) { $headers["pay_type"] } else { 17 }
        $invoiceNoCol = if ($headers.ContainsKey("invoice_number")) { $headers["invoice_number"] } elseif ($headers.ContainsKey("invoice_no")) { $headers["invoice_no"] } else { 18 }
        $routeCol = if ($headers.ContainsKey("route")) { $headers["route"] } else { 6 }
        $typeCol = if ($headers.ContainsKey("storage_location")) { $headers["storage_location"] } elseif ($headers.ContainsKey("type")) { $headers["type"] } else { 15 }
        $orderTypeCol = if ($headers.ContainsKey("order_type")) { $headers["order_type"] } else { 21 }

        # Search for trip mapping file (newly added order doc or fallback to old trip id file)
        $orderPaths = @(
            $workspaceDir,
            "C:\Users\rahul.malaganvi\Downloads",
            "C:\Users\rahul.malaganvi\Documents",
            "C:\Users\rahul.malaganvi\Desktop"
        )
        $tripMapPath = $null
        $latestTime = [DateTime]::MinValue
        $isCsv = $false
        
        # 1. Search for Order*.csv or Order*.xlsx (the new order doc)
        foreach ($folder in $orderPaths) {
            if (Test-Path $folder) {
                Get-ChildItem -Path $folder -Filter "Order*" -ErrorAction SilentlyContinue | ForEach-Object {
                    if ($_.Name -like "*.csv" -or $_.Name -like "*.xlsx") {
                        if ($_.LastWriteTime -gt $latestTime) {
                            $latestTime = $_.LastWriteTime
                            $tripMapPath = $_.FullName
                            $isCsv = ($_.Name -like "*.csv")
                        }
                    }
                }
            }
        }

        # 2. Fall back to old trip mapping files if no Order* file found
        if ($null -eq $tripMapPath) {
            $oldTripPaths = @(
                (Join-Path $workspaceDir "new trip id.xlsx"),
                (Join-Path $workspaceDir "trip id.xlsx"),
                (Join-Path (Split-Path $workspaceDir -Parent) "new trip id.xlsx"),
                (Join-Path (Split-Path $workspaceDir -Parent) "trip id.xlsx"),
                "C:\Users\rahul.malaganvi\Documents\new trip id.xlsx",
                "C:\Users\rahul.malaganvi\Documents\trip id.xlsx"
            )
            foreach ($p in $oldTripPaths) {
                if (Test-Path $p) {
                    $tripMapPath = $p
                    $isCsv = $false
                    break
                }
            }
        }

        $tripIdMap = @{}
        if ($null -ne $tripMapPath) {
            Write-Host "Found Trip ID mapping file at: $tripMapPath"
            Write-Host "Loading Trip ID mappings..."
            if ($isCsv) {
                try {
                    $csv = Import-Csv -Path $tripMapPath
                    $tripNumHeader = $null
                    $consNumHeader = $null
                    if ($csv.Count -gt 0) {
                        $firstRow = $csv[0]
                        foreach ($prop in $firstRow.PSObject.Properties) {
                            $name = $prop.Name.Trim().ToLower()
                            if ($name -eq "trip number" -or $name -eq "trip_number") {
                                $tripNumHeader = $prop.Name
                            }
                            if ($name -eq "consignment number" -or $name -eq "consignment_number" -or $name -eq "consignment no" -or $name -eq "docket no") {
                                $consNumHeader = $prop.Name
                            }
                        }
                    }
                    if ($null -eq $tripNumHeader) { $tripNumHeader = "Trip Number" }
                    if ($null -eq $consNumHeader) { $consNumHeader = "Consignment Number" }

                    foreach ($row in $csv) {
                        $tripNum = $row.$tripNumHeader
                        $consNum = $row.$consNumHeader
                        if ($null -ne $tripNum -and $null -ne $consNum) {
                            $t = $tripNum.ToString().Trim()
                            $c = $consNum.ToString().Trim().ToUpper()
                            if ($t -ne "" -and $c -ne "") {
                                $tripIdMap[$c] = $t
                            }
                        }
                    }
                } catch {
                    Write-Host "Error loading CSV with Import-Csv, attempting Excel open..."
                    $isCsv = $false
                }
            }

            if (-not $isCsv) {
                $wbTrip = $excel.Workbooks.Open($tripMapPath)
                $wsTripSrc = $wbTrip.Sheets.Item(1)
                $tripRows = $wsTripSrc.UsedRange.Rows.Count
                $tripCols = $wsTripSrc.UsedRange.Columns.Count
                
                # Resolve columns dynamically
                $tripColIdx = 1
                $consColIdx = 2
                for ($c = 1; $c -le $tripCols; $c++) {
                    $val = $wsTripSrc.Cells(1, $c).Text.Trim().ToLower()
                    if ($val -eq "trip number" -or $val -eq "trip_number" -or $val -eq "trip no") {
                        $tripColIdx = $c
                    }
                    if ($val -eq "consignment number" -or $val -eq "consignment_number" -or $val -eq "consignment no" -or $val -eq "docket no" -or $val -eq "docket_number") {
                        $consColIdx = $c
                    }
                }
                
                for ($tr = 2; $tr -le $tripRows; $tr++) {
                    $tripNum = $wsTripSrc.Cells($tr, $tripColIdx).Text.Trim()
                    $consNum = $wsTripSrc.Cells($tr, $consColIdx).Text.Trim().ToUpper()
                    if ($tripNum -ne "" -and $consNum -ne "") {
                        $tripIdMap[$consNum] = $tripNum
                    }
                }
                $wbTrip.Close($false)
            }
            Write-Host "Successfully loaded $($tripIdMap.Count) Trip ID mappings."
        } else {
            Write-Host "No Trip ID mapping file found. Trip ID fields will be left blank."
        }

        # Search for picklist mapping file dynamically in workspace, Downloads, Documents, or Desktop
        $picklistPaths = @(
            $workspaceDir,
            "C:\Users\rahul.malaganvi\Downloads",
            "C:\Users\rahul.malaganvi\Documents",
            "C:\Users\rahul.malaganvi\Desktop"
        )
        $picklistMapPath = $null
        $latestPickTime = [DateTime]::MinValue
        foreach ($folder in $picklistPaths) {
            if (Test-Path $folder) {
                Get-ChildItem -Path $folder -Filter "picklist*.xlsx" -ErrorAction SilentlyContinue | ForEach-Object {
                    if ($_.LastWriteTime -gt $latestPickTime) {
                        $latestPickTime = $_.LastWriteTime
                        $picklistMapPath = $_.FullName
                    }
                }
            }
        }

        $picklistMap = @{}
        if ($null -ne $picklistMapPath) {
            Write-Host "Found Picklist file at: $picklistMapPath"
            Write-Host "Loading Picklist mappings..."
            $wbPick = $excel.Workbooks.Open($picklistMapPath)
            $wsPickSrc = $wbPick.Sheets.Item(1)
            $pickRows = $wsPickSrc.UsedRange.Rows.Count
            $pickCols = $wsPickSrc.UsedRange.Columns.Count

            # Resolve columns dynamically
            $docketColIdx = -1
            $areaColIdx = -1
            for ($c = 1; $c -le $pickCols; $c++) {
                $val = $wsPickSrc.Cells(1, $c).Text.Trim().ToLower()
                if ($val -like "*docket*" -or $val -like "*consignment*" -or $val -like "*awb*") {
                    $docketColIdx = $c
                }
                if ($val -eq "area" -or $val -eq "type" -or $val -like "*slock*" -or $val -like "*storage_location*") {
                    $areaColIdx = $c
                }
            }

            if ($docketColIdx -eq -1) { $docketColIdx = 2 }
            if ($areaColIdx -eq -1) { $areaColIdx = 6 }

            for ($pr = 2; $pr -le $pickRows; $pr++) {
                $docketVal = $wsPickSrc.Cells($pr, $docketColIdx).Text.Trim().ToUpper()
                $areaVal = $wsPickSrc.Cells($pr, $areaColIdx).Text.Trim()
                if ($docketVal -ne "" -and $areaVal -ne "") {
                    $picklistMap[$docketVal] = $areaVal
                }
            }
            $wbPick.Close($false)
            Write-Host "Successfully loaded $($picklistMap.Count) Picklist mappings."
        } else {
            Write-Host "No Picklist file found."
        }

        # Load Service sheet data if it exists to help classify service dockets and populate SER-PARTS later
        $serviceSheetRows = @{}
        $wsServiceSrc = $null
        $sRows = 0
        try {
            $wsServiceSrc = $wb.Sheets.Item("Service")
        } catch {}
        
        if ($null -ne $wsServiceSrc) {
            Write-Host "Found Service sheet. Building service dockets mapping..."
            $sRows = $wsServiceSrc.UsedRange.Rows.Count
            $sCols = $wsServiceSrc.UsedRange.Columns.Count
            
            # Find column indices dynamically for Service sheet
            $sHeaders = @{}
            for ($c = 1; $c -le $sCols; $c++) {
                $rawVal = $wsServiceSrc.Cells(1, $c).Text
                if ($rawVal) {
                    $norm = $rawVal.Trim().ToLower() -replace '\s+', '_' -replace '_+', '_'
                    if ($norm -ne "" -and -not $sHeaders.ContainsKey($norm)) {
                        $sHeaders[$norm] = $c
                    }
                }
            }
            
            $sTypeColIdx = if ($sHeaders.ContainsKey("type")) { $sHeaders["type"] } elseif ($sHeaders.ContainsKey("ty")) { $sHeaders["ty"] } else { 1 }
            $sDocketColIdx = if ($sHeaders.ContainsKey("docket_number")) { $sHeaders["docket_number"] } elseif ($sHeaders.ContainsKey("docket_no")) { $sHeaders["docket_no"] } else { 6 }
            $sSkuColIdx = if ($sHeaders.ContainsKey("item_sku")) { $sHeaders["item_sku"] } else { 4 }
            $sInvoiceColIdx = if ($sHeaders.ContainsKey("invoice_number")) { $sHeaders["invoice_number"] } elseif ($sHeaders.ContainsKey("invoice_no")) { $sHeaders["invoice_no"] } else { 5 }
            
            $sOrderIdColIdx = if ($sHeaders.ContainsKey("order_random_id")) { $sHeaders["order_random_id"] } else { 3 }
            $sCxNameColIdx = if ($sHeaders.ContainsKey("cx_name")) { $sHeaders["cx_name"] } else { 7 }
            $sAddressColIdx = if ($sHeaders.ContainsKey("address")) { $sHeaders["address"] } else { 8 }
            $sPincodeColIdx = if ($sHeaders.ContainsKey("pincode")) { $sHeaders["pincode"] } else { 9 }
            $sMobileColIdx = if ($sHeaders.ContainsKey("mobile_number")) { $sHeaders["mobile_number"] } else { 10 }
            $sAltMobileColIdx = if ($sHeaders.ContainsKey("alternate_mobile_number")) { $sHeaders["alternate_mobile_number"] } else { 11 }
            $sRemarksColIdx = if ($sHeaders.ContainsKey("remarks")) { $sHeaders["remarks"] } else { 12 }
            $sServiceColIdx = if ($sHeaders.ContainsKey("service")) { $sHeaders["service"] } else { 2 }
            $sRouteColIdx = if ($sHeaders.ContainsKey("route")) { $sHeaders["route"] } else { $null }

            for ($sr = 2; $sr -le $sRows; $sr++) {
                $docketVal = $wsServiceSrc.Cells($sr, $sDocketColIdx).Text.Trim().ToUpper()
                if ($docketVal -ne "") {
                    $typeVal = $wsServiceSrc.Cells($sr, $sTypeColIdx).Text.Trim()
                    $skuVal = $wsServiceSrc.Cells($sr, $sSkuColIdx).Text.Trim()
                    $invoiceVal = $wsServiceSrc.Cells($sr, $sInvoiceColIdx).Text.Trim()
                    
                    $cell1 = $wsServiceSrc.Cells($sr, 1)
                    $sColor1 = $cell1.Interior.Color
                    $sColor1Idx = $cell1.Interior.ColorIndex
                    $isBlue1 = ($sColor1 -eq 15773696 -or $sColor1Idx -eq 33 -or $sColor1Idx -eq 34 -or $sColor1Idx -eq 41 -or $sColor1Idx -eq 42 -or ($sColor1 -eq 16773632) -or ($sColor1 -eq 16777164))

                    $cellDoc = $wsServiceSrc.Cells($sr, $sDocketColIdx)
                    $sColorDoc = $cellDoc.Interior.Color
                    $sColorDocIdx = $cellDoc.Interior.ColorIndex
                    $isBlueDoc = ($sColorDoc -eq 15773696 -or $sColorDocIdx -eq 33 -or $sColorDocIdx -eq 34 -or $sColorDocIdx -eq 41 -or $sColorDocIdx -eq 42 -or ($sColorDoc -eq 16773632) -or ($sColorDoc -eq 16777164))

                    $isRowBlue = ($isBlue1 -or $isBlueDoc)

                    $orderIdVal = $wsServiceSrc.Cells($sr, $sOrderIdColIdx).Text.Trim()
                    $cxNameVal = $wsServiceSrc.Cells($sr, $sCxNameColIdx).Text.Trim()
                    $addressVal = $wsServiceSrc.Cells($sr, $sAddressColIdx).Text.Trim()
                    $pincodeVal = $wsServiceSrc.Cells($sr, $sPincodeColIdx).Text.Trim()
                    $mobileVal = $wsServiceSrc.Cells($sr, $sMobileColIdx).Text.Trim()
                    $altMobileVal = $wsServiceSrc.Cells($sr, $sAltMobileColIdx).Text.Trim()
                    $remarksVal = $wsServiceSrc.Cells($sr, $sRemarksColIdx).Text.Trim()
                    $serviceVal = $wsServiceSrc.Cells($sr, $sServiceColIdx).Text.Trim()
                    $routeValFromService = if ($null -ne $sRouteColIdx) { $wsServiceSrc.Cells($sr, $sRouteColIdx).Text.Trim() } else { "" }

                    $serviceSheetRows[$docketVal] = [PSCustomObject]@{
                        Type                   = $typeVal
                        ItemSku                = $skuVal
                        InvoiceNumber          = $invoiceVal
                        IsBlue                 = $isRowBlue
                        OrderRandomId          = $orderIdVal
                        CxName                 = $cxNameVal
                        Address                = $addressVal
                        Pincode                = $pincodeVal
                        MobileNumber           = $mobileVal
                        AlternateMobileNumber  = $altMobileVal
                        Remarks                = $remarksVal
                        Service                = $serviceVal
                        Route                  = $routeValFromService
                    }
                }
            }
            Write-Host "Successfully mapped $($serviceSheetRows.Count) dockets from Service sheet."
        } else {
            Write-Host "Service sheet not found."
        }

        # Pre-pass to count duplicate customer names and addresses and resolve Route → Trip ID mappings
        $nameCounts   = @{}
        $addrCounts   = @{}
        $routeTripMap = @{}
        $docketRouteMap = @{}

        $routeByOrderId = @{}
        $routeByDocket = @{}
        $routeByCust = @{}
        $orderIdCol = if ($headers.ContainsKey("order_random_id")) { $headers["order_random_id"] } else { 2 }
        
        for ($r = 2; $r -le $routeRows; $r++) {
            $routeVal = $wsRoute.Cells($r, $routeCol).Text.Trim()
            if ($routeVal -eq "" -or $routeVal -eq "Not planned") { continue }
            
            $cxName = $wsRoute.Cells($r, $cxNameCol).Text.Trim().ToLower()
            $addr = $wsRoute.Cells($r, $addressCol).Text.Trim().ToLower()
            if ($cxName -ne "") {
                $nameCounts[$cxName]++
            }
            if ($addr -ne "") {
                $addrCounts[$addr]++
            }

            $orderId = $wsRoute.Cells($r, $orderIdCol).Text.Trim().ToUpper()
            $docketVal = $wsRoute.Cells($r, $docketCol).Text.Trim().ToUpper()
            if ($orderId -ne "") {
                $routeByOrderId[$orderId] = $routeVal
            }
            if ($docketVal -ne "") {
                $routeByDocket[$docketVal] = $routeVal
            }
            if ($cxName -ne "" -and $addr -ne "") {
                $routeByCust["$cxName|$addr"] = $routeVal
            }

            # Map Route to Trip Number if docket is in tripIdMap
            if ($tripIdMap.ContainsKey($docketVal)) {
                $tId = $tripIdMap[$docketVal]
                if ($tId -ne "") {
                    $routeTripMap[$routeVal] = $tId
                }
            }
        }

        # Initialize any remaining unmapped routes to blank
        for ($r = 2; $r -le $routeRows; $r++) {
            $routeVal = $wsRoute.Cells($r, $routeCol).Text.Trim()
            if ($routeVal -eq "" -or $routeVal -eq "Not planned") { continue }
            if (-not $routeTripMap.ContainsKey($routeVal)) {
                $routeTripMap[$routeVal] = ""
            }
        }
        
        $addedSerDockets = @{}
        $targetRouteRow = 2
        $targetSerRow   = 2
        $targetTripRow  = 2
        $targetInvRow   = 2
        
        for ($r = 2; $r -le $routeRows; $r++) {
            $routeVal = $wsRoute.Cells($r, $routeCol).Text.Trim()
            
            # Skip if Route is empty or "Not planned"
            if ($routeVal -eq "" -or $routeVal -eq "Not planned") {
                continue
            }
            
            # Map columns dynamically
            $platform   = $wsRoute.Cells($r, $platformCol).Text
            $docket     = $wsRoute.Cells($r, $docketCol).Text.Trim()
            $desc       = $wsRoute.Cells($r, $descCol).Text
            $cxName     = $wsRoute.Cells($r, $cxNameCol).Text
            $mob1       = $wsRoute.Cells($r, $mob1Col).Text
            $mob2       = $wsRoute.Cells($r, $mob2Col).Text
            $address    = $wsRoute.Cells($r, $addressCol).Text
            $code       = $wsRoute.Cells($r, $codeCol).Text
            $payType    = $wsRoute.Cells($r, $payTypeCol).Text
            $invoiceNo  = $wsRoute.Cells($r, $invoiceNoCol).Text
            $typeVal    = $wsRoute.Cells($r, $typeCol).Text.Trim()
            $orderTypeVal = if ($orderTypeCol) { $wsRoute.Cells($r, $orderTypeCol).Text.Trim() } else { "" }
            
            $cxKey = $cxName.Trim().ToLower()
            $addrKey = $address.Trim().ToLower()
            $isDuplicate = (($cxKey -ne "" -and $nameCounts[$cxKey] -gt 1) -or ($addrKey -ne "" -and $addrCounts[$addrKey] -gt 1))
            
            # Copy background color from source platform cell
            $platCell = $wsRoute.Cells($r, $platformCol)
            $sourceColor = $platCell.Interior.Color
            $sourceColorIdx = $platCell.Interior.ColorIndex
            
            # Check if it is a service part or service route or is present in Service sheet or has blue fill
            $platLower = $platform.ToLower().Trim()
            $isServicePart = ($platLower -ne "" -and $platLower -ne "wakefit" -and $platLower -ne "wakefit_retail" -and $platLower -ne "amazon" -and $platLower -ne "flipkart" -and $platLower -ne "offline")
            $isServiceRoute = ($routeVal -like "*Ser*" -or $routeVal -like "*Service*")
            $isServiceFromSheet = ($serviceSheetRows.ContainsKey($docket.ToUpper()) -and $serviceSheetRows[$docket.ToUpper()].IsBlue)
            $isBlueColor = ($sourceColor -eq 15773696 -or $sourceColorIdx -eq 33 -or $sourceColorIdx -eq 34 -or $sourceColorIdx -eq 41 -or $sourceColorIdx -eq 42 -or ($sourceColor -eq 16773632) -or ($sourceColor -eq 16777164))
            $isService = ($isServicePart -or $isServiceRoute -or $isServiceFromSheet -or $isBlueColor)
            
            $isReplacement = ($orderTypeVal -eq "Replacement")
            
            if ($isService) {
                $typeVal = "SER"
            } else {
                $docketKey = $docket.Trim().ToUpper()
                if ($picklistMap.Count -gt 0 -and $picklistMap.ContainsKey($docketKey)) {
                    $typeVal = $picklistMap[$docketKey]
                } else {
                    if ($typeVal -eq "") {
                        $typeVal = "FG01"
                    }
                }
            }
            
            # Get the resolved Trip Id for this route
            $tripId = $routeTripMap[$routeVal]

            # 1. Populate ROUTE sheet
            # Columns: Platform, docket_number, item_sku, invoice_number, Route, Type, Remarks, Bags, Roll, Trip Id
            $wsRouteOut.Cells($targetRouteRow, 1).Value2 = $platform
            $wsRouteOut.Cells($targetRouteRow, 2).Value2 = $docket
            $wsRouteOut.Cells($targetRouteRow, 3).Value2 = $desc
            $wsRouteOut.Cells($targetRouteRow, 4).Value2 = $invoiceNo
            $wsRouteOut.Cells($targetRouteRow, 5).Value2 = $routeVal
            $wsRouteOut.Cells($targetRouteRow, 6).Value2 = $typeVal
            $wsRouteOut.Cells($targetRouteRow, 7).Value2 = ""
            $wsRouteOut.Cells($targetRouteRow, 8).Value2 = ""
            $wsRouteOut.Cells($targetRouteRow, 9).Value2 = ""
            $wsRouteOut.Cells($targetRouteRow, 10).Value2 = $tripId
            
            $wsRouteOut.Rows.Item($targetRouteRow).RowHeight = 14.5
            
            $isSerType = ($typeVal.ToUpper() -eq "SER")
            $isRtpType = ($typeVal.ToUpper() -like "RTP*")
            
            for ($c = 1; $c -le 10; $c++) {
                $cell = $wsRouteOut.Cells($targetRouteRow, $c)
                $cell.Font.Name = "Calibri"
                $cell.Font.Size = 11
                $cell.Font.Bold = ($c -le 6 -and $isRtpType)
                $cell.Font.Color = 0
                $cell.HorizontalAlignment = -4108 # xlCenter
                $cell.VerticalAlignment = -4108   # xlCenter
                
                # Apply cell border
                $cell.Borders.LineStyle = 1       # xlContinuous
                $cell.Borders.Weight = 2          # xlThin
                $cell.Borders.Color = 0
                
                if ($c -le 5) {
                    if ($isSerType) {
                        $cell.Interior.Color = 0xF0B000 # Sky Blue hex 00B0F0 in BGR is F0B000
                    } elseif ($isReplacement) {
                        $cell.Interior.Color = 5296274 # Light Green hex 92D050 in BGR is 5296274 (0x50D092)
                    } elseif ($sourceColorIdx -ne -4142 -and $sourceColor -ne 16777215) {
                        if ($sourceColorIdx -gt 0) {
                            $cell.Interior.ColorIndex = $sourceColorIdx
                        } else {
                            $cell.Interior.Color = $sourceColor
                        }
                    } else {
                        $cell.Interior.ColorIndex = -4142
                    }
                } else {
                    $cell.Interior.ColorIndex = -4142
                }
            }
            $targetRouteRow++
            
            # Map docket to route for SER-PARTS sheet population later
            $docketKey = $docket.Trim().ToUpper()
            if ($docketKey -ne "" -and $routeVal -ne "" -and $routeVal -ne "Not planned") {
                $docketRouteMap[$docketKey] = $routeVal
                
                # If this row is a service row, populate SER-PARTS sheet
                if ($isService -and -not $addedSerDockets.ContainsKey($docketKey)) {
                    $serviceValSer = ""
                    $skuValSer = $desc
                    $invoiceValSer = $invoiceNo
                    if ($null -ne $wsServiceSrc -and $serviceSheetRows.ContainsKey($docketKey)) {
                        $sRowInfo = $serviceSheetRows[$docketKey]
                        $serviceValSer = $sRowInfo.Service
                        if ($sRowInfo.ItemSku -ne "") { $skuValSer = $sRowInfo.ItemSku }
                        if ($sRowInfo.InvoiceNumber -ne "") { $invoiceValSer = $sRowInfo.InvoiceNumber }
                    }
                    
                    $wsSerParts.Cells($targetSerRow, 1).Value2 = $serviceValSer
                    $wsSerParts.Cells($targetSerRow, 2).Value2 = $docket
                    $wsSerParts.Cells($targetSerRow, 3).Value2 = $skuValSer
                    $wsSerParts.Cells($targetSerRow, 4).Value2 = $invoiceValSer
                    $wsSerParts.Cells($targetSerRow, 5).Value2 = $routeVal
                    
                    $wsSerParts.Rows.Item($targetSerRow).RowHeight = 14.5
                    for ($c = 1; $c -le 5; $c++) {
                        $cell = $wsSerParts.Cells($targetSerRow, $c)
                        $cell.Font.Name = "Calibri"
                        $cell.Font.Size = 11
                        $cell.Font.Bold = $false
                        $cell.Font.Color = 0
                        $cell.HorizontalAlignment = -4108
                        $cell.VerticalAlignment = -4108
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.Weight = 2
                        $cell.Borders.Color = 0
                        $cell.Interior.ColorIndex = -4142
                    }
                    $targetSerRow++
                    $addedSerDockets[$docketKey] = $true
                }
            }

            # 3. Always populate INVOICE
            # Columns: PLATFORM, DOCKET NO, DESC, CUST NAME, MOB NO, MOB NO, ADDRESS, CODE, PAY TYPE, INVOICE NO, ROUTE
            $wsInv.Cells($targetInvRow, 1).Value2 = $platform
            $wsInv.Cells($targetInvRow, 2).Value2 = $docket
            $wsInv.Cells($targetInvRow, 3).Value2 = $desc
            $wsInv.Cells($targetInvRow, 4).Value2 = $cxName
            $wsInv.Cells($targetInvRow, 5).Value2 = $mob1
            $wsInv.Cells($targetInvRow, 6).Value2 = $mob2
            $wsInv.Cells($targetInvRow, 7).Value2 = $address
            $wsInv.Cells($targetInvRow, 8).Value2 = $code
            $wsInv.Cells($targetInvRow, 9).Value2 = $payType
            $wsInv.Cells($targetInvRow, 10).Value2 = $invoiceNo
            $wsInv.Cells($targetInvRow, 11).Value2 = $routeVal
            
            $wsInv.Rows.Item($targetInvRow).RowHeight = 219
            
            for ($c = 1; $c -le 11; $c++) {
                $cell = $wsInv.Cells($targetInvRow, $c)
                $cell.Font.Name = "Calibri"
                $cell.Font.Size = 30
                $cell.Font.Bold = $true
                $cell.Font.Color = 0
                $cell.HorizontalAlignment = -4108
                $cell.VerticalAlignment = -4108
                $cell.WrapText = $true
                
                $cell.Borders.LineStyle = 1
                $cell.Borders.Weight = 2
                $cell.Borders.Color = 0
                
                if ($isService) {
                    $cell.Interior.Color = 0xF0B000 # Sky Blue
                } else {
                    if ($c -eq 7 -and $isDuplicate) {
                        $cell.Interior.Color = 14212090 # Pink
                    } elseif ($sourceColorIdx -ne -4142 -and $sourceColor -ne 16777215) {
                        if ($sourceColorIdx -gt 0) {
                            $cell.Interior.ColorIndex = $sourceColorIdx
                        } else {
                            $cell.Interior.Color = $sourceColor
                        }
                    } else {
                        $cell.Interior.ColorIndex = -4142
                    }
                }
            }
            $targetInvRow++
        }
        
        # Populate TRIP ID sheet — one row per unique route, sorted numerically then alphabetically
        $sortedRoutes = $routeTripMap.Keys | Sort-Object {
            $n = 0.0
            if ([double]::TryParse($_, [ref]$n)) { $n } else { [double]::MaxValue }
        }, { $_ }
        foreach ($route in $sortedRoutes) {
            $tripNum = $routeTripMap[$route]
            $wsTrip.Cells($targetTripRow, 1).Value2 = $tripNum
            $wsTrip.Cells($targetTripRow, 2).Value2 = $route
            $wsTrip.Rows.Item($targetTripRow).RowHeight = 14.5
            for ($c = 1; $c -le 6; $c++) {
                $cell = $wsTrip.Cells($targetTripRow, $c)
                $cell.Font.Name   = "Calibri"
                $cell.Font.Size   = 11
                $cell.Font.Bold   = $false
                $cell.Font.Color  = 0
                $cell.HorizontalAlignment = -4108
                $cell.VerticalAlignment   = -4108
                $cell.Borders.LineStyle   = 1
                $cell.Borders.Weight      = 2
                $cell.Borders.Color       = 0
                $cell.Interior.ColorIndex = -4142
            }
            $targetTripRow++
        }
        Write-Host "TRIP ID sheet populated with $(($targetTripRow - 2)) unique routes."

        # Second pass: Process missing blue service dockets from the Service sheet
        if ($null -ne $wsServiceSrc) {
            for ($sr = 2; $sr -le $sRows; $sr++) {
                $docketVal = $wsServiceSrc.Cells($sr, $sDocketColIdx).Text.Trim()
                $docketKey = $docketVal.ToUpper()
                
                if ($docketKey -ne "") {
                    # Check if it was already processed in the first pass
                    if ($addedSerDockets.ContainsKey($docketKey)) {
                        continue
                    }
                    
                    # Check if it is blue in the Service sheet
                    $sRowInfo = $serviceSheetRows[$docketKey]
                    if ($null -eq $sRowInfo -or -not $sRowInfo.IsBlue) {
                        continue
                    }
                    
                    # Determine route
                    $orderId = $sRowInfo.OrderRandomId.ToUpper()
                    $cxName = $sRowInfo.CxName.ToLower()
                    $addr = $sRowInfo.Address.ToLower()
                    
                    $routeVal = $sRowInfo.Route
                    if ($routeVal -eq "") {
                        if ($orderId -ne "" -and $routeByOrderId.ContainsKey($orderId)) {
                            $routeVal = $routeByOrderId[$orderId]
                        } elseif ($routeByDocket.ContainsKey($docketKey)) {
                            $routeVal = $routeByDocket[$docketKey]
                        } elseif ($cxName -ne "" -and $addr -ne "" -and $routeByCust.ContainsKey("$cxName|$addr")) {
                            $routeVal = $routeByCust["$cxName|$addr"]
                        }
                    }
                    
                    # Extract values
                    $platform = $sRowInfo.Service
                    if ($platform -eq "") { $platform = $sRowInfo.Type }
                    $desc = $sRowInfo.ItemSku
                    $invoiceNo = $sRowInfo.InvoiceNumber
                    $typeVal = "SER"
                    
                    # 1. Always Populate SER-PARTS sheet
                    $wsSerParts.Cells($targetSerRow, 1).Value2 = $sRowInfo.Service
                    $wsSerParts.Cells($targetSerRow, 2).Value2 = $docketVal
                    $wsSerParts.Cells($targetSerRow, 3).Value2 = $desc
                    $wsSerParts.Cells($targetSerRow, 4).Value2 = $invoiceNo
                    $wsSerParts.Cells($targetSerRow, 5).Value2 = $routeVal
                    
                    $wsSerParts.Rows.Item($targetSerRow).RowHeight = 14.5
                    for ($c = 1; $c -le 5; $c++) {
                        $cell = $wsSerParts.Cells($targetSerRow, $c)
                        $cell.Font.Name = "Calibri"
                        $cell.Font.Size = 11
                        $cell.Font.Bold = $false
                        $cell.Font.Color = 0
                        $cell.HorizontalAlignment = -4108
                        $cell.VerticalAlignment = -4108
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.Weight = 2
                        $cell.Borders.Color = 0
                        $cell.Interior.ColorIndex = -4142
                    }
                    $targetSerRow++
                    $addedSerDockets[$docketKey] = $true

                    # 2. Only Populate ROUTE and INVOICE sheets if route is valid
                    if ($routeVal -ne "" -and $routeVal -ne "Not planned") {
                        # Populate ROUTE sheet
                        $wsRouteOut.Cells($targetRouteRow, 1).Value2 = $platform
                        $wsRouteOut.Cells($targetRouteRow, 2).Value2 = $docketVal
                        $wsRouteOut.Cells($targetRouteRow, 3).Value2 = $desc
                        $wsRouteOut.Cells($targetRouteRow, 4).Value2 = $invoiceNo
                        $wsRouteOut.Cells($targetRouteRow, 5).Value2 = $routeVal
                        $wsRouteOut.Cells($targetRouteRow, 6).Value2 = $typeVal
                        $wsRouteOut.Cells($targetRouteRow, 7).Value2 = ""
                        $wsRouteOut.Cells($targetRouteRow, 8).Value2 = ""
                        $wsRouteOut.Cells($targetRouteRow, 9).Value2 = ""
                        
                        $tripId = ""
                        if ($routeTripMap.ContainsKey($routeVal)) {
                            $tripId = $routeTripMap[$routeVal]
                        }
                        $wsRouteOut.Cells($targetRouteRow, 10).Value2 = $tripId
                        
                        $wsRouteOut.Rows.Item($targetRouteRow).RowHeight = 14.5
                        for ($c = 1; $c -le 10; $c++) {
                            $cell = $wsRouteOut.Cells($targetRouteRow, $c)
                            $cell.Font.Name = "Calibri"
                            $cell.Font.Size = 11
                            $cell.Font.Bold = $false
                            $cell.Font.Color = 0
                            $cell.HorizontalAlignment = -4108
                            $cell.VerticalAlignment = -4108
                            $cell.Borders.LineStyle = 1
                            $cell.Borders.Weight = 2
                            $cell.Borders.Color = 0
                            
                            if ($c -le 5) {
                                $cell.Interior.Color = 0xF0B000 # Sky Blue
                            } else {
                                $cell.Interior.ColorIndex = -4142
                            }
                        }
                        $targetRouteRow++
                        
                        # Populate INVOICE sheet
                        $wsInv.Cells($targetInvRow, 1).Value2 = $platform
                        $wsInv.Cells($targetInvRow, 2).Value2 = $docketVal
                        $wsInv.Cells($targetInvRow, 3).Value2 = $desc
                        $wsInv.Cells($targetInvRow, 4).Value2 = $sRowInfo.CxName
                        $wsInv.Cells($targetInvRow, 5).Value2 = $sRowInfo.MobileNumber
                        $wsInv.Cells($targetInvRow, 6).Value2 = $sRowInfo.AlternateMobileNumber
                        $wsInv.Cells($targetInvRow, 7).Value2 = $sRowInfo.Address
                        $wsInv.Cells($targetInvRow, 8).Value2 = $sRowInfo.Pincode
                        $wsInv.Cells($targetInvRow, 9).Value2 = "" # payType
                        $wsInv.Cells($targetInvRow, 10).Value2 = $invoiceNo
                        $wsInv.Cells($targetInvRow, 11).Value2 = $routeVal
                        
                        $wsInv.Rows.Item($targetInvRow).RowHeight = 219
                        for ($c = 1; $c -le 11; $c++) {
                            $cell = $wsInv.Cells($targetInvRow, $c)
                            $cell.Font.Name = "Calibri"
                            $cell.Font.Size = 30
                            $cell.Font.Bold = $true
                            $cell.Font.Color = 0
                            $cell.HorizontalAlignment = -4108
                            $cell.VerticalAlignment = -4108
                            $cell.WrapText = $true
                            $cell.Borders.LineStyle = 1
                            $cell.Borders.Weight = 2
                            $cell.Borders.Color = 0
                            $cell.Interior.Color = 0xF0B000 # Sky Blue
                        }
                        $targetInvRow++
                    }
                }
            }
        }
        
        # Ensure gridlines are visible on all four new sheets
        $wsInv.Activate()
        $excel.ActiveWindow.DisplayGridlines = $true
        $wsSerParts.Activate()
        $excel.ActiveWindow.DisplayGridlines = $true
        $wsTrip.Activate()
        $excel.ActiveWindow.DisplayGridlines = $true
        $wsRouteOut.Activate()
        $excel.ActiveWindow.DisplayGridlines = $true

        Write-Host "Successfully generated sheets (INVOICE: $(($targetInvRow - 2)) rows, SER-PARTS: $(($targetSerRow - 2)) rows, ROUTE: $(($targetRouteRow - 2)) rows)."
        
        # Save workbook
        Write-Host "Saving workbook..."
        $wb.Save()
        $wb.Close($false)
        Write-Host "Conversion completed successfully: data 1.xlsx has been created/updated."
        
    } else {
        # Converting data 1 -> data 2
        Write-Host "Copying data 1.xlsx to data 2.xlsx..."
        Copy-Item $data1Path $data2Path -Force
        
        Write-Host "Loading data 2.xlsx..."
        $wb = $excel.Workbooks.Open($data2Path)
        
        try {
            $ws = $wb.Sheets.Item("INVOICE")
            $ws.Delete()
            Write-Host "Removed INVOICE sheet."
        } catch {
            Write-Host "INVOICE sheet was not found in data 2.xlsx."
        }
        try {
            $ws = $wb.Sheets.Item("TRIP ID")
            $ws.Delete()
            Write-Host "Removed TRIP ID sheet."
        } catch {
            Write-Host "TRIP ID sheet was not found in data 2.xlsx."
        }
        try {
            $ws = $wb.Sheets.Item("SER-PARTS")
            $ws.Delete()
            Write-Host "Removed SER-PARTS sheet."
        } catch {
            Write-Host "SER-PARTS sheet was not found in data 2.xlsx."
        }
        try {
            $ws = $wb.Sheets.Item("ROUTE")
            $ws.Delete()
            Write-Host "Removed ROUTE sheet."
        } catch {
            Write-Host "ROUTE sheet was not found in data 2.xlsx."
        }
        
        # Save workbook
        Write-Host "Saving workbook..."
        $wb.Save()
        $wb.Close($false)
        Write-Host "Conversion completed successfully: data 2.xlsx has been created/updated."
    }
} catch {
    Write-Host "`n[ERROR] An error occurred during conversion:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    Write-Error $_
} finally {
    # Cleanup Excel Interop
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    # Clear variables to release handles
    Remove-Variable excel -ErrorAction SilentlyContinue
}

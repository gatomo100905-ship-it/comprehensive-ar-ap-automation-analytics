# =========================================================================
# FILE: create_excel_form.ps1
# DESCRIPTION: Khoi tao Excel COM Object, thiet ke giao dien Form
#              nhap lieu, luu vao thu muc Public va di chuyen ve thu muc du an.
# =========================================================================

$projectPath = "C:\Users\Hoàng\.gemini\antigravity\scratch\comprehensive-ar-ap-automation-analytics"
$xlsmPath = "$projectPath\ArAp_Manager.xlsm"
$tempPath = "C:\Users\Public\ArAp_Manager.xlsm"
$basPath = "$projectPath\vba\VBA_Invoice_Form.bas"

Write-Output "Starting Excel automation to build the form..."

# 1. Khai bao doi tuong Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # 2. Them workbook moi
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Sheets.Item(1)
    $ws.Name = "Invoice_Form"
    $excel.ActiveWindow.DisplayGridlines = $true

    # 3. Viet tieu de
    $ws.Range("A1:B1").Merge()
    $ws.Range("A1").Value = "NHAP HOA DON CONG NO (AR/AP)"
    $ws.Range("A1").Font.Name = "Segoe UI"
    $ws.Range("A1").Font.Size = 14
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1").Font.Color = 0xFFFFFF
    $ws.Range("A1").Interior.Color = 0x8C4B1B # Teal
    $ws.Range("A1").HorizontalAlignment = -4108
    $ws.Range("A1").RowHeight = 35

    # 4. Viet Nhan va Gia tri mau
    $ws.Range("A3").Value = "Loai hoa don * (AR/AP)"
    $ws.Range("B3").Value = "AR"
    
    $ws.Range("A4").Value = "So hoa don *"
    $ws.Range("B4").Value = "INV-26-0007"
    
    $ws.Range("A5").Value = "Ma doi tac *"
    $ws.Range("B5").Value = "CUST-001"
    
    $ws.Range("A6").Value = "Ngay hoa don * (YYYY-MM-DD)"
    $ws.Range("B6").Value = "2026-07-27"
    
    $ws.Range("A7").Value = "Ngay den han * (YYYY-MM-DD)"
    $ws.Range("B7").Value = "2026-08-27"
    
    $ws.Range("A8").Value = "So tien truoc thue *"
    $ws.Range("B8").Value = 15000000
    
    $ws.Range("A9").Value = "Thue suat *"
    $ws.Range("B9").Value = 0.10
    
    $ws.Range("A10").Value = "Noi dung dien giai"
    $ws.Range("B10").Value = "Ban tra sua si dot cuoi thang 7"

    # 5. Dinh dang do rong cot va chieu cao
    $ws.Columns.Item(1).ColumnWidth = 30
    $ws.Columns.Item(2).ColumnWidth = 35
    $ws.Range("A3:A10").RowHeight = 25

    # 6. Trang tri Border
    $inputRange = $ws.Range("A3:B10")
    $inputRange.Font.Name = "Segoe UI"
    $inputRange.Font.Size = 11
    $ws.Range("A3:A10").Font.Bold = $true
    $ws.Range("A3:A10").Interior.Color = 0xF2F2F2
    $inputRange.HorizontalAlignment = -4131
    $inputRange.Borders.LineStyle = 1
    $inputRange.Borders.Weight = 2
    $inputRange.Borders.Color = 0xD3D3D3

    # 7. Dinh dang so
    $ws.Range("B8").NumberFormat = "#,##0"
    $ws.Range("B9").NumberFormat = "0.0%"

    # 8. Import VBA (Neu duoc cap quyen)
    try {
        $wb.VBProject.VBComponents.Import($basPath) > $null
        Write-Output "VBA Module imported."
    }
    catch {
        Write-Output "VBA auto-import skipped."
    }

    # 9. Ve nut nhan
    $btn = $ws.Buttons().Add(430, 45, 160, 40)
    $btn.Text = "LUU HOA DON"
    $btn.OnAction = "SaveInvoice"
    $btn.Font.Name = "Segoe UI"
    $btn.Font.Bold = $true

    # 10. Luu vao thu muc temp (khong chua ki tu Unicode)
    if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
    $wb.SaveAs($tempPath, 52)
    Write-Output "Temporary Excel workbook saved."
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) > $null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) > $null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) > $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# 11. Di chuyen file ve thu muc du an bang PowerShell (ho tro tot Unicode)
if (Test-Path $tempPath) {
    if (Test-Path $xlsmPath) { Remove-Item $xlsmPath -Force }
    Move-Item -Path $tempPath -Destination $xlsmPath -Force
    Write-Output "Excel workbook successfully moved to project folder: $xlsmPath"
} else {
    Write-Error "Failed to locate temporary Excel file."
}

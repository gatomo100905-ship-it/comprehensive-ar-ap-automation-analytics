# =========================================================================
# FILE: verify_db.ps1
# DESCRIPTION: Kiểm thử hoạt động của CSDL (Views và Stored Procedure Netting).
# =========================================================================

$server = "localhost"
$database = "AR_AP_Analytics"
$connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True;"

function Run-Query {
    param([string]$sql)
    $conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $dt = New-Object System.Data.DataTable
    $da.Fill($dt) > $null
    $conn.Close()
    return $dt
}

Write-Output "=== KIỂM THỬ 1: TRUY VẤN BÁO CÁO TUỔI NỢ PHẢI THU (AR AGING) ==="
$arAging = Run-Query "SELECT InvoiceNumber, CustomerName, RemainingAmount, OverdueDays, AgingBucket FROM dbo.v_AR_Aging"
$arAging | Format-Table -AutoSize

Write-Output "=== KIỂM THỬ 2: TRUY VẤN BÁO CÁO TUỔI NỢ PHẢI TRẢ (AP AGING) ==="
$apAging = Run-Query "SELECT InvoiceNumber, SupplierName, RemainingAmount, OverdueDays, AgingBucket FROM dbo.v_AP_Aging"
$apAging | Format-Table -AutoSize

Write-Output "=== KIỂM THỬ 3: THỰC THI BÙ TRỪ CÔNG NỢ (NETTING) CHO ĐỐI TÁC IDP ==="
Write-Output "AR (Phải thu) truoc netting:"
$beforeAR = Run-Query "SELECT InvoiceNumber, RemainingAmount, Status FROM dbo.v_AR_Aging WHERE CustomerCode = 'CUST-005'"
$beforeAR | Format-Table -AutoSize

Write-Output "AP (Phải trả) truoc netting:"
$beforeAP = Run-Query "SELECT InvoiceNumber, RemainingAmount, Status FROM dbo.v_AP_Aging WHERE SupplierCode = 'SUPP-005'"
$beforeAP | Format-Table -AutoSize

Write-Output "Chay Stored Procedure sp_ProcessNetting..."
$conn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$cmd = $conn.CreateCommand()
$cmd.CommandText = "EXEC dbo.sp_ProcessNetting @CustomerCode = 'CUST-005', @SupplierCode = 'SUPP-005'"
$conn.Open()
$cmd.ExecuteNonQuery() > $null
$conn.Close()

Write-Output "Cấn trừ thành công. Kiểm tra kết quả sau cấn trừ:"

Write-Output "AR (Phải thu) sau netting:"
$afterAR = Run-Query "SELECT InvoiceNumber, RemainingAmount, Status FROM dbo.v_AR_Aging WHERE CustomerCode = 'CUST-005'"
$afterAR | Format-Table -AutoSize

Write-Output "AP (Phải trả) sau netting:"
$afterAP = Run-Query "SELECT InvoiceNumber, RemainingAmount, Status FROM dbo.v_AP_Aging WHERE SupplierCode = 'SUPP-005'"
$afterAP | Format-Table -AutoSize

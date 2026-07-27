# =========================================================================
# FILE: run_sql_scripts.ps1
# DESCRIPTION: Đọc và thực thi các file SQL trên SQL Server bằng .NET SqlClient.
#              Hỗ trợ phân tách lệnh bằng từ khóa 'GO'.
# =========================================================================

$server = "localhost"
$database = "master" # Ban đầu kết nối vào master để chạy script tạo DB
$connectionString = "Server=$server;Database=$database;Integrated Security=True;TrustServerCertificate=True;"

$sqlFiles = @(
    "sql\schema.sql",
    "sql\sample_data.sql",
    "sql\views_and_procedures.sql"
)

function Execute-SqlScript {
    param (
        [string]$filePath,
        [string]$connString
    )

    Write-Output "Executing file: $filePath"
    
    if (-not (Test-Path $filePath)) {
        Write-Error "File not found: $filePath"
        return
    }

    $content = Get-Content -Path $filePath -Raw
    
    # Phân tách file SQL thành các Batch độc lập dựa trên từ khóa GO nằm riêng lẻ trên một dòng
    # Sử dụng Regex để tách nhằm đảm bảo độ chính xác
    $batches = [regex]::Split($content, "(?m)^\s*GO\s*\r?\n?")

    $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
    try {
        $conn.Open()
        
        foreach ($batch in $batches) {
            $trimmedBatch = $batch.Trim()
            if ($trimmedBatch -eq "") { continue }

            # Nếu batch chứa lệnh USE AR_AP_Analytics, ta thay đổi Database của connection
            if ($trimmedBatch -match "(?i)^\s*USE\s+AR_AP_Analytics") {
                # Thay đổi CSDL kết nối sau khi tạo DB
                $conn.ChangeDatabase("AR_AP_Analytics")
                Write-Output "  -> Switch Connection Database to AR_AP_Analytics"
                continue
            }

            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $trimmedBatch
            $cmd.ExecuteNonQuery() > $null
        }
        Write-Output "  -> SUCCESS"
    }
    catch {
        Write-Error "  -> FAILED at Batch: $_"
        Write-Error "  -> Error Details: $($_.Exception.Message)"
        throw $_
    }
    finally {
        if ($conn.State -eq [System.Data.ConnectionState]::Open) {
            $conn.Close()
        }
    }
}

try {
    foreach ($file in $sqlFiles) {
        # Chạy từng file script
        Execute-SqlScript -filePath $file -connString $connectionString
    }
    Write-Output "`n=== ALL SQL SCRIPTS EXECUTED SUCCESSFULLY ==="
}
catch {
    Write-Error "`n=== EXECUTION FAILED ==="
    exit 1
}

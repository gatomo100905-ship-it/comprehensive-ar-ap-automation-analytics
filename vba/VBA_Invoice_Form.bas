Attribute VB_Name = "VBA_Invoice_Form"
' =========================================================================
' MODULE: VBA_Invoice_Form
' DESCRIPTION: Lập trình kết nối ADODB và kiểm soát dữ liệu nhập hóa đơn
'              từ Excel đẩy trực tiếp vào SQL Server cục bộ.
' =========================================================================

Public Sub SaveInvoice()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Invoice_Form")
    
    ' 1. Đọc dữ liệu từ các ô trên Form
    Dim invType As String, invNo As String, partnerCode As String
    Dim invDate As Variant, dueDate As Variant
    Dim amount As Double, taxRate As Double, desc As String
    
    invType = Trim(ws.Range("B3").Value)
    invNo = Trim(ws.Range("B4").Value)
    partnerCode = Trim(ws.Range("B5").Value)
    invDate = ws.Range("B6").Value
    dueDate = ws.Range("B7").Value
    
    ' 2. BẮT ĐẦU KIỂM TRA DỮ LIỆU ĐẦU VÀO (VALIDATION)
    
    ' Kiểm tra rỗng các trường bắt buộc
    If invType = "" Or invNo = "" Or partnerCode = "" Or IsEmpty(invDate) Or IsEmpty(dueDate) Or IsEmpty(ws.Range("B8").Value) Or IsEmpty(ws.Range("B9").Value) Then
        MsgBox "Loi: Vui long dien day du tat ca cac truong bat buoc (co dau *).", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    
    ' Kiểm tra loại hóa đơn hợp lệ
    If invType <> "AR" And invType <> "AP" Then
        MsgBox "Loi: Loai hoa don phai la 'AR' (Phai thu) hoac 'AP' (Phai tra).", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    
    ' Kiểm tra định dạng số tiền
    If Not IsNumeric(ws.Range("B8").Value) Then
        MsgBox "Loi: So tien truoc thue phai la kieu so.", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    amount = CDbl(ws.Range("B8").Value)
    If amount <= 0 Then
        MsgBox "Loi: So tien hoa don phai lon hon 0.", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    
    ' Kiểm tra định dạng thuế suất
    If Not IsNumeric(ws.Range("B9").Value) Then
        MsgBox "Loi: Thue suat phai la kieu so (vi du: 0.1 cho 10%).", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    taxRate = CDbl(ws.Range("B9").Value)
    If taxRate < 0 Then
        MsgBox "Loi: Thue suat khong duoc am.", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    
    ' Kiểm tra định dạng ngày tháng
    If Not IsDate(invDate) Then
        MsgBox "Loi: Ngay hoa don khong dung dinh dang ngay thang.", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    If Not IsDate(dueDate) Then
        MsgBox "Loi: Ngay den han khong dung dinh dang ngay thang.", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    
    ' Logic ngày đến hạn phải >= ngày hóa đơn
    If CDate(dueDate) < CDate(invDate) Then
        MsgBox "Loi: Ngay den han phai lon hon hoac bang Ngay hoa don.", vbCritical, "Kiem Tra Du Lieu"
        Exit Sub
    End If
    
    desc = Trim(ws.Range("B10").Value)
    
    ' 3. KHỞI TẠO KẾT NỐI DATABASE (ADODB)
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connStr As String
    
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' Chuỗi kết nối Windows Authentication kết nối trực tiếp vào SQL Server cục bộ
    connStr = "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=AR_AP_Analytics;Integrated Security=SSPI;Trust Server Certificate=True;"
    
    On Error GoTo ErrorHandler
    conn.Open connStr
    
    Dim sqlQuery As String
    Dim partnerID As Long
    
    ' 4. KIỂM TRA ĐỐI TÁC CÓ TỒN TẠI TRONG CSDL KHÔNG & LẤY ID
    If invType = "AR" Then
        sqlQuery = "SELECT CustomerID FROM dbo.DimCustomers WHERE CustomerCode = '" & Replace(partnerCode, "'", "''") & "' AND IsDeleted = 0;"
    Else
        sqlQuery = "SELECT SupplierID FROM dbo.DimSuppliers WHERE SupplierCode = '" & Replace(partnerCode, "'", "''") & "' AND IsDeleted = 0;"
    End If
    
    rs.Open sqlQuery, conn, adOpenStatic, adLockReadOnly
    
    If rs.EOF Then
        MsgBox "Loi: Ma doi tac '" & partnerCode & "' khong ton tai trong danh muc " & IIf(invType = "AR", "Khach hang", "Nha cung cap") & " tren CSDL SQL Server.", vbCritical, "Kiem Tra Doi Tac"
        rs.Close
        conn.Close
        Exit Sub
    End If
    
    partnerID = rs.Fields(0).Value
    rs.Close
    
    ' 5. KIỂM TRA SỐ HÓA ĐƠN CÓ BỊ TRÙNG LẶP TRÊN HỆ THỐNG KHÔNG
    If invType = "AR" Then
        sqlQuery = "SELECT InvoiceID FROM dbo.FactARInvoices WHERE InvoiceNumber = '" & Replace(invNo, "'", "''") & "' AND IsDeleted = 0;"
    Else
        sqlQuery = "SELECT InvoiceID FROM dbo.FactAPInvoices WHERE InvoiceNumber = '" & Replace(invNo, "'", "''") & "' AND IsDeleted = 0;"
    End If
    
    rs.Open sqlQuery, conn, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "Loi: So hoa don '" & invNo & "' da ton tai tren he thong. Vui long kiem tra lai.", vbCritical, "Trung Lap Hoa Don"
        rs.Close
        conn.Close
        Exit Sub
    End If
    rs.Close
    
    ' 6. TÍNH TOÁN THUẾ VÀ TỔNG TIỀN
    Dim taxAmount As Double, totalAmount As Double
    taxAmount = amount * taxRate
    totalAmount = amount + taxAmount
    
    ' Định dạng ngày theo chuẩn YYYY-MM-DD để đưa vào SQL Server
    Dim strInvDate As String, strDueDate As String
    strInvDate = Format(invDate, "YYYY-MM-DD")
    strDueDate = Format(dueDate, "YYYY-MM-DD")
    
    ' 7. THỰC THI LỆNH INSERT VÀO CSDL
    Dim insertSQL As String
    If invType = "AR" Then
        insertSQL = "INSERT INTO dbo.FactARInvoices (InvoiceNumber, CustomerID, InvoiceDate, DueDate, Amount, TaxAmount, TotalAmount, Description, Status) " & _
                    "VALUES ('" & Replace(invNo, "'", "''") & "', " & partnerID & ", '" & strInvDate & "', '" & strDueDate & "', " & _
                    amount & ", " & taxAmount & ", " & totalAmount & ", N'" & Replace(desc, "'", "''") & "', 'Unpaid');"
    Else
        insertSQL = "INSERT INTO dbo.FactAPInvoices (InvoiceNumber, SupplierID, InvoiceDate, DueDate, Amount, TaxAmount, TotalAmount, Description, Status) " & _
                    "VALUES ('" & Replace(invNo, "'", "''") & "', " & partnerID & ", '" & strInvDate & "', '" & strDueDate & "', " & _
                    amount & ", " & taxAmount & ", " & totalAmount & ", N'" & Replace(desc, "'", "''") & "', 'Unpaid');"
    End If
    
    conn.Execute insertSQL
    
    ' 8. THÀNH CÔNG: DỌN DẸP FORM & ĐÓNG KẾT NỐI
    conn.Close
    
    ' Xóa trắng form (giữ lại ô loại hóa đơn B3)
    ws.Range("B4:B10").ClearContents
    
    MsgBox "Thanh cong: Da luu hoa don '" & invNo & "' vao CSDL SQL Server!", vbInformation, "Thanh Cong"
    Exit Sub

ErrorHandler:
    MsgBox "Loi he thong: Khong the ket noi hoac ghi du lieu vao SQL Server." & vbNewLine & "Chi tiet loi: " & Err.Description, vbCritical, "Loi Ket Noi CSDL"
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
End Sub

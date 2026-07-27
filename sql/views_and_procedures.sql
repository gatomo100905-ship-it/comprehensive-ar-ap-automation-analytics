-- =========================================================================
-- FILE: views_and_procedures.sql
-- DESCRIPTION: Chứa định nghĩa các Views tính tuổi nợ và Stored Procedures
--              xử lý đối chiếu, cấn trừ công nợ tự động.
-- =========================================================================

USE AR_AP_Analytics;
GO

-- Xóa Views cũ nếu tồn tại
IF OBJECT_ID('dbo.v_AR_Aging', 'V') IS NOT NULL DROP VIEW dbo.v_AR_Aging;
IF OBJECT_ID('dbo.v_AP_Aging', 'V') IS NOT NULL DROP VIEW dbo.v_AP_Aging;
GO

-- Xóa Procedures cũ nếu tồn tại
IF OBJECT_ID('dbo.sp_GetCustomerAgingReport', 'P') IS NOT NULL DROP PROCEDURE dbo.sp_GetCustomerAgingReport;
IF OBJECT_ID('dbo.sp_ProcessNetting', 'P') IS NOT NULL DROP PROCEDURE dbo.sp_ProcessNetting;
GO

-- =========================================================================
-- 1. VIEW TÍNH TUỔI NỢ PHẢI THU (v_AR_Aging)
-- =========================================================================
CREATE VIEW dbo.v_AR_Aging AS
SELECT 
    i.InvoiceID,
    i.InvoiceNumber,
    c.CustomerID,
    c.CustomerCode,
    c.CompanyName AS CustomerName,
    i.InvoiceDate,
    i.DueDate,
    i.TotalAmount AS InvoiceAmount,
    ISNULL(p.TotalPaid, 0) AS PaidAmount,
    (i.TotalAmount - ISNULL(p.TotalPaid, 0)) AS RemainingAmount,
    CASE 
        WHEN (i.TotalAmount - ISNULL(p.TotalPaid, 0)) <= 0 THEN 0
        WHEN GETDATE() > i.DueDate THEN DATEDIFF(day, i.DueDate, GETDATE())
        ELSE 0 
    END AS OverdueDays,
    CASE 
        WHEN (i.TotalAmount - ISNULL(p.TotalPaid, 0)) <= 0 THEN 'Paid'
        WHEN GETDATE() <= i.DueDate THEN 'Current'
        ELSE
            CASE 
                WHEN DATEDIFF(day, i.DueDate, GETDATE()) BETWEEN 1 AND 30 THEN '1-30 Days'
                WHEN DATEDIFF(day, i.DueDate, GETDATE()) BETWEEN 31 AND 60 THEN '31-60 Days'
                WHEN DATEDIFF(day, i.DueDate, GETDATE()) BETWEEN 61 AND 90 THEN '61-90 Days'
                ELSE '90+ Days'
            END
    END AS AgingBucket,
    i.Status,
    i.IsDeleted
FROM dbo.FactARInvoices i
INNER JOIN dbo.DimCustomers c ON i.CustomerID = c.CustomerID
LEFT JOIN (
    SELECT InvoiceID, SUM(AmountPaid) AS TotalPaid
    FROM dbo.FactARPayments
    WHERE IsDeleted = 0
    GROUP BY InvoiceID
) p ON i.InvoiceID = p.InvoiceID
WHERE i.IsDeleted = 0 AND c.IsDeleted = 0;
GO

-- =========================================================================
-- 2. VIEW TÍNH TUỔI NỢ PHẢI TRẢ (v_AP_Aging)
-- =========================================================================
CREATE VIEW dbo.v_AP_Aging AS
SELECT 
    i.InvoiceID,
    i.InvoiceNumber,
    s.SupplierID,
    s.SupplierCode,
    s.CompanyName AS SupplierName,
    i.InvoiceDate,
    i.DueDate,
    i.TotalAmount AS InvoiceAmount,
    ISNULL(p.TotalPaid, 0) AS PaidAmount,
    (i.TotalAmount - ISNULL(p.TotalPaid, 0)) AS RemainingAmount,
    CASE 
        WHEN (i.TotalAmount - ISNULL(p.TotalPaid, 0)) <= 0 THEN 0
        WHEN GETDATE() > i.DueDate THEN DATEDIFF(day, i.DueDate, GETDATE())
        ELSE 0 
    END AS OverdueDays,
    CASE 
        WHEN (i.TotalAmount - ISNULL(p.TotalPaid, 0)) <= 0 THEN 'Paid'
        WHEN GETDATE() <= i.DueDate THEN 'Current'
        ELSE
            CASE 
                WHEN DATEDIFF(day, i.DueDate, GETDATE()) BETWEEN 1 AND 30 THEN '1-30 Days'
                WHEN DATEDIFF(day, i.DueDate, GETDATE()) BETWEEN 31 AND 60 THEN '31-60 Days'
                WHEN DATEDIFF(day, i.DueDate, GETDATE()) BETWEEN 61 AND 90 THEN '61-90 Days'
                ELSE '90+ Days'
            END
    END AS AgingBucket,
    i.Status,
    i.IsDeleted
FROM dbo.FactAPInvoices i
INNER JOIN dbo.DimSuppliers s ON i.SupplierID = s.SupplierID
LEFT JOIN (
    SELECT InvoiceID, SUM(AmountPaid) AS TotalPaid
    FROM dbo.FactAPPayments
    WHERE IsDeleted = 0
    GROUP BY InvoiceID
) p ON i.InvoiceID = p.InvoiceID
WHERE i.IsDeleted = 0 AND s.IsDeleted = 0;
GO

-- =========================================================================
-- 3. STORED PROCEDURE XUẤT BÁO CÁO TUỔI NỢ TẠI MỘT NGÀY BẤT KỲ (LỊCH SỬ)
-- =========================================================================
CREATE PROCEDURE dbo.sp_GetCustomerAgingReport
    @CustomerID INT = NULL,
    @ReportDate DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;

    -- Nếu không truyền ngày, lấy ngày hiện tại
    IF @ReportDate IS NULL
        SET @ReportDate = CAST(GETDATE() AS DATE);

    SELECT 
        c.CustomerCode,
        c.CompanyName AS CustomerName,
        i.InvoiceNumber,
        i.InvoiceDate,
        i.DueDate,
        i.TotalAmount AS InvoiceAmount,
        ISNULL(p.PaidToDate, 0) AS PaidAmount,
        (i.TotalAmount - ISNULL(p.PaidToDate, 0)) AS RemainingAmount,
        CASE 
            WHEN (i.TotalAmount - ISNULL(p.PaidToDate, 0)) <= 0 THEN 0
            WHEN @ReportDate > i.DueDate THEN DATEDIFF(day, i.DueDate, @ReportDate)
            ELSE 0 
        END AS OverdueDays,
        CASE 
            WHEN (i.TotalAmount - ISNULL(p.PaidToDate, 0)) <= 0 THEN 'Paid'
            WHEN @ReportDate <= i.DueDate THEN 'Current'
            ELSE
                CASE 
                    WHEN DATEDIFF(day, i.DueDate, @ReportDate) BETWEEN 1 AND 30 THEN '1-30 Days'
                    WHEN DATEDIFF(day, i.DueDate, @ReportDate) BETWEEN 31 AND 60 THEN '31-60 Days'
                    WHEN DATEDIFF(day, i.DueDate, @ReportDate) BETWEEN 61 AND 90 THEN '61-90 Days'
                    ELSE '90+ Days'
                END
        END AS AgingBucket
    FROM dbo.FactARInvoices i
    INNER JOIN dbo.DimCustomers c ON i.CustomerID = c.CustomerID
    LEFT JOIN (
        -- Chỉ tính các khoản thanh toán phát sinh trước hoặc bằng ngày báo cáo
        SELECT InvoiceID, SUM(AmountPaid) AS PaidToDate
        FROM dbo.FactARPayments
        WHERE IsDeleted = 0 AND PaymentDate <= @ReportDate
        GROUP BY InvoiceID
    ) p ON i.InvoiceID = p.InvoiceID
    WHERE i.IsDeleted = 0 
      AND c.IsDeleted = 0
      AND i.InvoiceDate <= @ReportDate -- Chỉ lấy hóa đơn xuất trước ngày báo cáo
      AND (@CustomerID IS NULL OR c.CustomerID = @CustomerID)
    ORDER BY c.CustomerCode, i.InvoiceDate;
END;
GO

-- =========================================================================
-- 4. STORED PROCEDURE BÙ TRỪ CÔNG NỢ TỰ ĐỘNG (Netting) - CHUẨN FIFO
-- =========================================================================
CREATE PROCEDURE dbo.sp_ProcessNetting
    @CustomerCode VARCHAR(50),
    @SupplierCode VARCHAR(50),
    @NettingDate DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON; -- Tự động Rollback nếu có lỗi nghiêm trọng

    IF @NettingDate IS NULL
        SET @NettingDate = CAST(GETDATE() AS DATE);

    -- 1. Lấy thông tin ID và mã số thuế
    DECLARE @CustomerID INT, @SupplierID INT;
    DECLARE @CustTaxCode VARCHAR(20), @SuppTaxCode VARCHAR(20);

    SELECT @CustomerID = CustomerID, @CustTaxCode = TaxCode FROM dbo.DimCustomers WHERE CustomerCode = @CustomerCode AND IsDeleted = 0;
    SELECT @SupplierID = SupplierID, @SuppTaxCode = TaxCode FROM dbo.DimSuppliers WHERE SupplierCode = @SupplierCode AND IsDeleted = 0;

    -- 2. Kiểm tra tính hợp lệ
    IF @CustomerID IS NULL OR @SupplierID IS NULL
    BEGIN
        RAISERROR(N'Lỗi: Không tìm thấy thông tin Khách hàng hoặc Nhà cung cấp hợp lệ.', 16, 1);
        RETURN;
    END

    IF @CustTaxCode IS NULL OR @SuppTaxCode IS NULL OR @CustTaxCode <> @SuppTaxCode
    BEGIN
        RAISERROR(N'Lỗi: Hai đối tác không có cùng Mã số thuế. Không thể thực hiện bù trừ công nợ (Netting).', 16, 1);
        RETURN;
    END

    -- 3. Tính toán tổng nợ chưa thanh toán của cả 2 bên
    DECLARE @TotalUnpaidAR DECIMAL(18,2), @TotalUnpaidAP DECIMAL(18,2);

    SELECT @TotalUnpaidAR = ISNULL(SUM(RemainingAmount), 0) FROM dbo.v_AR_Aging WHERE CustomerID = @CustomerID AND AgingBucket <> 'Paid';
    SELECT @TotalUnpaidAP = ISNULL(SUM(RemainingAmount), 0) FROM dbo.v_AP_Aging WHERE SupplierID = @SupplierID AND AgingBucket <> 'Paid';

    -- Số tiền cấn trừ tối đa là số tiền nhỏ hơn giữa 2 khoản nợ
    DECLARE @NettingAmount DECIMAL(18,2);
    SET @NettingAmount = CASE WHEN @TotalUnpaidAR < @TotalUnpaidAP THEN @TotalUnpaidAR ELSE @TotalUnpaidAP END;

    IF @NettingAmount <= 0
    BEGIN
        PRINT N'Thông báo: Không có khoản nợ nào cần cấn trừ.';
        RETURN;
    END

    PRINT N'Bắt đầu cấn trừ công nợ số tiền: ' + CAST(@NettingAmount AS VARCHAR(50)) + N' VNĐ';

    -- 4. BẮT ĐẦU GIAO DỊCH (TRANSACTION)
    BEGIN TRANSACTION;

    DECLARE @PaymentRef VARCHAR(100);
    SET @PaymentRef = 'NETTING-' + @CustomerCode + '-' + @SupplierCode + '-' + FORMAT(@NettingDate, 'yyyyMMdd');

    -- =========================================================================
    -- A. BÙ TRỪ CHO BÊN PHẢI THU (AR) - PHÂN BỔ DỮ LIỆU THEO FIFO (Hóa đơn cũ trả trước)
    -- =========================================================================
    DECLARE @RemainingARNetting DECIMAL(18,2) = @NettingAmount;
    
    -- Khai báo bảng tạm lưu danh sách hóa đơn AR chưa trả để duyệt
    DECLARE @ARInvoiceID INT, @ARInvoiceAmount DECIMAL(18,2), @ARInvoicePaid DECIMAL(18,2), @ARInvoiceRemaining DECIMAL(18,2);
    DECLARE @ARInvoiceNo VARCHAR(50);

    DECLARE AR_Cursor CURSOR FOR 
    SELECT InvoiceID, InvoiceNumber, InvoiceAmount, PaidAmount, RemainingAmount
    FROM dbo.v_AR_Aging 
    WHERE CustomerID = @CustomerID AND AgingBucket <> 'Paid'
    ORDER BY InvoiceDate ASC, InvoiceID ASC;

    OPEN AR_Cursor;
    FETCH NEXT FROM AR_Cursor INTO @ARInvoiceID, @ARInvoiceNo, @ARInvoiceAmount, @ARInvoicePaid, @ARInvoiceRemaining;

    WHILE @@FETCH_STATUS = 0 AND @RemainingARNetting > 0
    BEGIN
        DECLARE @PayAmountAR DECIMAL(18,2);
        SET @PayAmountAR = CASE WHEN @RemainingARNetting < @ARInvoiceRemaining THEN @RemainingARNetting ELSE @ARInvoiceRemaining END;

        -- Sinh mã số phiếu thanh toán ngẫu nhiên duy nhất cho AR
        DECLARE @ARPayNum VARCHAR(50) = 'PAY-AR-NET-' + CAST(@ARInvoiceID AS VARCHAR(10)) + '-' + FORMAT(GETDATE(), 'ssfff');

        -- Chèn lịch sử thanh toán
        INSERT INTO dbo.FactARPayments (PaymentNumber, InvoiceID, PaymentDate, AmountPaid, PaymentMethod, ReferenceNumber)
        VALUES (@ARPayNum, @ARInvoiceID, @NettingDate, @PayAmountAR, 'Netting', @PaymentRef);

        -- Cập nhật trạng thái hóa đơn
        DECLARE @NewPaidAR DECIMAL(18,2) = @ARInvoicePaid + @PayAmountAR;
        DECLARE @NewStatusAR VARCHAR(20) = CASE WHEN @NewPaidAR >= @ARInvoiceAmount THEN 'Paid' ELSE 'Partially Paid' END;

        UPDATE dbo.FactARInvoices 
        SET Status = @NewStatusAR 
        WHERE InvoiceID = @ARInvoiceID;

        SET @RemainingARNetting = @RemainingARNetting - @PayAmountAR;

        FETCH NEXT FROM AR_Cursor INTO @ARInvoiceID, @ARInvoiceNo, @ARInvoiceAmount, @ARInvoicePaid, @ARInvoiceRemaining;
    END;

    CLOSE AR_Cursor;
    DEALLOCATE AR_Cursor;

    -- =========================================================================
    -- B. BÙ TRỪ CHO BÊN PHẢI TRẢ (AP) - PHÂN BỔ DỮ LIỆU THEO FIFO (Hóa đơn cũ trả trước)
    -- =========================================================================
    DECLARE @RemainingAPNetting DECIMAL(18,2) = @NettingAmount;
    
    DECLARE @APInvoiceID INT, @APInvoiceAmount DECIMAL(18,2), @APInvoicePaid DECIMAL(18,2), @APInvoiceRemaining DECIMAL(18,2);
    DECLARE @APInvoiceNo VARCHAR(50);

    DECLARE AP_Cursor CURSOR FOR 
    SELECT InvoiceID, InvoiceNumber, InvoiceAmount, PaidAmount, RemainingAmount
    FROM dbo.v_AP_Aging 
    WHERE SupplierID = @SupplierID AND AgingBucket <> 'Paid'
    ORDER BY InvoiceDate ASC, InvoiceID ASC;

    OPEN AP_Cursor;
    FETCH NEXT FROM AP_Cursor INTO @APInvoiceID, @APInvoiceNo, @APInvoiceAmount, @APInvoicePaid, @APInvoiceRemaining;

    WHILE @@FETCH_STATUS = 0 AND @RemainingAPNetting > 0
    BEGIN
        DECLARE @PayAmountAP DECIMAL(18,2);
        SET @PayAmountAP = CASE WHEN @RemainingAPNetting < @APInvoiceRemaining THEN @RemainingAPNetting ELSE @APInvoiceRemaining END;

        -- Sinh mã số phiếu thanh toán ngẫu nhiên duy nhất cho AP
        DECLARE @APPayNum VARCHAR(50) = 'PAY-AP-NET-' + CAST(@APInvoiceID AS VARCHAR(10)) + '-' + FORMAT(GETDATE(), 'ssfff');

        -- Chèn lịch sử thanh toán
        INSERT INTO dbo.FactAPPayments (PaymentNumber, InvoiceID, PaymentDate, AmountPaid, PaymentMethod, ReferenceNumber)
        VALUES (@APPayNum, @APInvoiceID, @NettingDate, @PayAmountAP, 'Netting', @PaymentRef);

        -- Cập nhật trạng thái hóa đơn
        DECLARE @NewPaidAP DECIMAL(18,2) = @APInvoicePaid + @PayAmountAP;
        DECLARE @NewStatusAP VARCHAR(20) = CASE WHEN @NewPaidAP >= @APInvoiceAmount THEN 'Paid' ELSE 'Partially Paid' END;

        UPDATE dbo.FactAPInvoices 
        SET Status = @NewStatusAP 
        WHERE InvoiceID = @APInvoiceID;

        SET @RemainingAPNetting = @RemainingAPNetting - @PayAmountAP;

        FETCH NEXT FROM AP_Cursor INTO @APInvoiceID, @APInvoiceNo, @APInvoiceAmount, @APInvoicePaid, @APInvoiceRemaining;
    END;

    CLOSE AP_Cursor;
    DEALLOCATE AP_Cursor;

    -- Hoàn tất Giao dịch thành công
    COMMIT TRANSACTION;
    PRINT N'Cấn trừ công nợ thành công!';
END;
GO

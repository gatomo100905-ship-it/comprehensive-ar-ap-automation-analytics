-- =========================================================================
-- FILE: sample_data.sql
-- DESCRIPTION: Chèn dữ liệu mẫu cho danh mục và giao dịch hóa đơn/thanh toán.
-- =========================================================================

USE AR_AP_Analytics;
GO

-- Xóa dữ liệu cũ nếu chạy lại
DELETE FROM dbo.FactARPayments;
DELETE FROM dbo.FactAPPayments;
DELETE FROM dbo.FactARInvoices;
DELETE FROM dbo.FactAPInvoices;
DELETE FROM dbo.DimCustomers;
DELETE FROM dbo.DimSuppliers;
GO

-- 1. Chèn Khách hàng mẫu (DimCustomers)
-- Chú ý: Khách hàng CUST-005 và Nhà cung cấp SUPP-005 có cùng mã số thuế '0109999999' để test nghiệp vụ Netting (Cấn trừ)
INSERT INTO dbo.DimCustomers (CustomerCode, CompanyName, ContactName, Email, Phone, Address, TaxCode, CreditLimit)
VALUES
('CUST-001', N'Trường Đại học Bách Khoa', N'Nguyễn Văn A', 'bk-purchasing@edu.vn', '0912345678', N'Đại Cồ Việt, Hai Bà Trưng, Hà Nội', '0101234567', 100000000.00),
('CUST-002', N'Công ty CP Đầu tư & Công nghệ Hùng Vương', N'Trần Thị B', 'contact@hungvuongtech.com', '0987654321', N'Nguyễn Trãi, Thanh Xuân, Hà Nội', '0107654321', 150000000.00),
('CUST-003', N'Đại lý Trà Sữa Phúc Long Quận 1', N'Lê Hoàng C', 'phuclongq1@gmail.com', '0901234567', N'Lê Lợi, Quận 1, TP. HCM', '0301122334', 50000000.00),
('CUST-004', N'Chuỗi Siêu Thị Minimart Hà Đông', N'Phạm Minh D', 'minimart.hadong@outlook.com', '0944556677', N'Quang Trung, Hà Đông, Hà Nội', '0102233445', 75000000.00),
('CUST-005', N'Tập đoàn Sữa Quốc tế IDP (Netting Partner)', N'Hoàng Xuân E', 'idp-ar@idp.com.vn', '0933221100', N'Ba Vì, Hà Nội', '0109999999', 300000000.00);
GO

-- 2. Chèn Nhà cung cấp mẫu (DimSuppliers)
INSERT INTO dbo.DimSuppliers (SupplierCode, CompanyName, ContactName, Email, Phone, Address, TaxCode)
VALUES
('SUPP-001', N'Công ty TNHH Bột Sữa Trân Châu Đài Loan', N'Vương Chí Hào', 'taiwantea@sugar.tw', '0911223344', N'KCN Tân Tạo, Bình Tân, TP. HCM', '0309876543'),
('SUPP-002', N'Nhà máy Bao Bì & In Ấn Ánh Dương', N'Vũ Văn Nam', 'anhduong.packaging@gmail.com', '0922334455', N'Như Quỳnh, Văn Lâm, Hưng Yên', '0902233445'),
('SUPP-003', N'Tổng Kho Nông Sản Tây Nguyên', N'Nguyễn Thị Hoa', 'taynguyenagro@gmail.com', '0955667788', N'Buôn Ma Thuột, Đắk Lắk', '6001234567'),
('SUPP-004', N'Hãng Vận Chuyển Giao Hàng Nhanh', N'Đỗ Minh Trí', 'shipping@ghn.vn', '0966778899', N'Lương Yên, Hai Bà Trưng, Hà Nội', '0106677889'),
('SUPP-005', N'Tập đoàn Sữa Quốc tế IDP (Netting Partner)', N'Lê Quang Huy', 'idp-ap@idp.com.vn', '0933221122', N'Ba Vì, Hà Nội', '0109999999');
GO

-- 3. Chèn Hóa đơn Phải thu (FactARInvoices)
-- Ngày hiện tại trong CSDL là: 2026-07-27.
-- Hạn thanh toán thường là 30 ngày (Credit terms Net 30).
INSERT INTO dbo.FactARInvoices (InvoiceNumber, CustomerID, InvoiceDate, DueDate, Amount, TaxAmount, TotalAmount, Description, Status)
VALUES
-- Hóa đơn quá hạn > 90 ngày (Phát hành tháng 3/2026) -> ID = 1
('INV-26-0001', 1, '2026-03-01', '2026-03-31', 45000000.00, 4500000.00, 49500000.00, N'Hợp đồng cung cấp trà sữa HK1', 'Unpaid'),
-- Hóa đơn quá hạn 31-60 ngày (Phát hành tháng 5/2026) -> ID = 2
('INV-26-0002', 2, '2026-05-01', '2026-05-31', 70000000.00, 7000000.00, 77000000.00, N'Cung cấp thiết bị bếp trà sữa', 'Partially Paid'),
-- Hóa đơn quá hạn 1-30 ngày (Phát hành tháng 6/2026) -> ID = 3
('INV-26-0003', 3, '2026-06-10', '2026-07-10', 25000000.00, 2500000.00, 27500000.00, N'Giao hàng trà sữa tuần 2 tháng 6', 'Unpaid'),
-- Hóa đơn đã thanh toán hoàn toàn (Paid) -> ID = 4
('INV-26-0004', 4, '2026-06-01', '2026-07-01', 30000000.00, 3000000.00, 33000000.00, N'Cung cấp sỉ bánh ngọt tháng 6', 'Paid'),
-- Hóa đơn chưa đến hạn (Trong hạn - Not yet due) -> ID = 5
('INV-26-0005', 1, '2026-07-10', '2026-08-10', 50000000.00, 5000000.00, 55000000.00, N'Hợp đồng cung cấp trà sữa HK2', 'Unpaid'),
-- Hóa đơn của Netting Partner (Đang treo chờ đối chiếu) -> ID = 6
('INV-26-0006', 5, '2026-06-15', '2026-07-15', 120000000.00, 12000000.00, 132000000.00, N'Bán lô nguyên liệu vỏ hộp sữa', 'Unpaid');
GO

-- 4. Chèn Hóa đơn Phải trả (FactAPInvoices)
INSERT INTO dbo.FactAPInvoices (InvoiceNumber, SupplierID, InvoiceDate, DueDate, Amount, TaxAmount, TotalAmount, Description, Status)
VALUES
-- Hóa đơn quá hạn > 60 ngày -> ID = 1
('PINV-26-0001', 1, '2026-04-10', '2026-05-10', 80000000.00, 8000000.00, 88000000.00, N'Mua bột sữa béo Đài Loan đợt 1', 'Unpaid'),
-- Hóa đơn quá hạn 1-30 ngày -> ID = 2
('PINV-26-0002', 2, '2026-06-15', '2026-07-15', 15000000.00, 1500000.00, 16500000.00, N'In ấn nhãn mác ly trà sữa tháng 6', 'Partially Paid'),
-- Hóa đơn đã thanh toán hoàn toàn -> ID = 3
('PINV-26-0003', 3, '2026-05-20', '2026-06-20', 40000000.00, 0.00, 40000000.00, N'Mua lô trà đen Bảo Lộc', 'Paid'),
-- Hóa đơn chưa đến hạn -> ID = 4
('PINV-26-0004', 4, '2026-07-15', '2026-08-15', 8000.00, 800.00, 8800.00, N'Cước vận chuyển nội thành tháng 7', 'Unpaid'),
-- Hóa đơn của Netting Partner -> ID = 5
('PINV-26-0005', 5, '2026-06-20', '2026-07-20', 70000000.00, 7000000.00, 77000000.00, N'Mua sữa tươi nguyên liệu', 'Unpaid');
GO

-- 5. Chèn Lịch sử thanh toán thu nợ (FactARPayments)
-- Cho hóa đơn INV-26-0002 (Partially Paid): Tổng nợ 77tr, thanh toán trước 40tr
INSERT INTO dbo.FactARPayments (PaymentNumber, InvoiceID, PaymentDate, AmountPaid, PaymentMethod, ReferenceNumber)
VALUES
('PAY-AR-0001', 2, '2026-06-05', 40000000.00, 'Bank Transfer', 'FT261570992384');

-- Cho hóa đơn INV-26-0004 (Paid): Tổng nợ 33tr, thanh toán đủ
INSERT INTO dbo.FactARPayments (PaymentNumber, InvoiceID, PaymentDate, AmountPaid, PaymentMethod, ReferenceNumber)
VALUES
('PAY-AR-0002', 4, '2026-06-28', 33000000.00, 'Bank Transfer', 'FT261790012398');
GO

-- 6. Chèn Lịch sử thanh toán trả nợ (FactAPPayments)
-- Cho hóa đơn PINV-26-0002 (Partially Paid): Tổng nợ 16.5tr, trả trước 10tr
INSERT INTO dbo.FactAPPayments (PaymentNumber, InvoiceID, PaymentDate, AmountPaid, PaymentMethod, ReferenceNumber)
VALUES
('PAY-AP-0001', 2, '2026-07-10', 10000000.00, 'Bank Transfer', 'TXN-99887723');

-- Cho hóa đơn PINV-26-0003 (Paid): Tổng nợ 40tr, trả đủ
INSERT INTO dbo.FactAPPayments (PaymentNumber, InvoiceID, PaymentDate, AmountPaid, PaymentMethod, ReferenceNumber)
VALUES
('PAY-AP-0002', 3, '2026-06-18', 40000000.00, 'Cash', 'CASH-REC-502');
GO

-- =========================================================================
-- FILE: schema.sql
-- DESCRIPTION: Khởi tạo CSDL AR_AP_Analytics và các bảng, ràng buộc dữ liệu.
-- =========================================================================

-- 1. Tạo Database
IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = 'AR_AP_Analytics')
BEGIN
    CREATE DATABASE AR_AP_Analytics;
END
GO

USE AR_AP_Analytics;
GO

-- Xóa các bảng cũ nếu tồn tại để chạy lại sạch sẽ (Theo thứ tự khóa ngoại trước)
IF OBJECT_ID('dbo.FactARPayments', 'U') IS NOT NULL DROP TABLE dbo.FactARPayments;
IF OBJECT_ID('dbo.FactAPPayments', 'U') IS NOT NULL DROP TABLE dbo.FactAPPayments;
IF OBJECT_ID('dbo.FactARInvoices', 'U') IS NOT NULL DROP TABLE dbo.FactARInvoices;
IF OBJECT_ID('dbo.FactAPInvoices', 'U') IS NOT NULL DROP TABLE dbo.FactAPInvoices;
IF OBJECT_ID('dbo.DimCustomers', 'U') IS NOT NULL DROP TABLE dbo.DimCustomers;
IF OBJECT_ID('dbo.DimSuppliers', 'U') IS NOT NULL DROP TABLE dbo.DimSuppliers;
GO

-- 2. Tạo bảng Danh mục Khách hàng (DimCustomers)
CREATE TABLE dbo.DimCustomers (
    CustomerID INT IDENTITY(1,1) PRIMARY KEY,
    CustomerCode VARCHAR(50) NOT NULL UNIQUE,
    CompanyName NVARCHAR(150) NOT NULL,
    ContactName NVARCHAR(100) NULL,
    Email VARCHAR(100) NULL,
    Phone VARCHAR(20) NULL,
    Address NVARCHAR(250) NULL,
    TaxCode VARCHAR(20) NULL,
    CreditLimit DECIMAL(18,2) NOT NULL DEFAULT 0.00 CHECK (CreditLimit >= 0),
    IsDeleted BIT NOT NULL DEFAULT 0,
    CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    UpdatedAt DATETIME NOT NULL DEFAULT GETDATE()
);
GO

-- 3. Tạo bảng Danh mục Nhà cung cấp (DimSuppliers)
CREATE TABLE dbo.DimSuppliers (
    SupplierID INT IDENTITY(1,1) PRIMARY KEY,
    SupplierCode VARCHAR(50) NOT NULL UNIQUE,
    CompanyName NVARCHAR(150) NOT NULL,
    ContactName NVARCHAR(100) NULL,
    Email VARCHAR(100) NULL,
    Phone VARCHAR(20) NULL,
    Address NVARCHAR(250) NULL,
    TaxCode VARCHAR(20) NULL,
    IsDeleted BIT NOT NULL DEFAULT 0,
    CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    UpdatedAt DATETIME NOT NULL DEFAULT GETDATE()
);
GO

-- 4. Tạo bảng Hóa đơn Phải thu (FactARInvoices)
CREATE TABLE dbo.FactARInvoices (
    InvoiceID INT IDENTITY(1,1) PRIMARY KEY,
    InvoiceNumber VARCHAR(50) NOT NULL UNIQUE,
    CustomerID INT NOT NULL,
    InvoiceDate DATE NOT NULL,
    DueDate DATE NOT NULL,
    Amount DECIMAL(18,2) NOT NULL CHECK (Amount > 0),
    TaxAmount DECIMAL(18,2) NOT NULL DEFAULT 0.00 CHECK (TaxAmount >= 0),
    TotalAmount DECIMAL(18,2) NOT NULL CHECK (TotalAmount > 0),
    Description NVARCHAR(250) NULL,
    Status VARCHAR(20) NOT NULL DEFAULT 'Unpaid' CHECK (Status IN ('Unpaid', 'Partially Paid', 'Paid')),
    IsDeleted BIT NOT NULL DEFAULT 0,
    CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    UpdatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_FactARInvoices_DimCustomers FOREIGN KEY (CustomerID) REFERENCES dbo.DimCustomers(CustomerID),
    CONSTRAINT CK_ARInvoice_Dates CHECK (DueDate >= InvoiceDate)
);
GO

-- 5. Tạo bảng Hóa đơn Phải trả (FactAPInvoices)
CREATE TABLE dbo.FactAPInvoices (
    InvoiceID INT IDENTITY(1,1) PRIMARY KEY,
    InvoiceNumber VARCHAR(50) NOT NULL UNIQUE,
    SupplierID INT NOT NULL,
    InvoiceDate DATE NOT NULL,
    DueDate DATE NOT NULL,
    Amount DECIMAL(18,2) NOT NULL CHECK (Amount > 0),
    TaxAmount DECIMAL(18,2) NOT NULL DEFAULT 0.00 CHECK (TaxAmount >= 0),
    TotalAmount DECIMAL(18,2) NOT NULL CHECK (TotalAmount > 0),
    Description NVARCHAR(250) NULL,
    Status VARCHAR(20) NOT NULL DEFAULT 'Unpaid' CHECK (Status IN ('Unpaid', 'Partially Paid', 'Paid')),
    IsDeleted BIT NOT NULL DEFAULT 0,
    CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    UpdatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_FactAPInvoices_DimSuppliers FOREIGN KEY (SupplierID) REFERENCES dbo.DimSuppliers(SupplierID),
    CONSTRAINT CK_APInvoice_Dates CHECK (DueDate >= InvoiceDate)
);
GO

-- 6. Bảng lịch sử Thanh toán khoản phải thu (FactARPayments)
CREATE TABLE dbo.FactARPayments (
    PaymentID INT IDENTITY(1,1) PRIMARY KEY,
    PaymentNumber VARCHAR(50) NOT NULL UNIQUE,
    InvoiceID INT NOT NULL,
    PaymentDate DATE NOT NULL,
    AmountPaid DECIMAL(18,2) NOT NULL CHECK (AmountPaid > 0),
    PaymentMethod VARCHAR(50) NOT NULL DEFAULT 'Bank Transfer' CHECK (PaymentMethod IN ('Bank Transfer', 'Cash', 'Credit Card', 'Netting')),
    ReferenceNumber VARCHAR(100) NULL,
    IsDeleted BIT NOT NULL DEFAULT 0,
    CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    UpdatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_FactARPayments_FactARInvoices FOREIGN KEY (InvoiceID) REFERENCES dbo.FactARInvoices(InvoiceID)
);
GO

-- 7. Bảng lịch sử Thanh toán khoản phải trả (FactAPPayments)
CREATE TABLE dbo.FactAPPayments (
    PaymentID INT IDENTITY(1,1) PRIMARY KEY,
    PaymentNumber VARCHAR(50) NOT NULL UNIQUE,
    InvoiceID INT NOT NULL,
    PaymentDate DATE NOT NULL,
    AmountPaid DECIMAL(18,2) NOT NULL CHECK (AmountPaid > 0),
    PaymentMethod VARCHAR(50) NOT NULL DEFAULT 'Bank Transfer' CHECK (PaymentMethod IN ('Bank Transfer', 'Cash', 'Credit Card', 'Netting')),
    ReferenceNumber VARCHAR(100) NULL,
    IsDeleted BIT NOT NULL DEFAULT 0,
    CreatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    UpdatedAt DATETIME NOT NULL DEFAULT GETDATE(),
    CONSTRAINT FK_FactAPPayments_FactAPInvoices FOREIGN KEY (InvoiceID) REFERENCES dbo.FactAPInvoices(InvoiceID)
);
GO

-- =========================================================================
-- TRIGGERS TỰ ĐỘNG CẬP NHẬT CỘT UpdatedAt KHI CÓ SỰ THAY ĐỔI DỮ LIỆU
-- =========================================================================

CREATE TRIGGER dbo.trg_DimCustomers_Update ON dbo.DimCustomers
AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.DimCustomers SET UpdatedAt = GETDATE()
    FROM dbo.DimCustomers t INNER JOIN inserted i ON t.CustomerID = i.CustomerID;
END;
GO

CREATE TRIGGER dbo.trg_DimSuppliers_Update ON dbo.DimSuppliers
AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.DimSuppliers SET UpdatedAt = GETDATE()
    FROM dbo.DimSuppliers t INNER JOIN inserted i ON t.SupplierID = i.SupplierID;
END;
GO

CREATE TRIGGER dbo.trg_FactARInvoices_Update ON dbo.FactARInvoices
AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.FactARInvoices SET UpdatedAt = GETDATE()
    FROM dbo.FactARInvoices t INNER JOIN inserted i ON t.InvoiceID = i.InvoiceID;
END;
GO

CREATE TRIGGER dbo.trg_FactAPInvoices_Update ON dbo.FactAPInvoices
AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.FactAPInvoices SET UpdatedAt = GETDATE()
    FROM dbo.FactAPInvoices t INNER JOIN inserted i ON t.InvoiceID = i.InvoiceID;
END;
GO

CREATE TRIGGER dbo.trg_FactARPayments_Update ON dbo.FactARPayments
AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.FactARPayments SET UpdatedAt = GETDATE()
    FROM dbo.FactARPayments t INNER JOIN inserted i ON t.PaymentID = i.PaymentID;
END;
GO

CREATE TRIGGER dbo.trg_FactAPPayments_Update ON dbo.FactAPPayments
AFTER UPDATE AS
BEGIN
    SET NOCOUNT ON;
    UPDATE dbo.FactAPPayments SET UpdatedAt = GETDATE()
    FROM dbo.FactAPPayments t INNER JOIN inserted i ON t.PaymentID = i.PaymentID;
END;
GO

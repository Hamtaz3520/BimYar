CREATE TABLE tblInstallments
(
    InstallmentID AUTOINCREMENT CONSTRAINT PK_tblInstallments PRIMARY KEY,

    PolicyID LONG NOT NULL,

    InstallmentNo INTEGER NOT NULL,

    DueDate DATETIME,

    Amount CURRENCY,

    PaidAmount CURRENCY,

    RemainingAmount CURRENCY,

    Status TEXT(20),

    PaymentDate DATETIME,

    IsActive YESNO,

    Description LONGTEXT,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
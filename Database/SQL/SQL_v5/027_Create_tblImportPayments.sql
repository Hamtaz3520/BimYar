CREATE TABLE tblImportPayments (
    ImportPaymentID COUNTER
        CONSTRAINT PK_tblImportPayments PRIMARY KEY,
    ImportPolicyID LONG NOT NULL,
    ImportSessionID LONG,
    SourceRowNo LONG,
    PaymentSequence LONG NOT NULL,
    PaymentType TEXT(20),
    PaymentAmount CURRENCY,
    PaymentRefNo TEXT(255),
    PaymentDate TEXT(10),
    IsInstallment YESNO,
    CreatedDate DATETIME
);

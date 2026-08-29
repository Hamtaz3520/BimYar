CREATE TABLE tblPayments
(
    PaymentID AUTOINCREMENT
        CONSTRAINT PK_tblPayments PRIMARY KEY,

    PolicyID LONG,

    InstallmentID LONG,

    PaymentDate DATETIME,

    Amount CURRENCY,

    PaymentMethod TEXT(30),

    PayerName TEXT(100),

    PayerCardNo TEXT(25),

    PayerAccountNo TEXT(40),

    TrackingCode TEXT(100),

    BankTransactionID LONG,

    MatchStatus BYTE,

    IsConfirmed YESNO,

    Description LONGTEXT,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
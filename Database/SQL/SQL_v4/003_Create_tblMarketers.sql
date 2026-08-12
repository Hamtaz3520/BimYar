CREATE TABLE tblMarketers
(
    MarketerID AUTOINCREMENT PRIMARY KEY,

    MarketerCode TEXT(20),

    FirstName TEXT(50) NOT NULL,

    LastName TEXT(50) NOT NULL,

    NationalCode TEXT(10),

    Mobile TEXT(20),

    Phone TEXT(20),

    Email TEXT(100),

    Address MEMO,

    HireDate DATETIME,

    IsActive YESNO,

    Description MEMO,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
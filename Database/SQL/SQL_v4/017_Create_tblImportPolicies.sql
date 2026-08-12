CREATE TABLE tblImportPolicies (

    ImportPolicyID AUTOINCREMENT
        CONSTRAINT PK_tblImportPolicies PRIMARY KEY,

    ReportType BYTE NOT NULL,

    MarketerID LONG,

    ImportFileName TEXT(255),

    ImportDate DATETIME,

    ImportUserID LONG,

    RowNo LONG,

    PolicyNumber TEXT(50),

    InsuredName TEXT(150),

    PreviousInsurer TEXT(100),

    PreviousPolicyNumber TEXT(50),

    VehicleType TEXT(100),

    SystemType TEXT(100),

    VehicleModel TEXT(100),

    InsuredType TEXT(50),

    NationalCode TEXT(20),

    Phone TEXT(30),

    Mobile TEXT(30),

    Address LONGTEXT,

    StartDate DATETIME,

    EndDate DATETIME,

    IssueDate DATETIME,

    PremiumAmount CURRENCY,

    Profile TEXT(100),

    PaymentType TEXT(50),

    PaymentAmount CURRENCY,

    PaymentRefNo TEXT(100),

    PaymentDate DATETIME,

    IsImported YESNO

);
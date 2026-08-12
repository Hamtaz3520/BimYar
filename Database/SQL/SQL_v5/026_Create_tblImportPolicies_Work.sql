CREATE TABLE tblImportPolicies_Work (

    WorkID AUTOINCREMENT
        CONSTRAINT PK_tblImportPolicies_Work PRIMARY KEY,

    SourceImportPolicyID LONG,

    ReportType BYTE,
    MarketerID LONG,

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

    ParentPolicyNumber TEXT(50),

    IsPolicyRow YESNO,
    IsInstallmentRow YESNO,
    IsHeaderRow YESNO,

    IsMarketerMatched YESNO,
    IsManuallyAssigned YESNO,

    MatchStatus TEXT(30),
    ProcessStatus TEXT(30),

    ErrorMessage LONGTEXT

);
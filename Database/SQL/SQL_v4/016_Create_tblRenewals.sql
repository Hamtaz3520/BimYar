CREATE TABLE tblRenewals (

    RenewalID AUTOINCREMENT
        CONSTRAINT PK_tblRenewals PRIMARY KEY,

    PolicyID LONG NOT NULL,

    CustomerID LONG NOT NULL,

    ExpireDate DATETIME,

    ReminderDate DATETIME,

    RenewalStatus TEXT(30),

    PriorityLevel BYTE,

    ReminderCount INTEGER,

    LastReminderDate DATETIME,

    RenewedPolicyID LONG,

    Notes LONGTEXT,

    UserID LONG,

    IsClosed YESNO,

    CreatedDate DATETIME,

    ModifiedDate DATETIME

);
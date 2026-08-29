CREATE TABLE tblFollowUps (

    FollowUpID AUTOINCREMENT
        CONSTRAINT PK_tblFollowUps PRIMARY KEY,

    CustomerID LONG NOT NULL,

    PolicyID LONG,

    InstallmentID LONG,

    FollowUpType TEXT(30),

    FollowUpDate DATETIME,

    NextFollowUpDate DATETIME,

    Subject TEXT(100),

    Description LONGTEXT,

    Result TEXT(100),

    IsDone YESNO,

    UserID LONG,

    CreatedDate DATETIME,

    ModifiedDate DATETIME

);
CREATE TABLE tblRenewalStatuses
(
    RenewalStatusID AUTOINCREMENT CONSTRAINT PK_tblRenewalStatuses PRIMARY KEY,
    StatusCode TEXT(30),
    StatusTitle TEXT(100),
    IsActive YESNO,
    SortOrder INTEGER,
    CreatedDate DATETIME
);
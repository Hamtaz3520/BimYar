/*=========================================================
  Project : BimYar
  File    : 005_Create_tblBranches.sql
  Table   : tblBranches
  Version : 4.0
=========================================================*/

CREATE TABLE tblBranches
(
    BranchID AUTOINCREMENT
        CONSTRAINT PK_tblBranches PRIMARY KEY,

    CompanyID LONG NOT NULL,

    BranchCode TEXT(10),

    BranchName TEXT(100) NOT NULL,

    Province TEXT(50),

    City TEXT(50),

    Address LONGTEXT,

    Phone TEXT(30),

    ManagerName TEXT(100),

    IsActive YESNO,

    Description LONGTEXT,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
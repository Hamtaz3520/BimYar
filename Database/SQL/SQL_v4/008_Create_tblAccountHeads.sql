/*=========================================================
  Project : BimYar
  File    : 008_Create_tblAccountHeads.sql
  Table   : tblAccountHeads
  Version : 4.0
=========================================================*/

CREATE TABLE tblAccountHeads
(
    AccountHeadID AUTOINCREMENT
        CONSTRAINT PK_tblAccountHeads PRIMARY KEY,

    AccountCode TEXT(20),

    AccountTitle TEXT(100) NOT NULL,

    ParentAccountID LONG,

    AccountType TEXT(20),

    IsSystem YESNO,

    IsActive YESNO,

    Description LONGTEXT,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
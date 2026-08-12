/*=========================================================
  Project : BimYar
  File    : 002_Create_tblUsers.sql
  Table   : tblUsers
  Version : 4.0
=========================================================*/

CREATE TABLE tblUsers
(
    UserID AUTOINCREMENT
        CONSTRAINT PK_tblUsers PRIMARY KEY,

    Username TEXT(50) NOT NULL,

    PasswordHash TEXT(255) NOT NULL,

    FullName TEXT(100),

    RoleID LONG,

    IsActive YESNO,

    LastLogin DATETIME,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
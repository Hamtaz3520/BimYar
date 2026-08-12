/*=========================================================
  Project : BimYar
  File    : 001_Create_tblRoles.sql
  Table   : tblRoles
  Version : 4.0
=========================================================*/

CREATE TABLE tblRoles
(
    RoleID AUTOINCREMENT
        CONSTRAINT PK_tblRoles PRIMARY KEY,

    RoleName TEXT(50) NOT NULL,

    Description LONGTEXT,

    IsActive YESNO,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
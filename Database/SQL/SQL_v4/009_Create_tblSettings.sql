/*=========================================================
  Project : BimYar
  File    : 009_Create_tblSettings.sql
  Table   : tblSettings
  Version : 4.0
=========================================================*/

CREATE TABLE tblSettings
(
    SettingID AUTOINCREMENT
        CONSTRAINT PK_tblSettings PRIMARY KEY,

    SettingKey TEXT(50) NOT NULL,

    SettingValue LONGTEXT,

    Description LONGTEXT,

    IsSystem YESNO,

    ModifiedDate DATETIME
);
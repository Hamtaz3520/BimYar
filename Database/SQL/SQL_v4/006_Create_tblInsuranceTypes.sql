/*=========================================================
  Project : BimYar
  File    : 006_Create_tblInsuranceTypes.sql
  Table   : tblInsuranceTypes
  Version : 4.0
=========================================================*/

CREATE TABLE tblInsuranceTypes
(
    InsuranceTypeID AUTOINCREMENT
        CONSTRAINT PK_tblInsuranceTypes PRIMARY KEY,

    TypeCode TEXT(20),

    TypeName TEXT(100) NOT NULL,

    ParentTypeID LONG,

    HasInstallments YESNO,

    DefaultCommission DOUBLE,

    IsActive YESNO,

    Description LONGTEXT,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
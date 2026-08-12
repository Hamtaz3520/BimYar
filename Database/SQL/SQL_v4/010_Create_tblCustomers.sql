/*=========================================================
  Project : BimYar
  File    : 010_Create_tblCustomers.sql
  Table   : tblCustomers
  Version : 4.0
=========================================================*/

CREATE TABLE tblCustomers
(
    CustomerID AUTOINCREMENT
        CONSTRAINT PK_tblCustomers PRIMARY KEY,

    CustomerCode TEXT(20),

    CustomerType TEXT(20) NOT NULL,

    NationalCode TEXT(10),

    NationalID TEXT(20),

    FirstName TEXT(50),

    LastName TEXT(50),

    CompanyName TEXT(150),

    FatherName TEXT(50),

    BirthDate DATETIME,

    Mobile TEXT(20),

    Phone TEXT(20),

    Email TEXT(100),

    PostalCode TEXT(10),

    Province TEXT(50),

    City TEXT(50),

    Address LONGTEXT,

    IsActive YESNO,

    Description LONGTEXT,

    CreatedDate DATETIME,

    ModifiedDate DATETIME
);
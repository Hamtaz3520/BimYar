# SQL_v4 Schema Analysis

**Work item:** `BMY-CX-20260813-009`
**Phase / step:** Phase 3 / Step 3.5
**Analysis date:** 2026-08-13
**Authoritative scope:** `Database/SQL/SQL_v4/`

## 1. Scope, method, and headline result

Every repository file currently present in the target directory was inspected as text. The 18 non-placeholder `.sql` files are the authoritative source for this analysis; `.gitkeep` is excluded. Scripts numbered 001 through 017 are present, with two distinct files numbered 012: one table DDL file and one index-only file. No SQL logic was executed, repaired, inferred, or changed.

- **Real SQL_v4 files analyzed:** 18.
- **Tables represented by `CREATE TABLE`:** 17.
- **Other schema object statements:** five indexes in the additional 012 file (four non-unique and one unique composite index).
- **Declared primary keys:** 17 (one per table).
- **Declared foreign-key constraints:** 0.
- **Declared defaults:** 0.
- **Declared uniqueness beyond primary keys:** one composite unique index on `tblInstallments (PolicyID, InstallmentNo)`.
- **Explicit `NOT NULL`:** 25 non-PK fields; all other non-PK fields are nullable as far as the scripts determine. AutoNumber primary keys are inherently required in the resulting table.

This is a source analysis, not proof that the scripts ran successfully in Access or that the live `.accdb` matches them.

## 2. File-by-file schema inventory

In the field lists below, `!` means explicit `NOT NULL`; unmarked non-PK fields have no explicit required constraint. “Logical dependency” means an identifier strongly indicates a relationship, even though **none** is enforced by a `FOREIGN KEY` clause.

### 001 — `001_Create_tblRoles.sql`

- **Creates:** `tblRoles`.
- **Fields:** `RoleID AUTOINCREMENT` (PK); `RoleName TEXT(50)!`; `Description LONGTEXT`; `IsActive YESNO`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named `PK_tblRoles`; no other index or unique constraint.
- **Defaults:** none. **Dependencies:** none.
- **Risks/gaps:** role names are not unique; active/date fields have no defaults.

### 002 — `002_Create_tblUsers.sql`

- **Creates:** `tblUsers`.
- **Fields:** `UserID AUTOINCREMENT` (PK); `Username TEXT(50)!`; `PasswordHash TEXT(255)!`; `FullName TEXT(100)`; `RoleID LONG`; `IsActive YESNO`; `LastLogin DATETIME`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named `PK_tblUsers`; no other index or unique constraint.
- **Defaults:** none. **Logical dependency:** `RoleID` → `tblRoles.RoleID`.
- **Risks/gaps:** no FK or `RoleID` index; usernames are not unique; role is nullable.

### 003 — `003_Create_tblMarketers.sql`

- **Creates:** `tblMarketers`.
- **Fields:** `MarketerID AUTOINCREMENT` (PK); `MarketerCode TEXT(20)`; `FirstName TEXT(50)!`; `LastName TEXT(50)!`; `NationalCode TEXT(10)`; `Mobile TEXT(20)`; `Phone TEXT(20)`; `Email TEXT(100)`; `Address MEMO`; `HireDate DATETIME`; `IsActive YESNO`; `Description MEMO`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** unnamed PK; no secondary/unique indexes.
- **Defaults/dependencies:** none.
- **Risks/gaps:** no indexes for `MarketerCode`, `NationalCode`, `Mobile`, name, or `HireDate`; `MEMO` differs from the `LONGTEXT` spelling used elsewhere; `NationalCode` length 10 conflicts with length 20 in import tables.

### 004 — `004_Create_tblInsuranceCompanies.sql`

- **Creates:** `tblInsuranceCompanies`.
- **Fields:** `CompanyID AUTOINCREMENT` (PK); `CompanyCode TEXT(10)`; `CompanyName TEXT(100)!`; `ManagerName TEXT(100)`; `Phone TEXT(30)`; `Fax TEXT(30)`; `Email TEXT(100)`; `Website TEXT(100)`; `Address LONGTEXT`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults/dependencies:** none.
- **Risks/gaps:** company code/name have no uniqueness or lookup index.

### 005 — `005_Create_tblBranches.sql`

- **Creates:** `tblBranches`.
- **Fields:** `BranchID AUTOINCREMENT` (PK); `CompanyID LONG!`; `BranchCode TEXT(10)`; `BranchName TEXT(100)!`; `Province TEXT(50)`; `City TEXT(50)`; `Address LONGTEXT`; `Phone TEXT(30)`; `ManagerName TEXT(100)`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes.
- **Defaults:** none. **Logical dependency:** `CompanyID` → `tblInsuranceCompanies.CompanyID`.
- **Risks/gaps:** required parent identifier is neither enforced nor indexed; likely company/branch code uniqueness is undefined; geographic search fields are unindexed.

### 006 — `006_Create_tblInsuranceTypes.sql`

- **Creates:** `tblInsuranceTypes`.
- **Fields:** `InsuranceTypeID AUTOINCREMENT` (PK); `TypeCode TEXT(20)`; `TypeName TEXT(100)!`; `ParentTypeID LONG`; `HasInstallments YESNO`; `DefaultCommission DOUBLE`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes.
- **Defaults:** none. **Logical dependency:** self-reference `ParentTypeID` → `tblInsuranceTypes.InsuranceTypeID`.
- **Risks/gaps:** no self-FK or parent index; type code/name uniqueness undefined; `DOUBLE` may introduce rounding if commission is treated as exact financial data.

### 007 — `007_Create_tblBanks.sql`

- **Creates:** `tblBanks`.
- **Fields:** `BankID AUTOINCREMENT` (PK); `BankCode TEXT(10)`; `BankName TEXT(100)!`; `BranchName TEXT(100)`; `AccountNumber TEXT(50)`; `CardNumber TEXT(20)`; `IBAN TEXT(30)`; `AccountHolder TEXT(100)`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults/dependencies:** none.
- **Risks/gaps:** table mixes bank master and account details; no indexes/uniqueness on bank code, account number, card number, or IBAN. Its PK is `BankID`, while `tblBankTransactions` uses `BankAccountID`, leaving the intended relationship inconsistent.

### 008 — `008_Create_tblAccountHeads.sql`

- **Creates:** `tblAccountHeads`.
- **Fields:** `AccountHeadID AUTOINCREMENT` (PK); `AccountCode TEXT(20)`; `AccountTitle TEXT(100)!`; `ParentAccountID LONG`; `AccountType TEXT(20)`; `IsSystem YESNO`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes.
- **Defaults:** none. **Logical dependency:** self-reference `ParentAccountID` → `tblAccountHeads.AccountHeadID`.
- **Risks/gaps:** no self-FK/parent index; account code uniqueness undefined.

### 009 — `009_Create_tblSettings.sql`

- **Creates:** `tblSettings`.
- **Fields:** `SettingID AUTOINCREMENT` (PK); `SettingKey TEXT(50)!`; `SettingValue LONGTEXT`; `Description LONGTEXT`; `IsSystem YESNO`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults/dependencies:** none.
- **Risks/gaps:** `SettingKey` is neither indexed nor unique, permitting ambiguous duplicate settings; no created timestamp.

### 010 — `010_Create_tblCustomers.sql`

- **Creates:** `tblCustomers`.
- **Fields:** `CustomerID AUTOINCREMENT` (PK); `CustomerCode TEXT(20)`; `CustomerType TEXT(20)!`; `NationalCode TEXT(10)`; `NationalID TEXT(20)`; `FirstName TEXT(50)`; `LastName TEXT(50)`; `CompanyName TEXT(150)`; `FatherName TEXT(50)`; `BirthDate DATETIME`; `Mobile TEXT(20)`; `Phone TEXT(20)`; `Email TEXT(100)`; `PostalCode TEXT(10)`; `Province TEXT(50)`; `City TEXT(50)`; `Address LONGTEXT`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults/dependencies:** none.
- **Risks/gaps:** no indexes on `NationalCode`, `NationalID`, `Mobile`, customer code, names, or geography; person/company conditional requiredness is not encoded; import `NationalCode` is length 20/30-phone conventions differ.

### 011 — `011_Create_tblPolicies.sql`

- **Creates:** `tblPolicies`.
- **Fields:** `PolicyID AUTOINCREMENT` (PK); `PolicyNumber TEXT(50)!`; `UniquePolicyCode TEXT(50)`; `CustomerID LONG!`; `MarketerID LONG`; `CompanyID LONG!`; `BranchID LONG`; `InsuranceTypeID LONG!`; `PolicySeries TEXT(20)`; `ProposalNumber TEXT(50)`; `IssueDate DATETIME`; `StartDate DATETIME!`; `EndDate DATETIME!`; `PremiumAmount CURRENCY`; `TaxAmount CURRENCY`; `DiscountAmount CURRENCY`; `NetPremium CURRENCY`; `InstallmentCount INTEGER`; `PaymentMethod TEXT(30)`; `PolicyStatus TEXT(30)`; `SourceFileName TEXT(255)`; `SourceRowNumber LONG`; `IsImported YESNO`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults:** none.
- **Logical dependencies:** customer, marketer, company, branch, and insurance type identifiers point to their corresponding tables.
- **Risks/gaps:** no FK constraints and none of the five logical FKs is indexed; `PolicyNumber` and `UniquePolicyCode` lack indexes/uniqueness; issue/start/end/status and source provenance fields are unindexed; `INTEGER` is Access 16-bit and may be an unnecessarily narrow installment count.

### 012 table — `012_Create_tblInstallments.sql`

- **Creates:** `tblInstallments`.
- **Fields:** `InstallmentID AUTOINCREMENT` (PK); `PolicyID LONG!`; `InstallmentNo INTEGER!`; `DueDate DATETIME`; `Amount CURRENCY`; `PaidAmount CURRENCY`; `RemainingAmount CURRENCY`; `Status TEXT(20)`; `PaymentDate DATETIME`; `IsActive YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/defaults:** named PK; no defaults. **Logical dependency:** `PolicyID` → `tblPolicies.PolicyID`.
- **Risks/gaps:** no FK constraint; `Status` is a generic/reserved-word-risk identifier and should require brackets in some Access contexts. Indexes are supplied only by the companion 012 script described next.

### 012 indexes — `012_Create_tblInstallments -Indexes.sql`

- **Creates:** no table; five indexes on pre-existing `tblInstallments`:
  - `IX_tblInstallments_PolicyID` on `PolicyID`;
  - `IX_tblInstallments_DueDate` on `DueDate`;
  - `IX_tblInstallments_Status` on `Status`;
  - `IX_tblInstallments_PaymentDate` on `PaymentDate`;
  - unique `IX_tblInstallments_PolicyNo` on (`PolicyID`, `InstallmentNo`).
- **Dependencies/order:** must run after `012_Create_tblInstallments.sql`, and therefore after policies for the logical model.
- **Risks/gaps:** duplicate sequence number makes lexical/run ordering ambiguous; filename contains a space before `-Indexes`; the unique index is misleadingly named `PolicyNo` although its second field is `InstallmentNo`; five `CREATE INDEX` statements in one file may require statement-by-statement execution by an Access/DAO runner rather than a single saved query execution. This is the only SQL_v4 file containing secondary indexes.

### 013 — `013_Create_tblPayments.sql`

- **Creates:** `tblPayments`.
- **Fields:** `PaymentID AUTOINCREMENT` (PK); `PolicyID LONG`; `InstallmentID LONG`; `PaymentDate DATETIME`; `Amount CURRENCY`; `PaymentMethod TEXT(30)`; `PayerName TEXT(100)`; `PayerCardNo TEXT(25)`; `PayerAccountNo TEXT(40)`; `TrackingCode TEXT(100)`; `BankTransactionID LONG`; `MatchStatus BYTE`; `IsConfirmed YESNO`; `Description LONGTEXT`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults:** none.
- **Logical dependencies:** policies, installments, and bank transactions.
- **Risks/gaps:** no FK constraints/indexes; all relationship identifiers are nullable, so a payment can be unattached; payment date, tracking code, payer card/account, amount, match status, and confirmation searches are unindexed; `MatchStatus BYTE` conflicts in type with text match statuses in the work table.

### 014 — `014_Create_tblBankTransactions.sql`

- **Creates:** `tblBankTransactions` (header says Version 5.0 despite location in SQL_v4).
- **Fields:** `TransactionID AUTOINCREMENT` (PK); `BankAccountID LONG`; `TransactionDate DATETIME`; `TransactionTime DATETIME`; `DebitAmount CURRENCY`; `CreditAmount CURRENCY`; `Balance CURRENCY`; `TransactionType TEXT(30)`; `TransactionSource TEXT(30)`; `ReferenceNo TEXT(100)`; `TrackingCode TEXT(100)`; `PayerName TEXT(100)`; `PayerCardNo TEXT(25)`; `PayerAccountNo TEXT(40)`; `Description LONGTEXT`; `IsMatched YESNO`; `ImportBatchID LONG`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults:** none.
- **Unresolved dependencies:** `BankAccountID` has no matching `tblBankAccounts` table/PK in SQL_v4; `ImportBatchID` has no matching import-batch table.
- **Risks/gaps:** no matching/search indexes on date/time, reference, tracking code, payer identifiers, amounts, match flag, bank account, or batch; separate DATETIME date/time fields can become inconsistent; both debit and credit may be null/set because no validation is declared; version header conflicts with folder/version set.

### 015 — `015_Create_tblFollowUps.sql`

- **Creates:** `tblFollowUps`.
- **Fields:** `FollowUpID AUTOINCREMENT` (PK); `CustomerID LONG!`; `PolicyID LONG`; `InstallmentID LONG`; `FollowUpType TEXT(30)`; `FollowUpDate DATETIME`; `NextFollowUpDate DATETIME`; `Subject TEXT(100)`; `Description LONGTEXT`; `Result TEXT(100)`; `IsDone YESNO`; `UserID LONG`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults:** none.
- **Logical dependencies:** customers, policies, installments, users.
- **Risks/gaps:** no FK constraints/indexes; required customer is not enforced relationally; date/type/completion/user work-queue searches are unindexed.

### 016 — `016_Create_tblRenewals.sql`

- **Creates:** `tblRenewals`.
- **Fields:** `RenewalID AUTOINCREMENT` (PK); `PolicyID LONG!`; `CustomerID LONG!`; `ExpireDate DATETIME`; `ReminderDate DATETIME`; `RenewalStatus TEXT(30)`; `PriorityLevel BYTE`; `ReminderCount INTEGER`; `LastReminderDate DATETIME`; `RenewedPolicyID LONG`; `Notes LONGTEXT`; `UserID LONG`; `IsClosed YESNO`; `CreatedDate DATETIME`; `ModifiedDate DATETIME`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults:** none.
- **Logical dependencies:** policies (original and renewed), customers, users. SQL_v5 later creates `tblRenewalStatuses`, but this table stores status text and has no `RenewalStatusID`.
- **Risks/gaps:** no FK constraints/indexes; expiry/reminder/status/priority/closed/user work-queue fields unindexed; duplicated customer may disagree with policy customer; `ExpireDate` naming differs from policy `EndDate`; `INTEGER` is 16-bit; renewal status is denormalized and inconsistent with the SQL_v5 lookup table.

### 017 — `017_Create_tblImportPolicies.sql`

- **Creates:** `tblImportPolicies`.
- **Fields:** `ImportPolicyID AUTOINCREMENT` (PK); `ReportType BYTE!`; `MarketerID LONG`; `ImportFileName TEXT(255)`; `ImportDate DATETIME`; `ImportUserID LONG`; `RowNo LONG`; `PolicyNumber TEXT(50)`; `InsuredName TEXT(150)`; `PreviousInsurer TEXT(100)`; `PreviousPolicyNumber TEXT(50)`; `VehicleType TEXT(100)`; `SystemType TEXT(100)`; `VehicleModel TEXT(100)`; `InsuredType TEXT(50)`; `NationalCode TEXT(20)`; `Phone TEXT(30)`; `Mobile TEXT(30)`; `Address LONGTEXT`; `StartDate DATETIME`; `EndDate DATETIME`; `IssueDate DATETIME`; `PremiumAmount CURRENCY`; `Profile TEXT(100)`; `PaymentType TEXT(50)`; `PaymentAmount CURRENCY`; `PaymentRefNo TEXT(100)`; `PaymentDate DATETIME`; `IsImported YESNO`.
- **Keys/indexes/uniqueness:** named PK; no secondary/unique indexes. **Defaults:** none.
- **Logical dependencies:** `MarketerID` → marketers; `ImportUserID` → users. There is no `ImportSessionID`.
- **Risks/gaps:** neither logical FK is constrained/indexed; no import-session identity; file/date/row, policy number, national code, mobile, date, import-state, and payment matching fields are unindexed; `InsuredName TEXT(150)` conflicts with status documentation claiming it was changed to Long Text; operational customer/marketer phone and national-code lengths differ.

## 3. Relationships and dependency order

### Declared versus logical relationships

There are **zero declared foreign keys** across the set. At least 24 identifier occurrences imply relationships (including repeated references and self-references): users→roles; branches→companies; insurance types→self; account heads→self; policies→customers/marketers/companies/branches/types; installments→policies; payments→policies/installments/bank transactions; follow-ups→customers/policies/installments/users; renewals→policies twice/customers/users; import policies→marketers/users. Two more identifiers are unresolved (`BankAccountID`, `ImportBatchID`). Thus the scripts do not enforce referential integrity, cascade policy, or orphan prevention.

### Safe dependency interpretation (not an execution script)

The numeric order generally places parent tables before children. Exceptions/ambiguities are:

1. both 012 scripts share a number and the index script requires the table script first;
2. payments at 013 logically reference bank transactions created later at 014;
3. `tblRenewals.RenewalStatus` does not reference the later SQL_v5 status lookup;
4. no parent objects exist for `BankAccountID` and `ImportBatchID`;
5. no run-order guide or relationship-alteration scripts are present.

## 4. Index and constraint findings

### What is actually implemented

Only `012_Create_tblInstallments -Indexes.sql` defines secondary indexes. It gives installments a policy FK lookup, due-date/status/payment-date search indexes, and one business uniqueness rule per policy/installment number. No table script contains an inline secondary index. All table scripts define only their primary key; uniqueness and defaults are otherwise absent.

### Highest-impact index gaps (business rules still require approval)

- **Foreign keys:** every logical FK except `tblInstallments.PolicyID` lacks an index; none has an FK constraint.
- **Policy lookup/deduplication:** `tblPolicies.PolicyNumber` and `UniquePolicyCode`, plus import/work `PolicyNumber`/`ParentPolicyNumber`, have no indexes or uniqueness rule. Global uniqueness must not be assumed without company/branch/type scope rules.
- **Identity/contact:** operational and import `NationalCode` and `Mobile` fields are unindexed; length/normalization differ, so uniqueness should not be added until invalid, blank, shared, and formatted values are governed.
- **Marketers:** `tblPolicies.MarketerID` and `tblImportPolicies.MarketerID` are unindexed; the SQL_v5 work table also lacks the index.
- **Import session:** SQL_v4 has no `ImportSessionID` field or `tblImportSessions` script, so session indexing/relationship cannot be implemented from this set.
- **Dates/search:** policy issue/start/end, import dates and policy/payment dates, follow-up queue dates, renewal expiry/reminders, payment dates, and bank transaction date are unindexed. Add only indexes supported by actual query workload.
- **Payment/bank matching:** tracking/reference codes, payer card/account, amount/date combinations, match flags/status, bank account, and batch are entirely unindexed. Matching keys and tolerance/collision rules are undocumented.
- **Master/candidate keys:** username, setting key, role name, codes for marketers/companies/branches/types/banks/accounts/customers, and bank account identifiers lack unique constraints.

## 5. Access compatibility and schema consistency risks

1. **No explicit foreign keys or defaults.** Access-compatible types are mostly used, but behavior relies on application code and nullable fields.
2. **Reserved/generic identifiers.** `Status`, `Description`, `Address`, and similar common terms may require brackets in Access queries/VBA; `Status` is the clearest risk.
3. **Multi-statement index file.** Access execution paths commonly require each DDL statement to be executed separately; the runner must be verified rather than assuming the five-statement file runs atomically.
4. **Type aliases/style drift.** `MEMO` and `LONGTEXT` both appear; PK constraint naming is omitted only in 003; formatting/header metadata varies; 014 declares Version 5.0 within SQL_v4.
5. **Narrow numeric types.** Access `INTEGER` is 16-bit and `BYTE` is 0–255. Confirm these bounds for installment/reminder counts and encoded statuses/report types.
6. **Text/type inconsistency.** `NationalCode` is 10 in operational marketer/customer tables and 20 in import/work; phone/mobile lengths vary 20 versus 30; payment/match status is `BYTE` in payments but `TEXT(30)` in work; payment method/type naming and widths differ.
7. **Reported size change not reflected.** `tblImportPolicies.InsuredName` remains `TEXT(150)`, contrary to `PROJECT_STATUS.md` saying it was changed to Long Text.
8. **Bank relationship inconsistency.** `tblBanks.BankID` does not match `tblBankTransactions.BankAccountID`; no account table is present.
9. **Renewal inconsistency.** `ExpireDate` versus `EndDate`, and free-text `RenewalStatus` versus SQL_v5 `tblRenewalStatuses`.
10. **Audit convention inconsistency.** Some tables omit created/modified fields or only include one; import policies has neither general created/modified timestamps.

## 6. Reconciliation with project documentation

### Implemented from repository DDL

SQL_v4 directly represents 17 tables: `tblRoles`, `tblUsers`, `tblMarketers`, `tblInsuranceCompanies`, `tblBranches`, `tblInsuranceTypes`, `tblBanks`, `tblAccountHeads`, `tblSettings`, `tblCustomers`, `tblPolicies`, `tblInstallments`, `tblPayments`, `tblBankTransactions`, `tblFollowUps`, `tblRenewals`, and `tblImportPolicies`. SQL_v5 separately represents `tblRenewalStatuses` and `tblImportPolicies_Work`.

### Partially implemented

All 17 SQL_v4 tables are only **schema-source implemented** pending Access execution/live-schema verification. Tables with especially incomplete relational support are payments, bank transactions, follow-ups, renewals, and import policies. `tblInstallments` alone has secondary indexes, but still lacks a declared FK. The SQL_v5 work/status tables remain separate migration artifacts and do not resolve SQL_v4 relationships automatically.

### Referenced but missing from available DDL

`PROJECT_STATUS.md` reports `tblImportSessions`, `tblImportErrors`, and `tblImportLogs` as created, but no SQL_v4/v5 DDL for them is present. It lists `tblImportPayments`, `tblImportBank`, `tblUserLogs`, `tblPermissions`, and `tblRolePermissions` as planned; they are not implemented. No current repository source identifies exact filenames or intended objects for historical positions 018–023, and no such files are present. The task owner's report that the former local files were empty establishes only that these positions are known missing/unimplemented items; it is not authority to infer or create SQL for them.

### Comparison with the three baseline documents

| Baseline document | Comparison with the real SQL_v4 set |
| --- | --- |
| `PROJECT_STATUS.md` | Its 17 SQL_v4 table names are confirmed by DDL. The SQL_v5-only `tblRenewalStatuses` and `tblImportPolicies_Work` remain separate. The reported `tblImportSessions`, `tblImportErrors`, and `tblImportLogs` still have no available DDL. |
| `Documents/Progress/BIMYAR_GAP_ANALYSIS.md` | Its earlier statement that SQL_v4 contained only `.gitkeep` was accurate for that snapshot but is now superseded: 17 table definitions and five secondary indexes are inspectable. Its warnings about absent live-Access verification, declared relationships, broad index coverage, import-session DDL, VBA, samples, and business rules remain applicable. |
| `Documents/Progress/SQL_V4_INVENTORY.md` | The Step 3.4 zero-source snapshot is superseded by the current 18-file inventory. The inventory count, filenames, object count, duplicate 012 numbering, and missing 018–023 positions agree with this detailed analysis. |

The historical local files numbered 018 through 023 are reported by the task owner as having been empty. They are therefore treated as known missing/unimplemented items, not as SQL implementation evidence. They are absent from the current repository, their exact filenames and intended objects cannot be derived from the available sources, and no replacements are created or inferred in this step.

## 7. Highest-priority database fixes for a later implementation step

No fix is applied here. Recommended order after evidence and rules are approved:

1. Obtain/compare the sanitized Access schema export, actual relationships/indexes, and an authoritative run order; prove the scripts execute under Access 2016.
2. Define and enforce the relationship matrix, resolving bank account/import batch/session parents and delete/update/orphan policy.
3. Approve candidate-key and normalization rules for policy numbers, usernames/settings/codes, national code, mobile, and bank identifiers.
4. Add justified FK and query indexes, prioritizing policy/import-session paths and payment/bank reconciliation; measure against real queries and volumes.
5. Reconcile field sizes/types with Excel evidence to address error 3163 without truncation, especially insured name and national/phone data.
6. Normalize renewal status integration and reconcile bank/account naming/model.
7. Add explicit defaults/required validation only after application and migration behavior is agreed; avoid retroactive assumptions about null data.

## 8. Completion decision

**Phase 3 Step 3.5: Completed.** All 18 real SQL_v4 files currently present were inspected and the 17 represented tables, five indexes, constraints, dependencies, gaps, and inconsistencies were documented. This does **not** certify the live Access database, create missing scripts, repair SQL, complete Phase 3, or authorize Phase 4.

**Phase 3 remains In Progress.** Before closure: intake/verify the sanitized Access database and actual schema/relationships/indexes; verify SQL execution order and SQL_v5 completeness; intake remaining genuine VBA, Excel samples, and architecture/business-rule/diagnostic documents; reconcile live objects with sources; complete confidentiality/completeness checks and a baseline manifest; resolve documented discrepancies; and formally approve closure.

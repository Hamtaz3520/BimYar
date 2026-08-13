# Batch 1 — Protection and Read-only Baseline

**Work item:** `BMY-CX-20260813-012`

**Phase / step:** Phase 3 / Step 3.8

**Batch:** 1 — Protection and read-only preflight (`DB-01`, `DB-02`)

**Target:** Microsoft Access 2016

**Execution date:** 2026-08-13 UTC

**Gate result:** **BLOCKED — NOT PASSED**

## 1. Scope and safety boundary

This package defines the protection and evidence gate that must pass before any
schema-changing work. It authorizes only file protection, metadata capture, and
read-only inspection of an owner-approved sanitized production copy. It does
**not** authorize DDL, saved queries, compact/repair of the source, data updates,
normalization, truncation, deduplication, relationship or index changes, field
property changes, or SQL/VBA application-logic changes.

The repository preflight was executed on 2026-08-13. No `.accdb` or `.mdb` file,
sanitized workbook, live export, external backup location, Microsoft Access
runtime, or owner/privacy approval was available in the repository environment.
Consequently, none of the live baselines could be truthfully collected and the
gate is blocked. The package records that absence rather than substituting the
repository DDL for the live database.

The source-analysis basis is commit
`263c0a1f9e2ac05e28fdba238820ffc52ec6fe1f`. The 19 repository-represented tables
are a reconciliation target only; they are not a certification of the live
schema. Evidence containing production paths, names, identifiers, values, or
record keys must remain in the approved external evidence location and must not
be committed.

## 2. Entry controls and evidence convention

All entry controls must be satisfied before opening the database:

1. Record the database owner, operator, approver, ticket, confidentiality class,
   approved sanitized-copy path, external evidence path, and retention period.
2. Confirm all users are out of the source and no Access lock file remains.
   Record path, file name, byte size, UTC modified time, file format, and the
   SHA-256 digest without placing the database in this repository.
3. Confirm the operator is using Microsoft Access 2016 and record Office version,
   build, 32/64-bit bitness, ACE/DAO version, Windows version, locale, and Access
   trusted-location/security settings relevant to opening the copy.
4. Set the evidence directory to owner-only access. Use an immutable execution
   identifier of the form `BMY-CX-20260813-012_<UTC timestamp>` and retain a
   manifest containing relative evidence names, byte sizes, SHA-256 digests,
   producer, and timestamp.
5. Open only a disposable restored copy for inspection. Open it read-only, do not
   enable application startup actions/macros, and do not save any object or design
   change. Cancel immediately if Access offers to convert or repair the file.

Every result must use one of `PASS`, `FAIL`, `BLOCKED`, or `NOT_APPLICABLE`.
`PASS` requires an evidence reference and operator/reviewer sign-off. A zero
exception count is a measured result, never a blank cell. Exception evidence
should contain counts and privacy-safe row fingerprints; raw personal or banking
values are prohibited unless the approved evidence policy expressly permits them.

## 3. A–B — Backup and restore-test procedure

### A. Backup procedure

1. With the source closed and exclusively controlled, capture its metadata and
   SHA-256 digest (`source_before_sha256`).
2. Perform an operating-system byte-for-byte copy to the approved location
   outside both the repository and the source directory. Do not use Access
   Save-As, conversion, compact/repair, or export as the backup mechanism.
3. Hash the backup (`backup_sha256`) and compare byte size and SHA-256 to the
   source. Any mismatch is `FAIL`; stop without retrying against an open source.
4. Mark the backup read-only using the approved storage control, record access
   control/retention metadata, and write the two hashes to the external manifest.
5. Do not proceed to profiling until the restore test passes.

### B. Restore-test procedure

1. Copy the protected backup to a new, isolated restore-test directory; never
   restore over the source. Hash the restored file and require equality with the
   protected backup before opening it.
2. On the approved Access 2016 workstation, disable startup actions/macros and
   open the restored file read-only. Record whether Access reports corruption,
   conversion, repair, missing references, password, or workgroup errors.
3. Without saving, verify that the database window/catalog can enumerate tables,
   queries, forms, reports, modules, relationships, and linked-table definitions.
   Do not refresh or relink external tables during this test.
4. Close without saving, hash the restored file again, and require equality with
   the pre-open restored hash. Delete or retain the disposable copy according to
   the approved retention policy; preserve only the signed result and manifest.
5. A successful open is not enough: hash equality before and after, correct Access
   version/bitness, object enumeration, and two-person review are all required.

## 4. C–F — Structural and volume baselines

Collection must be performed from the disposable restored copy by a read-only
catalog/export method that creates no objects in the database. CSV or JSON is
preferred; each file must include the execution ID and be listed in the manifest.

### C. Schema baseline export

Export one row per table/field with: table and field name; ordinal; Access/DAO
type; declared size; precision/scale when exposed; Required, AllowZeroLength,
default, validation rule/text; AutoNumber/attributes; caption; and linked-table
source metadata with secrets removed. Separately inventory all object names/types
and reconcile the following 19 source-represented tables:

`tblRoles`, `tblUsers`, `tblMarketers`, `tblInsuranceCompanies`, `tblBranches`,
`tblInsuranceTypes`, `tblBanks`, `tblAccountHeads`, `tblSettings`, `tblCustomers`,
`tblPolicies`, `tblInstallments`, `tblPayments`, `tblBankTransactions`,
`tblFollowUps`, `tblRenewals`, `tblImportPolicies`, `tblRenewalStatuses`, and
`tblImportPolicies_Work`.

Record extra, missing, and differing objects explicitly. Also record whether the
reported but unverified `tblImportSessions`, `tblImportErrors`, and
`tblImportLogs` exist; do not infer their definitions from their names.

### D. Relationship baseline

Export every relation and field pair with relation name, parent/child table and
field ordinal, relationship attributes, referential-integrity enforcement, join
type where exposed, and cascade-update/delete flags. Zero relationships must be
recorded as an explicit count of zero. No relationship is to be created while
collecting this evidence.

### E. Index baseline

Export every table index and ordered index field with index name, primary,
unique, required, ignore-null, clustered/foreign attributes where exposed, and
sort direction. Reconcile the five repository-represented secondary indexes on
`tblInstallments`, but treat the live export as authoritative.

### F. Table row-count baseline

Record one read-only exact count per local table and, separately, each linked
table that can be read without refreshing credentials or links. Include success,
error, or `NOT_APPLICABLE`; never coerce an inaccessible link to zero. Repeat the
counts after all profiling and require equality. A difference blocks the gate
and triggers investigation from a new disposable restore, not data repair.

## 5. G–I — Read-only exception profiles

Execute transient parameterized `SELECT`/aggregate operations only. Do not save
queries in Access. Profiles must report table, field/rule, examined row count,
null count, blank count where applicable, exception count, minimum/maximum, and a
privacy-safe exception fingerprint. Raw values should not be exported by default.

### G. Error 3163 field-length profile

For `InsuredName`, `NationalCode`, `Phone`, and `Mobile` wherever present in
`tblImportPolicies`, `tblImportPolicies_Work`, `tblCustomers`, and
`tblMarketers`, capture declared capacity, maximum stored character length,
counts above each relevant destination capacity, null/blank counts, and a length
histogram. Also profile every actual source-to-destination text mapping discovered
in the genuine import module/export. Measure both stored length and the length
after the **currently implemented** mapping only; do not invent or apply a new
normalization rule. If the genuine mapping or sanitized failing workbook is not
available, mark mapping-specific results `BLOCKED`.

### H. Duplicate-key profile

Profile exact duplicate groups (including null/blank counts separately) for
declared live unique indexes and verified keys. At minimum report candidate
collisions for Username, SettingKey, policy-number/code fields, NationalCode,
Mobile, bank tracking/reference fields, and `(PolicyID, InstallmentNo)` where
those fields exist. Candidate profiles do **not** establish uniqueness and must
be labelled `CANDIDATE_ONLY` until scope, case/collation, normalization, null, and
blank rules receive business approval. Report group and affected-row counts.

### I. Orphan-record profile

For every exported live relationship and each repository-identified logical link,
count non-null child values with no parent. Include nullable-root handling for
self-links and do not treat null as an orphan. Required logical coverage includes
users/roles; branches/companies; insurance-type and account-head self-parents;
policy customer/marketer/company/branch/type; installments/policies; payment
policy/installment/bank-transaction; follow-up customer/policy/installment/user;
and renewal policy/customer/renewed-policy/user links. `BankAccountID` and
`ImportBatchID` must be reported as `BLOCKED_PARENT_UNVERIFIED`, not joined to an
invented parent.

## 6. J–L — Domain-specific and platform baselines

### J. Import-session baseline

Record existence, complete fields/indexes/relationships, row counts, date range,
distinct status values/counts, and participating-table session coverage for
`tblImportSessions`, `tblImportErrors`, and `tblImportLogs`. Record import/work
rows with null, missing, or orphan session identifiers only if an actual session
model exists. Capture genuine module/query references by name without changing
source. If objects or logic are absent, record `BLOCKED`/`NOT_PRESENT`; do not
design a replacement in Batch 1.

### K. Payment/bank baseline

For `tblPayments`, record row count and Currency sums, split by null/non-null
PolicyID, InstallmentID, and BankTransactionID; counts for neither/both attachment
patterns; link-orphan counts; MatchStatus/IsConfirmed distributions; and date
range. For `tblBankTransactions`, record row count, debit and credit Currency
sums, both/neither debit-credit counts, IsMatched distribution, date range,
duplicate tracking/reference candidate counts, and unresolved BankAccountID and
ImportBatchID coverage. Preserve exact Currency values in aggregate evidence;
do not export card/account numbers or infer that BankAccountID means BankID.

### L. Access version/bitness baseline

Record Microsoft Access product name, 2016 version/build, MSI or Click-to-Run,
32/64-bit bitness, executable path, database file format, ACE provider version,
DAO reference version, Windows version/architecture, locale/collation-relevant
settings, and whether the file is trusted/compiled. A non-Access substitute or a
different Access major version cannot pass the Access 2016 baseline.

## 7. Executed preflight record

| Control | Result | Evidence / reason |
|---|---|---|
| Repository source basis | PASS | Git commit `263c0a1f9e2ac05e28fdba238820ffc52ec6fe1f` inspected. |
| Safety boundary | PASS | Documentation only; no database artifact was opened and no DDL/DML was run. |
| Owner/privacy approval | BLOCKED | No signed approval or evidence policy was supplied. |
| Sanitized Access source | BLOCKED | No `.accdb` or `.mdb` exists in the repository workspace. |
| External backup/restore | BLOCKED | No approved source or external backup destination was supplied. |
| Access 2016/ACE baseline | BLOCKED | Microsoft Access/ACE is not available in this Linux execution environment. |
| C–K live exports/profiles | BLOCKED | Require the approved restored Access copy and read-only Access 2016 collector. |
| Batch 1 completion gate | **BLOCKED — NOT PASSED** | Required live evidence and sign-offs are absent. |

## 8. Exit gate and handoff

Batch 1 passes only when every required checklist item is `PASS`, source and
post-profile row counts agree, the backup/restore hashes agree, all A–L artifacts
are in the signed manifest, schema discrepancies and privacy-safe exception
counts are reviewed, and both the database owner and technical reviewer sign.

Until then, Phase 3 remains in progress and Batch 2, DDL, cleanup, and all schema
or application changes remain prohibited. To resume, supply an owner-approved
sanitized production copy on an Access 2016 workstation, approved external
evidence/backup locations, confidentiality rules, and named operator/reviewer.

| Role | Name | Decision | UTC timestamp | Signature/reference |
|---|---|---|---|---|
| Database owner | PENDING | BLOCKED | PENDING | PENDING |
| Technical reviewer | PENDING | BLOCKED | PENDING | PENDING |

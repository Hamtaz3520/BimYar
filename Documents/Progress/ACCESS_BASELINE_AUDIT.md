# Access Baseline Audit

**Work item:** `BMY-CX-20260813-013`  
**Phase / step:** Phase 3 / Step 3.9  
**Audit date (UTC):** 2026-08-13  
**Authoritative artifact:** `Database/Access/Baseline/BimYarCRM_v0.1.accdb`  
**Operating mode:** read-only evidence; no DDL, DML, object save, compact/repair, conversion, rename, move, delete, or replacement

## 1. Decision

The repository artifact was successfully verified at the filesystem level, but the Access database was **not successfully inspected at the Access object/schema level**. The environment has neither Microsoft Access/ACE/DAO nor LibreOffice, `mdbtools`, an Access ODBC driver, or an installed compatible Python parser. Attempts to obtain `mdbtools` and `access-parser` were blocked by the environment's package-network configuration. In accordance with the no-invention rule, binary-string observations were not treated as database evidence.

The Batch 1 gate is therefore **BLOCKED — NOT PASSED**. The artifact resolves the earlier “database not supplied” blocker, but it does not provide the required Access 2016 schema, object, property, row-profile, backup/restore, external-evidence, or sign-off evidence.

## 2. Artifact verification

| Check | Classification | Verified result |
| --- | --- | --- |
| Exact path and filename | VERIFIED | `Database/Access/Baseline/BimYarCRM_v0.1.accdb` |
| Repository readability | VERIFIED | Regular repository file; readable; mode `-rw-r--r--` |
| Non-zero size | VERIFIED | 1,069,056 bytes |
| SHA-256 | VERIFIED | `6e7f8f288b2f1e76c4717d3912104887a3f10afc2561d2fc775c87ec034f9088` |
| Access lock file | NOT PRESENT | No `*.laccdb` file existed in `Database/Access/Baseline/` at preflight or final integrity verification |
| Baseline unchanged by audit | VERIFIED | Final size and SHA-256 equal the initial values above |

Hashing and filesystem metadata reads do not constitute opening the database through Access. No backup, restored copy, or external evidence directory was supplied or created in this task.

## 3. Inspection capability and actual-object inventory

| Evidence requested | Classification | Result |
| --- | --- | --- |
| Tables and fields | NEEDS ACCESS 2016 VERIFICATION | Not inspectable in this environment; actual table count unknown |
| Field types/sizes, Required/nullable, defaults | NEEDS ACCESS 2016 VERIFICATION | Not inspectable |
| Primary keys and indexes | NEEDS ACCESS 2016 VERIFICATION | Not inspectable |
| Relationships, RI, and cascade settings | NEEDS ACCESS 2016 VERIFICATION | Not inspectable |
| Row counts and stored maximum lengths | NEEDS ACCESS 2016 VERIFICATION | Not inspectable; no data query was executed |
| Saved queries | NEEDS ACCESS 2016 VERIFICATION | Actual count unknown |
| Forms | NEEDS ACCESS 2016 VERIFICATION | Actual count unknown |
| Reports | NEEDS ACCESS 2016 VERIFICATION | Actual count unknown |
| VBA modules | NEEDS ACCESS 2016 VERIFICATION | Actual count unknown |

Consequently, the number of actual tables, queries, forms, reports, and modules found is **not available**, rather than zero.

## 4. Reconciliation with repository evidence

The following is a source expectation only; it is **not** a claim about objects in the `.accdb`:

- SQL_v4 contains 18 substantive SQL files representing 17 `CREATE TABLE` definitions, 17 primary keys, and five secondary indexes, all five on `tblInstallments`. The source contains no `FOREIGN KEY` and no `DEFAULT` clause.
- SQL_v5 adds definitions for `tblRenewalStatuses` and `tblImportPolicies_Work`. Thus SQL_v4/v5 represent 19 table definitions in total.
- `PROJECT_STATUS.md` reports `tblImportSessions`, `tblImportErrors`, and `tblImportLogs` as created but lacking repository DDL. Their actual existence and definitions remain unverified.
- The stabilization and MUST-fix plans correctly treat a real Access export and read-only profiles as prerequisites. Availability of the artifact changes the intake fact, but does not satisfy those prerequisites without a compatible inspector and Access 2016 runtime evidence.

No actual difference between Access and SQL_v4/v5 can be asserted. The possible differences previously recorded by the plans—including live relationships/indexes/properties, the three session objects, and post-DDL manual design changes—remain unresolved.

## 5. Named-object and 3163 preflight

| Object or field | Repository source expectation | Access classification |
| --- | --- | --- |
| `tblImportPolicies` | SQL_v4 definition present | NEEDS ACCESS 2016 VERIFICATION |
| `tblImportPolicies_Work` | SQL_v5 definition present | NEEDS ACCESS 2016 VERIFICATION |
| `tblRenewalStatuses` | SQL_v5 definition present | NEEDS ACCESS 2016 VERIFICATION |
| `tblImportSessions` | Reported in project status; no repository DDL | NEEDS ACCESS 2016 VERIFICATION |
| `tblImportErrors` | Reported in project status; no repository DDL | NEEDS ACCESS 2016 VERIFICATION |
| `tblImportLogs` | Reported in project status; no repository DDL | NEEDS ACCESS 2016 VERIFICATION |

For both repository definitions of `tblImportPolicies` and `tblImportPolicies_Work`, the declared expectations are `InsuredName TEXT(150)`, `NationalCode TEXT(20)`, `Phone TEXT(30)`, `Mobile TEXT(30)`, and `StartDate`/`EndDate`/`IssueDate DATETIME`. Those declarations do not prove the live Access field types, sizes, Required properties, or maximum stored/import values. Therefore every named 3163-related field is **NEEDS ACCESS 2016 VERIFICATION**, and the 3163 cause remains unresolved.

Existing Access indexes and relationships are likewise **NEEDS ACCESS 2016 VERIFICATION**. Only the source-level five `tblInstallments` indexes and absence of source-level foreign-key clauses are verified.

## 6. Batch 1 disposition and blockers

**Gate status: BLOCKED — NOT PASSED.** Verified evidence is limited to artifact identity, size, hash, readability, non-zero state, lack of a lock file, and preservation of the artifact. Exact blockers are:

1. No Microsoft Access 2016 runtime and no ACE/DAO metadata interface is available.
2. No compatible installed read-only `.accdb` inspection tool is available; package installation is blocked by HTTP 403 responses from configured package sources.
3. No byte-for-byte external backup, isolated restored copy, or approved external evidence location was supplied for controls A-02 through B-03 and sign-off.
4. No Access 2016 version/build/bitness, ACE/DAO version, locale, owner approval, operator, reviewer, or signed result is available.
5. Schema/object exports, row counts, maximum-length profiles, duplicate/orphan profiles, and financial baselines therefore could not be produced.

## 7. Exact recommended next task

On a workstation with Microsoft Access 2016 and compatible ACE/DAO, execute the existing Batch 1 checklist against an **isolated byte-for-byte copy** of this exact baseline after confirming the source SHA-256 above. Open only the copy read-only and without repair/conversion; export tables, fields/properties, PKs, indexes, relationships, queries/forms/reports/modules, and row counts; run the approved read-only 3163 length and integrity profiles; preserve privacy-safe evidence externally; verify pre/post hashes; and obtain owner/reviewer sign-off. Do not execute DDL/DML or begin Batch 2 unless every required gate item is verified and signed.

# SQL_v4 Source Inventory

**Work item:** `BMY-CX-20260813-009`
**Phase / step:** Phase 3 / Step 3.5
**Inspection date:** 2026-08-13
**Inspected location:** `Database/SQL/SQL_v4/`

## 1. Evidence rule

This inventory records the repository files actually present at inspection time. No SQL was reconstructed, invented, renamed, rewritten, or repaired. `.gitkeep` is a directory marker and is excluded from all real-source counts. Detailed field, key, relationship, index, compatibility, and inconsistency findings are in `SQL_V4_SCHEMA_ANALYSIS.md`.

## 2. Current authoritative inventory

**18 real SQL_v4 source files were inspected. They contain 17 `CREATE TABLE` definitions and one additional index-only script.** All 18 are non-empty, readable UTF-8 with BOM, and use a `.sql` extension.

| Seq. | Exact filename | Object/purpose | Kind |
| ---: | --- | --- | --- |
| 001 | `001_Create_tblRoles.sql` | `tblRoles` | Table |
| 002 | `002_Create_tblUsers.sql` | `tblUsers` | Table |
| 003 | `003_Create_tblMarketers.sql` | `tblMarketers` | Table |
| 004 | `004_Create_tblInsuranceCompanies.sql` | `tblInsuranceCompanies` | Table |
| 005 | `005_Create_tblBranches.sql` | `tblBranches` | Table |
| 006 | `006_Create_tblInsuranceTypes.sql` | `tblInsuranceTypes` | Table |
| 007 | `007_Create_tblBanks.sql` | `tblBanks` | Table |
| 008 | `008_Create_tblAccountHeads.sql` | `tblAccountHeads` | Table |
| 009 | `009_Create_tblSettings.sql` | `tblSettings` | Table |
| 010 | `010_Create_tblCustomers.sql` | `tblCustomers` | Table |
| 011 | `011_Create_tblPolicies.sql` | `tblPolicies` | Table |
| 012 | `012_Create_tblInstallments.sql` | `tblInstallments` | Table |
| 012 | `012_Create_tblInstallments -Indexes.sql` | Five `tblInstallments` indexes | Index-only |
| 013 | `013_Create_tblPayments.sql` | `tblPayments` | Table |
| 014 | `014_Create_tblBankTransactions.sql` | `tblBankTransactions` | Table |
| 015 | `015_Create_tblFollowUps.sql` | `tblFollowUps` | Table |
| 016 | `016_Create_tblRenewals.sql` | `tblRenewals` | Table |
| 017 | `017_Create_tblImportPolicies.sql` | `tblImportPolicies` | Table |

Directory marker: `.gitkeep` is present but is not SQL and is not included in the 18-file count.

## 3. Structure and numbering findings

- Sequence 001–017 is represented, but 012 is duplicated. The table must precede its index-only companion; filename sorting/run order alone is not a sufficient migration contract.
- The companion filename includes a space in `Installments -Indexes` and is the only secondary-index script. No table DDL contains an inline secondary index.
- The index file creates four non-unique indexes and one unique composite index. All 17 table files create a primary key.
- No 018–023 files exist in the current directory or inspected repository history. Per the task owner's historical evidence, the former local files at those sequence positions were empty; they are treated only as known missing/unimplemented items, are not counted as real sources, and no filename, object, or SQL content is inferred for them.
- `.gitkeep` remains harmless but no longer represents an empty directory.

## 4. Table reconciliation

The 17 SQL_v4 tables match 17 names in the “confirmed created” list in `PROJECT_STATUS.md`. SQL_v5 separately provides `tblRenewalStatuses` and `tblImportPolicies_Work`. Three status-listed created tables remain referenced without DDL: `tblImportSessions`, `tblImportErrors`, and `tblImportLogs`. Five tables remain explicitly planned without DDL: `tblImportPayments`, `tblImportBank`, `tblUserLogs`, `tblPermissions`, and `tblRolePermissions`.

The prior Step 3.4 inventory correctly recorded the earlier empty snapshot. This document now supersedes its zero-source count following upload of the real files.

## 5. Verification still required

- compare all DDL with the sanitized/live Access database, including field properties not expressible here;
- verify each script and the multi-statement index file executes under Access 2016;
- obtain an authoritative run order and identify any missing ALTER/relationship/index scripts;
- reconcile actual relationships, referential integrity, defaults, required properties, validation rules, and indexes;
- retain 018–023 as known missing/unimplemented historical positions unless authoritative non-empty artifacts are supplied; do not reconstruct them;
- complete SQL_v5 and remaining source intake.

## 6. Completion decision

**Phase 3 Step 3.5: Completed.** The current real SQL_v4 set—18 source files representing 17 tables—has been inventoried and analyzed. Phase 3 remains **In Progress**; this result does not certify the live database or complete remaining source intake.

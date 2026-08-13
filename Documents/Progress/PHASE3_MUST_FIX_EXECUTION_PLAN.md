# Phase 3 MUST-FIX Execution Plan

**Work item:** `BMY-CX-20260813-011-R1`

**Phase / step:** Phase 3 / Step 3.7-R1

**Target:** Microsoft Access 2016

**Authoritative plan:** `Documents/Progress/DATABASE_STABILIZATION_PLAN.md`

## 1. Reconciliation result and scope

This document translates the 14 **MUST FIX BEFORE PHASE 4** packages in the merged database stabilization plan into seven executable batches. It does not add a database finding or authorize a migration. The comparison basis was the authoritative stabilization plan, the SQL_v4 schema analysis, the repository gap analysis, and `PROJECT_STATUS.md`.

All **14 MUST-FIX items (DB-01 through DB-14) remain valid**, and all **seven implementation batches remain valid** after reconciliation. The batch boundaries below preserve the authoritative dependency order. The reconciliation makes the following planning clarifications:

1. Batch 1 is explicitly limited to protection, baseline discovery, and read-only profiling. It does not cleanse data or change schema.
2. Error 3163 work is separated from ImportSession design: DB-03 must prove the failing value and lossless destination capacity; DB-04 cannot proceed from reported table names alone.
3. Relationships and indexes are distinguished. A relationship means Access referential integrity; a normal index is not a relationship and does not establish uniqueness.
4. Payment relationship work is gated by attachment rules, while bank linkage is additionally gated by the unresolved account/batch model in DB-10.
5. DB-14 is a mandatory Access 2016 release gate, not merely a documentation review.

No SQL, VBA, Access database, or application logic is changed by this reconciliation.

## 2. Reconciled seven-batch register

| Batch | MUST items | Executable scope | Entry prerequisites | Completion gate |
|---|---|---|---|---|
| **1 — Protection and read-only preflight** | DB-01, DB-02 | Obtain exclusive-access metadata and a byte-for-byte backup outside the repository; prove the backup opens; export the live tables, fields, relationships, indexes, row counts, and validation results; reconcile the 19 source-represented tables; run privacy-safe orphan, duplicate, null, range, overlength, and domain profiles; create the signed baseline, migration manifest, and exception inventory. | Sanitized production copy or authoritative live export; owner-approved access, privacy-safe evidence format, backup location, and restore procedure. | Baseline discrepancies and exception counts/keys are recorded and signed. No data cleansing or structural change occurs. |
| **2 — Error 3163 capacity and mapping proof** | DB-03 | Reproduce error 3163 with a sanitized failing workbook; capture the exact source column, value, normalized length, mapped destination, and destination capacity; profile all relevant import/work/operational identity and contact lengths; propose only evidence-backed, lossless widening; test query/form/import binding compatibility. | Batch 1; failing sanitized fixture and expected result; actual mapping/module export; approved normalization; backup/restore rehearsal. | The root cause is demonstrated, no value is truncated, proposed capacity covers fixtures and live maxima, and a regression result is retained. Long Text is used only if evidence exceeds 255 and is not indexed. |
| **3 — ImportSession model verification and isolation design** | DB-04 | Intake and verify the actual `tblImportSessions`, `tblImportErrors`, and `tblImportLogs` DDL/live definitions and engine workflow; identify every participating staging/log table; then design `LONG ImportSessionID` fields, a verified session parent, no-cascade RI relationships, and normal child indexes. Define backfill, orphan, retention, cleanup/retry, and ownership rules before any Required property or FK. | Batch 1 exception evidence; actual DDL/live export and genuine engine sources; owner-approved retention, ownership, backfill, and orphan rules. | Session boundary and migration manifest are approved; legacy rows have an approved non-destructive disposition; no field, parent, or relationship is inferred from names alone. |
| **4 — Parent, policy, and installment integrity** | DB-05, DB-06, DB-07 | Validate and then add the no-cascade parent relationships for users/roles, branches/companies, insurance-type self-parent, and account-head self-parent; stabilize the five policy relationships and normal FK indexes; retain both installment indexes and add the policy RI relationship. Validate the company/branch pairing separately rather than inventing an unsupported composite FK. | Batches 1–3 as applicable; zero unresolved relevant orphans/duplicates; approved self-parent, user-role optionality, nullable marketer/branch, company-branch, and installment deletion rules; verified live indexes. | Actual Access RI flags and index definitions match the manifest; roots/nullable links behave as approved; no cascade delete exists; row counts and totals reconcile. |
| **5 — Dependent workflow, payment, and bank decisions** | DB-08, DB-09, DB-10 | Profile payment, follow-up, and renewal links; add only approved normal indexes and no-cascade relationships after parent stabilization. Stop bank FK work until the real bank-account/import-batch model and ownership/direction are approved. Quarantine rather than silently delete conflicting or unattached legacy records. | Batch 4 and DB-02 evidence; approved payment attachment, workflow ownership/deletion, duplicated-customer authority, bank/POS identity, account ownership/card/IBAN, batch/session, and reconciliation rules. | Payment policy/installment links have one approved semantic; bank linkage is either supported by verified parents/mapping or remains explicitly blocked; dependent records have signed exception dispositions. |
| **6 — Phase 4 lookup and matching indexes** | DB-11, DB-12, DB-13 | After normalization and collision profiling, add measured **normal, non-unique** indexes for policy/import lookup, identity/contact matching, and marketer matching. Add marketer RI only where unmatched/manual-assignment lifecycle permits it. Do not assume global policy-number, NationalCode, or Mobile uniqueness. | Batches 2–5 as applicable; measured query inventory; approved normalization, null/blank, renewal/insurer scope, shared contact/legal identity, and marketer/manual-assignment rules. | Index kind and field order match measured predicates; collision reports are reviewed; no speculative unique index or work-table RI is introduced. |
| **7 — Access 2016 rehearsal and Phase 4 gate** | DB-14 | Build the later implementation artifact from approved DB-01–13 outcomes and rehearse its exact order on a disposable copy with DAO/ACE used by Access 2016, one DDL statement at a time and within the safest supported transaction boundary. Verify RI/index properties, bracketed ambiguous names such as `[Status]`, compact/repair and reopen, actual forms/queries/import fixture, rollback, query performance, row counts, and financial totals. | Completed and approved outcomes for DB-01–13; exact ordered artifact; representative fixtures; duplicate-012 ordering resolved; rollback copy and sign-off owners. | Zero unexplained count/total changes and no new orphan/duplicate; restore is demonstrated; smoke/regression checks pass; signed Access 2016 result log approves the Phase 4 gate. |

## 3. Dependency and compatibility confirmations

### Runtime error 3163

DB-03 remains a MUST item. The reported widening of `InsuredName` and `NationalCode`, column-number mapping, and size diagnostics are status claims, not sufficient proof because the genuine modules, live schema, and failing workbook are absent. Batch 2 must identify the exact failing field/value and demonstrate a lossless correction. It must also reconcile the verified source differences: import/work `InsuredName` is `TEXT(150)`; NationalCode is 20 in import/work versus 10 operational; and phone/mobile are 30 versus 20. It may not assume that Long Text is the answer.

### Relationship prerequisites

Every RI change requires the Batch 1 live baseline and exception profile, compatible `LONG` child-to-`AUTOINCREMENT` parent types, verified parents, resolution of existing orphans, and approval of nullability/ownership/deletion semantics. Relationships use no cascade delete. DB-05 precedes DB-06; DB-06 precedes DB-07; those stabilized parents precede DB-08/09. DB-10 expressly forbids inventing a bank relationship.

### Index prerequisites

Live index definitions must first be exported. Normal indexes are used for FK and lookup performance unless an independently verified candidate-key rule supports uniqueness. Predicate evidence, normalization/collision reports, actual field types/lengths, ACE index-limit testing, field order for composites, and write/file-size measurements are required. Batch 6 therefore follows capacity, relationship, and business-rule decisions; it does not create speculative unique indexes.

### ImportSession prerequisites

The repository reports three session/error/log tables but contains no verified DDL for them, and the v4 import/work tables contain no `ImportSessionID`. Batch 3 requires actual definitions and engine workflow before designing the parent, `LONG` children, normal indexes, RI, Required properties, or backfill. Retention, session ownership, cleanup/retry, orphan handling, and legacy-row disposition remain business decisions.

### Payment and bank prerequisites

Payment indexes and relationships are not one indivisible action. Normal indexes may be proposed only after profiling; policy/installment RI waits for the approved rule for neither/one/both links and handling of conflicting legacy rows. Bank-transaction RI waits for DB-10 because no verified `tblBankAccounts` or import-batch parent exists and `BankAccountID` must not be treated as `BankID`. Bank/POS samples and reconciliation sign-off are mandatory.

### Access 2016 requirements

All implementation must use Access 2016-compatible DAO/ACE behavior: `LONG` FKs for AutoNumber parents, `CURRENCY` for exact money, awareness that `INTEGER` is 16-bit and `BYTE` is 0–255, brackets for ambiguous identifiers, tested text-index limits/collation, and an Access-verified mechanism for Required/default properties. Multi-statement files are not assumed atomic. Each DDL statement runs separately in the rehearsed order, followed by compact/reopen, application/import testing, reconciliation, and demonstrated rollback.

## 4. Business-confirmation blockers

The following approvals remain blockers; inclusion in the plan is not approval:

- Import-session retention, ownership, backfill, cleanup/retry, and orphan policy.
- Insurance-type/account-head self-parent roots and cycles; user-role optionality.
- Nullable policy marketer/branch behavior, company-branch consistency, and installment deletion handling.
- Whether payments require neither, exactly one, or both policy/installment links, and how conflicts are quarantined.
- Follow-up/renewal ownership and deletion rules, and whether duplicated `CustomerID` is authoritative or derived.
- The bank-account/import-batch model, identifier direction/ownership, card/IBAN rules, and reconciliation sign-off.
- Policy-number normalization, blank/null, renewal, and insurer/company uniqueness scope.
- NationalCode/Mobile normalization, shared-contact/legal-entity behavior, invalid values, and any future uniqueness.
- Marketer matching, unmatched/manual assignment, and whether import/work relationships are enforceable.

## 5. First-batch decision

**Batch 1 remains the smallest safe executable first batch.** DB-01 establishes what actually exists and the protected rollback/baseline evidence; DB-02 supplies the exception evidence required by every later capacity, relationship, Required, and uniqueness decision. Splitting either item into a structural batch would allow changes against an unverified live model or without knowing which existing rows would be rejected. Keeping them together is still read-only with respect to production data and schema, while the backup/export protocol protects against operational mistakes.

Batch 1 may begin only when its sanitized/live access, confidentiality, owner, and backup prerequisites are available. Until then it is the first planned batch but is blocked from execution.

## 6. Step status and exact next task

**Phase 3 Step 3.7-R1: Completed as a documentation reconciliation.** The authoritative 14-item order and the seven-batch execution mapping agree. Phase 3 itself remains **In Progress**, and Phase 4 remains blocked until every MUST prerequisite and acceptance gate is satisfied.

**Exact next implementation task:** execute **Batch 1 — Protection and read-only preflight (DB-01/DB-02)** on an owner-approved sanitized production copy: create and open-test the external byte-for-byte backup, export and sign the live schema/relationship/index baseline, then run and preserve the approved read-only exception profiles. Do not begin Batch 2 or deploy DDL until that gate is signed.

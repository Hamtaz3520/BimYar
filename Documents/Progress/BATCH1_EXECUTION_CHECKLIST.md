# Batch 1 Machine-readable Execution Checklist

This YAML document is the authoritative checklist instance for work item
`BMY-CX-20260813-013`. Enumerations are fixed so automation can reject blank or
ambiguous results. Evidence paths are relative to the approved external evidence
directory and must never contain production values or secrets.

```yaml
schema_version: "1.0"
work_item: "BMY-CX-20260813-013"
project: "BimYar"
phase: 3
step: "3.9"
batch: 1
title: "Real Access baseline read-only preflight"
target_runtime: "Microsoft Access 2016"
execution_date_utc: "2026-08-13"
source_commit: "b58767bd5c92c38f5daa4b3881e05ddb9ce2317a"
allowed_statuses: [VERIFIED, BLOCKED, NOT_PRESENT, NEEDS_ACCESS_2016_VERIFICATION]
overall_status: BLOCKED
gate_passed: false
safety:
  mode: READ_ONLY
  ddl_allowed: false
  dml_allowed: false
  source_object_saves_allowed: false
  compact_repair_source_allowed: false
  application_logic_changes_allowed: false
  raw_sensitive_evidence_in_repository_allowed: false
execution_context:
  database_owner: null
  operator: null
  reviewer: null
  approval_reference: null
  sanitized_source_path: "Database/Access/Baseline/BimYarCRM_v0.1.accdb"
  external_backup_path: null
  external_evidence_path: null
  access_version: null
  access_build: null
  access_bitness: null
  ace_version: null
  dao_version: null
  source_sha256: "6e7f8f288b2f1e76c4717d3912104887a3f10afc2561d2fc775c87ec034f9088"
  backup_sha256: null
  restore_preopen_sha256: null
  restore_postclose_sha256: null
items:
  - id: ENTRY-01
    requirement: "Owner access and privacy-safe evidence policy approved"
    required: true
    status: BLOCKED
    evidence: null
    note: "Approval was not supplied."
  - id: ENTRY-02
    requirement: "Design-only Access baseline database supplied as repository evidence"
    required: true
    status: VERIFIED
    evidence: "ACCESS_BASELINE_AUDIT.md#2-artifact-verification"
    note: "Exact artifact is present, non-zero, readable, and has no sibling .laccdb lock file."
  - id: A-01
    requirement: "Closed source metadata and SHA-256 captured"
    required: true
    status: VERIFIED
    evidence: "ACCESS_BASELINE_AUDIT.md#2-artifact-verification"
  - id: A-02
    requirement: "Byte-for-byte external backup created"
    required: true
    status: BLOCKED
    evidence: null
  - id: A-03
    requirement: "Source and backup byte size and SHA-256 match"
    required: true
    status: BLOCKED
    evidence: null
  - id: B-01
    requirement: "Backup restored to an isolated disposable path"
    required: true
    status: BLOCKED
    evidence: null
  - id: B-02
    requirement: "Restored copy opens read-only in Access 2016 without repair or conversion"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: B-03
    requirement: "Restore hashes match before open and after close-without-save"
    required: true
    status: BLOCKED
    evidence: null
  - id: C-01
    requirement: "Complete table/field/property schema baseline exported"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: C-02
    requirement: "Nineteen represented tables and three reported session objects reconciled"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: D-01
    requirement: "Relationship fields, RI, and cascade flags exported"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: E-01
    requirement: "Ordered index fields and index properties exported"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: F-01
    requirement: "Exact pre-profile local and readable linked-table counts captured"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: F-02
    requirement: "Exact post-profile counts equal pre-profile counts"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: G-01
    requirement: "3163-related declared-capacity and stored-length profiles captured"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: G-02
    requirement: "Genuine mapping-specific length profile captured without new normalization"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: H-01
    requirement: "Declared-key and candidate-only duplicate profiles captured"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: I-01
    requirement: "Live and logical-link orphan profiles captured without invented parents"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: J-01
    requirement: "Import session/error/log existence, definitions, counts, and coverage recorded"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: K-01
    requirement: "Payment linkage/status/count/Currency aggregate baseline captured"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: K-02
    requirement: "Bank matching/count/Currency aggregate and unresolved identifier baseline captured"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: L-01
    requirement: "Access 2016 version, build, bitness, ACE/DAO, OS, format, and locale recorded"
    required: true
    status: NEEDS_ACCESS_2016_VERIFICATION
    evidence: "ACCESS_BASELINE_AUDIT.md#3-inspection-capability-and-actual-object-inventory"
  - id: SAFE-01
    requirement: "No database artifact opened through Access and no DDL, DML, object save, or logic change performed"
    required: true
    status: VERIFIED
    evidence: "ACCESS_BASELINE_AUDIT.md#1-decision"
  - id: SIGN-01
    requirement: "External evidence manifest complete and SHA-256 verified"
    required: true
    status: BLOCKED
    evidence: null
  - id: SIGN-02
    requirement: "Database owner signed Batch 1 result"
    required: true
    status: BLOCKED
    evidence: null
  - id: SIGN-03
    requirement: "Technical reviewer signed Batch 1 result"
    required: true
    status: BLOCKED
    evidence: null
gate_rule:
  expression: "gate_passed == true only when every required item status == VERIFIED"
  current_evaluation: false
  next_action: "Use an isolated byte-for-byte copy on an Access 2016 workstation; execute remaining controls, preserve external evidence, verify hashes, and obtain owner/reviewer sign-off."
```

## Update rules

1. Do not delete checklist items. Change `BLOCKED` or `NEEDS_ACCESS_2016_VERIFICATION` only after attaching verified evidence.
2. Use `BLOCKED` when a prerequisite prevents execution; use `NOT_PRESENT` only when absence was directly verified; use `NEEDS_ACCESS_2016_VERIFICATION` when this environment cannot establish the Access result.
3. `NOT_PRESENT` is an evidence result, not a waiver of a required control.
4. Set `gate_passed: true` and `overall_status: VERIFIED` only when every required
   item is `VERIFIED`, the external manifest verifies, and both sign-offs exist.
5. Store sensitive artifacts outside Git. Commit only privacy-safe summaries,
   hashes, evidence identifiers, and the signed disposition.

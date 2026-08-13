# Batch 1 Machine-readable Execution Checklist

This YAML document is the authoritative checklist instance for work item
`BMY-CX-20260813-012`. Enumerations are fixed so automation can reject blank or
ambiguous results. Evidence paths are relative to the approved external evidence
directory and must never contain production values or secrets.

```yaml
schema_version: "1.0"
work_item: "BMY-CX-20260813-012"
project: "BimYar"
phase: 3
step: "3.8"
batch: 1
title: "Protection and read-only preflight"
target_runtime: "Microsoft Access 2016"
execution_date_utc: "2026-08-13"
source_commit: "263c0a1f9e2ac05e28fdba238820ffc52ec6fe1f"
allowed_statuses: [PASS, FAIL, BLOCKED, NOT_APPLICABLE]
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
  sanitized_source_path: null
  external_backup_path: null
  external_evidence_path: null
  access_version: null
  access_build: null
  access_bitness: null
  ace_version: null
  dao_version: null
  source_sha256: null
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
    requirement: "Owner-approved sanitized Access database supplied"
    required: true
    status: BLOCKED
    evidence: null
    note: "No .accdb or .mdb was present in the workspace."
  - id: A-01
    requirement: "Closed source metadata and SHA-256 captured"
    required: true
    status: BLOCKED
    evidence: null
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
    status: BLOCKED
    evidence: null
  - id: B-03
    requirement: "Restore hashes match before open and after close-without-save"
    required: true
    status: BLOCKED
    evidence: null
  - id: C-01
    requirement: "Complete table/field/property schema baseline exported"
    required: true
    status: BLOCKED
    evidence: null
  - id: C-02
    requirement: "Nineteen represented tables and three reported session objects reconciled"
    required: true
    status: BLOCKED
    evidence: null
  - id: D-01
    requirement: "Relationship fields, RI, and cascade flags exported"
    required: true
    status: BLOCKED
    evidence: null
  - id: E-01
    requirement: "Ordered index fields and index properties exported"
    required: true
    status: BLOCKED
    evidence: null
  - id: F-01
    requirement: "Exact pre-profile local and readable linked-table counts captured"
    required: true
    status: BLOCKED
    evidence: null
  - id: F-02
    requirement: "Exact post-profile counts equal pre-profile counts"
    required: true
    status: BLOCKED
    evidence: null
  - id: G-01
    requirement: "3163-related declared-capacity and stored-length profiles captured"
    required: true
    status: BLOCKED
    evidence: null
  - id: G-02
    requirement: "Genuine mapping-specific length profile captured without new normalization"
    required: true
    status: BLOCKED
    evidence: null
  - id: H-01
    requirement: "Declared-key and candidate-only duplicate profiles captured"
    required: true
    status: BLOCKED
    evidence: null
  - id: I-01
    requirement: "Live and logical-link orphan profiles captured without invented parents"
    required: true
    status: BLOCKED
    evidence: null
  - id: J-01
    requirement: "Import session/error/log existence, definitions, counts, and coverage recorded"
    required: true
    status: BLOCKED
    evidence: null
  - id: K-01
    requirement: "Payment linkage/status/count/Currency aggregate baseline captured"
    required: true
    status: BLOCKED
    evidence: null
  - id: K-02
    requirement: "Bank matching/count/Currency aggregate and unresolved identifier baseline captured"
    required: true
    status: BLOCKED
    evidence: null
  - id: L-01
    requirement: "Access 2016 version, build, bitness, ACE/DAO, OS, format, and locale recorded"
    required: true
    status: BLOCKED
    evidence: null
  - id: SAFE-01
    requirement: "No database artifact opened and no DDL, DML, object save, or logic change performed"
    required: true
    status: PASS
    evidence: "BATCH1_PROTECTION_BASELINE.md#7-executed-preflight-record"
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
  expression: "gate_passed == true only when every required item status == PASS"
  current_evaluation: false
  next_action: "Supply approved sanitized copy, Access 2016 workstation, external backup/evidence locations, approvals, operator, and reviewer; then execute A through L."
```

## Update rules

1. Do not delete checklist items. Change `BLOCKED` only after attaching evidence.
2. Use `FAIL` for an executed control whose acceptance condition was not met;
   use `BLOCKED` when a prerequisite prevents execution.
3. `NOT_APPLICABLE` is permitted only with reviewer-approved rationale and cannot
   be used for an item marked `required: true` unless this checklist is revised
   and approved.
4. Set `gate_passed: true` and `overall_status: PASS` only when every required
   item is `PASS`, the external manifest verifies, and both sign-offs exist.
5. Store sensitive artifacts outside Git. Commit only privacy-safe summaries,
   hashes, evidence identifiers, and the signed disposition.

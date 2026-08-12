# SQL_v4 Source Inventory

**Work item:** `BMY-CX-20260812-008`

**Phase / step:** Phase 3 / Step 3.4

**Inspection date:** 2026-08-12

**Inspected location:** `Database/SQL/SQL_v4/`

## 1. Scope and evidence rule

This inventory records only files physically present in the repository. Documentation references are comparison evidence, not proof that a source file exists. No SQL file was generated, reconstructed, inferred, renamed, rewritten, or repaired during this inspection.

The target directory and the repository-wide file list were inspected. File type, exact path, extension, byte size, emptiness, read access, and decoding were checked for every file in the target directory. Repository-wide filename and path checks were also used to look for misplaced SQL files and duplicate filenames.

## 2. VERIFIED REAL SQL_v4 FILES

**Total verified real SQL_v4 source files: 0.**

There are no SQL_v4 source files to list. In particular, no file with a `.sql` extension is present in `Database/SQL/SQL_v4/`. Consequently, there is no SQL_v4 source content whose purpose, encoding, or corruption can be assessed.

### Directory marker (not a real source file)

| Exact filename | Exact repository path | Extension | Byte size | Empty / non-empty | Readable / unreadable | Encoding / readability result | Obvious corruption | Purpose |
| --- | --- | --- | ---: | --- | --- | --- | --- | --- |
| `.gitkeep` | `Database/SQL/SQL_v4/.gitkeep` | No conventional extension (`Path.suffix` is empty) | 1 | Non-empty (one LF byte) | Readable | Valid UTF-8/ASCII newline; also decodes as UTF-8 with BOM handling; not a SQL source | None; it is a directory marker | Keeps the otherwise source-empty directory tracked by Git |

`.gitkeep` is explicitly excluded from the real-source count.

## 3. REFERENCED BUT NOT PRESENT

| Reference | Repository finding |
| --- | --- |
| `PROJECT_STATUS.md` describes `Database/SQL/SQL_v4` as the previous working collection and reference for existing scripts. | No real SQL_v4 script is present. The statement does not identify any exact filename and is not existence evidence. |
| `Documents/Progress/BIMYAR_GAP_ANALYSIS.md` refers to the complete/full SQL_v4 set and run order as intake still required. | The referenced set and run-order guide are not present. The gap analysis itself also records that the directory contains only `.gitkeep`. |
| `Documents/Progress/SOURCE_INTAKE_CHECKLIST.md` requests all current real SQL_v4 files and any available run-order guide. | The checklist does not name individual SQL_v4 files. No such file or guide is present. |
| `Documents/Progress/SOURCE_INTAKE_LOG.md` records earlier intake packages. | It records no received SQL_v4 source filename. |
| `PROJECT_STRUCTURE.md` defines `Database/SQL/` as the location for SQL scripts and explains `.gitkeep` as an empty-directory marker. | It provides structure only and does not name or prove the existence of a SQL_v4 source file. |

Because none of the reviewed documents gives an exact SQL_v4 filename, there is no defensible filename-level missing list. The authoritative finding is an **unnamed, uncounted SQL_v4 source set referenced for future intake but not present**.

The two real `.sql` files found elsewhere in the repository—`Database/SQL/SQL_v5/024_Create_tblRenewalStatuses.sql` and `Database/SQL/SQL_v5/026_Create_tblImportPolicies_Work.sql`—are SQL_v5 sources. They are not treated as SQL_v4 files or as proof of missing SQL_v4 counterparts.

## 4. NEEDS VERIFICATION

The following cannot be verified until the untouched authoritative source set or other direct evidence is received:

- the intended SQL_v4 filenames and total file count;
- whether any historical SQL_v4 source actually exists outside this repository;
- original byte sizes, encodings, readability, and content integrity;
- script purposes, dependencies, and execution order;
- whether a run-order guide exists;
- whether the complete historical set has been received.

These unknowns do not prevent completion of this repository-inventory step; they remain Phase 3 source-intake items.

## 5. NUMBERING / STRUCTURE ISSUES

- **Numbering gaps:** No SQL_v4 numbered source file exists, so a sequence cannot be established and specific missing numbers cannot be inferred. The entire numbering baseline remains unknown. The SQL_v5 numbers `024` and `026` must not be projected backward onto SQL_v4.
- **Duplicate filenames:** None among real SQL_v4 files, because the real-file set is empty. The repository-wide check found no duplicated SQL source filename.
- **Wrong extensions:** None in the SQL_v4 directory. `.gitkeep` has no conventional extension and is a directory marker, not a wrongly extended source.
- **Wrong-directory files:** No file elsewhere in the repository claims an SQL_v4 path or identity. The two repository SQL sources are correctly identifiable by their actual `SQL_v5` paths and were not reclassified.
- **Empty placeholders:** No zero-byte file exists in `Database/SQL/SQL_v4/`. The one-byte `.gitkeep` marker is placeholder infrastructure, not source.
- **Structural discrepancy:** `PROJECT_STATUS.md` describes SQL_v4 as a previous working/reference collection, while the repository contains no real files in that collection. This inventory preserves that distinction without attempting to resolve it by inference.

## 6. Completion decision

**Phase 3 Step 3.4: Completed.** The current repository SQL_v4 inventory was fully inspected and documented: zero real SQL_v4 sources are present, and `.gitkeep` is the only directory entry. Completion applies only to inventorying the present repository state; it does not mean the missing historical set was received, it does not close SQL_v4 intake, and it does not complete Phase 3.

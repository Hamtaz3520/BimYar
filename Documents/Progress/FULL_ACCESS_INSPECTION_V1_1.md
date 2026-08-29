# BimYar v1.1 - Full Access Inspection

Target database: `Database/Access/Working/BimYarCRM_v1.1.accdb`

## Purpose

Export a read-only inspection bundle from Microsoft Access so GitHub/ChatGPT can inspect the database without requiring direct binary `.accdb` parsing.

## Module

`Database/VBA/Modules/modFullAccessInspector.bas`

The module is intentionally written with ASCII-only comments/messages to avoid Persian code-page corruption in the VBA editor.

## Steps

1. Open `BimYarCRM_v1.1.accdb` from the trusted `Database` folder.
2. Open the VBA editor.
3. Import `modFullAccessInspector.bas` as a module.
4. Run `Debug -> Compile BimYarCRM_v1`.
5. Open the Immediate Window and run:

   `ExportBimYarFullInspection`

6. Choose an existing output folder, preferably the local BimYar project folder.
7. The module creates a timestamped folder named like:

   `BimYar_FullInspection_YYYYMMDD_HHMMSS`

8. Upload that entire folder to:

   `Database/Access/Inspection/v1.1/`

## Exported content

- Database summary
- User tables, fields, sizes, defaults and indexes
- Table record counts (no row contents)
- Relationships
- VBA References and broken-reference state when available
- Query SQL and Access text definitions
- VBA module text definitions
- Forms
- Reports
- Macros
- Short Text field maximum-length profile for diagnosing error 3163

## Safety

The inspector is read-only with respect to BimYar data and Access objects. It enumerates metadata, reads counts/lengths, and uses `Application.SaveAsText` to export object definitions. It does not run DELETE, UPDATE, INSERT, ALTER, DROP, or action queries.

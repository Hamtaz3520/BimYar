# Multi-Payment Import: Access 2016 Deployment and Validation

## Scope and schema decisions

This phase deliberately does not edit either `.accdb` file. The SQL and VBA must be
applied to a disposable local copy of `BimYarCRM_v1.1.accdb` before promotion.

The inspection shows that `tblImportPolicies` currently mixes policy fields and one
set of payment fields. Those legacy payment columns are retained for compatibility,
but the new importer does not populate them. Each policy is stored once and every
actual payment is stored in `tblImportPayments`.

`PaymentRefNo` is `TEXT(255)` and `PaymentDate` is `TEXT(10)`. This preserves long
identifiers as text and Persian `yyyy/mm/dd` dates without locale conversion.
`PaymentSequence` is unique within a policy. `ImportSessionID` is indexed for audit
queries. The policy/payment foreign key has referential integrity but no cascade
delete: deleting an import policy with payments must be an explicit, reviewed
operation. A relationship from payments to sessions is not added in this phase,
because the existing inspection contains no application-table relationships and a
second parent relationship should be introduced only after orphan data is audited.

`BalanceStatus TEXT(20)` is added to `tblImportPolicies`. The importer assigns
`Balanced` or `NeedsReview`; a mismatch also creates a non-sensitive
`tblImportErrors` row. Currency is retained because the inspected policy amount is
already Currency and the sample totals fit the Access Currency range.

## Local deployment in Access 2016

1. Back up the database and work only in a local disposable copy.
2. Open the copy exclusively in Access 2016.
3. For each file below, create a new query in SQL View, paste the **single** statement,
   and run it in this exact order:
   1. `Database/SQL/SQL_v5/027_Create_tblImportPayments.sql`
   2. `Database/SQL/SQL_v5/028_Index_tblImportPayments_PolicySequence.sql`
   3. `Database/SQL/SQL_v5/029_Index_tblImportPayments_ImportSessionID.sql`
   4. `Database/SQL/SQL_v5/030_Relation_tblImportPolicies_tblImportPayments.sql`
   5. `Database/SQL/SQL_v5/031_Add_tblImportPolicies_BalanceStatus.sql`
4. In the VBA editor, remove or rename the old `modImportEngine` and
   `modImportDiagnostic` modules. Use **File > Import File** to import the two `.bas`
   files from `Database/VBA/Modules/`.
5. Confirm **Tools > References** includes the Access database engine object library
   that provides DAO (normally already present in an `.accdb`). Excel is late-bound,
   so no Excel reference is required and the source has no Windows API declarations.
6. Run **Debug > Compile VBAProject**. Do not proceed if compilation fails.
7. Close and reopen the database, then compact and repair the disposable copy.

## Positive validation with the real issuance report

Run `ImportPolicyReport` from the Immediate Window. Select the real 22-column file
and enter its report metadata. After the completion message, run these Access queries.
Replace `[session]` with the new `ImportSessionID`.

```sql
SELECT SessionStatus, TotalRows, PolicyRows, InstallmentRows, ErrorRows, Notes
FROM tblImportSessions
WHERE ImportSessionID=[session];
```

Expected: `SessionStatus='Completed'`. For the known sample, policy count is 112 and
payment count is 133:

```sql
SELECT Count(*) AS PolicyCount
FROM tblImportPolicies
WHERE ImportSessionID=[session];
```

```sql
SELECT Count(*) AS PaymentCount
FROM tblImportPayments
WHERE ImportSessionID=[session];
```

Verify that no repeated header or total was imported:

```sql
SELECT ImportPolicyID, RowNo, PolicyNumber
FROM tblImportPolicies
WHERE ImportSessionID=[session]
  AND (PolicyNumber Is Null OR Trim(PolicyNumber)='');
```

Expected: zero rows. Verify sequences are complete and unique:

```sql
SELECT ImportPolicyID, Count(*) AS PaymentCount,
       Min(PaymentSequence) AS FirstSequence,
       Max(PaymentSequence) AS LastSequence
FROM tblImportPayments
WHERE ImportSessionID=[session]
GROUP BY ImportPolicyID
HAVING Min(PaymentSequence)<>1 OR Max(PaymentSequence)<>Count(*);
```

Expected: zero rows. Locate policies with multiple cash rows (the Unicode payment
label is entered directly in the Access query UI, not in VBA source):

```sql
SELECT ImportPolicyID, Count(*) AS CashRows, Sum(PaymentAmount) AS CashTotal
FROM tblImportPayments
WHERE ImportSessionID=[session] AND PaymentType='نقد'
GROUP BY ImportPolicyID
HAVING Count(*)>1;
```

Inspect at least one result and confirm its distinct card/reference identifiers were
preserved. Confirm installment classification:

```sql
SELECT ImportPaymentID, PaymentType, IsInstallment
FROM tblImportPayments
WHERE ImportSessionID=[session]
  AND ((PaymentType='قسط' AND IsInstallment=False)
    OR (PaymentType='نقد' AND IsInstallment=True));
```

Expected: zero rows. Confirm identifiers remain text and inspect 17-digit values:

```sql
SELECT PaymentRefNo, Len(PaymentRefNo) AS RefLength
FROM tblImportPayments
WHERE ImportSessionID=[session] AND Len(PaymentRefNo)>=17;
```

## Financial validation

```sql
SELECT P.ImportPolicyID, P.PolicyNumber, P.PremiumAmount,
       Sum(Nz(X.PaymentAmount,0)) AS PaymentTotal, P.BalanceStatus
FROM tblImportPolicies AS P LEFT JOIN tblImportPayments AS X
  ON P.ImportPolicyID=X.ImportPolicyID
WHERE P.ImportSessionID=[session]
GROUP BY P.ImportPolicyID, P.PolicyNumber, P.PremiumAmount, P.BalanceStatus
HAVING Nz(P.PremiumAmount,0)<>Sum(Nz(X.PaymentAmount,0))
    OR P.BalanceStatus<>'Balanced';
```

Expected for a fully balanced file: zero rows. For the known sample, both totals are
10,669,626,750 Rials:

```sql
SELECT Sum(PremiumAmount) AS PolicyTotal
FROM tblImportPolicies WHERE ImportSessionID=[session];
```

```sql
SELECT Sum(PaymentAmount) AS PaymentTotal
FROM tblImportPayments WHERE ImportSessionID=[session];
```

## Rollback and failed-session test

Make a test copy of the Excel file. In a payment continuation row, place more than
255 characters in `PaymentRefNo`, save it as text, and import it. This intentionally
exercises the pre-assignment size guard and Access error 3163 behavior.

1. Record the new session ID shown in `tblImportSessions`.
2. Confirm `SessionStatus='Failed'`.
3. Confirm the transaction left no policy or payment rows for that session:

```sql
SELECT Count(*) FROM tblImportPolicies WHERE ImportSessionID=[failed-session];
```

```sql
SELECT Count(*) FROM tblImportPayments WHERE ImportSessionID=[failed-session];
```

Both counts must be zero. Confirm one safe diagnostic exists:

```sql
SELECT ErrorCode, ErrorMessage, RowNo
FROM tblImportErrors
WHERE ImportSessionID=[failed-session];
```

Expected: error code `3163`, with field name, source row, actual length, and allowed
length, but no customer value. Also test a payment row before the first policy; it
must fail the session and roll back all policy/payment rows.

## Remaining risks and review points

- Excel stores only 15 significant digits in a numeric cell. A 17-digit reference
  must already be formatted/stored as text in the source workbook; no importer can
  recover digits Excel has previously rounded.
- Header detection relies on the report invariant that policy `RowNo` is numeric and
  repeated header labels are nonnumeric. Revalidate if the insurer changes layout.
- Blank `PaymentAmount` values are imported as Null and consequently make a policy
  `NeedsReview` unless the policy total is also zero.
- Unknown nonblank payment types are retained, classified as non-installment, and
  logged for review rather than discarded.
- Session creation is intentionally outside the data transaction so a fatal rollback
  can still leave an auditable `Failed` session. Policy and payment writes are inside
  one DAO Workspace transaction.

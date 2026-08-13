Attribute VB_Name = "modBaselineExporter"
Option Compare Database
Option Explicit

' BimYar read-only Access/DAO baseline exporter.
' This module only enumerates DAO metadata and opens SELECT recordsets. It never
' executes an action query, saves an Access object, or changes database content.

Private Const REPORT_PREFIX As String = "BimYar_Baseline_"

Public Sub ExportBimYarBaseline()
    Dim db As DAO.Database
    Dim reportText As String
    Dim folder As String
    Dim reportPath As String

    On Error GoTo ExportFailed
    Set db = CurrentDb

    folder = InputBox("Folder for the read-only baseline report:", _
                      "BimYar baseline exporter", CurrentProject.Path)
    If Len(Trim$(folder)) = 0 Then Exit Sub
    folder = RemoveTrailingSeparator(Trim$(folder))
    If Not FolderExists(folder) Then
        MsgBox "The selected folder does not exist. No report was written.", vbExclamation
        Exit Sub
    End If

    reportPath = folder & "\" & REPORT_PREFIX & _
                 Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    AddDatabaseInfo reportText
    AddTableInventory db, reportText
    AddFieldDefinitions db, reportText
    AddRelationships db, reportText
    AddIndexInventory db, reportText
    AddLengthProfile db, reportText
    AddRequiredTableCheck db, reportText
    AddDuplicateOrphanPrecheck db, reportText
    AddLine reportText, "END OF REPORT"

    WriteUtf8File reportPath, reportText
    MsgBox "Read-only baseline exported to:" & vbCrLf & reportPath, vbInformation
    Exit Sub

ExportFailed:
    MsgBox "Baseline export failed (" & CStr(Err.Number) & "): " & _
           Err.Description & vbCrLf & "No database changes were made.", vbCritical
End Sub

Private Sub AddDatabaseInfo(ByRef output As String)
    AddHeading output, "A. DATABASE INFO"
    AddPair output, "Database", CurrentProject.FullName
    AddPair output, "AccessVersion", Application.Version
    AddPair output, "VBAVersion", VbaVersionText()
#If Win64 Then
    AddPair output, "AccessBitness", "64-bit VBA host"
#ElseIf Win32 Then
    AddPair output, "AccessBitness", "32-bit VBA host"
#Else
    AddPair output, "AccessBitness", "Not determinable"
#End If
    AddPair output, "Timestamp", Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub AddTableInventory(ByVal db As DAO.Database, ByRef output As String)
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim indexes As String
    Dim primaryKey As String

    AddHeading output, "B. USER TABLE INVENTORY"
    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then
            indexes = ""
            primaryKey = "(none)"
            For Each idx In tdf.Indexes
                AppendList indexes, idx.Name
                If idx.Primary Then primaryKey = IndexFieldList(idx)
            Next idx
            If Len(indexes) = 0 Then indexes = "(none)"
            AddLine output, "Table=" & SafeText(tdf.Name) & _
                    " | Status=" & IIf(Len(tdf.Connect) > 0, "LINKED", "LOCAL") & _
                    " | RecordCount=" & ReadCount(db, tdf.Name) & _
                    " | PrimaryKey=" & SafeText(primaryKey) & _
                    " | Indexes=" & SafeText(indexes)
        End If
    Next tdf
End Sub

Private Sub AddFieldDefinitions(ByVal db As DAO.Database, ByRef output As String)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    AddHeading output, "C. FIELD DEFINITIONS"
    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then
            For Each fld In tdf.Fields
                AddLine output, "Table=" & SafeText(tdf.Name) & _
                        " | Field=" & SafeText(fld.Name) & _
                        " | DAOType=" & DaoTypeName(fld.Type) & " (" & CStr(fld.Type) & ")" & _
                        " | Size=" & CStr(fld.Size) & _
                        " | Required=" & PropertyText(fld, "Required", "N/A") & _
                        " | AllowZeroLength=" & PropertyText(fld, "AllowZeroLength", "N/A") & _
                        " | DefaultValue=" & SafeText(PropertyText(fld, "DefaultValue", "")) & _
                        " | ValidationRule=" & SafeText(PropertyText(fld, "ValidationRule", "")) & _
                        " | Indexed=" & FieldIndexedStatus(tdf, fld.Name)
            Next fld
        End If
    Next tdf
End Sub

Private Sub AddRelationships(ByVal db As DAO.Database, ByRef output As String)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim joins As String

    AddHeading output, "D. RELATIONSHIPS"
    If db.Relations.Count = 0 Then AddLine output, "(none)"
    For Each rel In db.Relations
        joins = ""
        For Each fld In rel.Fields
            AppendList joins, fld.Name & " -> " & fld.ForeignName
        Next fld
        AddLine output, "Relation=" & SafeText(rel.Name) & _
                " | Parent=" & SafeText(rel.Table) & _
                " | Child=" & SafeText(rel.ForeignTable) & _
                " | Fields=" & SafeText(joins) & _
                " | Attributes=" & CStr(rel.Attributes) & _
                " | EnforcedRI=" & YesNo((rel.Attributes And dbRelationDontEnforce) = 0) & _
                " | CascadeUpdate=" & YesNo((rel.Attributes And dbRelationUpdateCascade) <> 0) & _
                " | CascadeDelete=" & YesNo((rel.Attributes And dbRelationDeleteCascade) <> 0)
    Next rel
End Sub

Private Sub AddIndexInventory(ByVal db As DAO.Database, ByRef output As String)
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index

    AddHeading output, "E. INDEX INVENTORY"
    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then
            For Each idx In tdf.Indexes
                AddLine output, "Table=" & SafeText(tdf.Name) & _
                        " | Index=" & SafeText(idx.Name) & _
                        " | Fields=" & SafeText(IndexFieldList(idx)) & _
                        " | Primary=" & YesNo(idx.Primary) & _
                        " | Unique=" & YesNo(idx.Unique) & _
                        " | IgnoreNulls=" & YesNo(idx.IgnoreNulls) & _
                        " | Required=" & PropertyText(idx, "Required", "N/A")
            Next idx
        End If
    Next tdf
End Sub

Private Sub AddLengthProfile(ByVal db As DAO.Database, ByRef output As String)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    AddHeading output, "F. RUNTIME ERROR 3163 PROFILE"
    AddLine output, "Values are never printed; only lengths and a safe primary-key identifier are reported."
    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then
            For Each fld In tdf.Fields
                If IsProfileField(fld.Name) Then AddOneLengthProfile db, tdf, fld, output
            Next fld
        End If
    Next tdf
End Sub

Private Sub AddOneLengthProfile(ByVal db As DAO.Database, ByVal tdf As DAO.TableDef, _
                                ByVal fld As DAO.Field, ByRef output As String)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim idField As String
    Dim maxLength As Long
    Dim rowId As String
    Dim errorText As String

    On Error GoTo ProfileError
    idField = SinglePrimaryKeyField(tdf)
    sql = "SELECT TOP 1 "
    If Len(idField) > 0 Then sql = sql & Q(idField) & " AS RowIdentifier, "
    sql = sql & "Len(CStr(" & Q(fld.Name) & ")) AS ActualLength FROM " & Q(tdf.Name) & _
          " WHERE " & Q(fld.Name) & " Is Not Null ORDER BY Len(CStr(" & Q(fld.Name) & ")) DESC;"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        maxLength = 0
        rowId = "(no data)"
    Else
        maxLength = Nz(rs!ActualLength, 0)
        If Len(idField) > 0 Then
            rowId = SafeIdentifier(rs!RowIdentifier)
        Else
            rowId = "(not safely determinable: no single-field primary key)"
        End If
    End If
    rs.Close
    AddLine output, ProfileLine(tdf, fld, maxLength, rowId)
    Exit Sub
ProfileError:
    errorText = Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
    AddLine output, "Table=" & SafeText(tdf.Name) & " | Field=" & SafeText(fld.Name) & _
                    " | DefinedSize=" & CStr(fld.Size) & " | MaxLength=ERROR" & _
                    " | RowIdentifier=(unavailable) | Error=" & SafeText(errorText)
End Sub

Private Function VbaVersionText() As String
#If VBA7 Then
    VbaVersionText = "VBA 7"
#Else
    VbaVersionText = "VBA 6 or earlier"
#End If
End Function

Private Function ProfileLine(ByVal tdf As DAO.TableDef, ByVal fld As DAO.Field, _
                             ByVal maxLength As Long, ByVal rowId As String) As String
    ProfileLine = "Table=" & SafeText(tdf.Name) & " | Field=" & SafeText(fld.Name) & _
                  " | DefinedSize=" & CStr(fld.Size) & " | MaxLength=" & CStr(maxLength) & _
                  " | RowIdentifier=" & SafeText(rowId)
End Function

Private Sub AddRequiredTableCheck(ByVal db As DAO.Database, ByRef output As String)
    Dim requiredTables As Variant
    Dim item As Variant

    requiredTables = Array("tblCustomers", "tblPolicies", "tblInstallments", "tblPayments", _
        "tblBankTransactions", "tblRenewals", "tblRenewalStatuses", "tblImportPolicies", _
        "tblImportPolicies_Work", "tblImportSessions", "tblImportErrors", "tblImportLogs")
    AddHeading output, "G. REQUIRED TABLE CHECK"
    For Each item In requiredTables
        AddLine output, CStr(item) & "=" & IIf(TableExists(db, CStr(item)), "PRESENT", "NOT PRESENT")
    Next item
End Sub

Private Sub AddDuplicateOrphanPrecheck(ByVal db As DAO.Database, ByRef output As String)
    AddHeading output, "H. DUPLICATE / ORPHAN PRECHECK"
    AddLine output, "Counts only. Checks are limited to relationships/business rules evidenced by repository SQL and plans."

    AddDuplicateCheck db, output, "tblInstallments", Array("PolicyID", "InstallmentNo"), _
                      "Duplicate installment key (PolicyID, InstallmentNo)"

    AddOrphanCheck db, output, "tblUsers", "RoleID", "tblRoles", "RoleID"
    AddOrphanCheck db, output, "tblBranches", "CompanyID", "tblInsuranceCompanies", "CompanyID"
    AddOrphanCheck db, output, "tblInsuranceTypes", "ParentTypeID", "tblInsuranceTypes", "InsuranceTypeID"
    AddOrphanCheck db, output, "tblAccountHeads", "ParentAccountID", "tblAccountHeads", "AccountHeadID"
    AddOrphanCheck db, output, "tblPolicies", "CustomerID", "tblCustomers", "CustomerID"
    AddOrphanCheck db, output, "tblPolicies", "MarketerID", "tblMarketers", "MarketerID"
    AddOrphanCheck db, output, "tblPolicies", "CompanyID", "tblInsuranceCompanies", "CompanyID"
    AddOrphanCheck db, output, "tblPolicies", "BranchID", "tblBranches", "BranchID"
    AddOrphanCheck db, output, "tblPolicies", "InsuranceTypeID", "tblInsuranceTypes", "InsuranceTypeID"
    AddOrphanCheck db, output, "tblInstallments", "PolicyID", "tblPolicies", "PolicyID"
    AddOrphanCheck db, output, "tblPayments", "PolicyID", "tblPolicies", "PolicyID"
    AddOrphanCheck db, output, "tblPayments", "InstallmentID", "tblInstallments", "InstallmentID"
    AddOrphanCheck db, output, "tblPayments", "BankTransactionID", "tblBankTransactions", "TransactionID"
    AddOrphanCheck db, output, "tblFollowUps", "CustomerID", "tblCustomers", "CustomerID"
    AddOrphanCheck db, output, "tblFollowUps", "PolicyID", "tblPolicies", "PolicyID"
    AddOrphanCheck db, output, "tblFollowUps", "InstallmentID", "tblInstallments", "InstallmentID"
    AddOrphanCheck db, output, "tblFollowUps", "UserID", "tblUsers", "UserID"
    AddOrphanCheck db, output, "tblRenewals", "PolicyID", "tblPolicies", "PolicyID"
    AddOrphanCheck db, output, "tblRenewals", "CustomerID", "tblCustomers", "CustomerID"
    AddOrphanCheck db, output, "tblRenewals", "RenewedPolicyID", "tblPolicies", "PolicyID"
    AddOrphanCheck db, output, "tblRenewals", "UserID", "tblUsers", "UserID"
    AddOrphanCheck db, output, "tblImportPolicies", "MarketerID", "tblMarketers", "MarketerID"
    AddOrphanCheck db, output, "tblImportPolicies", "ImportUserID", "tblUsers", "UserID"
    AddOrphanCheck db, output, "tblImportPolicies_Work", "SourceImportPolicyID", "tblImportPolicies", "ImportPolicyID"
    AddOrphanCheck db, output, "tblImportPolicies_Work", "MarketerID", "tblMarketers", "MarketerID"
End Sub

Private Sub AddDuplicateCheck(ByVal db As DAO.Database, ByRef output As String, _
                              ByVal tableName As String, ByVal fields As Variant, ByVal label As String)
    Dim sql As String
    Dim fieldList As String
    Dim item As Variant
    If Not TableAndFieldsExist(db, tableName, fields) Then
        AddLine output, label & "=NOT CHECKED (table/field absent)"
        Exit Sub
    End If
    For Each item In fields
        AppendList fieldList, Q(CStr(item))
    Next item
    sql = "SELECT Count(*) AS N FROM (SELECT " & fieldList & " FROM " & Q(tableName) & _
          " GROUP BY " & fieldList & " HAVING Count(*) > 1) AS DuplicateGroups;"
    AddLine output, label & "=" & ScalarCount(db, sql)
End Sub

Private Sub AddOrphanCheck(ByVal db As DAO.Database, ByRef output As String, _
                           ByVal childTable As String, ByVal childField As String, _
                           ByVal parentTable As String, ByVal parentField As String)
    Dim sql As String
    Dim label As String
    label = "Orphans " & childTable & "." & childField & " -> " & parentTable & "." & parentField
    If Not FieldExists(db, childTable, childField) Or Not FieldExists(db, parentTable, parentField) Then
        AddLine output, label & "=NOT CHECKED (table/field absent)"
        Exit Sub
    End If
    sql = "SELECT Count(*) AS N FROM " & Q(childTable) & " AS C LEFT JOIN " & Q(parentTable) & _
          " AS P ON C." & Q(childField) & " = P." & Q(parentField) & _
          " WHERE C." & Q(childField) & " Is Not Null AND P." & Q(parentField) & " Is Null;"
    AddLine output, label & "=" & ScalarCount(db, sql)
End Sub

Private Function ScalarCount(ByVal db As DAO.Database, ByVal sql As String) As String
    Dim rs As DAO.Recordset
    On Error GoTo CountError
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    ScalarCount = CStr(rs!N)
    rs.Close
    Exit Function
CountError:
    ScalarCount = "ERROR: " & SafeText(Err.Description)
End Function

Private Function ReadCount(ByVal db As DAO.Database, ByVal tableName As String) As String
    ReadCount = ScalarCount(db, "SELECT Count(*) AS N FROM " & Q(tableName) & ";")
End Function

Private Function IsProfileField(ByVal fieldName As String) As Boolean
    Select Case LCase$(fieldName)
        Case "insuredname", "nationalcode", "phone", "mobile", "startdate", "enddate", "issuedate"
            IsProfileField = True
    End Select
End Function

Private Function IsUserTable(ByVal tdf As DAO.TableDef) As Boolean
    IsUserTable = (Left$(tdf.Name, 4) <> "MSys") And _
                  ((tdf.Attributes And dbSystemObject) = 0) And _
                  ((tdf.Attributes And dbHiddenObject) = 0)
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = db.TableDefs(tableName)
    TableExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function FieldExists(ByVal db As DAO.Database, ByVal tableName As String, _
                             ByVal fieldName As String) As Boolean
    Dim fld As DAO.Field
    On Error Resume Next
    Set fld = db.TableDefs(tableName).Fields(fieldName)
    FieldExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function TableAndFieldsExist(ByVal db As DAO.Database, ByVal tableName As String, _
                                     ByVal fields As Variant) As Boolean
    Dim item As Variant
    If Not TableExists(db, tableName) Then Exit Function
    For Each item In fields
        If Not FieldExists(db, tableName, CStr(item)) Then Exit Function
    Next item
    TableAndFieldsExist = True
End Function

Private Function SinglePrimaryKeyField(ByVal tdf As DAO.TableDef) As String
    Dim idx As DAO.Index
    For Each idx In tdf.Indexes
        If idx.Primary Then
            If idx.Fields.Count = 1 Then SinglePrimaryKeyField = idx.Fields(0).Name
            Exit Function
        End If
    Next idx
End Function

Private Function IndexFieldList(ByVal idx As DAO.Index) As String
    Dim fld As DAO.Field
    For Each fld In idx.Fields
        AppendList IndexFieldList, fld.Name & IIf((fld.Attributes And dbDescending) <> 0, " DESC", " ASC")
    Next fld
End Function

Private Function FieldIndexedStatus(ByVal tdf As DAO.TableDef, ByVal fieldName As String) As String
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim result As String
    For Each idx In tdf.Indexes
        For Each fld In idx.Fields
            If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
                AppendList result, idx.Name
                Exit For
            End If
        Next fld
    Next idx
    If Len(result) = 0 Then result = "NO" Else result = "YES (" & result & ")"
    FieldIndexedStatus = result
End Function

Private Function PropertyText(ByVal obj As Object, ByVal propertyName As String, _
                              ByVal unavailableText As String) As String
    Dim value As Variant
    On Error GoTo NotAvailable
    value = obj.Properties(propertyName).Value
    If IsNull(value) Then
        PropertyText = "Null"
    ElseIf VarType(value) = vbBoolean Then
        PropertyText = YesNo(CBool(value))
    Else
        PropertyText = CStr(value)
    End If
    Exit Function
NotAvailable:
    PropertyText = unavailableText
End Function

Private Function DaoTypeName(ByVal typeNumber As Integer) As String
    Select Case typeNumber
        Case dbBoolean: DaoTypeName = "Boolean"
        Case dbByte: DaoTypeName = "Byte"
        Case dbInteger: DaoTypeName = "Integer"
        Case dbLong: DaoTypeName = "Long"
        Case dbCurrency: DaoTypeName = "Currency"
        Case dbSingle: DaoTypeName = "Single"
        Case dbDouble: DaoTypeName = "Double"
        Case dbDate: DaoTypeName = "Date/Time"
        Case dbText: DaoTypeName = "Short Text"
        Case dbLongBinary: DaoTypeName = "Long Binary/OLE"
        Case dbMemo: DaoTypeName = "Long Text"
        Case dbGUID: DaoTypeName = "GUID"
        Case dbDecimal: DaoTypeName = "Decimal"
        Case dbBigInt: DaoTypeName = "BigInt"
        Case dbVarBinary: DaoTypeName = "VarBinary"
        Case Else: DaoTypeName = "Other"
    End Select
End Function

Private Function Q(ByVal identifier As String) As String
    Q = "[" & Replace(identifier, "]", "]]" ) & "]"
End Function

Private Function SafeText(ByVal value As String) As String
    SafeText = Replace(Replace(value, vbCr, " "), vbLf, " ")
End Function

Private Function SafeIdentifier(ByVal value As Variant) As String
    If IsNull(value) Then
        SafeIdentifier = "Null"
    ElseIf VarType(value) = vbString And Len(CStr(value)) > 64 Then
        SafeIdentifier = "(text identifier omitted; length=" & CStr(Len(CStr(value))) & ")"
    Else
        SafeIdentifier = CStr(value)
    End If
End Function

Private Function YesNo(ByVal value As Boolean) As String
    YesNo = IIf(value, "YES", "NO")
End Function

Private Sub AppendList(ByRef target As String, ByVal item As String)
    If Len(target) > 0 Then target = target & ", "
    target = target & item
End Sub

Private Sub AddHeading(ByRef output As String, ByVal heading As String)
    If Len(output) > 0 Then output = output & vbCrLf
    AddLine output, String$(72, "=")
    AddLine output, heading
    AddLine output, String$(72, "=")
End Sub

Private Sub AddPair(ByRef output As String, ByVal key As String, ByVal value As String)
    AddLine output, key & "=" & SafeText(value)
End Sub

Private Sub AddLine(ByRef output As String, ByVal value As String)
    output = output & value & vbCrLf
End Sub

Private Function RemoveTrailingSeparator(ByVal folder As String) As String
    Do While Len(folder) > 3 And (Right$(folder, 1) = "\" Or Right$(folder, 1) = "/")
        folder = Left$(folder, Len(folder) - 1)
    Loop
    RemoveTrailingSeparator = folder
End Function

Private Function FolderExists(ByVal folder As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folder) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Sub WriteUtf8File(ByVal filePath As String, ByVal contents As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText contents
    stream.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    stream.Close
End Sub

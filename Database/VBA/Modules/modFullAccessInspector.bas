Option Compare Database
Option Explicit

' BimYar Full Access Inspector
' Read-only exporter for Access 2016+
' Does not modify database objects or table data.

Public Sub ExportBimYarFullInspection()
    Dim root As String
    Dim outFolder As String
    Dim db As DAO.Database

    On Error GoTo FatalError

    root = InputBox("Output folder for BimYar full inspection:", _
                    "BimYar Full Inspector", CurrentProject.Path)
    If Len(Trim$(root)) = 0 Then Exit Sub
    root = RemoveTrailingSlash(Trim$(root))

    If Not FolderExists(root) Then
        MsgBox "Folder does not exist.", vbExclamation, "BimYar Full Inspector"
        Exit Sub
    End If

    outFolder = root & "\BimYar_FullInspection_" & Format$(Now, "yyyymmdd_hhnnss")
    MkDir outFolder
    MkDir outFolder & "\Modules"
    MkDir outFolder & "\Forms"
    MkDir outFolder & "\Reports"
    MkDir outFolder & "\Macros"
    MkDir outFolder & "\Queries"

    Set db = CurrentDb

    ExportSummary db, outFolder
    ExportTables db, outFolder
    ExportRelations db, outFolder
    ExportReferences outFolder
    ExportQueries db, outFolder
    ExportAccessObjects outFolder
    ExportFieldLengthProfile db, outFolder

    MsgBox "Full inspection exported to:" & vbCrLf & outFolder, _
           vbInformation, "BimYar Full Inspector"
    Exit Sub

FatalError:
    MsgBox "Inspector failed: " & Err.Number & " - " & Err.Description, _
           vbCritical, "BimYar Full Inspector"
End Sub

Private Sub ExportSummary(ByVal db As DAO.Database, ByVal outFolder As String)
    Dim s As String
    s = "BimYar Full Inspection" & vbCrLf & _
        "======================" & vbCrLf & _
        "Database=" & CurrentProject.FullName & vbCrLf & _
        "AccessVersion=" & Application.Version & vbCrLf & _
        "Timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
        "Tables=" & CountUserTables(db) & vbCrLf & _
        "Queries=" & CountUserQueries(db) & vbCrLf & _
        "Modules=" & CurrentProject.AllModules.Count & vbCrLf & _
        "Forms=" & CurrentProject.AllForms.Count & vbCrLf & _
        "Reports=" & CurrentProject.AllReports.Count & vbCrLf & _
        "Macros=" & CurrentProject.AllMacros.Count & vbCrLf
    WriteUtf8 outFolder & "\00_SUMMARY.txt", s
End Sub

Private Sub ExportTables(ByVal db As DAO.Database, ByVal outFolder As String)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim s As String

    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then
            s = s & vbCrLf & "TABLE: " & tdf.Name & vbCrLf
            s = s & "Linked=" & IIf(Len(tdf.Connect) > 0, "YES", "NO") & vbCrLf
            s = s & "RecordCount=" & SafeCount(db, tdf.Name) & vbCrLf

            For Each fld In tdf.Fields
                s = s & "FIELD|" & fld.Name & _
                    "|Type=" & DaoTypeName(fld.Type) & _
                    "|Size=" & fld.Size & _
                    "|Required=" & PropText(fld, "Required", "N/A") & _
                    "|AllowZeroLength=" & PropText(fld, "AllowZeroLength", "N/A") & _
                    "|DefaultValue=" & CleanText(PropText(fld, "DefaultValue", "")) & _
                    "|ValidationRule=" & CleanText(PropText(fld, "ValidationRule", "")) & vbCrLf
            Next fld

            For Each idx In tdf.Indexes
                s = s & "INDEX|" & idx.Name & _
                    "|Fields=" & IndexFields(idx) & _
                    "|Primary=" & YesNo(idx.Primary) & _
                    "|Unique=" & YesNo(idx.Unique) & vbCrLf
            Next idx
        End If
    Next tdf

    WriteUtf8 outFolder & "\01_TABLES_SCHEMA.txt", s
End Sub

Private Sub ExportRelations(ByVal db As DAO.Database, ByVal outFolder As String)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim s As String
    Dim joins As String

    For Each rel In db.Relations
        joins = ""
        For Each fld In rel.Fields
            If Len(joins) > 0 Then joins = joins & ", "
            joins = joins & fld.Name & "->" & fld.ForeignName
        Next fld

        s = s & "RELATION|" & rel.Name & _
            "|Parent=" & rel.Table & _
            "|Child=" & rel.ForeignTable & _
            "|Fields=" & joins & _
            "|Attributes=" & rel.Attributes & vbCrLf
    Next rel

    WriteUtf8 outFolder & "\02_RELATIONS.txt", s
End Sub

Private Sub ExportReferences(ByVal outFolder As String)
    Dim ref As Object
    Dim s As String

    On Error Resume Next
    For Each ref In Application.References
        Err.Clear
        s = s & "REFERENCE|Name=" & SafeRefValue(ref, "Name") & _
            "|FullPath=" & SafeRefValue(ref, "FullPath") & _
            "|GUID=" & SafeRefValue(ref, "GUID") & _
            "|Major=" & SafeRefValue(ref, "Major") & _
            "|Minor=" & SafeRefValue(ref, "Minor") & _
            "|IsBroken=" & SafeRefValue(ref, "IsBroken") & vbCrLf
    Next ref
    On Error GoTo 0

    WriteUtf8 outFolder & "\03_REFERENCES.txt", s
End Sub

Private Sub ExportQueries(ByVal db As DAO.Database, ByVal outFolder As String)
    Dim qdf As DAO.QueryDef
    Dim p As String

    For Each qdf In db.QueryDefs
        If Left$(qdf.Name, 1) <> "~" Then
            p = outFolder & "\Queries\" & SafeFileName(qdf.Name) & ".sql.txt"
            WriteUtf8 p, qdf.SQL
            On Error Resume Next
            Application.SaveAsText acQuery, qdf.Name, _
                outFolder & "\Queries\" & SafeFileName(qdf.Name) & ".access.txt"
            On Error GoTo 0
        End If
    Next qdf
End Sub

Private Sub ExportAccessObjects(ByVal outFolder As String)
    Dim ao As AccessObject

    For Each ao In CurrentProject.AllModules
        ExportOneObject acModule, ao.Name, outFolder & "\Modules"
    Next ao

    For Each ao In CurrentProject.AllForms
        ExportOneObject acForm, ao.Name, outFolder & "\Forms"
    Next ao

    For Each ao In CurrentProject.AllReports
        ExportOneObject acReport, ao.Name, outFolder & "\Reports"
    Next ao

    For Each ao In CurrentProject.AllMacros
        ExportOneObject acMacro, ao.Name, outFolder & "\Macros"
    Next ao
End Sub

Private Sub ExportOneObject(ByVal objectType As AcObjectType, _
                            ByVal objectName As String, _
                            ByVal targetFolder As String)
    Dim p As String
    p = targetFolder & "\" & SafeFileName(objectName) & ".txt"

    On Error Resume Next
    Err.Clear
    Application.SaveAsText objectType, objectName, p
    If Err.Number <> 0 Then
        WriteUtf8 targetFolder & "\ERROR_" & SafeFileName(objectName) & ".txt", _
                  CStr(Err.Number) & " - " & Err.Description
    End If
    On Error GoTo 0
End Sub

Private Sub ExportFieldLengthProfile(ByVal db As DAO.Database, ByVal outFolder As String)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim s As String

    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then
            For Each fld In tdf.Fields
                If fld.Type = dbText Then
                    s = s & "FIELD|Table=" & tdf.Name & _
                        "|Field=" & fld.Name & _
                        "|DefinedSize=" & fld.Size & _
                        "|MaxActualLength=" & MaxFieldLength(db, tdf.Name, fld.Name) & vbCrLf
                End If
            Next fld
        End If
    Next tdf

    WriteUtf8 outFolder & "\04_SHORTTEXT_LENGTH_PROFILE.txt", s
End Sub

Private Function MaxFieldLength(ByVal db As DAO.Database, _
                                ByVal tableName As String, _
                                ByVal fieldName As String) As String
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo Fail
    sql = "SELECT Max(Len(CStr(" & Q(fieldName) & "))) AS M FROM " & Q(tableName) & ";"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If IsNull(rs!M) Then
        MaxFieldLength = "0"
    Else
        MaxFieldLength = CStr(rs!M)
    End If
    rs.Close
    Exit Function
Fail:
    MaxFieldLength = "ERROR:" & Err.Number
End Function

Private Function CountUserTables(ByVal db As DAO.Database) As Long
    Dim tdf As DAO.TableDef
    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then CountUserTables = CountUserTables + 1
    Next tdf
End Function

Private Function CountUserQueries(ByVal db As DAO.Database) As Long
    Dim qdf As DAO.QueryDef
    For Each qdf In db.QueryDefs
        If Left$(qdf.Name, 1) <> "~" Then CountUserQueries = CountUserQueries + 1
    Next qdf
End Function

Private Function SafeCount(ByVal db As DAO.Database, ByVal tableName As String) As String
    Dim rs As DAO.Recordset
    On Error GoTo Fail
    Set rs = db.OpenRecordset("SELECT Count(*) AS N FROM " & Q(tableName) & ";", dbOpenSnapshot)
    SafeCount = CStr(rs!N)
    rs.Close
    Exit Function
Fail:
    SafeCount = "ERROR:" & Err.Number
End Function

Private Function IsUserTable(ByVal tdf As DAO.TableDef) As Boolean
    IsUserTable = (Left$(tdf.Name, 4) <> "MSys") And _
                  ((tdf.Attributes And dbSystemObject) = 0) And _
                  ((tdf.Attributes And dbHiddenObject) = 0)
End Function

Private Function DaoTypeName(ByVal n As Integer) As String
    Select Case n
        Case dbBoolean: DaoTypeName = "Boolean"
        Case dbByte: DaoTypeName = "Byte"
        Case dbInteger: DaoTypeName = "Integer"
        Case dbLong: DaoTypeName = "Long"
        Case dbCurrency: DaoTypeName = "Currency"
        Case dbSingle: DaoTypeName = "Single"
        Case dbDouble: DaoTypeName = "Double"
        Case dbDate: DaoTypeName = "DateTime"
        Case dbText: DaoTypeName = "ShortText"
        Case dbMemo: DaoTypeName = "LongText"
        Case dbGUID: DaoTypeName = "GUID"
        Case Else: DaoTypeName = "Other(" & n & ")"
    End Select
End Function

Private Function PropText(ByVal obj As Object, ByVal propName As String, _
                          ByVal fallback As String) As String
    Dim v As Variant
    On Error GoTo Fail
    v = obj.Properties(propName).Value
    If IsNull(v) Then
        PropText = "Null"
    Else
        PropText = CStr(v)
    End If
    Exit Function
Fail:
    PropText = fallback
End Function

Private Function SafeRefValue(ByVal ref As Object, ByVal propName As String) As String
    Dim v As Variant
    On Error GoTo Fail
    v = CallByName(ref, propName, VbGet)
    If IsNull(v) Then
        SafeRefValue = "Null"
    Else
        SafeRefValue = CStr(v)
    End If
    Exit Function
Fail:
    SafeRefValue = "UNAVAILABLE"
End Function

Private Function IndexFields(ByVal idx As DAO.Index) As String
    Dim fld As DAO.Field
    For Each fld In idx.Fields
        If Len(IndexFields) > 0 Then IndexFields = IndexFields & ","
        IndexFields = IndexFields & fld.Name
    Next fld
End Function

Private Function Q(ByVal s As String) As String
    Q = "[" & Replace(s, "]", "]]" ) & "]"
End Function

Private Function YesNo(ByVal b As Boolean) As String
    If b Then YesNo = "YES" Else YesNo = "NO"
End Function

Private Function CleanText(ByVal s As String) As String
    CleanText = Replace(Replace(s, vbCr, " "), vbLf, " ")
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant
    Dim item As Variant
    bad = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    SafeFileName = s
    For Each item In bad
        SafeFileName = Replace(SafeFileName, CStr(item), "_")
    Next item
End Function

Private Function RemoveTrailingSlash(ByVal s As String) As String
    Do While Len(s) > 3 And (Right$(s, 1) = "\" Or Right$(s, 1) = "/")
        s = Left$(s, Len(s) - 1)
    Loop
    RemoveTrailingSlash = s
End Function

Private Function FolderExists(ByVal s As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(s) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Sub WriteUtf8(ByVal filePath As String, ByVal contents As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText contents
    stm.SaveToFile filePath, 2
    stm.Close
End Sub

Attribute VB_Name = "modImportDiagnostic"
Option Compare Database
Option Explicit

Public Sub TestImportFieldSizes()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb
    Set tdf = db.TableDefs("tblImportPolicies")

    Debug.Print "===== tblImportPolicies Field Sizes ====="
    For Each fld In tdf.Fields
        If fld.Type = dbText Then
            Debug.Print fld.Name & " -> " & fld.Size
        Else
            Debug.Print fld.Name & " -> Type=" & fld.Type
        End If
    Next fld

    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
    MsgBox "Inspection completed. See the Immediate Window.", vbInformation, "BimYar"
End Sub

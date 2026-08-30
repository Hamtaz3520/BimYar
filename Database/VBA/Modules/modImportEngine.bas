Attribute VB_Name = "modImportEngine"
Option Compare Database
Option Explicit

Private Const REPORT_COLUMN_COUNT As Long = 22
Private Const ERR_TEXT_TOO_LONG As Long = 3163

Private mErrorField As String
Private mErrorRow As Long
Private mErrorActualLength As Long
Private mErrorAllowedLength As Long

Public Sub ImportPolicyReport()
    Const msoFileDialogFilePicker As Long = 3

    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rsPolicy As DAO.Recordset
    Dim rsPayment As DAO.Recordset
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim fd As Object
    Dim filePath As String
    Dim fileName As String
    Dim reportType As Integer
    Dim reportYear As Integer
    Dim reportMonth As Integer
    Dim marketerID As Variant
    Dim sessionID As Long
    Dim currentPolicyID As Long
    Dim paymentSequence As Long
    Dim lastRow As Long
    Dim sourceRow As Long
    Dim sourceRows As Long
    Dim policyRows As Long
    Dim paymentRows As Long
    Dim installmentRows As Long
    Dim errorRows As Long
    Dim transactionStarted As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo ErrHandler

    reportType = Val(InputBox("Report type (1=office, 2=marketer):", "BimYar Import"))
    If reportType <> 1 And reportType <> 2 Then Exit Sub
    reportYear = Val(InputBox("Report year (for example 1405):", "BimYar Import"))
    If reportYear = 0 Then Exit Sub
    reportMonth = Val(InputBox("Report month (1-12):", "BimYar Import"))
    If reportMonth < 1 Or reportMonth > 12 Then Exit Sub

    If reportType = 2 Then
        marketerID = Val(InputBox("MarketerID:", "BimYar Import"))
        If marketerID = 0 Then Exit Sub
    Else
        marketerID = Null
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the Iran Insurance issuance report"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    If Len(Dir$(filePath)) = 0 Then
        MsgBox "The selected file does not exist.", vbExclamation, "BimYar Import"
        Exit Sub
    End If

    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    sessionID = CreateImportSession(db, reportYear, reportMonth, reportType, _
                                    marketerID, fileName, filePath)

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlBook = xlApp.Workbooks.Open(filePath, False, True)
    Set xlSheet = xlBook.Worksheets(1)
    lastRow = xlSheet.UsedRange.Row + xlSheet.UsedRange.Rows.Count - 1

    ws.BeginTrans
    transactionStarted = True
    Set rsPolicy = db.OpenRecordset("tblImportPolicies", dbOpenDynaset)
    Set rsPayment = db.OpenRecordset("tblImportPayments", dbOpenDynaset)

    For sourceRow = 1 To lastRow
        If Not IsEmptyReportRow(xlSheet, sourceRow) Then
            If Not IsHeaderRow(xlSheet, sourceRow) And _
               Not IsTotalRow(xlSheet, sourceRow) Then
                sourceRows = sourceRows + 1

                If Len(GetCellText(xlSheet, sourceRow, 2)) > 0 Then
                    If currentPolicyID > 0 Then
                        ValidatePolicyBalance db, sessionID, currentPolicyID, _
                                              sourceRow - 1, errorRows
                    End If
                    currentPolicyID = InsertPolicy(rsPolicy, sessionID, reportType, _
                                                   marketerID, fileName, xlSheet, sourceRow)
                    policyRows = policyRows + 1
                    paymentSequence = 0
                End If

                If HasPaymentData(xlSheet, sourceRow) Then
                    If currentPolicyID = 0 Then
                        Err.Raise vbObjectError + 2101, "ImportPolicyReport", _
                                  "Payment row has no preceding policy at source row " & sourceRow
                    End If
                    paymentSequence = paymentSequence + 1
                    InsertPayment rsPayment, sessionID, currentPolicyID, _
                                  paymentSequence, xlSheet, sourceRow, errorRows, db
                    paymentRows = paymentRows + 1
                    If IsInstallmentType(GetCellText(xlSheet, sourceRow, 19)) Then
                        installmentRows = installmentRows + 1
                    End If
                End If
            End If
        End If
    Next sourceRow

    If currentPolicyID > 0 Then
        ValidatePolicyBalance db, sessionID, currentPolicyID, lastRow, errorRows
    End If

    rsPayment.Close
    Set rsPayment = Nothing
    rsPolicy.Close
    Set rsPolicy = Nothing
    ws.CommitTrans
    transactionStarted = False

    UpdateSession db, sessionID, "Completed", sourceRows, policyRows, _
                  installmentRows, errorRows, "Payment rows: " & paymentRows
    CloseExcel xlBook, xlApp
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set fd = Nothing
    Set ws = Nothing
    Set db = Nothing

    MsgBox "Import completed." & vbCrLf & _
           "Policies: " & policyRows & vbCrLf & _
           "Payments: " & paymentRows & vbCrLf & _
           "Review items: " & errorRows, vbInformation, "BimYar Import"
    Exit Sub

ErrHandler:
    errNumber = Err.Number
    errDescription = Err.Description

    If transactionStarted Then ws.Rollback
    transactionStarted = False
    CloseRecordset rsPayment
    CloseRecordset rsPolicy
    CloseExcel xlBook, xlApp

    If Not db Is Nothing And sessionID > 0 Then
        LogFatalImportError db, sessionID, sourceRow, errNumber, errDescription
        UpdateSession db, sessionID, "Failed", sourceRows, policyRows, _
                      installmentRows, errorRows + 1, _
                      "Fatal error " & CStr(errNumber) & " at source row " & CStr(sourceRow)
    End If

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set fd = Nothing
    Set ws = Nothing
    Set db = Nothing
    MsgBox "Import failed at source row " & sourceRow & "." & vbCrLf & _
           CStr(errNumber) & " - " & errDescription, vbCritical, "BimYar Import"
End Sub

Private Function CreateImportSession(ByVal db As DAO.Database, ByVal reportYear As Integer, _
    ByVal reportMonth As Integer, ByVal reportType As Integer, ByVal marketerID As Variant, _
    ByVal fileName As String, ByVal filePath As String) As Long

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblImportSessions", dbOpenDynaset)
    rs.AddNew
    rs!reportYear = reportYear
    rs!reportMonth = reportMonth
    rs!reportType = reportType
    rs!marketerID = marketerID
    SetTextValidated rs, "ImportFileName", fileName, 0
    SetTextValidated rs, "ImportFilePath", filePath, 0
    rs!ImportDate = Now()
    rs!SessionStatus = "Importing"
    rs.Update
    rs.Bookmark = rs.LastModified
    CreateImportSession = rs!ImportSessionID
    rs.Close
    Set rs = Nothing
End Function

Private Function InsertPolicy(ByRef rs As DAO.Recordset, ByVal sessionID As Long, _
    ByVal reportType As Integer, ByVal marketerID As Variant, ByVal fileName As String, _
    ByVal sheet As Object, ByVal sourceRow As Long) As Long

    rs.AddNew
    rs!ImportSessionID = sessionID
    rs!reportType = reportType
    rs!marketerID = marketerID
    SetTextValidated rs, "ImportFileName", fileName, sourceRow
    rs!ImportDate = Now()
    SetLongField rs, "RowNo", sheet.Cells(sourceRow, 1).Value2
    SetTextValidated rs, "PolicyNumber", GetCellText(sheet, sourceRow, 2), sourceRow
    SetTextValidated rs, "InsuredName", GetCellText(sheet, sourceRow, 3), sourceRow
    SetTextValidated rs, "PreviousInsurer", GetCellText(sheet, sourceRow, 4), sourceRow
    SetTextValidated rs, "PreviousPolicyNumber", GetCellText(sheet, sourceRow, 5), sourceRow
    SetTextValidated rs, "VehicleType", GetCellText(sheet, sourceRow, 6), sourceRow
    SetTextValidated rs, "SystemType", GetCellText(sheet, sourceRow, 7), sourceRow
    SetTextValidated rs, "VehicleModel", GetCellText(sheet, sourceRow, 8), sourceRow
    SetTextValidated rs, "InsuredType", GetCellText(sheet, sourceRow, 9), sourceRow
    SetTextValidated rs, "NationalCode", GetCellText(sheet, sourceRow, 10), sourceRow
    SetTextValidated rs, "Phone", GetCellText(sheet, sourceRow, 11), sourceRow
    SetTextValidated rs, "Mobile", GetCellText(sheet, sourceRow, 12), sourceRow
    SetTextValidated rs, "Address", GetCellText(sheet, sourceRow, 13), sourceRow
    SetTextValidated rs, "StartDate", GetCellText(sheet, sourceRow, 14), sourceRow
    SetTextValidated rs, "EndDate", GetCellText(sheet, sourceRow, 15), sourceRow
    SetTextValidated rs, "IssueDate", GetCellText(sheet, sourceRow, 16), sourceRow
    SetCurrencyField rs, "PremiumAmount", sheet.Cells(sourceRow, 17).Value2
    SetTextValidated rs, "Profile", GetCellText(sheet, sourceRow, 18), sourceRow
    rs!IsImported = True
    rs!BalanceStatus = "Pending"
    rs.Update
    rs.Bookmark = rs.LastModified
    InsertPolicy = rs!ImportPolicyID
End Function

Private Sub InsertPayment(ByRef rs As DAO.Recordset, ByVal sessionID As Long, _
    ByVal policyID As Long, ByVal sequence As Long, ByVal sheet As Object, _
    ByVal sourceRow As Long, ByRef errorRows As Long, ByVal db As DAO.Database)

    Dim paymentType As String
    paymentType = GetCellText(sheet, sourceRow, 19)
    rs.AddNew
    rs!ImportPolicyID = policyID
    rs!ImportSessionID = sessionID
    rs!SourceRowNo = sourceRow
    rs!PaymentSequence = sequence
    SetTextValidated rs, "PaymentType", paymentType, sourceRow
    SetCurrencyField rs, "PaymentAmount", sheet.Cells(sourceRow, 20).Value2
    SetTextValidated rs, "PaymentRefNo", GetCellText(sheet, sourceRow, 21), sourceRow
    SetTextValidated rs, "PaymentDate", GetCellText(sheet, sourceRow, 22), sourceRow
    rs!IsInstallment = IsInstallmentType(paymentType)
    rs!CreatedDate = Now()
    rs.Update

    If Len(paymentType) > 0 And Not IsCashType(paymentType) And _
       Not IsInstallmentType(paymentType) Then
        LogImportError db, sessionID, policyID, sourceRow, "PaymentType", _
                       "UNKNOWN_PAYMENT_TYPE", "Unrecognized payment type; row was retained."
        errorRows = errorRows + 1
    End If
End Sub

Private Sub ValidatePolicyBalance(ByVal db As DAO.Database, ByVal sessionID As Long, _
    ByVal policyID As Long, ByVal sourceRow As Long, ByRef errorRows As Long)

    Dim rs As DAO.Recordset
    Dim premium As Currency
    Dim paid As Currency
    Set rs = db.OpenRecordset( _
        "SELECT P.PremiumAmount, Sum(Nz(X.PaymentAmount,0)) AS PaidAmount " & _
        "FROM tblImportPolicies AS P LEFT JOIN tblImportPayments AS X " & _
        "ON P.ImportPolicyID=X.ImportPolicyID WHERE P.ImportPolicyID=" & policyID & _
        " GROUP BY P.PremiumAmount", dbOpenSnapshot)
    If Not rs.EOF Then
        premium = Nz(rs!PremiumAmount, 0)
        paid = Nz(rs!PaidAmount, 0)
    End If
    rs.Close
    Set rs = Nothing

    If premium = paid Then
        db.Execute "UPDATE tblImportPolicies SET BalanceStatus='Balanced' " & _
                   "WHERE ImportPolicyID=" & policyID, dbFailOnError
    Else
        db.Execute "UPDATE tblImportPolicies SET BalanceStatus='NeedsReview' " & _
                   "WHERE ImportPolicyID=" & policyID, dbFailOnError
        LogImportError db, sessionID, policyID, sourceRow, "FinancialBalance", _
                       "PAYMENT_MISMATCH", "Policy total and payment total differ."
        errorRows = errorRows + 1
    End If
End Sub

Private Sub SetTextValidated(ByRef rs As DAO.Recordset, ByVal fieldName As String, _
    ByVal fieldValue As Variant, ByVal sourceRow As Long)

    Dim valueText As String
    Dim allowedLength As Long
    If IsNull(fieldValue) Or IsEmpty(fieldValue) Then
        rs.Fields(fieldName).Value = Null
        Exit Sub
    End If
    valueText = Trim$(CStr(fieldValue))
    If Len(valueText) = 0 Then
        rs.Fields(fieldName).Value = Null
        Exit Sub
    End If

    If rs.Fields(fieldName).Type = dbText Then
        allowedLength = rs.Fields(fieldName).Size
        If Len(valueText) > allowedLength Then
            mErrorField = fieldName
            mErrorRow = sourceRow
            mErrorActualLength = Len(valueText)
            mErrorAllowedLength = allowedLength
            Err.Raise ERR_TEXT_TOO_LONG, "SetTextValidated", _
                      "Text exceeds field size: " & fieldName & "; actual=" & _
                      Len(valueText) & "; allowed=" & allowedLength
        End If
    End If
    rs.Fields(fieldName).Value = valueText
End Sub

Private Sub SetLongField(ByRef rs As DAO.Recordset, ByVal fieldName As String, _
    ByVal fieldValue As Variant)
    If IsNumeric(fieldValue) Then
        rs.Fields(fieldName).Value = CLng(fieldValue)
    Else
        rs.Fields(fieldName).Value = Null
    End If
End Sub

Private Sub SetCurrencyField(ByRef rs As DAO.Recordset, ByVal fieldName As String, _
    ByVal fieldValue As Variant)
    If IsNumeric(fieldValue) Then
        rs.Fields(fieldName).Value = CCur(fieldValue)
    Else
        rs.Fields(fieldName).Value = Null
    End If
End Sub

Private Function HasPaymentData(ByVal sheet As Object, ByVal sourceRow As Long) As Boolean
    Dim columnNumber As Long
    For columnNumber = 19 To 22
        If Len(GetCellText(sheet, sourceRow, columnNumber)) > 0 Then
            HasPaymentData = True
            Exit Function
        End If
    Next columnNumber
End Function

Private Function IsEmptyReportRow(ByVal sheet As Object, ByVal sourceRow As Long) As Boolean
    Dim columnNumber As Long
    IsEmptyReportRow = True
    For columnNumber = 1 To REPORT_COLUMN_COUNT
        If Len(GetCellText(sheet, sourceRow, columnNumber)) > 0 Then
            IsEmptyReportRow = False
            Exit Function
        End If
    Next columnNumber
End Function

Private Function IsHeaderRow(ByVal sheet As Object, ByVal sourceRow As Long) As Boolean
    Dim rowNoText As String
    Dim policyText As String
    rowNoText = GetCellText(sheet, sourceRow, 1)
    policyText = GetCellText(sheet, sourceRow, 2)
    IsHeaderRow = (Len(rowNoText) > 0 And Not IsNumeric(rowNoText) And _
                   Len(policyText) > 0 And _
                   Not IsNumeric(GetCellText(sheet, sourceRow, 17)) And _
                   Not IsNumeric(GetCellText(sheet, sourceRow, 20)))
End Function

Private Function IsTotalRow(ByVal sheet As Object, ByVal sourceRow As Long) As Boolean
    Dim columnNumber As Long
    For columnNumber = 1 To REPORT_COLUMN_COUNT
        If GetCellText(sheet, sourceRow, columnNumber) = TotalWord() Then
            IsTotalRow = True
            Exit Function
        End If
    Next columnNumber
End Function

Private Function GetCellText(ByVal sheet As Object, ByVal sourceRow As Long, _
    ByVal columnNumber As Long) As String
    Dim value As Variant
    value = sheet.Cells(sourceRow, columnNumber).Value2
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        GetCellText = vbNullString
    Else
        GetCellText = Trim$(CStr(value))
    End If
End Function

Private Function IsCashType(ByVal value As String) As Boolean
    IsCashType = (Trim$(value) = ChrW(1606) & ChrW(1602) & ChrW(1583))
End Function

Private Function IsInstallmentType(ByVal value As String) As Boolean
    IsInstallmentType = (Trim$(value) = ChrW(1602) & ChrW(1587) & ChrW(1591))
End Function

Private Function TotalWord() As String
    TotalWord = ChrW(1605) & ChrW(1580) & ChrW(1605) & ChrW(1608) & ChrW(1593)
End Function

Private Sub LogImportError(ByVal db As DAO.Database, ByVal sessionID As Long, _
    ByVal policyID As Long, ByVal sourceRow As Long, ByVal errorType As String, _
    ByVal errorCode As String, ByVal safeMessage As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblImportErrors", dbOpenDynaset)
    rs.AddNew
    rs!ImportSessionID = sessionID
    If policyID > 0 Then
        rs!SourceImportPolicyID = policyID
    Else
        rs!SourceImportPolicyID = Null
    End If
    rs!RowNo = sourceRow
    rs!errorType = errorType
    rs!ErrorCode = errorCode
    rs!ErrorMessage = safeMessage
    rs!IsResolved = False
    rs.Update
    rs.Close
    Set rs = Nothing
End Sub

Private Sub LogFatalImportError(ByVal db As DAO.Database, ByVal sessionID As Long, _
    ByVal sourceRow As Long, ByVal errorNumber As Long, ByVal errorDescription As String)
    Dim safeMessage As String
    If errorNumber = ERR_TEXT_TOO_LONG Then
        safeMessage = "Field=" & mErrorField & "; source row=" & mErrorRow & _
                      "; actual length=" & mErrorActualLength & _
                      "; allowed length=" & mErrorAllowedLength
    Else
        safeMessage = "Fatal import error at source row " & sourceRow & _
                      "; error=" & errorNumber & "; " & Left$(errorDescription, 500)
    End If
    LogImportError db, sessionID, 0, sourceRow, "FatalImport", _
                   CStr(errorNumber), safeMessage
End Sub

Private Sub UpdateSession(ByVal db As DAO.Database, ByVal sessionID As Long, _
    ByVal status As String, ByVal totalRows As Long, ByVal policyRows As Long, _
    ByVal installmentRows As Long, ByVal errorRows As Long, ByVal notes As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM tblImportSessions WHERE ImportSessionID=" & _
                              sessionID, dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        rs!TotalRows = totalRows
        rs!PolicyRows = policyRows
        rs!InstallmentRows = installmentRows
        rs!ErrorRows = errorRows
        rs!SessionStatus = status
        rs!notes = notes
        rs.Update
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub CloseRecordset(ByRef rs As DAO.Recordset)
    If rs Is Nothing Then Exit Sub
    On Error Resume Next
    rs.Close
    On Error GoTo 0
    Set rs = Nothing
End Sub

Private Sub CloseExcel(ByRef book As Object, ByRef app As Object)
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not app Is Nothing Then app.Quit
    On Error GoTo 0
End Sub

Attribute VB_Name = "modValidation"
Option Explicit

'===========================================================
' VALIDATE HEALTHCARE MASTER DATA
'===========================================================

Public Sub ValidateHealthcareData()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim qcRange As Range

    Dim TotalRecords As Long
    Dim PassCount As Long
    Dim ReviewCount As Long
    Dim ErrorCount As Long
    Dim BlankCount As Long

    Dim PassRate As Double

    Dim QualityText As String
    Dim ValidationText As String
    Dim MessageText As String
    Dim LogDetails As String

    Dim StartTime As Double

    Dim SavedErrNumber As Long
    Dim SavedErrDescription As String

    On Error GoTo ErrorHandler

    StartTime = Timer

    '=======================================================
    ' APPLICATION SETUP
    '=======================================================

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Validating healthcare claims data..."

    '=======================================================
    ' LOCATE MASTER DATA TABLE
    '=======================================================

    Set ws = ThisWorkbook.Worksheets("Master Data")
    Set lo = ws.ListObjects("tbl_ClaimsMaster")

    If lo.DataBodyRange Is Nothing Then

        Err.Raise _
            vbObjectError + 2000, _
            "ValidateHealthcareData", _
            "tbl_ClaimsMaster contains no records."

    End If

    '=======================================================
    ' LOCATE OVERALL QC COLUMN
    '=======================================================

    On Error Resume Next

    Set qcRange = _
        lo.ListColumns("Overall QC").DataBodyRange

    On Error GoTo ErrorHandler

    If qcRange Is Nothing Then

        Err.Raise _
            vbObjectError + 2001, _
            "ValidateHealthcareData", _
            "The 'Overall QC' column could not be found."

    End If

    '=======================================================
    ' COUNT RECORDS
    '=======================================================

    TotalRecords = lo.ListRows.Count

    PassCount = _
        Application.CountIf(qcRange, "PASS")

    ReviewCount = _
        Application.CountIf(qcRange, "REVIEW")

    ErrorCount = _
        Application.CountIf(qcRange, "ERROR")

    BlankCount = _
        Application.CountBlank(qcRange)

    '=======================================================
    ' CALCULATE PASS RATE
    '=======================================================

    If TotalRecords > 0 Then

        PassRate = PassCount / TotalRecords

    Else

        PassRate = 0

    End If

    '=======================================================
    ' BUILD DATA QUALITY STATUS
    '=======================================================

    QualityText = Format(PassRate, "0.0%") & " PASS"

    QualityText = QualityText & _
        " | " & Format(ReviewCount, "#,##0") & " REVIEW"

    QualityText = QualityText & _
        " | " & Format(ErrorCount, "#,##0") & " ERROR"

    QualityText = QualityText & _
        " | " & Format(BlankCount, "#,##0") & " UNCLASSIFIED"

    SetNamedCellValue _
        "DataQualityStatus", _
        QualityText

    '=======================================================
    ' DETERMINE OVERALL VALIDATION STATUS
    '=======================================================

    If ErrorCount > 0 Then

        ValidationText = "ATTENTION REQUIRED"

    ElseIf BlankCount > 0 Then

        ValidationText = "UNCLASSIFIED RECORDS"

    ElseIf ReviewCount > 0 Then

        ValidationText = "REVIEW REQUIRED"

    Else

        ValidationText = "VALIDATION PASSED"

    End If

    SetNamedCellValue _
        "ValidationStatus", _
        ValidationText

    '=======================================================
    ' FORMAT VALIDATION STATUS
    '=======================================================

    If NameExists("ValidationStatus") Then

        With ThisWorkbook.Names("ValidationStatus").RefersToRange

            .Font.Bold = True

            Select Case ValidationText

                Case "VALIDATION PASSED"

                    .Font.Color = RGB(0, 128, 0)

                Case "REVIEW REQUIRED"

                    .Font.Color = RGB(192, 128, 0)

                Case "UNCLASSIFIED RECORDS"

                    .Font.Color = RGB(192, 128, 0)

                Case "ATTENTION REQUIRED"

                    .Font.Color = RGB(192, 0, 0)

                Case Else

                    .Font.Color = RGB(0, 0, 0)

            End Select

        End With

    End If

    '=======================================================
    ' WRITE REFRESH LOG ENTRY
    '=======================================================

    LogDetails = _
        "Healthcare claims validation completed."

    LogActivity _
        ActionName:="Data Validation", _
        ResultText:=ValidationText, _
        DurationSeconds:=GetElapsedSeconds(StartTime), _
        DetailsText:=LogDetails

    '=======================================================
    ' RESTORE EXCEL
    '=======================================================

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    '=======================================================
    ' BUILD COMPLETION MESSAGE
    '=======================================================

    MessageText = "Validation complete."

    MessageText = MessageText & vbCrLf & vbCrLf
    MessageText = MessageText & _
        "Total Records: " & Format(TotalRecords, "#,##0")

    MessageText = MessageText & vbCrLf
    MessageText = MessageText & _
        "PASS: " & Format(PassCount, "#,##0")

    MessageText = MessageText & vbCrLf
    MessageText = MessageText & _
        "REVIEW: " & Format(ReviewCount, "#,##0")

    MessageText = MessageText & vbCrLf
    MessageText = MessageText & _
        "ERROR: " & Format(ErrorCount, "#,##0")

    MessageText = MessageText & vbCrLf
    MessageText = MessageText & _
        "Unclassified: " & Format(BlankCount, "#,##0")

    MessageText = MessageText & vbCrLf & vbCrLf
    MessageText = MessageText & _
        "Status: " & ValidationText

    MsgBox _
        MessageText, _
        vbInformation, _
        "Healthcare Data Validation"

    Exit Sub


'===========================================================
' ERROR HANDLER
'===========================================================

ErrorHandler:

    SavedErrNumber = Err.Number
    SavedErrDescription = Err.Description

    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    SetNamedCellValue _
        "ValidationStatus", _
        "VALIDATION FAILED"

    LogDetails = _
        "Error " & SavedErrNumber & ": " & SavedErrDescription

    On Error Resume Next

    LogActivity _
        ActionName:="Data Validation", _
        ResultText:="FAILED", _
        DurationSeconds:=GetElapsedSeconds(StartTime), _
        DetailsText:=LogDetails

    On Error GoTo 0

    MessageText = "Validation failed."
    MessageText = MessageText & vbCrLf & vbCrLf
    MessageText = MessageText & "Error " & SavedErrNumber
    MessageText = MessageText & vbCrLf
    MessageText = MessageText & SavedErrDescription

    MsgBox _
        MessageText, _
        vbCritical, _
        "Validation Error"

End Sub


Attribute VB_Name = "modLogging"
Option Explicit


'===========================================================
' REFRESH LOG CONFIGURATION
'===========================================================

Private Const LOG_SHEET As String = "Refresh Log"
Private Const LOG_TABLE As String = "tbl_RefreshLog"
Private Const EXPECTED_COLUMNS As Long = 12


'===========================================================
' MAIN PUBLIC LOGGING PROCEDURE
'
' Column order expected:
'
' 1  Timestamp
' 2  Action
' 3  Records
' 4  PASS
' 5  REVIEW
' 6  ERROR
' 7  Unclassified
' 8  Reports Generated
' 9  Duration (sec)
' 10 Result
' 11 Details
' 12 User
'
'===========================================================

Public Sub LogActivity( _
    ByVal ActionName As String, _
    ByVal ResultText As String, _
    Optional ByVal ReportsGenerated As Variant, _
    Optional ByVal DurationSeconds As Variant, _
    Optional ByVal DetailsText As String = "")

    Dim HasReports As Boolean
    Dim HasDuration As Boolean

    HasReports = Not IsMissing(ReportsGenerated)
    HasDuration = Not IsMissing(DurationSeconds)

    WriteRefreshLog _
        ActionName:=ActionName, _
        ResultText:=ResultText, _
        HasReports:=HasReports, _
        ReportsGenerated:=ReportsGenerated, _
        HasDuration:=HasDuration, _
        DurationSeconds:=DurationSeconds, _
        DetailsText:=DetailsText, _
        ShowErrors:=False

End Sub


'===========================================================
' CORE LOGGING PROCEDURE
'===========================================================

Private Sub WriteRefreshLog( _
    ByVal ActionName As String, _
    ByVal ResultText As String, _
    ByVal HasReports As Boolean, _
    ByVal ReportsGenerated As Variant, _
    ByVal HasDuration As Boolean, _
    ByVal DurationSeconds As Variant, _
    ByVal DetailsText As String, _
    ByVal ShowErrors As Boolean)

    Dim wsLog As Worksheet
    Dim loLog As ListObject
    Dim NewRow As ListRow

    Dim wsMaster As Worksheet
    Dim loMaster As ListObject
    Dim qcRange As Range

    Dim TotalRecords As Long
    Dim PassCount As Long
    Dim ReviewCount As Long
    Dim ErrorCount As Long
    Dim BlankCount As Long

    Dim WindowsUser As String

    Dim SavedErrNumber As Long
    Dim SavedErrDescription As String

    On Error GoTo ErrorHandler


    '=======================================================
    ' LOCATE REFRESH LOG WORKSHEET
    '=======================================================

    Set wsLog = _
        ThisWorkbook.Worksheets(LOG_SHEET)


    '=======================================================
    ' LOCATE REFRESH LOG TABLE
    '=======================================================

    Set loLog = _
        wsLog.ListObjects(LOG_TABLE)


    '=======================================================
    ' VALIDATE TABLE STRUCTURE
    '=======================================================

    If loLog.ListColumns.Count <> EXPECTED_COLUMNS Then

        Err.Raise _
            vbObjectError + 5000, _
            "WriteRefreshLog", _
            "tbl_RefreshLog must contain exactly 12 columns." & _
            vbCrLf & _
            "Current column count: " & _
            loLog.ListColumns.Count

    End If


    '=======================================================
    ' READ MASTER DATA COUNTS
    '=======================================================

    On Error Resume Next

    Set wsMaster = _
        ThisWorkbook.Worksheets("Master Data")

    Set loMaster = _
        wsMaster.ListObjects("tbl_ClaimsMaster")

    On Error GoTo ErrorHandler


    If Not loMaster Is Nothing Then

        If Not loMaster.DataBodyRange Is Nothing Then

            TotalRecords = _
                loMaster.ListRows.Count

        End If


        '---------------------------------------------------
        ' Overall QC
        '---------------------------------------------------

        On Error Resume Next

        Set qcRange = _
            loMaster.ListColumns( _
                "Overall QC" _
            ).DataBodyRange

        On Error GoTo ErrorHandler


        If Not qcRange Is Nothing Then

            PassCount = _
                Application.CountIf( _
                    qcRange, _
                    "PASS" _
                )

            ReviewCount = _
                Application.CountIf( _
                    qcRange, _
                    "REVIEW" _
                )

            ErrorCount = _
                Application.CountIf( _
                    qcRange, _
                    "ERROR" _
                )

            BlankCount = _
                Application.CountBlank( _
                    qcRange _
                )

        End If

    End If


    '=======================================================
    ' GET WINDOWS USER
    '=======================================================

    WindowsUser = _
        Environ$("Username")

    If Len(WindowsUser) = 0 Then

        WindowsUser = _
            Application.UserName

    End If


    '=======================================================
    ' DETERMINE WHICH TABLE ROW TO USE
    '
    ' Claude/Excel may have created one blank starter row.
    ' If so, use that row instead of creating another blank.
    '=======================================================

    If loLog.ListRows.Count = 0 Then

        Set NewRow = _
            loLog.ListRows.Add

    ElseIf _
        loLog.ListRows.Count = 1 And _
        Application.CountA( _
            loLog.ListRows(1).Range _
        ) = 0 Then

        Set NewRow = _
            loLog.ListRows(1)

    Else

        Set NewRow = _
            loLog.ListRows.Add

    End If


    '=======================================================
    ' POPULATE AUDIT ROW
    '
    ' Writing by POSITION intentionally avoids problems
    ' caused by spaces or slight header-name differences.
    '=======================================================

    With NewRow.Range

        '1 - Timestamp
        .Cells(1, 1).Value = Now

        '2 - Action
        .Cells(1, 2).Value = _
            ActionName

        '3 - Records
        .Cells(1, 3).Value = _
            TotalRecords

        '4 - PASS
        .Cells(1, 4).Value = _
            PassCount

        '5 - REVIEW
        .Cells(1, 5).Value = _
            ReviewCount

        '6 - ERROR
        .Cells(1, 6).Value = _
            ErrorCount

        '7 - Unclassified
        .Cells(1, 7).Value = _
            BlankCount


        '8 - Reports Generated
        If HasReports Then

            .Cells(1, 8).Value = _
                ReportsGenerated

        Else

            .Cells(1, 8).ClearContents

        End If


        '9 - Duration
        If HasDuration Then

            If IsNumeric(DurationSeconds) Then

                .Cells(1, 9).Value = _
                    Round( _
                        CDbl(DurationSeconds), _
                        1 _
                    )

            Else

                .Cells(1, 9).ClearContents

            End If

        Else

            .Cells(1, 9).ClearContents

        End If


        '10 - Result
        .Cells(1, 10).Value = _
            ResultText

        '11 - Details
        .Cells(1, 11).Value = _
            DetailsText

        '12 - User
        .Cells(1, 12).Value = _
            WindowsUser

    End With


    '=======================================================
    ' FORMAT NEW ROW
    '=======================================================

    NewRow.Range.Cells(1, 1).NumberFormat = _
        "mmm dd, yyyy hh:mm:ss AM/PM"

    NewRow.Range.Cells(1, 3).NumberFormat = _
        "#,##0"

    NewRow.Range.Cells(1, 4).NumberFormat = _
        "#,##0"

    NewRow.Range.Cells(1, 5).NumberFormat = _
        "#,##0"

    NewRow.Range.Cells(1, 6).NumberFormat = _
        "#,##0"

    NewRow.Range.Cells(1, 7).NumberFormat = _
        "#,##0"

    NewRow.Range.Cells(1, 8).NumberFormat = _
        "#,##0"

    NewRow.Range.Cells(1, 9).NumberFormat = _
        "0.0"


    Exit Sub


'===========================================================
' ERROR HANDLER
'===========================================================

ErrorHandler:

    SavedErrNumber = _
        Err.Number

    SavedErrDescription = _
        Err.Description


    Debug.Print _
        "Refresh Log Error " & _
        SavedErrNumber & _
        ": " & _
        SavedErrDescription


    If ShowErrors Then

        MsgBox _
            "Refresh Log test failed." & _
            vbCrLf & vbCrLf & _
            "Error " & _
            SavedErrNumber & _
            vbCrLf & _
            SavedErrDescription, _
            vbCritical, _
            "Refresh Log Error"

    End If

End Sub


'===========================================================
' TEST REFRESH LOG
'
' Run this manually before testing Validation/Refresh.
'===========================================================

Public Sub TestRefreshLog()

    WriteRefreshLog _
        ActionName:="SYSTEM TEST", _
        ResultText:="SUCCESS", _
        HasReports:=False, _
        ReportsGenerated:=Empty, _
        HasDuration:=True, _
        DurationSeconds:=1.2, _
        DetailsText:= _
            "Refresh Log connection test.", _
        ShowErrors:=True


    MsgBox _
        "Refresh Log test completed." & _
        vbCrLf & vbCrLf & _
        "Check the Refresh Log worksheet for a SYSTEM TEST row.", _
        vbInformation, _
        "Refresh Log Test"

End Sub


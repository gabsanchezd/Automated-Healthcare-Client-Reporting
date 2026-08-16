Attribute VB_Name = "modRefresh"
Option Explicit


'===========================================================
' GLOBAL EVENT HANDLER
'
' Keeps CQueryEvents alive while Power Query is refreshing.
'===========================================================

Public gQueryEvents As CQueryEvents


'===========================================================
' REFRESH START TIME
'===========================================================

Public gRefreshStartedAt As Date


'===========================================================
' REFRESH HEALTHCARE MASTER DATA
'===========================================================

Public Sub RefreshHealthcareData()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim qt As QueryTable

    Dim SavedErrNumber As Long
    Dim SavedErrDescription As String

    On Error GoTo ErrorHandler


    '=======================================================
    ' LOCATE POWER QUERY OUTPUT TABLE
    '=======================================================

    Set ws = _
        ThisWorkbook.Worksheets( _
            "Master Data" _
        )

    Set lo = _
        ws.ListObjects( _
            "tbl_ClaimsMaster" _
        )

    Set qt = _
        lo.QueryTable


    '=======================================================
    ' PREVENT DUPLICATE REFRESH
    '=======================================================

    If qt.Refreshing Then

        MsgBox _
            "Healthcare data is already refreshing.", _
            vbInformation, _
            "Refresh In Progress"

        Exit Sub

    End If


    '=======================================================
    ' INITIALIZE QUERY EVENT HANDLER
    '=======================================================

    Set gQueryEvents = _
        New CQueryEvents

    gQueryEvents.Initialize qt


    '=======================================================
    ' RECORD REFRESH START
    '=======================================================

    gRefreshStartedAt = Now


    '=======================================================
    ' UPDATE CONTROL PANEL
    '=======================================================

    SetNamedCellValue _
        "PowerQueryStatus", _
        "REFRESHING..."


    Application.StatusBar = _
        "Refreshing healthcare claims data..."


    'Events must remain enabled so AfterRefresh can fire.
    Application.EnableEvents = True


    '=======================================================
    ' START BACKGROUND REFRESH
    '=======================================================

    qt.Refresh _
        BackgroundQuery:=True


    Exit Sub


'===========================================================
' ERROR HANDLER
'===========================================================

ErrorHandler:

    SavedErrNumber = _
        Err.Number

    SavedErrDescription = _
        Err.Description


    Application.StatusBar = False


    SetNamedCellValue _
        "PowerQueryStatus", _
        "REFRESH ERROR"


    LogActivity _
        ActionName:="Data Refresh", _
        ResultText:="FAILED", _
        DurationSeconds:=0, _
        DetailsText:= _
            "Unable to start refresh. Error " & _
            SavedErrNumber & _
            ": " & _
            SavedErrDescription


    Set gQueryEvents = Nothing


    MsgBox _
        "Unable to start the data refresh." & _
        vbCrLf & vbCrLf & _
        "Error " & _
        SavedErrNumber & _
        vbCrLf & _
        SavedErrDescription, _
        vbCritical, _
        "Refresh Error"

End Sub


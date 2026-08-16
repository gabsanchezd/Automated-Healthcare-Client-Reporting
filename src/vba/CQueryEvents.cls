VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQueryEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'===========================================================
' POWER QUERY EVENT OBJECT
'===========================================================

Public WithEvents qt As QueryTable
Attribute qt.VB_VarHelpID = -1


'===========================================================
' INITIALIZE QUERY EVENT LISTENER
'===========================================================

Public Sub Initialize( _
    ByVal Query As QueryTable)

    Set qt = Query

End Sub


'===========================================================
' RUN AFTER POWER QUERY REFRESH COMPLETES
'===========================================================

Private Sub qt_AfterRefresh( _
    ByVal Success As Boolean)

    Dim DurationSeconds As Long

    On Error GoTo EventError


    '=======================================================
    ' CALCULATE REFRESH DURATION
    '=======================================================

    If gRefreshStartedAt > 0 Then

        DurationSeconds = _
            DateDiff( _
                "s", _
                gRefreshStartedAt, _
                Now _
            )

    Else

        DurationSeconds = 0

    End If


    '=======================================================
    ' SUCCESSFUL REFRESH
    '=======================================================

    If Success Then

        SetNamedCellValue _
            "LastRefreshTime", _
            Now


        SetNamedCellValue _
            "PowerQueryStatus", _
            "REFRESH COMPLETE"


        'Recalculate dashboard formulas
        Application.Calculate


        'Write audit log entry
        LogActivity _
            ActionName:="Data Refresh", _
            ResultText:="SUCCESS", _
            DurationSeconds:=DurationSeconds, _
            DetailsText:= _
                "qry_Claims_Master refreshed successfully."


        Application.StatusBar = _
            "Healthcare claims refresh complete."


    '=======================================================
    ' FAILED REFRESH
    '=======================================================

    Else

        SetNamedCellValue _
            "PowerQueryStatus", _
            "REFRESH FAILED"


        LogActivity _
            ActionName:="Data Refresh", _
            ResultText:="FAILED", _
            DurationSeconds:=DurationSeconds, _
            DetailsText:= _
                "Power Query refresh did not complete successfully."


        Application.StatusBar = _
            "Healthcare claims refresh failed."

    End If


    'Release event handler after refresh is finished.
    Set gQueryEvents = Nothing


    Exit Sub


'===========================================================
' EVENT ERROR HANDLER
'===========================================================

EventError:

    SetNamedCellValue _
        "PowerQueryStatus", _
        "REFRESH EVENT ERROR"


    Application.StatusBar = False


    Debug.Print _
        "Refresh Event Error " & _
        Err.Number & _
        ": " & _
        Err.Description

End Sub


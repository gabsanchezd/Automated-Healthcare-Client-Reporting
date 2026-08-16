Attribute VB_Name = "modUtilities"
Option Explicit


'===========================================================
' CHECK IF A WORKBOOK-LEVEL NAME EXISTS
'===========================================================

Public Function NameExists(ByVal NameText As String) As Boolean

    Dim nm As Name

    On Error Resume Next

    Set nm = ThisWorkbook.Names(NameText)

    NameExists = Not nm Is Nothing

    Set nm = Nothing

    On Error GoTo 0

End Function


'===========================================================
' SET A NAMED CELL VALUE SAFELY
'
' If the named range does not exist, the procedure simply
' skips it instead of stopping the automation.
'===========================================================

Public Sub SetNamedCellValue( _
    ByVal NameText As String, _
    ByVal NewValue As Variant)

    On Error Resume Next

    If NameExists(NameText) Then

        ThisWorkbook.Names(NameText) _
            .RefersToRange.Value = NewValue

    End If

    On Error GoTo 0

End Sub


'===========================================================
' RETURN PROJECT ROOT DIRECTORY
'
' Expected folder structure:
'
' Automated-Healthcare-Client-Reporting
'   |
'   +-- workbook
'   |     +-- Healthcare_Reporting_Automation.xlsm
'   |
'   +-- output
'
'===========================================================

Public Function GetProjectRoot() As String

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    GetProjectRoot = _
        fso.GetParentFolderName(ThisWorkbook.Path)

End Function


'===========================================================
' CREATE A COMPLETE FOLDER PATH IF IT DOES NOT EXIST
'===========================================================

Public Sub EnsureFolderPath(ByVal FolderPath As String)

    Dim fso As Object
    Dim ParentPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(FolderPath) Then Exit Sub

    ParentPath = _
        fso.GetParentFolderName(FolderPath)

    If Len(ParentPath) > 0 Then

        If Not fso.FolderExists(ParentPath) Then

            EnsureFolderPath ParentPath

        End If

    End If

    fso.CreateFolder FolderPath

End Sub


'===========================================================
' CLEAN TEXT SO IT CAN BE USED AS A WINDOWS FILE NAME
'===========================================================

Public Function CleanFileName( _
    ByVal TextValue As String) As String

    Dim InvalidCharacters As Variant
    Dim ch As Variant
    Dim Result As String

    InvalidCharacters = Array( _
        "\", _
        "/", _
        ":", _
        "*", _
        "?", _
        """", _
        "<", _
        ">", _
        "|" _
    )

    Result = Trim$(TextValue)

    For Each ch In InvalidCharacters

        Result = Replace( _
            Result, _
            CStr(ch), _
            "_" _
        )

    Next ch

    Result = Replace( _
        Result, _
        " ", _
        "_" _
    )

    Do While InStr(Result, "__") > 0

        Result = Replace( _
            Result, _
            "__", _
            "_" _
        )

    Loop

    CleanFileName = Result

End Function


'===========================================================
' WAIT FOR EXCEL CALCULATION TO COMPLETE
'===========================================================

Public Sub WaitForExcelCalculation()

    Do While _
        Application.CalculationState <> xlDone

        DoEvents

    Loop

End Sub


'===========================================================
' CALCULATE ELAPSED SECONDS
'
' Handles VBA Timer resetting at midnight.
'===========================================================

Public Function GetElapsedSeconds( _
    ByVal StartTimer As Double) As Double

    Dim CurrentTimer As Double

    CurrentTimer = Timer

    If CurrentTimer < StartTimer Then

        CurrentTimer = _
            CurrentTimer + 86400

    End If

    GetElapsedSeconds = _
        CurrentTimer - StartTimer

End Function


'===========================================================
' CONVERT REPORT FORMULAS TO STATIC VALUES
'===========================================================

Public Sub ConvertFormulasToValues( _
    ByVal ws As Worksheet)

    Dim rng As Range

    On Error GoTo SafeExit

    Set rng = ws.UsedRange

    If Not rng Is Nothing Then

        rng.Value = rng.Value

    End If

SafeExit:

    Set rng = Nothing

End Sub


'===========================================================
' BREAK EXTERNAL EXCEL LINKS
'===========================================================

Public Sub BreakWorkbookLinks( _
    ByVal wb As Workbook)

    Dim Links As Variant
    Dim i As Long

    On Error GoTo SafeExit

    Links = _
        wb.LinkSources( _
            Type:=xlLinkTypeExcelLinks _
        )

    If IsEmpty(Links) Then _
        GoTo SafeExit

    For i = _
        LBound(Links) To UBound(Links)

        wb.BreakLink _
            Name:=Links(i), _
            Type:=xlLinkTypeExcelLinks

    Next i

SafeExit:

End Sub


'===========================================================
' BASIC EXPORTED REPORT PAGE SETUP
'===========================================================

Public Sub PrepareClientReport( _
    ByVal wsSummary As Worksheet, _
    ByVal wsDetail As Worksheet)

    On Error Resume Next

    With wsSummary.PageSetup

        .Orientation = xlLandscape

        .Zoom = False

        .FitToPagesWide = 1

        .FitToPagesTall = False

    End With


    With wsDetail.PageSetup

        .Orientation = xlLandscape

        .Zoom = False

        .FitToPagesWide = 1

        .FitToPagesTall = False

    End With

    On Error GoTo 0

End Sub


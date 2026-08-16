Attribute VB_Name = "modReportGenerator"
Option Explicit


'===========================================================
' TEST MODE
'
' True  = Generate ONE report only
' False = Generate reports for ALL entities
'
' Keep True while testing.
'===========================================================

Private Const TEST_MODE As Boolean = False


'===========================================================
' GENERATE CLIENT REPORTS
'===========================================================

Public Sub GenerateClientReports()

    Dim wsMaster As Worksheet
    Dim loMaster As ListObject

    Dim entityRange As Range
    Dim cell As Range

    Dim dict As Object
    Dim EntityName As Variant

    Dim ProjectRoot As String
    Dim OutputRoot As String
    Dim OutputFolder As String

    Dim GenerationDate As String
    Dim ReportPath As String

    Dim wbReport As Workbook
    Dim wsSummary As Worksheet
    Dim wsDetail As Worksheet

    Dim GeneratedCount As Long

    Dim StartTime As Double

    Dim CurrentStage As String

    Dim SavedErrNumber As Long
    Dim SavedErrDescription As String

    Dim ActionLabel As String

    On Error GoTo ErrorHandler


    StartTime = Timer


    '=======================================================
    ' DETERMINE LOGGING ACTION NAME
    '=======================================================

    If TEST_MODE Then

        ActionLabel = _
            "Test Report Generation"

    Else

        ActionLabel = _
            "Client Report Generation"

    End If


    '=======================================================
    ' APPLICATION SETUP
    '=======================================================

    Application.ScreenUpdating = False

    Application.EnableEvents = False

    Application.DisplayAlerts = False

    Application.StatusBar = _
        "Preparing healthcare client reports..."


    SetNamedCellValue _
        "ReportGenerationStatus", _
        "PREPARING REPORTS"


    '=======================================================
    ' STEP 1 - LOCATE MASTER DATA
    '=======================================================

    CurrentStage = _
        "Locating Master Data table"


    Set wsMaster = _
        ThisWorkbook.Worksheets( _
            "Master Data" _
        )


    Set loMaster = _
        wsMaster.ListObjects( _
            "tbl_ClaimsMaster" _
        )


    If loMaster.DataBodyRange Is Nothing Then

        Err.Raise _
            vbObjectError + 3000, _
            "GenerateClientReports", _
            "tbl_ClaimsMaster contains no records."

    End If


    '=======================================================
    ' STEP 2 - VERIFY ENTITY SELECTOR
    '=======================================================

    CurrentStage = _
        "Checking selEntity named range"


    If Not NameExists("selEntity") Then

        Err.Raise _
            vbObjectError + 3001, _
            "GenerateClientReports", _
            "The workbook-level named range " & _
            "'selEntity' could not be found."

    End If


    '=======================================================
    ' STEP 3 - BUILD UNIQUE ENTITY LIST
    '=======================================================

    CurrentStage = _
        "Building healthcare entity list"


    Set dict = _
        CreateObject( _
            "Scripting.Dictionary" _
        )


    Set entityRange = _
        loMaster.ListColumns( _
            "Entity" _
        ).DataBodyRange


    For Each cell In entityRange.Cells

        If Len( _
            Trim$( _
                CStr(cell.Value) _
            ) _
        ) > 0 Then


            If Not dict.Exists( _
                Trim$( _
                    CStr(cell.Value) _
                ) _
            ) Then


                dict.Add _
                    Trim$( _
                        CStr(cell.Value) _
                    ), _
                    Trim$( _
                        CStr(cell.Value) _
                    )

            End If

        End If

    Next cell


    If dict.Count = 0 Then

        Err.Raise _
            vbObjectError + 3002, _
            "GenerateClientReports", _
            "No healthcare entities were found " & _
            "in tbl_ClaimsMaster."

    End If


    '=======================================================
    ' STEP 4 - CREATE OUTPUT FOLDER
    '=======================================================

    CurrentStage = _
        "Creating output folders"


    ProjectRoot = _
        GetProjectRoot()


    GenerationDate = _
        Format( _
            Date, _
            "yyyy-mm-dd" _
        )


    OutputRoot = _
        ProjectRoot & _
        "\output\client_reports"


    OutputFolder = _
        OutputRoot & _
        "\" & _
        GenerationDate


    EnsureFolderPath _
        OutputRoot


    EnsureFolderPath _
        OutputFolder


    '=======================================================
    ' STEP 5 - GENERATE REPORTS
    '=======================================================

    GeneratedCount = 0


    For Each EntityName In dict.Keys


        '---------------------------------------------------
        ' TEST MODE - STOP AFTER ONE REPORT
        '---------------------------------------------------

        If TEST_MODE Then

            If GeneratedCount >= 1 Then

                Exit For

            End If

        End If


        '---------------------------------------------------
        ' STATUS
        '---------------------------------------------------

        CurrentStage = _
            "Processing entity: " & _
            CStr(EntityName)


        Application.StatusBar = _
            "Generating report " & _
            (GeneratedCount + 1) & _
            " of " & _
            dict.Count & _
            ": " & _
            CStr(EntityName)


        SetNamedCellValue _
            "ReportGenerationStatus", _
            "GENERATING: " & _
            CStr(EntityName)


        DoEvents


        '===================================================
        ' STEP 5A - SET ENTITY
        '===================================================

        CurrentStage = _
            "Setting selEntity to " & _
            CStr(EntityName)


        ThisWorkbook.Names( _
            "selEntity" _
        ).RefersToRange.Value = _
            CStr(EntityName)


        '===================================================
        ' STEP 5B - UPDATE REPORT FORMULAS
        '===================================================

        CurrentStage = _
            "Updating Client Summary and Claim Detail"


        Application.Calculate


        WaitForExcelCalculation


        DoEvents


        '===================================================
        ' STEP 5C - COPY REPORT SHEETS
        '===================================================

        CurrentStage = _
            "Copying Client Summary and Claim Detail"


        ThisWorkbook.Worksheets( _
            Array( _
                "Client Summary", _
                "Claim Detail" _
            ) _
        ).Copy


        Set wbReport = _
            ActiveWorkbook


        Set wsSummary = _
            wbReport.Worksheets( _
                "Client Summary" _
            )


        Set wsDetail = _
            wbReport.Worksheets( _
                "Claim Detail" _
            )


        '===================================================
        ' STEP 5D - CONVERT TO VALUES
        '===================================================

        CurrentStage = _
            "Converting formulas to static values"


        ConvertFormulasToValues _
            wsSummary


        ConvertFormulasToValues _
            wsDetail


        '===================================================
        ' STEP 5E - BREAK EXTERNAL LINKS
        '===================================================

        CurrentStage = _
            "Removing external workbook links"


        BreakWorkbookLinks _
            wbReport


        '===================================================
        ' STEP 5F - FORMAT REPORT
        '===================================================

        CurrentStage = _
            "Formatting exported report"


        PrepareClientReport _
            wsSummary, _
            wsDetail


        '===================================================
        ' STEP 5G - BUILD FILE NAME
        '===================================================

        CurrentStage = _
            "Building report filename"


        ReportPath = _
            OutputFolder & _
            "\Client_Report_" & _
            CleanFileName( _
                CStr(EntityName) _
            ) & _
            "_" & _
            GenerationDate & _
            ".xlsx"


        '===================================================
        ' STEP 5H - SAVE REPORT
        '===================================================

        CurrentStage = _
            "Saving report: " & _
            ReportPath


        wbReport.SaveAs _
            Filename:=ReportPath, _
            FileFormat:=xlOpenXMLWorkbook


        wbReport.Close _
            SaveChanges:=False


        Set wbReport = Nothing

        Set wsSummary = Nothing

        Set wsDetail = Nothing


        GeneratedCount = _
            GeneratedCount + 1


    Next EntityName


    '=======================================================
    ' STEP 6 - RESET ENTITY SELECTOR
    '=======================================================

    CurrentStage = _
        "Resetting healthcare entity selector"


    ThisWorkbook.Names( _
        "selEntity" _
    ).RefersToRange.Value = _
        "(All Entities)"


    Application.Calculate


    WaitForExcelCalculation


    '=======================================================
    ' STEP 7 - UPDATE CONTROL PANEL
    '=======================================================

    CurrentStage = _
        "Updating Control Panel"


    SetNamedCellValue _
        "LastReportGeneration", _
        Now


    SetNamedCellValue _
        "ReportGenerationStatus", _
        "COMPLETE - " & _
        GeneratedCount & _
        " REPORT(S)"


    '=======================================================
    ' STEP 8 - WRITE REFRESH LOG
    '=======================================================

    LogActivity _
        ActionName:=ActionLabel, _
        ResultText:="SUCCESS", _
        ReportsGenerated:=GeneratedCount, _
        DurationSeconds:= _
            GetElapsedSeconds(StartTime), _
        DetailsText:= _
            "Reports saved to: " & _
            OutputFolder


    '=======================================================
    ' RESTORE APPLICATION
    '=======================================================

    Application.StatusBar = False

    Application.DisplayAlerts = True

    Application.EnableEvents = True

    Application.ScreenUpdating = True


    '=======================================================
    ' COMPLETION MESSAGE
    '=======================================================

    If TEST_MODE Then

        MsgBox _
            "Test report generated successfully." & _
            vbCrLf & vbCrLf & _
            "Reports generated: " & _
            GeneratedCount & _
            vbCrLf & vbCrLf & _
            "Folder:" & _
            vbCrLf & _
            OutputFolder & _
            vbCrLf & vbCrLf & _
            "A Refresh Log entry was also created.", _
            vbInformation, _
            "Test Report Complete"

    Else

        MsgBox _
            GeneratedCount & _
            " client reports generated successfully." & _
            vbCrLf & vbCrLf & _
            "Folder:" & _
            vbCrLf & _
            OutputFolder & _
            vbCrLf & vbCrLf & _
            "Elapsed time: " & _
            Format( _
                GetElapsedSeconds(StartTime), _
                "0.0" _
            ) & _
            " seconds.", _
            vbInformation, _
            "Report Generation Complete"

    End If


    Exit Sub


'===========================================================
' ERROR HANDLER
'===========================================================

ErrorHandler:

    'Capture original error BEFORE cleanup.
    SavedErrNumber = _
        Err.Number

    SavedErrDescription = _
        Err.Description


    '-------------------------------------------------------
    ' Restore Excel
    '-------------------------------------------------------

    Application.StatusBar = False

    Application.DisplayAlerts = True

    Application.EnableEvents = True

    Application.ScreenUpdating = True


    '-------------------------------------------------------
    ' Close partially generated workbook
    '-------------------------------------------------------

    On Error Resume Next

    If Not wbReport Is Nothing Then

        wbReport.Close _
            SaveChanges:=False

    End If


    '-------------------------------------------------------
    ' Reset entity selector
    '-------------------------------------------------------

    If NameExists("selEntity") Then

        ThisWorkbook.Names( _
            "selEntity" _
        ).RefersToRange.Value = _
            "(All Entities)"

    End If


    On Error GoTo 0


    '-------------------------------------------------------
    ' Control Panel status
    '-------------------------------------------------------

    SetNamedCellValue _
        "ReportGenerationStatus", _
        "FAILED"


    '-------------------------------------------------------
    ' Write failure to audit log
    '-------------------------------------------------------

    LogActivity _
        ActionName:=ActionLabel, _
        ResultText:="FAILED", _
        ReportsGenerated:=GeneratedCount, _
        DurationSeconds:= _
            GetElapsedSeconds(StartTime), _
        DetailsText:= _
            "Stage: " & _
            CurrentStage & _
            " | Error " & _
            SavedErrNumber & _
            ": " & _
            SavedErrDescription


    '-------------------------------------------------------
    ' Error message
    '-------------------------------------------------------

    MsgBox _
        "Client report generation failed." & _
        vbCrLf & vbCrLf & _
        "Stage:" & _
        vbCrLf & _
        CurrentStage & _
        vbCrLf & vbCrLf & _
        "Error " & _
        SavedErrNumber & _
        vbCrLf & _
        SavedErrDescription, _
        vbCritical, _
        "Report Generation Error"

End Sub


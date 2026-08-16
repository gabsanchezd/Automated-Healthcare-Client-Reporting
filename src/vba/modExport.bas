Attribute VB_Name = "modExport"
Option Explicit

Public Sub OpenReportFolder()

    Dim ProjectRoot As String
    Dim ReportFolder As String
    Dim fso As Object

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    ProjectRoot = GetProjectRoot()

    ReportFolder = _
        ProjectRoot & _
        "\output\client_reports"

    If Not fso.FolderExists(ReportFolder) Then

        MsgBox _
            "The client report folder does not exist yet." & _
            vbCrLf & vbCrLf & _
            "Generate client reports first.", _
            vbInformation, _
            "Report Folder"

        Exit Sub

    End If

    Shell _
        "explorer.exe """ & ReportFolder & """", _
        vbNormalFocus

    Exit Sub


ErrorHandler:

    MsgBox _
        "Unable to open the report folder." & _
        vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Folder Error"

End Sub


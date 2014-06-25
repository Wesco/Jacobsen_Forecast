Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : Sub ImportForecast
' Date : 6/25/2014
' Desc : Prompts user for forecast, imports it, then deletes the original file
'---------------------------------------------------------------------------------------
Sub ImportForecast(FileFilter As String, Title As String, Destination As Range)
    Dim Path As String

    Path = Application.GetOpenFilename(FileFilter, Title:=Title)

    If Path <> "False" Then
        Workbooks.Open Path
        ActiveSheet.UsedRange.Copy Destination:=Destination
        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close

        DeleteFile Path
    Else
        Err.Raise Errors.USER_INTERRUPT, "ImportForecast", "User Aborted Import"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportMaster
' Date : 6/25/2014
' Desc : Imports the Jacobsen master list
'---------------------------------------------------------------------------------------
Sub ImportMaster()
    Dim Path As String
    Dim File As String
    Dim MasterWkbk As Workbook
    Dim PrevDispAlerts As Boolean
    Dim PrevUpdateLnks As Boolean

    Path = "\\br3615gaps\gaps\Billy Mac-Master Lists\"
    File = "Jacobsen-Textron Master File " & Format(Date, "yyyy") & ".xls"
    PrevDispAlerts = Application.DisplayAlerts
    PrevUpdateLnks = Application.AskToUpdateLinks

    'Disable prompts
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    If FileExists(Path & File) Then
        Workbooks.Open FileName:=sPath
        Set MasterWkbk = ActiveWorkbook

        Sheets("Master").Select

        'Unhide data and remove filters
        ActiveSheet.AutoFilterMode = False
        ActivSheet.Columns.Hidden = False
        ActiveSheet.Rows.Hidden = False

        ActiveSheet.UsedRange.Copy

        ThisWorkbook.Activate
        Sheets("Master").Select
        Range("A1").PasteSpecial Paste:=xlPasteValues, _
                                 Operation:=xlNone, _
                                 SkipBlanks:=False, _
                                 Transpose:=False
        Application.CutCopyMode = False
        MasterWkbk.Close

        Application.DisplayAlerts = PrevDispAlerts
        Application.AskToUpdateLinks = PrevUpdateLnks
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "ImportMaster", File & " could not be found."
    End If
End Sub

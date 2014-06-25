Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportPdc
' Date  : 10/10/2012
' Desc  : Prompts user for pdc forecast and copies it to this workbook
'---------------------------------------------------------------------------------------
Public Sub ImportPdc()
    Dim sPath As String
    sPath = Application.GetOpenFilename("pdc (*.csv), pdc.csv", Title:="Open the Pdc forecast")

    If sPath <> "False" Then
        Workbooks.Open sPath
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Pdc").Range("A1")
        ActiveWorkbook.Close
        Kill sPath
    Else
        MsgBox Prompt:="File import canceled." & vbCrLf & Err.Description, Title:="User Aborted Operation"
        Err.Raise 95
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportMfg
' Date  : 10/10/2012
' Desc  : Prompts user for mfg forecast and copies it to this workbook
'---------------------------------------------------------------------------------------
Public Sub ImportMfg()
    Dim sPath As String
    sPath = Application.GetOpenFilename("mfg (*.csv), mfg.csv", Title:="Open the Mfg forecast")

    If sPath <> "False" Then
        Workbooks.Open sPath
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Mfg").Range("A1")
        ActiveWorkbook.Close
        Kill sPath
    Else
        MsgBox Prompt:="File import canceled." & vbCrLf & Err.Description, Title:="User Aborted Operation"
        Err.Raise 95
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

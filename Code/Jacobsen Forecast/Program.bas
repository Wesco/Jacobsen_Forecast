Attribute VB_Name = "Program"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub Main
' Date  : 10/11/2012
' Desc  : Main procedure, calls other methods and handles errors
'---------------------------------------------------------------------------------------
Sub Main()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    'Import external data sheets
    'On failure clear all data and
    'stop execution
    On Error GoTo Main_Error
    ImportMaster
    ImportGaps SimsAsText:=False
    ImportKitBOM
    ImportExpedite

    'Import Pdc forecast
    ImportForecast FileFilter:="pdc (*.csv), pdc.csv", _
                   Title:="Open the Pdc forecast", _
                   Destination:=Sheets("Pdc").Range("A1")

    'Import Mfg forecast
    ImportForecast FileFilter:="mfg (*.csv), mfg.csv", _
                   Title:="Open the Mfg forecast", _
                   Destination:=Sheets("Mfg").Range("A1")
    On Error GoTo 0

    FormatFcst Worksheets("Pdc")  'Format Pdc forecast
    FormatFcst Worksheets("Mfg")  'Format Mfg forecast
    MergeForecast   'Consolidate data
    FormatKitBOM    'Remove unused data from raw kit BOM
    CreateKitBOM    'Create the kit BOM using forecast data

    BuildKitFcst    'Create a kit forecast
    BuildFcst       'Create a forecast
    ExportFcst      'Save the forecast to the network
    Email SendTo:="JBarnhill@wesco.com", _
          CC:="ACoffey@wesco.com", _
          Subject:="Jacobsen Forecast", _
          Body:="""\\br3615gaps\gaps\Jacobsen-Textron\" & Format(Date, "yyyy") & " Alerts\Jacobsen Slink " & Format(Date, "m-dd-yy") & ".xlsx"""
    MsgBox "Complete!"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

Main_Error:
    If Err.Number <> 18 Then
        MsgBox Prompt:="Error " & Err.Number & " (" & Err.Description & ") occurred in " & Err.Source & ".", _
               Title:="Oops!"
    End If
    Clean
End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim s As Worksheet

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next

    Application.DisplayAlerts = PrevDispAlert

    Sheets("Macro").Select
    Range("C7").Select
End Sub

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
    On Error GoTo ErrHandler
    ImportMaster
    ImportGaps
    ImportPdc
    ImportMfg
    On Error GoTo 0

    'If all data imports are successful
    'Restructure the data into a useable
    'format for processing
    RestructFcst Worksheets("Pdc")
    RestructFcst Worksheets("Mfg")

    CombineFcst     'Consolidate data
    BuildFcst       'Process data and create a forecast
    SortByColor     'Sort the forecast by color
    AddNotes        'Add last weeks expedite notes
    ExportFcst      'Save the forecast to the network
    MsgBox "Complete!"
    Email SendTo:="JBarnhill@wesco.com", _
          CC:="ACoffey@wesco.com", _
          Subject:="Jacobsen Forecast", _
          Body:="""\\br3615gaps\gaps\Jacobsen-Textron\" & Format(Date, "yyyy") & " Alerts\Jacobsen Slink " & Format(Date, "m-dd-yy") & ".xlsx"""
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

    'On import error clean all sheets and stop execution
ErrHandler:
    Clean
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim s As Worksheet

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next
    
    Application.DisplayAlerts = PrevDispAlert
    
    Sheets("Macro").Select
    Range("C7").Select
End Sub

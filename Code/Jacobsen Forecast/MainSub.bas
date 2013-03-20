Attribute VB_Name = "MainSub"
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
          Subject:="Jacobsen Forecast", _
          Body:="""\\br3615gaps\gaps\Jacobsen-Textron\" & Format(Date, "yyyy") & " Alerts\Jacobsen Slink " & Format(Date, "m-dd-yy") & ".xlsx"""
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

    'On import error clean all sheets and stop execution
ErrHandler:
    ThisWorkbook.Worksheets("Gaps").Cells.Delete
    ThisWorkbook.Worksheets("Temp").Cells.Delete
    ThisWorkbook.Worksheets("Forecast").Cells.Delete
    ThisWorkbook.Worksheets("Combined").Cells.Delete
    ThisWorkbook.Worksheets("Pdc").Cells.Delete
    ThisWorkbook.Worksheets("Mfg").Cells.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

End Sub

Attribute VB_Name = "MergeData"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub CombineFcst
' Date  : 10/11/2012
' Desc  : Combines the pdc and mfg forecasts
'---------------------------------------------------------------------------------------
Sub CombineFcst()
    Dim iRows As Long
    Dim iCols As Integer
    Dim sColHeaders() As String
    Dim i As Integer

    'Moves both part forecasts onto one sheet
    Worksheets("Pdc").Select
    iRows = Rows(Rows.Count).End(xlUp).Row
    iCols = Columns(Columns.Count).End(xlToLeft).Column
    Range(Cells(1, 1), Cells(iRows, iCols)).Copy Destination:=Worksheets("Temp").Range("A1")

    Worksheets("Mfg").Select
    Range(Cells(2, 1), Cells(Rows(Rows.Count).End(xlUp).Row, Columns(Columns.Count).End(xlToLeft).Column)).Copy Destination:=Sheets("Temp").Cells(iRows + 1, 1)

    Worksheets("Temp").Select
    ReDim sColHeaders(1 To 1, 1 To iCols)

    For i = 1 To iCols
        sColHeaders(1, i) = Cells(1, i).Text
    Next i

    'Consolidate the data by creating a pivot table
    Columns("O:Z").Delete

    ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=ActiveSheet.UsedRange, _
            Version:=xlPivotTableVersion14) _
            .CreatePivotTable _
            TableDestination:=Worksheets("Combined").Range("A1"), _
            TableName:="PivotTable1", _
            DefaultVersion:=xlPivotTableVersion14

    Worksheets("Combined").Select
    Range("A1").Select

    'Setup the pivot tables fields
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .LayoutForm = xlTabular
    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Description")
        .Orientation = xlRowField
        .Position = 2
    End With

    On Error Resume Next
    For i = 3 To iCols
        With ActiveSheet.PivotTables("PivotTable1")
            .AddDataField .PivotFields(sColHeaders(1, i)), "Sum of " & sColHeaders(1, i), xlSum
        End With
    Next i
    On Error GoTo 0

    iRows = Rows(Rows.Count).End(xlUp).Row
    iCols = Columns(Columns.Count).End(xlToLeft).Column

    'Store the pivot table as values
    Range(Cells(1, 1), Cells(iRows, iCols)).Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Rows(iRows).Delete Shift:=xlUp
    Range(Cells(1, 1), Cells(1, iCols)) = sColHeaders

    'Match part numbers to SIM numbers
    Columns(1).Insert Shift:=xlToRight
    Range("A1").Value = "SIM"
    Range("A2:A" & iRows).Formula = "=IFERROR(IF(VLOOKUP(B2,Master!A:B,2,FALSE)=0, """", VLOOKUP(B2,Master!A:B,2,FALSE)), """")"
    Range("A2:A" & iRows).Value = Range("A2:A" & iRows).Value
End Sub


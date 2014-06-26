Attribute VB_Name = "MergeData"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : MergeForecast
' Date : 6/26/2014
' Desc : Combines the pdc and mfg forecasts
'---------------------------------------------------------------------------------------
Sub MergeForecast()
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim ColHeaders As Variant
    Dim PivField As String
    Dim i As Long

    'Check how many rows are on the Pdc forecast
    Sheets("Pdc").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Copy Mfg forecast to the Pdc sheet
    Sheets("Mfg").Select
    ActiveSheet.UsedRange.Copy Destination:=Sheets("Pdc").Cells(TotalRows + 1, 1)

    Sheets("Pdc").Select
    Rows(TotalRows + 1).Delete  'Remove Mfg headers
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    'Create pivot table
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                      SourceData:=Range(Cells(1, 1), Cells(TotalRows, TotalCols)), _
                                      Version:=xlPivotTableVersion14).CreatePivotTable _
                                      TableDestination:="Combined!R1C1", _
                                      TableName:="PivotTable1", _
                                      DefaultVersion:=xlPivotTableVersion14
    Sheets("Combined").Select
    With ActiveSheet.PivotTables("PivotTable1")
        .PivotFields("Item").Orientation = xlRowField
        .PivotFields("Item").Position = 1
        .ColumnGrand = False

        For i = 3 To UBound(ColHeaders, 2)
            PivField = Format(ColHeaders(1, i), "mmm yyyy")
            .AddDataField .PivotFields(PivField), "Sum of " & PivField, xlSum
        Next
    End With

    'Convert pivot table to a range
    ActiveSheet.UsedRange.Copy
    Range("A1").PasteSpecial xlPasteValues
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Fix column headers
    Range("A1").Value = "Part Number"
    For i = 2 To TotalCols
        Cells(1, i).Value = Replace(Cells(1, i).Value, "Sum of ", "")
        Cells(1, i).NumberFormat = "mmm yyyy"
    Next

    'Insert SIMs
    Columns(2).Insert
    Range("B1").Value = "SIM"
    Range("B2:B" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2, Master!A:B, 2, FALSE), """")"
    Range("B2:B" & TotalRows).Value = Range("B2:B" & TotalRows).Value
End Sub

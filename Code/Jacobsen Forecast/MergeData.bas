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
                                      TableDestination:="PivotTable!R1C1", _
                                      TableName:="PivotTable1", _
                                      DefaultVersion:=xlPivotTableVersion14
    Sheets("PivotTable").Select
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

'---------------------------------------------------------------------------------------
' Proc : MergeKitBOM
' Date : 6/26/2014
' Desc : Merges the kit BOM with the forecast
'---------------------------------------------------------------------------------------
Sub MergeKitBOM()
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim FcstRows As Long

    Sheets("PivotTable").Select
    FcstRows = Rows(Rows.Count).End(xlUp).Row

    'Copy kit data to combined forecast
    Sheets("Kit").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    Range("C2:C" & TotalRows).Copy Destination:=Sheets("PivotTable").Cells(FcstRows + 1, 2)
    Range("E2:Q" & TotalRows).Copy Destination:=Sheets("PivotTable").Cells(FcstRows + 1, 3)

    'Add part numbers to kit data
    Sheets("PivotTable").Select
    TotalRows = Columns(2).Rows(Rows.Count).End(xlUp).Row
    Range(Cells(FcstRows + 1, 1), Cells(TotalRows, 1)).Formula = "=IFERROR(INDEX(A$2:A$" & FcstRows & ",MATCH(B" & FcstRows + 1 & ",B$2:B$" & FcstRows & ",0)),"""")"
    Range(Cells(FcstRows + 1, 1), Cells(TotalRows, 1)).Value = Range(Cells(FcstRows + 1, 1), Cells(TotalRows, 1)).Value
End Sub

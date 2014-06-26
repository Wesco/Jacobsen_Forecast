Attribute VB_Name = "FormatData"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : FormatFcst
' Date : 6/25/2014
' Desc : Aggregates columns by month
' Ex   : FormatFcst Worksheets("SheetName")
'---------------------------------------------------------------------------------------
Sub FormatFcst(Source As Worksheet)
    Dim TotalRows As Long       'Total number of rows
    Dim TotalCols As Integer    'Total number of columns
    Dim StartCol As Integer     'Months starting column
    Dim EndCol As Integer       'Months ending column
    Dim CurrCell As Range       'Current loop iterations cell
    Dim PrevCell As Range       'Previous loop iterations cell
    Dim NextCell As Range       'Next loop iterations cell
    Dim i As Integer

    Source.Select

    'Remove report header
    Rows("1:5").Delete

    'Remove Agreement and Past Due columns
    Columns("E:F").Delete Shift:=xlToLeft

    'Remove Supplier Item and Item Rev columns
    Columns("B:C").Delete Shift:=xlToLeft

    'Add Item and Description to the column headers
    Range("A1:B1").Value = Array("Item", "Description")

    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column

    'Setup the first cell for processing
    Set CurrCell = Cells(1, TotalCols)
    CurrCell.Value = Replace(CurrCell.Value, "Day ", "")
    CurrCell.Value = Replace(CurrCell.Value, "Week ", "")
    CurrCell.Value = Replace(CurrCell.Value, "Buffer ", "")
    CurrCell.Value = Replace(CurrCell.Value, "Month ", "")
    CurrCell.Value = Format(CurrCell.Value, "mmm yyyy")
    CurrCell.NumberFormat = "mmm yyyy"

    'Combine columns by month
    For i = TotalCols To 3 Step -1
        Set PrevCell = Cells(1, i + 1)
        Set CurrCell = Cells(1, i)
        Set NextCell = Cells(1, i - 1)

        NextCell.Value = Replace(NextCell.Value, "Day ", "")
        NextCell.Value = Replace(NextCell.Value, "Week ", "")
        NextCell.Value = Replace(NextCell.Value, "Buffer ", "")
        NextCell.Value = Replace(NextCell.Value, "Month ", "")
        NextCell.Value = Format(NextCell.Value, "mmm yyyy")
        NextCell.NumberFormat = "mmm yyyy"

        If CurrCell.Value <> PrevCell.Value Then
            EndCol = i
        End If

        If CurrCell.Value <> NextCell.Value Then
            EndCol = EndCol + 1
            StartCol = i + 1

            Columns(i).Insert
            Cells(1, i).Value = Cells(1, StartCol).Value
            Range(Cells(2, i), Cells(TotalRows, i)).Formula = "=SUM(" & Range(Cells(2, StartCol), Cells(2, EndCol)).Address(False, False) & ")"
            Range(Cells(2, i), Cells(TotalRows, i)).Value = Range(Cells(2, i), Cells(TotalRows, i)).Value
            Range(Cells(1, StartCol), Cells(1, EndCol)).EntireColumn.Delete
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : FormatKitBOM
' Date : 6/26/2014
' Desc : Removes unused data from the kit bom
'---------------------------------------------------------------------------------------
Sub FormatKitBOM()
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim Kits() As String
    Dim i As Long

    Sheets("Kit").Select
    TotalCols = Rows(2).Columns(Columns.Count).End(xlToLeft).Column
    ColHeaders = Array("SIM", "Record Type", "Comp SIM", "Qty")

    'Remove report header
    Rows("1:4").Delete

    'Remove unused columns
    Columns("H:S").Delete
    Columns("D:D").Delete
    Columns("A:B").Delete

    'Add column headers
    Rows(1).Insert
    Range("A1:D1").Value = ColHeaders

    'Recount rows and columns
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    'Convert SIMs to numbers
    Columns(1).Insert
    Range("A1").Value = "SIM"
    Range("A2:A" & TotalRows).Formula = "=SUBSTITUTE(B2,""'"","""")"
    Range("A1:A" & TotalRows).Value = Range("A1:A" & TotalRows).Value
    Columns(2).Delete

    'Convert Comp SIMs to numbers
    Columns(3).Insert
    Range("C1").Value = "Comp SIM"
    Range("C2:C" & TotalRows).Formula = "=SUBSTITUTE(D2,""'"","""")"
    Range("C2:C" & TotalRows).Value = Range("C2:C" & TotalRows).Value
    Columns(4).Delete

    'Remove non-component lines
    ActiveSheet.UsedRange.AutoFilter Field:=2, Criteria1:="<>J", Criteria2:="<>I", Operator:=xlAnd
    Cells.Delete
    Rows(1).Insert
    Range("A1:D1").Value = ColHeaders
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    
    'Filter kits on the forecast
    Range("E1").Value = "On Forecast"
    Range("E2:E" & TotalRows).Formula = "=IFERROR(IF(B2=""J"",VLOOKUP(C2,PivotTable!B:B,1,FALSE),""""),"""")"
    Range("E2:E" & TotalRows).Value = Range("E2:E" & TotalRows).Value
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    
    ReDim Kits(1 To 1) As String
    For i = 2 To TotalRows
        If Not Cells(i, 5).Value = "" Then
            If Not Kits(1) = "" Then
                ReDim Preserve Kits(1 To UBound(Kits) + 1) As String
            End If
            
            Kits(UBound(Kits)) = Cells(i, 5).Value
        End If
        
        If Cells(i, 1).Value = Kits(UBound(Kits)) Then
            Cells(i, 5).Value = Kits(UBound(Kits))
        End If
    Next
    
    Range(Cells(1, 1), Cells(TotalRows, TotalCols)).AutoFilter Field:=5, Criteria1:="=", Operator:=xlAnd
    Cells.Delete
    Columns("E:E").Delete
    Rows(1).Insert
    Range("A1:D1").Value = ColHeaders
End Sub

'---------------------------------------------------------------------------------------
' Proc  : SortByColor
' Date  : 10/17/2012
' Desc  : Sorts the finished forecast by color to group bulk SIMs
'---------------------------------------------------------------------------------------
Sub SortByColor()
    Dim vCell As Variant
    Dim TotalRows As Long
    Dim PrevSheet As Worksheet

    Set PrevSheet = ActiveSheet
    Sheets("Forecast").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    With ActiveWorkbook.Worksheets("Forecast").ListObjects("Table1").Sort.SortFields
        .Clear
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(255, 235, 156)
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(204, 192, 218)
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(198, 239, 206)
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(183, 222, 232)
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(230, 184, 183)
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(252, 213, 180)
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(196, 189, 151)
        .Add(Range("Table1[SIM]"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color = RGB(184, 204, 228)
        .Add Key:=Range("Table1[LT/Days]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    End With
    With ActiveWorkbook.Worksheets("Forecast").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    For Each vCell In Range(Cells(2, 1), Cells(TotalRows, 1))
        If vCell.Value = "99923698005" Or vCell.Value = "99923698006" Or _
           vCell.Value = "99923697662" Or vCell.Value = "99420498967" Then
            Rows(vCell.Row).Cut
            Rows(vCell.Offset(-1).Row).Insert Shift:=xlDown
        End If
    Next

    PrevSheet.Select
End Sub

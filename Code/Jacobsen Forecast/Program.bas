Attribute VB_Name = "Program"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub FormatFcst
' Date  : 10/11/2012
' Desc  : Aggregates fractured months
' Ex    : RestructFcst Worksheets("SheetName")
'---------------------------------------------------------------------------------------
Sub RestructFcst(ByVal WS As Worksheet)
    Dim sCell1 As String           'Address of a months first column, 2nd row
    Dim sCell2 As String           'Address of a months last column, 2nd row
    Dim sRange As String           'sCell1 and sCell2 combined for use as a range
    Dim aColHeaders() As Variant   'Column headers containing dates
    Dim iColNum() As Integer       'The last column number of each fractured month
    Dim iDist As Integer           'Number of columns from a months first column to its last
    Dim sLast As String            'Used to compare the previous column to the current column
    Dim iRows As Long              'Used range row count
    Dim n As Integer               'iColNum counter
    Dim i As Integer               'aColHeaders counter

    'Select worksheet
    WS.Select
    'Remove data that is not needed
    Rows("1:5").Delete
    Columns("E:F").Delete Shift:=xlToLeft
    Columns("B:C").Delete Shift:=xlToLeft
    'Fix column headers
    Range("A1:B1").Value = Array("Item", "Description")
    'Store column headers
    aColHeaders = Range(Cells(1, 3), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    'Set initial values
    sLast = Format(Replace(aColHeaders(1, 1), "Day ", ""), "mmm yyyy")
    iRows = ActiveSheet.UsedRange.Rows.Count
    'Instantiate arrays
    ReDim iColNum(1 To UBound(aColHeaders, 2))


    'Remove prefixes from column headers
    'Format column header dates
    For i = 1 To UBound(aColHeaders, 2)
        If InStr(aColHeaders(1, i), "Day ") Then
            aColHeaders(1, i) = Format(Replace(aColHeaders(1, i), "Day ", ""), "mmm yyyy")
            If aColHeaders(1, i) <> sLast Then
                n = n + 1
                iColNum(n) = i + 2
            End If
        End If

        If InStr(aColHeaders(1, i), "Week ") Then
            aColHeaders(1, i) = Format(Replace(aColHeaders(1, i), "Week ", ""), "mmm yyyy")
            If aColHeaders(1, i) <> sLast Then
                n = n + 1
                iColNum(n) = i + 2
            End If
        End If

        If InStr(aColHeaders(1, i), "Buffer ") Then
            aColHeaders(1, i) = Format(Replace(aColHeaders(1, i), "Buffer ", ""), "mmm yyyy")
            If aColHeaders(1, i) <> sLast Then
                n = n + 1
                iColNum(n) = i + 2
            End If
        End If

        If InStr(aColHeaders(1, i), "Month ") Then
            aColHeaders(1, i) = Format(Replace(aColHeaders(1, i), "Month ", ""), "mmm yyyy")
            If aColHeaders(1, i) <> sLast Then
                n = n + 1
                iColNum(n) = i + 2
            End If
        End If
        sLast = aColHeaders(1, i)
    Next i

    'Set date column headers
    Range(Cells(1, 3), Cells(1, ActiveSheet.UsedRange.Columns.Count)) = aColHeaders

    'Insert a new column at the end of each split month
    'Subtotal the months part usage and remove the columns
    'that are no longer needed
    For i = UBound(iColNum) To 1 Step -1
        If iColNum(i) <> 0 Then
            Columns(iColNum(i)).Insert Shift:=xlToRight
            Cells(1, iColNum(i)).Value = Cells(1, iColNum(i) - 1)

            If i = 1 Then
                sCell1 = Cells(2, 3).Address(False, False)
                sCell2 = Cells(2, iColNum(i) - 1).Address(False, False)

                Cells(2, iColNum(i)).Formula = "=SUM(" & sCell1 & ":" & sCell2 & ")"
                Cells(2, iColNum(i)).AutoFill _
                        Destination:=Range(Cells(2, iColNum(i)), Cells(ActiveSheet.UsedRange.Rows.Count, iColNum(i)))

                With Range(Cells(2, iColNum(i)), Cells(ActiveSheet.UsedRange.Rows.Count, iColNum(i)))
                    .Value = .Value
                End With
                sRange = sCell1 & ":" & sCell2
                Range(sRange).EntireColumn.Delete Shift:=xlToLeft
            Else
                iDist = iColNum(i) - iColNum(i - 1)
                iDist = iColNum(i) - iDist

                sCell1 = Cells(2, iDist).Address(False, False)
                sCell2 = Cells(2, iColNum(i) - 1).Address(False, False)
                Cells(2, iColNum(i)).Formula = "=SUM(" & sCell1 & ":" & sCell2 & ")"

                Cells(2, iColNum(i)).AutoFill _
                        Destination:=Range(Cells(2, iColNum(i)), Cells(ActiveSheet.UsedRange.Rows.Count, iColNum(i)))

                With Range(Cells(2, iColNum(i)), Cells(ActiveSheet.UsedRange.Rows.Count, iColNum(i)))
                    .Value = .Value
                End With

                sRange = sCell1 & ":" & sCell2
                Range(sRange).EntireColumn.Delete Shift:=xlToLeft
            End If
        End If
    Next i

    'Add bulk kit item to the bottom of the raw forecast
    'Pull in the part description from gaps
    'Write the formula as a value
    'Set the part forecast to 0 for each month
    AddBulkSIM "40309495373"
    AddBulkSIM "78923694616"
    AddBulkSIM "78420420014"
    AddBulkSIM "78420420179"
    AddBulkSIM "78923693664"
    AddBulkSIM "78923693663"
    AddBulkSIM "63285098955"
    AddBulkSIM "63285098954"
    AddBulkSIM "78862198856"
    AddBulkSIM "78923693770"
    AddBulkSIM "78923693769"
    AddBulkSIM "78420498874"
    AddBulkSIM "78923693548"
    AddBulkSIM "78008097520"
    AddBulkSIM "78862198844"
    AddBulkSIM "78862198843"
    AddBulkSIM "78862198842"
End Sub

Private Sub AddBulkSIM(ItemNum As String)
    Dim TotalCols As Integer
    Dim TotalRows As Long

    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = Rows(Rows.Count).End(xlUp).Row + 1

    'Add SIM
    Cells(TotalRows, 1).Value = ItemNum

    'Lookup description
    Cells(TotalRows, 2).Formula = "=IFERROR(VLOOKUP(" & Cells(TotalRows, 1).Address(False, False) & ", Gaps!D:E, 2, FALSE),"""")"
    Cells(TotalRows, 2).Value = Cells(TotalRows, 2).Value

    Range(Cells(TotalRows, 3), Cells(TotalRows, TotalCols)).Value = 0
End Sub

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

'---------------------------------------------------------------------------------------
' Proc  : Sub BuildFcst
' Date  : 10/17/2012
' Desc  : Takes the combined forecasts and creates a final forecast sheet for use by wesco
'---------------------------------------------------------------------------------------
Sub BuildFcst()
    Dim iCombinedCols As Integer
    Dim iCombinedRows As Long
    Dim iRows As Long
    Dim iCols As Long
    Dim i As Long
    Dim Item As Object
    Dim aArray() As Variant
    Dim rRange As Range
    Dim sCol1 As String
    Dim sCol2 As String
    Dim TotalRows As Long

    Worksheets("Combined").Select
    iCombinedCols = ActiveSheet.UsedRange.Columns.Count
    iCombinedRows = ActiveSheet.UsedRange.Rows.Count

    With ActiveSheet.UsedRange
        Range(Cells(2, 1), Cells(.CurrentRegion.Rows.Count, 3)).Copy Destination:=Worksheets("Forecast").Range("A2")
    End With

    'Set column headers & number formats
    Worksheets("Forecast").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Range("A1:C1") = Array("SIM", "Part", "Description")
    Range("L1").Value = "Stock Visualization"


    'Add data
    Range("D1:K1") = Array("On Hand", "Reserve", "On Order", "BO", "WDC", "Last Cost", "UOM", "Supplier")

    Range("D2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:F,3,FALSE),""0"")"
    Range("E2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:G,4,FALSE),""0"")"
    Range("F2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:I,6,FALSE),""0"")"
    Range("G2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!G:H,5,FALSE),""0"")"
    Range("H2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:AJ,33,FALSE),""0"")"
    Range("I2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:AE,28,FALSE),""0"")"
    Range("J2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:AI,32,FALSE),"""")"
    Range("K2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:AL,35,FALSE),"""")"

    Range("D2").AutoFill Destination:=Range(Cells(2, 4), Cells(TotalRows, 4))
    Range("E2").AutoFill Destination:=Range(Cells(2, 5), Cells(TotalRows, 5))
    Range("F2").AutoFill Destination:=Range(Cells(2, 6), Cells(TotalRows, 6))
    Range("G2").AutoFill Destination:=Range(Cells(2, 7), Cells(TotalRows, 7))
    Range("H2").AutoFill Destination:=Range(Cells(2, 8), Cells(TotalRows, 8))
    Range("I2").AutoFill Destination:=Range(Cells(2, 9), Cells(TotalRows, 9))
    Range("J2").AutoFill Destination:=Range(Cells(2, 10), Cells(TotalRows, 10))
    'Supplier
    Range(Cells(2, 11), Cells(TotalRows, 11)).NumberFormat = "@"
    Range("K2").AutoFill Destination:=Range(Cells(2, 11), Cells(TotalRows, 11))

    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
    iRows = ActiveSheet.UsedRange.Rows.Count

    sCol1 = "B"
    sCol2 = Columns(iCombinedCols).Address(False, False)
    sCol2 = Left(sCol2, 2)
    If Right(sCol2, 1) = ":" Then
        sCol2 = Left(sCol2, 1)
    End If

    'Autofill data
    Range("M1").Value = Worksheets("Combined").Range("D1").Value
    Range("M2").Formula = "=IFERROR(VLOOKUP(A2,Gaps!D:F,3,FALSE),0)-VLOOKUP(B2,Combined!" & sCol1 & ":" & sCol2 & ",3,FALSE)"
    Cells(2, 13).AutoFill Destination:=Range(Cells(2, 13), Cells(iRows, 13))

    For i = 5 To iCombinedCols
        Cells(1, i + 9).Value = Worksheets("Combined").Cells(i).Value
        Cells(2, i + 9).Formula = _
        "=" & Cells(2, i + 8).Address(False, False) & "-VLOOKUP(B2,Combined!" & sCol1 & ":" & sCol2 & "," & i - 1 & ",FALSE)"
        Cells(2, i + 9).AutoFill Destination:=Range(Cells(2, i + 9), Cells(iRows, i + 9))
    Next i

    iCols = ActiveSheet.UsedRange.Columns.Count + 1
    Cells(1, iCols).Value = "Notes"
    Cells(2, iCols).Formula = "=IFERROR(IF(VLOOKUP(B2,Master!A:L,12,FALSE)=0,"""",VLOOKUP(B2,Master!A:L,12,FALSE)),"""")"
    Cells(2, iCols).AutoFill Destination:=Range(Cells(2, iCols), Cells(iRows, iCols))


    With ActiveSheet.UsedRange
        For i = 1 To .CurrentRegion.Columns.Count
            Range(Cells(1, i), Cells(ActiveSheet.UsedRange.Rows.Count, i)).Value = _
            Range(Cells(1, i), Cells(ActiveSheet.UsedRange.Rows.Count, i)).Value
        Next
    End With

    'Add sparklines
    Range("L2").Select
    Range("L2").SparklineGroups.Add _
            Type:=xlSparkColumn, _
            SourceData:=Range(Cells(2, 13), Cells(2, ActiveSheet.UsedRange.Columns.Count - 1)).Address(False, False)

    Selection.SparklineGroups.Item(1).Points.Negative.Visible = True
    Selection.SparklineGroups.Item(1).SeriesColor.Color = 3289650
    Selection.SparklineGroups.Item(1).SeriesColor.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Negative.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Negative.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Markers.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Highpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lowpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Firstpoint.Color.TintAndShade = 0
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.Color = 208
    Selection.SparklineGroups.Item(1).Points.Lastpoint.Color.TintAndShade = 0
    With Range("L:L")
        Range("L2").AutoFill Destination:=Range(Cells(2, 12), Cells(.CurrentRegion.Rows.Count, 12))
    End With

    Range(Cells(2, 13), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count - 1)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    'Create table
    With Range("A:A")
        ActiveSheet.ListObjects.Add( _
                xlSrcRange, Range(Cells(1, 1), Cells(.CurrentRegion.Rows.Count, .CurrentRegion.Columns.Count)), , _
                xlYes).Name = "Table1"
    End With

    '''NOTE''''''''''''''''''''''''''''''''''''''''''''''
    ' The formula autofills since it is part of a table '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Range("K:K").Insert Shift:=xlToRight
    Range("K1").Value = "LT/Weeks"
    Range("K2").Formula = "=IFERROR(VLOOKUP(B2,Master!A:N,14,FALSE)/7,"""")"
    Range(Cells(1, 11), Cells(ActiveSheet.UsedRange.Rows.Count, 11)).Value = _
    Range(Cells(1, 11), Cells(ActiveSheet.UsedRange.Rows.Count, 11)).Value

    Range("K:K").Insert Shift:=xlToRight
    Range("K1").Value = "LT/Days"
    Range("K2").Formula = "=IFERROR(VLOOKUP(B2,Master!A:N,14,FALSE),"""")"
    Range(Cells(1, 11), Cells(ActiveSheet.UsedRange.Rows.Count, 11)).Value = _
    Range(Cells(1, 11), Cells(ActiveSheet.UsedRange.Rows.Count, 11)).Value

    Range("K:K").Insert Shift:=xlToRight
    Range("K1").Value = "Min/Mult"
    Range("K2").Formula = "=IFERROR(VLOOKUP([@Part],Master!A:M,13,FALSE),"""")"
    Range(Cells(1, 11), Cells(ActiveSheet.UsedRange.Rows.Count, 11)).Value = _
    Range(Cells(1, 11), Cells(ActiveSheet.UsedRange.Rows.Count, 11)).Value

    'Fix text alignment
    Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count)).HorizontalAlignment = xlCenter
    Range(Cells(2, 2), Cells(ActiveSheet.UsedRange.Rows.Count, 2)).HorizontalAlignment = xlCenter
    Range(Cells(2, 4), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count - 1)).HorizontalAlignment = xlCenter
    Range(Cells(2, ActiveSheet.UsedRange.Columns.Count), _
          Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).HorizontalAlignment = xlLeft
    Cells.EntireColumn.AutoFit

    Columns("G:G").Insert
    Range("G1").Value = "Net Stock"
    Range("G2").Formula = "=SUM(D2,F2)"
    Range("G2").AutoFill Destination:=Range(Cells(2, 7), Cells(ActiveSheet.UsedRange.Rows.Count, 7))
    With Range(Cells(2, 7), Cells(ActiveSheet.UsedRange.Rows.Count, 7))
        .Value = .Value
    End With

    'Color bulk item SIMs
    Set rRange = Range(Cells(1, 1), Cells(Rows(Rows.Count).End(xlUp).Row, 1))
    aArray = Range(Cells(1, 2), Cells(Rows(Rows.Count).End(xlUp).Row, 2))
    For i = 1 To UBound(aArray, 1)
        If aArray(i, 1) = "4193360" Or aArray(i, 1) = "40309495373" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "10284031"

        ElseIf aArray(i, 1) = "3005286" Or aArray(i, 1) = "78420420014" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "13561798"

        ElseIf aArray(i, 1) = "4265710" Or aArray(i, 1) = "78923694616" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "14336204"

        ElseIf aArray(i, 1) = "3010331" Or aArray(i, 1) = "78420420179" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "11851260"

        ElseIf aArray(i, 1) = "4187221" Or aArray(i, 1) = "78923693663" Or aArray(i, 1) = "78923693664" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "12040422"

        ElseIf aArray(i, 1) = "4283654" Or aArray(i, 1) = "63285098954" Or aArray(i, 1) = "63285098955" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "15261367"

        ElseIf aArray(i, 1) = "4292871" Or aArray(i, 1) = "78862198856" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "9944516"

        ElseIf aArray(i, 1) = "4283892" Or aArray(i, 1) = "78923693770" Or aArray(i, 1) = "78923693769" Or aArray(i, 1) = "78420498874" Then
            Range(rRange(i, 1), rRange(i, 15)).Interior.Color = "14994616"
        End If
    Next

    Range("Q1:AC1").Value = Range("Q1:AC1").Value
    Range("Q1:AC1").NumberFormat = "mmm-yyyy"
End Sub

'---------------------------------------------------------------------------------------
' Proc  : Sub SortByColor
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

Private Sub ColorToRGB()
    Dim HexColor As String
    Dim Red As String
    Dim Green As String
    Dim Blue As String

    HexColor = Hex(ActiveCell.Interior.Color)
    HexColor = Replace(HexColor, "#", "")
    Red = Val("&H" & Mid(HexColor, 5, 2))
    Green = Val("&H" & Mid(HexColor, 3, 2))
    Blue = Val("&H" & Mid(HexColor, 1, 2))
    Debug.Print "RGB(" & Red & ", " & Green & ", " & Blue & ")"
End Sub

'---------------------------------------------------------------------------------------
' Proc : AddNotes
' Date : 1/17/2013
' Desc : Add previous weeks expedite notes to the forecast
'---------------------------------------------------------------------------------------
Sub AddNotes()
    Dim sPath As String
    Dim sWkBk As String
    Dim sYear As String
    Dim iRows As Long
    Dim iCols As Integer
    Dim i As Integer

    Sheets("Temp").Cells.Delete

    For i = 1 To 30
        sYear = Date - i
        sWkBk = "Jacobsen Slink " & Format(sYear, "m-dd-yy") & ".xlsx"
        sPath = "\\br3615gaps\gaps\Jacobsen-Textron\" & Format(sYear, "yyyy") & " Alerts\"

        If FileExists(sPath & sWkBk) = True Then
            Workbooks.Open sPath & sWkBk

            Sheets("Expedite").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            iCols = ActiveSheet.UsedRange.Columns.Count

            Range(Cells(1, 1), Cells(iRows, 1)).Copy Destination:=ThisWorkbook.Sheets("Temp").Range("A1")
            Range(Cells(1, iCols), Cells(iRows, iCols)).Copy Destination:=ThisWorkbook.Sheets("Temp").Range("B1")
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True

            Sheets("Forecast").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            iCols = ActiveSheet.UsedRange.Columns.Count + 1

            Cells(1, iCols).Value = "Expedite Notes"
            Cells(2, iCols).Formula = "=IFERROR(IF(VLOOKUP(A2,Temp!A:B,2,FALSE)=0,"""",VLOOKUP(A2,Temp!A:B,2,FALSE)),"""")"
            Cells(2, iCols).AutoFill Destination:=Range(Cells(2, iCols), Cells(iRows, iCols))
            Range(Cells(2, iCols), Cells(iRows, iCols)).Value = Range(Cells(2, iCols), Cells(iRows, iCols)).Value
            Columns(iCols).EntireColumn.AutoFit

            Range("Q1:AC1").NumberFormat = "mmm-yyyy"
            Range("Q1:AC1").Value = Range("Q1:AC1").Value
            Exit For
        End If
    Next

    Columns("G:G").Delete
End Sub





Attribute VB_Name = "FormatData"
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

Attribute VB_Name = "CreateReport"
Option Explicit

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
' Proc : CreateKitBOM
' Date : 6/26/2014
' Desc : Fills out the kit bom to calculate part requirements
'---------------------------------------------------------------------------------------
Sub CreateKitBOM()
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim Addr As String
    Dim i As Long
    Dim j As Long
    
    Sheets("Combined").Select
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    Range("C1:O1").Copy Destination:=Sheets("Kit").Range("E1")
    
    Sheets("Kit").Select
    TotalRows = Rows(Rows.Count).End(xlUp).Row
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    
    For j = 5 To TotalCols
        For i = 2 To TotalRows
            If Cells(i, 2).Value = "J" Then
                Addr = Cells(i, j).Address(False, False)    'Address of the current KIT total
                'vlookup KIT SIM on combined forecast to get total needed for the current month
                Cells(i, j).Formula = "=IFERROR(VLOOKUP(" & Cells(i, 3).Address(False, False) & ",'Combined'!B:O," & j - 2 & ",FALSE),0)"
            Else
                'Multiply the kit total by the number of components needed per kit
                Cells(i, j).Formula = "=" & Addr & "*" & Cells(i, 4).Address(False, False)
            End If
        Next
    Next
    
    Range("E2:Q" & TotalRows).Value = Range("E2:Q" & TotalRows).Value
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

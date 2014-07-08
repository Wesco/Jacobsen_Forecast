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
    Dim TotalCols As Integer


    'Copy The Parts and SIMs
    Sheets("Combined").Select
    TotalRows = Columns(3).Rows(Rows.Count).End(xlUp).Row
    Range(Cells(1, 1), Cells(TotalRows, 2)).Copy Destination:=Sheets("Forecast").Range("A1")

    Sheets("Forecast").Select

    'Set column headers
    Range("C1:O1") = Array("Description", _
                           "On Hand", _
                           "Reserve", _
                           "On Order", _
                           "BO", _
                           "WDC", _
                           "Last Cost", _
                           "UOM", _
                           "Min/Mult", _
                           "LT/Days", _
                           "LT/Weeks", _
                           "Supplier", _
                           "Stock Visualization")

    'Add column data
    Range("C2:N" & TotalRows).Formula = Array("=IFERROR(IF(VLOOKUP(B2,Gaps!A:F,6,FALSE)=0,"""",VLOOKUP(B2,Gaps!A:F,6,FALSE)),"""")", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:G,7,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:H,8,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:J,10,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:I,9,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AK,37,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AF,32,FALSE),0)", _
                                              "=IFERROR(IF(VLOOKUP(B2,Gaps!A:AJ,36,FALSE)=0,"""",VLOOKUP(B2,Gaps!A:AJ,36,FALSE)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Master!B:M,12,FALSE)=0,"""",VLOOKUP(B2,Master!B:M,12,FALSE)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Master!B:N,13,FALSE)=0,"""",VLOOKUP(B2,Master!B:N,13,FALSE)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Master!B:N,13,FALSE)=0,"""",ROUNDUP(VLOOKUP(B2,Master!B:N,13,FALSE)/7,0)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Gaps!A:AM,39,FALSE)=0,"""",VLOOKUP(B2,Gaps!A:AM,39,FALSE)),"""")")

    'Set text formatting
    Range("A2:A" & TotalRows).NumberFormat = "@"            'Part
    Range("B2:B" & TotalRows).NumberFormat = "0000000000#"  'SIM
    Range("N2:N" & TotalRows).NumberFormat = "@"            'Supplier
    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

    'Add forecast month headers
    Range("P1:AA1").Formula = "=PivotTable!C1"
    Range("P1:AA1").NumberFormat = "mmm yyyy"
    Range("P1:AA1").Value = Range("P1:AA1").Value

    'Add forecast month data
    Range("P2:P" & TotalRows).Formula = "=D2-IFERROR(VLOOKUP(B2,PivotTable!B:N,2,FALSE),0)"
    Range("P2:P" & TotalRows).NumberFormat = "General"
    Range("P2:P" & TotalRows).Value = Range("P2:P" & TotalRows).Value

    'Columns Q to AA
    For i = 17 To 27
        Range(Cells(2, i), Cells(TotalRows, i)).Formula = "=" & Cells(2, i - 1).Address(False, False) & "-IFERROR(VLOOKUP(B2,PivotTable!B:N," & i - 14 & ",FALSE),0)"
        Range(Cells(2, i), Cells(TotalRows, i)).Value = Range(Cells(2, i), Cells(TotalRows, i)).Value
    Next

    'Add stock visualization
    With Range("O2:O" & TotalRows).SparklineGroups
        .Add Type:=xlSparkColumn, SourceData:=Range("P2:AA" & TotalRows).Address(False, False)
        With .Item(1)
            .Points.Negative.Visible = True
            .SeriesColor.Color = 3289650
            .SeriesColor.TintAndShade = 0
            .Points.Negative.Color.Color = 208
            .Points.Negative.Color.TintAndShade = 0
            .Points.Markers.Color.Color = 208
            .Points.Markers.Color.TintAndShade = 0
            .Points.Highpoint.Color.Color = 208
            .Points.Highpoint.Color.TintAndShade = 0
            .Points.Lowpoint.Color.Color = 208
            .Points.Lowpoint.Color.TintAndShade = 0
            .Points.Firstpoint.Color.Color = 208
            .Points.Firstpoint.Color.TintAndShade = 0
            .Points.Lastpoint.Color.Color = 208
            .Points.Lastpoint.Color.TintAndShade = 0
        End With
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

    Sheets("PivotTable").Select
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
                Cells(i, j).Formula = "=IFERROR(VLOOKUP(" & Cells(i, 3).Address(False, False) & ",'PivotTable'!B:O," & j - 2 & ",FALSE),0)"
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

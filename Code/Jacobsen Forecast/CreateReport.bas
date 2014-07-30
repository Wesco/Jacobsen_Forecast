Attribute VB_Name = "CreateReport"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub BuildFcst
' Date  : 7/8/2014
' Desc  : Takes the combined forecasts and creates a final forecast sheet for use by wesco
'---------------------------------------------------------------------------------------
Sub BuildFcst()
    Dim TotalRows As Long
    Dim i As Long


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
    Range("C2:N" & TotalRows).Formula = Array("=IFERROR(IF(VLOOKUP(A2,Gaps!A:F,6,FALSE)=0,"""",VLOOKUP(A2,Gaps!A:F,6,FALSE)),"""")", _
                                              "=IFERROR(VLOOKUP(A2,Gaps!A:G,7,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(A2,Gaps!A:H,8,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(A2,Gaps!A:J,10,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(A2,Gaps!A:I,9,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(A2,Gaps!A:AK,37,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(A2,Gaps!A:AF,32,FALSE),0)", _
                                              "=IFERROR(IF(VLOOKUP(A2,Gaps!A:AJ,36,FALSE)=0,"""",VLOOKUP(A2,Gaps!A:AJ,36,FALSE)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(A2,Master!B:M,12,FALSE)=0,"""",VLOOKUP(A2,Master!B:M,12,FALSE)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(A2,Master!B:N,13,FALSE)=0,"""",VLOOKUP(A2,Master!B:N,13,FALSE)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(A2,Master!B:N,13,FALSE)=0,"""",ROUNDUP(VLOOKUP(A2,Master!B:N,13,FALSE)/7,0)),"""")", _
                                              "=IFERROR(IF(VLOOKUP(A2,Gaps!A:AM,39,FALSE)=0,"""",VLOOKUP(A2,Gaps!A:AM,39,FALSE)),"""")")

    'Set text formatting
    Range("A2:A" & TotalRows).NumberFormat = "0000000000#"  'SIM
    Range("B2:B" & TotalRows).NumberFormat = "@"            'Part
    Range("N2:N" & TotalRows).NumberFormat = "@"            'Supplier
    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value

    'Add notes from master
    Range("AB1").Value = "Notes"

    'Try to lookup by part number, if the part number is not found try to lookup by SIM, if both are not found, return nothing
    Range("AB2:AB" & TotalRows).Formula = "=IFERROR(IFERROR(IF(VLOOKUP(B2,Master!A:L,12,FALSE)=0,"""",VLOOKUP(B2,Master!A:L,12,FALSE)),IF(VLOOKUP(A2,Master!A:L,12,FALSE)=0,"""",VLOOKUP(A2,Master!A:L,12,FALSE))),"""")"
    Range("AB2:AB" & TotalRows).Value = Range("AB2:AB" & TotalRows).Value

    'Add notes from previous expedite sheet
    Range("AC1").Value = "Expedite Notes"
    Range("AC2:AC" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,Expedite!A:B,2,FALSE)=0,"""",VLOOKUP(A2,Expedite!A:B,2,FALSE)),"""")"
    Range("AC2:AC" & TotalRows).Value = Range("AC2:AC" & TotalRows).Value

    'Add forecast month headers
    Range("P1:AA1").Formula = "=Combined!C1"
    Range("P1:AA1").NumberFormat = "mmm yyyy"
    Range("P1:AA1").Value = Range("P1:AA1").Value

    'Add forecast month data
    Range("P2:P" & TotalRows).Formula = "=D2-IFERROR(VLOOKUP(B2,Combined!B:N,2,FALSE),0)"
    Range("P2:P" & TotalRows).NumberFormat = "General"
    Range("P2:P" & TotalRows).Value = Range("P2:P" & TotalRows).Value

    'Columns Q to AA
    For i = 17 To 27
        Range(Cells(2, i), Cells(TotalRows, i)).Formula = "=" & Cells(2, i - 1).Address(False, False) & "-IFERROR(VLOOKUP(B2,Combined!B:N," & i - 14 & ",FALSE),0)"
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

    'If inventory is less than 0 highlight the cell
    Range("P2:AA" & TotalRows).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    With Range("P2:AA" & TotalRows).FormatConditions(1)
        .Font.Color = -16383844
        .Font.TintAndShade = 0
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 13551615
        .Interior.TintAndShade = 0
        .StopIfTrue = False
    End With
    
    'Sort by lead time
    ActiveSheet.UsedRange.Sort Range("L1:L" & TotalRows), xlDescending, Header:=xlYes

    'Create table
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:AC" & TotalRows), , xlYes).Name = "Table1"
    ActiveSheet.ListObjects(1).Unlist

    'Set text alignment
    Range("A1:AC1").HorizontalAlignment = xlCenter
    Range("C2:C" & TotalRows).HorizontalAlignment = xlLeft
    Range("D2:AA" & TotalRows).HorizontalAlignment = xlCenter
    Range("AB2:AC" & TotalRows).HorizontalAlignment = xlLeft
    ActiveSheet.UsedRange.Columns.AutoFit
End Sub

'---------------------------------------------------------------------------------------
' Proc : BuildKitFcst
' Date : 7/29/2014
' Desc : Creates a forecast using data from the kit bom
'---------------------------------------------------------------------------------------
Sub BuildKitFcst()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Long
    Dim j As Integer

    Sheets("Kit").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    For i = 2 To TotalRows
        'Calculate the first month
        Cells(i, 5).Formula = "=IFERROR(VLOOKUP(C" & i & ",Gaps!A:G,7,FALSE),0)-" & Cells(i, 5).Value
        Cells(i, 5).Value = Cells(i, 5).Value

        'Calculate the remaining months
        For j = 6 To TotalCols
            Cells(i, j).Formula = Cells(i, j - 1).Value - Cells(i, j).Value
        Next

        'Convert J/I lines to K(it)/C(omponent)
        If Cells(i, 2).Value = "J" Then Cells(i, 2).Value = "K"
        If Cells(i, 2).Value = "I" Then Cells(i, 2).Value = "C"
    Next

    Range("D:D").Delete
    Range("A:A").Delete
    Range("C:O").Insert

    Range("A1").Value = "Type"
    Range("B1").Value = "SIM"

    'Set column headers
    Range("C1:O1").Value = Array("Part", _
                                 "OH", _
                                 "OR", _
                                 "OO", _
                                 "BO", _
                                 "WDC", _
                                 "LC", _
                                 "UOM", _
                                 "Min/Mult", _
                                 "LT/Days", _
                                 "LT/Weeks", _
                                 "Sup", _
                                 "Stock Visualization")

    Range("C2:N" & TotalRows).Formula = Array("=IFERROR(IF(INDEX(Master!A:A,MATCH(B2,Master!B:B,0))=B2,"""",INDEX(Master!A:A,MATCH(B2,Master!B:B,0))),"""")", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:G,7,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:H,8,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:J,10,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:I,9,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AK,37,FALSE),0)", _
                                              "=IFERROR(VLOOKUP(B2,Gaps!A:AF,32,FALSE),""-"")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Gaps!A:AJ,36,FALSE)=0,""-"",VLOOKUP(B2,Gaps!A:AJ,36,FALSE)),""-"")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Master!B:M,12,FALSE)=0,""-"",VLOOKUP(B2,Master!B:M,12,FALSE)),""-"")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Master!B:N,13,FALSE)=0,""-"",VLOOKUP(B2,Master!B:N,13,FALSE)),""-"")", _
                                              "=IFERROR(IF(VLOOKUP(B2,Master!B:N,13,FALSE)=0,""-"",ROUNDUP(VLOOKUP(B2,Master!B:N,13,FALSE)/7,0)),""-"")", _
                                              "=IFERROR(""=""&""""""""&VLOOKUP(B2,Gaps!A:AM,39,FALSE)&"""""""",""-"")")

    Range("C2:N" & TotalRows).Value = Range("C2:N" & TotalRows).Value

    'Add stock visualization
    With Range("O2:O" & TotalRows).SparklineGroups
        .Add Type:=xlSparkColumn, SourceData:=Range("P2:AB" & TotalRows).Address(False, False)
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

    'Add expedite notes
    Range("AC1").Value = "Expedite Notes"
    Range("AC2:AC" & TotalRows).Formula = "=IFERROR(IF(VLOOKUP(A2,Expedite!A:B,2,FALSE)=0,"""",VLOOKUP(A2,Expedite!A:B,2,FALSE)),"""")"
    Range("AC2:AC" & TotalRows).Value = Range("AC2:AC" & TotalRows).Value

    'If inventory is less than 0 highlight the cell
    Range("P2:AB" & TotalRows).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    With Range("P2:AB" & TotalRows).FormatConditions(1)
        .Font.Color = -16383844
        .Font.TintAndShade = 0
        .Interior.PatternColorIndex = xlAutomatic
        .Interior.Color = 13551615
        .Interior.TintAndShade = 0
        .StopIfTrue = False
    End With

    'Set number formats
    Range("I2:I" & TotalRows).NumberFormat = "0.00"

    'Set text alignment
    Range("A1:AB1").HorizontalAlignment = xlCenter
    Range("C2:C" & TotalRows).HorizontalAlignment = xlLeft
    Range("D2:AB" & TotalRows).HorizontalAlignment = xlCenter
    Range("I2:I" & TotalRows).HorizontalAlignment = xlLeft

    'Create table
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:AC" & TotalRows), , xlYes).Name = "Table1"
    ActiveSheet.ListObjects(1).Unlist
    ActiveSheet.UsedRange.Columns.AutoFit
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
                'Vlookup KIT SIM on combined forecast to get total needed for the current month
                Cells(i, j).Formula = "=IFERROR(VLOOKUP(" & Cells(i, 3).Address(False, False) & ",'Combined'!A:O," & j - 2 & ",FALSE),0)"
            Else
                'Multiply the kit total by the number of components needed per kit
                Cells(i, j).Formula = "=" & Addr & "*" & Cells(i, 4).Address(False, False)
            End If
        Next
    Next

    Range("E2:Q" & TotalRows).Value = Range("E2:Q" & TotalRows).Value
End Sub

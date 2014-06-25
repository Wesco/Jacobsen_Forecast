Attribute VB_Name = "Imports"
Option Explicit
Public Gaps As Collection

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportPdc
' Date  : 10/10/2012
' Desc  : Prompts user for pdc forecast and copies it to this workbook
'---------------------------------------------------------------------------------------
Public Sub ImportPdc()
    Dim sPath As String
    sPath = Application.GetOpenFilename("pdc (*.csv), pdc.csv", Title:="Open the Pdc forecast")

    If sPath <> "False" Then
        Workbooks.Open sPath
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Pdc").Range("A1")
        ActiveWorkbook.Close
        Kill sPath
    Else
        MsgBox Prompt:="File import canceled." & vbCrLf & Err.Description, Title:="User Aborted Operation"
        Err.Raise 95
    End If
End Sub


'---------------------------------------------------------------------------------------
' Proc  : Sub ImportMfg
' Date  : 10/10/2012
' Desc  : Prompts user for mfg forecast and copies it to this workbook
'---------------------------------------------------------------------------------------
Public Sub ImportMfg()
    Dim sPath As String
    sPath = Application.GetOpenFilename("mfg (*.csv), mfg.csv", Title:="Open the Mfg forecast")

    If sPath <> "False" Then
        Workbooks.Open sPath
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Mfg").Range("A1")
        ActiveWorkbook.Close
        Kill sPath
    Else
        MsgBox Prompt:="File import canceled." & vbCrLf & Err.Description, Title:="User Aborted Operation"
        Err.Raise 95
    End If
End Sub


'---------------------------------------------------------------------------------------
' Proc  : Sub ImportGaps
' Date  : 10/10/2012
' Desc  : Opens gaps inventory file and copies it to this workbook
'---------------------------------------------------------------------------------------
Public Sub ImportGaps()
    Dim QOH As New GapsObject       'Quantity On Hand
    Dim QOR As New GapsObject       'Quantity On Reserve
    Dim QOO As New GapsObject       'Quantity On Order
    Dim QBO As New GapsObject       'Quantity On BO
    Dim WDC As New GapsObject       'Quantity at WDC
    Dim LC As New GapsObject        'Last Cost
    Dim UOM As New GapsObject       'Unit Of Measure
    Dim SUP As New GapsObject       'Supplier Number
    Dim sName As String             'Gaps Filename
    Dim i As Integer                'Counter
    Dim iRows As Long               'Number of rows on gaps
    Dim m_Gaps As New Collection    'Private GAPS collection
    Dim bFileFound As Boolean       'True if GAPS file is found
    Dim sPath As String             'Gaps File location
    
    sPath = "\\BR3615GAPS\GAPS\3615 GAPS DOWNLOAD\" & Format(Date, "yyyy") & "\"

    For i = 0 To 10
        sName = "3615 " & Format(Date - i, "yyyy-mm-dd") & ".csv"
        If FileExists(sPath & sName) Then
            bFileFound = True
            Exit For
        End If
    Next

    If bFileFound = False Then
        MsgBox Title:="Gaps file not found.", Prompt:="Please make sure you can connect to br3615gaps."
        Err.Raise 95
    Else
        Workbooks.Open sPath & sName
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Worksheets("Gaps").Range("A1")
        ActiveWorkbook.Close
        Worksheets("Gaps").Select
        Columns("D:D").ClearContents
        Range("D1").Value = "SIM"
        Range("D2").Formula = "=B2&C2"
        Range("D2").AutoFill Destination:=Range(Cells(2, 4), Cells(ActiveSheet.UsedRange.Rows.Count, 4))
        Range("D:D").Value = Range("D:D").Value
        Range("D:D").EntireColumn.AutoFit
        iRows = ActiveSheet.UsedRange.Rows.Count

        With ActiveSheet.UsedRange

            QOH.Address = Range(Cells(1, 4), Cells(iRows, 6)).Address(False, False)
            QOH.Column = 3
            QOH.Name = "On Hand"

            QOR.Address = Range(Cells(1, 4), Cells(iRows, 7)).Address(False, False)
            QOR.Column = 4
            QOR.Name = "Reserve"

            QOO.Address = Range(Cells(1, 4), Cells(iRows, 9)).Address(False, False)
            QOO.Column = 6
            QOO.Name = "On Order"

            QBO.Address = Range(Cells(1, 4), Cells(iRows, 8)).Address(False, False)
            QBO.Column = 5
            QBO.Name = "BO"

            WDC.Address = Range(Cells(1, 4), Cells(iRows, 36)).Address(False, False)
            WDC.Column = 33
            WDC.Name = "WDC"

            LC.Address = Range(Cells(1, 4), Cells(iRows, 31)).Address(False, False)
            LC.Column = 28
            LC.Name = "Last Cost"

            UOM.Address = Range(Cells(1, 4), Cells(iRows, 35)).Address(False, False)
            UOM.Column = 32
            UOM.Name = "UOM"

            SUP.Address = Range(Cells(1, 4), Cells(iRows, 38)).Address(False, False)
            SUP.Column = 35
            SUP.Name = "Supplier"

            m_Gaps.Add QOH
            m_Gaps.Add QOR
            m_Gaps.Add QOO
            m_Gaps.Add QBO
            m_Gaps.Add WDC
            m_Gaps.Add LC
            m_Gaps.Add UOM
            m_Gaps.Add SUP

            Set Gaps = m_Gaps
        End With
    End If
End Sub

Sub ImportMaster()
    Dim sPath As String
    Dim Wkbk As Workbook

    sPath = "\\br3615gaps\gaps\Billy Mac-Master Lists\Jacobsen-Textron Master File " & Format(Date, "yyyy") & ".xls"
    ThisWorkbook.Sheets("Master").Cells.Delete
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Workbooks.Open FileName:=sPath
    ActiveSheet.AutoFilterMode = False
    Set Wkbk = ActiveWorkbook
    ActiveSheet.UsedRange.Copy
    ThisWorkbook.Activate
    Sheets("Master").Range("A1").PasteSpecial Paste:=xlPasteValues, _
                                              Operation:=xlNone, _
                                              SkipBlanks:=False, _
                                              Transpose:=False
    Application.CutCopyMode = False
    Wkbk.Close
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Sheets("Macro").Select
End Sub




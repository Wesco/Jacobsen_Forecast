Attribute VB_Name = "Exports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub ExportFcst
' Date  : 10/17/2012
' Desc  : Copies the forecast sheet to a new workbook and saves it to the network
'---------------------------------------------------------------------------------------
Sub ExportFcst()
    Dim PrevDispAlert As Boolean
    Dim sPath As String
    Dim sName As String

    PrevDispAlert = Application.DisplayAlerts
    sName = "Jacobsen Slink " & Format(Date, "m-dd-yy") & ".xlsx"
    sPath = "\\br3615gaps\gaps\Jacobsen-Textron\" & Format(Date, "yyyy") & " Alerts\"

    If FolderExists(sPath) = False Then
        RecMkDir sPath
    End If

    Sheets("Forecast").Copy
    Sheets.Add After:=Sheets(Sheets.Count), Count:=2
    Sheets("Sheet2").Name = "Order"
    Sheets("Sheet3").Name = "Expedite"
    Sheets("Forecast").Select
    Range("A1").Select

    Application.DisplayAlerts = True
    On Error Resume Next
    ActiveWorkbook.SaveAs FileName:=sPath & sName, FileFormat:=xlOpenXMLWorkbook
    On Error GoTo 0

    Application.DisplayAlerts = False
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert
End Sub

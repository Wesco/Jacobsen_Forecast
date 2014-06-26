Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : Sub ImportForecast
' Date : 6/25/2014
' Desc : Prompts user for forecast, imports it, then deletes the original file
'---------------------------------------------------------------------------------------
Sub ImportForecast(FileFilter As String, Title As String, Destination As Range)
    Dim Path As String

    Path = Application.GetOpenFilename(FileFilter, Title:=Title)

    If Path <> "False" Then
        Workbooks.Open Path
        ActiveSheet.UsedRange.Copy Destination:=Destination
        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close

        DeleteFile Path
    Else
        Err.Raise Errors.USER_INTERRUPT, "ImportForecast", "User Aborted Import"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportMaster
' Date : 6/25/2014
' Desc : Imports the Jacobsen master list
'---------------------------------------------------------------------------------------
Sub ImportMaster()
    Dim Path As String
    Dim File As String
    Dim MasterWkbk As Workbook
    Dim PrevDispAlerts As Boolean
    Dim PrevUpdateLnks As Boolean

    Path = "\\br3615gaps\gaps\Billy Mac-Master Lists\"
    File = "Jacobsen-Textron Master File " & Format(Date, "yyyy") & ".xls"
    PrevDispAlerts = Application.DisplayAlerts
    PrevUpdateLnks = Application.AskToUpdateLinks

    'Disable prompts
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False

    If FileExists(Path & File) Then
        Workbooks.Open FileName:=Path & File
        Set MasterWkbk = ActiveWorkbook

        Sheets("Master").Select

        'Unhide data and remove filters
        ActiveSheet.AutoFilterMode = False
        ActiveSheet.Columns.Hidden = False
        ActiveSheet.Rows.Hidden = False

        ActiveSheet.UsedRange.Copy

        ThisWorkbook.Activate
        Sheets("Master").Select
        Range("A1").PasteSpecial Paste:=xlPasteValues, _
                                 Operation:=xlNone, _
                                 SkipBlanks:=False, _
                                 Transpose:=False
        Application.CutCopyMode = False
        MasterWkbk.Close

        Application.DisplayAlerts = PrevDispAlerts
        Application.AskToUpdateLinks = PrevUpdateLnks
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "ImportMaster", File & " could not be found."
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportKitBOM
' Date : 6/26/2014
' Desc : Imports kit bill of materials
'---------------------------------------------------------------------------------------
Sub ImportKitBOM()
    Dim Path As String
    Dim File As String
    Dim ColHeaders As Variant
    Dim TotalCols As Integer
    Dim i As Integer

    Path = "\\br3615gaps\gaps\3615 Kit BOM\"
    File = "Kit BOM " & Format(Date, "yyyy") & ".csv"
    ColHeaders = Array("Transaction Code    ", _
                       "* Branch *          ", _
                       "SIM                 ", _
                       "Line No.            ", _
                       "Record Type         ", _
                       "Comp SIM            ", _
                       "Quantity            ", _
                       "Cat No.             ", _
                       "Description         ", _
                       "UOM                 ", _
                       "Supp. No.           ", _
                       "Supp. Name          ", _
                       "Product             ", _
                       "Unit Cost           ", _
                       "BO                  ", _
                       "GST                 ", _
                       "Taxable             ", _
                       "BOM Type", _
                       "Long Description")

    'Import the Kit BOM
    If FileExists(Path & File) Then
        Workbooks.Open Path & File
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Kit").Range("A1")
        ActiveWorkbook.Saved = True
        ActiveWorkbook.Close
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "ImportKitBOM", "Kit BOM not found"
    End If

    Sheets("Kit").Select
    Columns("T:T").Delete 'Remove column that contains nothing but spaces
    TotalCols = Rows(2).Columns(Columns.Count).End(xlToLeft).Column

    'Make sure the correct number of columns exist
    If Not TotalCols = UBound(ColHeaders) + 1 Then
        Err.Raise CustErr.MODIFIEDREP, "ImportKitBOM", "Kit BOM has been modified"
    End If

    'Make sure none of the columns have changed
    For i = 1 To TotalCols
        If Not Cells(2, i).Value = ColHeaders(i - 1) Then
            Err.Raise CustErr.MODIFIEDREP, "ImportKitBOM", "Kit BOM has been modified"
        End If
    Next
End Sub

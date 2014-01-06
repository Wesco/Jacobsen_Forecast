Attribute VB_Name = "All_Helper_Functions"
Option Explicit
'Pauses for x# of milliseconds
'Used for email function to prevent
'All emails from being sent at once
'Example: "Sleep 1500" will pause for 1.5 seconds
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------------------------------------------------
' Proc  : Function FileExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a file exists
' Ex    : FileExists "C:\autoexec.bat"
'---------------------------------------------------------------------------------------
Public Function FileExists(ByVal sPath As String) As Boolean
    'Remove trailing backslash
    If InStr(Len(sPath), sPath, "\") > 0 Then sPath = Left(sPath, Len(sPath) - 1)
    'Check to see if the directory exists and return true/false
    If Dir(sPath, vbDirectory) <> "" Then FileExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Function FolderExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a folder exists
' Ex    : FolderExists "C:\Program Files\"
'---------------------------------------------------------------------------------------
Public Function FolderExists(ByVal sPath As String) As Boolean
    ' Add trailing backslash
    If InStr(Len(sPath), sPath, "\") = 0 Then sPath = sPath & "\"
    ' If the folder exists return true
    If Dir(sPath, vbDirectory) <> "" Then FolderExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Sub RecMkDir
' Date  : 10/10/2012
' Desc  : Creates an entire directory tree
' Ex    : RecMkDir "C:\Dir1\Dir2\Dir3\"
'---------------------------------------------------------------------------------------
Sub RecMkDir(ByVal sPath As String)
    Dim sDirArray() As String   'Folder names
    Dim sDrive As String        'Base drive
    Dim sNewPath As String      'Path builder
    Dim LoopStart As Long       'Loop start number
    Dim i As Long               'Counter

    'Add trailing slash
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    'Split at each \
    If Left(sPath, 2) <> "\\" Then
        sDirArray = Split(sPath, "\")
        sDrive = sDirArray(0) & "\"
    Else
        sDirArray = Split(sPath, "\")
        sDrive = "\\" & sDirArray(2) & "\"
    End If

    'Determine where in the array to start the loop
    If sDrive = "\\" & sDirArray(2) & "\" Then
        LoopStart = 3
    Else
        LoopStart = 1
    End If

    'Loop through each directory
    For i = LoopStart To UBound(sDirArray) - 1
        If Len(sNewPath) = 0 Then
            sNewPath = sDrive & sNewPath & sDirArray(i) & "\"
        Else
            sNewPath = sNewPath & sDirArray(i) & "\"
        End If

        If Not FolderExists(sNewPath) And Len(sNewPath) > 3 Then
            MkDir sNewPath
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc  : Function Email
' Date  : 10/11/2012
' Type  : Variant
' Desc  : Sends an email
' Ex    : Email SendTo:=email@example.com, Subject:="example email", Body:="Email Body"
'---------------------------------------------------------------------------------------
Function Email(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, Optional Attachment As String)
    Dim Mail_Object, Mail_Single As Variant
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
    With Mail_Single
        .Subject = Subject
        If Attachment <> "" Then
            'Attachment must contain file path
            .Attachments.Add Attachment
        End If
        .To = SendTo
        .CC = CC
        .BCC = BCC
        .HTMLbody = Body
        .Send
    End With
    Sleep 1500
End Function

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    
    AddReferences
    codeFolder = CombinePaths(GetWorkbookPath, "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5))
    
    On Error Resume Next
    RecMkDir codeFolder
    On Error GoTo 0

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                FileName = CombinePaths(codeFolder, comp.Name & ".bas")
                DeleteFile FileName
                comp.Export FileName
            Case 2
                FileName = CombinePaths(codeFolder, comp.Name & ".cls")
                DeleteFile FileName
                comp.Export FileName
            Case 3
                FileName = CombinePaths(codeFolder, comp.Name & ".frm")
                DeleteFile FileName
                comp.Export FileName
        End Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : DeleteFile
' Date : 3/19/2013
' Desc : Deletes a file
'---------------------------------------------------------------------------------------
Sub DeleteFile(FileName As String)
    On Error Resume Next
    Kill FileName
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : CombinePaths
' Date : 3/19/2013
' Desc : Adds folders onto the end of a file path
'---------------------------------------------------------------------------------------
Function CombinePaths(ByVal Path1 As String, ByVal Path2 As String) As String
    If Not EndsWith(Path1, "\") Then
        Path1 = Path1 & "\"
    End If
    CombinePaths = Path1 & Path2
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds references required for helper functions
'---------------------------------------------------------------------------------------
Sub AddReferences()
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean

    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = "{0002E157-0000-0000-C000-000000000046}" And Ref.Major = 5 And Ref.Minor = 3 Then
            Result = True
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    End If
End Sub

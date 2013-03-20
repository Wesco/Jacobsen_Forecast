VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GapsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_iColumn As Integer
Private m_sAddress As String
Private m_sName As String

'---------------------------------------------------------------------------------------
' Proc  : Property Column
' Date  : 10/12/2012
' Type  : Integer
' Desc  : Defines column retrieval
'---------------------------------------------------------------------------------------
Public Property Get Column() As Integer
    Column = m_iColumn
End Property

'---------------------------------------------------------------------------------------
' Proc  : Property Column
' Date  : 10/12/2012
' Type  : Integer
' Desc  : Defines column assignment
'---------------------------------------------------------------------------------------
Public Property Let Column(ByVal iColumn As Integer)
    m_iColumn = iColumn
End Property

'---------------------------------------------------------------------------------------
' Proc  : Property Address
' Date  : 10/12/2012
' Type  : String
' Desc  : Defines address retrieval
'---------------------------------------------------------------------------------------
Public Property Get Address() As String
    Address = m_sAddress
End Property

'---------------------------------------------------------------------------------------
' Proc  : Property Address
' Date  : 10/23/2012
' Type  : String
' Desc  : Defines address assignment
'---------------------------------------------------------------------------------------
Public Property Let Address(ByVal sAddress As String)
    m_sAddress = sAddress
End Property

'---------------------------------------------------------------------------------------
' Proc  : Property Name
' Date  : 10/12/2012
' Type  : String
' Desc  : Defines name retrieval
'---------------------------------------------------------------------------------------
Public Property Get Name() As String
    Name = m_sName
End Property

'---------------------------------------------------------------------------------------
' Proc  : Property Name
' Date  : 10/12/2012
' Type  : String
' Desc  : Defines name assignment
'---------------------------------------------------------------------------------------
Public Property Let Name(ByVal sName As String)
    m_sName = sName
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_RootElement As CXmlElement

Public Sub AppendNode(ByVal Item As CXmlElement)
Call m_RootElement.AppendNode(Item)
End Sub

' This routine does not check for the existance of a Document Prolog (I just don't need it yet)
' Also It only accepts ANSI Characters (Hope you speak english)
Public Sub LoadData(Data As Variant)
Dim bData() As Byte
On Error Resume Next

' Coerce Data into a Byte Array
Select Case VarType(Data)
Case vbString
'If Len(Data & vbNullString) = 0 Then _
  '    Err.Raise 13 & vbObjectError, "vbXml", "Unable to Parse Empty Data" ' Type Mismatch
    bData() = StrConv(Data & vbNullString, vbFromUnicode)

Case (vbArray And vbByte)
    bData() = Data

Case Else
'Err.Raise 13, "vbXml" ' Type Mismatch

End Select

Call m_RootElement.Parse(bData())

On Error GoTo 0
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_RootElement.NewEnum
End Property

Public Property Get Node(ByVal Index As Long) As CXmlElement
Set Node = m_RootElement.Node(Index)
End Property

Public Sub RemoveNode(ByVal Index As Long)
Call m_RootElement.RemoveNode(Index)
End Sub


Public Property Get Root() As CXmlElement
Set Root = m_RootElement
End Property

Public Function Serialize() As String
Serialize = m_RootElement.Serialize
End Function

Private Sub Class_Initialize()
Set m_RootElement = New CXmlElement
End Sub


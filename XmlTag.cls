VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum TagStyles
    STag        ' Start Tag
    ETag        ' End Tag
    EmptyTag    ' Empty Element Tag
End Enum

Private m_eStyle As TagStyles
Private m_colAttributes As Collection
Private m_sType As String

Friend Function Serialize(Optional eType As TagStyles = STag) As String
Dim oAttrib As CXmlAttribute

If (eType = STag) Or (eType = EmptyTag) Then
    Serialize = "<" & m_sType
    For Each oAttrib In m_colAttributes
        Serialize = Serialize & " " & oAttrib.Serialize
    Next
    Serialize = Serialize & IIf(eType = EmptyTag, "/>", ">")
Else
    Serialize = "</" & m_sType & ">"
End If

End Function

Public Sub AppendAttribute(ByVal Attrib As CXmlAttribute)
On Error GoTo Err_Trap

Call m_colAttributes.Add(Attrib, Attrib.KeyWord)
Exit Sub

Err_Trap:
Select Case Err.Number
' Exceptions Raised:
Case 457
'Duplicate Attribute in Tag [add attribute with same key]
    Call Err.Raise(vbObjectError Or Err.Number, _
                   "AppendAttribute", _
                   "Attempt to Append Duplicate Attribute")

' Exceptions Forwarded:
Case 91
'Object Variable not Set
    Err.Raise Err.Number

Case Else
' ToDo . . . Log All Other Errors
End Select
End Sub
Public Property Get AttributeCount() As Long
AttributeCount = m_colAttributes.count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_colAttributes.[_NewEnum]
End Property

Public Sub RemoveAttribute(ByVal Index As Variant)
' Forward All Exceptions
Call m_colAttributes.Remove(Index)
End Sub

Public Property Get TagAttribute(ByVal Index As Variant) As CXmlAttribute
' Forward All Exceptions
Set TagAttribute = m_colAttributes(Index)
End Property

' Start is the position of the Tag Start Element
' Upon Return it should be positioned on the end tag on the Tag Parsed
Friend Sub Parse(Data() As Byte, Start As Long)
Dim oAttrib As CXmlAttribute

On Error GoTo Err_Trap

Start = Start + 1
Do
    Select Case Data(Start)
    Case ascTagTerm
        If Data(Start - 1) = ascTagBegin Then
            m_eStyle = ETag
        ElseIf Data(Start + 1) = ascTagEnd Then
            m_eStyle = EmptyTag
        End If

' Delimiting Characters
    Case ascSpace, ascTab, ascLf, ascCr
' Ignored

' Valid Name Starting Characetrs
    Case ascUpperFirst To ascUpperLast, _
         ascLowerFirst To ascLowerLast, _
         ascUnderScore, ascColon

        If m_sType <> vbNullString Then
            Set oAttrib = New CXmlAttribute
            Call oAttrib.Parse(Data(), Start)
            Call AppendAttribute(oAttrib)
            Start = Start - 1
        Else
            m_sType = ParseName(Data(), Start)
            Start = Start - 1
        End If

    Case ascTagEnd
        Exit Do

    Case Else
' Error Condition

    End Select
    Start = Start + 1
'   DoEvents
Loop    ' Until we reach the End of the Tag
Exit Sub

Err_Trap:
'    Select Case Err.Number
'        ' Exceptions Raised:
'        Case 9
'            'Unexpected End of Data [array index out of bounds]
'            Err.Raise vbObjectError Or Err.Number, "Parse", "Unexpected end of data", vbNullString, 0
'
'        ' Exceptions Forwarded:
'        Case vbObjectError Or 457, vbObjectError Or 9
'            'CXmlAttribute.Parse Exceptions
'            'ParseName Exceptions
'            'AppendAttribute Exceptions
'            Err.Raise Err.Number
'
'        Case Else
'        ' ToDo . . . Log All Other Errors
'
'    End Select
End Sub

Public Property Get TagStyle() As TagStyles
TagStyle = m_eStyle
End Property

Public Property Let TagType(ByVal RHS As String)
' ToDo . . . XML Validate RHS Data
m_sType = RHS
End Property

Public Property Get TagType() As String
TagType = m_sType
End Property

Private Sub Class_Initialize()
Set m_colAttributes = New Collection
End Sub

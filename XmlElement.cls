VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXmlElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private m_colChildren As Collection     ' Child Elements of this Element
Private m_oTag As CXmlTag               ' The tag defining this element, along
' with all it's associated properties
Public Body As String                   ' The CharData section of the Element

Public Sub AppendAtribute(ByVal Attrib As CXmlAttribute)
' Forward All Exceptions
If m_oTag Is Nothing Then _
   Set m_oTag = New CXmlTag
Call m_oTag.AppendAttribute(Attrib)
End Sub

Public Sub AppendNode(ByVal Element As CXmlElement, Optional ByVal Before As Long)
'Forward All Exceptions
If Before = 0 Then
    Call m_colChildren.Add(Element)
Else
    Call m_colChildren.Add(Element, , Before)
End If
End Sub

Public Property Get AttributeCount() As Long
If m_oTag Is Nothing Then Exit Property
AttributeCount = m_oTag.AttributeCount
End Property

Public Property Get ElementAttribute(ByVal Index As Variant) As CXmlAttribute
' Forward All Excpetions
If m_oTag Is Nothing Then Exit Property
Set ElementAttribute = m_oTag.TagAttribute(Index)
End Property

Public Property Let Name(ByVal RHS As String)
' Todo . . . XML Validate RHS Data
If m_oTag Is Nothing Then _
   Set m_oTag = New CXmlTag
m_oTag.TagType = RHS
End Property

Public Property Get Name() As String
If m_oTag Is Nothing Then Exit Property
Name = m_oTag.TagType
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = m_colChildren.[_NewEnum]
End Property

Public Property Get Node(ByVal Index As Long) As CXmlElement
' Forward All Exceptions
Set Node = m_colChildren(Index)
End Property

Friend Sub Parse(Data() As Byte, Optional Start As Long)
Dim lPos As Long, bStart As Boolean
Dim oTag As CXmlTag
Dim oChild As CXmlElement

On Error GoTo Err_Trap

If Start < 0 Then Exit Sub
If Start = 0 Then Start = LBound(Data)

' Ensure we clear out the Current Data (if any)
Set m_colChildren = New Collection

' Start Parsing the Element Body
Do
    Select Case Data(Start)
' The first things we check for are the Start Tag Character '<'
' and the Amperstand '&'. These are the only characters not allowed
' in a character data section of an XML element. Also the sequence ']]>'
' is not allowed as it marks the end of a '<[CDATA[' markup section.
    Case ascTagBegin
' I am going to check if the last characters of the body are white
' space and trim them if they are.
        If bStart Then
            Do
                lPos = Len(Body)
                Select Case Asc(right$(Body, 1))
                Case ascSpace, ascTab, ascCr, ascLf
                    Body = left$(Body, lPos - 1)
                End Select
                If Len(Body) = 0 Then Exit Do
            Loop Until lPos = Len(Body)    'until we don't change the expression
        Else
            bStart = True
        End If

' Save the current seek value
        lPos = Start
' Parse the Tag By Itself
        Set oTag = New CXmlTag
        Call oTag.Parse(Data(), Start)

' Start by getting the sTag for this element
        If m_oTag Is Nothing Then
            Set m_oTag = oTag

' Technically, the First Tag could be an Empty Element Tag
            If m_oTag.TagStyle = EmptyTag Then
                Exit Sub
            ElseIf m_oTag.TagStyle = ETag Then
                Set m_oTag = Nothing
' I suppose it would be possible for an End Tag to exist without
' an associated Start Tag, but it wouldn't be well formed if it was.
            End If

' Check to see what Kind of Tag it is
        ElseIf oTag.TagStyle = ETag Then     ' We found our End Tag . . . Hurray!!!!
            Exit Do

' New Tag . . . Time to Recurse
        Else
' Reset the Seek Value
            Start = lPos
' Create a Child Element
            Set oChild = New CXmlElement
' Tell it to Parse
            Call oChild.Parse(Data, Start)
' Append it to the Collection
            Call AppendNode(oChild)
        End If


' Escape Sequence
    Case ascAmper
        If bStart Then Body = Body & DecodeEscape(Data(), Start)

    Case Else
' Everything Else is CharData
        If bStart Then Body = Body & Chr(Data(Start))

    End Select
    Start = Start + 1
'DoEvents
Loop
Exit Sub

Err_Trap:
Select Case Err.Number
' Exceptions Raised:
Case 9
'Unexpected End of Data [array index out of bounds]
    Err.Raise vbObjectError Or Err.Number, "Parse", "Missing closing tag </" & m_oTag.TagType & ">", vbNullString, 0

' Exceptions Forwarded:
Case vbObjectError Or 9, vbObjectError Or 457
'Tag Parse Exceptions
'Element Parse Exceptions
    Err.Raise Err.Number

Case Else
' ToDo . . . Log all other Errors
End Select
End Sub
Public Sub RemoveAttribute(ByVal Index As Variant)
If m_oTag Is Nothing Then Exit Sub
Call m_oTag.RemoveAttribute(Index)
End Sub

Public Sub RemoveNode(ByVal Index As Long)
Call m_colChildren.Remove(Index)
End Sub

Public Function Serialize() As String
Dim oElement As CXmlElement
Dim lPos As Long

If m_oTag Is Nothing Then Exit Function

Serialize = m_oTag.Serialize
For Each oElement In m_colChildren
    Serialize = Serialize & oElement.Serialize
Next
' ToDo ... XML Encode Body Data
lPos = InStr(Body, "]]>")
If lPos > 0 Then
    Body = left$(Body, lPos - 1) & "]]&gt;" & right$(Body, Len(Body) - (lPos - 1) - 3)
End If

Serialize = Serialize & "<![CDATA[" & Body & "]]>" & m_oTag.Serialize(ETag) & vbCrLf
End Function

Private Sub Class_Initialize()
Set m_colChildren = New Collection
End Sub

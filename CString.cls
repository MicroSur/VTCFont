VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CStringBuilder Class
' ====================
' String Builder class for VB6 to vastly speed up the concatination
' of large strings

' This will be a 1-based array
Dim m_bytes() As Byte

Public Sub concat(ByVal s As String)
' 19/01/09 - Concatinate the specified string
'            to the string held in this object
Dim lOldUbound As Long
Dim l As Long
On Error GoTo errError1

lOldUbound = UBound(m_bytes)
ReDim Preserve m_bytes(0 To lOldUbound + Len(s))
For l = lOldUbound + 1 To lOldUbound + Len(s)
    m_bytes(l) = Asc(Mid$(s, l - lOldUbound, 1))
Next

Exit Sub
errError1:
MsgBox Err.Number, , Err.Description
End Sub

Public Property Get length() As Long
' 19/01/01 - Return the length of the string
'            currently held in this object
On Error GoTo errError1

' The array is 1-based so will match the string length
length = UBound(m_bytes)
Exit Property

errError1:
MsgBox Err.Number, , Err.Description
End Property

Public Property Get Text() As String
' 19/01/01 - Return the string held in this
'            object
On Error GoTo errError1

If Me.length = 0 Then
    Text = vbNullString
Else
    Text = right$(StrConv(m_bytes, vbUnicode), Me.length)
End If

Exit Property

errError1:
MsgBox Err.Number, , Err.Description
End Property
Public Property Get Char(Ind As Long) As String
' Return one char from ind position
On Error Resume Next 'GoTo errError1

Char = Chr(m_bytes(Ind))
Exit Property

errError1:
MsgBox Err.Number, , Err.Description
End Property

Private Sub Class_Initialize()
' 19/01/01 - Prepare the array for first use
On Error GoTo errError1

Call reset
Exit Sub
errError1:
MsgBox Err.Number, , Err.Description
End Sub

Public Sub reset()
' 19/01/01 - Reset the array ready to be used again
On Error GoTo errError1

ReDim m_bytes(0) As Byte
Exit Sub

errError1:
MsgBox Err.Number, , Err.Description
End Sub

Private Sub Class_Terminate()
' 19/01/01 - Clear up memory used by byte array
'            Not really necessary
On Error GoTo errError1

'Call reset
Exit Sub

errError1:
MsgBox Err.Number, , Err.Description
End Sub

Attribute VB_Name = "INI"
Option Explicit
Private Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
' This first line is the declaration from win32api.txt
Private Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Private Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
'Private Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)
'Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Const maxLen = 4096
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal olestr As Long, ByVal bLen As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal NBytes As Long)

Public Function AllocString_ADV(ByVal lSize As Long) As String
RtlMoveMemory ByVal VarPtr(AllocString_ADV), _
              SysAllocStringByteLen(0&, lSize + lSize), 4&
End Function

'WriteKey "section", "param", "value", iniFileName
'value = VBGetPrivateProfileString("section", "param", iniFileName)

Function VBGetPrivateProfileString(section As String, Key As String, file As String, Optional ByRef SameStr As String) As String

Dim KeyValue As String
Dim characters As Long

'KeyValue = String$(1024, 0)
KeyValue = AllocString_ADV(maxLen + 1)

characters = GetPrivateProfileStringByKeyName(section, Key, "=", KeyValue, maxLen, file)
'Debug.Print Asc(KeyValue)
If (AscW(KeyValue) = 0) Or (left$(KeyValue, 1) = "=") Then     ' не нашли
    VBGetPrivateProfileString = SameStr
    Exit Function
End If

If characters > 0 Then
    KeyValue = left$(KeyValue, characters)
End If

VBGetPrivateProfileString = KeyValue

End Function


Public Function GetSectionNames(filename As String, SectionNames() As String) As Integer
'GetSectionNames Return Number of Section in file
'SectionNames return all section names

Dim characters As Long
Dim SectionList As String
Dim ArrSection() As String
Dim i As Integer
Dim NullOffset As Integer

SectionList = String(maxLen + 1, 0)  '128

' Retrieve the list of keys in the section
characters = GetPrivateProfileStringSections(0, 0, "", SectionList, maxLen, filename)    '127

' Load sections into Arrey
i = 0
Do
    NullOffset = InStr(SectionList, vbNullChar)
    If NullOffset > 1 Then
        ReDim Preserve ArrSection(i)
        ArrSection(i) = Mid$(SectionList, 1, NullOffset - 1)
        SectionList = Mid$(SectionList, NullOffset + 1)
        i = i + 1
    End If
Loop While NullOffset > 1
GetSectionNames = i - 1
SectionNames = ArrSection

End Function

Public Function GetKeyNames(SectionName As String, filename As String, KeyNames As Variant) As Integer
'GetKeyNames Return Number of key in section или -1
'KeyNames Return list of keyNames in section массив

Dim characters As Long
Dim KeyList As String
Dim ArrName() As String
Dim NullOffset As Integer
Dim i As Integer

KeyList = String(maxLen + 1, 0)
' Retrieve the list of keys in the section

characters = GetPrivateProfileStringKeys(SectionName, 0, "", KeyList, maxLen, filename)

' Load Keys into Arrey
i = 0
Do
    NullOffset = InStr(KeyList, vbNullChar)
    If NullOffset > 1 Then
        ReDim Preserve ArrName(i)
        ArrName(i) = Mid$(KeyList, 1, NullOffset - 1)
        KeyList = Mid$(KeyList, NullOffset + 1)
        i = i + 1
    End If
Loop While NullOffset > 1
GetKeyNames = i - 1
KeyNames = ArrName

End Function

Public Function DeleteKey(KeyName As String, SectionName As String, filename As String) As Long
'Return 0 if Deletion not sucsesful или нечего удалять (не нашел)
' Delete the selected key
DeleteKey = WritePrivateProfileStringToDeleteKey(SectionName, KeyName, 0, filename)
End Function

Public Function WriteKey(SectionName As String, KeyName As String, KeyValue As String, filename As String) As Long
If Len(KeyValue) = 0 Then KeyValue = vbNullString
WriteKey = WritePrivateProfileStringByKeyName(SectionName, KeyName, KeyValue, filename)
End Function

'Public Function WriteSection(SectionName As String, filename As String) As Long
'    WriteSection = WritePrivateProfileSection(SectionName, "", filename)
'End Function
'
'Public Function DeleteSection(SectionName, filename) As Long
'    DeleteSection = WritePrivateProfileStringToDeleteSection(SectionName, 0&, 0&, filename)
'End Function





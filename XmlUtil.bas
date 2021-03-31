Attribute VB_Name = "XmlUtil"
Option Explicit

' White Space Characters
Public Const ascSpace As Byte = 32
Public Const ascTab As Byte = 9
Public Const ascCr As Byte = 13
Public Const ascLf As Byte = 10

' Tag Characters
Public Const ascTagBegin As Byte = 60
Public Const ascTagEnd As Byte = 62
Public Const ascTagTerm As Byte = 47
Public Const ascAmper As Byte = 38
Public Const ascSemiColon As String = 59

' Letter Characters (Begining And Ending for Simplicity)
Public Const ascLowerFirst As Byte = 97
Public Const ascLowerLast As Byte = 122
Public Const ascUpperFirst As Byte = 65
Public Const ascUpperLast As Byte = 90
Public Const ascUnderScore As Byte = 95
Public Const ascColon As Byte = 58

' Digit Characters
Public Const ascNumFirst As Byte = 48
Public Const ascNumLast As Byte = 57

' Other Characters
Public Const ascEquals As Byte = 61
Public Const ascApos As Byte = 39      ' Single Quote
Public Const ascQuote As Byte = 34     ' Double Quote
Public Const ascPound As Byte = 35

' Special Strings
Public Const strAmp As String = "amp"
Public Const strLessThan As String = "lt"
Public Const strMoreThan As String = "gt"
Public Const strApostrophe As String = "apos"
Public Const strQuote As String = "quot"

Public Function DecodeEscape(Data() As Byte, Start As Long) As String
On Error GoTo Err_Trap

Do      ' Until we find a semicolon
    Start = Start + 1
    If Data(Start) = ascSemiColon Then _
       Exit Do
    DecodeEscape = DecodeEscape & Chr(Data(Start))
Loop

Select Case DecodeEscape
Case strAmp
    DecodeEscape = "&"

Case strApostrophe
    DecodeEscape = "'"

Case strLessThan
    DecodeEscape = "<"

Case strMoreThan
    DecodeEscape = ">"

Case strQuote
    DecodeEscape = """"

Case Else
    If Data(Start - Len(DecodeEscape)) = ascPound Then
' Numeric Escape Sequence
        If Data(Start - (Len(DecodeEscape) + 1)) = Chr("x") Then
' Hexadecimal
            DecodeEscape = right$(DecodeEscape, Len(DecodeEscape) - 2)
        Else
' Decimal
            DecodeEscape = right$(DecodeEscape, Len(DecodeEscape) - 1)
        End If
    Else
' Custom Entity Reference
' Not Currently Supported
        DecodeEscape = vbNullString
    End If
End Select
Exit Function

Err_Trap:
Select Case Error
' Exceptions Raised:
Case 9
'Unexpected End of Data [array index out of bounds]
    Err.Raise vbObjectError Or Err.Number, "DecodeEscape", "Unexpected end of data", vbNullString, 0

Case Else
' Todo . . . Log all other errors

End Select

End Function

' Parses a value contained within quotes
' Start identifies the begining quote and
' will identify the closing quote on exit
Public Function ParseValue(Data() As Byte, Start As Long) As String
Dim bEnd As Boolean
Dim QuoteChar As Byte

On Error GoTo Err_Trap

QuoteChar = Data(Start)

Do
    Select Case Data(Start)
    Case QuoteChar
        bEnd = Not bEnd
        If Not bEnd Then Exit Do

    Case Is <> ascTagBegin, Is <> ascAmper
        ParseValue = ParseValue & Chr(Data(Start))

    Case ascAmper
        ParseValue = ParseValue & DecodeEscape(Data(), Start)

    Case Else
' The only other case is the Begin Tag which is invalid in this context

    End Select
    Start = Start + 1
Loop
Exit Function

Err_Trap:
Select Case Error
' Exceptions Raised:
Case 9
'Unexpected End of Data [array index out of bounds]
    Err.Raise vbObjectError Or Err.Number, "ParseValue", "Unexpected end of data", vbNullString, 0

' Exceptions Forwarded:
Case vbObjectError Or 9
' DecodeEscape Exceptions
    Err.Raise Err.Number

Case Else
' Todo . . . Log all other errors

End Select

End Function

' Start Identifies the First Character to Check
' Upon completion, Start should point to the first
' non-delimitng character after the Name Value is read
Public Function ParseName(Data() As Byte, Start As Long) As String
Dim bEnd As Boolean

On Error GoTo Err_Trap

Do
    Select Case Data(Start)
' Delimitng Characters
    Case ascSpace, ascTab, ascCr, ascLf, ascEquals, ascSemiColon
        bEnd = True

    Case ascTagEnd, ascApos, ascQuote
        Exit Do

' Letter Characetrs
    Case ascUpperFirst To ascUpperLast, _
         ascLowerFirst To ascLowerLast, _
         ascUnderScore, ascColon, _
         ascNumFirst To ascNumLast, &H41 To &H5A, _
         &H61 To &H7A, &HC0 To &HD6, _
         &HD8 To &HF6, &HF8 To &HFF, &HB7

        If bEnd Then
            Exit Do
        Else
            ParseName = ParseName & Chr(Data(Start))
        End If

    Case Else
' Error . . . Normally not too many charater
' types can be used for the Name Identifier

    End Select
    Start = Start + 1
Loop
Exit Function

Err_Trap:
Select Case Error
' Exceptions Raised:
Case 9
'Unexpected End of Data [array index out of bounds]
    Err.Raise vbObjectError Or Err.Number, "ParseName", "Unexpected end of data", vbNullString, 0

Case Else
' Todo . . . Log all other errors

End Select

End Function



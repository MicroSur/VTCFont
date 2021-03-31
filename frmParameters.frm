VERSION 5.00
Begin VB.Form frmParameters 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   4260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3420
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtValidator 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   14
      Top             =   5940
      Width           =   1575
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   2580
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1740
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdModifyPatch 
      Caption         =   "Modify"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   5940
      Width           =   3015
   End
   Begin VB.TextBox txtParam 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   9
      Left            =   1140
      TabIndex        =   22
      Top             =   4320
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   8
      Left            =   1140
      TabIndex        =   20
      Top             =   3900
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   7
      Left            =   1140
      TabIndex        =   18
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   6
      Left            =   1140
      TabIndex        =   16
      Top             =   3060
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   5
      Left            =   1140
      TabIndex        =   13
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   4
      Left            =   1140
      TabIndex        =   12
      Top             =   2220
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   3
      Left            =   1140
      TabIndex        =   11
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   10
      Top             =   1380
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   1
      Left            =   1140
      TabIndex        =   9
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label lblParam 
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   8
      Top             =   540
      Width           =   5175
   End
   Begin VB.Label lblPatchName 
      AutoSize        =   -1  'True
      Caption         =   "patch name"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private arrAllParamAddrs() As String
Private arrAllParamBytes() As String
Private Function ValidateInput() As Boolean
Dim i As Integer
Dim muls As Integer
Dim ds As String

On Error GoTo frmErr

ds = left$(DecimalSeparator, 1)
For i = 0 To iParamNum - 1
    'If InStr(txtParam(i).Text, ",") Then
        txtParam(i).Text = Replace(txtParam(i).Text, ",", ds)
        txtParam(i).Text = Replace(txtParam(i).Text, ".", ds)
    'End If
Next i

ValidateInput = True
For i = 1 To iParamNum

    If arrMulByte(i) = 0 Then
        muls = 1
    Else
        muls = 10 ^ arrMulByte(i)
    End If


    txtParam(i - 1).ForeColor = &H80000008
    If Not IsNumeric(txtParam(i - 1).Text) Then
        ValidateInput = False
        txtParam(i - 1).ForeColor = vbRed
        MsgBoxEx ArrMsg(44), , , CenterOwner, vbCritical
        Exit Function
    Else
        txtParam(i - 1).Text = Abs(txtParam(i - 1).Text)
    End If

    If arrByteCount(i) = 1 Then
    
        If (CDbl(txtParam(i - 1).Text) * muls) > 255 Then
            ValidateInput = False
            txtParam(i - 1).ForeColor = vbRed
            MsgBoxEx ArrMsg(43), , , CenterOwner, vbCritical
            Exit Function
        End If
        
    ElseIf arrByteCount(i) = 2 Then
    
        If (CDbl(txtParam(i - 1).Text) * muls) > 65535 Then
            ValidateInput = False
            txtParam(i - 1).ForeColor = vbRed
            MsgBoxEx ArrMsg(43), , , CenterOwner, vbCritical
            Exit Function
        End If
        
    End If
Next i

'''
Exit Function
frmErr:
ValidateInput = False
MsgBox Err.Description & ": Param_ValidateInput"
End Function
Private Sub Changes2Arrays()
Dim i As Integer
Dim Tmp As String
Dim r As Integer
Dim muls As Integer
Dim n As Integer
Dim m As Integer

On Error GoTo frmErr

For i = 1 To iParamNum
    If arrMulByte(i) = 0 Then
        muls = 1
    Else
        muls = 10 ^ arrMulByte(i)
    End If

    Tmp = vbNullString

    If arrByteCount(i) = 1 Then    '1 byte

        For j = 1 To UBound(carrParamBytes(i))
            Tmp = "0" & Hex((CDbl(txtParam(i - 1).Text)) * muls)
            carrParamBytes(i)(j) = right$(Tmp, 2)
            'carrParamAddrs()
        Next j

    Else    'word

        Tmp = Hex((txtParam(i - 1).Text) * muls)
        Tmp = AddZeroesLeft(Tmp, arrByteCount(i))

        For j = 0 To UBound(carrParamBytes(i)) - 1 Step arrByteCount(i)
            r = 1

            For n = arrByteCount(i) + j To j + 1 Step -1
                m = carrParamBytePos(i)(n) + j
                carrParamBytes(i)(m) = Mid$(Tmp, r, 2)
                r = r + 2
            Next n
        Next j
    End If
Next i


'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Param_Changes2Arrays"
End Sub

Private Sub cmdModifyPatch_Click()

If ValidateInput Then
    Call Changes2Arrays
    Call All2OneArray
    Call Write2File(frmPatch.lstPatches.ListIndex)

    Unload Me
End If

End Sub
Private Sub All2OneArray()
Dim i As Integer
Dim n As Integer
On Error GoTo frmErr

For i = 1 To UBound(carrParamBytes)
    n = n + UBound(carrParamBytes(i))
Next i

ReDim arrAllParamAddrs(n)
ReDim arrAllParamBytes(n)
n = 1
For i = 1 To UBound(carrParamBytes)
    For j = 1 To UBound(carrParamBytes(i))
        arrAllParamBytes(n) = carrParamBytes(i)(j)
        arrAllParamAddrs(n) = carrParamAddrs(i)(j)
        n = n + 1
    Next j
Next i


'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Param_All2OneArray"
End Sub
Private Sub Write2File(ByRef Ind As Integer)
Dim f As Integer
Dim Tmp As String
Dim s() As String
Dim TotalFile As String
Dim Lines() As String
Dim l As Integer
Dim i As Integer
Dim fn As String
On Error GoTo frmErr

f = FreeFile

'Open arrPtcFileName(Ind) For Input Access RWrite As #f
Open arrPtcFileName(Ind) For Binary Access Read Shared As #f
TotalFile = Space(LOF(f))
Get #f, , TotalFile
Close #f


Lines = Split(TotalFile, vbLf)    ' for if no CR in file, only LF
'   Do Until EOF(f)
For l = 0 To UBound(Lines)
    '  Line Input #f, tmp ' need CR
    Tmp = Lines(l)

    For i = 1 To UBound(arrAllParamBytes)
        If InStr(1, Tmp, arrAllParamAddrs(i)) Then

            s = Split(Tmp, ": ", 3)
            If UBound(s) = 2 Then

                Tmp = right$(s(2), Len(s(2)) - 2)
                Tmp = s(0) & ": " & s(1) & ": " & arrAllParamBytes(i) & Tmp
                Lines(l) = Tmp

            End If
        End If
    Next i

Next l
' Loop

'ReDim s(UBound(Lines))
'For i = 0 To UBound(Lines)
's(i) = Lines(i)
'Next i

f = FreeFile
'fn = frmPatch.FilePatches.Path & "\" & "my_" & frmPatch.FilePatches.List(Ind)
fn = arrPtcFileName(Ind)
Open fn For Binary Access Write As #f
Put #f, , Join(Lines, vbLf)
Close #f

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Param_Write2File"
End Sub

Private Function AddZeroesLeft(ByRef s As String, ByRef n As Integer) As String
On Error GoTo frmErr
AddZeroesLeft = Replace(Space(2 * n - Len(s)) & s, mySpace, "0")

'''
Exit Function
frmErr:
MsgBox Err.Description & ": AddZeroesLeft"
End Function
Private Sub Form_Load()
Dim i As Integer
On Error GoTo frmErr

Me.Icon = frmmain.Icon
Me.Caption = "VTCFont: " & frmmain.cmdPatcher.Caption & " / " & frmPatch.cmdParam.Caption

If Not fFileOpen Then Exit Sub

Call GetLanguage(4)

With frmPatch
    Call GetPatchParam(.lstPatches.ListIndex)

    '        If Not CheckParams Then
    '        MsgBoxEx "err in param", , , centerowner, vbCritical
    '        Exit Sub
    '        End If

    For i = 0 To iParamNum - 1
        txtParam(i).Visible = True
    Next i

    Call fillParamWindow(.lstPatches.ListIndex)

End With

frmParameters.Height = 2115
For i = 1 To iParamNum - 1
frmParameters.Height = frmParameters.Height + 420
Next i

txtValidator.top = frmParameters.Height - 900
cmdModifyPatch.top = txtValidator.top

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": frmParam_Load"
End Sub


'Private Function CheckParams() As Boolean
'Dim i As Integer
'
'CheckParams = True
'For i = 1 To iParamNum
'
'If arrMulByte(i) = 0 Then CheckParams = False: Exit Function
'
'Next i
'
'End Function


Private Sub fillParamWindow(ByRef Ind As Integer)
Dim i As Integer
Dim a() As String
Dim Tmp As String
Dim v As Long
Dim muls As Integer

On Error GoTo frmErr

lblPatchName.Caption = "> " & arrPtcName(Ind)

For i = 1 To iParamNum

    If arrMulByte(i) = 0 Then
        muls = 1
    Else
        muls = 10 ^ arrMulByte(i) '10 100 1000...
    End If

    Tmp = vbNullString

    lblParam(i - 1).Caption = arrParamDescr(i)

    If arrByteCount(i) = 1 Then    '1 byte

        txtParam(i - 1).Text = ("&H" & carrParamBytes(i)(1)) / muls


    Else    'word

        ReDim a(arrByteCount(i))
        For j = 1 To UBound(a)
            a(carrParamBytePos(i)(j)) = carrParamBytes(i)(j)
        Next j

        For j = UBound(a) To 1 Step -1
            Tmp = Tmp & a(j)
        Next j

        v = ("&H" & Tmp) / muls
        txtParam(i - 1).Text = v


    End If

Next i

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": fillParamWindow"
End Sub
Private Sub GetPatchParam(Ind As Integer)
Dim f As Integer
Dim Tmp As String
Dim s() As String
Dim n As Integer
Dim dataFlag As Boolean
Dim fSkip As Boolean
Dim TotalFile As String
Dim Lines() As String
Dim l As Integer
Dim max_iParamNum As Integer
Dim iAllBytesCount As Integer
Dim p() As String
Dim r As Integer

On Error GoTo frmErr

ReDim arrPtcAdr(0)
ReDim arrPtcOld(0)
ReDim arrPtcNew(0)
ReDim arrPtcSkip(0)

ReDim arrByteCount(0)    '1 base
ReDim arrMulByte(0)
ReDim arrParamDescr(0)
ReDim arrBytesAll(0)
ReDim arrAddrAll(0)
ReDim arrBytePos(0)

ReDim carrParamBytes(0)
ReDim carrParamAddrs(0)
ReDim carrParamBytePos(0)

f = FreeFile
'Open arrPtcFileName(Ind) For Input Access Read Shared As #f
Open arrPtcFileName(Ind) For Binary Access Read Shared As #f
TotalFile = Space(LOF(f))
Get #f, , TotalFile
Close #f

Lines = Split(TotalFile, vbLf)    ' for if no CR in file, only LF

Select Case arrPatchFormat(Ind)

Case 1    'vtc
    ' Do Until EOF(f)

    iAllBytesCount = 0

    For l = 0 To UBound(Lines)
        Tmp = Lines(l)
        'Line Input #f, Tmp ' need CR

        If InStr(1, Tmp, ": ") Then
            s = Split(Tmp, ": ", 3)
            If UBound(s) = 2 Then

                s(0) = Trim(s(0))
                s(2) = left$(Trim(s(2)), 2)

                If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then

                    If InStr(10, Tmp, "@PARAM@") Then
                        p = Split(Tmp, "@")

                        If UBound(p) = 7 Then    'main param string
                            iParamNum = Val(p(2))    'current param number

                            If max_iParamNum < iParamNum Then max_iParamNum = iParamNum

                            iAllBytesCount = 0

                            ReDim Preserve arrByteCount(iParamNum)
                            ReDim Preserve arrMulByte(iParamNum)
                            ReDim Preserve arrParamDescr(iParamNum)

                            arrByteCount(iParamNum) = Val(p(4))
                            arrMulByte(iParamNum) = Val(p(5))

                            Select Case LCase(Language)
                            Case "en"
                                arrParamDescr(iParamNum) = p(6)
                            Case Else
                                If Len(p(7)) <= 1 Then
                                    arrParamDescr(iParamNum) = p(6)
                                Else
                                    arrParamDescr(iParamNum) = p(7)
                                End If
                            End Select

                            iAllBytesCount = iAllBytesCount + 1

                            ReDim Preserve arrBytesAll(iAllBytesCount)
                            ReDim Preserve arrAddrAll(iAllBytesCount)
                            ReDim Preserve arrBytePos(iAllBytesCount)

                            arrBytesAll(iAllBytesCount) = s(2)
                            arrAddrAll(iAllBytesCount) = s(0)
                            arrBytePos(iAllBytesCount) = Val(p(3))

                            For r = l To UBound(Lines)    'search other parts of this param
                                Tmp = Lines(r)
                                If InStr(1, Tmp, ": ") Then
                                    s = Split(Tmp, ": ", 3)

                                    If UBound(s) = 2 Then
                                        s(0) = Trim(s(0))
                                        s(2) = left$(Trim(s(2)), 2)

                                        If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then
                                            If InStr(10, Tmp, "@PARAM@") Then

                                                p = Split(Tmp, "@")

                                                If UBound(p) = 2 Then    'other addr of CURRENT 1byte param @PARAM@1

                                                    If iParamNum = p(2) Then
                                                   
                                                        iAllBytesCount = iAllBytesCount + 1

                                                        ReDim Preserve arrBytesAll(iAllBytesCount)
                                                        ReDim Preserve arrAddrAll(iAllBytesCount)
                                                        'ReDim Preserve arrBytePos(iAllBytesCount)

                                                        arrBytesAll(iAllBytesCount) = s(2)
                                                        arrAddrAll(iAllBytesCount) = s(0)
                                                        'arrBytePos(iAllBytesCount) = Val(p(3))

                                                    End If

                                                ElseIf UBound(p) = 3 Then    'other addr of current Word param @PARAM@1@1

                                                    If iParamNum = p(2) Then
                                                        iAllBytesCount = iAllBytesCount + 1

                                                        ReDim Preserve arrBytesAll(iAllBytesCount)
                                                        ReDim Preserve arrAddrAll(iAllBytesCount)
                                                        ReDim Preserve arrBytePos(iAllBytesCount)

                                                        arrBytesAll(iAllBytesCount) = s(2)
                                                        arrAddrAll(iAllBytesCount) = s(0)
                                                        arrBytePos(iAllBytesCount) = Val(p(3))

                                                    End If

                                                End If    'If UBound(p)
                                            End If    'If InStr(1, Tmp, "@PARAM@") Then
                                        End If    ' IsNumeric("&
                                    End If    'UBound(s) = 2
                                End If    'InStr(1,
                            Next r

                            If UBound(carrParamBytes) < max_iParamNum Then
                                ReDim Preserve carrParamBytes(max_iParamNum)
                                ReDim Preserve carrParamAddrs(max_iParamNum)
                                ReDim Preserve carrParamBytePos(max_iParamNum)
                            End If
                            carrParamBytes(iParamNum) = arrBytesAll    '?carrParamBytes(1)(2)
                            carrParamAddrs(iParamNum) = arrAddrAll
                            carrParamBytePos(iParamNum) = arrBytePos


                        End If    'If UBound(p) = 7 Then


                    End If    'If InStr(1, Tmp, "@PARAM@") Then


                End If    'If IsNumeric

            End If    'If UBound(s) = 2
        End If    'If InStr(1, Tmp
    Next l

    '   Debug.Print iParamNum

Case 2    'nfe
    dataFlag = False

    For l = 0 To UBound(Lines)
        Tmp = Lines(l)
        '    Do Until EOF(f)
        '       Line Input #f, Tmp

        If dataFlag Then
            If InStr(1, Tmp, ": ") Then
                s = Split(Tmp, " ", 4)    '00001188: 1E - 1D ;# comment
                If UBound(s) = 3 Then
                    s(0) = left$(Trim(s(0)), Len(s(0)) - 1)

                    s(1) = Trim(s(1))
                    fSkip = False
                    If InStr(1, s(1), "null", vbTextCompare) Or InStr(1, s(1), "*") Then
                        s(1) = &H0
                        fSkip = True
                    End If

                    s(2) = left$(Trim(s(3)), 2)

                    If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then

                        n = n + 1
                        ReDim Preserve arrPtcAdr(n)
                        ReDim Preserve arrPtcOld(n)
                        ReDim Preserve arrPtcNew(n)
                        ReDim Preserve arrPtcSkip(n)

                        arrPtcAdr(n) = "&H" & s(0)
                        arrPtcOld(n) = "&H" & s(1)
                        arrPtcNew(n) = "&H" & s(2)

                        If fSkip Then arrPtcSkip(n) = True

                    End If
                End If
            End If
        Else
            If InStr(1, Tmp, "<Data>") Then dataFlag = True
        End If
        'Loop
    Next l

Case 3    'crk , dif
    For l = 0 To UBound(Lines)
        Tmp = Lines(l)
        'Do Until EOF(f)
        '   Line Input #f, Tmp
        If InStr(1, Tmp, ": ") Then    '000037C8: 46 4B ;comment
            s = Split(Tmp, " ", 3)
            If UBound(s) = 2 Then
                s(0) = left$(Trim(s(0)), Len(s(0)) - 1)
                s(1) = Trim(s(1))

                fSkip = False
                If InStr(1, s(1), "null", vbTextCompare) Or InStr(1, s(1), "*") Then
                    s(1) = &H0
                    fSkip = True
                End If

                s(2) = left$(Trim(s(2)), 2)
                If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then
                    n = n + 1
                    ReDim Preserve arrPtcAdr(n)
                    ReDim Preserve arrPtcOld(n)
                    ReDim Preserve arrPtcNew(n)
                    ReDim Preserve arrPtcSkip(n)
                    arrPtcAdr(n) = "&H" & s(0)
                    arrPtcOld(n) = "&H" & s(1)
                    arrPtcNew(n) = "&H" & s(2)

                    If fSkip Then arrPtcSkip(n) = True

                End If
            End If
        End If
        ' Loop
    Next l

End Select

Close #f

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GetPatchParam()"
End Sub


Private Sub txtParam_GotFocus(Index As Integer)
'Dim i As Integer
Dim muls As Integer
On Error GoTo frmErr

'For i = 1 To iParamNum
txtValidator.Text = vbNullString

If arrMulByte(Index + 1) = 0 Then
    muls = 1
Else
    muls = 10 ^ arrMulByte(Index + 1)
End If

Select Case arrByteCount(Index + 1)
Case 1
    txtValidator.Text = "0 - " & 255 / muls
Case 2
    txtValidator.Text = "0 - " & 65535 / muls
Case 4
    txtValidator.Text = "0 - " & 4294967295# / muls
End Select
'Next i

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": txtParam_Change"
End Sub

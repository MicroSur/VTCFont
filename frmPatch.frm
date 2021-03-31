VERSION 5.00
Begin VB.Form frmPatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VTCFont: Patcher"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFWUpdate 
      Caption         =   "cmdFWUpdater"
      Height          =   315
      Left            =   4320
      TabIndex        =   19
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdIda 
      Caption         =   "ida block"
      Height          =   315
      Left            =   -780
      TabIndex        =   18
      Top             =   4500
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMakePatchFolder 
      Caption         =   "Create folder"
      Height          =   315
      Left            =   4020
      TabIndex        =   16
      ToolTipText     =   "Create folder in Patches for this HW"
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton cmdConvert2VTCF 
      Caption         =   "Convert format"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Convert to VTCFont format"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton cmdReload 
      Height          =   315
      Left            =   5640
      Picture         =   "frmPatch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   555
   End
   Begin VB.CommandButton cmdParam 
      Caption         =   "Parameters"
      Enabled         =   0   'False
      Height          =   315
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Edit patch parameters"
      Top             =   5220
      Width           =   1935
   End
   Begin VB.CommandButton cmdPtcConflict 
      Caption         =   "Show conflicts"
      Height          =   315
      Left            =   2220
      TabIndex        =   4
      ToolTipText     =   "Show conflicted pathces for selected patch"
      Top             =   5220
      Width           =   1935
   End
   Begin VB.CommandButton cmdPtcInstalled 
      Caption         =   "Show installed"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Show patches installed in firmware"
      Top             =   5220
      Width           =   1935
   End
   Begin VB.CommandButton cmdViewPatch 
      Caption         =   "View patch in notepad"
      Height          =   315
      Left            =   2220
      TabIndex        =   5
      ToolTipText     =   "View patch text in notepad"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CheckBox chkPtcList 
      Appearance      =   0  'Flat
      Caption         =   "Check/Uncheck all"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4860
      Width           =   2955
   End
   Begin VB.FileListBox FilePatches 
      Height          =   870
      Left            =   6480
      TabIndex        =   15
      Top             =   3060
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstPatches 
      Appearance      =   0  'Flat
      Height          =   4305
      ItemData        =   "frmPatch.frx":058A
      Left            =   120
      List            =   "frmPatch.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   6135
   End
   Begin VB.CommandButton cmdRevert 
      Caption         =   "Revert "
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Revert choosen patches to original version"
      Top             =   7680
      Width           =   1995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      ToolTipText     =   "Close this window"
      Top             =   7680
      Width           =   975
   End
   Begin VB.CommandButton cmdApplyPatch 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   8
      ToolTipText     =   "Apply choosen patches to current firmware"
      Top             =   7680
      Width           =   1995
   End
   Begin VB.TextBox txtPatDescr 
      Appearance      =   0  'Flat
      Height          =   1515
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   6060
      Width           =   6135
   End
   Begin VB.Label lblCheched 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   60
      TabIndex        =   17
      Top             =   4890
      Width           =   255
   End
   Begin VB.Label lblCurrentFW 
      AutoSize        =   -1  'True
      Caption         =   "FW"
      Height          =   195
      Left            =   2280
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lstPatchList 
      AutoSize        =   -1  'True
      Caption         =   "List of patches for:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label lblPatchDescr 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   4140
      Width           =   1515
   End
End
Attribute VB_Name = "frmPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrDescrEn() As String
Private arrDescrRu() As String
'Private arrPatchFormat() As Byte    '1 vtc, 2 nfe, 3 crk
'in module Private arrPtcAdr() As Long    'hex address of changes
Private arrPtcOld() As Byte    'hex values
'Private arrPtcNew() As Byte
Private arrPtcSkip() As Boolean

'Private arrPtcFileName() As String
'Private arrPtcName() As String
Private fShowInstalledPatches As Boolean
Private fShowConflictedPatches As Boolean

Private arrPtcAdrCollection() As Variant
Private ListMarkersInstalled() As String
Private ListMarkersConflicts() As String

Private NoPatchesFlag As Boolean


Private Sub chkPtcList_Click()
Dim i As Integer
Dim Tmp As Integer

On Error GoTo frmErr
Tmp = lstPatches.ListIndex

lstPatches.Enabled = False

For i = 0 To lstPatches.ListCount - 1
    lstPatches.Selected(i) = chkPtcList.Value
Next i


lstPatches.ListIndex = Tmp    '-1 not work in exe
lstPatches.Enabled = True
Call CalcChecked

'''
Exit Sub
frmErr:
lstPatches.Enabled = True
MsgBox Err.Description & ": chkPtcList_Click()"
End Sub

Private Sub cmdApplyPatch_Click()
If NoPatchesFlag Then Exit Sub
Call GoPatch(True)
End Sub

Private Sub GetPatch(Ind As Integer)
Dim f As Integer
Dim Tmp As String
Dim s() As String
Dim n As Integer
Dim dataFlag As Boolean
Dim fSkip As Boolean
Dim TotalFile As String
Dim Lines() As String
Dim l As Integer
Dim sComment As String

On Error GoTo frmErr

ReDim arrPtcAdr(0)
ReDim arrPtcOld(0)
ReDim arrPtcNew(0)
ReDim arrPtcSkip(0)
ReDim arrPtcComment(0)    'for convert

n = -1
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
    '6A8A: 04: 0D ;comments

    For l = 0 To UBound(Lines)
        Tmp = Lines(l)
        'Line Input #f, Tmp ' need CR


        If InStr(1, Tmp, ": ") Then
            s = Split(Tmp, ": ", 3)
            If UBound(s) = 2 Then
                s(0) = Trim(s(0))

                s(1) = Trim(s(1))
                fSkip = False
                If InStr(1, s(1), "null", vbTextCompare) Or InStr(1, s(1), "*") Then
                    s(1) = &H0
                    fSkip = True
                End If

                s(2) = Trim(s(2))
                'sComment = right(s(2), Len(s(2)) - 2)
                s(2) = left$(s(2), 2)

                If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then

                    n = n + 1
                    ReDim Preserve arrPtcAdr(n)
                    ReDim Preserve arrPtcOld(n)
                    ReDim Preserve arrPtcNew(n)
                    ReDim Preserve arrPtcSkip(n)
                    'ReDim Preserve arrPtcComment(n)
                    arrPtcAdr(n) = "&H" & s(0)
                    arrPtcOld(n) = "&H" & s(1)
                    arrPtcNew(n) = "&H" & s(2)
                    'arrPtcComment(n) = sComment 'with crlf, not use in vtc

                    If fSkip Then arrPtcSkip(n) = True    ' skip redo

                End If
            End If
        End If
        ' Loop
    Next l

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

                    s(3) = Trim(s(3))
                    sComment = right$(s(3), Len(s(3)) - 2)
                    s(2) = left$(s(3), 2)

                    If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then

                        n = n + 1
                        ReDim Preserve arrPtcAdr(n)
                        ReDim Preserve arrPtcOld(n)
                        ReDim Preserve arrPtcNew(n)
                        ReDim Preserve arrPtcSkip(n)
                        ReDim Preserve arrPtcComment(n)

                        arrPtcAdr(n) = "&H" & s(0)
                        arrPtcOld(n) = "&H" & s(1)
                        arrPtcNew(n) = "&H" & s(2)
                        arrPtcComment(n) = sComment    'with crlf

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

                s(2) = Trim(s(2))
                sComment = right$(s(2), Len(s(2)) - 2)
                s(2) = left$(s(2), 2)

                If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then
                    n = n + 1
                    ReDim Preserve arrPtcAdr(n)
                    ReDim Preserve arrPtcOld(n)
                    ReDim Preserve arrPtcNew(n)
                    ReDim Preserve arrPtcSkip(n)
                    ReDim Preserve arrPtcComment(n)

                    arrPtcAdr(n) = "&H" & s(0)
                    arrPtcOld(n) = "&H" & s(1)
                    arrPtcNew(n) = "&H" & s(2)
                    arrPtcComment(n) = sComment    'with crlf

                    If fSkip Then arrPtcSkip(n) = True

                End If
            End If
        End If
        ' Loop
    Next l

End Select

If fShowConflictedPatches Then    'fill collection
    arrPtcAdrCollection(Ind) = arrPtcAdr
End If

Close #f

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GetPatch()"
End Sub
Private Function CheckPatchPresent() As Boolean
'true if patch is in fw
Dim i As Long
Dim X As Byte
Dim b As Byte
On Error GoTo frmErr

If UBound(arrPtcAdr) > -1 Then
    CheckPatchPresent = True
    For i = 0 To UBound(arrPtcAdr)

        Seek #bFileIn, arrPtcAdr(i) + 1
        Get #bFileIn, , X
        b = (X Xor (arrPtcAdr(i) + lngBytes + uMagic - lngBytes \ uMagic)) And 255

        If b <> arrPtcNew(i) Then
            CheckPatchPresent = False
            Exit For
        End If

    Next i
End If

If UBound(arrPtcAdr) = 0 And arrPtcAdr(0) = 0 Then CheckPatchPresent = False

'''
Exit Function
frmErr:
MsgBox Err.Description & ": CheckPatchPresent()"
End Function
Private Sub WritePatch(ByRef fApply As Boolean)
'char 2 file
Dim i As Long
Dim X As Byte
Dim arrX() As Byte
Dim arrB() As Byte
Dim MaxAdr As Long
On Error GoTo frmErr

For i = 0 To UBound(arrPtcAdr)
    If arrPtcAdr(i) > MaxAdr Then MaxAdr = arrPtcAdr(i)
Next i
MaxAdr = MaxAdr + 1
If MaxAdr > lngBytes Then    'enlarge fw file

    ReDim arrX(lngBytes - 1)
    ReDim arrB(lngBytes - 1)
    Seek #bFileIn, 1
    Get #bFileIn, , arrX()

    For i = 0 To lngBytes - 1
        arrB(i) = (arrX(i) Xor (i + lngBytes + uMagic - lngBytes \ uMagic)) And 255
    Next i

    ReDim Preserve arrB(MaxAdr - 1)
    ReDim arrX(MaxAdr - 1)
    For i = 0 To MaxAdr - 1
        arrX(i) = (arrB(i) Xor (i + MaxAdr + uMagic - MaxAdr \ uMagic)) And 255
    Next i

    Seek #bFileIn, 1
    Put #bFileIn, , arrX()

    lngBytes = MaxAdr

End If

For i = 0 To UBound(arrPtcAdr)

    If fApply Then
        X = (arrPtcNew(i) Xor (arrPtcAdr(i) + lngBytes + uMagic - lngBytes \ uMagic)) And 255
    Else    'redo
        If Not arrPtcSkip(i) Then
            X = (arrPtcOld(i) Xor (arrPtcAdr(i) + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        End If
    End If

    Seek #bFileIn, arrPtcAdr(i) + 1
    Put #bFileIn, , X

Next i

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": WritePatch()"
End Sub
Private Sub cmdCancel_Click()
Unload Me

End Sub
Private Sub ShowAllConflicts()
Dim i As Integer, j As Integer, k As Integer, m As Integer    ', n As Integer
Dim h As Integer
Dim sConfFlag As Boolean
Dim arrConflict() As Boolean
'Dim arrIndex() As Integer
Dim sMark As String
Dim Tmp() As String
Dim sTmp As String

Dim indCount As Integer
'Dim SelInd As Integer

'Dim sTab As String

On Error GoTo frmErr

'tmp = lstPatches.ListIndex

fShowConflictedPatches = Not fShowConflictedPatches
'lstPatches.Visible = False
lstPatches.Enabled = False

Me.MousePointer = vbHourglass

indCount = lstPatches.ListCount - 1
'SelInd = lstPatches.ListIndex

'fShowInstalledPatches = False
'For i = 0 To lstPatches.ListCount - 1
'       lstPatches.List(i) = arrPtcName(i)
'Next i

If fShowConflictedPatches Then

    txtPatDescr.Text = vbNullString
    ReDim arrPtcAdrCollection(indCount)
    ReDim Tmp(indCount)
    ReDim arrConflict(indCount, indCount)
    ReDim ListMarkersConflicts(indCount)

    For i = 0 To indCount
        Call GetPatch(i)
    Next i

    For i = 0 To indCount
        For j = i + 1 To indCount

            For k = 0 To UBound(arrPtcAdrCollection(i))
                For m = 0 To UBound(arrPtcAdrCollection(j))
                    sConfFlag = False
                    If arrPtcAdrCollection(i)(k) = arrPtcAdrCollection(j)(m) Then
                        arrConflict(i, j) = True
                        sConfFlag = True
                        Exit For
                    End If
                Next m
                If sConfFlag Then Exit For
            Next k

        Next j
    Next i

'go mark
    h = 1
    For i = 0 To indCount
        sConfFlag = False
        k = 0
'   ReDim arrIndex(indCount)    ' store other conflicted patches

        sTmp = h & " "

        For j = i + 1 To indCount

            If arrConflict(i, j) Then

                sConfFlag = True
                Tmp(j) = Tmp(j) & sTmp
                ListMarkersConflicts(j) = "[" & Trim(Tmp(j)) & "] "

                If fShowInstalledPatches Then
                    sMark = ListMarkersInstalled(j) & ListMarkersConflicts(j)
                Else
                    sMark = ListMarkersConflicts(j)
                End If

                lstPatches.List(j) = sMark & arrPtcName(j)  'mark other

'                arrIndex(k) = j
'                k = k + 1

            End If

        Next j

        If sConfFlag Then
            Tmp(i) = Tmp(i) & sTmp
            ListMarkersConflicts(i) = "[" & Trim(Tmp(i)) & "] "

            If fShowInstalledPatches Then
                sMark = ListMarkersInstalled(i) & ListMarkersConflicts(i)
            Else
                sMark = ListMarkersConflicts(i)
            End If

            lstPatches.List(i) = sMark & arrPtcName(i)  'mark first
            h = h + 1


'no
'            For m = 0 To k - 1 'false to already marked
'                For n = m + 1 To k - 1
'                    arrConflict(arrIndex(m), arrIndex(n)) = False
'                Next n
'            Next m

        End If
    Next i

    txtPatDescr.Text = "(" & h - 1 & ")"
    If h = 1 Then fShowConflictedPatches = False
Else

    For i = 0 To indCount
        lstPatches.List(i) = arrPtcName(i)
    Next i
    fShowConflictedPatches = False
    fShowInstalledPatches = False

End If

Me.MousePointer = vbNormal
'lstPatches.Visible = True
lstPatches.Enabled = True
'lstPatches.ListIndex = tmp    '-1 not work in exe

'''
Exit Sub
frmErr:
'lstPatches.Visible = True
lstPatches.Enabled = True
Me.MousePointer = vbNormal
MsgBox Err.Description & ": ShowAllConflicts"
End Sub

Private Sub cmdConvert2VTCF_Click()
'make tmp file with old patch
'write vtc format
Dim s As New CString
Dim Ind As Integer
Dim i As Integer
Dim f As Integer
Dim Tmp As String
'Dim fn As String
Dim Ret As Long

On Error GoTo frmErr
'<PatchName>
'<Eng Description>
'<Rus Description>
'# RX23 4.02 4.12
'...patch.data.. 49B2: 05: 03

Ret = MsgBoxEx(ArrMsg(46), , , CenterOwner, vbQuestion Or vbOKCancel)
If Ret <> 1 Then Exit Sub

Ind = lstPatches.ListIndex
Select Case arrPatchFormat(Ind)
Case 2, 3
    Tmp = Replace(arrPtcName(Ind), vbCrLf, vbNullString)
    Tmp = Replace(Tmp, vbLf, vbNullString)
    s.concat "<PatchName>" & Tmp & vbCrLf
    Tmp = Replace(arrDescrEn(Ind), vbCrLf, vbNullString)
    Tmp = Replace(Tmp, vbLf, vbNullString)
    s.concat "<Eng Description>" & Tmp & vbCrLf
    s.concat "<Rus Description>" & Tmp & vbCrLf
    s.concat "# " & Hardtext & vbCrLf
    s.concat vbCrLf

    Call GetPatch(Ind)

   'no If UBound(arrPtcAdr) = 0 Then Exit Sub

    For i = 0 To UBound(arrPtcAdr)
        s.concat right$("00000000" & Hex(arrPtcAdr(i)), 8)
        s.concat ": "
        
        s.concat right$("0" & Hex(arrPtcOld(i)), 2)
        s.concat ": "
        
        s.concat right$("0" & Hex(arrPtcNew(i)), 2)
        
        Tmp = Replace(arrPtcComment(i), vbCrLf, vbNullString)
        Tmp = Replace(Tmp, vbLf, vbNullString)
        s.concat Tmp
        
        s.concat vbCrLf
    Next i
    
Case Else
    Exit Sub
End Select

'tmp = App.Path & "\Patches\" & Hardtext
'fn = tmp & "\" & FilePatches.List(ind)
'arrPtcFileName


f = FreeFile
Open arrPtcFileName(Ind) For Binary Access Read Shared As #f
Tmp = Space(LOF(f))
Get #f, , Tmp
Close #f

f = FreeFile
Open arrPtcFileName(Ind) & ".bak" For Output As #f
Print #f, Tmp;
Close #f

'tmp = s.Text
'tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)

f = FreeFile
'fn = GetName(arrPtcFileName(ind))
'fn = GetPathFromPathAndName(arrPtcFileName(ind)) & fn & ".patch"
Open arrPtcFileName(Ind) For Output As #f
Print #f, s.Text;
Close #f

cmdConvert2VTCF.Enabled = False
arrPatchFormat(Ind) = 1    'vtcf now
txtPatDescr.Text = "Ok."

'''
Exit Sub
frmErr:
Close #f
MsgBox Err.Description & ": cmdConvert2VTCF"
End Sub

Private Sub cmdFWUpdate_Click()
frmmain.cmdFWUpdater_Click
End Sub

Private Sub cmdIda_Click()

Dim i As Integer, j As Integer
Dim Tmp As New CString
On Error Resume Next

'ReDim arrPtcAdr(0)
'ReDim arrPtcOld(0)
'ReDim arrPtcNew(0)
'ReDim arrPtcSkip(0)
'ReDim arrPtcComment(0)    'for convert

fFileOpen = OpenFW_write
If Not fFileOpen Then Exit Sub

Call GetPatch(lstPatches.ListIndex)

Tmp.concat Hex(arrPtcAdr(0)) & ": " & right$("0" & Hex(arrPtcNew(0)), 2) & " "

For i = 1 To UBound(arrPtcAdr)

    If arrPtcAdr(i) = arrPtcAdr(i - 1) + 1 Then

        j = j + 1
        If j = 16 Then
            Tmp.concat vbCrLf & Hex(arrPtcAdr(i)) & ": "
            j = 0
        End If

    Else
    
        Tmp.concat vbCrLf & Hex(arrPtcAdr(i)) & ": "
        j = 0
        
    End If

    Tmp.concat right$("0" & Hex(arrPtcNew(i)), 2) & " "


Next i

'Debug.Print tmp.Text
Clipboard.Clear
Clipboard.SetText Tmp.Text

End Sub


Private Sub cmdMakePatchFolder_Click()
Dim Tmp As String
On Error GoTo frmErr
Tmp = App.Path & "\Patches\" & Hardtext
If Not FileExists(Tmp) Then
    'create
    MkDir Tmp
    cmdMakePatchFolder.Enabled = False
Else
    cmdMakePatchFolder.Enabled = False
    Exit Sub
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": MakePatchFolder"
End Sub

Private Sub cmdParam_Click()
frmParameters.Show 1, frmPatch
End Sub

Private Sub cmdPtcConflict_Click()
'
Dim i As Integer, j As Integer, k As Integer, m As Integer    ', n As Integer
'Dim h As Integer
Dim sConfFlag As Boolean
Dim arrConflict() As Boolean
'Dim arrIndex() As Integer
Dim sMark As String
Dim Tmp() As String
'Dim sTmp As String
'Dim sTab As String
Dim indCount As Integer
Dim SelInd As Integer

On Error GoTo frmErr

'tmp = lstPatches.ListIndex
If lstPatches.ListIndex < 0 Then Exit Sub

'fShowConflictedPatches = Not fShowConflictedPatches
fShowConflictedPatches = True
txtPatDescr.Text = vbNullString

indCount = lstPatches.ListCount - 1
SelInd = lstPatches.ListIndex

Me.MousePointer = vbHourglass
'lstPatches.Visible = False
lstPatches.Enabled = False

'fShowInstalledPatches = False
'For i = 0 To lstPatches.ListCount - 1
'       lstPatches.List(i) = arrPtcName(i)
'Next i


If fShowConflictedPatches Then
    txtPatDescr.Text = vbNullString

    ReDim ListMarkersConflicts(indCount)
    ReDim arrPtcAdrCollection(indCount)
    ReDim Tmp(indCount)
    ReDim arrConflict(indCount, indCount)
    ReDim ListMarkersConflicts(indCount)

    For i = 0 To indCount
        Call GetPatch(i)
    Next i

    i = 0
    For j = 0 To indCount

        sMark = vbNullString
        If fShowInstalledPatches Then sMark = ListMarkersInstalled(j)
        lstPatches.List(j) = sMark & arrPtcName(j)  'unmark

        For k = 0 To UBound(arrPtcAdrCollection(SelInd))
            For m = 0 To UBound(arrPtcAdrCollection(j))
                sConfFlag = False
' If j <> lstPatches.ListIndex Then
                If arrPtcAdrCollection(SelInd)(k) = arrPtcAdrCollection(j)(m) Then
                    i = i + 1

                    ListMarkersConflicts(j) = "[@] "

                    If fShowInstalledPatches Then
                        sMark = ListMarkersInstalled(j) & ListMarkersConflicts(j)
                    Else
                        sMark = ListMarkersConflicts(j)
                    End If

                    lstPatches.List(j) = sMark & arrPtcName(j)  'mark
                    txtPatDescr.Text = txtPatDescr.Text & arrPtcName(j) & vbCrLf
                    sConfFlag = True
                    Exit For

                End If
'End If
            Next m
            If sConfFlag Then Exit For
        Next k

    Next j

    txtPatDescr.Text = txtPatDescr.Text & "(+" & i - 1 & ")"
    If i = 1 Then    '0
        fShowConflictedPatches = False

        sMark = vbNullString
        If fShowInstalledPatches Then sMark = ListMarkersInstalled(SelInd)
        lstPatches.List(SelInd) = sMark & arrPtcName(SelInd)  'unmark selected

    Else
        ListMarkersConflicts(SelInd) = "[@] "
        If fShowInstalledPatches Then
            sMark = ListMarkersInstalled(SelInd) & ListMarkersConflicts(SelInd)
        Else
            sMark = ListMarkersConflicts(SelInd)
        End If
        lstPatches.List(SelInd) = sMark & arrPtcName(SelInd)  'mark selected

    End If

Else

    For i = 0 To indCount
        lstPatches.List(i) = arrPtcName(i)
    Next i
    fShowConflictedPatches = False
    fShowInstalledPatches = False

End If

'lstPatches.Visible = True
lstPatches.Enabled = True
Me.MousePointer = vbNormal

'''
Exit Sub
frmErr:
'lstPatches.Visible = True
lstPatches.Enabled = True
Me.MousePointer = vbNormal
MsgBox Err.Description & ": cmdPtcConflict"
End Sub

Private Sub cmdPtcConflict_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Call ShowAllConflicts
End Sub

Private Sub cmdPtcInstalled_Click()
Call CheckInstalledPatches
End Sub

Private Sub cmdReload_Click()
lstPatches.Clear
Call LoadCurrentPatches
If NoPatchesFlag Then
    cmdParam.Enabled = False
    cmdConvert2VTCF.Enabled = False
Exit Sub
End If
If OldPatchIndex <= lstPatches.ListCount Then lstPatches.ListIndex = OldPatchIndex
Call ListClick

fShowInstalledPatches = False: cmdPtcInstalled.Caption = ArrMsg(41)

End Sub

Private Sub cmdRevert_Click()
If NoPatchesFlag Then Exit Sub
Call GoPatch(False)

End Sub
Private Sub GoPatch(fApply As Boolean)
Dim i As Integer
Dim flag As Boolean
On Error GoTo frmErr

fFileOpen = OpenFW_write
If Not fFileOpen Then Exit Sub

For i = 0 To lstPatches.ListCount - 1
    If lstPatches.Selected(i) = True Then    ' checked
        Call GetPatch(i)
        Call WritePatch(fApply)
        flag = True
    End If
Next i

fFileOpen = OpenFW_read

If flag Then
    txtPatDescr.Text = ArrMsg(9) & " (" & lstPatches.SelCount & ")"
Else
    txtPatDescr.Text = ArrMsg(10)
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GoPatch"
End Sub
Private Sub CheckInstalledPatches()
Dim i As Integer
'Dim tmp As Integer
'Dim sTab As String
Dim sMark As String
Dim noInstalled As Boolean

On Error GoTo frmErr

'tmp = lstPatches.ListIndex

fShowInstalledPatches = Not fShowInstalledPatches
'fShowConflictedPatches = False
'    For i = 0 To lstPatches.ListCount - 1
'        lstPatches.List(i) = arrPtcName(i)
'    Next i

'lstPatches.Visible = False
lstPatches.Enabled = False


If fShowInstalledPatches Then
    cmdPtcInstalled.Caption = ArrMsg(42)

    txtPatDescr.Text = vbNullString
    noInstalled = True

    For i = 0 To lstPatches.ListCount - 1
        Call GetPatch(i)

        If CheckPatchPresent Then
            noInstalled = False

            ListMarkersInstalled(i) = vbTab & "[v] "

            If fShowConflictedPatches Then
                sMark = ListMarkersInstalled(i) & ListMarkersConflicts(i)
            Else
                sMark = ListMarkersInstalled(i)
            End If

            lstPatches.List(i) = sMark & arrPtcName(i)
            txtPatDescr.Text = txtPatDescr.Text & arrPtcName(i) & vbCrLf
'Else
'    lstPatches.List(i) = lstPatches.List(i)
        End If
    Next i

Else    'hide
    cmdPtcInstalled.Caption = ArrMsg(41)
    For i = 0 To lstPatches.ListCount - 1
        lstPatches.List(i) = arrPtcName(i)
    Next i
    fShowConflictedPatches = False
    fShowInstalledPatches = False
    txtPatDescr.Text = vbNullString
End If

If noInstalled Then
    txtPatDescr.Text = "(0)"
    fShowConflictedPatches = False
    fShowInstalledPatches = False
    cmdPtcInstalled.Caption = ArrMsg(41)
End If

'lstPatches.Visible = True
lstPatches.Enabled = True

'LockWindowUpdate 0
'lstPatches.ListIndex = tmp    '-1 not work in exe

'''
Exit Sub
frmErr:
'lstPatches.Visible = True
lstPatches.Enabled = True
MsgBox Err.Description & ": CheckInstalledPatches"
End Sub
Private Sub cmdViewPatch_Click()

On Error GoTo frmErr
If NoPatchesFlag Then Exit Sub
If lstPatches.ListIndex < 0 Then Exit Sub

Shell "notepad.exe " & arrPtcFileName(lstPatches.ListIndex), vbNormalFocus
'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdViewPatch_Click()"
End Sub


Private Sub Form_Load()
On Error GoTo frmErr

DoEvents
Me.Icon = frmmain.Icon
Me.Caption = "VTCFont: " & frmmain.cmdPatcher.Caption
lblCurrentFW.Caption = Hardtext

chkPtcList.Value = vbUnchecked

lstPatches.Clear
If Len(lngFileName) <> 0 Then Call GetLanguage(2)
cmdFWUpdate.Caption = frmmain.cmdFWUpdater.Caption
cmdFWUpdate.ToolTipText = frmmain.cmdFWUpdater.ToolTipText

lblCurrentFW.left = lstPatchList.left + lstPatchList.Width + 6

fShowInstalledPatches = False
fShowConflictedPatches = False

'in loadFW OldPatchIndex = 0
Call LoadCurrentPatches
If NoPatchesFlag Then
    cmdParam.Enabled = False
    cmdConvert2VTCF.Enabled = False
Exit Sub
End If

If OldPatchIndex <= lstPatches.ListCount Then lstPatches.ListIndex = OldPatchIndex
'lstPatches.ListIndex = 0 'OldPatchIndex
Call ListClick
'lstPatches.SetFocus

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Patcher Form_Load()"
End Sub

Private Sub LoadCurrentPatches()
'for current hardware
'setup only Names and descriptions here
Dim i As Integer ', j As Integer
Dim f As Integer
Dim l As Integer
Dim fn As String
Dim sLine As String    'line input
Dim sBlock As String
Dim s() As String
Dim count As Integer
Dim fVTCPatch As Boolean
Dim fNfePatch As Boolean
Dim fCrkPatch As Boolean
Dim noNameflag As Boolean
Dim Tmp As String
Dim TotalFile As String
Dim Lines() As String

On Error GoTo frmErr

'If Not Me.Visible Then Exit Sub
NoPatchesFlag = False
f = FreeFile

cmdMakePatchFolder.Enabled = False
Tmp = App.Path & "\Patches\" & Hardtext
If Not FileExists(Tmp) Then
    NoPatchesFlag = True    'no dir
    cmdMakePatchFolder.Enabled = True
    Exit Sub
End If



FilePatches.Path = Tmp
FilePatches.Pattern = "*.patch;*.dif"
FilePatches.Refresh

If FilePatches.ListCount = 0 Then
    NoPatchesFlag = True    'empty dir
    Exit Sub
End If

ReDim arrDescrEn(FilePatches.ListCount - 1)
ReDim arrDescrRu(FilePatches.ListCount - 1)
ReDim arrPatchFormat(FilePatches.ListCount - 1)

ReDim arrPtcFileName(FilePatches.ListCount - 1)
ReDim arrPtcName(FilePatches.ListCount - 1)

lstPatches.Enabled = False
lstPatches.Visible = False
Me.MousePointer = vbHourglass
'DoEvents

For i = 0 To FilePatches.ListCount - 1
    count = 0
    noNameflag = True

'Debug.Print FilePatches.List(i)
    fn = FilePatches.Path & "\" & FilePatches.List(i)
    arrPtcFileName(i) = fn

'Open fn For Input Access Read As #f
    Open fn For Binary Access Read Shared As #f

    TotalFile = Space(LOF(f))
    Get #f, , TotalFile
    Close #f
    Lines = Split(TotalFile, vbLf)

    fVTCPatch = False: fNfePatch = False: fCrkPatch = False

If UBound(Lines) < 1 Then
lstPatches.AddItem "!ERROR! in: " & FilePatches.List(i)
End If

' Do Until EOF(f)
    For l = 0 To UBound(Lines)
        sLine = Lines(l)
        sLine = Replace(sLine, vbCr, vbNullString)

'    Line Input #f, sLine

''''''''''''''''''''''''''''''''''VTCFont patch
'<patchname> 1 line
'<eng description> 1 line
'<rus description> 1 line
'#...
'6A8A: 04: 0D ;#...

        If (Not fNfePatch) And (Not fCrkPatch) Then
            If InStr(1, LCase(sLine), "<patchname>") > 0 Then
                fVTCPatch = True
                arrPatchFormat(i) = 1
                s = Split(sLine, ">")
                If UBound(s) > 0 Then
                    arrPtcName(i) = s(1)
                    lstPatches.AddItem s(1)
'                lstPatches.ItemData(lstPatches.NewIndex) = j
'                j = j + 1
                End If
'count = count + 1

            ElseIf InStr(1, LCase(sLine), "<eng description>") > 0 Then
                s = Split(sLine, ">", 2)
                If UBound(s) > 0 Then

                    If Len(arrDescrEn(i)) = 0 Then    '0ne or more Description lines
                        arrDescrEn(i) = mySpace & s(1)
                    Else
                        arrDescrEn(i) = arrDescrEn(i) & vbCrLf & mySpace & s(1)
                    End If
                End If

'count = count + 1
            ElseIf InStr(1, LCase(sLine), "<rus description>") > 0 Then
                s = Split(sLine, ">", 2)

                If UBound(s) > 0 Then
                    If Len(arrDescrRu(i)) = 0 Then
                        arrDescrRu(i) = mySpace & s(1)
                    Else
                        arrDescrRu(i) = arrDescrRu(i) & vbCrLf & mySpace & s(1)
                    End If
                End If

            Else
                If left$(sLine, 1) = "#" Then Exit For
'count = count + 1
            End If

'If count = 3 Then Exit Do

        End If

''''''''''''''''''''''''''''''Nfe
'nfe
'<Patch Definition="Evic VTC Mini 3.01" Name="Splash 2" Version="1.1" Author="Team">
'<Description>..........
'...........Charge Screen Mod is recommended to be installed.
'</Description>
'  <Data>
'xx: yy - zz #...
'#.....
'xx: yy - zz
'</Data>
'</Patch>
        If (Not fVTCPatch) And (Not fCrkPatch) Then
            If InStr(1, sLine, "<Patch Definition=") > 0 Then
                fNfePatch = True
                arrPatchFormat(i) = 2
                If InStr(1, sLine, "Name=""") > 0 Then
                    sBlock = GetBlockFromText(sLine, "Name=""", """")
                    If InStr(1, sLine, "Version=""") > 0 Then
                        sBlock = sBlock & " v" & GetBlockFromText(sLine, "Version=""", """")
                    End If
                    If Len(sBlock) > 0 Then
                        arrPtcName(i) = sBlock
                        lstPatches.AddItem sBlock
                        sBlock = vbNullString
                    End If
                End If

            ElseIf InStr(1, sLine, "<Description>") > 0 Then

                sBlock = sBlock & GetBlockFromText(sLine, "<Description>", "</Description>") & vbCrLf
'save it, if 1 line descr
                sBlock = DecodeUTF8(sBlock)
                arrDescrEn(i) = sBlock
                arrDescrRu(i) = sBlock


            ElseIf InStr(1, sLine, "</Description>") > 0 Then
                sBlock = sBlock & GetBlockFromText(sLine, "<Description>", "</Description>") & vbCrLf
'multiline descr
                sBlock = DecodeUTF8(sBlock)
                arrDescrEn(i) = sBlock
                arrDescrRu(i) = sBlock
                Exit For

            Else
                If InStr(1, sLine, "<Data>") > 0 Then Exit For
                sBlock = sBlock & GetBlockFromText(sLine, "<Description>", "</Description>") & vbCrLf

            End If
        End If



''''''''''''''''''''''''''dif''''crk
'crk
'Заголовок патча, автор, версия
'Пустая строка
'Патч
'xx: yy zz
'xx: yy zz
'Патч#2
'xx: yy zz
'xx: yy zz
        If (Not fVTCPatch) And (Not fNfePatch) Then
            If left$(sLine, 1) <> "<" And Len(sLine) <> 0 And noNameflag Then
                fCrkPatch = True
                arrPatchFormat(i) = 3
                arrPtcName(i) = sLine
                lstPatches.AddItem sLine
                sBlock = vbNullString
                noNameflag = False
'Exit Do
            End If
            If fCrkPatch Then
                If count = 1 Then
                    sBlock = sBlock & sLine & vbCrLf
                    count = 0
                End If
                If Len(sLine) = 0 Then count = 1

'If EOF(f) Then
                If l = UBound(Lines) Then
                    arrDescrEn(i) = sBlock
                    arrDescrRu(i) = sBlock
                End If

            End If
        End If
''''''''''''''''''''''''''''''''''''''


'   Loop
    Next l
'   Close #f
Next i

'LockWindowUpdate 0
Me.MousePointer = vbNormal
lstPatches.Enabled = True
lstPatches.Visible = True

    ReDim ListMarkersInstalled(lstPatches.ListCount - 1)
    ReDim ListMarkersConflicts(lstPatches.ListCount - 1)

'''
Exit Sub
frmErr:
'LockWindowUpdate 0
Me.MousePointer = vbNormal
lstPatches.Enabled = True
lstPatches.Visible = True
MsgBox Err.Description & ": LoadCurrentPatches()"
NoPatchesFlag = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
DoEvents
End Sub

Private Sub lblCheched_Click()
Call cmdIda_Click
End Sub

Private Sub lstPatches_ItemCheck(Item As Integer)
Call CalcChecked
End Sub

Private Sub lstPatches_KeyUp(KeyCode As Integer, Shift As Integer)
'SAME in mousedown
Call ListClick
Call CalcChecked
End Sub
Private Sub CalcChecked()
'On Error Resume Next
  Dim i As Integer, X As Integer

    For i = 0 To lstPatches.ListCount - 1
        If lstPatches.Selected(i) = True Then    ' if the item is selected(checked)
        'If lstPatches.ListIndex <> i Then
        X = X + 1
        End If
    Next
    
lblCheched.Caption = X

End Sub
Private Sub ListClick()
Dim Ind As Integer
On Error GoTo frmErr

If NoPatchesFlag Then Exit Sub
Ind = lstPatches.ListIndex


Select Case LCase(Language)
Case "ru"
    txtPatDescr.Text = arrDescrRu(Ind)
Case Else
    txtPatDescr.Text = arrDescrEn(Ind)
End Select

txtPatDescr.Text = txtPatDescr.Text & vbCrLf & vbCrLf & "***" & vbCrLf & _
arrPtcName(Ind) & vbCrLf & FilePatches.List(Ind)

If CheckPatchParam(Ind) Then
    cmdParam.Enabled = True
Else
    cmdParam.Enabled = False
End If

Select Case arrPatchFormat(Ind)
Case 2, 3
    cmdConvert2VTCF.Enabled = True
Case 1
    'myformat
    cmdConvert2VTCF.Enabled = False
End Select

OldPatchIndex = Ind

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": ListClick"
End Sub

Private Sub lstPatches_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'SAME in keyUp
Call ListClick
End Sub


Private Function CheckPatchParam(Ind As Integer) As Boolean
Dim f As Integer
Dim Tmp As String
Dim s() As String
Dim n As Integer
'Dim dataFlag As Boolean
'Dim fSkip As Boolean
Dim TotalFile As String
Dim Lines() As String
Dim l As Integer

'Dim iParamNum As Integer
Dim iAllBytesCount As Integer
Dim p() As String

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
ReDim arrParamsCount(0)
ReDim arrBytePos(0)

ReDim carrParamBytes(0)
ReDim carrParamAddrs(0)
ReDim carrParamBytePos(0)

'n = -1
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
    '   arrParamsCount(Ind) = 0
    
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

                    n = n + 1

                    If InStr(10, Tmp, "@PARAM@") Then
                        p = Split(Tmp, "@")

                        If UBound(p) = 7 Then
                        CheckPatchParam = True
                        Exit Function
                        End If
                       
                        End If    'If UBound(p) = 7 Then

                    End If    'If InStr(1, Tmp, "@PARAM@") Then

                End If


            End If
        
    Next l

  '    Debug.Print iParamNum 'ubound array in collection

Case 2    'nfe
'    dataFlag = False
'
'    For l = 0 To UBound(Lines)
'        tmp = Lines(l)
'        '    Do Until EOF(f)
'        '       Line Input #f, Tmp
'
'        If dataFlag Then
'            If InStr(1, tmp, ": ") Then
'                s = Split(tmp, " ", 4)    '00001188: 1E - 1D ;# comment
'                If UBound(s) = 3 Then
'                    s(0) = left(Trim(s(0)), Len(s(0)) - 1)
'
'                    s(1) = Trim(s(1))
'                    fSkip = False
'                    If InStr(1, s(1), "null", vbTextCompare) Or InStr(1, s(1), "*") Then
'                        s(1) = &H0
'                        fSkip = True
'                    End If
'
'                    s(2) = left(Trim(s(3)), 2)
'
'                    If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then
'
'                        n = n + 1
''''
'                    End If
'                End If
'            End If
'        Else
'            If InStr(1, tmp, "<Data>") Then dataFlag = True
'        End If
'        'Loop
'    Next l

Case 3    'crk , dif
'    For l = 0 To UBound(Lines)
'        tmp = Lines(l)
'        'Do Until EOF(f)
'        '   Line Input #f, Tmp
'        If InStr(1, tmp, ": ") Then    '000037C8: 46 4B ;comment
'            s = Split(tmp, " ", 3)
'            If UBound(s) = 2 Then
'                s(0) = left(Trim(s(0)), Len(s(0)) - 1)
'                s(1) = Trim(s(1))
'
'                fSkip = False
'                If InStr(1, s(1), "null", vbTextCompare) Or InStr(1, s(1), "*") Then
'                    s(1) = &H0
'                    fSkip = True
'                End If
'
'
'                s(2) = left(Trim(s(2)), 2)
'                If IsNumeric("&H" & s(0)) And IsNumeric("&H" & s(1)) And IsNumeric("&H" & s(2)) Then
'                    n = n + 1
''''
'
'                End If
'            End If
'        End If
'        ' Loop
'    Next l

End Select

Close #f

'''
Exit Function
frmErr:
MsgBox Err.Description & ": CheckPatchParam()"
End Function


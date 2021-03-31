VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMagnifyBy 
      Height          =   315
      Left            =   2880
      TabIndex        =   20
      Text            =   "2"
      Top             =   1620
      Width           =   1935
   End
   Begin VB.TextBox txtResizeWidth 
      Height          =   315
      Left            =   2880
      TabIndex        =   18
      Text            =   "64"
      Top             =   2700
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "close"
      Height          =   315
      Left            =   5880
      TabIndex        =   16
      Top             =   2220
      Width           =   675
   End
   Begin VB.ComboBox cboWordsInLine 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1980
      Width           =   1935
   End
   Begin VB.ComboBox cboCheckCharSize 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   1935
   End
   Begin VB.CommandButton cmdApplyRestart 
      Caption         =   "Apply and Restart"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cboDither 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2340
      Width           =   1935
   End
   Begin VB.ComboBox cboMouse 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   900
      Width           =   1935
   End
   Begin VB.ComboBox cboMagnify 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1260
      Width           =   1935
   End
   Begin VB.FileListBox FileOptions 
      Height          =   1260
      Left            =   5880
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cboLang 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   1935
   End
   Begin VB.CommandButton cmdApplyOpt 
      Caption         =   "Apply "
      Default         =   -1  'True
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
      Left            =   120
      TabIndex        =   0
      Top             =   3180
      Width           =   2355
   End
   Begin VB.CommandButton cmdShowIni 
      Caption         =   "Show INI"
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   3180
      Width           =   2355
   End
   Begin VB.Label lblMagnifyBy 
      Alignment       =   1  'Right Justify
      Caption         =   "Magnify by"
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label lblResizeW 
      Alignment       =   1  'Right Justify
      Caption         =   "Quick resize to width"
      Height          =   315
      Left            =   60
      TabIndex        =   17
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblWordsInLine 
      Alignment       =   1  'Right Justify
      Caption         =   "Words in one column"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   2595
   End
   Begin VB.Label lblCheckCharSize 
      Alignment       =   1  'Right Justify
      Caption         =   "Check Char Size"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   2595
   End
   Begin VB.Label lblDither 
      Alignment       =   1  'Right Justify
      Caption         =   "Picture dithering method"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   2595
   End
   Begin VB.Label lblMouse 
      Alignment       =   1  'Right Justify
      Caption         =   "Mouse buttons mode"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   2595
   End
   Begin VB.Label lblMagnify 
      Alignment       =   1  'Right Justify
      Caption         =   "Magnify preview"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   2595
   End
   Begin VB.Label lblLanguage 
      Alignment       =   1  'Right Justify
      Caption         =   "Interface language"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2595
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private fNeedRestart As Boolean

Private Sub cmdApplyOpt_Click()
On Error GoTo frmErr

QwickResizeWidth = 64
If IsNumeric(txtResizeWidth.Text) Then
    If Val(txtResizeWidth.Text) < 257 And Val(txtResizeWidth.Text) > 0 Then QwickResizeWidth = Val(txtResizeWidth.Text)
End If

MagnifyBy = 3
If IsNumeric(txtMagnifyBy.Text) Then
    If Val(txtMagnifyBy.Text) < 33 And Val(txtMagnifyBy.Text) > 0 Then MagnifyBy = Val(txtMagnifyBy.Text)
End If

Call WriteOptIni
Call frmmain.reloadIni
If Len(lngFileName) <> 0 Then Call GetLanguage(3)

'1 or 2
'Call FillAllCombos
Unload Me

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdApplyOpt"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

'Private Sub cmdApplyRestart_Click()
'Unload Me
'Call RestartProg
'End Sub

Private Sub cmdShowIni_Click()
Shell "notepad.exe " & App.Path & "\VTCFont.ini", vbNormalFocus
End Sub

Private Sub Form_Load()
On Error GoTo frmErr

Me.Icon = frmmain.Icon
Me.Caption = "VTCFont: " & ArrMsg(35)

If Len(lngFileName) <> 0 Then Call GetLanguage(3)

Call FillAllCombos
txtResizeWidth.Text = QwickResizeWidth
txtMagnifyBy.Text = MagnifyBy

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Options_Load"
End Sub

Private Sub FillAllCombos()
With frmOptions
    For Each Contrl In .Controls
        If TypeOf Contrl Is ComboBox Then
            Contrl.Clear
        End If
    Next
End With

Call FillcboLang
Call FillcboMagnify
Call FillcboMouse
Call FillcboDither
Call FillcboCheckCharSize
Call FillcboWordsInLine
End Sub

Private Sub WriteOptIni()
On Error GoTo frmErr

'Dim tmpLanguage As String
Dim tmpMagnify As String

'tmpLanguage = right(cboLang.Text, 6)
'tmpLanguage = left(tmpLanguage, 2)

If cboMagnify.ListIndex = 0 Then    'yes
    Magnify = True
Else
    Magnify = False
End If

Select Case cboMouse.ListIndex
Case 0: InvertMouseB = 0
Case 1: InvertMouseB = 1
Case 2: InvertMouseB = 2
End Select

Select Case cboDither.ListIndex
Case 0: fPicDithered = 0
Case 1: fPicDithered = 1
Case 2: fPicDithered = 2
Case 3: fPicDithered = 3
End Select

If cboCheckCharSize.ListIndex = 0 Then    'yes
    CheckCharSizeFlag = True
Else
    CheckCharSizeFlag = False
End If

If cboWordsInLine.ListIndex = 0 Then    'yes
    AllWordsInLineFlag = True
Else
    AllWordsInLineFlag = False
End If

If Not NoIniFlag Then
    WriteKey "Global", "Language", cboLang.Text, iniFileName
    WriteKey "Global", "Magnify", CStr(Abs(CInt(Magnify))), iniFileName
    WriteKey "Global", "MagnifyBy", CStr(Abs(CInt(MagnifyBy))), iniFileName
    WriteKey "Global", "InvertMouseB", CStr(Abs(CInt(InvertMouseB))), iniFileName
    WriteKey "Global", "PicDithered", CStr(Abs(CInt(fPicDithered))), iniFileName
    WriteKey "Global", "CheckCharSize", CStr(Abs(CInt(CheckCharSizeFlag))), iniFileName
    WriteKey "Global", "WordsInLine", CStr(Abs(CInt(AllWordsInLineFlag))), iniFileName
    WriteKey "Global", "ResizeWidth", CStr(Abs(CInt(QwickResizeWidth))), iniFileName
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": WriteOptIni"
End Sub
Private Sub FillcboWordsInLine()
On Error GoTo frmErr

cboWordsInLine.AddItem ArrMsg(36)    'col
cboWordsInLine.AddItem ArrMsg(37)    'box
If AllWordsInLineFlag Then
    cboWordsInLine.Text = ArrMsg(36)
Else
    cboWordsInLine.Text = ArrMsg(37)
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillcboWordsInLine"
End Sub
Private Sub FillcboCheckCharSize()
On Error GoTo frmErr

cboCheckCharSize.AddItem ArrMsg(26)    'es
cboCheckCharSize.AddItem ArrMsg(27)    'no
If CheckCharSizeFlag Then
    cboCheckCharSize.Text = ArrMsg(26)
Else
    cboCheckCharSize.Text = ArrMsg(27)
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillcboCheckCharSize"
End Sub
Private Sub FillcboDither()
On Error GoTo frmErr

cboDither.AddItem ArrMsg(31)    'CustomShades
cboDither.AddItem ArrMsg(32)    'AtkinsonGS
cboDither.AddItem ArrMsg(33)    'AtkinsonBW
cboDither.AddItem ArrMsg(34)    'ShadesDithered
Select Case fPicDithered
Case 0: cboDither.Text = ArrMsg(31)
Case 1: cboDither.Text = ArrMsg(32)
Case 2: cboDither.Text = ArrMsg(33)
Case 3: cboDither.Text = ArrMsg(34)
End Select

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillcboDither"
End Sub
Private Sub FillcboMouse()
On Error GoTo frmErr

cboMouse.AddItem ArrMsg(28)    'normal
cboMouse.AddItem ArrMsg(29)    'Swap
cboMouse.AddItem ArrMsg(30)    'inverse
Select Case InvertMouseB
Case 0: cboMouse.Text = ArrMsg(28)
Case 1: cboMouse.Text = ArrMsg(29)
Case 2: cboMouse.Text = ArrMsg(30)
End Select

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillcboMouse"
End Sub
Private Sub FillcboMagnify()
On Error GoTo frmErr

cboMagnify.AddItem ArrMsg(26)    'es
cboMagnify.AddItem ArrMsg(27)    'no
If Magnify Then
    cboMagnify.Text = ArrMsg(26)
Else
    cboMagnify.Text = ArrMsg(27)
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillcboMagnify"
End Sub
Private Sub FillcboLang()
Dim i As Integer
Dim tmpLanguage As String

On Error GoTo frmErr

FileOptions.Path = App.Path
FileOptions.Pattern = "*.lng"

If FileOptions.ListCount <> 0 Then
    For i = 0 To FileOptions.ListCount - 1
        tmpLanguage = right$(FileOptions.List(i), 6)
        tmpLanguage = left$(tmpLanguage, 2)
        cboLang.AddItem tmpLanguage
    Next i
    cboLang.Text = Language    'lngFileNameOnly
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillcboLang"
End Sub

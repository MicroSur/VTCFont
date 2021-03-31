VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "VTCFont"
   ClientHeight    =   7635
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   12060
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   804
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVocabSL 
      Height          =   315
      Index           =   1
      Left            =   5400
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Export All words. Save data to text file"
      Top             =   1680
      Width           =   315
   End
   Begin VB.CommandButton cmdVocabSL 
      Height          =   315
      Index           =   0
      Left            =   5040
      Picture         =   "Form1.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Import All words. Open text data from file"
      Top             =   1680
      Width           =   315
   End
   Begin VB.ComboBox cmbVocAdr 
      Height          =   315
      Left            =   3420
      TabIndex        =   68
      Top             =   2100
      Width           =   975
   End
   Begin VB.CommandButton cmdResize 
      Height          =   315
      Left            =   6660
      Picture         =   "Form1.frx":0E1E
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Quick resize to Width specified in Options"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdFWUpdater 
      Caption         =   "FWUpdater"
      Height          =   375
      Left            =   5820
      TabIndex        =   16
      Top             =   2100
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "id test"
      Height          =   375
      Left            =   10800
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdShowAllDict 
      Height          =   315
      Left            =   2940
      Picture         =   "Form1.frx":13A8
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Show All words"
      Top             =   2100
      Width           =   375
   End
   Begin VB.CommandButton cmdUndoRedo 
      Height          =   315
      Index           =   1
      Left            =   6240
      Picture         =   "Form1.frx":1932
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Redo"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdUndoRedo 
      Height          =   315
      Index           =   0
      Left            =   5880
      Picture         =   "Form1.frx":1EBC
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Undo"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdINIShow 
      Height          =   315
      Index           =   1
      Left            =   5040
      Picture         =   "Form1.frx":2446
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "View readme file"
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picRectSel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8400
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   59
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkSelection 
      Height          =   315
      Left            =   1740
      Picture         =   "Form1.frx":29D0
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Selection tool"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdReloadFW 
      Height          =   315
      Left            =   4560
      Picture         =   "Form1.frx":2F5A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Reload current opened firmware"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   13
      Left            =   6060
      Picture         =   "Form1.frx":34E4
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Export. Save data to text file"
      Top             =   2640
      Width           =   315
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   12
      Left            =   5700
      Picture         =   "Form1.frx":3A6E
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Import. Open text data from file"
      Top             =   2640
      Width           =   315
   End
   Begin VB.CommandButton cmdPatcher 
      Caption         =   "Patch Lab"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Show patchers window"
      Top             =   600
      Width           =   1155
   End
   Begin VB.PictureBox PicScroll 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   1860
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   57
      Top             =   3600
      Width           =   2535
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   0
         ScaleHeight     =   2775
         ScaleWidth      =   2175
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.HScrollBar HScrollDraw 
      Height          =   255
      Left            =   2760
      TabIndex        =   56
      Top             =   6720
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.VScrollBar VScrollDraw 
      Height          =   2535
      Left            =   4860
      TabIndex        =   55
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicTTF 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   6540
      ScaleHeight     =   219
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   549
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   8235
      Begin VB.ComboBox cmbTTF_Char 
         Height          =   315
         Left            =   0
         TabIndex        =   64
         Text            =   "5"
         Top             =   420
         Width           =   2175
      End
      Begin VB.ComboBox cmb_SysFonts 
         Height          =   315
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   2175
      End
      Begin VB.CheckBox chkTTFBIU 
         Height          =   315
         Index           =   2
         Left            =   3000
         Picture         =   "Form1.frx":3FF8
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Underline"
         Top             =   420
         Width           =   315
      End
      Begin VB.CheckBox chkTTFBIU 
         Height          =   315
         Index           =   1
         Left            =   2640
         Picture         =   "Form1.frx":4582
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Italic"
         Top             =   420
         Width           =   315
      End
      Begin VB.CheckBox chkTTFBIU 
         Height          =   315
         Index           =   0
         Left            =   2280
         Picture         =   "Form1.frx":4B0C
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Bold"
         Top             =   420
         Width           =   315
      End
      Begin VB.VScrollBar VScroll_Y 
         Height          =   2415
         Left            =   360
         TabIndex        =   52
         Top             =   840
         Width           =   255
      End
      Begin VB.VScrollBar VScroll_S 
         Height          =   2415
         Left            =   0
         TabIndex        =   50
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdLoadFont 
         Caption         =   "Load TTF"
         Height          =   315
         Left            =   2280
         TabIndex        =   53
         ToolTipText     =   "Load TrueType font"
         Top             =   0
         Width           =   1515
      End
      Begin VB.HScrollBar HScroll_X 
         Height          =   255
         Left            =   780
         TabIndex        =   51
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblTTF 
         Height          =   2055
         Left            =   780
         TabIndex        =   54
         Top             =   1200
         Width           =   7125
      End
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   11
      Left            =   6780
      Picture         =   "Form1.frx":5096
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Open picture"
      Top             =   2640
      Width           =   315
   End
   Begin VB.ComboBox cmbVocab 
      Height          =   315
      ItemData        =   "Form1.frx":5620
      Left            =   1740
      List            =   "Form1.frx":5622
      TabIndex        =   14
      ToolTipText     =   "Vocabulary"
      Top             =   1680
      Width           =   3195
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   10
      Left            =   6420
      Picture         =   "Form1.frx":5624
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Use TrueType fonts"
      Top             =   2640
      Width           =   315
   End
   Begin VB.CommandButton cmdXY 
      Caption         =   "&Y"
      Height          =   315
      Index           =   1
      Left            =   3840
      TabIndex        =   24
      Top             =   2640
      Width           =   555
   End
   Begin VB.CommandButton cmdXY 
      Caption         =   "&X"
      Height          =   315
      Index           =   0
      Left            =   3300
      TabIndex        =   23
      Top             =   2640
      Width           =   555
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   9
      Left            =   4320
      Picture         =   "Form1.frx":5BAE
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Flip up-down"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   8
      Left            =   4680
      Picture         =   "Form1.frx":6138
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Flip left-right"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdViewWord 
      Height          =   315
      Left            =   2460
      Picture         =   "Form1.frx":66C2
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "View word"
      Top             =   2100
      Width           =   375
   End
   Begin VB.CheckBox chkByNumber 
      Appearance      =   0  'Flat
      Caption         =   "Paste by number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      ToolTipText     =   "Get char number from paste data, else from selection"
      Top             =   780
      Width           =   2055
   End
   Begin VB.PictureBox PicReal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000005&
      Height          =   315
      Left            =   7260
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picX3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   7260
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   48
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdINIShow 
      Height          =   315
      Index           =   0
      Left            =   5460
      Picture         =   "Form1.frx":6C4C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "View VTCFont *.ini file"
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdChar 
      Caption         =   "&nn"
      Height          =   315
      Left            =   1740
      TabIndex        =   17
      ToolTipText     =   "Add chars to word"
      Top             =   2100
      Width           =   615
   End
   Begin VB.CommandButton test 
      Caption         =   "test barr"
      Height          =   315
      Left            =   10800
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox tmpPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7860
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   7
      Left            =   5460
      Picture         =   "Form1.frx":71D6
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Clear"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   6
      Left            =   5100
      Picture         =   "Form1.frx":7760
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Inverse"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdSaveAll 
      Caption         =   "Save &All"
      Height          =   375
      Left            =   1740
      TabIndex        =   10
      ToolTipText     =   "Save changed chars to file"
      Top             =   1140
      Width           =   1755
   End
   Begin VB.CommandButton cmdLoadFile 
      Caption         =   "&Load"
      Height          =   375
      Left            =   1740
      TabIndex        =   0
      ToolTipText     =   "Load firmware file"
      Top             =   600
      Width           =   1755
   End
   Begin VB.CommandButton cmdSaveWord 
      Caption         =   "Save &word"
      Height          =   375
      Left            =   5820
      TabIndex        =   15
      ToolTipText     =   "Save current word"
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   5
      Left            =   3960
      Picture         =   "Form1.frx":8162
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Rotate right"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   4
      Left            =   3600
      Picture         =   "Form1.frx":86EC
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Rotate left"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdPasteCfont 
      Caption         =   "Paste C font"
      Height          =   315
      Left            =   10800
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   3
      Left            =   3240
      Picture         =   "Form1.frx":8C76
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Shift down"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   2
      Left            =   2880
      Picture         =   "Form1.frx":9200
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Shift up"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   1
      Left            =   2520
      Picture         =   "Form1.frx":978A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Shift right"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdToolBar 
      Height          =   315
      Index           =   0
      Left            =   2160
      Picture         =   "Form1.frx":9D14
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Shift left"
      Top             =   3120
      Width           =   315
   End
   Begin VB.CheckBox chkDupFont 
      Appearance      =   0  'Flat
      Caption         =   "Save to all blocks"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "Save Char to all blocks"
      Top             =   540
      Width           =   2055
   End
   Begin VB.ComboBox cmbHard 
      Height          =   315
      ItemData        =   "Form1.frx":A29E
      Left            =   1740
      List            =   "Form1.frx":A2A0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Load for this hardware"
      Top             =   120
      Width           =   2775
   End
   Begin VB.OptionButton optBlock 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   6
      ToolTipText     =   "Second Font block"
      Top             =   180
      Width           =   255
   End
   Begin VB.OptionButton optBlock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Blocks:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5940
      TabIndex        =   5
      ToolTipText     =   "Font blocks selection"
      Top             =   180
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      ToolTipText     =   "Paste this character"
      Top             =   1140
      Width           =   915
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      ToolTipText     =   "Copy this character"
      Top             =   1140
      Width           =   1155
   End
   Begin VB.ComboBox cmbAdr 
      Height          =   315
      ItemData        =   "Form1.frx":A2A2
      Left            =   4500
      List            =   "Form1.frx":A2A4
      TabIndex        =   25
      ToolTipText     =   "List of Fonts in block"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox chkGrid 
      Appearance      =   0  'Flat
      Caption         =   "Grid"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2460
      TabIndex        =   22
      Top             =   2640
      Value           =   1  'Checked
      Width           =   795
   End
   Begin VB.CommandButton cmdScale 
      Height          =   315
      Index           =   1
      Left            =   2040
      Picture         =   "Form1.frx":A2A6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Editor size +"
      Top             =   2640
      Width           =   315
   End
   Begin VB.CommandButton cmdScale 
      Height          =   315
      Index           =   0
      Left            =   1740
      Picture         =   "Form1.frx":A830
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Editor window size -"
      Top             =   2640
      Width           =   315
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save char"
      Height          =   375
      Left            =   5820
      TabIndex        =   13
      ToolTipText     =   "Save this character to file"
      Top             =   1140
      Width           =   1275
   End
   Begin VTCFont.McListBox McListBox1 
      Height          =   6135
      Left            =   120
      TabIndex        =   66
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   10821
      Picture         =   "Form1.frx":ADBA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StrechIcon      =   0   'False
      BorderStyle     =   0
      GridLines       =   0   'False
      BackGradient    =   1
   End
   Begin VB.Label lblWordSize 
      Height          =   255
      Left            =   4560
      TabIndex        =   45
      ToolTipText     =   "Word size x/y"
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Menu mnu_RealPic 
      Caption         =   "RealPicMenu"
      Visible         =   0   'False
      Begin VB.Menu mnu_SaveBMP 
         Caption         =   "Save to BMP real"
         Index           =   0
      End
      Begin VB.Menu mnu_SaveBMP 
         Caption         =   "Save to BMP increased"
         Index           =   1
      End
      Begin VB.Menu mnu_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_CopyPic 
         Caption         =   "Copy real to clipboard"
         Index           =   0
      End
      Begin VB.Menu mnu_CopyPic 
         Caption         =   "Copy enlarged to clipboard"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_MCList 
      Caption         =   "MCList"
      Visible         =   0   'False
      Begin VB.Menu mnu_Export 
         Caption         =   "Add to Export"
         Index           =   0
      End
      Begin VB.Menu mnu_Export 
         Caption         =   "Remove from Export"
         Index           =   1
      End
      Begin VB.Menu mnu_Export 
         Caption         =   "Select Added"
         Index           =   2
      End
      Begin VB.Menu mnu_Export 
         Caption         =   "Clear Export List"
         Index           =   3
      End
      Begin VB.Menu mnu_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_GoExport 
         Caption         =   "Export to file"
      End
      Begin VB.Menu mnu_UpdateExport 
         Caption         =   "Update an existing file"
      End
      Begin VB.Menu mnu_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SelectAll 
         Caption         =   "Select All"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'(C)MicroSuR 2015-2017
'Joyetech/Wismec/Eleaf/Vortex firmware resource editor

'Private Sub PicReal_Change() todo 2 раза

'конфликтуют пустые патчи

'чистить и не давать работать со словарем, если он остался с пред прошивки, а в текущем словаря нет
'что с пастом си шрифта
'drawword без использования piccontainer (долго)
'долгие сабы
''Call bArr2PicDraw(-1)    'fill draw box with bArr data
'

'Dim startTimer As Long
'startTimer = GetTickCount()
'Debug.Print GetTickCount() - startTimer

'note:
'vocab last addr is at the end of last word
'font last addr is at the start of last glyph

Option Explicit

Private m_oDoc As New XMLDocument
Private m_oCurrentElement As CXmlElement
Private arrXML_DataBody() As String
Private arrXML_ImageNum() As Integer
Private arrXML_ImageCol() As Long
Private arrXML_ImageRow() As Long
Private XML_Image_Count As Integer

Private XYspace As Single
Private lcBackColor As Long, lcForeColor As Long
Attribute lcForeColor.VB_VarUserMemId = 1073938440
Private bArr() As Byte    'draw array, (sRow,sCol)
Attribute bArr.VB_VarUserMemId = 1073938435
Private sCol As Long, sRow As Long    'array res
Attribute sCol.VB_VarUserMemId = 1073938436
Attribute sRow.VB_VarUserMemId = 1073938436
Private sRowArr() As Long    'for store it
Attribute sRowArr.VB_VarUserMemId = 1073938435
Private sColArr() As Long
Attribute sColArr.VB_VarUserMemId = 1073938436
Private FontData As New CString    'current 101011111010...
Attribute FontData.VB_VarUserMemId = 1073938438
Private FontDataArr() As String
Attribute FontDataArr.VB_VarUserMemId = 1073938437
Private sCharData As New CString    ''2,8,4,33,33,22
Attribute sCharData.VB_VarUserMemId = 1073938438
Private CharDataArr() As String    'from sCharData
Attribute CharDataArr.VB_VarUserMemId = 1073938439
Private CharDataHEXArr() As String 'store myevic format font data

Private arrExport() As Boolean    'all chars, true if char go to export
Private bTmp() As Byte    'decrypted block
Attribute bTmp.VB_VarUserMemId = 1073938441
Private xTmp() As Byte    'encrypted block
Attribute xTmp.VB_VarUserMemId = 1073938442
Private bTmpCollection() As Variant    ' array of bTmp()
Attribute bTmpCollection.VB_VarUserMemId = 1073938441
Private startAddr As Long    'start adr of char block in file
Attribute startAddr.VB_VarUserMemId = 1073938443

Private oldX As Integer, oldY As Integer
Attribute oldX.VB_VarUserMemId = 1073938443
Attribute oldY.VB_VarUserMemId = 1073938443
Private oldColor As Long
Attribute oldColor.VB_VarUserMemId = 1073938437

Private FontBlock1IndArr() As String    'addresses of pointers to fonts
Private FontBlock2IndArr() As String
Private FontBlock1Arr() As String    '3 bytes hex, combo list
Attribute FontBlock1Arr.VB_VarUserMemId = 1073938438
Private FontBlock2Arr() As String
Attribute FontBlock2Arr.VB_VarUserMemId = 1073938439

Private VortexMod As Boolean 'flag if current fw is vortex
Private FontBlock1VortexWidthArr() As Byte 'width of every glyphs
Private FontBlock2VortexWidthArr() As Byte 'do not change after fill
Private VortexBlock1Height As Integer
Private VortexBlock2Height As Integer 'do not change after fill

Private NoVocabFlag As Boolean 'no vocab
Private VocBlock1Arr() As String    'hex, combo list
Attribute VocBlock1Arr.VB_VarUserMemId = 1073938444
Private VocBlock2Arr() As String
Attribute VocBlock2Arr.VB_VarUserMemId = 1073938445
Private VocBlock1Arr0x() As String
Private VocBlock2Arr0x() As String
Private Word1StartArr() As Long  '(addr) of word from fw
Attribute Word1StartArr.VB_VarUserMemId = 1073938446
Private Word2StartArr() As Long
Attribute Word2StartArr.VB_VarUserMemId = 1073938447
Private Word1LenArr() As Integer    '(length of word)
Attribute Word1LenArr.VB_VarUserMemId = 1073938448
Private Word2LenArr() As Integer
Attribute Word2LenArr.VB_VarUserMemId = 1073938449
Private WordCharNumBytes As Integer    ' 1 byte or 2 bytes
Private CurrentWordInd As Integer    'cmb word listindex
Attribute CurrentWordInd.VB_VarUserMemId = 1073938450
Private cmbLastIndex As Integer
Attribute cmbLastIndex.VB_VarUserMemId = 1073938440
Private sColPos As Integer
Attribute sColPos.VB_VarUserMemId = 1073938451
Private sRowPos As Integer    ' to shift chars down2 in DrawWord
Attribute sRowPos.VB_VarUserMemId = 1073938452
Private sRowMax As Integer
Attribute sRowMax.VB_VarUserMemId = 1073938453
Private sColMax As Integer

Private CurrentChar As String    'hex, 01 02..
Attribute CurrentChar.VB_VarUserMemId = 1073938454
Private ShiftChar() As String    'from ini
Attribute ShiftChar.VB_VarUserMemId = 1073938455
Private selArr() As Integer    ' array of selection in MCL, 0 not used
Attribute selArr.VB_VarUserMemId = 1073938456
'inmodule Private Hardtext As String    'store fo current file open
Private VocSelStart As Integer
Attribute VocSelStart.VB_VarUserMemId = 1073938457
Private TTFontName As String    'for TTF
Attribute TTFontName.VB_VarUserMemId = 1073938458
Private TTFontBold As Boolean
Attribute TTFontBold.VB_VarUserMemId = 1073938459
Private TTFontItalic As Boolean
Attribute TTFontItalic.VB_VarUserMemId = 1073938460
Private TTFontUnderline As Boolean
Attribute TTFontUnderline.VB_VarUserMemId = 1073938461
Private TTF_X As Single
Attribute TTF_X.VB_VarUserMemId = 1073938462
Private TTF_Y As Single
Attribute TTF_Y.VB_VarUserMemId = 1073938463
Private TTF_Size As Integer
Attribute TTF_Size.VB_VarUserMemId = 1073938464
Private TTF_Char As String
Attribute TTF_Char.VB_VarUserMemId = 1073938465
Private m_Preview As CFontPreview
Attribute m_Preview.VB_VarUserMemId = 1073938466
'Private VScroll_Y_Value As Integer
'Private HScroll_X_Value As Integer
'Private VScroll_S_Value As Integer

Private FirstPasteByNum As Integer    'to point in mclist
Attribute FirstPasteByNum.VB_VarUserMemId = 1073938467
Private start_sel_X As Integer  'stores the coord of the mousedown event
Attribute start_sel_X.VB_VarUserMemId = 1073938468
Private start_sel_Y As Integer
Attribute start_sel_Y.VB_VarUserMemId = 1073938469
Private old_X1 As Integer
Attribute old_X1.VB_VarUserMemId = 1073938470
Private old_Y1 As Integer
Attribute old_Y1.VB_VarUserMemId = 1073938471
Private old_X2 As Integer
Attribute old_X2.VB_VarUserMemId = 1073938472
Private old_Y2 As Integer
Attribute old_Y2.VB_VarUserMemId = 1073938473
Private old_xx1 As Integer
Attribute old_xx1.VB_VarUserMemId = 1073938474
Private old_yy1 As Integer
Attribute old_yy1.VB_VarUserMemId = 1073938475
Private old_xx2 As Integer
Attribute old_xx2.VB_VarUserMemId = 1073938476
Private old_yy2 As Integer
Attribute old_yy2.VB_VarUserMemId = 1073938477
Private old_Xm As Integer
Attribute old_Xm.VB_VarUserMemId = 1073938478
Private old_Ym As Integer
Attribute old_Ym.VB_VarUserMemId = 1073938479
Private LineWidth As Integer
Attribute LineWidth.VB_VarUserMemId = 1073938480

Private isSelection As Boolean
Attribute isSelection.VB_VarUserMemId = 1073938481
Private fMoveSelRect As Boolean
Attribute fMoveSelRect.VB_VarUserMemId = 1073938482
Private fCopySelection As Boolean
Attribute fCopySelection.VB_VarUserMemId = 1073938483
Private fMoveSelection As Boolean
Attribute fMoveSelection.VB_VarUserMemId = 1073938484
'Private Selecting As Boolean
Private X1Region As Integer
Attribute X1Region.VB_VarUserMemId = 1073938485
Private Y1Region As Integer
Attribute Y1Region.VB_VarUserMemId = 1073938486
Private X2Region As Integer
Attribute X2Region.VB_VarUserMemId = 1073938487
Private Y2Region As Integer
Attribute Y2Region.VB_VarUserMemId = 1073938488

Private UndoBuffer() As StdPicture
Attribute UndoBuffer.VB_VarUserMemId = 1073938489
Private picCount As Integer
Attribute picCount.VB_VarUserMemId = 1073938490
Private MaxUndoCircle As Integer    '4
Attribute MaxUndoCircle.VB_VarUserMemId = 1073938491
Private UndoClicksCount As Integer
Attribute UndoClicksCount.VB_VarUserMemId = 1073938492
'Private UndoFlag As Boolean

Private IDpassword() As String    'identify FW, array all HW
Attribute IDpassword.VB_VarUserMemId = 1073938493
Private IDsecuredAdr() As Long    'identify FW, array all HW
Private IDsecuredByte() As Byte    'identify FW, array all HW
Private ISVortex() As Boolean

Private IDanswer(2) As Byte    ', 80 51 01
Attribute IDanswer.VB_VarUserMemId = 1073938494
Private IDVortex(2) As Byte    'VOR tex in IDsecured,

Private NoBlock2Flag As Boolean
Attribute NoBlock2Flag.VB_VarUserMemId = 1073938441
Private Block1Flag As Boolean    ' true to show font for last block in FW (old screen), false for first
Attribute Block1Flag.VB_VarUserMemId = 1073938443
Private AllWordsShow As Boolean 'flag all word show


Private AllWordsStartCoord_X() As Integer    'coords of all words show
Private AllWordsStartCoord_Y() As Integer
Private AllWordsFinishCoord_X() As Integer
Private AllWordsFinishCoord_Y() As Integer
Private oldAllWordsStartCoord_X As Integer
'Private oldAllWordsStartCoord_Y As Integer
Private oldAllWordsFinishCoord_X As Integer
Private oldAllWordsFinishCoord_Y As Integer
Private AllWordsInd As Integer

Private DrawWordFlag As Boolean
Attribute DrawWordFlag.VB_VarUserMemId = 1073938497
Private DrawAllWordsFlag As Boolean
Private NoRealDraw As Boolean    'no realmap this
Attribute NoRealDraw.VB_VarUserMemId = 1073938498
'Private FillMCListFlag As Boolean '
Private MCListFilling As Boolean '
Attribute MCListFilling.VB_VarUserMemId = 1073938499
Private PasteByNumber As Boolean    'current index from file or selection
Attribute PasteByNumber.VB_VarUserMemId = 1073938500

Private noResize As Boolean    ' no res in batch save
Private noChangeBlockFlag As Boolean    'changing block not allow
Private GoExportFlag As Boolean


'Private Const MagnifyBy As Integer = 2
'Private MagnifyBy As Integer
Private EscFlag As Boolean
Attribute EscFlag.VB_VarUserMemId = 1073938502
Private shiftFlag As Boolean
Attribute shiftFlag.VB_VarUserMemId = 1073938503
Private chkGridFlag As Integer
Attribute chkGridFlag.VB_VarUserMemId = 1073938504
Private LoadFontFlag As Boolean
Attribute LoadFontFlag.VB_VarUserMemId = 1073938505

Private AlwaysHex As Boolean 'in x_y

Private pngClass As New LoadPNG
Attribute pngClass.VB_VarUserMemId = 1073938507

Private Vocab1Start As Long 'from ini
Private Vocab2Start As Long
Private Vocab1End As Long
Private Vocab2End As Long
'Private oldPicRealH As Long 'for picreal no change
'Private oldPicRealW As Long

Private redGridFlag As Boolean



Private Sub chkByNumber_Click()
PasteByNumber = chkByNumber.Value
End Sub

Private Sub GetArray(ByRef SelInd As Integer)
'bArr from FontDataArr
On Error GoTo frmErr
'Debug.Print "> GetArray"

sCol = sColArr(SelInd)
sRow = sRowArr(SelInd)
FontData.reset
FontData.concat FontDataArr(SelInd)

Call FileFontData2bArr(SelInd)

If MCListFilling Then
    Call bArr2PicReal(SelInd)
Else
'  If Not DrawWordFlag Then Call bArr2PicDraw(SelInd)

    If Not NoRealDraw Then
        If DrawWordFlag Then
            Call PicDraw2PicReal(SelInd)    'and draw word too
        Else
            Call bArr2PicDraw(SelInd)
            Call bArr2PicReal(SelInd)
        End If
    End If

End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GetArray()"
'lblChar = vbNullString
cmdChar.Caption = "&nn"

End Sub


Private Sub GetArray_Vortex(ByRef SelInd As Integer)
'bArr from FontDataArr
On Error GoTo frmErr
'Debug.Print "> GetArray"

sCol = sColArr(SelInd)
sRow = sRowArr(SelInd)
FontData.reset
FontData.concat FontDataArr(SelInd)

Call FileFontData2bArr_Vortex(SelInd)

If MCListFilling Then
    Call bArr2PicReal(SelInd)
Else
'  If Not DrawWordFlag Then Call bArr2PicDraw(SelInd)

    If Not NoRealDraw Then
        If DrawWordFlag Then
            Call PicDraw2PicReal(SelInd)    'and draw word too
        Else
            Call bArr2PicDraw(SelInd)
            Call bArr2PicReal(SelInd)
        End If
    End If

End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GetArray_Vortex"
'lblChar = vbNullString
cmdChar.Caption = "&nn"

End Sub

Private Sub GetBlock(ByRef SelInd As Integer)
'from file
'filesize + uMagic + index - filesize \ uMagic
'data[i] = (self.data[i] ^ self._genfun(len(self.data), i)) & 0xFF
Dim i As Long
Dim Ind As Long    'pos in file
Dim n As Long
Dim sRowCurrent As Long
Dim sColCurrent As Long
On Error GoTo frmErr
'Debug.Print "> GetBlock"

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If

'ReDim bTmp(&H1000)    'todo
'ReDim xTmp(&H1000)
ReDim bTmp(1)
'ReDim xTmp(1)

Ind = startAddr
Seek bFileIn, Ind + 1
Get #bFileIn, , xTmp()

'get size
For i = 0 To 1
    bTmp(i) = (xTmp(i) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
    Ind = Ind + 1
Next i

sColCurrent = bTmp(0): sRowCurrent = bTmp(1)

If startAddr = 0 Then
    sColCurrent = 1: sRowCurrent = 1
ElseIf sColCurrent = 0 Or sColCurrent > 96 Then 'correct errors
    sColCurrent = 1: sRowCurrent = 1
ElseIf sRowCurrent = 0 Or sRowCurrent > 128 Then
    sColCurrent = 1: sRowCurrent = 1
End If


sRowArr(SelInd) = sRowCurrent    'before FileFontData2bArr
sColArr(SelInd) = sColCurrent

If Block1Flag Then
    'If sRowCurrent < 8 Then sRowCurrent = 8
    ' ReDim Preserve bTmp(sColCurrent * sRowCurrent / 8 + 1)
    'n = sRowCurrent \ 8
    n = intBytes(sRowCurrent)
    'ReDim Preserve bTmp(sColCurrent * n / 8 + 1)
    ReDim bTmp(n * sColCurrent / 8 + 1)
Else
    n = intBytes(sColCurrent)
    'ReDim Preserve bTmp(n * sRowCurrent / 8 + 1)
    ReDim bTmp(n * sRowCurrent / 8 + 1)
End If

'get all bytes (6  8  0  62  65  65  62  0 )
ReDim xTmp(UBound(bTmp)) ' IsRavage = + 1)
Ind = startAddr
Seek bFileIn, Ind + 1
Get #bFileIn, , xTmp()

' IsRavage Dim r As Integer

For i = 0 To UBound(xTmp)

' IsRavage If i <> 2 Then
' IsRavage bTmp(r)
    bTmp(i) = (xTmp(i) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
    ' IsRavage r = r + 1
' IsRavage End If
    Ind = Ind + 1
    
Next i

''''''''''''''''''''''''''''''''''
bTmpCollection(SelInd) = bTmp

Call bTmp2FontData(SelInd)
Call FileFontData2bArr(SelInd)

Call bArr2CharData(SelInd)    'for copy-paste

If MCListFilling Then
    Call bArr2PicReal(SelInd)
Else
    'If cmbLastIndex = SelInd Then
    If Not DrawWordFlag Then Call bArr2PicDraw(SelInd)
    If Not NoRealDraw Then
        If DrawWordFlag Then
            Call PicDraw2PicReal(SelInd)    'and draw word too
        Else
            Call bArr2PicReal(SelInd)
        End If
    End If
    'End If
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GetBlock()"
'lblChar = vbNullString
cmdChar.Caption = "&nn"
End Sub


Private Sub GetBlock_Vortex(ByRef SelInd As Integer)
'from file
Dim i As Long
Dim Ind As Long    'pos in file
Dim n As Long
Dim sRowCurrent As Long
Dim sColCurrent As Long
On Error GoTo frmErr
'Debug.Print "> GetBlock_V"

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If

'ReDim bTmp(1)
'
'Ind = startAddr
'Seek bFileIn, Ind + 1
'Get #bFileIn, , xTmp()
'
''get size
'For i = 0 To 1
'    bTmp(i) = (xTmp(i) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
'    Ind = Ind + 1
'Next i
'
'sColCurrent = bTmp(0): sRowCurrent = bTmp(1)
'
'If startAddr = 0 Then
'    sColCurrent = 1: sRowCurrent = 1
'ElseIf sColCurrent = 0 Or sColCurrent > 96 Then 'correct errors
'    sColCurrent = 1: sRowCurrent = 1
'ElseIf sRowCurrent = 0 Or sRowCurrent > 128 Then
'    sColCurrent = 1: sRowCurrent = 1
'End If


sRowArr(SelInd) = sRowCurrent    'before FileFontData2bArr
sColArr(SelInd) = sColCurrent

If Block1Flag Then

sRowArr(SelInd) = VortexBlock1Height
sColArr(SelInd) = FontBlock1VortexWidthArr(SelInd)
'sRowArr(SelInd) = FontBlock1VortexWidthArr(SelInd)
'sColArr(SelInd) = VortexBlock1Height

'    n = intBytes(sRowArr(SelInd))
'    ReDim bTmp(sColArr(SelInd) * n / 8 + 1)
    n = intBytes(sColArr(SelInd))
    ReDim bTmp(n * sRowArr(SelInd) / 8 - 1)
    
Else

'sRowArr(SelInd) = FontBlock2VortexWidthArr(SelInd)
'sColArr(SelInd) = VortexBlock2Height
sRowArr(SelInd) = VortexBlock2Height
sColArr(SelInd) = FontBlock2VortexWidthArr(SelInd)

    n = intBytes(sRowArr(SelInd))
    ReDim bTmp(n * sColArr(SelInd) / 8 - 1)
 '   n = intBytes(sColArr(SelInd))
  '  ReDim bTmp(n * sRowArr(SelInd) / 8 - 1)
End If



'get all bytes
'ReDim xTmp(UBound(bTmp))
Ind = startAddr + 1
Get #bFileIn, Ind, bTmp()

''''''''''''''''''''''''''''''''''
bTmpCollection(SelInd) = bTmp
Call bTmp2FontData_Vortex(SelInd)
Call FileFontData2bArr_Vortex(SelInd)
Call bArr2CharData(SelInd)    'for copy-paste

If MCListFilling Then
    Call bArr2PicReal(SelInd)
Else
    'If cmbLastIndex = SelInd Then
    If Not DrawWordFlag Then Call bArr2PicDraw(SelInd)
    If Not NoRealDraw Then
        If DrawWordFlag Then
            Call PicDraw2PicReal(SelInd)    'and draw word too
        Else
            Call bArr2PicReal(SelInd)
        End If
    End If
    'End If
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GetBlock_Vortex"
'lblChar = vbNullString
cmdChar.Caption = "&nn"
End Sub



Private Sub PicDraw2PicReal(ByRef SelInd As Integer)
'call picContainerMappingTo with param
'and draw word
Dim i As Long
Dim sRowCurrent As Integer
Dim sColCurrent As Integer
Dim shiftFlag As Boolean

On Error GoTo frmErr

If SelInd = -1 Then    'no file
    sRowCurrent = sRow
    sColCurrent = sCol
Else
    sRowCurrent = sRowArr(SelInd)
    sColCurrent = sColArr(SelInd)
End If

If DrawWordFlag Then
'in call   sRowPos = 0
    For i = 0 To UBound(ShiftChar)
        If CurrentChar = ShiftChar(i) Then
            sRowPos = sRowPos + 2    ' todo shift down in ini?
            shiftFlag = True
            Exit For
        End If
    Next i
    
    
'sRowPos sColPos uses in picContainerMappingTo
    'sRowMax = IIf(sRowMax > sRowCurrent, sRowMax, sRowCurrent)
    If sRowMax <= sRowCurrent Then sRowMax = sRowCurrent
    'sColMax = IIf(sColCurrent + sColPos > sColMax, sColCurrent + sColPos, sColMax)
    If sColCurrent + sColPos > sColMax Then sColMax = sColCurrent + sColPos

    With PicReal                ' 10 word with black frame 5-5
        If DrawAllWordsFlag Then
            .Width = sColMax + 10
        Else
            .Width = sColCurrent + sColPos + 10
        End If

        .Height = sRowMax + 10

    End With

    picContainerMappingTo PicReal, sColCurrent, sRowCurrent, 5, 5  'word to real bitmap


    sColPos = sColPos + sColCurrent

    If VortexMod Then
        sColPos = sColPos + 2
    End If
    
    If shiftFlag Then sRowPos = sRowPos - 2    'back for draw all words

Else
    sRowPos = 0
    sColPos = 0
    With PicReal
        .Width = sColCurrent
        .Height = sRowCurrent
        .Picture = Nothing    'need
    End With
    picContainerMappingTo PicReal, sColCurrent, sRowCurrent    'to real bitmap
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": PicDraw2PicReal()"
End Sub
Private Sub bArr2PicReal(ByRef SelInd As Integer)
Dim i As Long, j As Long
Dim sRowCurrent As Integer
Dim sColCurrent As Integer
Dim aTmp() As Long
Dim X As Long, Y As Long

On Error GoTo frmErr

'Debug.Print ">   bArr2PicReal"
If SelInd = -1 Then    'no file
    sRowCurrent = sRow
    sColCurrent = sCol
Else
    sRowCurrent = sRowArr(SelInd)
    sColCurrent = sColArr(SelInd)
    'sRowCurrent = sColArr(SelInd)
    'sColCurrent = sRowArr(SelInd)
End If

With PicReal
    .Width = sColCurrent
    .Height = sRowCurrent
    .Picture = Nothing
End With

ReDim aTmp(sColCurrent, sRowCurrent)
X = sRowCurrent - 1
For j = 0 To sColCurrent - 1
    For i = 0 To sRowCurrent - 1
        If bArr(i, j) = 1 Then
            aTmp(Y, X) = lcBackColor     'inverse color
'Else
'    aTmp(Y, X) = lcForeColor
        End If
        X = X - 1
    Next i
    X = sRowCurrent - 1: Y = Y + 1
Next j
Call SetBitmapData(PicReal.hdc, sColCurrent + 1, sRowCurrent, VarPtr(aTmp(0, 0)))

PicReal.Picture = PicReal.Image

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": bArr2PicReal()"
End Sub

Private Sub WriteBlock(adr As Long, SelInd)
'char 2 file
Dim i As Long
Dim UBbTmp As Integer
On Error GoTo frmErr
'Debug.Print ">   WriteBlock"

If adr = 0 Then Exit Sub

UBbTmp = UBound(bTmpCollection(SelInd))
ReDim xTmp(UBbTmp)

'bTmpCollection(SelInd)(i) => btmp(i)
Seek bFileIn, adr + 1

If VortexMod Then

    For i = 0 To UBbTmp
        xTmp(i) = bTmpCollection(SelInd)(i)
    Next i

Else

    For i = 0 To UBbTmp
        xTmp(i) = (bTmpCollection(SelInd)(i) Xor (adr + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        adr = adr + 1
    Next i
End If

Put #bFileIn, , xTmp()

If VortexMod Then
    'write new width
    If Block1Flag Then
        If FontBlock1VortexWidthArr(SelInd) <> sColArr(SelInd) Then
            Put #bFileIn, FontBlock1IndArr(SelInd), CByte(sColArr(SelInd))
            FontBlock1VortexWidthArr(SelInd) = sColArr(SelInd)
        End If
        
    Else
        If FontBlock2VortexWidthArr(SelInd) <> sColArr(SelInd) Then
            Put #bFileIn, FontBlock2IndArr(SelInd), CByte(sColArr(SelInd))
            FontBlock2VortexWidthArr(SelInd) = sColArr(SelInd)
        End If

    End If
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": WriteBlock()"
End Sub


Private Sub chkGrid_Click()
chkGridFlag = chkGrid.Value
If XYspace = 1 Then Exit Sub
Call bArr2PicDraw(-1)
End Sub

Private Sub chkGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then redGridFlag = Not redGridFlag
Call bArr2PicDraw(-1)
'DrawGrid picContainer, vbGrayText
End Sub

Private Sub chkSelection_Click()
On Error GoTo frmErr
If fFileOpen Then

    If VortexMod Then
        Call GetArray_Vortex(cmbLastIndex)    'if word show
    Else
        Call GetArray(cmbLastIndex)    'if word show
    End If

    
End If
isSelection = False    '1
Call bArr2PicDraw(-1)    '2
X1Region = 0: Y1Region = 0: X2Region = 0: Y2Region = 0
old_X1 = 0: old_Y1 = 0
'''
Exit Sub
frmErr:
MsgBox Err.Description & ": chkSelection_Click()"
End Sub

Private Sub chkTTFBIU_Click(Index As Integer)
Select Case Index
Case 0
    TTFontBold = chkTTFBIU(0).Value
Case 1
    TTFontItalic = chkTTFBIU(1).Value
Case 2
    TTFontUnderline = chkTTFBIU(2).Value
End Select
Call cmbTTF_Char_Click
End Sub

Private Sub cmb_SysFonts_Change()
Call cmb_SysFonts_Click
End Sub

Private Sub cmb_SysFonts_Click()
On Error Resume Next

TTFontName = cmb_SysFonts.Text
'TTFontBold = m_Preview.Bold
'TTFontItalic = m_Preview.Italic
'TTFontUnderline = m_Preview.Underlined
TTF_Char = cmbTTF_Char.Text


LoadFontFlag = True
VScroll_Y.Max = 150
VScroll_Y.Min = -350
HScroll_X.Max = 250
HScroll_X.Min = -200
VScroll_S.Max = 1
VScroll_S.Min = 500

'VScroll_Y.Value = VScroll_Y_Value '0
'HScroll_X.Value = HScroll_X_Value '0
' VScroll_S.Value = 250

'TTF_Size = sCol * XYspace
'If TTF_Size < VScroll_S.Min Then VScroll_S.Value = TTF_Size       '250
TTF_Size = VScroll_S.Value

LoadFontFlag = False

Call cmbTTF_Char_Click

On Error GoTo 0
End Sub


Private Sub cmb_SysFonts_DropDown()
SetDropdownHeight cmb_SysFonts, ScaleHeight / 2
End Sub

Private Sub cmbAdr_Click()
On Error GoTo frmErr

cmbLastIndex = cmbAdr.ListIndex

If VortexMod Then
    Call GetArray_Vortex(cmbLastIndex)
Else
    Call GetArray(cmbLastIndex)
End If

If Not MCListFilling Then

    'no scroll mouse wheel McListBox1.Visible = False

    McListBox1.ClearSelectionAll
    McListBox1.SelectItem (cmbLastIndex)

    If UBound(selArr) <= 1 Then    'sel first
        ReDim selArr(1)
        selArr(1) = cmbLastIndex
    End If
    ' McListBox1.Visible = True

    McListBox1.Refresh    ' need for scroll in cmbadr


End If

'lblChar = Right("0" & Hex(CStr(cmbAdr.ItemData(cmbLastIndex))), 2)    'get char number
'cmdChar.Caption = Right("0" & Hex(CStr(cmbAdr.ItemData(cmbLastIndex))), 2)    '& " ^"

'If AlwaysHex And Not VortexMod Then
If AlwaysHex Then
    cmdChar.Caption = "H" & Hex(CStr(cmbAdr.ItemData(cmbLastIndex)))
Else
    cmdChar.Caption = "#" & CStr(cmbAdr.ItemData(cmbLastIndex))
End If

Call XYcaptionSet(sCol, sRow)

If Block1Flag Then
    startAddr = "&H" & FontBlock1Arr(cmbLastIndex)
    cmbAdr.ToolTipText = Hex(FontBlock1IndArr(cmbLastIndex))
Else
    startAddr = "&H" & FontBlock2Arr(cmbLastIndex)
    cmbAdr.ToolTipText = Hex(FontBlock2IndArr(cmbLastIndex))
End If

cmdCopy.Caption = ArrMsg(13) & " (" & McListBox1.SelCount & ")"

'Call Form_Resize
Call SetUpPicScroll
Call SetUpScrollBars

If PicTTF.Visible Then
    PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2
End If

Call UndoBufferClear
Call StoreInUndoBuffer

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmbAdr_Click()"
End Sub

Private Sub cmbAdr_DropDown()
SetDropdownHeight cmbAdr, ScaleHeight
End Sub

Private Sub cmbAdr_GotFocus()
MouseUpBug = True 'not select in font list with mouse move from here
End Sub

'Private Sub cmbAdr_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error GoTo frmErr
'
''not work now
''If Not fFileOpen Then Exit Sub
''If Not IsNumeric("&H" & cmbAdr.Text) Then Exit Sub
''
''If KeyCode = 13 Then
''    startAddr = "&H" & cmbAdr.Text
''    cmbLastIndex = cmbAdr.ListIndex
'''Call GetBlock(cmbLastIndex)
''    Call GetArray(cmbLastIndex)
''    Call XYcaptionSet(sCol, sRow)
''End If
'
''''
'Exit Sub
'frmErr:
'MsgBox Err.Description & ": cmbAdr_KeyDown()"
'End Sub


Private Sub cmbHard_DropDown()
SetDropdownHeight cmbHard, ScaleHeight
End Sub

Private Sub SetDropdownHeight(cbo As ComboBox, ByVal max_extent As Integer)
' Adjust height of combobox dropdown part; call in response to DropDown event
' max_extent is the absolute maximum clientY value that the dropdown may extend to
' case 1: nItems <= 8 : do nothing - vb standard behaviour
' case 2: Items will fit in defined max area : resize to fit
' case 3: Items will not fit : resize to defined max height
On Error GoTo frmErr

If cbo.ListCount > 8 Then
    Dim max_fit As Integer    ' maximum number of items that will fit in maximum extent
    Dim item_ht As Integer    ' Calculated height of an item in the dropdown

    item_ht = ScaleY(cbo.Height, ScaleMode, vbPixels) - 8
    max_fit = (max_extent - cbo.top - cbo.Height) \ ScaleY(item_ht, vbPixels, ScaleMode)

    If cbo.ListCount <= max_fit Then
        MoveWindow cbo.hWnd, ScaleX(cbo.left, ScaleMode, vbPixels), _
                   ScaleY(cbo.top, ScaleMode, vbPixels), _
                   ScaleX(cbo.Width, ScaleMode, vbPixels), _
                   ScaleY(cbo.Height, ScaleMode, vbPixels) + (item_ht * cbo.ListCount) + 2, 0
    Else
        MoveWindow cbo.hWnd, ScaleX(cbo.left, ScaleMode, vbPixels), _
                   ScaleY(cbo.top, ScaleMode, vbPixels), _
                   ScaleX(cbo.Width, ScaleMode, vbPixels), _
                   ScaleY(cbo.Height, ScaleMode, vbPixels) + (item_ht * max_fit) + 2, 0
    End If
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": SetDropdownHeight()"
End Sub

Private Sub cmbTTF_Char_Change()
Call cmbTTF_Char_Click
End Sub

Private Sub cmbTTF_Char_Click()
On Error GoTo frmErr

If LoadFontFlag Then Exit Sub
If Len(TTFontName) = 0 Then Exit Sub

TTF_Char = cmbTTF_Char.Text
TTFontDraw TTF_Char, TTF_Size, TTF_X, TTF_Y, TTFontBold, TTFontItalic, TTFontUnderline

lblTTF.Font.Name = TTFontName
lblTTF.FontSize = 36
lblTTF.FontBold = TTFontBold
lblTTF.FontItalic = TTFontItalic
lblTTF.FontUnderline = TTFontUnderline
'lblTTF.FontStrikethru
lblTTF = TTF_Char

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmbTTF_Char_Click()"
End Sub

Private Sub cmbTTF_Char_DropDown()
SetDropdownHeight cmbTTF_Char, ScaleHeight / 2
End Sub

Private Sub cmbTTF_Char_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call cmbTTF_Char_Click
End Sub

Private Sub cmbVocab_Change()
'for check only
Dim arr() As String
Dim Tmp As String
On Error GoTo frmErr

If Not fFileOpen Then Exit Sub

Tmp = cmbVocab.Text
Do While InStr(Tmp, "  ")
    Tmp = Replace(Tmp, "  ", mySpace)
Loop
Tmp = Trim(Tmp)

arr = Split(Tmp, mySpace)

'lblWordSize = cmbVocab.ItemData(CurrentWordInd) & ", " & lblWordSize

If VortexMod And (Not Block1Flag) Then Exit Sub
If NoVocabFlag Then Exit Sub

If UBound(arr) <> cmbVocab.ItemData(CurrentWordInd) - 1 Then
    lblWordSize.ForeColor = vbRed
Else
    lblWordSize.ForeColor = &H80000012
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmbVocab_Change()"
End Sub

Private Sub cmbVocab_Click()
Dim arr() As String
Dim s As String
'Dim bTmp As Boolean

On Error GoTo frmErr


PicReal.Picture = Nothing
CurrentWordInd = cmbVocab.ListIndex
cmbVocAdr.ListIndex = CurrentWordInd

s = cmbVocab.Text
'in drawword
'Do While InStr(s, "  ")
'    s = Replace(s, "  ", mySpace)
'Loop
's = Trim(s)

'bTmp = PicTTF.Visible
'PicTTF.Visible = False

Call DrawWord(s)

'PicTTF.Visible = bTmp

arr = Split(s, mySpace)
lblWordSize = UBound(arr) + 1 & " (" & cmbVocab.ItemData(CurrentWordInd) & "), " & lblWordSize
lblWordSize.ForeColor = &H80000012
VocSelStart = Len(cmbVocab.Text)

'If Block1Flag Then
'    'lblVocAdr.Caption = Hex(Word1StartArr(CurrentWordInd))
'    cmbVocAdr.Text = Hex(Word1StartArr(CurrentWordInd))
'
'Else
'    'lblVocAdr.Caption = Hex(Word2StartArr(CurrentWordInd))
'    cmbVocAdr.Text = Hex(Word2StartArr(CurrentWordInd))
'End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmbVocab_Click()"
End Sub

Private Sub cmbVocab_DropDown()
SetDropdownHeight cmbVocab, ScaleHeight
End Sub

Private Sub cmbVocab_GotFocus()
MouseUpBug = True 'not select in font list with mouse move from here
End Sub

Private Sub cmbVocab_KeyPress(KeyAscii As Integer)
Dim arr() As String
Dim s As String
On Error GoTo frmErr
'Call cmbVocab_Change

If Not fFileOpen Then Exit Sub
If Len(cmbVocab.Text) = 0 Then Exit Sub


If KeyAscii = 13 Then
    PicReal.Picture = Nothing

    s = cmbVocab.Text
'in drawword
'    Do While InStr(s, "  ")
'        s = Replace(s, "  ", mySpace)
'    Loop
'    s = Trim(s)

    Call DrawWord(s)
    
    If NoVocabFlag Then Exit Sub
    If VortexMod And (Not Block1Flag) Then Exit Sub
    
    arr = Split(s, mySpace)

    lblWordSize = UBound(arr) + 1 & " (" & cmbVocab.ItemData(CurrentWordInd) & "), " & lblWordSize
Else
    VocSelStart = cmbVocab.SelStart
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmbVocab_KeyPress()"
End Sub

Private Sub cmbVocAdr_Click()
On Error GoTo frmErr

cmbVocab.ListIndex = cmbVocAdr.ListIndex

Exit Sub
frmErr:
MsgBox Err.Description & ": cmbVocAdr_Click()"
End Sub

Private Sub cmbVocAdr_DropDown()
SetDropdownHeight cmbVocAdr, ScaleHeight
End Sub

Private Sub cmbVocAdr_GotFocus()
MouseUpBug = True 'not select in font list with mouse move from here
End Sub

Private Sub cmdChar_Click()
Dim Tmp As String
Dim arr() As String
Dim i As Long

On Error GoTo frmErr

If Not fFileOpen Then Exit Sub
'If VortexMod And (Not Block1Flag) Then Exit Sub

'cmbVocab.SelStart = Len(cmbVocab.text)
'cmbVocab.SelText = " " & cmdChar.Caption & " "
'For i = 0 To McListBox1.SelCount - 1
'    Tmp = Tmp & mySpace & McListBox1.List(McListBox1.SelItem(i)) & mySpace
'Next i

If VortexMod Then
    For i = 1 To UBound(selArr)
        Tmp = Tmp & mySpace & Hex(selArr(i) + 32) & mySpace
    Next i

Else
    For i = 1 To UBound(selArr)
        Tmp = Tmp & mySpace & Hex(selArr(i) + 1) & mySpace
    Next i
End If

cmbVocab.SelStart = VocSelStart
cmbVocab.SelText = Tmp

Tmp = cmbVocab.Text
Do While InStr(Tmp, "  ")
    Tmp = Replace(Tmp, "  ", mySpace)
Loop
Tmp = Trim(Tmp)

cmbVocab.Text = Tmp

VocSelStart = Len(cmbVocab.Text)
PicReal.Picture = Nothing
Call DrawWord(Tmp)

If NoVocabFlag Then Exit Sub
If VortexMod And (Not Block1Flag) Then
Else

arr = Split(Tmp, mySpace)
lblWordSize = cmbVocab.ItemData(CurrentWordInd) & ", " & lblWordSize
If UBound(arr) <> cmbVocab.ItemData(CurrentWordInd) - 1 Then
    lblWordSize.ForeColor = vbRed
Else
    lblWordSize.ForeColor = &H80000012
End If
End If

If Block1Flag Then
    'lblVocAdr.Caption = Hex(Word1StartArr(CurrentWordInd))
    cmbVocAdr.Text = Hex(Word1StartArr(CurrentWordInd))
Else
    'lblVocAdr.Caption = Hex(Word2StartArr(CurrentWordInd))
    cmbVocAdr.Text = Hex(Word2StartArr(CurrentWordInd))
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdChar_Click()"
End Sub

Private Sub CopyData(ByRef strData As String, ByRef FromArrExport As Boolean, Optional fHex As Boolean = False)
'fHex - right+Copy for myevic source
'CharDataArr(i) -> sClipFont.concat -> strData
Dim i As Integer
Dim sClipFont As New CString
'Dim tmpCharDataArr As New CString
Dim ListCount As Long
Dim SelCount As Long
Dim ItemNum As Long
'Dim baseI As Integer

On Error GoTo frmErr

SelCount = McListBox1.SelCount
ListCount = McListBox1.ListCount

Me.MousePointer = vbHourglass

If ListCount = 0 Or SelCount = 0 Then
    'copy draw, no fw open
    
    ReDim sColArr(0)
    ReDim sRowArr(0)
    sColArr(0) = sCol
    sRowArr(0) = sRow

    If fHex Then
        sClipFont.concat "0 = " & CharDataHEXArr(0) & vbCrLf
    Else
        Call bArr2CharData(0)    '(0)
        sClipFont.concat "0," & sCharData.Text & vbCrLf
    End If


Else
    'mass

    If ListCount = 0 Or SelCount = 0 And fFileOpen Then Exit Sub

    If FromArrExport Then    'from menu

        For i = 0 To UBound(arrExport)
        
        
            If arrExport(i) Then
            
'If VortexMod Then baseI = i + 31 Else baseI = i 'for vortex base?

               ' no happend from menu/// If fHex Then
               '     sClipFont.concat Hex(i + 1) & " = " & CharDataHEXArr(i) & vbCrLf
               ' Else
                    sClipFont.concat i + 1 & "," & CharDataArr(i) & vbCrLf
                    'CharDataArr and CharDataHEXArr count from 0
               ' End If
            End If

        Next i

    Else
        'copy current selections in list
        'For i = 0 To SelCount - 1    '0-226
        'ItemNum = McListBox1.SelItem(i)

        For i = 1 To UBound(selArr)
            ItemNum = selArr(i)
            
'If VortexMod Then baseI = ItemNum + 31 Else baseI = ItemNum

            If fHex Then
                sClipFont.concat Hex(ItemNum + 1) & " = " & CharDataHEXArr(ItemNum) & vbCrLf
            Else
                sClipFont.concat ItemNum + 1 & "," & CharDataArr(ItemNum) & vbCrLf
            End If

        Next i

    End If
End If

strData = left$(sClipFont.Text, Len(sClipFont.Text) - 2)  '- last vbCrLf

Set sClipFont = Nothing
Me.MousePointer = vbNormal

'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": CopyData()"
End Sub
Private Sub cmdCopy_Click()
Dim strData As String

On Error GoTo frmErr

If isSelection Then
    If old_X1 >= sCol Then Exit Sub
    If old_Y1 >= sRow Then Exit Sub
    If old_X2 > sCol Then old_X2 = sCol
    If old_Y2 > sRow Then old_Y2 = sRow
    Call Selection_bArr2sClipFont(strData, old_X1, old_Y1, old_X2, old_Y2)
Else
    Call CopyData(strData, False, False)
End If

Clipboard.Clear
Clipboard.SetText strData

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdCopy_Click()"
End Sub
Private Function writeAutoIni() As Boolean
On Error GoTo frmErr
Dim i As Long
Dim b() As Byte
Dim v() As Byte
Dim Ret As Long
Dim bStr As String
Dim bSearch() As Byte
Dim bl1start As Long
Dim bl2start As Long
Dim bl1fin As Long
Dim bl2fin As Long
Dim v1start As Long
Dim v1fin As Long
Dim v2start As Long
Dim v2fin As Long
Dim unknNumBytes As Boolean

Dim sNumBytes As String
Dim ShiftDownChar As String

ReDim b(4)

ReDim xTmp(lngBytes - 1)
ReDim bTmp(lngBytes - 1)

Seek #bFileIn, 1
Get #bFileIn, , xTmp()

For i = 0 To lngBytes - 1    'decript
    bTmp(i) = (xTmp(i) Xor (i + lngBytes + uMagic - lngBytes \ uMagic)) And 255
Next i

sNumBytes = "2"
Call FillByteArray("01EB800050F8040C", bSearch)    'to find blocks start from pointers

Ret = InStrB(1, bTmp, bSearch, vbBinaryCompare) - 1
If Ret = -1 Then Exit Function

Call FillByteArray("2DE9F0", bSearch)
Ret = InStrB(Ret, bTmp, bSearch, vbBinaryCompare) - 1
If Ret = -1 Then Exit Function

bl2start = b4toLong_BE(Ret - 4, bTmp)
bl1start = b4toLong_BE(Ret - 8, bTmp)
If bl1start = 18288 Then
    bl1start = bl2start
    bl2start = 0
End If

If bl1start <> 0 Then
    For i = bl1start To lngBytes - 3 Step 4
        'find block end

        b(0) = bTmp(i + 3)    ' AA AA 00 _00_

        If (b(0) <> 0) Then    'Or (bTmp(i + 4) = 0) - no (
            'Debug.Print Hex(i - 4)
            bl1fin = i - 4
            'If bl2start = 0 Then bl1fin = i - 88    'dirty correct len

            Exit For    'found font block end
        End If

    Next i
End If

If bl2start <> 0 Then
    For i = bl2start To lngBytes - 3 Step 4
        'find block end
        b(0) = bTmp(i + 3)
        If b(0) <> 0 Then
            '  Debug.Print Hex(i - 4)
            bl2fin = i - 4
            Exit For    'found font block end
        End If
    Next i
End If

'correct len of 1 block to match len of second block
If bl2start <> 0 Then
    i = bl2fin - bl2start
    bl1fin = bl1start + i
End If


'''''''''''find vocab
v2start = 0: v2fin = 0
v1start = 0: v1fin = 0

If bl2start = 0 Then
    'for 2bytes eleaf small screen

    Call FillByteArray("707372203D20307825780A00", bSearch)
    Ret = InStrB(1, bTmp, bSearch, vbBinaryCompare) - 1
    If Ret <> -1 Then
        v1start = Ret + 12

        Call FillByteArray("00000A00140023003C", bSearch)
        Ret = InStrB(v1start, bTmp, bSearch, vbBinaryCompare) - 1

        If Ret <> -1 Then
            v1fin = Ret

        Else

            Call FillByteArray("000A00140023003C", bSearch)
            Ret = InStrB(v1start, bTmp, bSearch, vbBinaryCompare) - 1
            If (Ret <> -1) And (v1fin > v1start) Then
                v1fin = Ret
                sNumBytes = "1"    'old
            End If
        End If

    End If

Else
    'for 2bytes other (only 1 vocab)
    'nuvoton

    Call FillByteArray("4E00750076006F0074006F006E00", bSearch)    'N u v o t o n
    Ret = InStrB(1, bTmp, bSearch, vbBinaryCompare) - 1
    If Ret <> -1 Then
        v1start = Ret + 82

        For i = v1start To lngBytes - 3
            b(0) = bTmp(i)
            b(1) = bTmp(i + 1)

            If bTmp(i) = 0 And bTmp(i + 1) = 0 And bTmp(i + 2) = 0 And bTmp(i + 3) = 0 And bTmp(i + 4) = 0 Then    'search 5 zero bytes
                v1fin = i + 1
                Exit For
            End If

            If b(0) <> 0 And b(1) <> 0 Then    'search 2 non zero bytes
                v1fin = i - 2
                Exit For
            End If

        Next i
    End If
End If

If v1fin <= v1start Then
    'old
    '1-bytes vocab! (v1fin = v1start - 2)
    '2 vocabs?)

    sNumBytes = "1"
    v2start = 0: v2fin = 0
    v1start = 0: v1fin = 0
    'nuvoton
    Call FillByteArray("4E00750076006F0074006F006E00", bSearch)    'N u v o t o n
    Ret = InStrB(1, bTmp, bSearch, vbBinaryCompare) - 1
    If Ret <> -1 Then
        v1start = Ret + 82

        Call FillByteArray("00001BB7000080", bSearch)
        Ret = InStrB(v1start, bTmp, bSearch, vbBinaryCompare) - 1
        If Ret <> -1 Then
            v1fin = Ret
        End If

        For i = v1start To lngBytes - 3 Step 2

            If bTmp(i) <> 0 And bTmp(i + 1) <> 0 Then
                If i < v1fin Then    'check shorter (for ikonn)
                    v1fin = i - 2
                End If
                Exit For
            End If

        Next i

    End If

    If (v1fin > v1start) And (v1fin - v1start > 32) Then

        'dirty check 99 74 00 93 74 00 989800
        v2start = 0: v2fin = 0
        Call FillByteArray("997400937400989800", bSearch)
        Ret = InStrB(1, bTmp, bSearch, vbBinaryCompare) - 1

        If Ret <> -1 Then
            v2start = Ret

            Call FillByteArray("000608", bSearch)
            Ret = InStrB(v2start, bTmp, bSearch, vbBinaryCompare) - 1
            If Ret <> -1 Then v2fin = Ret
        End If

    Else    'v1fin <= v1start
        'very old 1 vocab upper (evic mini 102)
        'simple
        v1start = 0: v1fin = 0
        v2start = 0: v2fin = 0
        'dirty check
        Call FillByteArray("462D004C2700462700", bSearch)
        Ret = InStrB(1, bTmp, bSearch, vbBinaryCompare) - 1
        If Ret <> -1 Then
            v1start = Ret

            Call FillByteArray("000608", bSearch)
            Ret = InStrB(v1start, bTmp, bSearch, vbBinaryCompare) - 1
            If Ret <> -1 Then v1fin = Ret

        End If
    End If

End If

'recheck numbytes (dirty, not work if first words edited with 00 00)
unknNumBytes = False
If v1fin > v1start Then
    ReDim v(31)    '(lVocab1End - lVocab1Start - 1)
    Call FillByteArray("0000", bSearch)
    For i = 0 To UBound(v)
        v(i) = bTmp(v1start + i)
    Next i
End If
Ret = InStrB(1, v(), bSearch, vbBinaryCompare) - 1
If Ret = -1 Then
    If sNumBytes = "2" Then unknNumBytes = True
    'sNumBytes = "1"
Else
    If sNumBytes = "1" Then unknNumBytes = True
    ' sNumBytes = "2"
End If

'Debug.Print
'Debug.Print "bl 1 start= " & Hex(bl1start)
'Debug.Print "bl 1 fin= " & Hex(bl1fin)
'Debug.Print "bl 2 start= " & Hex(bl2start)
'Debug.Print "bl 2 fin= " & Hex(bl2fin)
'
'Debug.Print "v 1 start= " & Hex(v1start)
'Debug.Print "v 1 fin= " & Hex(v1fin)
'Debug.Print "v 2 start= " & Hex(v2start)
'Debug.Print "v 2 fin= " & Hex(v2fin)


If bl1start > 0 And bl1fin > bl1start Then

    WriteKey "Auto", "Block1Start", mySpace & Hex(bl1start), iniFileName
    WriteKey "Auto", "Block1End", mySpace & Hex(bl1fin), iniFileName
    WriteKey "Auto", "Block2Start", mySpace & Hex(bl2start), iniFileName
    WriteKey "Auto", "Block2End", mySpace & Hex(bl2fin), iniFileName

    WriteKey "Auto", "Vocab1Start", mySpace & Hex(v1start), iniFileName
    WriteKey "Auto", "Vocab1End", mySpace & Hex(v1fin), iniFileName
    WriteKey "Auto", "Vocab2Start", mySpace & Hex(v2start), iniFileName
    WriteKey "Auto", "Vocab2End", mySpace & Hex(v2fin), iniFileName

    If Not unknNumBytes Then
        WriteKey "Auto", "NumBytes", mySpace & sNumBytes, iniFileName
    End If

    For i = 0 To lngBytes - 14
        'search for ShiftDownChar
        b(0) = bTmp(i)
        b(1) = bTmp(i + 4)
        b(2) = bTmp(i + 8)
        b(3) = bTmp(i + 12)

        If b(0) = 66 And b(1) = 75 And b(2) = 57 Then
            ShiftDownChar = "42,4B,39": Exit For
        End If
        If b(0) = 145 And b(1) = 136 And b(2) = 154 And b(3) = 146 Then
            ShiftDownChar = "91,88,9A,92": Exit For
        End If
        If b(0) = 123 And b(1) = 114 And b(2) = 132 And b(3) = 124 Then
            ShiftDownChar = "7B,72,84,7C": Exit For
        End If
        If b(0) = 113 And b(1) = 104 And b(2) = 122 And b(3) = 114 Then
            ShiftDownChar = "71,68,7A,72": Exit For
        End If
        If b(0) = 46 And b(1) = 37 And b(2) = 55 And b(3) = 47 Then
            ShiftDownChar = "2E,25,37,2F": Exit For
        End If
        If b(0) = 46 And b(1) = 55 And b(2) = 37 Then
            ShiftDownChar = "2E,37,25": Exit For
        End If
        If b(0) = 157 And b(1) = 148 And b(2) = 166 And b(3) = 158 Then
            ShiftDownChar = "9D,94,A6,9E": Exit For
        End If
    Next i

    WriteKey "Auto", "ShiftDownChar", mySpace & ShiftDownChar, iniFileName

    writeAutoIni = True
End If

'''
Exit Function
frmErr:
MsgBox Err.Description & ": writeAutoIni()"
End Function

Private Function writeMyEvicIni() As Boolean
On Error GoTo frmErr
Dim i As Long
Dim X() As Byte
'Dim b() As Byte
'Dim d() As String
Dim Ind As Long
Dim bSearch() As Byte
Dim Ret As Long
Dim sBlock1Start As String
Dim sBlock1End As String
Dim sBlock2Start As String
Dim sBlock2End As String
Dim sVocab1Start As String
Dim sVocab1End As String

Dim lBlock1Start As Long
Dim lBlock1End As Long
Dim lBlock2Start As Long
Dim lBlock2End As Long
Dim lVocab1Start As Long
Dim lVocab1End As Long

Dim sNumBytes As String

Dim ShiftDownChar As String
Dim ME_flag As Boolean

'ReDim X(3): ReDim b(3): ReDim d(3)
ReDim b(4)

ReDim xTmp(lngBytes - 1)
ReDim bTmp(lngBytes - 1)

Seek #bFileIn, 1
Get #bFileIn, , xTmp()

For i = 0 To lngBytes - 1    'decript
    bTmp(i) = (xTmp(i) Xor (i + lngBytes + uMagic - lngBytes \ uMagic)) And 255
Next i

Ind = &H140
'Seek #bFileIn, Ind + 1
'Get #bFileIn, , X()
'b(0) = (X(0) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
'b(1) = (X(1) Xor (Ind + 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255
'b(2) = (X(2) Xor (Ind + 2 + lngBytes + uMagic - lngBytes \ uMagic)) And 255
'b(3) = (X(3) Xor (Ind + 3 + lngBytes + uMagic - lngBytes \ uMagic)) And 255

'If b(0) = 77 And b(1) = 89 And b(2) = 70 And b(3) = 87 Then
If bTmp(Ind) = 77 And bTmp(Ind + 1) = 89 And bTmp(Ind + 2) = 70 And bTmp(Ind + 3) = 87 Then
    ShiftDownChar = "91,88,9A,92,8B"
    ME_flag = True    'MYFW
'sNumBytes = 1
ElseIf bTmp(Ind) = 65 And bTmp(Ind + 1) = 70 And bTmp(Ind + 2) = 79 And bTmp(Ind + 3) = 88 Then
    ShiftDownChar = "53,56,5C,5D,65"
    ME_flag = True    'AFOX
'sNumBytes = 1
End If

If ME_flag Then
    Ind = &H144

    lBlock1Start = b4toLong_BE(Ind, bTmp)
    sBlock1Start = Hex(lBlock1Start)

    Ind = &H148

    lBlock1End = b4toLong_BE(Ind, bTmp) - 4
    sBlock1End = Hex(lBlock1End)

    Ind = &H14C

    lBlock2Start = b4toLong_BE(Ind, bTmp)
    sBlock2Start = Hex(lBlock2Start)

    Ind = &H150

    lBlock2End = b4toLong_BE(Ind, bTmp) - 4
    sBlock2End = Hex(lBlock2End)

    Ind = &H154

    lVocab1Start = b4toLong_BE(Ind, bTmp)
    sVocab1Start = Hex(lVocab1Start)

    Ind = &H158

    lVocab1End = b4toLong_BE(Ind, bTmp) - 2
    sVocab1End = Hex(lVocab1End)

    If lBlock2End < lBlock2Start Then
        NoBlock2Flag = True
        optBlock(1).Enabled = False
        sBlock2Start = "0"    'Block1Start
        sBlock2End = "0"
    End If

    'check numbytes (dirty, not work if first words edited with 00 00)
    If lVocab1End > lVocab1Start Then
        ReDim X(31) '(lVocab1End - lVocab1Start - 1)
        Call FillByteArray("0000", bSearch)
        For i = 0 To UBound(X)
            X(i) = bTmp(lVocab1Start + i)
        Next i
    End If
    Ret = InStrB(1, X(), bSearch, vbBinaryCompare) - 1
    If Ret = -1 Then
        sNumBytes = "1"
    Else
        sNumBytes = "2"
    End If

    WriteKey "MyEvic", "Block1Start", mySpace & sBlock1Start, iniFileName
    WriteKey "MyEvic", "Block1End", mySpace & sBlock1End, iniFileName
    WriteKey "MyEvic", "Block2Start", mySpace & sBlock2Start, iniFileName
    WriteKey "MyEvic", "Block2End", mySpace & sBlock2End, iniFileName
    WriteKey "MyEvic", "Vocab1Start", mySpace & sVocab1Start, iniFileName
    WriteKey "MyEvic", "Vocab1End", mySpace & sVocab1End, iniFileName

    WriteKey "MyEvic", "ShiftDownChar", mySpace & ShiftDownChar, iniFileName

    WriteKey "MyEvic", "NumBytes", mySpace & sNumBytes, iniFileName

    writeMyEvicIni = True
End If

'''
Exit Function
frmErr:
MsgBox Err.Description & ": writeMyEvicIn()"
End Function

Private Function CheckVortex(ByRef Ind As Long) As Boolean
Dim b() As Byte

ReDim b(2)
Get #bFileIn, Ind + 1, b()

If b(0) = IDVortex(0) And b(1) = IDVortex(1) And b(2) = IDVortex(2) Then
    CheckVortex = True
End If

End Function

Private Function CheckIdFw() As Boolean
'select HW in cmbHard.Text
Dim i As Integer
Dim X() As Byte
Dim b() As Byte
Dim Ind As Long
Dim MatchIndex As Integer

On Error GoTo frmErr

If Not fFileOpen Then Exit Function

MatchIndex = -1

For i = 0 To UBound(IDpassword)
    ReDim X(4)
    ReDim b(4)

    If Len(IDpassword(i)) <> 0 Then
        Ind = "&H" & IDpassword(i)


        If ISVortex(i) Then
            If CheckVortex(Ind) Then
                VortexMod = True
                MatchIndex = i
                Exit For
            End If
            
        'for myevic def
        ElseIf Ind = 0 Then
            If writeMyEvicIni Then
                MatchIndex = i
                Exit For
            End If

            'Auto
        ElseIf Ind = 1 Then
            If writeAutoIni Then
                MatchIndex = i
                Exit For
            End If


        Else
            Seek #bFileIn, Ind    '+ 1 'for get 0 before/  0 128 81 1 0
            Get #bFileIn, , X()

            b(0) = (X(0) Xor (Ind - 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255

            b(1) = (X(1) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
            b(2) = (X(2) Xor (Ind + 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255
            b(3) = (X(3) Xor (Ind + 2 + lngBytes + uMagic - lngBytes \ uMagic)) And 255

            b(4) = (X(4) Xor (Ind + 3 + lngBytes + uMagic - lngBytes \ uMagic)) And 255
            'b(0) = (X(4) Xor (Ind + 4 + lngBytes + uMagic - lngBytes \ uMagic)) And 255

            '   If IDanswer(0) = b(0) And IDanswer(1) = b(1) And IDanswer(2) = b(2) And b(3) = 0 And b(4) = 0 Then
            '"HID " for new
            If ((uMagic = &H3745B6) _
            And (&H48 = b(1) And &H49 = b(2) And &H44 = b(3) And b(4) = &H20)) _
            Or ((uMagic = &H63B38) _
            And (b(0) = 0 And IDanswer(0) = b(1) And IDanswer(1) = b(2) And IDanswer(2) = b(3) And b(4) = 0)) _
            Then

                If IDsecuredAdr(i) <> 0 Then

                    Ind = IDsecuredAdr(i)
                    Seek #bFileIn, Ind + 1
                    Get #bFileIn, , X()
                    b(0) = (X(0) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
                    If b(0) = IDsecuredByte(i) Then
                        MatchIndex = i
                        Exit For
                    End If

                Else
                    MatchIndex = i
                    Exit For
                End If

            End If

        End If    'If Ind = 0 Then
    End If    'If Len(IDpassword(i)) <> 0
Next i

If MatchIndex > -1 Then
    CheckIdFw = True
    cmbHard.ListIndex = MatchIndex
Else
    CheckIdFw = False
End If


'''
Exit Function
frmErr:
MsgBox Err.Description & ": CheckIdFw()"
End Function
Private Sub FillVocab()
Dim i As Long, n As Integer, k As Integer, j As Long
Dim Ind As Long    'pos in file
Dim endBlock As Long
Dim Tmp As String
Dim sHex As New CString
Dim s0xHex As New CString
Dim No2VocFlag As Boolean
Dim WordSize As Long
On Error GoTo frmErr
'Debug.Print ">   FillVocab"

If Not fFileOpen Then Exit Sub

cmbVocab.Clear
cmbVocAdr.Clear
ReDim VocBlock1Arr(0)
ReDim VocBlock2Arr(0)
ReDim VocBlock1Arr0x(0)
ReDim VocBlock2Arr0x(0)
ReDim Word1StartArr(0)
ReDim Word2StartArr(0)
ReDim Word1LenArr(0)
ReDim Word2LenArr(0)
'lblVocAdr.Caption = vbNullString

Tmp = VBGetPrivateProfileString(Hardtext, "NumBytes", iniFileName)
If Len(Tmp) = 0 Then
    WordCharNumBytes = 1
ElseIf IsNumeric(Tmp) Then
    WordCharNumBytes = Val(Tmp)
End If

Tmp = VBGetPrivateProfileString(Hardtext, "Vocab1Start", iniFileName)
If Len(Tmp) < 2 Then
    NoVocabFlag = True
    Exit Sub
End If
Vocab1Start = "&H" & Tmp

Tmp = VBGetPrivateProfileString(Hardtext, "Vocab1End", iniFileName)
If Len(Tmp) = 0 Then
    NoVocabFlag = True
    Exit Sub
End If
Vocab1End = "&H" & Tmp

NoVocabFlag = False

Tmp = VBGetPrivateProfileString(Hardtext, "Vocab2Start", iniFileName)
No2VocFlag = False
If Len(Tmp) < 2 Then
    No2VocFlag = True
Else
    Vocab2Start = "&H" & Tmp
End If

Tmp = VBGetPrivateProfileString(Hardtext, "Vocab2End", iniFileName)
If Len(Tmp) = 0 Then
    'Vocab2End = Vocab1End
ElseIf UCase(Tmp) = "EOF" Then
    Vocab2End = lngBytes - 1
Else
    Vocab2End = "&H" & Tmp
End If


If No2VocFlag Then k = 1 Else k = 2

Ind = Vocab1Start
endBlock = Vocab1End

ReDim bTmp(1)
ReDim xTmp(lngBytes - 1)

Seek #bFileIn, 1
Get #bFileIn, , xTmp()

For i = 1 To k
    n = 0
    Do While Ind <= endBlock   ' lngBytes

        ' Seek bFileIn, Ind

        'need parse both
        'AF    86    8E    91    00
        'AF 00 86 00 8E 00 91 00 00 00
        If WordCharNumBytes = 1 Then

            For j = Ind To endBlock
                bTmp(0) = (xTmp(j) Xor (j + lngBytes + uMagic - lngBytes \ uMagic)) And 255    'xor magic

                Tmp = right$("0" & Hex(bTmp(0)), 2)
                If bTmp(0) = 0 Then  '# word end
                    Exit For
                Else
                    sHex.concat Tmp & mySpace
                    s0xHex.concat "0x" & Tmp & ","
                    WordSize = WordSize + 1
                End If
            Next j

        Else    'If WordCharNumBytes = 2 Then

            For j = Ind To endBlock Step 2

                bTmp(0) = (xTmp(j) Xor (j + lngBytes + uMagic - lngBytes \ uMagic)) And 255    'lo
                bTmp(1) = (xTmp(j + 1) Xor (j + 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255    'hi

                If (bTmp(0) = 0) And (bTmp(1) = 0) Then  '## word end
                    Exit For
                Else

                    Tmp = right$("0" & Hex(bTmp(0)), 2)
                    If bTmp(1) = 0 Then

                        sHex.concat Tmp & mySpace
                        s0xHex.concat "0x" & Tmp & ","

                    Else
                        sHex.concat Hex(bTmp(1)) & Tmp & mySpace
                        s0xHex.concat "0x" & Hex(bTmp(1)) & Tmp & ","
                    End If
                    WordSize = WordSize + 1

                End If
            Next j


        End If

        If sHex.Text <> vbNullString Then

            If i = 1 Then    '1 block
                ReDim Preserve VocBlock1Arr(n)
                ReDim Preserve VocBlock1Arr0x(n)
                VocBlock1Arr(n) = Trim(sHex.Text)
                VocBlock1Arr0x(n) = Trim(s0xHex.Text)
                ReDim Preserve Word1StartArr(n)
                ReDim Preserve Word1LenArr(n)
                Word1StartArr(n) = Ind
                Word1LenArr(n) = WordSize
                If No2VocFlag Then    ' same set if not specified
                    ReDim Preserve VocBlock2Arr(n)
                    ReDim Preserve VocBlock2Arr0x(n)
                    VocBlock2Arr(n) = Trim(sHex.Text)
                    VocBlock2Arr0x(n) = Trim(s0xHex.Text)
                    ReDim Preserve Word2StartArr(n)
                    ReDim Preserve Word2LenArr(n)
                    Word2StartArr(n) = Ind
                    Word2LenArr(n) = WordSize
                End If
            Else      '2 block
                ReDim Preserve VocBlock2Arr(n)
                ReDim Preserve VocBlock2Arr0x(n)
                VocBlock2Arr(n) = Trim(sHex.Text)
                VocBlock2Arr0x(n) = Trim(s0xHex.Text)
                ReDim Preserve Word2StartArr(n)
                ReDim Preserve Word2LenArr(n)
                Word2StartArr(n) = Ind
                Word2LenArr(n) = WordSize
            End If

            sHex.reset: s0xHex.reset
            n = n + 1

        End If

        Ind = Ind + WordCharNumBytes * (WordSize + 1)
        WordSize = 0

    Loop

    Ind = Vocab2Start
    endBlock = Vocab2End
Next i

If Block1Flag Then    '1 block to combo
    For i = 0 To UBound(VocBlock1Arr)
        cmbVocab.AddItem VocBlock1Arr(i)
        cmbVocab.ItemData(i) = Word1LenArr(i)
        cmbVocAdr.AddItem Hex(Word1StartArr(i))

    Next i
Else
    For i = 0 To UBound(VocBlock2Arr)
        cmbVocab.AddItem VocBlock2Arr(i)
        cmbVocab.ItemData(i) = Word2LenArr(i)
        cmbVocAdr.AddItem Hex(Word2StartArr(i))
    Next i
End If

'cmbVocab.text = cmbVocab.List(0)
'VocSelStart = Len(cmbVocab.text)
'lblWordSize.ForeColor = &H80000012
lblWordSize = vbNullString
'Set sHex = Nothing
'cmbVocab.ListIndex = cmbLastIndex
'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillVocab()"
End Sub


Private Sub FillVocab_Vortex()
Dim i As Long, n As Integer, k As Integer, j As Long
Dim Ind As Long    'pos in file
Dim endBlock As Long
Dim Tmp As String
Dim sHex As New CString
Dim s0xHex As New CString
Dim No2VocFlag As Boolean
Dim WordSize As Long
On Error GoTo frmErr
'Debug.Print ">   FillVocab"

If Not fFileOpen Then Exit Sub

cmbVocab.Clear
cmbVocAdr.Clear
ReDim VocBlock1Arr(0)
ReDim VocBlock2Arr(0)
ReDim VocBlock1Arr0x(0)
ReDim VocBlock2Arr0x(0)
ReDim Word1StartArr(0)
ReDim Word2StartArr(0)
ReDim Word1LenArr(0)
ReDim Word2LenArr(0)
'lblVocAdr.Caption = vbNullString

Tmp = VBGetPrivateProfileString(Hardtext, "NumBytes", iniFileName)
If Len(Tmp) = 0 Then
    WordCharNumBytes = 1
ElseIf IsNumeric(Tmp) Then
    WordCharNumBytes = Val(Tmp)
End If

Tmp = VBGetPrivateProfileString(Hardtext, "Vocab1Start", iniFileName)
If Len(Tmp) < 2 Then
    NoVocabFlag = True
    Exit Sub
End If
Vocab1Start = "&H" & Tmp

Tmp = VBGetPrivateProfileString(Hardtext, "Vocab1End", iniFileName)
If Len(Tmp) = 0 Then
    NoVocabFlag = True
    Exit Sub
End If
Vocab1End = "&H" & Tmp

NoVocabFlag = False

'Tmp = VBGetPrivateProfileString(Hardtext, "Vocab2Start", iniFileName)
'No2VocFlag = False
'If Len(Tmp) < 2 Then
'    No2VocFlag = True
'Else
'    Vocab2Start = "&H" & Tmp
'End If
'
'Tmp = VBGetPrivateProfileString(Hardtext, "Vocab2End", iniFileName)
'If Len(Tmp) = 0 Then
'    'Vocab2End = Vocab1End
'ElseIf UCase(Tmp) = "EOF" Then
'    Vocab2End = lngBytes - 1
'Else
'    Vocab2End = "&H" & Tmp
'End If


'If No2VocFlag Then k = 1 Else k = 2

Ind = Vocab1Start
endBlock = Vocab1End

'ReDim bTmp(1)
ReDim xTmp(lngBytes - 1)

Seek #bFileIn, 1
Get #bFileIn, , xTmp()

If Block1Flag Then
    'For i = 1 To k
    n = 0
    Do While Ind <= endBlock   ' lngBytes

        'Seek bFileIn, Ind


        '        If WordCharNumBytes = 1 Then

        For j = Ind To endBlock
            'bTmp(0) = (xTmp(j) Xor (j + lngBytes + uMagic - lngBytes \ uMagic)) And 255    'xor magic

            Tmp = right$("0" & Hex(xTmp(j)), 2)
            If xTmp(j) = 0 Then  '# word end
                Exit For
            Else
                sHex.concat Tmp & mySpace
                s0xHex.concat "0x" & Tmp & ","
                WordSize = WordSize + 1
            End If
        Next j

        'Debug.Print sHex.Text
        'Debug.Print s0xHex.Text
        'Debug.Print

        '        Else    'If WordCharNumBytes = 2 Then

        '            For j = Ind To endBlock Step 2
        '
        '                bTmp(0) = (xTmp(j) Xor (j + lngBytes + uMagic - lngBytes \ uMagic)) And 255    'lo
        '                bTmp(1) = (xTmp(j + 1) Xor (j + 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255    'hi
        '
        '                If bTmp(0) = 0 And bTmp(1) = 0 Then  '## word end
        '                    Exit For
        '                Else
        '
        '                    Tmp = right$("0" & Hex(bTmp(0)), 2)
        '                    If bTmp(1) = 0 Then
        '
        '                        sHex.concat Tmp & mySpace
        '                        s0xHex.concat "0x" & Tmp & ","
        '
        '                    Else
        '                        sHex.concat Hex(bTmp(1)) & Tmp & mySpace
        '                        s0xHex.concat "0x" & Hex(bTmp(1)) & Tmp & ","
        '                    End If
        '                    WordSize = WordSize + 1
        '
        '                End If
        '            Next j
        '       End If

        '  If i = 1 Then    '1 block
        If WordSize > 0 Then
            ReDim Preserve VocBlock1Arr(n)
            ReDim Preserve VocBlock1Arr0x(n)
            VocBlock1Arr(n) = Trim(sHex.Text)
            VocBlock1Arr0x(n) = Trim(s0xHex.Text)
            ReDim Preserve Word1StartArr(n)
            ReDim Preserve Word1LenArr(n)
            Word1StartArr(n) = Ind
            Word1LenArr(n) = WordSize



            '            If No2VocFlag Then    ' same set if not specified
            '                ReDim Preserve VocBlock2Arr(n)
            '                ReDim Preserve VocBlock2Arr0x(n)
            '                VocBlock2Arr(n) = Trim(sHex.Text)
            '                VocBlock2Arr0x(n) = Trim(s0xHex.Text)
            '                ReDim Preserve Word2StartArr(n)
            '                ReDim Preserve Word2LenArr(n)
            '                Word2StartArr(n) = Ind
            '                Word2LenArr(n) = WordSize
            '            End If
            '        Else      '2 block
            '            ReDim Preserve VocBlock2Arr(n)
            '            ReDim Preserve VocBlock2Arr0x(n)
            '            VocBlock2Arr(n) = Trim(sHex.Text)
            '            VocBlock2Arr0x(n) = Trim(s0xHex.Text)
            '            ReDim Preserve Word2StartArr(n)
            '            ReDim Preserve Word2LenArr(n)
            '            Word2StartArr(n) = Ind
            '            Word2LenArr(n) = WordSize
            '        End If


            n = n + 1
        End If
        sHex.reset: s0xHex.reset
        Ind = Ind + WordCharNumBytes * (WordSize + 1)
        WordSize = 0
    Loop

    'Ind = Vocab2Start
    'endBlock = Vocab2End
    'Next i
End If

If Block1Flag Then    '1 block to combo
    For i = 0 To UBound(VocBlock1Arr)
        cmbVocab.AddItem VocBlock1Arr(i)
        cmbVocab.ItemData(i) = Word1LenArr(i)
        cmbVocAdr.AddItem Hex(Word1StartArr(i))

    Next i
Else
    cmbVocab.Clear

    '    For i = 0 To UBound(VocBlock2Arr)
    '        cmbVocab.AddItem VocBlock2Arr(i)
    '        cmbVocab.ItemData(i) = Word2LenArr(i)
    '        cmbVocAdr.AddItem Hex(Word2StartArr(i))
    '    Next i
End If

lblWordSize = vbNullString

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillVocab_Vortex"
End Sub



Public Sub LoadFWfile()
Dim i As Integer
Dim Ret As Long

On Error GoTo frmErr

' to cmdReloadFW_Click()
'If FileNameFW = vbNullString Then
'    If LastOpenedFW = vbNullString Then
'        Exit Sub
'    Else
'        FileNameFW = LastOpenedFW
'        'FileTitle = GetNameExt(FileNameFW)
'    End If
'    'Else
'    'FileTitle = GetNameExt(FileNameFW)
'End If

If Not FileExists(FileNameFW) Then Exit Sub

FileTitle = GetNameExt(FileNameFW)

If fFileOpen Then    'need
    Close #bFileIn
    fFileOpen = False
End If

bFileIn = FreeFile
If Not OpenFW_read Then Exit Sub
LastPath = GetPathFromPathAndName(FileNameFW)
lngBytes = LOF(bFileIn)    '1
fFileOpen = True

VortexMod = False
'''''''''''' !!!!!!!!!!!!!!!!!!!!!!!!!!             go check                !!!!!!!!!!!!!!!
If Not CheckIdFw Then   'search for definition
    Ret = MsgBoxEx(ArrMsg(19), , , CenterOwner, vbCritical)
    'If Ret <> 1 Then
    Close #bFileIn
    fFileOpen = False
    FileNameFW = vbNullString
    Exit Sub    'unknown FW
    'End If
End If

LastOpenedFW = FileNameFW
Me.Caption = "VTCFont v" & App.Major & "." & App.Minor & "." & App.Revision & ": " & FileTitle

Hardtext = cmbHard.Text    'store if changed while process
NoVocabFlag = True

If VortexMod Then

    chkDupFont.Value = False
    
    Call FillFontList_Vortex
    Call FillVocab_Vortex

Else
    Call FillVocab
    Call FillFontList    'fill cmbAdr, char addr list with char number
End If

If (Not Block1Flag) And NoBlock2Flag Then
    Me.MousePointer = vbNormal
    optBlock(0).Value = vbChecked
    Exit Sub
End If

picContainer.Visible = False    'need
McListBox1.Visible = False
Me.MousePointer = vbHourglass

Call FillMCList    'fill MClist same

picContainer.Visible = True
McListBox1.Visible = True


If cmbLastIndex > cmbAdr.ListCount - 1 Then cmbLastIndex = 0
If UBound(selArr) < 2 Then
    cmbAdr.ListIndex = cmbLastIndex    'cmbAdr_Click return first
Else

    If selArr(1) <= cmbAdr.ListCount Then
        cmbAdr.ListIndex = selArr(1)
        'Else
        'no, later cmbAdr.ListIndex = cmbLastIndex
    End If

    McListBox1.ClearSelectionAll
    If UBound(selArr) <= cmbAdr.ListCount Then
        For i = 1 To UBound(selArr)
            If selArr(i) <= cmbAdr.ListCount Then
                McListBox1.SelectItem (selArr(i))
            Else
                ReDim selArr(0)
                cmbAdr.ListIndex = cmbLastIndex    'click
                Exit For
            End If
        Next i
    End If
    McListBox1.Refresh
End If

ReDim arrExport(cmbAdr.ListCount - 1)
GoExportFlag = False

cmdCopy.Caption = ArrMsg(13) & " (" & McListBox1.SelCount & ")"

Me.MousePointer = vbNormal

'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": LoadFWfile()"
End Sub


Private Sub cmdCopy_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim strData As String
Dim i As Integer

On Error Resume Next

If Button = 2 Then

      '  For i = 1 To UBound(selArr)
      '  Call bArr2HEXCharData(selArr(i) + 1)
            
      '  Next i
        
    Call bArr2HEXCharData(cmbAdr.ListIndex)
    
    Call CopyData(strData, False, True)

    Clipboard.Clear
    Clipboard.SetText strData

End If
End Sub

Public Sub cmdFWUpdater_Click()
    Dim fwu As String    'path
    Dim fu_hwnd As Long
    On Error GoTo frmErr
    fwu = App.Path & "/FWUpdater.exe"

    fu_hwnd = FindWindow(vbNullString, "FWUpdater")

    If FileExists(fwu) Then

        If fu_hwnd <> 0 Then
            'ShowWindow hwnd, 9 'SW_SHOWNORMAL
            SetForegroundWindow fu_hwnd
        Else

            If fFileOpen And Len(FileNameFW) <> 0 Then

                ShellExecute Me.hWnd, "Open", fwu, FileNameFW, App.Path, SW_SHOWNORMAL
            Else
                ShellExecute Me.hWnd, "Open", fwu, vbNullString, App.Path, SW_SHOWNORMAL
            End If
            
        End If
    Else
        'no fwu
        MsgBoxEx ArrMsg(45), , , CenterOwner, vbCritical
    End If


    '''
    Exit Sub
frmErr:
    MsgBox Err.Description & ": cmdFWUpdater()"
End Sub

Private Sub cmdINIShow_Click(Index As Integer)
'Dim oldLangEng As Boolean
'Dim oldMagn As Boolean
'Dim oldDither As Boolean
'Dim oldMouse As Integer

On Error GoTo frmErr

Select Case Index
Case 0    'opt
    frmOptions.Show 1, frmmain
'    Shell "notepad.exe " & App.Path & "\VTCFont.ini", vbNormalFocus

'    oldLangEng = LanguageEng
'    Call LoadIniGlobal
'    If LanguageEng <> oldLangEng Then Call GetLanguage(1)


Case 1
    Shell "notepad.exe " & App.Path & "\readme.txt", vbNormalFocus
End Select

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdINIShow_Click()"
End Sub
Public Sub reloadIni()
'reread some parts of ini
'Dim sectArr() As String
'Dim n As Integer, i As Integer
Dim WFD As WIN32_FIND_DATA
Dim Ret As Long
Dim Tmp As String
On Error GoTo frmErr



Tmp = VBGetPrivateProfileString("Global", "CheckCharSize", iniFileName)
CheckCharSizeFlag = True
If Len(Tmp) <> 0 Then CheckCharSizeFlag = Val(Tmp)

Tmp = VBGetPrivateProfileString("Global", "Magnify", iniFileName)
'PicReal.Visible = True
If Len(Tmp) <> 0 Then
    Magnify = Val(Tmp)
End If
'    If Magnify Then
'        picX3.Visible = True
'        PicReal.Visible = False
'    Else
'        picX3.Visible = False
'        PicReal.Visible = True
'    End If

Tmp = VBGetPrivateProfileString("Global", "WordsInLine", iniFileName)
If Len(Tmp) <> 0 Then AllWordsInLineFlag = CBool(Tmp)

Tmp = VBGetPrivateProfileString("Global", "Language", iniFileName)
If Len(Tmp) <> 0 Then Language = LCase(Trim(Tmp))

Tmp = VBGetPrivateProfileString("Global", "InvertMouseB", iniFileName)
If Len(Tmp) <> 0 Then InvertMouseB = Val(Tmp)

Tmp = VBGetPrivateProfileString("Global", "PicDithered", iniFileName)
If IsNumeric(Tmp) Then fPicDithered = Val(Tmp)
If fPicDithered > 3 Then fPicDithered = 1

'    Select Case Language
'    Case "ru"
'        'lngFileName = App.Path & "\VTCFont_Ru.lng"
'    Case "en"
'        'lngFileName = App.Path & "\VTCFont_En.lng"
'    End Select

If Len(Language) = 0 Then Language = "en"
lngFileNameOnly = "VTCFont_" & Language & ".lng"
lngFileName = App.Path & "\" & lngFileNameOnly

Ret = FindFirstFile(lngFileName, WFD)
If Ret < 0 Then
'lngFileName = App.Path & "\VTCFont_En.lng"
    MsgBoxEx ArrMsg(11) & mySpace & lngFileName, , CenterScreen, , vbExclamation   '"Warning: lang file not found!"
    FindClose Ret
    lngFileName = vbNullString
    cmdINIShow_Click (0)
End If

If Len(lngFileName) <> 0 Then Call GetLanguage(1)



If fFileOpen Then
    cmdSaveAll.Caption = ArrMsg(12) & " (" & GetChangesCount & ")"
    cmdCopy.Caption = ArrMsg(13) & " (" & McListBox1.SelCount & ")"
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": reloadIni"
End Sub

'Private Sub cmdINIShow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Index = 0 And Button = vbRightButton Then
'    Call reloadIni
'End If
'End Sub

Private Sub cmdLoadFile_Click()
Dim Ret As Long
On Error GoTo frmErr

If fFileOpen Then
    If GetChangesCount > 0 Then
        Ret = MsgBoxEx(ArrMsg(23), , , CenterOwner, vbOKCancel Or vbQuestion)    'unsaved
        If Ret <> 1 Then Exit Sub
    End If
End If

EscFlag = False
FileNameFW = vbNullString
FileNameFW = pLoadDialog(ArrMsg(15), FileTitle)

DoEvents 'click on (, bugfixed

If Len(FileNameFW) = 0 Then Exit Sub

'encrypt if decrypted
If Not EncryptFW Then
    Ret = MsgBoxEx(ArrMsg(19), , , CenterOwner, vbCritical)
        'Close #bFileIn
        fFileOpen = False
        Exit Sub    'unknown FW
End If

Call LoadFWfile

OldPatchIndex = 0

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdLoadFile()"
End Sub

Private Sub FillFontList()
Dim i As Long, n As Integer, k As Integer
Dim Ind As Long    'pos in file
Dim endBlock As Long
Dim Tmp As String
Dim Block1Start As Long
Dim Block2Start As Long
Dim Block1End As Long
Dim Block2End As Long
Dim sHex As String
On Error GoTo frmErr
'Debug.Print ">   FillFontList"

If Not fFileOpen Then Exit Sub

Tmp = VBGetPrivateProfileString(Hardtext, "ShiftDownChar", iniFileName)
If Len(Tmp) <> 0 Then
    ShiftChar = Split(Tmp, ",")
Else
    ReDim ShiftChar(0)
End If

optBlock(1).Enabled = True

Tmp = VBGetPrivateProfileString(Hardtext, "Block1Start", iniFileName)
If Len(Tmp) = 0 Then
    MsgBoxEx ArrMsg(1), , , CenterOwner    '"No block 1 specified in INI file."
    Exit Sub
End If
Block1Start = "&H" & Tmp
Tmp = VBGetPrivateProfileString(Hardtext, "Block1End", iniFileName)
If Len(Tmp) = 0 Then Exit Sub
Block1End = "&H" & Tmp

Tmp = VBGetPrivateProfileString(Hardtext, "Block2Start", iniFileName)
NoBlock2Flag = False
If Len(Tmp) = 0 Or Tmp = "0" Then
    NoBlock2Flag = True
    optBlock(1).Enabled = False
    Block2Start = Block1Start
Else
    Block2Start = "&H" & Tmp
End If

Tmp = VBGetPrivateProfileString(Hardtext, "Block2End", iniFileName)
If Len(Tmp) = 0 Then
    Block2End = Block1End
Else
    Block2End = "&H" & Tmp
End If

If Block1Flag Then    ' block 1, second in firmare
'    ind = Block1Start
'    endBlock = Block1End
Else    'block 2 first in firmware
    If NoBlock2Flag Then
        MsgBoxEx ArrMsg(2), , , CenterOwner    '"No block 2 specified in INI file."
'        cmbAdr.Clear
'        ReDim FontBlock1Arr(0)
'        ReDim FontBlock2Arr(0)
        Exit Sub
    End If
'    ind = Block2Start
'    endBlock = Block2End
End If

ReDim bTmp(2)    '(lngBytes - 1)
ReDim xTmp(lngBytes - 1)
cmbAdr.Clear
ReDim FontBlock1Arr(0)
ReDim FontBlock2Arr(0)
ReDim FontBlock1IndArr(0)
ReDim FontBlock2IndArr(0)

'Seek #bFileIn, 1
Get #bFileIn, 1, xTmp()

If NoBlock2Flag Then k = 1 Else k = 2
Ind = Block1Start
endBlock = Block1End

For i = 1 To k
    n = 0
    Do While Ind <= endBlock   ' lngBytes
        'Seek bFileIn, Ind
        bTmp(0) = (xTmp(Ind) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255    'xor magic
        bTmp(1) = (xTmp(Ind + 1) Xor (Ind + 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        bTmp(2) = (xTmp(Ind + 2) Xor (Ind + 2 + lngBytes + uMagic - lngBytes \ uMagic)) And 255

        If bTmp(2) = 0 Then
            sHex = Hex(bTmp(1)) & right$("0" & Hex(bTmp(0)), 2)
        Else
            sHex = Hex(bTmp(2)) & right$("0" & Hex(bTmp(1)), 2) & right$("0" & Hex(bTmp(0)), 2)
        End If

'If bTmp(2) = 0 Then
'        sHex = Right("0" & Hex(bTmp(1)), 2) & Right("0" & Hex(bTmp(0)), 2)
'Else
'        sHex = Right("0" & Hex(bTmp(2)), 2) & Right("0" & Hex(bTmp(1)), 2) & Right("0" & Hex(bTmp(0)), 2)
'End If

        If i = 1 Then    '1 block
            ReDim Preserve FontBlock1Arr(n)
            
            If uMagic = &H3745B6 Then 'new Tour FW
                FontBlock1Arr(n) = Hex(CLng("&H" & sHex) - &H3800) '&H3800 = start addr of FW on stm32
            Else
                FontBlock1Arr(n) = sHex
            End If
            
            ReDim Preserve FontBlock1IndArr(n)
            FontBlock1IndArr(n) = Ind    'adr of pointer to current font
            
        Else    '2 block
            ReDim Preserve FontBlock2Arr(n)
            FontBlock2Arr(n) = sHex
            ReDim Preserve FontBlock2IndArr(n)
            FontBlock2IndArr(n) = Ind
        End If
        
        n = n + 1
        Ind = Ind + 4
        
    Loop

    Ind = Block2Start
    endBlock = Block2End
Next i

If Block1Flag Then    '1 block to combo
    ReDim bTmpCollection(UBound(FontBlock1Arr))
    For i = 0 To UBound(FontBlock1Arr)
        cmbAdr.AddItem FontBlock1Arr(i)
        cmbAdr.ItemData(i) = i + 1
    Next i
Else
    ReDim bTmpCollection(UBound(FontBlock2Arr))
    For i = 0 To UBound(FontBlock2Arr)
        cmbAdr.AddItem FontBlock2Arr(i)
        cmbAdr.ItemData(i) = i + 1
    Next i
End If

'cmbAdr.ListIndex = cmbLastIndex    'cmbAdr_Click > startAddr = cmbAdr.text : Call GetBlock

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillFontList()"
End Sub

Private Sub FillFontList_Vortex()
Dim i As Long, n As Integer, k As Integer
Dim Ind As Long    'pos in file
Dim endBlock As Long
Dim Tmp As String
Dim lTmp As Long
Dim Block1Start As Long
Dim Block2Start As Long
Dim Block1End As Long
Dim Block2End As Long
Dim sHex As String

On Error GoTo frmErr
'Debug.Print ">   FillFontList"

If Not fFileOpen Then Exit Sub

Tmp = VBGetPrivateProfileString(Hardtext, "ShiftDownChar", iniFileName)
If Len(Tmp) <> 0 Then
    ShiftChar = Split(Tmp, ",")
Else
    ReDim ShiftChar(0)
End If

optBlock(1).Enabled = True
'NoBlock2Flag = True

Tmp = VBGetPrivateProfileString(Hardtext, "Block1Start", iniFileName)
If Len(Tmp) = 0 Then
    MsgBoxEx ArrMsg(1), , , CenterOwner    '"No block 1 specified in INI file."
    Exit Sub
End If
Block1Start = "&H" & Tmp

ReDim xTmp(0)
Get #bFileIn, Block1Start + 7, xTmp()
VortexBlock1Height = xTmp(0)

Tmp = VBGetPrivateProfileString(Hardtext, "Block1End", iniFileName)
If Len(Tmp) = 0 Then Exit Sub
Block1End = "&H" & Tmp

Tmp = VBGetPrivateProfileString(Hardtext, "Block2Start", iniFileName)
NoBlock2Flag = False
If Len(Tmp) = 0 Or Tmp = "0" Then
    NoBlock2Flag = True
    optBlock(1).Enabled = False
    Block2Start = Block1Start
Else
    Block2Start = "&H" & Tmp
End If

ReDim xTmp(0)
Get #bFileIn, Block2Start + 7, xTmp()
VortexBlock2Height = xTmp(0)

Tmp = VBGetPrivateProfileString(Hardtext, "Block2End", iniFileName)
If Len(Tmp) = 0 Then
    Block2End = Block1End
Else
    Block2End = "&H" & Tmp
End If

If Block1Flag Then    ' block 1, second in firmare
'
Else    'block 2 first in firmware
    If NoBlock2Flag Then
        MsgBoxEx ArrMsg(2), , , CenterOwner    '"No block 2 specified in INI file."
        Exit Sub
    End If
End If

ReDim bTmp(lngBytes - 1)    '(lngBytes - 1)
'ReDim xTmp(lngBytes - 1)
cmbAdr.Clear
ReDim FontBlock1Arr(0)
ReDim FontBlock1IndArr(0) 'adr of pointer to current font
ReDim FontBlock2Arr(0)
ReDim FontBlock2IndArr(0)
ReDim FontBlock1VortexWidthArr(0)
ReDim FontBlock2VortexWidthArr(0)

'Seek #bFileIn, 1
Get #bFileIn, 1, bTmp()

If NoBlock2Flag Then k = 1 Else k = 2

Ind = Block1Start + 9
endBlock = Block1End

For i = 1 To k
    n = 0
    Do While Ind <= endBlock   ' lngBytes

        'If bTmp(2) = 0 Then
        lTmp = bTmp(Ind) + 256 * bTmp(Ind + 1)
            
'        Else
'            sHex = Hex(bTmp(2)) & right$("0" & Hex(bTmp(1)), 2) & right$("0" & Hex(bTmp(0)), 2)
'        End If

        If i = 1 Then    '1 block
            ReDim Preserve FontBlock1Arr(n)
            FontBlock1Arr(n) = Hex(Block1Start + lTmp)
            ReDim Preserve FontBlock1IndArr(n)
            FontBlock1IndArr(n) = Ind    'adr of pointer to current font
            ReDim Preserve FontBlock1VortexWidthArr(n)
            FontBlock1VortexWidthArr(n) = bTmp(Ind - 1)
        Else    '2 block
            ReDim Preserve FontBlock2Arr(n)
            FontBlock2Arr(n) = Hex(Block2Start + lTmp)
            ReDim Preserve FontBlock2IndArr(n)
            FontBlock2IndArr(n) = Ind
            ReDim Preserve FontBlock2VortexWidthArr(n)
            FontBlock2VortexWidthArr(n) = bTmp(Ind - 1)
        End If
        n = n + 1
        Ind = Ind + 4
    Loop

    Ind = Block2Start + 9
    endBlock = Block2End
Next i

If Block1Flag Then    '1 block to combo
    ReDim bTmpCollection(UBound(FontBlock1Arr))
    For i = 0 To UBound(FontBlock1Arr)
        cmbAdr.AddItem FontBlock1Arr(i)
        cmbAdr.ItemData(i) = i + 32
    Next i
Else
    ReDim bTmpCollection(UBound(FontBlock2Arr))
    For i = 0 To UBound(FontBlock2Arr)
        cmbAdr.AddItem FontBlock2Arr(i)
        cmbAdr.ItemData(i) = i + 1
    Next i
End If

'cmbAdr.ListIndex = cmbLastIndex    'cmbAdr_Click > startAddr = cmbAdr.text : Call GetBlock

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillFontList_Vortex()"
End Sub


Private Sub PasteFontData2bArr(sRowIn As Long, sColIn As Long, sRowCurrent As Long, sColCurrent As Long)
'from Paste cmd
'fill barr from FontData.text
Dim i As Integer, j As Integer, X As Long, n As Integer
Dim bArrTmp() As Byte
On Error GoTo frmErr
'Debug.Print ">   PasteFontData2bArr"

ReDim bArrTmp(sRowIn - 1, sColIn - 1)
If shiftFlag Then
    ReDim Preserve bArr(sRowCurrent - 1, sColCurrent - 1)
Else
    ReDim bArr(sRowCurrent - 1, sColCurrent - 1)
End If

X = 1
n = intBytes(sColIn)
For i = 0 To sRowIn - 1
    For j = 0 To sColIn - 1
'bArrTmp(i, j) = Val(Mid(FontData.Text, x, 1))
        bArrTmp(i, j) = Val(FontData.Char(X))
        X = X + 1
    Next
    X = X + n - sColIn
Next

For i = 0 To sRowIn - 1
    For j = 0 To sColIn - 1
        If i <= UBound(bArr, 1) And j <= UBound(bArr, 2) Then
            If shiftFlag Then
                If bArrTmp(i, j) = 1 Then bArr(i, j) = 1
            Else
                bArr(i, j) = bArrTmp(i, j)
            End If
        End If
    Next j
Next i

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": PasteFontData2bArr()"
End Sub

Private Sub PasteFontData2bArr_Selection(ByRef sClipFont As String, ByRef X1 As Integer, ByRef X2 As Integer, ByRef Y1 As Integer, ByRef Y2 As Integer)
'from Paste cmd
'fill barr from FontData.text
Dim i As Integer, j As Integer, X As Long, n As Integer
Dim k As Integer, m As Integer
Dim bArrTmp() As Byte
Dim sRowIn As Long
Dim sColIn As Long
Dim s() As String
Dim sRowCurrent As Integer
Dim sColCurrent As Integer
Dim sClipBytes() As String
Dim SelInd As Integer

On Error GoTo frmErr

sClipBytes = Split(sClipFont, ",")

'check valid
s = Split(sClipFont, ",")
If UBound(s) = 0 Then Exit Sub
If Not IsNumeric(s(0)) Then Exit Sub

sRowIn = Val(sClipBytes(2))
sColIn = Val(sClipBytes(1))

FontData.reset
For i = 3 To UBound(sClipBytes)
    FontData.concat dec2bin(Val(sClipBytes(i)))    'this FontData - simple, without formatting for block and headers
Next i

sRowCurrent = sRow
sColCurrent = sCol

ReDim bArrTmp(sRowIn - 1, sColIn - 1)
'If shiftFlag Then
ReDim Preserve bArr(sRowCurrent - 1, sColCurrent - 1)
'Else
'    ReDim bArr(sRowCurrent - 1, sColCurrent - 1)
'End If

X = 1
n = intBytes(sColIn)
'For i = Y1 To Y1 + sRowIn - 1
'    For j = X1 To X1 + sColIn - 1
'        bArr(i, j) = Val(FontData.Char(X))
'        X = X + 1
'    Next
'    X = X + n - sColIn
'Next

For i = 0 To sRowIn - 1
    For j = 0 To sColIn - 1
        bArrTmp(i, j) = Val(FontData.Char(X))
        X = X + 1
    Next
    X = X + n - sColIn
Next

For i = Y1 To Y2 - 1
    For j = X1 To X2 - 1
        If i <= UBound(bArr, 1) And j <= UBound(bArr, 2) Then
            If k <= UBound(bArrTmp, 1) And m <= UBound(bArrTmp, 2) Then
                If shiftFlag Then
                    If bArrTmp(k, m) = 1 Then bArr(i, j) = 1
                Else
                    bArr(i, j) = bArrTmp(k, m)
                End If
            End If
        End If
        m = m + 1
    Next j
    k = k + 1: m = 0
Next i

SelInd = cmbAdr.ListIndex
Call StoreAfterPaste(SelInd)

If fFileOpen Then

    If VortexMod Then
        Call GetArray_Vortex(SelInd)
    Else
        Call GetArray(SelInd)
    End If

    
Else
    Call bArr2PicDraw(0)    'fill draw box with bArr data
    Call bArr2PicReal(0)
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": PasteFontData2bArr_Selection()"
End Sub

Private Sub FileFontData2bArr(ByRef SelInd As Integer)
Dim i As Long, j As Long, X As Long, k As Long, n As Long, r As Long
Dim sRowCurrent As Long
Dim sColCurrent As Long
On Error GoTo frmErr
'Debug.Print ">   FileFontData2bArr"

sRowCurrent = sRowArr(SelInd)
sColCurrent = sColArr(SelInd)
If sRowCurrent = 0 Then sRowCurrent = 1
If sColCurrent = 0 Then sColCurrent = 1
'Exit Sub

ReDim bArr(sRowCurrent - 1, sColCurrent - 1)    'fill this from FontData.text

If Block1Flag Then
    X = 1
    n = intBytes(sRowCurrent)
'n = sRowCurrent / 8
    n = n / 8
    For k = 1 To n
        r = 8 * k
        For j = 0 To sColCurrent - 1
            For i = r - 1 To r - 8 Step -1
' If i <= UBound(bArr, 1) Then bArr(i, j) = Val(Mid(FontData.Text, x, 1))
                If i <= UBound(bArr, 1) Then bArr(i, j) = Val(FontData.Char(X))
' Debug.Print Val(Mid(FontData.Text, x, 1)), FontData.Char(x)

                X = X + 1
            Next
        Next
'Debug.Print "FileFontData2bArr" & vbCrLf: Command1_Click
    Next
Else
    X = 1
    n = intBytes(sColCurrent)
    For i = 0 To sRowCurrent - 1
        For j = 0 To sColCurrent - 1
'bArr(i, j) = Val(Mid(FontData.Text, x, 1))
            bArr(i, j) = Val(FontData.Char(X))
            X = X + 1
        Next
        X = X + n - sColCurrent
    Next
End If


'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FileFontData2bArr()"
End Sub


Private Sub FileFontData2bArr_Vortex(ByRef SelInd As Integer)
Dim i As Long, j As Long, X As Long, k As Long, n As Long, r As Long
Dim sRowCurrent As Long
Dim sColCurrent As Long
Dim arr() As Byte

On Error GoTo frmErr
'Debug.Print ">   FileFontData2bArr"

sRowCurrent = sRowArr(SelInd)
sColCurrent = sColArr(SelInd)
If sRowCurrent = 0 Then sRowCurrent = 1
If sColCurrent = 0 Then sColCurrent = 1
'Exit Sub

    ReDim bArr(sRowCurrent - 1, sColCurrent - 1)    'fill this from FontData.text
    
If Block1Flag Then

    X = 1
    n = intBytes(sColCurrent)

    For i = 0 To sRowCurrent - 1
        For j = 0 To sColCurrent - 1

            bArr(i, j) = Val(FontData.Char(X))
            X = X + 1

        Next j
        X = X + n - sColCurrent
    Next i
    
Else


    X = 1
    n = intBytes(sRowCurrent)

    For j = 0 To sColCurrent - 1

        For i = 0 To sRowCurrent - 1

            bArr(i, j) = Val(FontData.Char(X))
            'Debug.Print dec2binByte(arr(j, i));
            X = X + 1
            
        Next i
        X = X + n - sRowCurrent
    Next j

End If
'bArr(i, j) = Val(Mid(FontData.Text, x, 1))


'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FileFontData2bArr_V()"
End Sub



Private Sub bArr2HEXCharData(ByRef SelInd As Integer)
'fill CharDataHEXArr(i)
'myevic font format

Dim i As Long, j As Long, k As Long, n As Long, r As Long
Dim Tmp As New CString
Dim sZero As String
Dim e As Long
Dim sRowCurrent As Long
Dim sColCurrent As Long
On Error GoTo frmErr
'Debug.Print ">   bArr2HEXCharData"

If Not fFileOpen Then
    SelInd = 0
    ReDim CharDataHEXArr(0)
End If

sRowCurrent = sRowArr(SelInd)
sColCurrent = sColArr(SelInd)

sCharData.reset
sCharData.concat "{" & CStr(sColCurrent) & ","
sCharData.concat CStr(sRowCurrent) & ",{"

'по столбцам, снизу вверх по рядам по слолбцов/8

For n = 0 To sRowCurrent \ 8 - 1

    For i = 0 To sColCurrent - 1
        For j = 7 + 8 * n To 8 * n Step -1
            Tmp.concat bArr(j, i)
        Next j
        sCharData.concat Bin2Dec(Tmp.Text) & ","
        Tmp.reset
    Next i
Next n

Tmp.concat left$(sCharData.Text, Len(sCharData.Text) - 1)    '- ","
Tmp.concat "}}"

CharDataHEXArr(SelInd) = Tmp.Text

'Set Tmp = Nothing
'''
Exit Sub
frmErr:
MsgBox Err.Description & ": bArr2HEXCharData()"
End Sub

Private Sub bArr2CharData(ByRef SelInd As Integer)
'and arr
Dim i As Long, j As Long, k As Long, n As Long, r As Long
Dim Tmp As New CString
Dim sZero As String
Dim e As Long
Dim sRowCurrent As Long
Dim sColCurrent As Long
On Error GoTo frmErr
'Debug.Print ">   bArr2CharData"

sRowCurrent = sRowArr(SelInd)
sColCurrent = sColArr(SelInd)

sCharData.reset
sCharData.concat CStr(sColCurrent) & ","
sCharData.concat CStr(sRowCurrent) & ","

n = intBytes(sColCurrent)
sZero = Space$(n - sColCurrent)
sZero = Replace(sZero, mySpace, "0")

For i = 0 To sRowCurrent - 1
    For k = 0 To n / 8 - 1
        r = 8 * k
        e = 8
        If sColCurrent - r < 8 Then e = sColCurrent - r
        For j = r To r + e - 1     '17   8 + 8 + 1
            Tmp.concat bArr(i, j)
        Next j
        If e < 8 Then Tmp.concat sZero    '1+7
        sCharData.concat Bin2Dec(Tmp.Text) & ","
        Tmp.reset
    Next k
Next i

Tmp.concat left$(sCharData.Text, Len(sCharData.Text) - 1)    '- ","
sCharData.reset
sCharData.concat Tmp.Text


If fFileOpen Then
    CharDataArr(SelInd) = sCharData.Text
End If

'Set Tmp = Nothing
'''
Exit Sub
frmErr:
MsgBox Err.Description & ": bArr2CharData()"
End Sub




Private Sub Selection_bArr2sClipFont(ByRef strData As String, ByRef X1 As Integer, ByRef Y1 As Integer, ByRef X2 As Integer, ByRef Y2 As Integer)
'and arr
Dim i As Long, j As Long, k As Long, n As Long, r As Long
Dim Tmp As New CString
Dim sZero As String
Dim e As Long
Dim sRowCurrent As Long
Dim sColCurrent As Long
On Error GoTo frmErr
'Debug.Print ">   bArr2CharData"

sRowCurrent = Y2 - Y1
sColCurrent = X2 - X1

sCharData.reset
sCharData.concat CStr(sColCurrent) & ","
sCharData.concat CStr(sRowCurrent) & ","

n = intBytes(sColCurrent)
sZero = Space$(n - sColCurrent)
sZero = Replace(sZero, mySpace, "0")

For i = Y1 To Y1 + sRowCurrent - 1
    For k = 0 To n / 8 - 1
        r = 8 * k
        e = 8
        If sColCurrent - r < 8 Then e = sColCurrent - r
        For j = X1 + r To X1 + r + e - 1  '17   8 + 8 + 1
            Tmp.concat bArr(i, j)
        Next j
        If e < 8 Then Tmp.concat sZero    '1+7
        sCharData.concat Bin2Dec(Tmp.Text) & ","
        Tmp.reset
    Next k
Next i

Tmp.concat left$(sCharData.Text, Len(sCharData.Text) - 1)    '- ","
strData = "0," & Tmp.Text

'Set Tmp = Nothing
'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Selection_bArr2sClipFont()"
End Sub
Private Sub bArr2PicDraw(ByRef SelInd As Integer)
'fill draw box with bArr data
Dim i As Long, j As Long
'Dim ptColor As Long
Dim Xm As Integer, Ym As Integer
Dim sRowCurrent As Integer
Dim sColCurrent As Integer
Dim g As Integer
On Error GoTo frmErr
'Debug.Print ">   bArr2PicDraw"

If SelInd = -1 Then    'no file
    sRowCurrent = sRow
    sColCurrent = sCol
Else
    sRowCurrent = sRowArr(SelInd)
    sColCurrent = sColArr(SelInd)
   ' sRowCurrent = sColArr(SelInd)
   ' sColCurrent = sRowArr(SelInd)
End If

If chkGridFlag Or chkSelection.Value = vbChecked Then g = 1


'picContainer.Visible = False


With picContainer
    Set .Picture = Nothing
    .Height = sRowCurrent * XYspace + g
    .Width = sColCurrent * XYspace + g

'.Cls
End With
'
'            Dim pt As POINTAPI
'    Dim hOldPen As Long, hPen As Long
'
'    Dim logBR As LOGBRUSH
'    With logBR
'        .lbColor = lcForeColor
'        .lbStyle = 0
'        .lbHatch = 0&
'    End With
'
'   hPen = ExtCreatePen(PS_GEOMETRIC Or PS_ENDCAP_SQUARE Or PS_SOLID, XYspace, logBR, 0, ByVal 0&)

' picContainer.DrawWidth = XYspace
'             hPen = CreatePen(PS_SOLID, XYspace, lcForeColor)
'    hOldPen = SelectObject(picContainer.hdc, hPen)
'    DeleteObject SelectObject(picContainer.hdc, CreatePen(0, 5, lcForeColor))

For i = 0 To sRowCurrent - 1
    For j = 0 To sColCurrent - 1
        If bArr(i, j) = 1 Then
            Xm = j * XYspace
            Ym = i * XYspace
' picContainer.PSet (Xm, Ym), lcForeColor
'    RectangleX picContainer.hdc, Xm + XYspace, Ym, Xm + XYspace, Ym + XYspace
'     SetPixelV picContainer.hdc, Xm, Ym, lcForeColor
'   DeleteObject SelectObject(picContainer.hdc, hOldPen)
            picContainer.Line (Xm, Ym)-(Xm + XYspace - 1, Ym + XYspace - 1), lcForeColor, BF
        End If
    Next j
Next i

'picContainer.DrawWidth = 1

'  ptColor = lcForeColor

'         Xm = j * XYspace - XYspace \ 2
'        Ym = i * XYspace - XYspace \ 2
''            Dim pt As POINTAPI
'            MoveToEx picContainer.hdc, Xm, Ym, pt
'     LineTo picContainer.hdc, 1 + Xm, Ym
''RectangleX picContainer.hDC, Xm, Ym, Xm + 5, Ym

'SetPixel picContainer.hdc, Xm, Ym, lcForeColor

'    RectangleX picContainer.hdc, Xm, Ym, Xm + XYspace, Ym + XYspace
'    MoveToEx picContainer.hdc, Xm, 0, pt
'     LineTo picContainer.hdc, Xm, picContainer.ScaleHeight - 1
'         MoveToEx picContainer.hdc, 0, Ym, pt
'     LineTo picContainer.hdc, picContainer.ScaleWidth - 1, Ym

' LineTo picContainer.hdc, Xm + XYspace, Ym ' + XYspace

'   DeleteObject SelectObject(picContainer.hdc, hOldPen)

'        Else
'            ptColor = lcBackColor


'DeleteObject SelectObject(picContainer.hdc, CreatePen(0, XYspace, ptColor))
'RectangleX picContainer.hdc, Xm, Ym, Xm + XYspace, Ym + XYspace

'    Next j
'Next i

' DeleteObject SelectObject(picContainer.hdc, hOldPen)

If Not DrawWordFlag Then
    If chkGridFlag = vbChecked Then DrawGrid picContainer, vbGrayText
    If isSelection Then Call DrawOldSelRect

    Call Form_Resize    '2 for correct
End If

'picContainer.Visible = True

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": bArr2PicDraw()"
End Sub
Private Sub ClipBytesParse(ByRef sClipFont As String, ByRef SelInd As Integer, Optional f_respack As Boolean = False)
Dim i As Integer
Dim sClipBytes() As String
Dim Ret As Long
Dim sColIn As Long
Dim sRowIn As Long
Dim SelIndFromClip As Integer    '1-st byte in sClipChars
Dim arr_Ind() As Integer
Dim BaseI As Integer

On Error GoTo frmErr
'Debug.Print ">   ClipBytesParse"

sClipBytes = Split(sClipFont, ",")
If UBound(sClipBytes) = 0 Then Exit Sub

If SelInd < 0 Then    ' draw without open file
    'ReDim Preserve sClipBytes(0) 'one first char if many
    ReDim sColArr(0)
    sColArr(0) = sCol
    ReDim sRowArr(0)
    sRowArr(0) = sRow
    SelInd = 0
    PasteByNumber = False
Else
    PasteByNumber = chkByNumber.Value
    SelIndFromClip = Val(sClipBytes(0))
    If PasteByNumber Then SelInd = SelIndFromClip - 1
    If SelInd < 0 Then Exit Sub    'invalid index in text data

    'no here    If PasteByNumber Then
    '    If SelInd <> cmbAdr.ListIndex Then Exit Sub 'or paste current to selind
    '    End If

End If

sRowIn = Val(sClipBytes(2))
sColIn = Val(sClipBytes(1))

If VortexMod Then BaseI = SelInd + 31 Else BaseI = SelInd

If sColArr(SelInd) <> sColIn Or sRowArr(SelInd) <> sRowIn Then
    ReDim arr_Ind(0)
    Call GetAllIndexesOfSameGlyph(SelInd, arr_Ind)

    If CheckCharSizeFlag Then
        Call McListBox1.ViewBoldItem(SelInd)    'point to problem char
        McListBox1.Refresh

        Ret = MsgBoxEx("(&h_" & Hex(BaseI + 1) & ") " & ArrMsg(3) & vbCrLf & sColIn & "x" & sRowIn & " -> " & sColArr(SelInd) & "x" & sRowArr(SelInd) & vbCrLf & vbCrLf & ArrMsg(24), , , CenterOwner, vbYesNoCancel Or vbQuestion)

        Select Case Ret
        Case 6  'yes change size
            For i = 0 To UBound(arr_Ind)
                'sRowArr(SelInd) = sRowIn
                'sColArr(SelInd) = sColIn
                sRowArr(arr_Ind(i)) = sRowIn
                sColArr(arr_Ind(i)) = sColIn
            Next i
            sCol = sColIn
            sRow = sRowIn
        Case 7    'no only paste
        Case Else
            Exit Sub
        End Select

    Else    ' yes change size

        For i = 0 To UBound(arr_Ind)
            'sRowArr(SelInd) = sRowIn
            'sColArr(SelInd) = sColIn
            sRowArr(arr_Ind(i)) = sRowIn
            sColArr(arr_Ind(i)) = sColIn
        Next i

        sCol = sColIn
        sRow = sRowIn

    End If
End If

FontData.reset
If f_respack Then
    FontData.concat sClipBytes(3)
Else
    For i = 3 To UBound(sClipBytes)
        'If VortexMod Then
        '    FontData.concat StrReverse(dec2bin(Val(sClipBytes(i))))
        'Else
            FontData.concat dec2bin(Val(sClipBytes(i)))    'this FontData - simple, without formatting for block and headers
        'End If
    Next i
End If

Call PasteFontData2bArr(sRowIn, sColIn, sRowArr(SelInd), sColArr(SelInd))    'fill bArr from FontData.text

Call StoreAfterPaste(SelInd)

'''
Exit Sub
frmErr:
'LockWindowUpdate 0
McListBox1.Visible = True
Me.MousePointer = vbNormal
MsgBox Err.Description & ": ClipBytesParse()"
End Sub

Private Sub StoreAfterPaste(ByRef SelInd As Integer)
'no draw to container
' some diff to Private Sub StoreCurrentChar() TODO?
Dim arr_Ind() As Integer
Dim i As Integer

On Error GoTo frmErr

If fFileOpen Then

    McListBox1.Visible = False
    ReDim arr_Ind(0)
    Call GetAllIndexesOfSameGlyph(SelInd, arr_Ind)

    '  no!  For i = 0 To UBound(arr_Ind)
    '        'sRowArr(SelInd) = sRow
    '        'sColArr(SelInd) = sCol
    '        sRowArr(arr_Ind(i)) = sRow
    '        sColArr(arr_Ind(i)) = sCol
    '    Next i

    If PicNotEqual(SelInd) Then

        '        ChangesIndArr(SelInd) = True
        '        McListBox1.ListBold(SelInd) = True

        For i = 0 To UBound(arr_Ind)
            ChangesIndArr(arr_Ind(i)) = True
            McListBox1.ListBold(arr_Ind(i)) = True
        Next i

        If FirstPasteByNum < 0 Then FirstPasteByNum = SelInd    'first changed in last paste

    Else

        '        ChangesIndArr(SelInd) = False
        '        McListBox1.ListBold(SelInd) = False

        For i = 0 To UBound(arr_Ind)
            ChangesIndArr(arr_Ind(i)) = False
            McListBox1.ListBold(arr_Ind(i)) = False
        Next i

    End If

    cmdSaveAll.Caption = ArrMsg(12) & " (" & GetChangesCount & ")"

    Call bArr2bTmp(sRowArr(SelInd), sColArr(SelInd))

    '    bTmpCollection(SelInd) = bTmp
    '    Call bTmp2FontData(SelInd)
    '    Call bTmp2CharData(SelInd)

    For i = 0 To UBound(arr_Ind)
        bTmpCollection(arr_Ind(i)) = bTmp

        If VortexMod Then
            Call bTmp2FontData_Vortex(arr_Ind(i))
'            Call bArr2CharData_Vortex(arr_Ind(i))
        Else
            Call bTmp2FontData(arr_Ind(i))
 '           Call bArr2CharData(arr_Ind(i))
        End If
Call bArr2CharData(arr_Ind(i))

    Next i

    McListBox1.Visible = True

Else    'no file
    ReDim sColArr(0)
    ReDim sRowArr(0)

    sColArr(0) = sCol
    sRowArr(0) = sRow

    SelInd = 0

    Call bArr2CharData(0)

End If

Exit Sub
frmErr:
McListBox1.Visible = True
MsgBox Err.Description & ": DrawAfterPaste()"
End Sub
Private Sub cmdLoadFont_Click()
Dim filename As String
Dim FileTitle As String
On Error GoTo frmErr

filename = vbNullString
filename = fLoadDialog(ArrMsg(16), FileTitle)
If filename = vbNullString Then Exit Sub
LastPath = GetPathFromPathAndName(filename)

Set m_Preview = New CFontPreview
m_Preview.FontFile = filename
If Len(m_Preview.FaceName) = 0 Then Exit Sub

TTFontName = m_Preview.FaceName
TTFontBold = m_Preview.Bold
TTFontItalic = m_Preview.Italic
TTFontUnderline = m_Preview.Underlined

TTF_Char = cmbTTF_Char.Text

cmb_SysFonts.Visible = False
Call FillListWithFonts(cmb_SysFonts)
cmb_SysFonts.Text = TTFontName
cmb_SysFonts.Visible = True

Call cmbTTF_Char_Click

'TTF_Size = sCol * XYspace

'LoadFontFlag = True
'VScroll_Y.Max = 150
'VScroll_Y.Min = -350
'HScroll_X.Max = 250
'HScroll_X.Min = -150
'VScroll_S.Max = 1
'VScroll_S.Min = 1000
'VScroll_Y.value = 0
'HScroll_X.value = 0
'VScroll_S.value = TTF_Size    '250
'LoadFontFlag = False


'Call StoreInUndoBuffer

'TTFontDraw TTF_Char, TTF_Size, 0, 0, TTFontBold, TTFontItalic
'      If Len(m_Preview.FaceName) Then
'         Me.Caption = App.Title & ": " & m_Preview.FaceName
'         Set fnt = New StdFont
'         Set Picture1.Font = fnt
'         fnt.Name = m_Preview.FaceName
'         fnt.Bold = m_Preview.Bold
'         fnt.Italic = m_Preview.Italic
'         Sizes = Array(60, 48, 36, 24, 18, 14, 12, 10, 8)
'         Picture1.Cls
'         For i = LBound(Sizes) To UBound(Sizes)
'            fnt.Size = Sizes(i)
'            Picture1.Print Pangram; " ("; Sizes(i); ")"
'         Next i
'         fnt.Size = 24
'         Picture1.Print "0123456789!@#$%^&*()~-_+=:;""',<.>/?"
'      Else
'         'Me.Caption = App.Title
'      End If

'     Set Picture1.Font = Me.Font
'      Set str = New CStringBuilder
'      str.Append "FaceName: " & m_Preview.FaceName & vbCrLf
'      str.Append "Family Name: " & m_Preview.FamilyName & vbCrLf
'      str.Append "Subfamily Name: " & m_Preview.SubFamilyName & vbCrLf
'      str.Append "Full Name: " & m_Preview.FullName & vbCrLf
'      str.Append "Unique Identifier: " & m_Preview.UniqueIdentifier & vbCrLf
'      str.Append "Postscript Name: " & m_Preview.PostscriptName & vbCrLf
'      str.Append "Copyright: " & m_Preview.Copyright & vbCrLf
'      str.Append "Trademark: " & m_Preview.Trademark & vbCrLf
'      str.Append "Version: " & m_Preview.VersionString & vbCrLf
'      str.Append "Installed: " & m_Preview.Installed & vbCrLf
'      Picture1.Print str.ToString



'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdLoadFont_Click"
End Sub

Private Sub TTFontDraw(fChar As String, fSize As Integer, fX As Single, fY As Single, fBold As Boolean, fItalic As Boolean, fUnderline As Boolean)
On Error GoTo frmErr

If Len(TTFontName) = 0 Then Exit Sub

If fSize = 0 Then Exit Sub
picContainer.Cls
If chkGridFlag = vbChecked Then DrawGrid picContainer, vbGrayText

With picContainer
    .Font.Name = TTFontName
    .FontSize = fSize

    .FontBold = fBold
    .FontItalic = fItalic
    .FontUnderline = fUnderline
    .CurrentX = fX
    .CurrentY = fY
    .FontTransparent = True

End With
picContainer.Print fChar

Call PicDraw2PicReal(-1)    '1 mapping font to real bitmap
Call Draw2bArr(PicReal, sCol, sRow)  '2 write to array bArr
Call StoreCurrentChar(cmbAdr.ListIndex)
Call bArr2PicDraw(-1)

Call StoreInUndoBuffer

'VScroll_S_Value = fSize
'VScroll_Y_Value = fY
'HScroll_X_Value = fX

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": TTFontDraw"
End Sub

Private Sub cmdPaste_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = 1 Then
    shiftFlag = True
Else
    shiftFlag = False
End If
End Sub

Private Sub cmdPatcher_Click()

If Not fFileOpen Then cmdReloadFW_Click

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If

frmPatch.Show 1, frmmain

End Sub

Private Sub cmdReloadFW_Click()
Dim Ret As Long

If fFileOpen Then
    If GetChangesCount > 0 Then
        Ret = MsgBoxEx(ArrMsg(23), , , CenterOwner, vbOKCancel Or vbQuestion)    'unsaved
        If Ret <> 1 Then Exit Sub
    End If
End If

'FileNameFW
reloadFW_flag = True 'for setup old position in list

If FileNameFW = vbNullString Then
    If LastOpenedFW = vbNullString Then
        Exit Sub
    Else
        FileNameFW = LastOpenedFW
        'FileTitle = GetNameExt(FileNameFW)
    End If
    'Else
    'FileTitle = GetNameExt(FileNameFW)
End If

'encrypt if decrypted
If Not EncryptFW Then
    Ret = MsgBoxEx(ArrMsg(19), , , CenterOwner, vbCritical)
        'Close #bFileIn
        fFileOpen = False
        Exit Sub    'unknown FW
End If

Call LoadFWfile

reloadFW_flag = False

End Sub



Private Sub cmdResize_Click()
'Dim w As Long
Dim h As Long
'Dim ratioWidth As Single, ratioHeight As Single,
Dim ratio As Single
On Error GoTo frmErr
If PicReal.Picture = 0 Then Exit Sub

    'Calgulate AspectRatio
    ratio = (QwickResizeWidth / PicReal.Width)
    'ratioHeight = (Height / PicReal.Height)

    'Calgulate newWidth and newHeight
    'w = PicReal.Width * ratio
    h = PicReal.Height * ratio
    
'ratio = PicReal.Width / QwickResizeWidth
'h = PicReal.Height / ratio
PicReal.Width = QwickResizeWidth
PicReal.Height = h
PicReal.PaintPicture PicReal.Picture, 0, 0, QwickResizeWidth, h


'PicReal.Picture = PicReal.Image
            sCol = PicReal.Width
            sRow = PicReal.Height
            Call XYcaptionSet(sCol, sRow)
            ReDim bArr(sRow - 1, sCol - 1)
Call Draw2bArr(PicReal, sCol, sRow)
Call StoreCurrentChar(cmbAdr.ListIndex)
Call bArr2PicDraw(-1)
Call bArr2PicReal(-1)
Call StoreInUndoBuffer

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdResize"
End Sub

Private Sub cmdScale_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo frmErr

If Button <> vbRightButton Then Exit Sub

Select Case Index
Case 0
If XYspace > 2 Then
    XYspace = XYspace \ 2
    If XYspace < 2 Then
        chkGridFlag = vbUnchecked
    End If
End If
Case 1
If XYspace * 2 <= 40 Then
    XYspace = XYspace * 2
    If XYspace >= 2 Then
        chkGridFlag = chkGrid.Value
    End If
End If
End Select

TTF_Size = sCol * XYspace
'TTF_Size = TTF_Size + XYspace

With picContainer
    .Cls
    .Height = sRow * XYspace
    .Width = sCol * XYspace
End With

Call bArr2PicDraw(-1)    '1
If PicTTF.Visible Then
    PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2
End If

'chkGridFlag = chkGrid.Value
If chkGridFlag = vbChecked Then DrawGrid picContainer, vbGrayText

Form_Resize
'Call SetUpPicScroll
'Call SetUpScrollBars

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdScale_MouseDown()"
End Sub

Private Sub cmdShowAllDict_Click()
On Error GoTo frmErr

If Not fFileOpen Then Exit Sub
If NoVocabFlag Then Exit Sub

PicReal.Visible = False
picX3.Visible = False
Call DrawAllWords
'Call PicReal_Change
'picX3.Visible = True
PicReal.Visible = True

'''
Exit Sub
frmErr:
PicReal.Visible = True
MsgBox Err.Description & ": ShowAllDict"
End Sub

Private Sub cmdToolBar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next 'GoTo frmErr
If Button = 2 Then

    Select Case Index
    Case 12    'import from txt file
        Call LoadImport
        Exit Sub

    Case 13    'export to txt file
        Call SaveExport(False)
        Exit Sub
    End Select
End If

End Sub

Private Sub cmdUndoRedo_Click(Index As Integer)
On Error GoTo frmErr

Select Case Index

Case 0    'undo
    If UndoClicksCount < UBound(UndoBuffer) Then
        If picCount > -1 Then
            picCount = IIf(picCount = LBound(UndoBuffer), UBound(UndoBuffer), picCount - 1)
            'Debug.Print "undo to " & picCount
            'PicReal.Width = UndoBuffer(picCount).Width
            'PicReal.Height = UndoBuffer(picCount).Height
            PicReal.AutoSize = True
            Set PicReal.Picture = UndoBuffer(picCount)
            PicReal.AutoSize = False
            'PicReal.Picture = PicReal.Image
            UndoClicksCount = UndoClicksCount + 1
            'sCol = sColArr(cmbLastIndex): sRow = sRowArr(cmbLastIndex)
            sCol = PicReal.Width
            sRow = PicReal.Height
            Call XYcaptionSet(sCol, sRow)
            ReDim bArr(sRow - 1, sCol - 1)    'if rotate was we lost right dim - restore from pic
            Call Draw2bArr(PicReal, sCol, sRow)
            Call StoreCurrentChar(cmbAdr.ListIndex)
            Call bArr2PicDraw(-1)
            'Call XYcaptionSet(sCol, sRow)
            '
            'Call bArr2PicDraw(-1)    '0
            'Call PicDraw2PicReal(-1)    '1
            'Call Copy2PicRectSel    '2
            'Call StoreCurrentChar


        End If
    Else
        If UndoClicksCount = 0 Then
            If fFileOpen Then
                If McListBox1.ListBold(cmbLastIndex) = True Then
                    Call UndoBufferClear
                    Call StoreInUndoBuffer
                    If VortexMod Then
                        Call GetBlock_Vortex(cmbLastIndex)
                    Else
                        Call GetBlock(cmbLastIndex)
                    End If
                    sCol = sColArr(cmbLastIndex): sRow = sRowArr(cmbLastIndex)
                    Call XYcaptionSet(sCol, sRow)
                    Call StoreCurrentChar(cmbAdr.ListIndex)
                    'Debug.Print sCol, sRow

                    'Call XYcaptionSet(sCol, sRow)
                    '
                    'Call bArr2PicDraw(-1)    '0
                    'Call PicDraw2PicReal(-1)    '1
                    'Call Copy2PicRectSel    '2
                    'Call StoreCurrentChar

                    UndoClicksCount = UndoClicksCount + 1
                End If
            End If
        End If
    End If

Case 1    ' redo
    If UndoClicksCount > 0 Then
        If picCount > -1 Then
            picCount = IIf(picCount = UBound(UndoBuffer), LBound(UndoBuffer), picCount + 1)
            '   Debug.Print "redo to " & picCount
            PicReal.AutoSize = True
            Set PicReal.Picture = UndoBuffer(picCount)
            PicReal.AutoSize = False
            'PicReal.Picture = PicReal.Image

            sCol = PicReal.Width
            sRow = PicReal.Height
            Call XYcaptionSet(sCol, sRow)
            ReDim bArr(sRow - 1, sCol - 1)

            UndoClicksCount = UndoClicksCount - 1

            Call Draw2bArr(PicReal, sCol, sRow)
            Call StoreCurrentChar(cmbAdr.ListIndex)
            Call bArr2PicDraw(-1)

        End If
    End If
End Select

'Debug.Print UndoClicksCount

If PicTTF.Visible Then
    PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdUndoRedo_Click()"
End Sub

Private Sub cmdVocabSL_Click(Index As Integer)
'import export
On Error GoTo frmErr

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If

If VortexMod And (Not Block1Flag) Then Exit Sub

Select Case Index
Case 0
    Call VocabLoad
Case 1
    Call VocabSave
End Select

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdVocabSL_Click"
End Sub

Private Sub VocabLoad()
'import vocab from file
Dim fn As String
Dim FileTitle As String
Dim ExportDir As String
Dim sClipFont As String    'addr,len,c c c...[vbcr]addr,len,c c c ...[vbcr]
Dim f As Integer
'Dim Tmp As String
'Dim sFileData As String    'xml
'Dim s() As String
Dim Ret As Long
Dim sClipChars() As String
Dim sClipBytes() As String
Dim sClipBlock() As String
Dim n As Long    ', i As Long
Dim AllOk As Boolean
Dim delim As String    ' , for C, " " for simple

On Error GoTo frmErr

f = FreeFile
ExportDir = App.Path & "\Export"

fn = vbNullString
fn = ImportLoadDialog(ExportDir, ArrMsg(17), FileTitle)
If fn = vbNullString Then Exit Sub
'LastPath = GetPathFromPathAndName(filename)

DoEvents    '2 close dialog
Me.MousePointer = vbHourglass

Open fn For Input Access Read As #f
sClipFont = Input$(LOF(f), f)    'all file into string
Close #f

'''''''''''''''''''''''''''
sClipFont = Replace(sClipFont, vbLf, vbNullString)

delim = mySpace
If InStr(1, sClipFont, "0x") Then
    sClipFont = Replace(sClipFont, "{", vbNullString)
    sClipFont = Replace(sClipFont, "}", vbNullString)
    sClipFont = Replace(sClipFont, "0x", vbNullString)
    delim = ","
End If

sClipChars = Split(sClipFont, vbCr)
Ret = UBound(sClipChars)

Me.MousePointer = vbHourglass
AllOk = True

If Ret > 0 Then

    For n = 0 To Ret

        sClipBlock = Split(sClipChars(n), ",", 3)
        If UBound(sClipBlock) = 2 Then

            'sClipBlock(0) 'addr
            'sClipBlock(1) 'len
            If Block1Flag Then    'check valid
                If sClipBlock(0) < Vocab1Start Then Me.MousePointer = vbNormal: AllOk = False: Exit For
                If sClipBlock(0) > Vocab1End Then Me.MousePointer = vbNormal: AllOk = False: Exit For
            Else
                If sClipBlock(0) < Vocab2Start Then Me.MousePointer = vbNormal: AllOk = False: Exit For
                If sClipBlock(0) > Vocab2End Then Me.MousePointer = vbNormal: AllOk = False: Exit For
            End If

            sClipBytes = Split(sClipBlock(2), delim)

            'try 2 save
            If VortexMod Then
                If Not SaveWord_Vortex(CLng(sClipBlock(0)), sClipBytes()) Then AllOk = False
            Else
                If Not SaveWord(CLng(sClipBlock(0)), sClipBytes()) Then AllOk = False
            End If

        End If

    Next n
End If

'Debug.Print "AllOk=" & AllOk
If VortexMod Then
    Call FillVocab_Vortex
Else
    Call FillVocab
End If

cmbVocab.ListIndex = CurrentWordInd
Me.MousePointer = vbNormal

'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": VocabLoad"
End Sub

Private Sub VocabSave(Optional fHex As Boolean = False)
'export vocab to file
Dim fn As String
'Dim FileTitle As String
Dim ExportDir As String
'Dim sClipFont As String
Dim f As Integer
Dim strData As String
Dim i As Integer
Dim sClipFont As New CString

On Error GoTo frmErr

If VortexMod And (Not Block1Flag) Then Exit Sub

If Not fHex Then

    If Block1Flag Then    '1 block to combo
        For i = 0 To UBound(VocBlock1Arr)
            If Word1LenArr(i) = 0 Then
                sClipFont.concat Word1StartArr(i) & "," & Word1LenArr(i) & "," & VocBlock1Arr(i) & "00" & vbCrLf
            Else
                sClipFont.concat Word1StartArr(i) & "," & Word1LenArr(i) & "," & VocBlock1Arr(i) & " 00" & vbCrLf
            End If
        Next i
    Else
        For i = 0 To UBound(VocBlock2Arr)
            If Word2LenArr(i) = 0 Then
                sClipFont.concat Word2StartArr(i) & "," & Word2LenArr(i) & "," & VocBlock2Arr(i) & "00" & vbCrLf
            Else
                sClipFont.concat Word2StartArr(i) & "," & Word2LenArr(i) & "," & VocBlock2Arr(i) & " 00" & vbCrLf
            End If
        Next i
    End If

Else

    If Block1Flag Then    '1 block to combo
        For i = 0 To UBound(VocBlock1Arr0x)
            sClipFont.concat Word1StartArr(i) & "," & Word1LenArr(i) & "," & "{" & VocBlock1Arr0x(i) & "0}" & vbCrLf
        Next i
    Else
        For i = 0 To UBound(VocBlock2Arr0x)
            sClipFont.concat Word2StartArr(i) & "," & Word2LenArr(i) & "," & "{" & VocBlock2Arr0x(i) & "0}" & vbCrLf
        Next i
    End If

End If

strData = left$(sClipFont.Text, Len(sClipFont.Text) - 2)    '- vbCrLf

'If Len(strData) <> 0 Then 'always present
f = FreeFile
ExportDir = App.Path & "\Export"
fn = vbNullString
fn = ExportSaveDialog(ExportDir, ArrMsg(18))
If fn = vbNullString Then Exit Sub
'LastPath = GetPathFromPathAndName(FileName)

Open fn For Output Access Write As #f
Print #f, strData
Close #f
'End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": VocabSave"
End Sub

Private Sub cmdVocabSL_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Not fFileOpen Then Exit Sub

If Button = 2 Then
Select Case Index
'Case 0
'    Call VocabLoad
Case 1
    Call VocabSave(True) '0xHex
End Select

End If

End Sub

Private Sub cmdXY_Click(Index As Integer)
Dim X As Integer, Y As Integer
Dim m As Integer
Dim n As Integer
Dim iTemp() As Integer
On Error GoTo frmErr


ReDim iTemp(sRow - 1, sCol - 1)
For m = LBound(bArr, 1) To UBound(bArr, 1)      'Loop for 1st dimension 'increase
    For n = LBound(bArr, 2) To UBound(bArr, 2)  'Loop for 2nd dimension
        iTemp(m, n) = bArr(m, n)
    Next n
Next m

Select Case Index
Case 0
    X = Val(InputBox(ArrMsg(39), ArrMsg(38), sCol, Me.left + 3400, Me.top + 3300))    'Set width to:
    If X > 0 And X <> sCol Then
        sCol = X
    Else
        Exit Sub
    End If
Case 1
    Y = Val(InputBox(ArrMsg(40), ArrMsg(38), sRow, Me.left + 3900, Me.top + 3300))    'Set height to
    If Y > 0 And Y <> sRow Then
        sRow = Y
    Else
        Exit Sub
    End If
End Select

Call XYcaptionSet(sCol, sRow)

ReDim bArr(sRow - 1, sCol - 1)
'back 2 array
For m = LBound(iTemp, 1) To UBound(iTemp, 1)
    For n = LBound(iTemp, 2) To UBound(iTemp, 2)
        If m < sRow And n < sCol Then bArr(m, n) = iTemp(m, n)
    Next n
Next m

Call bArr2PicDraw(-1)    'fill draw box with bArr data
'Call PicDraw2PicReal(-1)    ' mapping font to real bitmap
Call bArr2PicReal(-1)
Call StoreCurrentChar(cmbAdr.ListIndex)
'Call StoreInUndoBuffer

If PicTTF.Visible Then
    PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdlXY_Click()"
End Sub

Private Sub XYcaptionSet(X As Long, Y As Long)
On Error Resume Next
If AlwaysHex Then
    cmdXY(0).Caption = "x" & Hex(X)
    cmdXY(1).Caption = "y" & Hex(Y)
Else
    cmdXY(0).Caption = "X" & X
    cmdXY(1).Caption = "Y" & Y
End If
On Error GoTo 0
End Sub
Private Sub PasteData(ByRef sClipFont As String, Optional f_respack As Boolean)
Dim sClipChars() As String
'Dim sClipFont As String
Dim n As Long
Dim Ret As Long
Dim SelCount As Long
Dim SelInd As Integer
Dim sListCount As Long
Dim s() As String
Dim sClipBytes() As String
Dim j As Integer

On Error GoTo frmErr

sClipFont = Replace(sClipFont, vbLf, vbNullString)

If InStr(1, sClipFont, "0x", vbTextCompare) Then
    'paste hex like 0x78,0x88,0x80,0x60,0x10,0x08,0x88,0xF0
    Call cmdPasteCfont_Click
    Exit Sub
End If

'check valid
s = Split(sClipFont, ",")
If UBound(s) = 0 Then Exit Sub
If Not IsNumeric(s(0)) Then Exit Sub

SelCount = McListBox1.SelCount
sListCount = cmbAdr.ListCount

'one, no fw
If SelCount < 1 Then
    If InStr(1, sClipFont, vbCr) Then
        sClipChars = Split(sClipFont, vbCr)
        Call ClipBytesParse(sClipChars(0), -5, f_respack)   'get one first line only
    Else
        Call ClipBytesParse(sClipFont, -5, f_respack)     'past without FW
    End If

    Call bArr2PicDraw(0)    'fill draw box with bArr data
    'Call PicDraw2PicReal(0)    ' mapping font to real bitmap
    Call bArr2PicReal(0)
    Exit Sub
End If

'mass
sClipChars = Split(sClipFont, vbCr)
Ret = UBound(sClipChars)

FirstPasteByNum = -1    'reset first changed in current paste

If Ret > 0 Then
    shiftFlag = False
    Me.MousePointer = vbHourglass

    For n = 0 To Ret    'paste many from paste data

        If Not PasteByNumber Then
            If SelCount - 1 < n Then Exit For
            SelInd = McListBox1.SelItem(n)
        End If    'else SelInd any >=0

        If n > sListCount - 1 Then Exit For    'if wrong import

        If SelCount > 1 And PasteByNumber Then    'paste by number to selected scope (if sel > 1) only
            sClipBytes = Split(sClipChars(n), ",", 2)
            If UBound(sClipBytes) <> 0 Then
                If SelCount - 1 < j Then Exit For
                SelInd = McListBox1.SelItem(j)
                If Val(sClipBytes(0) - 1) = SelInd Then
                    j = j + 1
                    If sClipChars(n) <> vbNullString Then Call ClipBytesParse(sClipChars(n), SelInd, f_respack)
                End If
            End If

        Else
            'SelInd any >=0
            If sClipChars(n) <> vbNullString Then Call ClipBytesParse(sClipChars(n), SelInd, f_respack)
        End If
    Next n
    Me.MousePointer = vbNormal
    
Else    'only one is in paste data

    If Not PasteByNumber Then SelInd = McListBox1.ListIndex    'or SelInd any >=0
'    sClipChars = Split(sClipFont, vbCr)
'sClipChars(0)
    If PasteByNumber Then
    'SelInd=0 if copy paste from other instance
        If SelInd <> 0 And SelInd <> cmbAdr.ListIndex Then Exit Sub   'or paste current to selind
    End If

    Call ClipBytesParse(sClipFont, SelInd, f_respack)

End If

If fFileOpen Then
    'lastsur    cmbAdr_Click    'redraw current view
    
    If VortexMod Then
        Call GetArray_Vortex(cmbLastIndex)
    Else
        Call GetArray(cmbLastIndex)
    End If

    If PasteByNumber Then    'point to first bold in mclist
        'For n = 0 To McListBox1.ListCount - 1
        '   If McListBox1.ListBold(n) Then
        Call McListBox1.ViewBoldItem(FirstPasteByNum)
        McListBox1.Refresh
        '      Exit For
        '   End If
        'Next n
    End If

    'Else    'no fw open
    '    Call bArr2PicDraw(SelInd)    'fill draw box with bArr data
    '    'Call PicDraw2PicReal(SelInd)    ' mapping font to real bitmap
    '    Call bArr2PicReal(SelInd)
End If


'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": PasteData"
End Sub

Private Sub cmdPaste_Click()
Dim sClipFont As String
On Error GoTo frmErr

If XYspace < 2 Then
    chkGridFlag = vbUnchecked
End If

If chkSelection.Value = vbChecked Then
    '    If picRectSel.Picture = 0 Then Exit Sub
    '    Copy2PicRealFromSelRect

    'If Clipboard.GetFormat(vbCFBitmap) Then Exit Sub

    If isSelection Then

        If Clipboard.GetFormat(vbCFBitmap) Then
        
            Set tmpPic = Nothing    'or overlay in png transparent (cool)
            tmpPic.Picture = Clipboard.GetData
            Call OrderCorners
            'PicReal.PaintPicture _
             tmpPic.Picture, _
             old_X1, old_Y1, _
             tmpPic.Width, tmpPic.Height, _
             0, 0, tmpPic.Width, tmpPic.Height, vbSrcCopy    ', vbSrcCopy And vbSrcInvert
             
            PicReal.PaintPicture _
                    tmpPic.Picture, _
                    old_X1, old_Y1, _
                    , , , , old_X2 - old_X1, old_Y2 - old_Y1, vbSrcCopy
            PicReal.Picture = PicReal.Image

            Call Draw2bArr(PicReal, sCol, sRow)
            Call StoreCurrentChar(cmbAdr.ListIndex)
            Call bArr2PicDraw(-1)

        Else

            sClipFont = Clipboard.GetText
            If Len(sClipFont) <> 0 Then
                Call PasteFontData2bArr_Selection(sClipFont, old_X1, old_X2, old_Y1, old_Y2)
            End If
        End If
    End If

Else

    If Clipboard.GetFormat(vbCFBitmap) Then
        LoadClipboardBMP
    Else
        sClipFont = Clipboard.GetText
        Call PasteData(sClipFont)

    End If

End If
XYcaptionSet sCol, sRow
Call StoreInUndoBuffer

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdPaste_Click"
End Sub

Private Sub bArr2bTmp(ByRef sRowCurrent As Long, ByRef sColCurrent As Long)
'code back like in file
'for single only
'bArr to bTmp
Dim i As Integer, j As Integer
Dim k As Integer, n As Long, r As Integer
Dim X As Integer, e As Integer
Dim Tmp As New CString
Dim s As String
On Error GoTo frmErr
'Debug.Print ">   bArr2bTmp"

If Not fFileOpen Then Exit Sub

'WTF is going on... 8(

If VortexMod Then

    '    For i = 0 To sRowCurrent - 1
    '        For j = 0 To sColCurrent - 1
    '            Debug.Print bArr(i, j);
    '        Next
    '        Debug.Print
    '    Next

    If Block1Flag Then

        n = intBytes(sColCurrent)
        ReDim bTmp(n * sRowCurrent / 8 - 1)
        s = Space$(n - sColCurrent)
        s = Replace(s, mySpace, "0")
        X = 0

        For i = 0 To sRowCurrent - 1
            For k = 0 To n / 8 - 1
                r = 8 * k
                e = 8
                If sColCurrent - r < 8 Then e = sColCurrent - r
                For j = r To r + e - 1     '17 -  8 8 1
                    If i <= UBound(bArr, 1) And j <= UBound(bArr, 2) Then
                        Tmp.concat bArr(i, j)
                    End If
                Next j
                If e < 8 Then Tmp.concat s
                bTmp(X) = Bin2Dec(StrReverse(Tmp.Text))
                X = X + 1
                Tmp.reset
            Next k
        Next i

        '        For i = 0 To UBound(bTmp)
        'Debug.Print Hex(bTmp(i))
        'Next

    Else
        n = intBytes(sRowCurrent)
        ReDim bTmp(n * sColCurrent / 8 - 1)
        s = Space$(n - sRowCurrent)
        s = Replace(s, mySpace, "0")
        X = 0

        For j = 0 To sColCurrent - 1
            For k = 0 To n / 8 - 1
                r = 8 * k
                e = 8
                If sRowCurrent - r < 8 Then e = sRowCurrent - r
                For i = r To r + e - 1     '17 -  8 8 1
                    If i <= UBound(bArr, 1) And j <= UBound(bArr, 2) Then
                        Tmp.concat bArr(i, j)
                    End If
                Next i
                If e < 8 Then Tmp.concat s
                bTmp(X) = Bin2Dec(StrReverse(Tmp.Text))
                X = X + 1
                Tmp.reset
            Next k
        Next j

    End If

Else

    If Block1Flag Then    'block1

        n = intBytes(sRowCurrent)
        ReDim bTmp(n * sColCurrent \ 8 + 1)
        bTmp(0) = sColCurrent
        bTmp(1) = sRowCurrent

        X = 2

        n = n / 8
        For k = 1 To n
            r = 8 * k
            For j = 0 To sColCurrent - 1
                For i = r - 1 To r - 8 Step -1    'row
                    If i <= UBound(bArr, 1) And j <= UBound(bArr, 2) Then
                        Tmp.concat bArr(i, j)
                    End If
                Next i
                If X > UBound(bTmp) Then
                    ReDim Preserve bTmp(UBound(bTmp) + 1)
                End If
                bTmp(X) = Bin2Dec(Tmp.Text)
                X = X + 1
                Tmp.reset
            Next j

        Next k
        ReDim Preserve bTmp(X - 1)    '? in second

    Else    ' for  block2
        n = intBytes(sColCurrent)
        ReDim bTmp(n * sRowCurrent / 8 + 1)
        bTmp(0) = sColCurrent
        bTmp(1) = sRowCurrent

        X = 2
        'n = intBytes(sColCurrent)
        s = Space$(n - sColCurrent)
        s = Replace(s, mySpace, "0")

        For i = 0 To sRowCurrent - 1
            For k = 0 To n / 8 - 1
                r = 8 * k
                e = 8
                If sColCurrent - r < 8 Then e = sColCurrent - r
                For j = r To r + e - 1     '17 -  8 8 1
                    If i <= UBound(bArr, 1) And j <= UBound(bArr, 2) Then
                        Tmp.concat bArr(i, j)
                    End If
                Next j
                If e < 8 Then Tmp.concat s
                bTmp(X) = Bin2Dec(Tmp.Text)
                X = X + 1
                Tmp.reset
            Next k
        Next i
    End If
End If

'Set Tmp = Nothing

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": bArr2bTmp()"
End Sub
Private Function CheckFontSize(ByRef sAdr As Long, ByRef xCol As Byte, ByRef yRow As Byte, ByRef SelInd As Integer, ByRef Block As Boolean) As Boolean
'xCol yRow for out
'block true - block1Flag true
Dim xInp(1) As Byte, bInp(1) As Byte
Dim Ind As Long
Dim i As Long
On Error GoTo frmErr
'Debug.Print ">   CheckFontSize"

If VortexMod Then
    If Block Then
        If sColArr(SelInd) = FontBlock1VortexWidthArr(SelInd) And sRowArr(SelInd) = VortexBlock1Height Then
            CheckFontSize = True    'same size
        Else
            CheckFontSize = False
        End If
        xCol = FontBlock1VortexWidthArr(SelInd): yRow = VortexBlock1Height
    Else
        If sColArr(SelInd) = FontBlock2VortexWidthArr(SelInd) And sRowArr(SelInd) = VortexBlock2Height Then
            CheckFontSize = True    'same size
        Else
            CheckFontSize = False
        End If
        xCol = FontBlock2VortexWidthArr(SelInd): yRow = VortexBlock2Height
    End If

Else

    Ind = sAdr
    Seek bFileIn, Ind + 1
    Get #bFileIn, , xInp()

    For i = 0 To 1
        bInp(i) = (xInp(i) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        Ind = Ind + 1
    Next i

    xCol = bInp(0): yRow = bInp(1)
    If sColArr(SelInd) = bInp(0) And sRowArr(SelInd) = bInp(1) Then
        CheckFontSize = True    'same size
    Else
        CheckFontSize = False
    End If

End If

'''
Exit Function
frmErr:
MsgBox Err.Description & ": CheckFontSize()"
End Function

Private Sub cmdPasteCfont_Click()
Dim i As Integer
Dim sClip() As String
Dim sClipFont As String
On Error GoTo frmErr

'cmbAdr.SetFocus
sClipFont = Clipboard.GetText
i = InStr(1, sClipFont, "/")
If i > 0 Then sClipFont = left$(sClipFont, i - 1)

sClipFont = Replace(sClipFont, vbTab, vbNullString)
Do While InStr(1, sClipFont, mySpace) <> 0
    sClipFont = Replace(sClipFont, mySpace, vbNullString)
Loop
sClipFont = Replace(sClipFont, "{", vbNullString)
sClipFont = Replace(sClipFont, "}", vbNullString)
sClipFont = Replace(sClipFont, vbCr, vbNullString)
sClipFont = Replace(sClipFont, vbLf, vbNullString)
sClipFont = Replace(sClipFont, "0x", vbNullString)

sClip = Split(sClipFont, ",")
sRow = "&H" & sClip(1)
sCol = "&H" & sClip(0)
FontData.reset
'only data
For i = 2 To UBound(sClip)
    FontData.concat Hex2Bin(sClip(i))
Next i

Call PasteFontData2bArr(sRow, sCol, sRow, sCol)  'fill bArr from FontData.text
Call bArr2PicDraw(-1)    'fill draw box with bArr data
'Call PicDraw2PicReal(-1)    ' mapping font to real bitmap
Call bArr2PicReal(-1)
Call XYcaptionSet(sCol, sRow)

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdPasteCfont_Click()"
End Sub
Private Sub ReloadSavedCurrent(ByRef SelInd As Integer)
Dim Tmp As String
On Error GoTo frmErr
'Debug.Print ">   ReloadSavedCurrent"

If Not fFileOpen Then Exit Sub

'Ind = cmbAdr.ListIndex
If Block1Flag Then
    startAddr = "&H" & FontBlock1Arr(SelInd)
Else
    startAddr = "&H" & FontBlock2Arr(SelInd)
End If

If VortexMod Then
    Call GetBlock_Vortex(SelInd)
Else
    Call GetBlock(SelInd)
End If


Set ImArr(SelInd) = PicReal.Picture

McListBox1.Remove SelInd
'Tmp = Right("0" & Hex(SelInd + 1), 2)

If VortexMod Then
Tmp = Hex(SelInd + 32)
'    If Block1Flag Then
'        'Tmp = SelInd + 32
'        Tmp = Hex(SelInd + 32)
'    Else
'        Tmp = Hex(SelInd + 1)
'        'Tmp = SelInd + 1
'    End If
Else
    Tmp = Hex(SelInd + 1)
End If
'Tmp = Hex(SelInd + 1)

McListBox1.AddItem Tmp, CLng(SelInd), CLng(SelInd), False

'McListBox1.Refresh
'cmbAdr_Click

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": ReloadSavedCurrent()"
End Sub

Private Function SaveChars(ByRef SelInd As Integer, sRowCurrent As Long, sColCurrent As Long) As Boolean
Dim xCol As Byte, yRow As Byte
Dim Ret As Long
Dim BaseI As Integer

On Error GoTo frmErr
'Debug.Print ">   SaveChars"

'cmbAdr.SetFocus

If VortexMod Then BaseI = SelInd + 31 Else BaseI = SelInd

If Block1Flag Then
    'check font dim
    If (Not CheckFontSize("&H" & FontBlock1Arr(SelInd), xCol, yRow, SelInd, True)) And CheckCharSizeFlag Then
        Call McListBox1.ViewBoldItem(SelInd)    'point to problem char
        McListBox1.Refresh

        Ret = MsgBoxEx("(&h_" & Hex(BaseI + 1) & ") " & ArrMsg(4) & vbCrLf & sColCurrent & "x" & sRowCurrent & " -> " & xCol & "x" & yRow & vbCrLf & vbCrLf & ArrMsg(25), , , CenterOwner, vbOKCancel Or vbQuestion)
        If Ret <> 1 Then Exit Function
    End If
    startAddr = "&H" & FontBlock1Arr(SelInd)
Else
    If (Not CheckFontSize("&H" & FontBlock2Arr(SelInd), xCol, yRow, SelInd, False)) And CheckCharSizeFlag Then
        Call McListBox1.ViewBoldItem(SelInd)    'point to problem char
        McListBox1.Refresh

        Ret = MsgBoxEx("(&h_" & Hex(BaseI + 1) & ") " & ArrMsg(4) & vbCrLf & sColCurrent & "x" & sRowCurrent & " -> " & xCol & "x" & yRow & vbCrLf & vbCrLf & ArrMsg(25), , , CenterOwner, vbOKCancel Or vbQuestion)
        If Ret <> 1 Then Exit Function
    End If
    startAddr = "&H" & FontBlock2Arr(SelInd)
End If

'Call bArr2bTmp(sRowCurrent, sColCurrent)
''''''''''''''''''''''''''''''''''''''''''''WriteBlock''''''''''''''''''''''''''''
Call WriteBlock(startAddr, SelInd)    'save to current block
SaveChars = True    'for reload if saved

If NoBlock2Flag Then Exit Function
'save 2 other block...
If chkDupFont.Value = vbChecked Then    'change all need for other block writing

    If Block1Flag Then

        'check 2 block! here
        If (Not CheckFontSize("&H" & FontBlock2Arr(SelInd), xCol, yRow, SelInd, False)) And CheckCharSizeFlag Then
            Call McListBox1.ViewBoldItem(SelInd)    'point to problem char
            McListBox1.Refresh

            Ret = MsgBoxEx("(&h_" & Hex(BaseI + 1) & ") " & ArrMsg(5) & vbCrLf & sColCurrent & "x" & sRowCurrent & " -> " & xCol & "x" & yRow & vbCrLf & vbCrLf & ArrMsg(25), , , CenterOwner, vbOKCancel Or vbQuestion)
            If Ret <> 1 Then Exit Function
        End If
        startAddr = "&H" & FontBlock2Arr(SelInd)    'for other block
        
    Else 'chek 1 block
        If (Not CheckFontSize("&H" & FontBlock1Arr(SelInd), xCol, yRow, SelInd, True)) And CheckCharSizeFlag Then
            Call McListBox1.ViewBoldItem(SelInd)    'point to problem char
            McListBox1.Refresh

            Ret = MsgBoxEx("(&h_" & Hex(BaseI + 1) & ") " & ArrMsg(5) & vbCrLf & sColCurrent & "x" & sRowCurrent & " -> " & xCol & "x" & yRow & vbCrLf & vbCrLf & ArrMsg(25), , , CenterOwner, vbOKCancel Or vbQuestion)
            If Ret <> 1 Then Exit Function
        End If
        startAddr = "&H" & FontBlock1Arr(SelInd)
    End If


    FontData.reset
    FontData.concat FontDataArr(SelInd)
    'Call ParseInputBin

    If VortexMod Then
        Call FileFontData2bArr_Vortex(SelInd)
    Else
        Call FileFontData2bArr(SelInd)
    End If


    Block1Flag = Not Block1Flag
    Call bArr2bTmp(sRowCurrent, sColCurrent)    'save font to other block too
    bTmpCollection(SelInd) = bTmp
    '''''''''''''''''''''''''''''''''''''''''''WriteBlock'''''''''''''''''''''''''
    Call WriteBlock(startAddr, SelInd)   'save to current block


    'get all changed back
    Block1Flag = Not Block1Flag
    FontData.reset
    FontData.concat FontDataArr(SelInd)
    'Call ParseInputBin

    If VortexMod Then
        Call FileFontData2bArr_Vortex(SelInd)
    Else
        Call FileFontData2bArr(SelInd)
    End If

    Call bArr2bTmp(sRowCurrent, sColCurrent)
    bTmpCollection(SelInd) = bTmp
    'startAddr = "&H" & cmbAdr.List(SelInd)

End If

'''
Exit Function
frmErr:
MsgBox Err.Description & ": SaveChars()"
End Function
Private Sub cmdSave_Click()
'Dim oldSBvalue As Long
Dim arr_Ind() As Integer
Dim i As Long

On Error GoTo frmErr

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If

fFileOpen = OpenFW_write
If Not fFileOpen Then
    fFileOpen = OpenFW_read
    Exit Sub
End If

If SaveChars(cmbAdr.ListIndex, sRow, sCol) Then
    'reload saved items to lists
    'oldSBvalue = McListBox1.SBValue(efsVertical)
    'Call ReloadSavedCurrent(cmbAdr.ListIndex)    'and - bold
    'ChangesIndArr(cmbAdr.ListIndex) = False

    ReDim arr_Ind(0)
    Call GetAllIndexesOfSameGlyph(cmbAdr.ListIndex, arr_Ind)

    For i = 0 To UBound(arr_Ind)
        Call ReloadSavedCurrent(arr_Ind(i))    'and - bold
        ChangesIndArr(arr_Ind(i)) = False
    Next i

    cmdSaveAll.Caption = ArrMsg(12) & " (" & GetChangesCount & ")"

    McListBox1.Refresh
End If

fFileOpen = OpenFW_read

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdSave_Click()"
End Sub


Private Sub cmdSaveAll_Click()

Dim arr_Ind() As Integer
Dim i As Integer, j As Integer

On Error GoTo frmErr

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If


fFileOpen = OpenFW_write
If Not fFileOpen Then
    fFileOpen = OpenFW_read
    Exit Sub
End If
Me.MousePointer = vbHourglass
'DoEvents

PicScroll.Visible = False
noResize = True


For i = 0 To UBound(ChangesIndArr) - 1
    If ChangesIndArr(i) Then    'i -> current selind
        If SaveChars(i, sRowArr(i), sColArr(i)) Then

            ReDim arr_Ind(0)
            Call GetAllIndexesOfSameGlyph(i, arr_Ind)

            '            Call ReloadSavedCurrent(i)
            '            ChangesIndArr(i) = False

            For j = 0 To UBound(arr_Ind)
                Call ReloadSavedCurrent(arr_Ind(j))
                ChangesIndArr(arr_Ind(j)) = False 'dubles false also
            Next j
        End If

    End If
    ' ChangesIndArr(i) = False
    'McListBox1.ListBold(i) = False
Next i

noResize = False
PicScroll.Visible = True

    If VortexMod Then
        Call GetArray_Vortex(cmbLastIndex)    ' recover current unsaved
    Else
        Call GetArray(cmbLastIndex)    ' recover current unsaved
    End If


fFileOpen = OpenFW_read

cmdSaveAll.Caption = ArrMsg(12) & " (" & GetChangesCount & ")"
McListBox1.Refresh


Me.MousePointer = vbNormal
'''
Exit Sub
frmErr:
noResize = False
PicScroll.Visible = True
Me.MousePointer = vbNormal
MsgBox Err.Description & ": cmdSaveAll_Click()"
End Sub

Private Function SaveWord(ByRef Ind As Long, ByRef s() As String) As Boolean
Dim i As Long
Dim n As Integer
'Dim Ind As Long
Dim X() As Byte
'Dim s() As String
'Dim Ret As Long
Dim Tmp As String
Dim sLo As String
Dim sHi As String
Dim NBytes As Integer    '-1

On Error GoTo frmErr

If WordCharNumBytes = 2 Then
    NBytes = WordCharNumBytes * (UBound(s) + 1) - 1
Else
    NBytes = UBound(s)
End If
If NBytes < 0 Then Exit Function

ReDim X(NBytes)

fFileOpen = OpenFW_write
If Not fFileOpen Then Exit Function

Seek #bFileIn, Ind + 1
For i = 0 To NBytes Step WordCharNumBytes

    If WordCharNumBytes = 2 Then
    
'low byte first
        Tmp = right$("000" & s(n), 4)    '####
        sHi = left$(Tmp, 2)
        sLo = right$(Tmp, 2)
        X(i) = (Val("&H" & sLo) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        X(i + 1) = (Val("&H" & sHi) Xor (Ind + 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        
    Else

        X(i) = (Val("&H" & right$("0" & s(i), 2)) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255

    End If
    
    n = n + 1    'for s()
    Ind = Ind + WordCharNumBytes    '1 or 2 bytes
    
Next i

Put #bFileIn, , X()
fFileOpen = OpenFW_read
SaveWord = True

'''
Exit Function
frmErr:
MsgBox Err.Description & ": SaveWord"
End Function

Private Function SaveWord_Vortex(ByRef Ind As Long, ByRef s() As String) As Boolean
Dim i As Long
Dim n As Integer
'Dim Ind As Long
Dim X() As Byte
'Dim s() As String
'Dim Ret As Long
Dim Tmp As String
Dim sLo As String
Dim sHi As String
Dim NBytes As Integer    '-1

On Error GoTo frmErr

'If WordCharNumBytes = 2 Then
'    NBytes = WordCharNumBytes * (UBound(s) + 1) - 1
'Else
    NBytes = UBound(s)
'End If
If NBytes < 0 Then Exit Function

ReDim X(NBytes)

fFileOpen = OpenFW_write
If Not fFileOpen Then Exit Function

Seek #bFileIn, Ind + 1
For i = 0 To NBytes Step WordCharNumBytes

'    If WordCharNumBytes = 2 Then
'
''low byte first
'        Tmp = right$("000" & s(n), 4)    '####
'        sHi = left$(Tmp, 2)
'        sLo = right$(Tmp, 2)
'        X(i) = (Val("&H" & sLo) Xor (Ind + lngBytes + uMagic - lngBytes \ uMagic)) And 255
'        X(i + 1) = (Val("&H" & sHi) Xor (Ind + 1 + lngBytes + uMagic - lngBytes \ uMagic)) And 255
'
'    Else

        X(i) = Val("&H" & right$("0" & s(i), 2))

'    End If
    
    n = n + 1    'for s()
    Ind = Ind + WordCharNumBytes    '1 or 2 bytes
    
Next i

Put #bFileIn, , X()
fFileOpen = OpenFW_read
SaveWord_Vortex = True

'''
Exit Function
frmErr:
MsgBox Err.Description & ": SaveWord_Vortex"
End Function



Private Sub cmdSaveWord_Click()
Dim i As Long
'Dim n As Integer
Dim Ind As Long
'Dim X() As Byte
Dim s() As String
Dim Ret As Long
Dim Tmp As String
'Dim sLo As String
'Dim sHi As String

On Error GoTo frmErr

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If
If CurrentWordInd < 0 Then Exit Sub

If VortexMod And (Not Block1Flag) Then Exit Sub

If Block1Flag Then
    Ind = Word1StartArr(CurrentWordInd)
Else
    Ind = Word2StartArr(CurrentWordInd)
End If

'cmbVocab.text = cmbVocab.text & " 0"
Tmp = cmbVocab.Text

Do While InStr(Tmp, "  ")
    Tmp = Replace(Tmp, "  ", mySpace)
Loop
s = Split(Trim$(Tmp), mySpace)
cmbVocab.Text = Tmp

'check 1.5 bytes max
For i = 0 To UBound(s)
    If Len(s(i)) > 3 Then Exit Sub
Next i

If Block1Flag Then
    If CheckCharSizeFlag And (UBound(s) + 1 <> Word1LenArr(CurrentWordInd)) Then
        Ret = MsgBoxEx(ArrMsg(6) & mySpace & Word1LenArr(CurrentWordInd) & mySpace & ArrMsg(7), , , CenterOwner, vbOKCancel Or vbQuestion)
        If Ret <> 1 Then Exit Sub
    End If

Else
    If CheckCharSizeFlag And (UBound(s) + 1 <> Word2LenArr(CurrentWordInd)) Then
        Ret = MsgBoxEx(ArrMsg(6) & mySpace & Word2LenArr(CurrentWordInd) & mySpace & ArrMsg(7), , , CenterOwner, vbOKCancel Or vbQuestion)
        If Ret <> 1 Then Exit Sub
    End If
End If

If VortexMod Then
    If Not SaveWord_Vortex(Ind, s()) Then Exit Sub
Else
    If Not SaveWord(Ind, s()) Then Exit Sub
End If
'''cut'''

If VortexMod Then
    Call FillVocab_Vortex
Else
    Call FillVocab
End If

cmbVocab.ListIndex = CurrentWordInd

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdSaveWord()"
End Sub

Private Sub cmdScale_Click(Index As Integer)
On Error GoTo frmErr

Select Case Index
Case 0
    If XYspace < 2 Then Exit Sub

    XYspace = XYspace - 1
    If XYspace < 2 Then
        chkGridFlag = vbUnchecked
    End If
Case 1
    If XYspace > 40 Then Exit Sub
    chkGridFlag = chkGrid.Value
    XYspace = XYspace + 1
End Select

TTF_Size = sCol * XYspace
'TTF_Size = TTF_Size + XYspace

With picContainer
    .Cls
    .Height = sRow * XYspace
    .Width = sCol * XYspace
End With
If chkGridFlag = vbChecked Then DrawGrid picContainer, vbGrayText

Call bArr2PicDraw(-1)    '1
If PicTTF.Visible Then
    PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2
End If

'Call Form_Resize in bArr2PicDraw

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdScale_Click()"
End Sub

Private Sub ShiftRotateChar(ByRef Index As Integer, ByRef SelInd As Integer, oneSelFlag As Boolean)
Dim bArrTmp() As Byte
Dim X As Long, Y As Long
Dim i As Long
Dim j As Long
Dim sTmp As Long

On Error GoTo frmErr

Me.MousePointer = vbHourglass
'save old bArr
ReDim bArrTmp(sRow - 1, sCol - 1)
For i = 0 To sRow - 1
    For j = 0 To sCol - 1
        bArrTmp(i, j) = bArr(i, j)
    Next j
Next i

Select Case Index
Case 0    '<

    If oneSelFlag And isSelection Then
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                X = j - 1
                If X < old_X1 Then X = old_X2 - 1
                bArr(i, X) = bArrTmp(i, j)
            Next j
        Next i
    Else
        For i = 0 To sRow - 1    'y
            For j = 0 To sCol - 1    'x
                X = j - 1
                If X < 0 Then X = sCol - 1
                bArr(i, X) = bArrTmp(i, j)
            Next j
        Next i
    End If

Case 1    '>
    If oneSelFlag And isSelection Then
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                X = j + 1
                If X > old_X2 - 1 Then X = old_X1
                bArr(i, X) = bArrTmp(i, j)
            Next j
        Next i
    Else
        For i = 0 To sRow - 1    'y
            For j = 0 To sCol - 1    'x
                X = j + 1
                If X > sCol - 1 Then X = 0
                bArr(i, X) = bArrTmp(i, j)
            Next j
        Next i
    End If

Case 2    ' ^
    If oneSelFlag And isSelection Then
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                X = i - 1
                If X < old_Y1 Then X = old_Y2 - 1
                bArr(X, j) = bArrTmp(i, j)
            Next j
        Next i
    Else
        For i = 0 To sRow - 1    'y
            For j = 0 To sCol - 1    'x
                X = i - 1
                If X < 0 Then X = sRow - 1
                bArr(X, j) = bArrTmp(i, j)
            Next j
        Next i
    End If

Case 3    ' v
    If oneSelFlag And isSelection Then
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                X = i + 1
                If X > old_Y2 - 1 Then X = old_Y1
                bArr(X, j) = bArrTmp(i, j)
            Next j
        Next i

    Else
        For i = 0 To sRow - 1    'y
            For j = 0 To sCol - 1    'x
                X = i + 1
                If X > sRow - 1 Then X = 0
                bArr(X, j) = bArrTmp(i, j)
            Next j
        Next i
    End If

Case 4    ' UCW
    If oneSelFlag And isSelection Then
        If old_X2 - old_X1 = old_Y2 - old_Y1 Then

            Y = old_Y2 - 1
            X = old_X1
            For i = old_Y1 To old_Y2 - 1
                For j = old_X1 To old_X2 - 1
                    bArr(Y, X) = bArrTmp(i, j)
                    Y = Y - 1
                Next j
                X = X + 1: Y = old_Y2 - 1
            Next i
        End If
    Else
        ReDim bArr(sCol - 1, sRow - 1)
        Y = sCol - 1    'new row
        For i = 0 To sRow - 1
            For j = 0 To sCol - 1
                bArr(Y, X) = bArrTmp(i, j)
                Y = Y - 1
            Next j
            X = X + 1: Y = sCol - 1
        Next i
        sTmp = sRow: sRow = sCol: sCol = sTmp
    End If

Case 5    ' CW
    If oneSelFlag And isSelection Then
        If old_X2 - old_X1 = old_Y2 - old_Y1 Then

            Y = old_Y1
            X = old_X2 - 1    'new col

            For j = old_X1 To old_X2 - 1
                For i = old_Y1 To old_Y2 - 1
                    bArr(Y, X) = bArrTmp(i, j)
                    X = X - 1
                Next i
                X = old_X2 - 1: Y = Y + 1
            Next j
        End If

    Else
        ReDim bArr(sCol - 1, sRow - 1)
        X = sRow - 1    'new col
        For j = 0 To sCol - 1
            For i = 0 To sRow - 1
                bArr(Y, X) = bArrTmp(i, j)
                X = X - 1
            Next i
            X = sRow - 1: Y = Y + 1
        Next j
        sTmp = sRow: sRow = sCol: sCol = sTmp
    End If

Case 6
    'inv
    If oneSelFlag And isSelection Then
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                If bArr(i, j) = 1 Then
                    bArr(i, j) = 0
                Else
                    bArr(i, j) = 1
                End If

            Next j
        Next i

    Else
        For i = 0 To sRow - 1
            For j = 0 To sCol - 1
                If bArr(i, j) = 1 Then
                    bArr(i, j) = 0
                Else
                    bArr(i, j) = 1
                End If

            Next j
        Next i
    End If

Case 7
    'cls
    If oneSelFlag And isSelection Then
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                bArr(i, j) = 0
            Next j
        Next i
    Else
        For i = 0 To sRow - 1
            For j = 0 To sCol - 1
                bArr(i, j) = 0
            Next j
        Next i
    End If

Case 8    'flip <>
    If oneSelFlag And isSelection Then
        X = old_X2 - 1
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                bArr(i, X) = bArrTmp(i, j)
                X = X - 1
            Next j
            X = old_X2 - 1
        Next i

    Else
        ReDim bArr(sRow - 1, sCol - 1)
        X = sCol - 1
        For i = 0 To sRow - 1
            For j = 0 To sCol - 1
                bArr(i, X) = bArrTmp(i, j)
                X = X - 1
            Next j
            X = sCol - 1
        Next i
    End If

Case 9    'flip ^v
    If oneSelFlag And isSelection Then
        Y = old_Y2 - 1
        For i = old_Y1 To old_Y2 - 1
            For j = old_X1 To old_X2 - 1
                bArr(Y, j) = bArrTmp(i, j)
            Next j
            Y = Y - 1
        Next i

    Else
        ReDim bArr(sRow - 1, sCol - 1)
        Y = sRow - 1
        For i = 0 To sRow - 1
            For j = 0 To sCol - 1
                bArr(Y, j) = bArrTmp(i, j)
            Next j
            Y = Y - 1
        Next i
    End If

End Select

Call XYcaptionSet(sCol, sRow)

If Not oneSelFlag Then
                sColArr(SelInd) = sCol
                sRowArr(SelInd) = sRow
End If
       
Call bArr2PicDraw(SelInd)    '(-1)    '0
'Call PicDraw2PicReal(-1)    '1
Call bArr2PicReal(SelInd)    '(-1)

If oneSelFlag Then


    Call Copy2PicRectSel    '2   'if 1 selected
    Call StoreInUndoBuffer
    If PicTTF.Visible Then
        PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2
    End If
End If

Call StoreCurrentChar(SelInd)    '(cmbAdr.ListIndex)

Me.MousePointer = vbNormal

'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
ReDim bArr(sRow - 1, sCol - 1)
For i = 0 To sRow - 1
    For j = 0 To sCol - 1
        bArr(i, j) = bArrTmp(i, j)
    Next j
Next i
MsgBox Err.Description & ": ShiftRotateChar()"
End Sub

Private Sub cmdToolBar_click(Index As Integer)
Dim i As Long
Dim n As Integer

On Error GoTo frmErr

Select Case Index
Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9

    If fFileOpen Then
        If UBound(selArr) = 1 Then
            Call ShiftRotateChar(Index, -1, True)
        Else
            For n = 1 To UBound(selArr)

                'getdata
                sCol = sColArr(selArr(n))
                sRow = sRowArr(selArr(n))
                FontData.reset
                FontData.concat FontDataArr(selArr(n))

                If VortexMod Then
                    Call FileFontData2bArr_Vortex(selArr(n))
                Else
                    Call FileFontData2bArr(selArr(n))
                End If

                'go shift
                Call ShiftRotateChar(Index, selArr(n), False)

            Next n
            'get current back
            If cmbAdr.ListIndex > -1 Then

                If VortexMod Then
                    Call GetArray_Vortex(cmbAdr.ListIndex)
                Else
                    Call GetArray(cmbAdr.ListIndex)
                End If

            End If
            
        End If

    Else

        Call ShiftRotateChar(Index, -1, True)
    End If
    Exit Sub

Case 10    'ttf
    PicTTF.Visible = Not PicTTF.Visible

    If PicTTF.Visible Then
        PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2


        'fill system fonts
        cmb_SysFonts.Visible = False
        Call FillListWithFonts(cmb_SysFonts)
        cmb_SysFonts.Visible = True
        If Len(TTFontName) <> 0 Then    'this init scrrolls
            cmb_SysFonts.Text = TTFontName
        Else
            'cmb_SysFonts.ListIndex = 0
            cmb_SysFonts.Text = "Arial"
        End If

        'TTF_Size = sCol * XYspace
        'If TTF_Size < VScroll_S.Min Then VScroll_S.Value = TTF_Size       '250
        'TTF_Size = VScroll_S.Value

        On Error Resume Next
        If sCol > sRow Then
            VScroll_S.Value = sRow * XYspace
        Else
            VScroll_S.Value = sCol * XYspace
        End If
        On Error GoTo frmErr

        VScroll_Y.Value = -20
        'VScroll_Y.Value = (VScroll_Y.Max - Abs(VScroll_Y.Min)) / 2
        'HScroll_X.Value = (HScroll_X.Max - Abs(HScroll_X.Min)) / 2 '=0

        cmbTTF_Char.Visible = False
        cmbTTF_Char.Clear
        For i = 33 To 255
            cmbTTF_Char.AddItem Chr(i)
        Next i
        cmbTTF_Char.Visible = True
        LoadFontFlag = True
        cmbTTF_Char.Text = 5
        LoadFontFlag = False
        'Call Form_Resize

    End If
    Exit Sub

Case 11    'bmp
    Call LoadBMP
    Exit Sub

Case 12    'import from txt file
    Call LoadImport
    Exit Sub

Case 13    'export to txt file
    Call SaveExport(False)
    Exit Sub
End Select

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": cmdToolBar_Click()"
End Sub

Private Sub SaveExport(ByRef FromArrExport As Boolean, Optional fHex As Boolean = False)
Dim fn As String
'Dim FileTitle As String
Dim ExportDir As String
'Dim sClipFont As String
Dim f As Integer
Dim strData As String

On Error GoTo frmErr

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If

Call CopyData(strData, FromArrExport, fHex)

'If Len(strData) <> 0 Then 'always present
f = FreeFile
ExportDir = App.Path & "\Export"
fn = vbNullString
fn = ExportSaveDialog(ExportDir, ArrMsg(18))
If fn = vbNullString Then Exit Sub
'LastPath = GetPathFromPathAndName(FileName)

Open fn For Output Access Write As #f
Print #f, strData
Close #f
'End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": SaveExport"
End Sub

Private Sub ParseXML()
Dim i As Integer, m As Integer  'j As Integer
Dim Tmp As String
Dim strData() As String
'Dim numData() As Integer
Dim sFontData() As String
Dim arrClipFont() As String
Dim sClipFont As String
Dim sZero As String

On Error GoTo frmErr

Set m_oCurrentElement = m_oDoc.Root

ReDim arrXML_ImageNum(0)
ReDim arrXML_ImageRow(0)
ReDim arrXML_ImageCol(0)
ReDim arrXML_DataBody(0)
XML_Image_Count = -1

Call XML_Fill_Arrays(m_oCurrentElement)

ReDim sFontData(XML_Image_Count)
ReDim arrClipFont(XML_Image_Count)

For i = 0 To XML_Image_Count
    Tmp = Replace(arrXML_DataBody(i), vbCr, vbNullString)
    Tmp = Replace(Tmp, ".", 0)
    Tmp = Replace(Tmp, "X", 1)
    strData() = Split(Tmp, vbLf)


    For m = 1 To arrXML_ImageRow(i)
        sZero = Space(intBytes(arrXML_ImageCol(i)) - arrXML_ImageCol(i))
        sZero = Replace(sZero, mySpace, 0)
        strData(m) = Trim(strData(m))
        strData(m) = left$(strData(m), arrXML_ImageCol(i)) & sZero
    Next m
'
'    For j = 1 To arrXML_ImageRow(i)
    sFontData(0) = Join(strData(), vbNullString)
    arrClipFont(i) = arrXML_ImageNum(i) & "," & arrXML_ImageCol(i) & "," & arrXML_ImageRow(i) & "," & sFontData(0)
'    Next j

Next i

sClipFont = Join(arrClipFont(), vbCr)

'FontData.Reset
'For i = 3 To UBound(sClipBytes)
'    FontData.concat dec2bin(Val(sClipBytes(i)))    'this FontData - simple, without formatting for block and headers
'Next i
'Call PasteFontData2bArr

'i=223: ?arrXML_ImageNum(i),arrXML_Imagecol(i),arrXML_Imagerow(i)
'?arrXML_DataBody(223)
Set m_oDoc = Nothing

Call PasteData(sClipFont, True)
Call StoreInUndoBuffer

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": ParseXML"
End Sub

Private Sub LoadImport()
Dim fn As String
Dim FileTitle As String
Dim ExportDir As String
Dim sClipFont As String    '1,1,1...[vbcr]2,2,2...[vbcr]
Dim f As Integer
Dim Tmp As String
Dim sFileData As String    'xml

On Error GoTo frmErr

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If

f = FreeFile
ExportDir = App.Path & "\Export"

fn = vbNullString
fn = ImportLoadDialog(ExportDir, ArrMsg(17), FileTitle)
If fn = vbNullString Then Exit Sub
'LastPath = GetPathFromPathAndName(FileName)

DoEvents    '2 close dialog
Me.MousePointer = vbHourglass

Tmp = GetNameExt(fn)    'can select preview bmp, but load txt

If InStr(1, Tmp, ".respack") Then

    Open fn For Input Access Read As #f
    sFileData = Input$(LOF(f), f)    'all file into string
    m_oDoc.LoadData (sFileData)
    Close #f
    If m_oDoc Is Nothing Then Exit Sub
    Call ParseXML
    Me.MousePointer = vbNormal
    Exit Sub

ElseIf InStr(1, Tmp, ".bmp") Then
    Tmp = GetName(Tmp)
    Tmp = Tmp & ".txt"
    Call searchForFile(ExportDir, Tmp)
    If Len(FindFilePath) = 0 Then
        Exit Sub
    Else
        fn = FindFilePath
        FindFilePath = vbNullString
    End If
'fn = ExportDir & "\" & tmp & ".txt"
'    If Not FileExists(fn) Then Exit Sub

End If

Open fn For Input Access Read As #f
sClipFont = Input$(LOF(f), f)    'all file into string
Close #f

Call PasteData(sClipFont)
Call StoreInUndoBuffer

Me.MousePointer = vbNormal
'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": LoadImport"
End Sub
Private Sub AppendImport()
Dim fn As String

Dim ExportDir As String
Dim sClipFont As String    '1,1,1...[vbcr]2,2,2...[vbcr]
Dim f As Integer
Dim sClipChars() As String
Dim Ret As Long
Dim n As Integer
Dim arrFromFile() As String
Dim nCharNumber As Integer
Dim s() As String
Dim sClipBytes() As String
Dim cClipFont As New CString
Dim strData As String

On Error GoTo frmErr
f = FreeFile
ExportDir = App.Path & "\Export"
fn = vbNullString
fn = ExportSaveDialog(ExportDir, ArrMsg(18))    'save dialog but first read exists data
If fn = vbNullString Then Exit Sub

'LastPath = GetPathFromPathAndName(FileName)

DoEvents    '2 close dialog


Open fn For Input Access Read As #f
sClipFont = Input$(LOF(f), f)    'all file into string

sClipFont = Replace(sClipFont, vbLf, vbNullString)
'check valid
s = Split(sClipFont, ",")
If UBound(s) = 0 Then Exit Sub
If Not IsNumeric(s(0)) Then Exit Sub

sClipChars = Split(sClipFont, vbCr)

Ret = UBound(sClipChars)
ReDim arrFromFile(0)
Me.MousePointer = vbHourglass

If Ret > -1 Then
    For n = 0 To Ret
'save old from file to array (start with 1) arrFromFile(num)=all_line
        sClipBytes = Split(sClipChars(n), ",", 2)
        If UBound(sClipBytes) = 1 Then

            If IsNumeric(sClipBytes(0)) Then
                nCharNumber = Val(sClipBytes(0))
                If nCharNumber >= 0 Then
                    If UBound(arrFromFile) < nCharNumber Then ReDim Preserve arrFromFile(nCharNumber)
                    arrFromFile(nCharNumber) = sClipChars(n)
                End If
            End If
        End If

    Next n

End If
Close #f

'   If FromArrExport Then 'from menu
For n = 0 To UBound(arrExport)    '0-226
    If arrExport(n) Then
        If UBound(arrFromFile) < n + 1 Then ReDim Preserve arrFromFile(n + 1)
        arrFromFile(n + 1) = n + 1 & "," & CharDataArr(n)
    End If
Next n

For n = 1 To UBound(arrFromFile)
    If Len(arrFromFile(n)) <> 0 Then cClipFont.concat arrFromFile(n) & vbCrLf
Next n

strData = left$(cClipFont.Text, Len(cClipFont.Text) - 2)  '- last vbCrLf
'Set cClipFont = Nothing

Open fn For Output Access Write As #f
Print #f, strData
Close #f


Me.MousePointer = vbNormal
'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": AppendImport"
End Sub
'Convert an image to a specific number of shades of gray (WITH error-diffusion dithering support!)
Public Sub drawGrayscaleCustomShadesDithered(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal numOfShades As Long = 256)
On Error GoTo frmErr

'These arrays will hold the source and destination image's pixel data, respectively
Dim ImageData() As Byte

'Coordinate variables
Dim X As Long, Y As Long

'Image dimensions
Dim iWidth As Long, iHeight As Long

'Instantiate a FastDrawing class and gather the image's data (into ImageData())
Dim fDraw As New FastDrawing
iWidth = fDraw.GetImageWidth(srcPic)
iHeight = fDraw.GetImageHeight(srcPic)
fDraw.GetImageData2D srcPic, ImageData()

'These variables will hold temporary pixel color values (red, green, blue)
Dim r As Byte, g As Byte, b As Byte

'This value will hold the grayscale value of each pixel
Dim Gray As Byte

'This look-up table holds all possible totals of adding together the R, G, and B values of an image (0 to 255*3 - for pure white)
Dim grayLookup(0 To 765) As Byte

'Populate the look-up table
For X = 0 To 765
    grayLookup(X) = X \ 3
Next X

'This conversionFactor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
Dim conversionFactor As Long    'Single
conversionFactor = 255    '(255 / (numOfShades - 1))

'Unfortunately, this algorithm (unlike its non-dithering counterpart) is not well-suited to using a look-up table, so all calculations have been moved into the loop
Dim grayTempCalc As Long

'This value tracks the drifting error of our conversions, which allows us to dither
Dim errorValue As Long
errorValue = 0

'Loop through the image, adjusting pixel values as we go
Dim quickX As Long

'Note that I have reversed the loop order (now we go horizontally instead of vertically).
' This is because I want my dithering algorithm to work from left-to-right instead of top-to-bottom.
'    For y = 0 To iHeight - 1
'    For x = 0 To iWidth - 1

For X = 0 To iWidth - 1
    For Y = 0 To iHeight - 1


        quickX = X * 3

'Get the source image pixels
        r = ImageData(quickX + 2, Y)
        g = ImageData(quickX + 1, Y)
        b = ImageData(quickX, Y)

'First, generate a raw grayscale value
        Gray = grayLookup(CLng(r) + CLng(g) + CLng(b))
        grayTempCalc = Gray

'Add the error value (a cumulative value of the difference between actual gray values and gray values we've selected) to the current gray value
        grayTempCalc = grayTempCalc + errorValue

'Rebuild our temporary calculation variable using the shade reduction formula
        grayTempCalc = Int((CDbl(grayTempCalc) / conversionFactor) + 0.5) * conversionFactor

'Adjust our error value to include this latest calculation
        errorValue = CLng(Gray) + errorValue - grayTempCalc

        Gray = ByteMeL(grayTempCalc)

'Assign all color channels to the new gray value
        ImageData(quickX + 2, Y) = Gray
        ImageData(quickX + 1, Y) = Gray
        ImageData(quickX, Y) = Gray

    Next Y
'Reset our error value after each row
    errorValue = 0
Next X

'Draw the new image data to the screen
fDraw.SetImageData2D dstPic, iWidth, iHeight, ImageData()

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": ShadesDithered"
End Sub
Public Sub DrawGrayscaleAtkinsonGS(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal numOfShades As Long = 256)
'Atkinson gs

'Instantiate a FastDrawing class and gather the image's data (into ImageData())
Dim fDraw As New FastDrawing
Dim ImageData() As Byte
Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
initX = 0
initY = 0
finalX = fDraw.GetImageWidth(srcPic)
finalY = fDraw.GetImageHeight(srcPic)
fDraw.GetImageData2D srcPic, ImageData()

'These values will help us access locations in the array more quickly.
' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
Dim QuickVal As Long, qvDepth As Long
qvDepth = 3    'srcDIB.GetDIBColorDepth \ 8

'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
' based on the size of the area to be processed.
Dim progBarCheck As Long

'Color variables
Dim g As Long, grayVal As Long

'This conversion factor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
Dim conversionFactor As Double
conversionFactor = 255    '(255 / (numOfShades - 1))

'Build a look-up table for our custom grayscale conversion results
Dim LookUp(0 To 255) As Long

For X = 0 To 255
    grayVal = Int((CDbl(X) / conversionFactor) + 0.5) * conversionFactor
    If grayVal > 255 Then grayVal = 255
    LookUp(X) = grayVal
Next X

Dim DitherTable() As Byte
Dim xLeft As Long, xRight As Long, yDown As Long
Dim errorVal As Double
Dim dDivisor As Double
Dim l As Long, newL As Long

'First, prepare a dither table
ReDim DitherTable(-1 To 2, 0 To 2) As Byte
DitherTable(1, 0) = 1
DitherTable(2, 0) = 1
DitherTable(-1, 1) = 1
DitherTable(0, 1) = 1
DitherTable(1, 1) = 1
DitherTable(0, 2) = 1
dDivisor = 8
'Second, mark the size of the array in the left, right, and down directions
xLeft = -1
xRight = 2
yDown = 2

'First, we need a dithering table the same size as the image.  We make it of Single type to prevent rounding errors.
' (This uses a lot of memory, but on modern systems it shouldn't be a problem.)
Dim dErrors() As Single
ReDim dErrors(0 To finalX, 0 To finalY) As Single

Dim i As Long, j As Long
Dim quickX As Long, QuickY As Long

'Now loop through the image, calculating errors as we go
For Y = initY To finalY - 1
    For X = initX To finalX - 1

        QuickVal = X * qvDepth

'Get the source pixel color values.  Because we know the image we're handed is already going to be grayscale,
' we can shortcut this calculation by only grabbing the red channel.
        g = ImageData(QuickVal + 2, Y)

'Convert those to a luminance value and add the value of the error at this location
        l = g + dErrors(X, Y)

'Convert that to a lookup-table-safe luminance (e.g. 0-255)
        If l < 0 Then
            newL = 0
        ElseIf l > 255 Then
            newL = 255
        Else
            newL = l
        End If

'Write the new luminance value out to the image array
        ImageData(QuickVal + 2, Y) = LookUp(newL)
        ImageData(QuickVal + 1, Y) = LookUp(newL)
        ImageData(QuickVal, Y) = LookUp(newL)

'Calculate an error for this calculation
        errorVal = l - LookUp(newL)

'If there is an error, spread it
        If errorVal <> 0 Then

'Now, spread that error across the relevant pixels according to the dither table formula
            For i = xLeft To xRight
                For j = 0 To yDown

'First, ignore already processed pixels
                    If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel

'Second, ignore pixels that have a zero in the dither table
                    If DitherTable(i, j) = 0 Then GoTo NextDitheredPixel

                    quickX = X + i
                    QuickY = Y + j

'Next, ignore target pixels that are off the image boundary
                    If quickX < initX Then GoTo NextDitheredPixel
                    If quickX > finalX Then GoTo NextDitheredPixel
                    If QuickY > finalY Then GoTo NextDitheredPixel

'If we've made it all the way here, we are able to actually spread the error to this location
                    dErrors(quickX, QuickY) = dErrors(quickX, QuickY) + (errorVal * (CSng(DitherTable(i, j)) / dDivisor))

NextDitheredPixel:                     Next j
            Next i
        End If
    Next X
Next Y

'    'With our work complete, point ImageData() away from the DIB and deallocate it
'    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
'    Erase ImageData

'Draw the new image data to the screen
fDraw.SetImageData2D dstPic, finalX, finalY, ImageData()

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": AtkinsonGS"
End Sub

'Convert an image to a specific number of shades of gray; any value in the range [2,256] is acceptable
Public Sub DrawGrayscaleAtkinsonBW(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal numOfShades As Long = 256)
'Atkinson test bw
On Error GoTo 0    'frmErr

Dim DitherTable() As Byte
Dim xLeft As Long, xRight As Long, yDown As Long
Dim errorVal As Double
Dim dDivisor As Double
Dim lowR As Long, lowG As Long, lowB As Long
Dim highR As Long, highG As Long, highB As Long

'Image dimensions
Dim iWidth As Long, iHeight As Long
Dim X As Long, Y As Long, i As Long, j As Long
Dim quickX As Long, QuickY As Long
Dim dErrors() As Double

Dim fDraw As New FastDrawing
Dim QuickVal As Long
Dim ImageData() As Byte
Dim r As Byte, g As Byte, b As Byte
Dim l As Long, newL As Long

lowR = ExtractR(&H0)
lowG = ExtractG(&H0)
lowB = ExtractB(&H0)

highR = ExtractR(&HFFFFFF)
highG = ExtractG(&HFFFFFF)
highB = ExtractB(&HFFFFFF)

'First, prepare a dither table
ReDim DitherTable(-1 To 2, 0 To 2) As Byte

DitherTable(1, 0) = 1
DitherTable(2, 0) = 1

DitherTable(-1, 1) = 1
DitherTable(0, 1) = 1
DitherTable(1, 1) = 1

DitherTable(0, 2) = 1

dDivisor = 8

'Second, mark the size of the array in the left, right, and down directions
xLeft = -1
xRight = 2
yDown = 2

'First, we need a dithering table the same size as the image.  We make it of Single type to prevent rounding errors.
' (This uses a lot of memory, but on modern systems it shouldn't be a problem.)


iWidth = fDraw.GetImageWidth(srcPic)
iHeight = fDraw.GetImageHeight(srcPic)
fDraw.GetImageData2D srcPic, ImageData()

ReDim dErrors(0 To iWidth, 0 To iHeight) As Double


'Now loop through the image, calculating errors as we go
For X = 0 To iWidth - 1
    QuickVal = X * 3    'qvDepth 24 bit
    For Y = 0 To iHeight - 1

'Get the source pixel color values
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)

'Convert those to a luminance value and add the value of the error at this location
        l = getLuminance(r, g, b)

        newL = l + dErrors(X, Y)

'Check our modified luminance value against the threshold, and set new values accordingly
        If newL >= 128 Then     'cThreshold?
            errorVal = newL - 255
            ImageData(QuickVal + 2, Y) = highR
            ImageData(QuickVal + 1, Y) = highG
            ImageData(QuickVal, Y) = highB
        Else
            errorVal = newL
            ImageData(QuickVal + 2, Y) = lowR
            ImageData(QuickVal + 1, Y) = lowG
            ImageData(QuickVal, Y) = lowB
        End If

'If there is an error, spread it
        If errorVal <> 0 Then

'Now, spread that error across the relevant pixels according to the dither table formula
            For i = xLeft To xRight
                For j = 0 To yDown

'First, ignore already processed pixels
                    If (j = 0) And (i <= 0) Then GoTo NextDitheredPixel

'Second, ignore pixels that have a zero in the dither table
                    If DitherTable(i, j) = 0 Then GoTo NextDitheredPixel

                    quickX = X + i
                    QuickY = Y + j

'Next, ignore target pixels that are off the image boundary
                    If quickX < 0 Then GoTo NextDitheredPixel
                    If quickX > iWidth - 1 Then GoTo NextDitheredPixel
                    If QuickY > iHeight - 1 Then GoTo NextDitheredPixel

'If we've made it all the way here, we are able to actually spread the error to this location
                    dErrors(quickX, QuickY) = dErrors(quickX, QuickY) + (errorVal * (CSng(DitherTable(i, j)) / dDivisor))

NextDitheredPixel:                     Next j
            Next i
        End If
    Next Y
Next X
'With our work complete, point ImageData() away from the DIB and deallocate it
'   CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
'   Erase ImageData

'Draw the new image data to the screen
fDraw.SetImageData2D dstPic, iWidth, iHeight, ImageData()

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": AtkinsonBW"
End Sub

'Convert an image to a specific number of shades of gray; any value in the range [2,256] is acceptable
Public Sub drawGrayscaleCustomShades(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal numOfShades As Long = 256)
On Error GoTo frmErr

'These arrays will hold the source and destination image's pixel data, respectively
Dim ImageData() As Byte

'Coordinate variables
Dim X As Long, Y As Long

'Image dimensions
Dim iWidth As Long, iHeight As Long

'Instantiate a FastDrawing class and gather the image's data (into ImageData())
Dim fDraw As New FastDrawing
iWidth = fDraw.GetImageWidth(srcPic)
iHeight = fDraw.GetImageHeight(srcPic)
fDraw.GetImageData2D srcPic, ImageData()

'These variables will hold temporary pixel color values (red, green, blue)
Dim r As Byte, g As Byte, b As Byte

'This value will hold the grayscale value of each pixel
Dim Gray As Byte

'This conversionFactor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
Dim conversionFactor As Single
conversionFactor = 255    '(255 / (numOfShades - 1))

'This algorithm is well-suited to using a look-up table, so let's build one and (obviously!) prepopulate it
Dim grayLookup(0 To 255) As Byte
Dim grayTempCalc As Long

For X = 0 To 255
    grayTempCalc = Int((CDbl(X) / conversionFactor) + 0.5) * conversionFactor
    grayLookup(X) = ByteMeL(grayTempCalc)
Next X

'Loop through the image, adjusting pixel values as we go
Dim quickX As Long

For X = 0 To iWidth - 1
    quickX = X * 3
    For Y = 0 To iHeight - 1

'Get the source image pixels
        r = ImageData(quickX + 2, Y)
        g = ImageData(quickX + 1, Y)
        b = ImageData(quickX, Y)

'Look up this pixel's value in the lookup table
        Gray = grayLookup((CLng(r) + CLng(g) + CLng(b)) \ 3)

'Assign all color channels to the new gray value
        ImageData(quickX + 2, Y) = Gray
        ImageData(quickX + 1, Y) = Gray
        ImageData(quickX, Y) = Gray

    Next Y
Next X

'Draw the new image data to the screen
fDraw.SetImageData2D dstPic, iWidth, iHeight, ImageData()

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": CustomShades"
End Sub
'This function ensures that a long-type variable falls into the range of 0-255
Public Function ByteMeL(ByRef TempVar As Long) As Byte
If TempVar > 255 Then
    ByteMeL = 255
ElseIf TempVar < 0 Then
    ByteMeL = 0
Else
    ByteMeL = CByte(TempVar)
End If
End Function
Private Sub LoadClipboardBMP()
'Dim oldXYspace As Integer
Dim sColIn As Long, sRowIn As Long
Dim SelInd As Integer
Dim Ret As Long
Dim arr_Ind() As Integer
Dim i As Integer

On Error GoTo frmErr
'oldXYspace = XYspace
'XYspace = 1
Set tmpPic = Nothing    'or overlay in png transparent (cool)
tmpPic.Picture = Clipboard.GetData

If fFileOpen Then
    If tmpPic.Height > 255 Or tmpPic.Width > 255 Then
        MsgBoxEx ArrMsg(22), , , CenterOwner, vbOKOnly Or vbExclamation
        Exit Sub
    End If
End If

sRowIn = tmpPic.Height
sColIn = tmpPic.Width
SelInd = cmbAdr.ListIndex
ReDim arr_Ind(0)
Call GetAllIndexesOfSameGlyph(SelInd, arr_Ind)

If sCol <> sColIn Or sRow <> sRowIn Then
If CheckCharSizeFlag Then
    'Me.MousePointer = vbNormal
    Ret = MsgBoxEx(ArrMsg(3) & vbCrLf & sColIn & "x" & sRowIn & " -> " & sCol & "x" & sRow & vbCrLf & vbCrLf & ArrMsg(24), , , CenterOwner, vbYesNoCancel Or vbQuestion)

    Select Case Ret
    Case 6  'yes change size
        If fFileOpen Then
            For i = 0 To UBound(arr_Ind)
                'sRowArr(SelInd) = sRowIn
                'sColArr(SelInd) = sColIn
                sRowArr(arr_Ind(i)) = sRowIn
                sColArr(arr_Ind(i)) = sColIn
            Next i
        End If
        sCol = sColIn
        sRow = sRowIn
    Case 7    'no only paste
    Case Else
        Exit Sub
    End Select

Else    ' yes change size
    If fFileOpen Then
        For i = 0 To UBound(arr_Ind)
            'sRowArr(SelInd) = sRowIn
            'sColArr(SelInd) = sColIn
            sRowArr(arr_Ind(i)) = sRowIn
            sColArr(arr_Ind(i)) = sColIn
        Next i
    End If
    sCol = sColIn
    sRow = sRowIn

End If
End If

Me.MousePointer = vbHourglass

picContainer.Visible = False
picContainer.Cls
picContainer.Height = tmpPic.Height
picContainer.Width = tmpPic.Width

Select Case fPicDithered
Case 0
    drawGrayscaleCustomShades tmpPic, picContainer, 2
Case 1
    DrawGrayscaleAtkinsonGS tmpPic, picContainer, 2
Case 2
    DrawGrayscaleAtkinsonBW tmpPic, picContainer, 2
Case 3
    drawGrayscaleCustomShadesDithered tmpPic, picContainer, 2
End Select

'picContainer.AutoSize = True
'picContainer.Picture = LoadPicture(FileName)
'picContainer.AutoSize = False

picContainer.Picture = picContainer.Image
'picContainer.Visible = True

'sRow = picContainer.Height
'sCol = picContainer.Width

Me.MousePointer = vbHourglass
ReDim bArr(sRow - 1, sCol - 1)

'chkGridFlag = vbUnchecked
Call Draw2bArr(picContainer, sCol, sRow)  ' write to array bArr
Call StoreCurrentChar(cmbAdr.ListIndex) '-1 when no fileopen

Call XYcaptionSet(sCol, sRow)
'XYspace = oldXYspace
Call bArr2PicDraw(-1)
'Call PicDraw2PicReal(-1)
Call bArr2PicReal(-1)
' in call Call StoreInUndoBuffer

'chkGridFlag = chkGrid.Value
If chkGridFlag Then DrawGrid picContainer, vbGrayText
'If chkGrid.Value = vbChecked Then DrawGrid picContainer, vbGrayText

picContainer.Visible = True
Me.MousePointer = vbNormal

'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": LoadClipboardBMP()"
End Sub
Private Sub LoadBMP()
Dim filename As String
Dim FileTitle As String
On Error GoTo frmErr

filename = vbNullString
filename = BMPLoadDialog(ArrMsg(14), FileTitle)
If filename = vbNullString Then Exit Sub
LastPath = GetPathFromPathAndName(filename)

BMPFileName = FileTitle
'XYspace = 1
Set tmpPic = Nothing    'or overlay in png transparent (cool)
If LCase(right$(filename, 4)) = ".png" Then
    pngClass.PicBox = tmpPic    'form or Picturebox
'pngClass.SetToBkgrnd True, 100, 50 'set to Background (True or false), x and y
    pngClass.BackgroundPicture = tmpPic    'same Backgroundpicture
    pngClass.SetAlpha = True    'when Alpha then alpha
    pngClass.SetTrans = True    'when transparent Color then transparent Color
    pngClass.OpenPNG filename    'Open and display Picture
Else
    tmpPic.Picture = LoadPicture(filename)
End If

If fFileOpen Then
    If tmpPic.Height > 255 Or tmpPic.Width > 255 Then
        MsgBoxEx ArrMsg(22), , , CenterOwner, vbOKOnly Or vbExclamation
        Exit Sub
    End If
End If

Me.MousePointer = vbHourglass

picContainer.Visible = False
picContainer.Cls
picContainer.Height = tmpPic.Height
picContainer.Width = tmpPic.Width

Select Case fPicDithered
Case 0
    drawGrayscaleCustomShades tmpPic, picContainer, 2
Case 1
    DrawGrayscaleAtkinsonGS tmpPic, picContainer, 2
Case 2
    DrawGrayscaleAtkinsonBW tmpPic, picContainer, 2
Case 3
    drawGrayscaleCustomShadesDithered tmpPic, picContainer, 2
End Select

'picContainer.AutoSize = True
'picContainer.Picture = LoadPicture(FileName)
'picContainer.AutoSize = False

picContainer.Picture = picContainer.Image
picContainer.Visible = True

sRow = picContainer.Height
sCol = picContainer.Width

ReDim bArr(sRow - 1, sCol - 1)

'chkGridFlag = vbUnchecked

Call Draw2bArr(picContainer, sCol, sRow)  ' write to array bArr
Call StoreCurrentChar(cmbAdr.ListIndex)

Call XYcaptionSet(sCol, sRow)

Call bArr2PicDraw(-1)
'Call PicDraw2PicReal(-1)
Call bArr2PicReal(-1)
Call StoreInUndoBuffer

If chkGrid.Value = vbChecked Then DrawGrid picContainer, vbGrayText

'LockWindowUpdate 0
Me.MousePointer = vbNormal

'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": LoadBMP()"
End Sub

Private Sub StoreCurrentChar(ByRef SelInd As Integer)
'same in storeAfterPaste
'Dim SelInd As Integer
Dim arr_Ind() As Integer
Dim i As Integer

On Error GoTo frmErr
'Debug.Print ">   StoreCurrentChar"
If fFileOpen Then

    ' SelInd = cmbAdr.ListIndex
    If SelInd = -1 Then SelInd = cmbAdr.ListIndex

    ReDim arr_Ind(0)
    Call GetAllIndexesOfSameGlyph(SelInd, arr_Ind)

    For i = 0 To UBound(arr_Ind)
        'sRowArr(SelInd) = sRow
        'sColArr(SelInd) = sCol
        sRowArr(arr_Ind(i)) = sRow
        sColArr(arr_Ind(i)) = sCol
    Next i


    If PicNotEqual(SelInd) Then
        'ChangesIndArr(SelInd) = True    'for GetChangesCount

        For i = 0 To UBound(arr_Ind)
            McListBox1.ListBold(arr_Ind(i)) = True
            ChangesIndArr(arr_Ind(i)) = True
        Next i
        'McListBox1.ListBold(SelInd) = True
        'McListBox1.ListBold(280) = True

    Else
        'ChangesIndArr(SelInd) = False
        For i = 0 To UBound(arr_Ind)
            McListBox1.ListBold(arr_Ind(i)) = False
            ChangesIndArr(arr_Ind(i)) = False
        Next i
        'McListBox1.ListBold(SelInd) = False
        'McListBox1.ListBold(280) = False

    End If

    cmdSaveAll.Caption = ArrMsg(12) & " (" & GetChangesCount & ")"

    Call bArr2bTmp(sRowArr(SelInd), sColArr(SelInd))

    For i = 0 To UBound(arr_Ind)
        '    bTmpCollection(SelInd) = bTmp
        '    Call bTmp2FontData(SelInd)
        '    Call bTmp2CharData(SelInd)
        bTmpCollection(arr_Ind(i)) = bTmp
        
        If VortexMod Then
            Call bTmp2FontData_Vortex(arr_Ind(i))
         '   Call bArr2CharData_Vortex(arr_Ind(i))
        Else
            Call bTmp2FontData(arr_Ind(i))
         '   Call bArr2CharData(arr_Ind(i))
        End If
        Call bArr2CharData(arr_Ind(i))
        
    Next i

Else 'no file
    'SelInd = 0
    ReDim sColArr(0)
    sColArr(0) = sCol
    ReDim sRowArr(0)
    sRowArr(0) = sRow
    
    Call bArr2CharData(0)
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": StoreCurrentChar()"
End Sub

Private Sub GetAllIndexesOfSameGlyph(Ind As Integer, ByRef arr_Ind() As Integer)
'ind = selind
'fill array with dups: many pointers to 1 glyph
Dim i As Long, n As Long
On Error GoTo frmErr

If Not fFileOpen Then Exit Sub

If Block1Flag Then
    For i = 0 To UBound(FontBlock1Arr)
        If FontBlock1Arr(i) = FontBlock1Arr(Ind) Then
            ReDim Preserve arr_Ind(n)
            arr_Ind(n) = i
            n = n + 1
        End If
    Next i

Else
'2block
    For i = 0 To UBound(FontBlock2Arr)
        If FontBlock2Arr(i) = FontBlock2Arr(Ind) Then
            ReDim Preserve arr_Ind(n)
            arr_Ind(n) = i
            n = n + 1
        End If
    Next i

End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": GetAllIndexesOfSameGlyph()"
End Sub
Private Sub LoadIni()
'and lang
Dim sectArr() As String
Dim s() As String
Dim n As Integer, i As Integer
Dim WFD As WIN32_FIND_DATA
Dim Ret As Long
Dim Tmp As String
Dim t As Single, l As Single, w As Single, h As Single
On Error GoTo frmErr
'Debug.Print ">   LoadIni"

iniFileName = App.Path & "\VTCFont.ini"
Ret = FindFirstFile(iniFileName, WFD)
If Ret < 0 Then
    MsgBoxEx ArrMsg(8), , , CenterScreen, vbCritical    '"Error: VTCFont.ini not found!"
    FindClose Ret
    NoIniFlag = True
'Unload Me

Else    'load from Ini

    Tmp = VBGetPrivateProfileString("Global", "Top", iniFileName)
    If Len(Tmp) <> 0 Then t = Val(Tmp) 'Me.top = Val(Tmp)
    Tmp = VBGetPrivateProfileString("Global", "Left", iniFileName)
    If Len(Tmp) <> 0 Then l = Val(Tmp) 'Me.left = Val(Tmp)
    Tmp = VBGetPrivateProfileString("Global", "Width", iniFileName)
    If Len(Tmp) <> 0 Then w = Val(Tmp) 'Me.Width = Val(Tmp)
    Tmp = VBGetPrivateProfileString("Global", "Height", iniFileName)
    If Len(Tmp) <> 0 Then h = Val(Tmp) 'Me.Height = Val(Tmp)
    Me.Move l, t, w, h
    
    Tmp = VBGetPrivateProfileString("Global", "SaveBoth", iniFileName)
    If Len(Tmp) <> 0 Then chkDupFont.Value = Val(Tmp)
    Tmp = VBGetPrivateProfileString("Global", "WordsInLine", iniFileName)
    If Len(Tmp) <> 0 Then AllWordsInLineFlag = CBool(Tmp)


'fill cmbHard
    n = GetSectionNames(iniFileName, sectArr)
    ReDim IDpassword(n - 1)
    ReDim IDsecuredAdr(n - 1)
    ReDim IDsecuredByte(n - 1)
    ReDim ISVortex(n - 1)
    For i = 0 To n
        If LCase(sectArr(i)) <> "global" Then
            cmbHard.AddItem sectArr(i)    'add HW name
'get all IDpassword
            Tmp = VBGetPrivateProfileString(sectArr(i), "IDpassword", iniFileName)
            IDpassword(cmbHard.NewIndex) = Trim(Tmp)
            
            Tmp = Trim(VBGetPrivateProfileString(sectArr(i), "ISVortex", iniFileName))
            If Len(Tmp) = 1 Then
                ISVortex(cmbHard.NewIndex) = CBool(Tmp)
            End If
            
            Tmp = VBGetPrivateProfileString(sectArr(i), "IDsecured", iniFileName)
            If Len(Tmp) <> 0 Then
                s = Split(Trim(Tmp), ",")
                If UBound(s) = 1 Then
                    IDsecuredAdr(cmbHard.NewIndex) = Val("&H" & s(0))
                    IDsecuredByte(cmbHard.NewIndex) = Val("&H" & s(1))
                End If
            End If
        End If
    Next i
    
    Tmp = VBGetPrivateProfileString("Global", "LastHardware", iniFileName)
    If Len(Tmp) <> 0 Then
        cmbHard.ListIndex = Val(Tmp)
    End If

    Tmp = VBGetPrivateProfileString("Global", "CheckCharSize", iniFileName)
    CheckCharSizeFlag = True
    If Len(Tmp) <> 0 Then CheckCharSizeFlag = Val(Tmp)

    Tmp = VBGetPrivateProfileString("Global", "EditorSize", iniFileName)
    If Len(Tmp) <> 0 And IsNumeric(Val(Tmp)) Then XYspace = Val(Tmp)

    Tmp = VBGetPrivateProfileString("Global", "EditorWidth", iniFileName)
    If Len(Tmp) <> 0 And IsNumeric(Val(Tmp)) Then sCol = Val(Tmp)
    Tmp = VBGetPrivateProfileString("Global", "EditorHeight", iniFileName)
    If Len(Tmp) <> 0 And IsNumeric(Val(Tmp)) Then sRow = Val(Tmp)

    Tmp = VBGetPrivateProfileString("Global", "MagnifyBy", iniFileName)
    If Len(Tmp) <> 0 And IsNumeric(Val(Tmp)) Then MagnifyBy = Val(Tmp)

    Tmp = VBGetPrivateProfileString("Global", "Magnify", iniFileName)
'PicReal.Visible = True
    If Len(Tmp) <> 0 Then Magnify = CBool(Tmp)
    If Magnify Then Call PicReal_Change
'        picX3.Visible = True
'        PicReal.Visible = False
'
'    Else
'        picX3.Visible = False
'        PicReal.Visible = True
'    End If

    Tmp = VBGetPrivateProfileString("Global", "PasteByNumber", iniFileName)
    If Len(Tmp) <> 0 Then chkByNumber.Value = Val(Tmp)
    PasteByNumber = chkByNumber.Value

    Tmp = VBGetPrivateProfileString("Global", "Language", iniFileName)
    If Len(Tmp) <> 0 Then Language = LCase(Trim(Tmp))

    Tmp = VBGetPrivateProfileString("Global", "InvertMouseB", iniFileName)
    If Len(Tmp) <> 0 Then InvertMouseB = Val(Tmp)

'    tmp = VBGetPrivateProfileString("Global", "AlwaysHex", iniFileName)
'    If Len(tmp) <> 0 Then AlwaysHex = Val(tmp)
    Tmp = VBGetPrivateProfileString("Global", "PicDithered", iniFileName)
    If IsNumeric(Tmp) Then fPicDithered = Val(Tmp)
    If fPicDithered > 3 Then fPicDithered = 1

    Tmp = VBGetPrivateProfileString("Global", "LastOpenedFW", iniFileName)
    If Len(Tmp) <> 0 Then LastOpenedFW = Tmp

    Tmp = VBGetPrivateProfileString("Global", "ResizeWidth", iniFileName)
    If IsNumeric(Tmp) Then QwickResizeWidth = Val(Tmp) Else QwickResizeWidth = 64

End If

'Select Case Language
'Case "ru"
'    lngFileName = App.Path & "\VTCFont_Ru.lng"
'Case "en"
'    lngFileName = App.Path & "\VTCFont_En.lng"
'End Select
If Len(Language) = 0 Then Language = "en"
lngFileNameOnly = "VTCFont_" & Language & ".lng"
lngFileName = App.Path & "\" & lngFileNameOnly

Ret = FindFirstFile(lngFileName, WFD)
If Ret < 0 Then
'lngFileName = App.Path & "\VTCFont_En.lng"
    MsgBoxEx ArrMsg(11) & mySpace & lngFileName, , , CenterScreen, vbExclamation     '"Warning: lang file not found!"
    FindClose Ret
    lngFileName = vbNullString
'cmdINIShow_Click (0)
End If


'''
Exit Sub
frmErr:
MsgBox Err.Description & ": LoadIni()"
End Sub

Private Sub FillMCList()
'all fill sub
Dim X As Integer    'SelInd
Dim Tmp As String
'Dim startTimer As Long
'startTimer = GetTickCount()
'Debug.Print GetTickCount() - startTimer

On Error GoTo frmErr
'Debug.Print ">   FillMCList"

McListBox1.Clear

If cmbAdr.ListCount = 0 Then Exit Sub
If (Not Block1Flag) And NoBlock2Flag Then Exit Sub

ReDim ImArr(cmbAdr.ListCount)


PicReal.Visible = False
picX3.Visible = False

Me.MousePointer = vbHourglass
MCListFilling = True

ReDim sRowArr(cmbAdr.ListCount)
ReDim sColArr(cmbAdr.ListCount)
ReDim CharDataArr(cmbAdr.ListCount)
ReDim CharDataHEXArr(cmbAdr.ListCount)
ReDim FontDataArr(cmbAdr.ListCount)
ReDim ChangesIndArr(cmbAdr.ListCount)

For X = 0 To cmbAdr.ListCount - 1

    'DoEvents
    If EscFlag Then Exit For

    startAddr = "&H" & cmbAdr.List(X)

    If VortexMod Then
        Call GetBlock_Vortex(X)
    Else
        Call GetBlock(X)
    End If

    Set ImArr(X) = PicReal.Picture

    'Tmp = Right("00" & Hex(X + 1), 3)

    If VortexMod Then
    Tmp = Hex(X + 32)
'        If Block1Flag Then
'            'Tmp = X + 32
'            Tmp = Hex(X + 32)
'        Else
'            Tmp = Hex(X + 1)
'            'Tmp = X + 1
'        End If
    Else
        Tmp = Hex(X + 1)
    End If
    'McListBox1.AddItem tmp & "/" & cmbAdr.List(X), -1, CLng(X), False

    McListBox1.AddItem Tmp, -1, CLng(X), False

Next X

If Magnify Then
    picX3.Visible = True
Else
    PicReal.Visible = True
End If

MCListFilling = False

cmdSaveAll.Caption = ArrMsg(12) & " (" & GetChangesCount & ")"

Me.MousePointer = vbNormal


McListBox1.Refresh


'Debug.Print GetTickCount() - startTimer
'''
Exit Sub
frmErr:
Me.MousePointer = vbNormal
MsgBox Err.Description & ": FillMCList()"
End Sub


Private Sub cmdViewWord_Click()
Call cmbVocab_KeyPress(13)
End Sub



Private Sub cmdXY_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
AlwaysHex = Not AlwaysHex
If AlwaysHex Then
    cmdXY(0).Caption = "x" & Hex(sCol)
    cmdXY(1).Caption = "y" & Hex(sRow)
    
Else
    cmdXY(0).Caption = "X" & sCol
    cmdXY(1).Caption = "Y" & sRow

End If
End If

End Sub

Private Sub Command1_Click()
'check dups
Dim i As Integer, j As Integer ', n As Integer
'Dim aIDP() As Long
'Dim aIDSA() As String
'Dim aIDSB() As String
'Dim aTMP() As Long

For i = 0 To UBound(IDpassword)
    For j = 0 To UBound(IDpassword)
        If i <> j Then
        
            If IDpassword(j) = IDpassword(i) Then
                Debug.Print IDpassword(i), Hex(IDsecuredAdr(i)), Hex(IDsecuredByte(i)), cmbHard.List(i)
'                ReDim Preserve aTMP(n)
'                ReDim Preserve aIDP(n)
'                ReDim Preserve aIDSA(n)
'                ReDim Preserve aIDSB(n)
'                aTMP(n) = "&H" & IDpassword(i)
'                aIDP(n) = "&H" & IDpassword(i)
'                aIDSA(n) = Hex(IDsecuredAdr(i))
'                aIDSB(n) = Hex(IDsecuredByte(i))
                'n = n + 1
            End If
     
        End If
    Next j
Next i

''sort
'Call BubbleSort(aTMP)
''print
'For i = 0 To UBound(aTMP)
'    For j = 0 To UBound(aIDP)
'        If aIDP(j) = aTMP(i) Then
'            Debug.Print Hex(aIDP(j)), aIDSA(j), aIDSB(j)
'            Exit For
'        End If
'    Next j
'Next i

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call XYcaptionSet(sCol, sRow)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Ret As Long
On Error Resume Next

If fFileOpen Then
    If GetChangesCount > 0 Then
        Ret = MsgBoxEx(ArrMsg(20), , , CenterOwner, vbOKCancel Or vbQuestion)
        If Ret <> 1 Then Cancel = True
    End If
End If
End Sub

Private Sub HScrollDraw_Change()
Call HScrollDraw_Scroll
End Sub

Private Sub HScrollDraw_Scroll()
picContainer.left = HScrollDraw.Value
End Sub

Private Sub lblTTF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTTF.ToolTipText = TTFontName
End Sub

Private Sub McListBox1_DbClick()
cmdChar_Click
End Sub

Private Sub McListBox1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 65 And Shift = 2 Then Call mnu_SelectAll_Click
End Sub

Private Sub McListBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseUpBug = False

If Button = vbRightButton Then
    If GoExportFlag Then
        mnu_GoExport.Enabled = True
        mnu_Export(1).Enabled = True
        mnu_Export(2).Enabled = True
        mnu_Export(3).Enabled = True
        mnu_UpdateExport.Enabled = True
    Else
        mnu_GoExport.Enabled = False
        mnu_Export(1).Enabled = False
        mnu_Export(2).Enabled = False
        mnu_Export(3).Enabled = False
        mnu_UpdateExport.Enabled = False
    End If
    Me.PopupMenu mnu_MCList, vbPopupMenuCenterAlign
End If
End Sub


Private Sub mnu_GoExport_Click()
Call SaveExport(True)
End Sub


Private Sub mnu_SelectAll_Click()
Dim i As Integer
On Error GoTo frmErr

McListBox1.Visible = False
Call McListBox1.SelectAll
McListBox1.Visible = True
cmdCopy.Caption = ArrMsg(13) & " (" & McListBox1.SelCount & ")"

ReDim selArr(McListBox1.ListCount)
For i = 0 To UBound(selArr) - 1
    selArr(i + 1) = i
Next i

'''
Exit Sub
frmErr:
McListBox1.Visible = True
MsgBox Err.Description & ": mnu_SelectAll"
End Sub

Private Sub mnu_Export_Click(Index As Integer)
Dim i As Integer
On Error GoTo frmErr

'LockWindowUpdate frmmain.hWnd
McListBox1.Visible = False

Select Case Index
Case 0
    If McListBox1.SelCount = 1 Then
        arrExport(cmbLastIndex) = True
    Else
        For i = 1 To UBound(selArr)
            arrExport(selArr(i)) = True
'McListBox1.ListBold(selArr(i)) = True
        Next i
    End If

Case 1
    If McListBox1.SelCount = 1 Then
        arrExport(cmbLastIndex) = False
'McListBox1.ListBold(cmbLastIndex) = False
    Else
        For i = 1 To UBound(selArr)
            arrExport(selArr(i)) = False
' McListBox1.ListBold(selArr(i)) = False
        Next i
    End If

Case 2
    McListBox1.ClearSelectionAll
    For i = 0 To UBound(arrExport)
        If arrExport(i) Then
            McListBox1.SelectItem (i)
'selArr(i + 1) = i + 1 'selarr change on right click
        End If
    Next i
    McListBox1.Refresh
    cmdCopy.Caption = ArrMsg(13) & " (" & McListBox1.SelCount & ")"
Case 3    'clear list

    For i = 0 To UBound(arrExport)
        arrExport(i) = False
    Next i

End Select
'LockWindowUpdate 0
McListBox1.Visible = True

GoExportFlag = False
For i = 0 To UBound(arrExport)
    If arrExport(i) Then
        GoExportFlag = True
        Exit For
    End If
Next i

McListBox1.SetFocus

'''
Exit Sub
frmErr:
McListBox1.Visible = True
MsgBox Err.Description & ": mnu_Export"
End Sub

Private Sub mnu_UpdateExport_Click()
Call AppendImport
End Sub

Private Sub optBlock_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'todo and into keydown...
Dim Ret As Long
'Dim Ind As Integer
On Error GoTo frmErr

If Block1Flag Then
    If Index = 0 Then Exit Sub
Else
    If Index = 1 Then Exit Sub
End If

noChangeBlockFlag = False
If fFileOpen Then
    If GetChangesCount > 0 Then
        Ret = MsgBoxEx(ArrMsg(23), , , CenterOwner, vbOKCancel Or vbQuestion)   'unsaved
        If Ret <> 1 Then
            noChangeBlockFlag = True
'If Index = 0 Then ind = 1 Else ind = 0
'optBlock_Click (ind)
            Exit Sub
        Else
            optBlock(Index).Value = vbChecked

'optBlock_Click (Index)
        End If
    End If
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": optBlock"
End Sub

Private Sub picContainer_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode 'del 46
On Error GoTo frmErr

If KeyCode = 46 And isSelection Then

' Erase rect from picreal
    PicReal.Line _
            (old_X1, old_Y1)-(old_X2 - 1, old_Y2 - 1), _
            PicReal.ForeColor, BF

    PicReal.Picture = PicReal.Image

    Call Draw2bArr(PicReal, sCol, sRow)
    Call StoreCurrentChar(cmbAdr.ListIndex)
    Call bArr2PicDraw(-1)

    Call StoreInUndoBuffer
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": picContainer_KeyDown()"
End Sub



Private Sub picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Xm As Integer, Ym As Integer
On Error GoTo frmErr
MouseUpBug = False

If chkSelection.Value = vbChecked Then
    Xm = (X \ XYspace)
    Ym = (Y \ XYspace)
    old_Xm = Xm
    old_Ym = Ym
    Call OrderCorners
    If isSelection And ((X >= X1Region) And (X <= X2Region) And (Y >= Y1Region) And (Y <= Y2Region)) Then
'move sel rect
        Select Case Button

        Case vbRightButton
'move content of rect
            fMoveSelRect = True
            fCopySelection = False
            fMoveSelection = True

            If fMoveSelection Then
' Erase rect from picreal
                PicReal.Line _
                        (old_X1, old_Y1)-(old_X2 - 1, old_Y2 - 1), _
                        PicReal.BackColor, BF
                PicReal.Picture = PicReal.Image
            End If

        Case vbMiddleButton
'move only rect
            fMoveSelRect = True
            fCopySelection = False
            fMoveSelection = False

        Case vbLeftButton

            If Shift = 0 Then
'copy content of rect
                fMoveSelRect = True
                fCopySelection = True
                fMoveSelection = False

            ElseIf Shift = 2 Then    'ctrl
'move content of rect
                fMoveSelRect = True
                fCopySelection = False
                fMoveSelection = True

                If fMoveSelection Then
' Erase rect from picreal
                    PicReal.Line _
                            (old_X1, old_Y1)-(old_X2 - 1, old_Y2 - 1), _
                            PicReal.BackColor, BF
                    PicReal.Picture = PicReal.Image
                End If

            ElseIf Shift = 1 Then    'shift
'move only rect
                fMoveSelRect = True
                fCopySelection = False
                fMoveSelection = False

            End If

        End Select

    Else
'start Selecting
        fMoveSelRect = False
        fCopySelection = False
        fMoveSelection = False

        isSelection = False    'rem prew selection
        Call bArr2PicDraw(-1)
        X1Region = 0: Y1Region = 0: X2Region = 0: Y2Region = 0

        If Button = vbLeftButton Then

            start_sel_X = Xm
            start_sel_Y = Ym
'  isSelection = False

        End If
    End If
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": picContainer_MouseDown()"
End Sub


Private Sub PicReal_Change()
'todo 2 раза
Dim autoMagnify As Long
Dim i As Long

On Error GoTo frmErr
'no If Not PicReal.Visible Then Exit Sub
'If picX3.Visible Then    'copy to magnify pic
'If PicReal.Height = oldPicRealH And PicReal.Width = oldPicRealW Then Exit Sub
'oldPicRealH = PicReal.Height: oldPicRealW = PicReal.Width

If MagnifyBy < 2 Then MagnifyBy = 2
autoMagnify = MagnifyBy
If MCListFilling Or DrawAllWordsFlag Then
'
Else

    AllWordsShow = False

    For i = MagnifyBy To 2 Step -1
        If PicReal.ScaleHeight * autoMagnify > 222 Then
            autoMagnify = autoMagnify - 1
        Else
            Exit For
        End If
    Next i


    If PicReal.ScaleHeight * autoMagnify <= 222 Then
        If Magnify Then
            picX3.Visible = True
            PicReal.Visible = False
        Else
            picX3.Visible = False
            PicReal.Visible = True
        End If
    Else
        picX3.Visible = False
        PicReal.Visible = True
    End If

    picX3.Width = autoMagnify * PicReal.Width
    picX3.Height = autoMagnify * PicReal.Height

    If picX3.Visible Then    'lastsur032017
        StretchBlt picX3.hdc, 0, 0, picX3.Width, picX3.Height, _
                   PicReal.hdc, 0, 0, PicReal.Width, PicReal.Height, vbSrcCopy

        picX3.Picture = picX3.Image
    End If

End If
'End If


'''
Exit Sub
frmErr:
MsgBox Err.Description & ": FillMCList()"
End Sub

Private Sub PicReal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim i As Integer, j As Integer
On Error Resume Next

If Button = vbRightButton Then

    If AllWordsShow Then    'all words picture present
        PicReal.Line (oldAllWordsStartCoord_X, oldAllWordsFinishCoord_Y)-(oldAllWordsFinishCoord_X, oldAllWordsFinishCoord_Y), lcForeColor
    End If

    '    mnu_SaveBMP(1).Visible = False
    '   mnu_CopyPic(1).Visible = False
    Me.PopupMenu mnu_RealPic, vbPopupMenuCenterAlign

ElseIf Button = vbLeftButton Then

    If AllWordsShow Then    'all words picture present
        If AllWordsInd > -1 Then
            If cmbVocab.ListIndex = AllWordsInd Then 'same
                cmdViewWord_Click
            Else
                cmbVocab.ListIndex = AllWordsInd
                cmbVocAdr.ListIndex = AllWordsInd
            End If
        End If

        '        For i = 0 To UBound(AllWordsStartCoord_X)
        '            If X - 5 > AllWordsStartCoord_X(i) And Y - 5 > AllWordsStartCoord_Y(i) Then
        '                If X - 5 < AllWordsFinishCoord_X(i) And Y - 5 < AllWordsFinishCoord_Y(i) Then
        '                    cmbVocab.ListIndex = i
        '                End If
        '            End If
        '        Next i
    End If

End If
End Sub
Private Sub mnu_CopyPic_click(Index As Integer)
On Error GoTo frmErr

Clipboard.Clear
If Index = 1 Then

    Set tmpPic = Nothing
    tmpPic.Width = MagnifyBy * PicReal.Width
    tmpPic.Height = MagnifyBy * PicReal.Height

    StretchBlt tmpPic.hdc, 0, 0, tmpPic.Width, tmpPic.Height, _
               PicReal.hdc, 0, 0, PicReal.Width, PicReal.Height, vbSrcCopy

    tmpPic.Picture = tmpPic.Image

    Clipboard.SetData tmpPic.Picture
    
Else
'PicReal.Picture = PicReal.Image
    Clipboard.SetData PicReal.Picture
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": mnu_CopyReal"
End Sub
Private Sub mnu_SaveBMP_click(Index As Integer)
Dim fn As String
'Dim FileTitle As String
Dim ExportDir As String
'Dim sClipFont As String
'Dim f As Integer
'Dim strData As String
On Error GoTo frmErr

'f = FreeFile
ExportDir = App.Path & "\Export"

fn = vbNullString
fn = BMPSaveDialog(ExportDir, ArrMsg(21))
If fn = vbNullString Then Exit Sub
'LastPath = GetPathFromPathAndName(FileName)

If Index = 1 Then


    Set tmpPic = Nothing
    tmpPic.Width = MagnifyBy * PicReal.Width
    tmpPic.Height = MagnifyBy * PicReal.Height

    StretchBlt tmpPic.hdc, 0, 0, tmpPic.Width, tmpPic.Height, _
               PicReal.hdc, 0, 0, PicReal.Width, PicReal.Height, vbSrcCopy

    tmpPic.Picture = tmpPic.Image


    Call SavePictureBW(tmpPic, fn)
    ' Call SaveBMP1bit(picX3, fn)
    'SavePicture picX3.Image, fn

Else
    Call SavePictureBW(PicReal, fn)
    'Call SaveBMP1bit(PicReal, fn)
    'SavePicture PicReal.Image, fn
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": mnu_SaveBMP"
End Sub

Private Sub PicReal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
On Error Resume Next

If AllWordsShow Then    'all words picture present
    AllWordsInd = -1

    For i = 0 To UBound(AllWordsStartCoord_X)

        If X - 5 > AllWordsStartCoord_X(i) And Y - 5 > AllWordsStartCoord_Y(i) Then
            If X - 5 < AllWordsFinishCoord_X(i) And Y - 5 < AllWordsFinishCoord_Y(i) Then

                PicReal.Line (oldAllWordsStartCoord_X, oldAllWordsFinishCoord_Y)-(oldAllWordsFinishCoord_X, oldAllWordsFinishCoord_Y), lcForeColor

'PicReal.Line (AllWordsStartCoord_X(i) + 5, AllWordsStartCoord_Y(i) + 5)-(AllWordsFinishCoord_X(i) + 5, AllWordsFinishCoord_Y(i) + 5), lcBackColor, B

                oldAllWordsStartCoord_X = AllWordsStartCoord_X(i) + 3
'oldAllWordsStartCoord_Y
                oldAllWordsFinishCoord_X = AllWordsFinishCoord_X(i) + 2
                oldAllWordsFinishCoord_Y = AllWordsFinishCoord_Y(i) + 6
'If oldAllWordsFinishCoord_X - oldAllWordsStartCoord_X < 25 Then oldAllWordsFinishCoord_X = oldAllWordsFinishCoord_X + 30

                PicReal.Line (oldAllWordsStartCoord_X, oldAllWordsFinishCoord_Y)-(oldAllWordsFinishCoord_X, oldAllWordsFinishCoord_Y), lcBackColor
'PicReal.Line (AllWordsStartCoord_X(i) + 5, AllWordsFinishCoord_Y(i) + 6)-(AllWordsFinishCoord_X(i), AllWordsFinishCoord_Y(i) + 6), lcBackColor

'                    cmbVocab.ListIndex = i
                AllWordsInd = i
            End If
        End If
    Next i
End If

End Sub



Private Sub picX3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim i As Integer
On Error Resume Next

If Button = vbRightButton Then
'    mnu_SaveBMP(1).Visible = False
'   mnu_CopyPic(1).Visible = False
    Me.PopupMenu mnu_RealPic, vbPopupMenuCenterAlign

'ElseIf Button = vbLeftButton Then
'    If AllWordsShow Then    'all words picture present
'
'        For i = 0 To UBound(AllWordsStartCoord_X)
'
'            If X - 5 > AllWordsStartCoord_X(i) And Y - 5 > AllWordsStartCoord_Y(i) Then
'                If X - 5 < AllWordsFinishCoord_X(i) And Y - 5 < AllWordsFinishCoord_Y(i) Then
'
'                    cmbVocab.ListIndex = i
'                End If
'            End If
'        Next i
'    End If
End If
End Sub

Private Sub test_click()
'Dim r As Integer, c As Integer
'Debug.Print "-----start------"
'For r = 0 To UBound(bArr, 1)
'For c = 0 To UBound(bArr, 2)
'Debug.Print bArr(r, c);
'Next c
'Debug.Print
'Next r
'Debug.Print "-----FIN--------"

'Dim Contrl As Control
'For Each Contrl In frmPatch.Controls
'Debug.Print Contrl.Name
'Next

End Sub
Private Sub SetUpPicScroll()
Dim w1 As Long, h1 As Long
Dim w2 As Long, h2 As Long
On Error Resume Next

'''''''''
'Call SetUpPicScroll

'        PicTTF.Left = PicScroll.Left + PicScroll.Width + 40    '2
'    End If

If HScrollDraw.Visible Then
    h1 = Me.ScaleHeight - PicScroll.top - 30
Else
    h1 = Me.ScaleHeight - PicScroll.top - 8
End If

h2 = picContainer.top + picContainer.Height - VScrollDraw.Value
If h1 > h2 Then
    PicScroll.Height = h2
Else
    PicScroll.Height = h1
End If

'If PicTTF.Visible Then
'    PicScroll.Width = PicTTF.Left - PicScroll.Left - 30
'
'Else
If VScrollDraw.Visible Then
    w1 = Me.ScaleWidth - PicScroll.left - 30
Else
    w1 = Me.ScaleWidth - PicScroll.left - 8
End If

w2 = picContainer.left + picContainer.Width - HScrollDraw.Value
If w1 > w2 Then
    PicScroll.Width = w2
Else
    PicScroll.Width = w1
End If
'End If
On Error GoTo 0
End Sub
Private Sub SetUpScrollBars()
'and scroll bars
On Error Resume Next

'PicTTF.Visible
'VScrollDraw
'HScrollDraw

VScrollDraw.top = PicScroll.top
VScrollDraw.left = PicScroll.left + PicScroll.Width + 4
VScrollDraw.Height = PicScroll.Height

HScrollDraw.top = PicScroll.top + PicScroll.Height + 4
HScrollDraw.left = PicScroll.left
HScrollDraw.Width = PicScroll.Width

VScrollDraw.Max = PicScroll.Height - picContainer.Height
HScrollDraw.Max = PicScroll.Width - picContainer.Width

If VScrollDraw.Max >= 0 Then
    VScrollDraw.Visible = False
Else
    VScrollDraw.Visible = True
End If
If HScrollDraw.Max >= 0 Then
    HScrollDraw.Visible = False
Else
    HScrollDraw.Visible = True
End If
VScrollDraw.SmallChange = XYspace
HScrollDraw.SmallChange = XYspace

VScrollDraw.LargeChange = PicScroll.Height
HScrollDraw.LargeChange = PicScroll.Width

Call VScrollDraw_Scroll
Call HScrollDraw_Scroll

On Error GoTo 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print KeyCode, Shift
On Error Resume Next

If KeyCode = 27 Then
    EscFlag = True
    If isSelection Then
        isSelection = False
        Call bArr2PicDraw(-1)    '2
        X1Region = 0: Y1Region = 0: X2Region = 0: Y2Region = 0
    Else
        If chkSelection.Value = vbChecked Then chkSelection.Value = vbUnchecked
    End If
    
End If

If Shift <> 4 Then Exit Sub
'Debug.Print KeyCode
Select Case KeyCode

Case 86
    Call cmdPaste_Click
Case 67
    Call cmdCopy_Click
Case 37
    Call cmdToolBar_click(0)
    picContainer.SetFocus
Case 39
    Call cmdToolBar_click(1)
    picContainer.SetFocus
Case 38
    Call cmdToolBar_click(2)
    picContainer.SetFocus
Case 40
    Call cmdToolBar_click(3)
    picContainer.SetFocus
Case 36
    Call cmdToolBar_click(4)
    picContainer.SetFocus
Case 33
    Call cmdToolBar_click(5)
    picContainer.SetFocus
Case 12
    Call cmdToolBar_click(6)
    picContainer.SetFocus
Case 45
    Call cmdToolBar_click(7)
    picContainer.SetFocus
Case 35
    Call cmdToolBar_click(9)
    picContainer.SetFocus
Case 34
    Call cmdToolBar_click(8)
    picContainer.SetFocus
Case 88    'x
    Call cmdXY_Click(0)
'picContainer.SetFocus
Case 89    'y
    Call cmdXY_Click(1)
'picContainer.SetFocus
Case 72    'h
    Call cmdChar_Click
'picContainer.SetFocus
Case 83    's
    Call cmdSave_Click
'picContainer.SetFocus
Case 79    'o
    Call cmdINIShow_Click(0)
Case 80    'P
    Call cmdPatcher_Click

End Select

On Error GoTo 0
End Sub

Private Sub Form_Load()
Dim g As Integer

On Error GoTo frmErr


Me.Caption = "VTCFont v" & App.Major & "." & App.Minor & "." & App.Revision

ReDim ImArr(0)
ReDim selArr(0)
IDanswer(0) = 128: IDanswer(1) = 81: IDanswer(2) = 1 '80 51 01 hex

'VORTEX 56 4F 52 54 45 58
IDVortex(0) = &H56: IDVortex(1) = &H4F: IDVortex(2) = &H52

Me.ScaleMode = vbPixels
Block1Flag = True

sCol = 8: sRow = 16
XYspace = 10
lcBackColor = vbWhite    'vbWindowBackground
lcForeColor = vbBlack    'vbButtonFace
'no PicReal.Picture = PicReal.Image

Call SetMessages    'for msgbox
Call LoadIni
If Len(lngFileName) <> 0 Then Call GetLanguage(1)

MaxUndoCircle = 20
picCount = -1
ReDim UndoBuffer(0)


ReDim bArr(sRow - 1, sCol - 1)

If XYspace = 1 Then
    chkGridFlag = vbUnchecked
Else
    chkGridFlag = chkGrid.Value
End If
If chkGridFlag = vbChecked Or chkSelection.Value = vbChecked Then g = 1

With picContainer
    .ScaleMode = vbPixels
    .AutoSize = False
    .AutoRedraw = True
    .Height = sRow * XYspace + g
    .Width = sCol * XYspace + g
    .BackColor = lcBackColor
    .ForeColor = lcForeColor
End With
With PicReal
    .ScaleMode = vbPixels
    .Width = sCol
    .Height = sRow
    .AutoSize = False
    .AutoRedraw = True
    .BackColor = lcForeColor    'lcBackColor
    .ForeColor = lcForeColor
'.Picture = PicReal.Image
End With
With picX3
    .ScaleMode = vbPixels
    .Width = sCol * 2 ' for empty start * MagnifyBy
    .Height = sRow * 2 '* MagnifyBy
    .AutoSize = False
    .AutoRedraw = True
    .BackColor = lcForeColor    'lcBackColor
    .ForeColor = lcForeColor
'.Picture = PicReal.Image
End With
picContainer.Cls
picContainer.Picture = picContainer.Image

Call StoreInUndoBuffer

If chkGridFlag = vbChecked Then DrawGrid picContainer, vbGrayText

Call XYcaptionSet(sCol, sRow)

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Load()"
End Sub


Private Sub Form_Resize()
Dim i As Integer
On Error Resume Next

If Me.Visible Then
    If Not noResize Then

        i = ((Me.ScaleHeight - 16) \ McListBox1.RowHeight) * McListBox1.RowHeight + 1
        If i <> McListBox1.Height Then
            McListBox1.Height = i
        End If

        Call SetUpPicScroll
        Call SetUpScrollBars
        Call SetUpPicScroll
        Call SetUpScrollBars

        If PicTTF.Visible Then
            PicTTF.left = PicScroll.left + PicScroll.Width + 40    '2
        End If

    End If
End If

On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

Close #bFileIn
'Set FontData = Nothing
'Set sCharData = Nothing
Set m_Preview = Nothing

Call WriteIni

End Sub
Private Sub WriteIni()
On Error Resume Next
If Not NoIniFlag Then
    WriteKey "Global", "Top", CStr(Me.top), iniFileName
    WriteKey "Global", "Left", CStr(Me.left), iniFileName
    WriteKey "Global", "Width", CStr(Me.Width), iniFileName
    WriteKey "Global", "Height", CStr(Me.Height), iniFileName

    WriteKey "Global", "EditorSize", CStr(XYspace), iniFileName

    WriteKey "Global", "PasteByNumber", CStr(chkByNumber.Value), iniFileName
    WriteKey "Global", "SaveBoth", CStr(chkDupFont.Value), iniFileName
    WriteKey "Global", "LastHardware", CStr(cmbHard.ListIndex), iniFileName

    If Len(FileNameFW) <> 0 Then WriteKey "Global", "LastOpenedFW", FileNameFW, iniFileName

End If
On Error GoTo 0
End Sub
Private Sub optBlock_Click(Index As Integer)
Dim i As Integer

On Error GoTo frmErr

If noChangeBlockFlag Then Exit Sub

Select Case Index
Case 0
    Block1Flag = True    'block1
Case 1
    Block1Flag = False
End Select

If Not fFileOpen Then
    MsgBoxEx ArrMsg(0), , , CenterOwner, vbExclamation    '"Open firmware file first!"
    Exit Sub
End If


'ReDim selArr(0)    '1 based
'If McListBox1.SelCount > 0 Then
'    ReDim selArr(McListBox1.SelCount)
'    For i = 1 To McListBox1.SelCount
'        selArr(i) = McListBox1.SelItem(i - 1)
'    Next i
'End If
'LockWindowUpdate frmmain.hWnd

picContainer.Visible = False    'need
McListBox1.Visible = False

If VortexMod Then
    Call FillVocab_Vortex
    Call FillFontList_Vortex

Else

    Call FillVocab
    Call FillFontList
End If

Call FillMCList

picContainer.Visible = True    'need
McListBox1.Visible = True
'LockWindowUpdate 0

reloadFW_flag = True    'for setup old position in list
If cmbLastIndex > cmbAdr.ListCount - 1 Then cmbLastIndex = 0
If UBound(selArr) < 2 Then
    'McListBox1.ClearSelectionAll
    cmbAdr.ListIndex = cmbLastIndex    'cmbAdr_Click return first
    'cmbAdr.SetFocus
Else
    cmbAdr.ListIndex = selArr(1)
    McListBox1.ClearSelectionAll
    If McListBox1.SelCount + UBound(selArr) <= cmbAdr.ListCount Then
        For i = 1 To UBound(selArr)
            McListBox1.SelectItem (selArr(i))
        Next i
    End If
End If
reloadFW_flag = False

McListBox1.Refresh

cmdCopy.Caption = ArrMsg(13) & " (" & McListBox1.SelCount & ")"

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": optBlock_Click()"
End Sub
Private Sub DrawOldSelRect()
On Error GoTo frmErr

If Not isSelection Then Exit Sub
picContainer.DrawMode = vbInvert
picContainer.DrawWidth = LineWidth
picContainer.Line _
        (old_X1 * XYspace + old_xx1, old_Y1 * XYspace + old_yy1)- _
        (old_X2 * XYspace + old_xx2, old_Y2 * XYspace + old_yy2), _
        , B
picContainer.DrawMode = vbCopyPen
picContainer.DrawWidth = 1

X1Region = old_X1 * XYspace + old_xx1
X2Region = old_X2 * XYspace + old_xx2
Y1Region = old_Y1 * XYspace + old_yy1
Y2Region = old_Y2 * XYspace + old_yy2

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": DrawOldSelRect()"
End Sub
Private Sub DrawSelection(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Xm As Integer, Ym As Integer
Dim X1 As Integer, Y1 As Integer
Dim X2 As Integer, Y2 As Integer
Dim xx1 As Integer, yy1 As Integer
Dim xx2 As Integer, yy2 As Integer
'Dim ptColor As Long
On Error GoTo frmErr

If fMoveSelRect Then Exit Sub

LineWidth = 2

If Button = vbLeftButton Then

    picContainer.DrawMode = vbInvert    'no colors
    picContainer.DrawWidth = LineWidth
'    picContainer.DrawStyle = 1 'vbDot

    Xm = X \ XYspace
    Ym = Y \ XYspace
    If Xm >= sCol Then Xm = sCol - 1
    If Ym >= sRow Then Ym = sRow - 1
    If Xm < 0 Then Xm = 0
    If Ym < 0 Then Ym = 0

    If isSelection Then

'del last
        picContainer.Line _
                (old_X1 * XYspace + old_xx1, old_Y1 * XYspace + old_yy1)- _
                (old_X2 * XYspace + old_xx2, old_Y2 * XYspace + old_yy2), _
                , B
    End If

    xx1 = 0: yy1 = 0: xx2 = 0: yy2 = 0

'draw new sel
    If Xm = start_sel_X And Ym = start_sel_Y Then
'.
        X1 = start_sel_X
        Y1 = start_sel_Y
        X2 = Xm + 1
        Y2 = Ym + 1
        xx1 = 1
        yy1 = 1
    End If

    If Xm > start_sel_X And Ym = start_sel_Y Then
'>
        X1 = start_sel_X
        Y1 = start_sel_Y
        X2 = Xm + 1
        Y2 = Ym + 1
        yy1 = 1
        xx1 = 1
    End If

    If Xm < start_sel_X And Ym = start_sel_Y Then
'<
        X1 = Xm
        Y1 = Ym
        X2 = start_sel_X + 1
        Y2 = start_sel_Y + 1
        yy1 = 1
        xx1 = 1
    End If

    If Ym > start_sel_Y And Xm = start_sel_X Then
'down
        X1 = start_sel_X
        Y1 = start_sel_Y
        X2 = Xm + 1
        Y2 = Ym + 1
        yy1 = 1
        xx1 = 1
    End If

    If Ym < start_sel_Y And Xm = start_sel_X Then
'up
        X1 = Xm
        Y1 = Ym
        X2 = start_sel_X + 1
        Y2 = start_sel_Y + 1
        yy1 = 1
        xx1 = 1

    End If

    If Ym > start_sel_Y And Xm > start_sel_X Then
'> down
        X1 = start_sel_X
        Y1 = start_sel_Y
        X2 = Xm + 1
        Y2 = Ym + 1
        yy1 = 1
        xx1 = 1

    End If

    If Ym < start_sel_Y And Xm < start_sel_X Then
'up <
        X1 = Xm
        Y1 = Ym
        X2 = start_sel_X + 1
        Y2 = start_sel_Y + 1
        yy1 = 1

        xx1 = 1
    End If

    If Ym > start_sel_Y And Xm < start_sel_X Then
'down <
        X1 = Xm
        Y1 = Ym + 1
        X2 = start_sel_X + 1
        Y2 = start_sel_Y
        xx1 = 1
' yy2 = 1

    End If

    If Ym < start_sel_Y And Xm > start_sel_X Then
'up >
        X1 = start_sel_X
        Y1 = start_sel_Y + 1
        X2 = Xm + 1
        Y2 = Ym

        yy1 = 1
        xx1 = 1
        yy2 = 1
    End If

    picContainer.Line _
            (X1 * XYspace + xx1, Y1 * XYspace + yy1)- _
            (X2 * XYspace + xx2, Y2 * XYspace + yy2), _
            , B

' picContainer.DrawStyle = vbSolid
    picContainer.DrawMode = vbCopyPen
    picContainer.DrawWidth = 1
    isSelection = True
    X1Region = X1 * XYspace + xx1
    X2Region = X2 * XYspace + xx2
    Y1Region = Y1 * XYspace + yy1
    Y2Region = Y2 * XYspace + yy2

    old_X1 = X1: old_Y1 = Y1: old_X2 = X2: old_Y2 = Y2
'Call OrderCorners 'no

    old_xx1 = xx1: old_yy1 = yy1: old_xx2 = xx2: old_yy2 = yy2


    Call XYcaptionSet(Abs(X2 - X1), Abs(Y2 - Y1))


End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": DrawSelection()"
End Sub
Private Sub picContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim tmp1 As Integer, tmp2 As Integer
'Dim Xm As Integer, Ym As Integer
Dim Xm As Long, Ym As Long
On Error GoTo frmErr

If Button = 1 And MouseUpBug Then Exit Sub

If chkSelection Then
    If Button = 0 Then
        Xm = Int(X \ XYspace) + 1
        Ym = Int(Y \ XYspace) + 1
        If Xm > sCol Then Xm = sCol
        If Ym > sRow Then Ym = sRow
        Call XYcaptionSet(Xm, Ym)
        Exit Sub
    End If

'setfocus to SB if need scroll
'Debug.Print X, -HScrollDraw.Value
    If Y > PicScroll.Height - VScrollDraw.Value Then
        If VScrollDraw.Visible And VScrollDraw.Value <> VScrollDraw.Max Then VScrollDraw.SetFocus

    ElseIf Y < -VScrollDraw.Value Then
        If VScrollDraw.Visible And VScrollDraw.Value <> 0 Then VScrollDraw.SetFocus

    ElseIf X > PicScroll.Width - HScrollDraw.Value Then
        If HScrollDraw.Visible And HScrollDraw.Value <> HScrollDraw.Max Then HScrollDraw.SetFocus

    ElseIf X < -HScrollDraw.Value Then
        If HScrollDraw.Visible And HScrollDraw.Value <> 0 Then HScrollDraw.SetFocus

    Else
        picContainer.SetFocus
    End If

    If fMoveSelRect Then
        Call MoveSelection(Button, Shift, X, Y)
    Else
        Call DrawSelection(Button, Shift, X, Y)
    End If

Else
    Call DrawGraf(Button, Shift, X, Y)
End If

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": picContainer_MouseMove()"
End Sub
Private Sub MoveSelection(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Xm As Integer, Ym As Integer
Dim tmp1 As Integer
Dim tmp2 As Integer

On Error GoTo frmErr
'If Button <> vbLeftButton Then Exit Sub
'If Button <> vbMiddleButton Then Exit Sub

Xm = (X \ XYspace)
Ym = (Y \ XYspace)
picContainer.DrawMode = vbInvert
picContainer.DrawWidth = LineWidth

Call OrderCorners

picContainer.Line _
        (old_X1 * XYspace + old_xx1, old_Y1 * XYspace + old_yy1)- _
        (old_X2 * XYspace + old_xx2, old_Y2 * XYspace + old_yy2), _
        , B


tmp1 = old_X1 + (Xm - old_Xm)    ' Prevent Selection from being moved of the left or right
tmp2 = old_X2 + (Xm - old_Xm)
If (tmp1 >= 0) And (tmp2 <= sCol) Then
    old_X1 = old_X1 + (Xm - old_Xm)
    old_X2 = old_X2 + (Xm - old_Xm)
End If

tmp1 = old_Y1 + (Ym - old_Ym)
tmp2 = old_Y2 + (Ym - old_Ym)
If (tmp1 >= 0) And (tmp2 <= sRow) Then
    old_Y1 = old_Y1 + (Ym - old_Ym)
    old_Y2 = old_Y2 + (Ym - old_Ym)
End If

'old_X1 = old_X1 + (Xm - old_Xm)
'old_X2 = old_X2 + (Xm - old_Xm)
'old_Y1 = old_Y1 + (Ym - old_Ym)
'old_Y2 = old_Y2 + (Ym - old_Ym)

old_Xm = Xm: old_Ym = Ym

picContainer.Line _
        (old_X1 * XYspace + old_xx1, old_Y1 * XYspace + old_yy1)- _
        (old_X2 * XYspace + old_xx2, old_Y2 * XYspace + old_yy2), _
        , B

X1Region = old_X1 * XYspace + old_xx1
X2Region = old_X2 * XYspace + old_xx2
Y1Region = old_Y1 * XYspace + old_yy1
Y2Region = old_Y2 * XYspace + old_yy2

'old_X1 = X1: old_Y1 = Y1: old_X2 = X2: old_Y2 = Y2
'old_xx1 = xx1: old_yy1 = yy1: old_xx2 = xx2: old_yy2 = yy2

picContainer.DrawMode = vbCopyPen
picContainer.DrawWidth = 1

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": MoveSelection()"
End Sub


Private Sub DrawGraf(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ptColor As Long
Dim Xm As Long, Ym As Long
On Error GoTo frmErr

Xm = Int(X \ XYspace) + 1
Ym = Int(Y \ XYspace) + 1

If Xm > sCol Then Xm = sCol
If Ym > sRow Then Ym = sRow

Call XYcaptionSet(Xm, Ym)

If InvertMouseB = 1 Then
    Select Case Button
    Case vbLeftButton
        Button = vbRightButton
    Case vbRightButton
        Button = vbLeftButton
    End Select
ElseIf InvertMouseB = 2 Then
    Select Case Button
    Case vbLeftButton
        Button = vbMiddleButton
    Case vbMiddleButton
        Button = vbLeftButton
    End Select
End If

If Button = vbLeftButton Then
    ptColor = lcForeColor
ElseIf Button = vbRightButton Then
    ptColor = lcBackColor
ElseIf Button = vbMiddleButton Then
    ptColor = oldColor
    If oldX = Xm And oldY = Ym Then
'nop
    Else
'        If picContainer.Point((Xm - 1) * XYspace + 1, (Ym - 1) * XYspace + 1) = lcForeColor Then
'            ptColor = lcBackColor
'            oldColor = lcBackColor
'        ElseIf picContainer.Point((Xm - 1) * XYspace + 1, (Ym - 1) * XYspace + 1) = lcBackColor Then
'            ptColor = lcForeColor
'            oldColor = lcForeColor
'        End If
        If GetPixel(picContainer.hdc, (Xm - 1) * XYspace + 1, (Ym - 1) * XYspace + 1) = lcForeColor Then
            ptColor = lcBackColor
            oldColor = lcBackColor
        ElseIf GetPixel(picContainer.hdc, (Xm - 1) * XYspace + 1, (Ym - 1) * XYspace + 1) = lcBackColor Then
            ptColor = lcForeColor
            oldColor = lcForeColor
        End If
        oldX = Xm: oldY = Ym
    End If
Else
    Exit Sub
End If

'Debug.Print ">   DrawGraf"

'Xm = (X \ XYspace) + 1
'Ym = (Y \ XYspace) + 1
If Shift = 1 Then    'large pixel
    If chkGridFlag = vbChecked Then
        picContainer.Line ((Xm - 3) * XYspace + 1, (Ym - 3) * XYspace + 1)- _
                          ((Xm + 2) * XYspace - 1, (Ym + 2) * XYspace - 1), ptColor, BF
'DrawGrid picContainer
    Else
        picContainer.Line ((Xm - 3) * XYspace + 0, (Ym - 3) * XYspace + 0)- _
                          ((Xm + 2) * XYspace - 1, (Ym + 2) * XYspace - 1), ptColor, BF
    End If
Else
    If chkGridFlag = vbChecked Then
        picContainer.Line ((Xm - 1) * XYspace + 1, (Ym - 1) * XYspace + 1)- _
                          (Xm * XYspace - 1, Ym * XYspace - 1), ptColor, BF
    Else
        picContainer.Line ((Xm - 1) * XYspace + 0, (Ym - 1) * XYspace + 0)- _
                          (Xm * XYspace - 1, Ym * XYspace - 1), ptColor, BF
    End If
End If

'If chkGridFlag = vbChecked Then DrawGrid picContainer, vbGrayText, True

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": DrawGraf()"
End Sub
Private Function PicNotEqual(SelInd As Integer) As Boolean
Dim i As Long, j As Long
Dim ptColor As Long
'compare pic from file ImArr and bArr
On Error GoTo frmErr

With tmpPic
    .ScaleMode = vbPixels
    .AutoSize = True
    .Picture = Nothing
    Set .Picture = ImArr(SelInd)
    .Picture = .Image

    If UBound(bArr, 1) < .Height - 1 Or UBound(bArr, 2) < .Width - 1 Then
        PicNotEqual = True
        Exit Function
    End If
End With


For j = 0 To sRowArr(SelInd) - 1
    For i = 0 To sColArr(SelInd) - 1
'ptColor = tmpPic.Point(i, j)
        ptColor = GetPixel(tmpPic.hdc, i, j)
        If ptColor > 0 Then ptColor = 1
        If j > UBound(bArr, 1) Or i > UBound(bArr, 2) Then
            PicNotEqual = True
            Exit Function
        End If
        If ptColor <> bArr(j, i) Then
            PicNotEqual = True
            Exit Function
        End If
    Next i
Next j

'''
Exit Function
frmErr:
MsgBox Err.Description & ": PicNotEqual()"
End Function
' Make sure X1 <= X2 and Y1 <= Y2.
Private Sub OrderCorners()
Dim Tmp As Single

If old_X1 > old_X2 Then
    Tmp = old_X1
    old_X1 = old_X2
    old_X2 = Tmp
End If
If old_Y1 > old_Y2 Then
    Tmp = old_Y1
    old_Y1 = old_Y2
    old_Y2 = Tmp
End If

If X1Region > X2Region Then
    Tmp = X1Region
    X1Region = X2Region
    X2Region = Tmp
End If
If Y1Region > Y2Region Then
    Tmp = Y1Region
    Y1Region = Y2Region
    Y2Region = Tmp
End If


End Sub
Private Sub Copy2PicRectSel()
On Error GoTo frmErr

If chkSelection.Value <> vbChecked Then Exit Sub
If PicReal.Picture = 0 Then Exit Sub
'If picRectSel.Picture = 0 Then Exit Sub

'no And isSelection
'copy pic rect to picRectSel.picture
picRectSel.Cls
Call OrderCorners
picRectSel.Width = old_X2 - old_X1
picRectSel.Height = old_Y2 - old_Y1
'BitBlt picRectSel.hdc, 0, 0, picRectSel.Width, picRectSel.Height, PicReal, picRectSel.Width, picRectSel.Height, SRCCOPY
picRectSel.PaintPicture PicReal.Picture, 0, 0, picRectSel.Width, picRectSel.Height, old_X1, old_Y1, picRectSel.Width, picRectSel.Height    ', vbSrcInvert

picRectSel.Picture = picRectSel.Image

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Copy2PicRectSel()"
End Sub

Private Sub Copy2PicRealFromSelRect()
On Error GoTo frmErr

'If PicReal.Picture = 0 Then Exit Sub
If picRectSel.Picture = 0 Then Exit Sub

Call OrderCorners
PicReal.PaintPicture _
        picRectSel.Picture, _
        old_X1, old_Y1, _
        picRectSel.Width, picRectSel.Height, _
        0, 0, picRectSel.Width, picRectSel.Height, vbSrcCopy    ', vbSrcCopy And vbSrcInvert

PicReal.Picture = PicReal.Image

Call Draw2bArr(PicReal, sCol, sRow)
Call StoreCurrentChar(cmbAdr.ListIndex)
Call bArr2PicDraw(-1)

Call StoreInUndoBuffer

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Copy2PicRealFromSelRect()"
End Sub

Private Sub picContainer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo frmErr
If MouseUpBug Then Exit Sub

If chkSelection.Value = vbChecked Then

    If fCopySelection Or fMoveSelection Then
'copy/move content of rect
        Call Copy2PicRealFromSelRect
    Else
'move rect only
        Call Copy2PicRectSel
    End If

    Exit Sub
End If

'no select rect, Drawing

Call DrawGraf(Button, Shift, X, Y)
oldX = 0: oldY = 0
If chkGridFlag = vbChecked Then DrawGrid picContainer
DoEvents

Call PicDraw2PicReal(-1)    '1 mapping font to real bitmap
Call Draw2bArr(PicReal, sCol, sRow)  '2 write to array bArr
Call StoreCurrentChar(cmbAdr.ListIndex)
Call StoreInUndoBuffer

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": picContainer_MouseUp()"
End Sub


Private Sub bTmp2FontData(ByRef SelInd As Integer)
'and arr
Dim i As Long
On Error GoTo frmErr
'Debug.Print ">   bTmp2FontData"

If Not fFileOpen Then Exit Sub
FontData.reset
For i = 2 To UBound(bTmp)
    FontData.concat dec2binByte(bTmp(i))
Next i
FontDataArr(SelInd) = FontData.Text

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": bTmp2FontData()"
End Sub

Private Sub bTmp2FontData_Vortex(ByRef SelInd As Integer)
'and arr
Dim i As Long
On Error GoTo frmErr
'Debug.Print ">   bTmp2FontData_V"

If Not fFileOpen Then Exit Sub
FontData.reset
For i = 0 To UBound(bTmp)
    FontData.concat StrReverse(dec2binByte(bTmp(i)))
    'FontData.concat dec2binByte(bTmp(i))
Next i
FontDataArr(SelInd) = FontData.Text

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": bTmp2FontData_Vortex"
End Sub

Private Sub picContainerMappingTo(canvas As PictureBox, sColIn As Integer, sRowIn As Integer, Optional sColSt As Integer, Optional sRowSt As Integer)
Dim i As Integer, j As Integer
'Dim gridCorrection As Single
Dim ptColor As Long
On Error GoTo frmErr
'Debug.Print ">   picContainerMappingTo"

'If chkGridFlag = vbChecked Then gridCorrection = 0.5
Dim XYspace_div2 As Integer, sColSt_sColPos As Integer, sRowSt_sRowPos As Integer
XYspace_div2 = XYspace \ 2
sColSt_sColPos = sColSt + sColPos
sRowSt_sRowPos = sRowSt + sRowPos
'no x3 canvas.Visible = False
If DrawWordFlag Then
    For j = 0 To sRowIn - 1
        For i = 0 To sColIn - 1
            If bArr(j, i) = 1 Then
                Call SetPixelV(canvas.hdc, i + sColSt_sColPos, j + sRowSt_sRowPos, lcBackColor)
            End If
        Next i
    Next j
Else
    For j = 0 To sRowIn - 1
        For i = 0 To sColIn - 1
            ptColor = GetPixel(picContainer.hdc, i * XYspace + XYspace_div2, j * XYspace + XYspace_div2)
'If ptColor = lcForeColor Then ptColor = lcBackColor Else ptColor = lcForeColor
            If ptColor = lcForeColor Then
'Call SetPixel(canvas.hdc, i + sColSt_sColPos, j + sRowSt_sRowPos, ptColor)
                Call SetPixelV(canvas.hdc, i + sColSt_sColPos, j + sRowSt_sRowPos, lcBackColor)
            End If
        Next i
    Next j
End If
canvas.Picture = canvas.Image

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": picContainerMappingTo()"
End Sub


Private Sub Draw2bArr(canvas As PictureBox, sColIn As Long, sRowIn As Long)
Dim i As Integer, j As Integer
Dim ptColor As Long
On Error GoTo frmErr
'Debug.Print ">   Draw2bArr"

'fill bArr after manual draw , from 1-1pic

For j = 0 To sRowIn - 1
    For i = 0 To sColIn - 1
        ptColor = GetPixel(canvas.hdc, i, j)
        If Abs(ptColor - vbWhite) <= 1 Then
            bArr(j, i) = 1    'inverse
        Else
            bArr(j, i) = 0
        End If
    Next i
Next j

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": Draw2bArr()"
End Sub


Private Sub DrawGrid(pic As PictureBox, _
                     Optional ByVal LineColor As Long = vbGrayText)
Dim SW As Long, SH As Long
Dim s As Long
Dim i As Long    'floats here since pic may be any size
Dim pt As POINTAPI
Dim redStep As Long

On Error GoTo frmErr

SW = pic.ScaleWidth    '- 1 - XYspace
SH = pic.ScaleHeight    '- 1 - XYspace
s = XYspace    'SW / Cols  'horizontal step
'VS = XYspace    ' SH / Rows   'vertical step
If s = 0 Then Exit Sub

pic.ForeColor = LineColor
redStep = 8 * s

For i = 0 To SW Step s    'draw cols

    MoveToEx pic.hdc, i, 0, pt
    LineTo pic.hdc, i, SH
    If redGridFlag Then If i Mod redStep = 0 Then pic.Line (i, 0)-(i, SH), vbRed
Next i
For i = 0 To SH Step s    'draw rows

    MoveToEx pic.hdc, 0, i, pt
    LineTo pic.hdc, SW, i
    If redGridFlag Then If i Mod redStep = 0 Then pic.Line (0, i)-(SW, i), vbRed
Next i
'If Border Then
'  '  pic.Line (0, 0)-(SW + 0, SH + 0), LineColor, B
'End If
pic.ForeColor = lcForeColor
'''
Exit Sub
frmErr:
MsgBox Err.Description & ": DrawGrid()"
End Sub
Private Sub DrawAllWords()
Dim sChar() As String
Dim i As Integer, j As Integer, n As Integer
Dim s As String
Dim curMaxRow As Integer
Dim curColMaxRow As Integer
Dim curColPos As Integer
Dim allowChangeHeight As Boolean
Dim MaxRow As Integer    'box height
Dim Offset As Integer

On Error GoTo frmErr

If VortexMod Then
    MaxRow = 420
    Offset = 32
Else
    If PicScroll.Width < 350 Then
        MaxRow = 250    '145
    Else
        MaxRow = 145
    End If
    Offset = 1
End If
'CurrentWordInd = cmbVocab.ListIndex

DrawWordFlag = True
DrawAllWordsFlag = True

PicReal.Picture = Nothing

CurrentChar = vbNullString
sRowMax = 0: sColMax = 0

picContainer.Visible = False
Me.MousePointer = vbHourglass

allowChangeHeight = True
curColPos = 0

ReDim AllWordsStartCoord_X(UBound(VocBlock1Arr))
ReDim AllWordsStartCoord_Y(UBound(VocBlock1Arr))
ReDim AllWordsFinishCoord_X(UBound(VocBlock1Arr))
ReDim AllWordsFinishCoord_Y(UBound(VocBlock1Arr))


If Block1Flag Then    'for VocBlock1Arr

    For i = 0 To UBound(VocBlock1Arr)
        s = Trim(VocBlock1Arr(i))
        sChar = Split(s, mySpace)

        AllWordsStartCoord_X(i) = sColPos
        AllWordsStartCoord_Y(i) = sRowPos

        If UBound(sChar) > -1 Then    'else skip empty

            curMaxRow = 0
            For n = 0 To UBound(sChar)

                j = ("&H" & sChar(n)) - Offset

                If j > cmbAdr.ListCount - 1 Then Exit For    'if bug in structure

                CurrentChar = sChar(n)    'for check ShiftChar and so on

                If sRowArr(j) > curMaxRow Then
                    If allowChangeHeight Then sRowMax = sRowMax + sRowArr(j) + 3       'get max height
                    curMaxRow = sRowArr(j)
                End If

                If VortexMod Then
                    Call GetArray_Vortex(j)    'go draw
                Else
                    Call GetArray(j)    'go draw
                End If

            Next n

            If j <= cmbAdr.ListCount - 1 Then
                AllWordsFinishCoord_X(i) = sColPos + sColArr(j)
                AllWordsFinishCoord_Y(i) = sRowPos + sRowArr(j)

                curColMaxRow = curColMaxRow + curMaxRow
                If Not AllWordsInLineFlag And curColMaxRow > MaxRow Then

                    curColMaxRow = 0
                    sColPos = sColMax + 10
                    sRowPos = 0
                    allowChangeHeight = False
                    curColPos = sColPos
                    allowChangeHeight = False
                Else
                    sColPos = curColPos
                    sRowPos = sRowPos + curMaxRow + 3
                End If
            End If

        End If
    Next i



Else    'same for VocBlock2Arr

    For i = 0 To UBound(VocBlock2Arr)
        s = Trim(VocBlock2Arr(i))
        sChar = Split(s, mySpace)
        AllWordsStartCoord_X(i) = sColPos
        AllWordsStartCoord_Y(i) = sRowPos

        If UBound(sChar) > -1 Then    'else skip empty
            curMaxRow = 0
            For n = 0 To UBound(sChar)

                'If VortexMod Then    'not used if vortex
                j = ("&H" & sChar(n)) - Offset

                If j > cmbAdr.ListCount - 1 Then Exit For    'if bug in structure

                CurrentChar = sChar(n)    'for check ShiftChar and so on
                If sRowArr(j) > curMaxRow Then
                    If allowChangeHeight Then sRowMax = sRowMax + sRowArr(j) + 3       'get max height
                    curMaxRow = sRowArr(j)
                End If

                If VortexMod Then
                    Call GetArray_Vortex(j)    'go draw
                Else
                    Call GetArray(j)    'go draw
                End If

            Next n

            If j <= cmbAdr.ListCount - 1 Then
                AllWordsFinishCoord_X(i) = sColPos + sColArr(j)
                AllWordsFinishCoord_Y(i) = sRowPos + sRowArr(j)

                curColMaxRow = curColMaxRow + curMaxRow
                If Not AllWordsInLineFlag And curColMaxRow > MaxRow Then

                    curColMaxRow = 0
                    sColPos = sColMax + 10
                    sRowPos = 0
                    allowChangeHeight = False
                    curColPos = sColPos
                    allowChangeHeight = False
                Else
                    sColPos = curColPos
                    sRowPos = sRowPos + curMaxRow + 3
                End If
            End If

        End If
    Next i

End If

Me.MousePointer = vbNormal
picContainer.Visible = True

'With PicReal                ' 10 word with black frame 5-5
'    lblWordSize = .Width - 10 & "x" & .Height - 10     ' = sRowMax + 10
'End With
AllWordsShow = True
DrawWordFlag = False
DrawAllWordsFlag = False
'return to current char
NoRealDraw = True    'load but no draw 'todo?
startAddr = "&H" & cmbAdr.Text

If cmbAdr.ListIndex > -1 Then
    If VortexMod Then
        Call GetArray_Vortex(cmbAdr.ListIndex)
    Else
        Call GetArray(cmbAdr.ListIndex)
    End If
End If
NoRealDraw = False
sColPos = 0: sRowPos = 0
'sRowMax = 0: sColMax = 0

'''
Exit Sub
frmErr:
picContainer.Visible = True
Me.MousePointer = vbNormal
DrawWordFlag = False
DrawAllWordsFlag = False
NoRealDraw = False
sColPos = 0: sRowPos = 0
'startAddr = "&H" & cmbAdr.Text
MsgBox Err.Description & ": DrawAllWords()"
End Sub
Private Sub DrawWord(s As String)
Dim sChar() As String
Dim i As Integer
Dim j As Integer
On Error GoTo frmErr
'Debug.Print ">   DrawWord"

If Not fFileOpen Then Exit Sub

DrawWordFlag = True
CurrentChar = vbNullString

Do While InStr(s, "  ")
    s = Replace(s, "  ", mySpace)
Loop
s = Trim(s)

sChar = Split(s, mySpace)
'search in list

sRowMax = 0
sColMax = 0

For i = 0 To UBound(sChar)
    '    For j = 0 To cmbAdr.ListCount - 1
    '        If cmbAdr.ItemData(j) = "&H" & sChar(i) Then
    '            CurrentChar = sChar(i)
    '            'draw real char
    '           ' startAddr = "&H" & cmbAdr.List(j)            'Call GetBlock(j)
    '            Call GetArray(j) +vortex
    '            Exit For
    '        End If
    '    Next j

    sRowPos = 0

    CurrentChar = sChar(i)     'for check ShiftChar and so on

    If VortexMod Then
        j = ("&H" & sChar(i)) - 32
        If j > cmbAdr.ListCount - 1 Then Exit For
        Call GetArray_Vortex(j)
    Else
        j = ("&H" & sChar(i)) - 1
        If j > cmbAdr.ListCount - 1 Then Exit For
        Call GetArray(j)
    End If


Next i

With PicReal                ' 10 word with black frame 5-5
    lblWordSize = .Width - 10 & "x" & .Height - 10     ' = sRowMax + 10
End With

DrawWordFlag = False
'return to current char
NoRealDraw = True    'load but no draw current char
'startAddr = "&H" & cmbAdr.Text
'sCol = sColArr(cmbAdr.ListIndex)
'sRow = sRowArr(cmbAdr.ListIndex)

If cmbAdr.ListIndex > -1 Then
    If VortexMod Then
        Call GetArray_Vortex(cmbAdr.ListIndex)
    Else
        Call GetArray(cmbAdr.ListIndex)
    End If


End If

NoRealDraw = False

sColPos = 0: sRowPos = 0

'''
Exit Sub
frmErr:
DrawWordFlag = False
NoRealDraw = False
sColPos = 0
sRowPos = 0
'startAddr = "&H" & cmbAdr.Text
MsgBox Err.Description & ": DrawWord()"
End Sub
Public Function EncryptFW() As Boolean
'true if vortex
'check if decrypted, encrypt, save e*_enc.* and use
Dim FWstring As String    'decripted
Dim arrFirmwareDecr() As Byte
Dim arrFirmwareEncr() As Byte
'Dim bFileIn As Integer
Dim i As Long, n As Long

'Dim NewFileNameFW As String    '*_enc.*
Dim fn As String
Dim sfile As String

On Error GoTo frmErr

If fFileOpen Then    'need
    Close #bFileIn
    fFileOpen = False
End If

bFileIn = FreeFile
If Not OpenFW_read Then Exit Function

lngBytes = LOF(bFileIn)

ReDim arrFirmwareEncr(lngBytes - 1)
ReDim arrFirmwareDecr(lngBytes - 1)
Get #bFileIn, 1, arrFirmwareEncr()
Get #bFileIn, 1, arrFirmwareDecr()
Close #bFileIn

FWstring = String(lngBytes, Chr(0))
Call CopyMemoryString(FWstring, arrFirmwareEncr(0), Len(FWstring))

If InStr(1, FWstring, "VORTEX") Then
    EncryptFW = True    'for this func ok
    'VortexMod = True 'other check will be
    Exit Function
End If

uMagic = &H63B38    'old
For n = 0 To 1
    If InStr(1, FWstring, "Joyetech AP") Then    '"Joyetech APROM"

        'decripted, convert arrFirmwareEncr
        For i = 0 To lngBytes - 1
            arrFirmwareEncr(i) = (arrFirmwareDecr(i) Xor (i + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        Next i

        'save encr file to
        sfile = GetNameExt(FileNameFW)

        sfile = "enc_" & sfile
        LastPath = GetPathFromPathAndName(FileNameFW)
        fn = FWSaveDialog(sfile, ArrMsg(47))
        If Len(fn) <> 0 Then
            If FileExists(fn) Then Kill fn
            FileNameFW = fn

            bFileIn = FreeFile
            If Not OpenFW_write Then Exit Function
            Put #bFileIn, 1, arrFirmwareEncr()
            Close #bFileIn

            EncryptFW = True
            Exit For    'n
        Else
            FileNameFW = vbNullString    'LastOpenedFW
        End If

    Else
        'encripted?, decrypt for check
        For i = 0 To lngBytes - 1
            arrFirmwareDecr(i) = (arrFirmwareEncr(i) Xor (i + lngBytes + uMagic - lngBytes \ uMagic)) And 255
        Next i

        Call CopyMemoryString(FWstring, arrFirmwareDecr(0), Len(FWstring))
        If InStr(1, FWstring, "Joyetech AP", vbTextCompare) Then
            EncryptFW = True
            Exit For    'n
        Else
            'FW not supported
            EncryptFW = False
        End If

    End If

    uMagic = &H3745B6 ' new
Next n

'''
Exit Function
frmErr:
MsgBox Err.Description & ": EncryptFW()"

End Function

Private Sub McListBox1_SelChange(Shift As Integer)
'main click sub
Dim i As Integer, j As Integer, n As Integer
Dim aTmp() As Integer
Dim mli As Long, msi As Long, msc As Long
On Error GoTo frmErr

'Debug.Print "sh= " & Shift

msc = McListBox1.SelCount
mli = McListBox1.ListIndex

If msc = 1 Then
    If mli = cmbLastIndex Then
    
    If VortexMod Then
            Call GetArray_Vortex(cmbLastIndex)    'reclick item to show pic
    Else
        Call GetArray(cmbLastIndex)    'reclick item to show pic
End If

        Call XYcaptionSet(sCol, sRow)
    Else
        cmbAdr.ListIndex = mli
    End If
    
    ReDim selArr(1)
    selArr(1) = mli
    cmdCopy.Caption = ArrMsg(13) & " (1)"
    'Debug.Print McListBox1.SelCount & " = " & mli
    Exit Sub
End If


If Shift = 1 Then
    ReDim selArr(msc)
    If cmbLastIndex <= mli Then

        For i = 1 To msc
            selArr(i) = McListBox1.SelItem(i - 1)
        Next i

    Else

        For i = 1 To msc
            selArr(msc - i + 1) = McListBox1.SelItem(i - 1)
        Next i
    End If

    'Debug.Print "sh add ";: For i = 1 To UBound(selArr): Debug.Print selArr(i);: Next i



''' need NOT fast scroll select (((

ElseIf msc > UBound(selArr) Then

    'If UBound(selArr) <= 1 Then
    '    ReDim selArr(msc)
    'Else
    ReDim Preserve selArr(msc)
    'End If

    ''Debug.Print McListBox1.SelItem(McListBox1.ListIndex)
    selArr(msc) = mli
    'Debug.Print "add.. " & msc & " = " & mli

End If

''''''''' --- '''''''''
If msc < UBound(selArr) Then
    'rem item McListBox1.ListIndex

    'Debug.Print "was ";
    ReDim aTmp(UBound(selArr))
    For i = 1 To UBound(selArr)
        aTmp(i) = selArr(i)
        'Debug.Print aTmp(i);
    Next i

    'Debug.Print "now ";

    ReDim selArr(msc)

    For i = 1 To UBound(aTmp)
        For j = 0 To msc - 1

            msi = McListBox1.SelItem(j)    '+ 1
            If aTmp(i) = msi Then
                n = n + 1
                selArr(n) = aTmp(i)
                'Debug.Print selArr(n);
                Exit For
            End If

        Next j
    Next i
End If


cmdCopy.Caption = ArrMsg(13) & " (" & msc & ")"

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": McListBox1_SelChange()"
End Sub
Private Sub HScroll_X_Change()
On Error GoTo frmErr

If Not LoadFontFlag Then Call HScroll_X_Scroll

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": HScroll_X_Change()"
End Sub

Private Sub HScroll_X_Scroll()
On Error GoTo frmErr

TTF_X = HScroll_X.Value
TTFontDraw TTF_Char, TTF_Size, TTF_X, TTF_Y, TTFontBold, TTFontItalic, TTFontUnderline

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": HScroll_X_Scroll"
End Sub



Private Sub VScroll_Y_Change()
On Error GoTo frmErr

If Not LoadFontFlag Then VScroll_Y_Scroll

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": VScroll_Y_Change"
End Sub

Private Sub VScroll_Y_Scroll()
On Error GoTo frmErr

TTF_Y = VScroll_Y.Value
TTFontDraw TTF_Char, TTF_Size, TTF_X, TTF_Y, TTFontBold, TTFontItalic, TTFontUnderline

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": VScroll_Y_Scroll"
End Sub
Private Sub VScroll_S_Change()
On Error GoTo frmErr

If Not LoadFontFlag Then VScroll_S_Scroll

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": VScroll_S_Change"
End Sub

Private Sub VScroll_S_Scroll()
On Error GoTo frmErr

TTF_Size = VScroll_S.Value
TTFontDraw TTF_Char, TTF_Size, TTF_X, TTF_Y, TTFontBold, TTFontItalic, TTFontUnderline

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": VScroll_S_Scroll"
End Sub

Private Sub VScrollDraw_Change()
Call VScrollDraw_Scroll
End Sub

Private Sub VScrollDraw_Scroll()
picContainer.top = VScrollDraw.Value
End Sub
Private Sub StoreInUndoBuffer()
On Error GoTo frmErr

'PicReal.Picture = PicReal.Image

If picCount < MaxUndoCircle Then
    picCount = picCount + 1
Else
    picCount = 0
End If

If UBound(UndoBuffer) < MaxUndoCircle Then
    ReDim Preserve UndoBuffer(picCount)
End If

If fFileOpen Then
    If PicReal.Width < 257 And PicReal.Height < 257 Then
        Set UndoBuffer(picCount) = PicReal.Picture
    End If
Else
    Set UndoBuffer(picCount) = PicReal.Picture
End If

'Debug.Print "prev to " & picCount
UndoClicksCount = 0

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": StoreInUndoBuffer"
End Sub

Private Sub UndoBufferClear()
ReDim UndoBuffer(0)
picCount = -1
UndoClicksCount = 0

End Sub

Private Sub XML_Fill_Arrays(oElement As CXmlElement)
'Dim i As Integer
Dim oChild As CXmlElement
On Error GoTo frmErr

Select Case oElement.Name
Case "Image"
    XML_Image_Count = XML_Image_Count + 1
    ReDim Preserve arrXML_ImageNum(XML_Image_Count)
    ReDim Preserve arrXML_ImageCol(XML_Image_Count)
    ReDim Preserve arrXML_ImageRow(XML_Image_Count)
'if For i = 1 To oElement.AttributeCount
    arrXML_ImageNum(XML_Image_Count) = "&H" & oElement.ElementAttribute(1).Value
    arrXML_ImageCol(XML_Image_Count) = oElement.ElementAttribute(2).Value
    arrXML_ImageRow(XML_Image_Count) = oElement.ElementAttribute(3).Value

'Next i

Case "Data"
'after image must be
    ReDim Preserve arrXML_DataBody(XML_Image_Count)
    arrXML_DataBody(XML_Image_Count) = oElement.Body

End Select

'Debug.Print oElement.Name,
'Debug.Print oElement.Body 'chars

'For i = 1 To oElement.AttributeCount
'Debug.Print oElement.ElementAttribute(i).Value,
'Next i

For Each oChild In oElement
    Call XML_Fill_Arrays(oChild)
Next

'''
Exit Sub
frmErr:
MsgBox Err.Description & ": XML_Fill_Arrays"
End Sub

VERSION 5.00
Begin VB.UserControl McListBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   FillColor       =   &H80000018&
   FillStyle       =   0  'Solid
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   147
End
Attribute VB_Name = "McListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------------
' SourceCode : McListBox 1.7
' Auther     : Jim Jose
' Email      : jimjosev33@yahoo.com
' Date       : 3-9-2005
' Purpose    : An upgraded version of VBListBox with Icons, Item HighLight and many more
' CopyRight  : JimJose © Gtech Creations - 2005
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
' About :
'        This control is a replacement for the vb's inbuilt 'ListBox'
' control. This version comes with,

'   1. Unicode' support
'   2. MultiSelect Option
'   3. Custom Icons for each item
'   4. Item highlight by making the fond bold
'   5. Nice gradient effects
'   6. Custom RowHeight
'   7. SortItems(Ascendindg/Descending)
'   8. Grid Lines
'   9. API SCrollBars
'  10. XP style
'  11. Mouse Wheel
'  12. And many more...
'
'       This is my second attempt on this purpose. The first one
' 'ListBoxEx' is there on PSC. The primary objective of this update
' was multiple icons and Item HighLight. I tried to implement most
' of the suggessions I got on previous submission.
'
'       But, this version also will not support 'Multiple Columns' and
' 'Column headers'. I agree.. that will not cause much effort. But, this
' control is ment for easy use in small & nice projects. So I like to
' keep it small!!
'
'       This control also uses 'McImageList', by which the images can store
' externally and make referance to them. This way we can greatly reduse the
' memory use caused by a big 'Picture Array', where each item needs a seperate
' picture object.
'
' Special Note:
'       Please 'Refresh' the listbox after 'item adding'. I didn't include
' 'Sorting' and 'Redrawing' on 'ItemAdd' for better speed..

'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
'
' Credits/Thanks :
'
'    Paul Carton    -   For his unbeatable self subclasser.
'    Gary Noble     -   For his API ScrollBars and ColorBlending code.
'    Carls P.V      -   For his excellent DIB-gradient routine.
'    Dana Seaman    -   The Master of Unicode support.
'
'-----------------------------------------------------------------------------------------------------------
'Bug fixes by Sur

Option Explicit

'[APIs]
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lprect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lprect As RECT, ByVal wFormat As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lprect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetRect Lib "user32.dll" (ByRef lprect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
'Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' for subclassing
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

' for API scroll bars
'Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
'Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
'Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
'Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
'Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, ByVal Index As Long, ByVal newValue As Long, ByVal fRedraw As Boolean) As Long
'Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long

'Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


'[Types]
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128    ' Maintenance string
End Type

Private Type tSubData                                                                   'Subclass data type
    hWnd As Long                                   'Handle of the window being subclassed
    nAddrSub As Long                                       'The address of our new WndProc (allocated memory).
    nAddrOrig As Long                                        'The address of the pre-existing WndProc
    nMsgCntA As Long                                       'Msg after table entry count
    nMsgCntB As Long                                       'Msg before table entry count
    aMsgTblA() As Long                                         'Msg after table array
    aMsgTblB() As Long                                         'Msg Before table array
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As TRACKMOUSEEVENT_FLAGS
    hwndTrack As Long
    dwHoverTime As Long
End Type

' Scroll bar
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

'[Enums]
Public Enum GradientDirectionEnum
    [Fill_None] = 0
    [Fill_Horizontal] = 1
    [Fill_HorizontalMiddleOut] = 2
    [Fill_Vertical] = 3
    [Fill_VerticalMiddleOut] = 4
    [Fill_DownwardDiagonal] = 5
    [Fill_UpwardDiagonal] = 6
End Enum

Public Enum AppearanceEnum
    [Flat] = 0
    [3D] = 1
End Enum

Public Enum BorderEnum
    [None] = 0
    [Fixed Single] = 1
End Enum

Public Enum SortOrderEnum
    [Sort_None] = 0
    [Sort_Ascending] = -1
    [Sort_Desending] = 1
End Enum

Public Enum SelectionStyleEnum
    [Style_Normal] = 0
    [Style_XP] = 1
End Enum

' for subclassing
Private Enum eMsgWhen
    MSG_AFTER = 1       'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

' for scroll bars
Public Enum ScrollBarOrienationEnum
    Scroll_Horizontal
    Scroll_Vertical
    Scroll_Both
End Enum

Public Enum ScrollBarStyleEnum
    Style_Regular = 1&    ' FSB_REGULAR_MODE
    Style_Flat = 0&    'FSB_FLAT_MODE
End Enum

Public Enum EFSScrollBarConstants
    efsHorizontal = 0    'SB_HORZ
    efsVertical = 1    'SB_VERT
End Enum

'[Local Variables]
Private m_SelItem As Long
Private m_RowHeight As Long    'Double '
Private m_iCount As Long
Private m_hMode As Long
Private m_KeyControl As Boolean
Private m_bIsNT As Boolean
Private m_Inside As Boolean
Private m_HasFocus As Boolean
Private m_PicComplete As PictureBox
Private m_Left As Long  'Double
Private m_LastComplete As Long

'[Data Storage]
Private m_Items As New Collection
Private m_ImageList As Object
Private m_Selected As New Collection

'[Property Variables]
Private m_Picture As New StdPicture
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_Font As New StdFont
Private m_SelColor As OLE_COLOR
Private m_FullRowSel As Boolean
Private m_SortOrder As SortOrderEnum
Private m_SelForeColor As OLE_COLOR
Private m_StrechIcon As Boolean
Private m_IconFocus As Boolean
Private m_GridLines As Boolean
Private m_GridColor As OLE_COLOR
Private m_BackGradient As GradientDirectionEnum
Private m_SelGradient As GradientDirectionEnum
Private m_MultiSelect As Boolean
Private m_SelStart As Long
Private m_BackGradientCol As OLE_COLOR
Private m_SelGradientCol As OLE_COLOR
Private m_TextAllineMent As AlignmentConstants
Private m_ShowIcon As Boolean
Private m_FocusRectangle As Boolean
Private m_SelectionStyle As SelectionStyleEnum
Private m_FlatScrollBar As Boolean
Private m_AutoHideScrollBars As Boolean

'[Referance Constants]
Private Const DIB_RGB_ColS As Long = 0
Private Const VER_PLATFORM_WIN32_NT As Long = 2
Private Const SPLITER As String = "<%S%>"

'[Default Property Values]
Private Const m_def_BackColor = &HFFFFFF
Private Const m_def_ForeColor = &H80000012
Private Const m_def_SelColor = &HFF8C1A
Private Const m_def_SelForeColor = &HFFFFFF
Private Const m_def_StrechIcon = True
Private Const m_def_Appearance = 1
Private Const m_def_BorderStyle = 1
Private Const m_def_FullRowSel = True
Private Const m_def_SortOrder = Sort_None
'Private Const m_def_SortItems = False
Private Const m_def_IconFocus = False
Private Const m_def_TextAllignMent = vbLeftJustify
Private Const m_Def_RowHeight = 26
Private Const m_def_GridLines = True
Private Const m_Def_GridColor = &H8000000F
Private Const m_def_BackGradient = 0
Private Const m_def_SelGradient = 0
Private Const m_def_MultiSelect = True
'Private Const m_def_SelCount = 0
Private Const m_def_BackGradientCol = &HC0C0C0
Private Const m_def_SelGradientCol = &HFFFFFF
Private Const m_def_ShowIcon = True
Private Const m_def_FocusRectangle = False
Private Const m_def_SelectionStyle = Style_XP
Private Const m_def_FlatScrollBar = False
Private Const m_def_AutoHideScrollBars = True
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = &H80&
'Private Const SWP_SHOWWINDOW As Long = &H40

' for subclassing
'Private Const WM_GETMINMAXINFO      As Long = &H24
'Private Const WM_WINDOWPOSCHANGED   As Long = &H47
'Private Const WM_WINDOWPOSCHANGING  As Long = &H46
'Private Const WM_LBUTTONDOWN        As Long = &H201
'Private Const WM_SIZE               As Long = &H5
'Private Const WM_LBUTTONDBLCLK      As Long = &H203
'Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_MOUSEMOVE As Long = &H200
'Private Const WM_SETFOCUS           As Long = &H7
'Private Const WM_KILLFOCUS          As Long = &H8
'Private Const WM_MOVE               As Long = &H3
'Private Const WM_TIMER              As Long = &H113
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEHOVER As Long = &H2A1

Private Const ALL_MESSAGES As Long = -1                             'All messages added or deleted
Private Const GMEM_FIXED As Long = 0                            'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC As Long = -4                            'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04 As Long = 88                         'Table B (before) address patch offset
Private Const PATCH_05 As Long = 93                         'Table B (before) entry count patch offset
Private Const PATCH_08 As Long = 132                        'Table A (after) address patch offset
Private Const PATCH_09 As Long = 137                        'Table A (after) entry count patch offset

Private sc_aSubData() As tSubData
Private bTrack As Boolean
Private bTrackUser32 As Boolean
Private bInCtrl As Boolean

' for scroll bars
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114

Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
Private Const MK_CONTROL = &H8
'Private Const WSB_PROP_VSTYLE = &H100&
'Private Const WSB_PROP_HSTYLE = &H200&
'Private Const FSB_FLAT_MODE = 1&
'Private Const FSB_REGULAR_MODE = 0&

Private m_bInitialised As Boolean
Private m_eOrientation As ScrollBarOrienationEnum
Private m_eStyle As ScrollBarStyleEnum
Private m_hWnd As Long
Private m_lSmallChangeHorz As Long
Private m_lSmallChangeVert As Long
Private m_bEnabledHorz As Boolean
Private m_bEnabledVert As Boolean
Private m_bVisibleHorz As Boolean
Private m_bVisibleVert As Boolean
Private m_bNoFlatScrollBars As Boolean

Private iPrev As Long

'[Events]
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseClick()
Public Event SelChange(Shift As Integer)
Public Event DbClick()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event mouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)



'[ Subclassed events receiver ]
'------------------------------------------------------------------------------------------
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef LParam As Long)
Dim eBar As EFSScrollBarConstants
Dim lV As Long, lSC As Long
Dim lScrollCode As Long
Dim tSI As SCROLLINFO
Dim mPrev As Long

mPrev = SBValue(efsVertical)

' Gray Nobles's API scrollbar code
Select Case uMsg
Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
    lScrollCode = (wParam And &HFFFF&)

    Select Case uMsg

    Case WM_HSCROLL    ' Get the scrollbar type
        eBar = efsHorizontal

    Case WM_VSCROLL
        eBar = efsVertical

    Case Else     'WM_MOUSEWHEEL
        eBar = IIf(lScrollCode And MK_CONTROL, efsHorizontal, efsVertical)
        lScrollCode = IIf(wParam / 65536 < 0, SB_LINEDOWN, SB_LINEUP)

    End Select

    Select Case lScrollCode

    Case SB_THUMBTRACK

' Is vertical/horizontal?
        pSBGetSI eBar, tSI, SIF_TRACKPOS
        SBValue(eBar) = tSI.nTrackPos

    Case SB_LEFT, SB_BOTTOM
        SBValue(eBar) = IIf(lScrollCode = 7, SBMax(eBar), SBMin(eBar))

    Case SB_RIGHT, SB_TOP
        SBValue(eBar) = SBMin(eBar)

    Case SB_LINELEFT, SB_LINEUP

        If SBVisible(eBar) Then

            lV = SBValue(eBar)
            If (eBar = efsHorizontal) Then
                lSC = m_lSmallChangeHorz
            Else
                lSC = m_lSmallChangeVert
            End If

            If (lV - lSC < SBMin(eBar)) Then
                SBValue(eBar) = SBMin(eBar)
            Else
                SBValue(eBar) = lV - lSC
            End If

        End If

    Case SB_LINERIGHT, SB_LINEDOWN
        If SBVisible(eBar) Then

            lV = SBValue(eBar)

            If (eBar = efsHorizontal) Then
                lSC = m_lSmallChangeHorz
            Else
                lSC = m_lSmallChangeVert
            End If

            If (lV + lSC > SBMax(eBar)) Then
                SBValue(eBar) = SBMax(eBar)
            Else
                SBValue(eBar) = lV + lSC
            End If
        End If

    Case SB_PAGELEFT, SB_PAGEUP
        SBValue(eBar) = SBValue(eBar) - SBLargeChange(eBar)

    Case SB_PAGERIGHT, SB_PAGEDOWN
        SBValue(eBar) = SBValue(eBar) + SBLargeChange(eBar)

    Case SB_ENDSCROLL

    End Select

    If Not mPrev = SBValue(efsVertical) Then ReDrawList

End Select

Select Case uMsg
Case WM_MOUSEWHEEL
    m_PicComplete.Visible = False
    m_LastComplete = -100

Case WM_MOUSEMOVE
    If Not bInCtrl Then
        bInCtrl = True
        Call TrackMouseLeave(lng_hWnd)
        RaiseEvent MouseEnter
    End If

Case WM_MOUSELEAVE
    bInCtrl = False
    m_PicComplete.Visible = False
    m_LastComplete = -100
    RaiseEvent MouseLeave

End Select

End Sub


'------------------------------------------------------------------------------------------
'Domain     : UserControl Properties
'------------------------------------------------------------------------------------------

Public Function ListCount() As Long
On Error GoTo handle
ListCount = m_Items.count

Exit Function
handle:
ListCount = 0
End Function


Public Property Get ListIcon(ByVal vIndex As Long) As Long
On Error Resume Next

ListIcon = Split(m_Items(vIndex + 1), SPLITER)(1)
End Property

Public Property Let ListIcon(ByVal vIndex As Long, ByVal vNewPicture As Long)
Dim TxtData() As String
On Error Resume Next

TxtData() = Split(m_Items(vIndex + 1), SPLITER)
TxtData(1) = vNewPicture
m_Items.Remove vIndex + 1
m_Items.Add Join(TxtData, SPLITER), , vIndex + 1
PropertyChanged "ListIcon"
ReDrawList

End Property


Public Property Get ListBold(ByVal vIndex As Long) As Boolean
Dim mBold As String
On Error Resume Next

mBold = Split(m_Items(vIndex + 1), SPLITER)(2)
If mBold = "True" Then
    ListBold = True
Else
    ListBold = False
End If

End Property

Public Property Let ListBold(ByVal vIndex As Long, ByVal vNewVlaue As Boolean)
Dim TxtData() As String
On Error Resume Next

TxtData = Split(m_Items(vIndex + 1), SPLITER)
TxtData(2) = vNewVlaue
m_Items.Remove vIndex + 1

If vIndex + 1 > m_Items.count Then
    m_Items.Add Join(TxtData, SPLITER)
Else
    m_Items.Add Join(TxtData, SPLITER), , vIndex + 1
End If

PropertyChanged "ListBold"
ReDrawList

handle:
End Property

Public Property Get Picture() As Picture
Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal vNewPicture As Picture)
Set m_Picture = vNewPicture
PropertyChanged "Picture"
ReDrawList
End Property


Public Property Get Font() As Font
Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
Set m_Font = vNewFont
UserControl.Font = vNewFont
If TextHeight("A") > m_RowHeight Then m_RowHeight = TextHeight("A")
PropertyChanged "Font"
Me.Refresh
End Property


Public Property Get ForeColor() As OLE_COLOR
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewCol As OLE_COLOR)
m_ForeColor = vNewCol
PropertyChanged "ForeColor"
ReDrawList
End Property


Public Property Get BackColor() As OLE_COLOR
BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal vNewCol As OLE_COLOR)
m_BackColor = vNewCol
PropertyChanged "BackColor"
UserControl.BackColor = vNewCol
ReDrawList
End Property


Public Property Get SelColor() As OLE_COLOR
SelColor = m_SelColor
End Property

Public Property Let SelColor(ByVal vNewCol As OLE_COLOR)
m_SelColor = vNewCol
PropertyChanged "SelColor"
ReDrawList
End Property


Public Property Get SelForeColor() As OLE_COLOR
SelForeColor = m_SelForeColor
End Property

Public Property Let SelForeColor(ByVal vNewCol As OLE_COLOR)
m_SelForeColor = vNewCol
PropertyChanged "SelForeColor"
ReDrawList
End Property


Public Property Get StrechIcon() As Boolean
StrechIcon = m_StrechIcon
End Property

Public Property Let StrechIcon(ByVal vNewValue As Boolean)
m_StrechIcon = vNewValue
PropertyChanged "StrechIcon"
ReDrawList
End Property


Public Property Get Appearance() As AppearanceEnum
Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal vNewAppearance As AppearanceEnum)
UserControl.Appearance = vNewAppearance
PropertyChanged "Appearance"
ReDrawList
End Property


Public Property Get BorderStyle() As BorderEnum
BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewBorder As BorderEnum)
UserControl.BorderStyle = vNewBorder
PropertyChanged "BorderStyle"
ReDrawList
End Property


Public Property Get ListIndex() As Long
If m_Items.count = 0 Then
    ListIndex = -1
Else
    ListIndex = m_SelItem
End If
End Property

Public Property Let ListIndex(ByVal vNewValue As Long)
'Dim i As Integer
m_SelItem = vNewValue

If SBVisible(efsVertical) Then
    If m_SelItem < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, m_SelItem + 2 - ScaleHeight \ m_RowHeight)  'm_SelItem)
    If m_SelItem >= SBValue(efsVertical) + ScaleHeight \ m_RowHeight Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), m_SelItem - 1)
End If
'Debug.Print ScaleHeight \ m_RowHeight

Selection_Clear
m_Selected.Add m_SelItem

PropertyChanged "ListIndex"
CheckSelected
ReDrawList
RaiseEvent SelChange(0)

End Property


Public Property Get Text() As String
If ListCount = 0 Then Exit Property
Text = Split(m_Items(m_SelItem + 1), SPLITER)(0)
End Property


Public Property Get List(ByVal vIndex As Long) As String

If vIndex > ListCount - 1 Or vIndex < 0 Then Exit Sub
List = Split(m_Items(vIndex + 1), SPLITER)(0)

End Property

Public Property Let List(ByVal vIndex As Long, ByVal vNewValue As String)
Dim TxtData() As String

TxtData = Split(m_Items(vIndex + 1), SPLITER)
TxtData(0) = vNewValue
m_Items(vIndex + 1) = Join(TxtData, SPLITER)

End Property


Public Property Get FullRowSelect() As Boolean
FullRowSelect = m_FullRowSel
End Property

Public Property Let FullRowSelect(ByVal vNewValue As Boolean)
m_FullRowSel = vNewValue
PropertyChanged "FullRowSelect"
ReDrawList
End Property


Public Property Get MultiSelect() As Boolean
MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)

m_MultiSelect = New_MultiSelect
Selection_Clear
m_Selected.Add m_SelItem

PropertyChanged "MultiSelect"
ReDrawList

End Property


Public Property Get SelCount() As Long
SelCount = m_Selected.count
End Property

Public Property Set ImageList(ByVal vNewValue As Object)
Set m_ImageList = vNewValue
End Property


Public Property Get FocusRectangle() As Boolean
FocusRectangle = m_FocusRectangle
End Property

Public Property Let FocusRectangle(ByVal New_FocusRectangle As Boolean)
m_FocusRectangle = New_FocusRectangle
PropertyChanged "FocusRectangle"
ReDrawList
End Property


Public Property Get ShowIcon() As Boolean
ShowIcon = m_ShowIcon
End Property

Public Property Let ShowIcon(ByVal New_ShowIcon As Boolean)
m_ShowIcon = New_ShowIcon
PropertyChanged "ShowIcon"
ReDrawList
End Property


Public Property Get SortOrder() As SortOrderEnum
SortOrder = m_SortOrder
End Property

Public Property Let SortOrder(ByVal vNewValue As SortOrderEnum)
m_SortOrder = vNewValue
PropertyChanged "SortOrder"
SortCollection m_Items, m_SortOrder
ReDrawList
End Property


Public Property Get IconFocus() As Boolean
IconFocus = m_IconFocus
End Property

Public Property Let IconFocus(ByVal vNewValue As Boolean)
m_IconFocus = vNewValue
PropertyChanged "IconFocus"
ReDrawList
End Property


Public Property Get TextAlignment() As AlignmentConstants
TextAlignment = m_TextAllineMent
End Property

Public Property Let TextAlignment(ByVal vNewValue As AlignmentConstants)
m_TextAllineMent = vNewValue
PropertyChanged "TextAlignment"
ReDrawList
End Property


Public Property Get GridLines() As Boolean
GridLines = m_GridLines
End Property

Public Property Let GridLines(ByVal vNewValue As Boolean)
m_GridLines = vNewValue
PropertyChanged "GridLines"
ReDrawList
End Property


Public Property Get GridColor() As OLE_COLOR
GridColor = m_GridColor
End Property

Public Property Let GridColor(ByVal vNewValue As OLE_COLOR)
m_GridColor = vNewValue
PropertyChanged "GridColor"
ReDrawList
End Property


Public Property Get RowHeight() As Long
RowHeight = m_RowHeight
End Property

Public Property Let RowHeight(ByVal vNewValue As Long)
If vNewValue >= TextHeight("A") Then
    m_RowHeight = vNewValue
Else
    m_RowHeight = TextHeight("A")
End If
PropertyChanged "RowHeight"
Me.Refresh
End Property


Public Property Get BackGradient() As GradientDirectionEnum
BackGradient = m_BackGradient
End Property

Public Property Let BackGradient(ByVal New_BackGradient As GradientDirectionEnum)
m_BackGradient = New_BackGradient
PropertyChanged "BackGradient"
ReDrawList
End Property


Public Property Get SelGradient() As GradientDirectionEnum
SelGradient = m_SelGradient
End Property

Public Property Let SelGradient(ByVal New_SelGradient As GradientDirectionEnum)
m_SelGradient = New_SelGradient
PropertyChanged "SelGradient"
ReDrawList
End Property

Public Property Get BackGradientCol() As OLE_COLOR
BackGradientCol = m_BackGradientCol
End Property

Public Property Let BackGradientCol(ByVal New_BackGradientCol As OLE_COLOR)
m_BackGradientCol = New_BackGradientCol
PropertyChanged "BackGradientCol"
ReDrawList
End Property


Public Property Get SelGradientCol() As OLE_COLOR
SelGradientCol = m_SelGradientCol
End Property

Public Property Let SelGradientCol(ByVal New_SelGradientCol As OLE_COLOR)
m_SelGradientCol = New_SelGradientCol
PropertyChanged "SelGradientCol"
ReDrawList
End Property


Public Property Get SelItem(ByVal vSelIndex As Long) As Long
SelItem = m_Selected(vSelIndex + 1)
End Property


Public Property Get SelectionStyle() As SelectionStyleEnum
SelectionStyle = m_SelectionStyle
End Property

Public Property Let SelectionStyle(ByVal vNewValue As SelectionStyleEnum)
m_SelectionStyle = vNewValue
PropertyChanged "SelectionStyle"
ReDrawList
End Property


Public Property Get FlatScrollBar() As Boolean
FlatScrollBar = m_FlatScrollBar
End Property

Public Property Let FlatScrollBar(ByVal vNewValue As Boolean)

m_FlatScrollBar = vNewValue
pSBClearUp
m_bNoFlatScrollBars = Not m_FlatScrollBar
If m_FlatScrollBar Then SBCreate UserControl.hWnd
PropertyChanged "FlatScrollBar"
DoEvents
Me.Refresh

End Property


Public Property Get AutoHideScrollBars() As Boolean
AutoHideScrollBars = m_AutoHideScrollBars
End Property

Public Property Let AutoHideScrollBars(ByVal New_AutoHideScrollBars As Boolean)
m_AutoHideScrollBars = New_AutoHideScrollBars
PropertyChanged "AutoHideScrollBars"
Me.Refresh
End Property


' By 'Gary Noble'
Public Sub ClearSelection()

Selection_Clear
m_Selected.Add m_SelItem
ReDrawList

End Sub
Public Sub ClearSelectionAll()

Selection_Clear
'm_Selected.Add m_SelItem

'ReDrawList True 'False

End Sub

'' By 'Gary Noble'
Public Sub SelectAll()
On Error Resume Next
Dim X As Long
Selection_Clear
For X = 0 To m_Items.count - 1
    m_Selected.Add X
Next
ReDrawList

On Error GoTo 0
End Sub

Public Sub SelectItem(Ind As Integer)
'need refresh
On Error Resume Next
Dim X As Long
'Selection_Clear
m_SelItem = Ind

For X = 0 To m_Items.count
    If X = Ind Then
        m_Selected.Add X
        Exit For
    End If
Next
If SBVisible(efsVertical) Then
    'Debug.Print SBValue(efsVertical)
    'no SBValue(efsVertical) = old_SBValue_efsVertical
    
    If reloadFW_flag Then
    'only if reload current
        SBValue(efsVertical) = old_SBValue_efsVertical
    Else
        If m_SelItem < SBValue(efsVertical) Then
            'scroll up in cmbAdr
            SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, old_SBValue_efsVertical, m_SelItem + 2 - ScaleHeight \ m_RowHeight)  'm_SelItem)
'            Debug.Print 1
        End If
        If m_SelItem >= SBValue(efsVertical) + ScaleHeight \ m_RowHeight Then
            'scroll down in cmbAdr
            SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), m_SelItem - 1)
'            Debug.Print 2
        End If

    End If
End If
old_SBValue_efsVertical = SBValue(efsVertical)

CheckSelected

'after all in frmmain > ReDrawList

On Error GoTo 0
End Sub
Public Sub ViewBoldItem(Ind As Integer)
'sur scroll to ind without sel
'need refresh to list
On Error Resume Next

If Ind < 0 Then Exit Sub

If SBVisible(efsVertical) Then
    If Ind < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, Ind + 2 - ScaleHeight \ m_RowHeight)
    If Ind >= SBValue(efsVertical) + ScaleHeight \ m_RowHeight Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), Ind - 1)
End If

On Error GoTo 0
End Sub


'------------------------------------------------------------------------------------------
' Domain    : Events
'------------------------------------------------------------------------------------------

Private Sub UserControl_Click()
RaiseEvent MouseClick
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DbClick
End Sub

Private Sub UserControl_GotFocus()
m_HasFocus = True
ReDrawList    'sur
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mNew As Long
'Debug.Print "KeyDown shift: " & Shift

' Select each Key
Select Case KeyCode
Case vbKeyUp
    mNew = m_SelItem - 1

Case vbKeyDown
    mNew = m_SelItem + 1

Case vbKeyEnd
    mNew = ListCount - 1

Case vbKeyHome
    mNew = 0

Case vbKeyPageDown
    mNew = m_SelItem + m_iCount

Case vbKeyPageUp
    mNew = m_SelItem - m_iCount

'Case 65 'A
'If Shift = 2 Then RaiseEvent KeyDown(KeyCode, Shift)

Case Else
    RaiseEvent KeyDown(KeyCode, Shift)
    Exit Sub
End Select

If mNew > ListCount - 1 Then mNew = ListCount - 1
If mNew < 0 Then mNew = 0

' Refrech Control
If Not mNew = m_SelItem And Not mNew = -1 Then

    If SBVisible(efsVertical) Then
        If mNew < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, mNew)
        If mNew >= SBValue(efsVertical) + ScaleHeight \ m_RowHeight Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), mNew - m_iCount + 1)
    End If

    UserControl_MouseDown 999, Shift, Val(mNew), 0

End If

End Sub

Private Sub UserControl_LostFocus()
m_HasFocus = False
ReDrawList
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mRowHeight As Long    'Double
Dim xStart As Long
Dim mIndex As Long
Dim xEnd As Long
Dim X1 As Long
'Debug.Print "MouseDown shift: " & Shift
'If MouseUpBug Then Exit Sub

If m_SelItem > ListCount - 1 Or m_SelItem < 0 Then Exit Sub

'Debug.Print Button '"Selecting Item!"
If Button = 999 Then    ' Event send from KEyDown
    m_SelItem = X
ElseIf Button = vbRightButton Then
    GoTo Skip
Else
    mRowHeight = m_Def_RowHeight  '(ScaleHeight \ m_iCount) 'sur

    If Y \ mRowHeight > m_Items.count Then Exit Sub

    m_SelItem = SBValue(efsVertical) + Y \ mRowHeight
End If
If m_SelItem > m_Items.count - 1 Then m_SelItem = m_Items.count - 1    'sur from unselect wrong last

If GetKeyState(vbKeyControl) < 0 And Button = vbKeyLButton And m_MultiSelect Then
    mIndex = Selection_Find(m_SelItem)
    If mIndex = 0 Then
        m_Selected.Add m_SelItem
        Selection_Sort
    Else
        m_Selected.Remove mIndex
    End If
    m_SelStart = m_SelItem
    GoTo Skip
End If

If GetKeyState(vbKeyShift) < 0 And m_MultiSelect Then

    If m_SelStart < m_SelItem Then
        xStart = m_SelStart
        xEnd = m_SelItem
    Else
        xStart = m_SelItem
        xEnd = m_SelStart
    End If

    Selection_Clear

    For X1 = xStart To xEnd
        m_Selected.Add X1
    Next X1

    Selection_Sort

Else
    m_SelStart = m_SelItem
    Selection_Clear
    m_Selected.Add m_SelItem    'sur sur
'Debug.Print m_Selected.Count
End If


Skip:
'DoEvents

'If Button = 999 Then
ReDrawList True '1 redraw on click Need for shift click select
'End If

iPrev = m_SelItem

RaiseEvent SelChange(Shift)
RaiseEvent mouseDown(Button, Shift, X, Y)

'Debug.Print "upmd " & m_Selected.Count
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mRowHeight As Long    'Double
'Static mPrev As Long
Dim xStart As Long
Dim mIndex As Long
Dim xEnd As Long
Dim X1 As Long
Dim SBValue_efsVertical As Long

'Debug.Print "M mouse move" & Time
If MouseUpBug Then Exit Sub

If Y < 0 Or Y > ScaleHeight Then m_Inside = False Else m_Inside = True
mRowHeight = m_Def_RowHeight    '(ScaleHeight \ m_iCount) sur
mIndex = Y \ mRowHeight
'ShowItemComplete mIndex + 1 'sur?

' In multiselection mode, user is trying to select more than one row
' by (LeftButton Down + MouseMove)
If GetKeyState(vbKeyLButton) < 0 And m_MultiSelect And GetKeyState(vbKeyControl) >= 0 And GetKeyState(vbKeyShift) >= 0 Then

    SBValue_efsVertical = SBValue(efsVertical)

    m_SelItem = SBValue_efsVertical + mIndex
    If Not IsValidSelection(m_SelItem) Then Exit Sub
    If m_SelItem > m_Items.count - 1 Then m_SelItem = m_Items.count - 1    'sur from unselect wrong last
    If m_SelItem = iPrev Then Exit Sub    'sur same selection click

    'If SBVisible(efsVertical) Then 'always big list

    If m_SelItem < SBValue_efsVertical Then

        If SBValue_efsVertical <> 0 Then SBValue_efsVertical = SBValue_efsVertical - 1
        'SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, SBValue(efsVertical) - 1)

    ElseIf m_SelItem >= SBValue_efsVertical + ScaleHeight \ m_RowHeight Then

        If SBValue_efsVertical <> SBMax(efsVertical) Then
            SBValue_efsVertical = SBValue_efsVertical + 1
        Else
            SBValue_efsVertical = SBMax(efsVertical)
            'SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), SBValue(efsVertical) + 1)
        End If
        
    End If
    SBValue(efsVertical) = SBValue_efsVertical
    
    'End If

    If m_SelStart < m_SelItem Then
        xStart = m_SelStart
        xEnd = m_SelItem
    Else
        xStart = m_SelItem
        xEnd = m_SelStart
    End If

    Selection_Clear
    For X1 = xStart To xEnd
        If X1 >= 0 Then m_Selected.Add X1
    Next X1

 

    'If Not m_MultiSelect Then RaiseEvent SelChange

    If Not m_SelItem = iPrev Then RaiseEvent SelChange(0)

   Selection_Sort
   
    Call ReDrawList(True)
    
    'RaiseEvent SelChange
    'If m_MultiSelect Then RaiseEvent SelChange

End If

' Not in multiselect mode. The LeftButton is down and mouse is moving..
' So scroll down the list

NextSelection:

If GetKeyState(vbKeyLButton) < 0 And Not m_MultiSelect Then

    mRowHeight = m_Def_RowHeight  '(ScaleHeight / m_iCount) 'sur
    m_SelItem = SBValue(efsVertical) + Y \ mRowHeight
    If Not IsValidSelection(m_SelItem) Then Exit Sub

    If m_SelItem >= 0 And m_SelItem < m_Items.count Then

        If Not m_SelItem = iPrev Then

            iPrev = m_SelItem
            Selection_Clear
            m_Selected.Add m_SelItem

            If SBVisible(efsVertical) Then
                If m_SelItem < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, m_SelItem)
                If m_SelItem >= SBValue(efsVertical) + ScaleHeight \ m_RowHeight Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), m_SelItem - m_iCount + 1)
            End If

            ReDrawList
            RaiseEvent SelChange(0)
        End If

        DoEvents
        If Not m_Inside Then GoTo NextSelection

    End If

End If

If Y \ m_RowHeight > Me.ListCount - 1 Then
    MousePointer = vbArrow
Else
    MousePointer = vbCustom
End If

iPrev = m_SelItem
RaiseEvent mouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If MouseUpBug Then Exit Sub

If m_SelItem > ListCount - 1 Or m_SelItem < 0 Then Exit Sub
'If m_MultiSelect Then RaiseEvent SelChange
RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Initialize()

' Used to prevent crashes on XP
'Debug.Print "Initializing..."
m_hMode = LoadLibrary("shell32.dll")
m_KeyControl = True

End Sub

Private Sub UserControl_InitProperties()

'Debug.Print "Initilizing Properties..."

m_BackColor = m_def_BackColor
m_ForeColor = m_def_ForeColor
m_SelColor = m_def_SelColor
m_SelForeColor = m_def_SelForeColor
Set m_Picture = Nothing
Set m_Font = Ambient.Font
m_StrechIcon = m_def_StrechIcon
m_RowHeight = m_Def_RowHeight
m_FullRowSel = m_def_FullRowSel
m_SortOrder = m_def_SortOrder
m_IconFocus = m_def_IconFocus
m_TextAllineMent = m_def_TextAllignMent
m_GridLines = m_def_GridLines
m_GridColor = m_Def_GridColor

m_BackGradient = m_def_BackGradient
m_SelGradient = m_def_SelGradient
m_BackGradientCol = m_def_BackGradientCol
m_SelGradientCol = m_def_SelGradientCol
m_MultiSelect = m_def_MultiSelect
m_FocusRectangle = m_def_FocusRectangle
m_ShowIcon = m_def_ShowIcon
m_FlatScrollBar = m_def_FlatScrollBar
m_bNoFlatScrollBars = Not m_def_FlatScrollBar
m_AutoHideScrollBars = m_def_AutoHideScrollBars
m_SelectionStyle = Style_XP
Me.BorderStyle = m_def_BorderStyle
Me.Appearance = m_def_Appearance
Me.BackColor = vbWhite
'Debug.Print "Properties Initialized!"

If m_hWnd = 0 Then InitializeSubClassing

End Sub

Private Sub UserControl_Resize()
'Debug.Print "Resizing..."
Me.Refresh
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch

FreeLibrary m_hMode
Me.Clear
pSBClearUp
Set m_Items = Nothing
Call Subclass_StopAll

Catch:
End Sub

Private Sub CheckSelected()

If m_Items.count = 0 Then m_SelItem = -1
If m_SelItem > m_Items.count - 1 Then m_SelItem = m_Items.count - 1
If m_SelItem < 0 Then m_SelItem = 0

End Sub

'------------------------------------------------------------------------------------------
' Procedure  : AddItem
' Auther     : Jim Jose
' Input      : New item
' OutPut     : None
' Purpose    : To add an item to listBox
'------------------------------------------------------------------------------------------

Public Sub AddItem(vText As String, _
                   Optional vIndex As Long = -1, _
                   Optional vIcon As Long = -1, _
                   Optional vBold As Boolean = False)

If vIndex = -1 Or vIndex = m_Items.count Then
' Index not specified , add to last
    m_Items.Add vText & SPLITER & vIcon & SPLITER & vBold
Else
' add to specified index
    m_Items.Add vText & SPLITER & vIcon & SPLITER & vBold, , vIndex + 1

End If

End Sub

'------------------------------------------------------------------------------------------
' Procedure  : Remove
' Auther     : Jim Jose
' Input      : Index
' OutPut     : None
' Purpose    : To remove an item from List
'------------------------------------------------------------------------------------------

Public Sub Remove(Optional ByVal vIndex As Long = -1)

If vIndex = -1 Then
' Index not specifid, remove selected item
    m_Items.Remove m_SelItem + 1
Else
' Remove specified item
    m_Items.Remove vIndex + 1
End If

' Sort If needed
'SortCollection m_Items, m_SortOrder
'Me.Refresh 'sur

End Sub

'------------------------------------------------------------------------------------------
' Procedure  : Clear
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Clear List
'------------------------------------------------------------------------------------------

Public Sub Clear()
Dim X As Long

' Remove each Item
For X = 1 To m_Items.count
    m_Items.Remove (1)
Next X
Selection_Clear

' Redraw
UserControl.Cls
Me.Refresh

End Sub

'------------------------------------------------------------------------------------------
' Procedure  : Refresh
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Arrage control and calculate local variables
'------------------------------------------------------------------------------------------

Public Sub Refresh()
On Error GoTo handle

' Determine item height & item count per Screen
'  m_iCount = Int(ScaleHeight \ m_RowHeight)
If m_RowHeight = 0 Then Exit Sub
m_iCount = ScaleHeight \ m_RowHeight
' Arrange\Set controls
If m_Items.count > m_iCount Then

    SBVisible(efsVertical) = True
    SBEnabled(efsVertical) = True
    SBMax(efsVertical) = m_Items.count - m_iCount
    SBMax(efsVertical) = SBMax(efsVertical) - 6
    SBSmallChange(efsVertical) = 1
    SBLargeChange(efsVertical) = 5

Else

    SBValue(efsVertical) = 0
    SBSmallChange(efsVertical) = 1
    SBLargeChange(efsVertical) = 5
    If m_AutoHideScrollBars Then
        SBVisible(efsVertical) = False
        SBEnabled(efsVertical) = True
    Else
        SBVisible(efsVertical) = True
        SBEnabled(efsVertical) = False
    End If

End If

handle:

'SortCollection m_Items, m_SortOrder
ReDrawList

End Sub


'------------------------------------------------------------------------------------------
' Procedure  : ReDrawList
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To draw the entire region
'------------------------------------------------------------------------------------------

Private Sub ReDrawList(Optional mouseMove As Boolean = False)
Dim mRowHeight As Long    'Double
Dim mIcon As StdPicture
Dim TxtData() As String
Dim mBold As Boolean
Dim mText As String
Dim mIndex As Long
Dim xStart As Long
Dim xMax As Long
Dim mTop As Long    'Double 'Long
Dim Rct As RECT
Dim X As Long
Dim TmpSelCol1 As Long
Dim TmpSelCol2 As Long
Dim TmpBorderCol As Long

'Debug.Print "Redrawing.."
On Error GoTo handle
UserControl.Cls

m_Left = 0

Call CheckSelected

xStart = SBValue(efsVertical)
xMax = xStart + m_iCount
mRowHeight = m_Def_RowHeight    '+ 0.6 'ScaleHeight \ m_iCount    'sur /

UserControl.DrawStyle = vbTransparent 'need

If m_ShowIcon Then m_Left = mRowHeight + 3

'    ' To view style on design mode
'    If Ambient.UserMode = False And m_Items.Count = 0 Then
'        m_Items.Add "McListBox" & SPLITER & "-1" & SPLITER & "False"
'        m_Selected.Add 0
'    End If
'
' Draw back

UserControl.Picture = m_Picture

If m_BackGradient Then FillGradient hdc, 0, 0, ScaleWidth, ScaleHeight, m_BackColor, m_BackGradientCol, m_BackGradient
If m_Items.count = 0 Then Exit Sub
' Set selection colors

'If m_SelectionStyle = Style_XP Then
    TmpBorderCol = BlendColor(m_SelColor, vbBlack, 250)
    If m_HasFocus Then
        TmpSelCol1 = BlendColor(m_SelColor, vbWhite, 150)    '150
        TmpSelCol2 = TmpSelCol1
'TmpSelCol2 = BlendColor(m_SelColor, vbWhite, 150) '100
    Else
        TmpSelCol1 = BlendColor(m_SelColor, vbWhite, 75)    '75)
'TmpSelCol2 = BlendColor(m_SelColor, vbWhite, 100)
        TmpSelCol2 = TmpSelCol1
    End If
'Else
'    TmpBorderCol = m_SelColor
'    TmpSelCol1 = m_SelColor
'    TmpSelCol2 = TmpSelCol1
'End If


For X = xStart To xMax - 1

' Get the data
    TxtData = Split(m_Items(X + 1), SPLITER)
    mText = TxtData(0)

    If Not TxtData(1) = -1 And m_ShowIcon Then Set mIcon = ImArr(Val(TxtData(1)))

    If TxtData(2) = "True" Then mBold = True Else mBold = False

' Find if Selected!?
    mIndex = Selection_Find(X)
    mTop = (X - xStart) * mRowHeight

    If mIndex > 0 Then

' Fill Selection

      '  If m_SelGradient = Fill_None Then

            UserControl.ForeColor = TmpBorderCol
            If X = m_SelItem Then UserControl.FillColor = TmpSelCol1 Else UserControl.FillColor = TmpSelCol2

            If m_FullRowSel Then
                Rectangle hdc, 1, mTop + 1, ScaleWidth, mTop + mRowHeight
            Else
                Rectangle hdc, m_Left + 1, mTop + 1, ScaleWidth, mTop + mRowHeight
            End If

'        Else
'            If m_FullRowSel Then
'                FillGradient hdc, 0, mTop + 1, ScaleWidth, mRowHeight, m_SelColor, m_SelGradientCol, m_SelGradient
'            Else
'                FillGradient hdc, m_Left, mTop + 1, ScaleWidth - m_Left, mRowHeight, m_SelColor, m_SelGradientCol, m_SelGradient
'            End If
'
'        End If

' Draw Icon Focus
        If m_IconFocus And Not m_FullRowSel And m_ShowIcon And m_HasFocus Then
            UserControl.ForeColor = vbBlack
            SetRect Rct, 1, mTop + 1, mRowHeight, mTop + mRowHeight
            DrawFocusRect UserControl.hdc, Rct
        End If

        UserControl.ForeColor = m_SelForeColor

    Else
        UserControl.ForeColor = m_ForeColor

    End If

' Draw the Text
    UserControl.FontBold = mBold
    DrawText mText, m_Left, mTop + 1

' Draw Icon
    If Not TxtData(1) = -1 And m_ShowIcon Then DrawPicture mIcon, 10, mTop + 1, mRowHeight - 1, mRowHeight - 0

'Draw Grid
    If m_GridLines Then
        UserControl.ForeColor = m_GridColor
        Line (0, mTop)-(ScaleWidth, mTop)
    End If

    If X = m_SelItem Then

' Draw Focus Rect
'        If m_FocusRectangle Then
'            UserControl.ForeColor = vbBlack
'            If m_FullRowSel Or Not m_ShowIcon Then
'                SetRect Rct, -1, mTop, ScaleWidth + 1, mTop + mRowHeight + 1
'            Else
'                SetRect Rct, m_Left, mTop, ScaleWidth + 1, mTop + mRowHeight + 1
'            End If
'            DrawFocusRect hdc, Rct
'        End If

    End If
Next X

'Line (0, mTop)-(ScaleWidth, mTop)

If Not mouseMove Then
UserControl.Refresh 'rem good, но эффект тормоза при скролле
End If

'Debug.Print "Redrawing Completed!!"
handle:
End Sub


Private Sub DrawText(ByVal lpStr As String, _
                     ByVal X As Long, ByVal Y As Long)
Dim Rct As RECT

' Set the Rect
Rct.left = X + 20
Rct.top = Y + (m_RowHeight - TextHeight("A")) / 2
Rct.right = ScaleWidth
Rct.bottom = Y + m_RowHeight

' Draw the text
If IsNT Then
    DrawTextW hdc, StrPtr(lpStr), -1, Rct, m_TextAllineMent
Else
    DrawTextA hdc, lpStr, -1, Rct, m_TextAllineMent
End If

End Sub


Private Sub DrawPicture(vPicture As StdPicture, _
                        ByVal X As Long, ByVal Y As Long, _
                        ByVal vWidth As Long, ByVal vHeight As Long)

Dim prop As Integer

'If Not m_StrechIcon Then
If vPicture = Null Then Exit Sub
If vPicture.Height = 0 Then Exit Sub
If vPicture.Width = 0 Then Exit Sub
If vHeight = 0 Then Exit Sub
If vWidth = 0 Then Exit Sub

If vPicture.Width >= vPicture.Height Then
    prop = vPicture.Width / vPicture.Height
    If ScaleX(vPicture.Width) > vWidth Then
'vWidth = ScaleX(vPicture.Width)
        vHeight = vHeight / prop
    Else
        vHeight = ScaleY(vPicture.Height)
        vWidth = ScaleX(vPicture.Width)
    End If

Else
    prop = vPicture.Height / vPicture.Width
    If ScaleY(vPicture.Height) > vHeight Then
        vWidth = vWidth / prop
'vHeight = Scaley(vPicture.Height)
    Else
        vHeight = ScaleY(vPicture.Height)
        vWidth = ScaleX(vPicture.Width)
    End If
End If

If vHeight = 0 Then Exit Sub
If vWidth = 0 Then Exit Sub
'X = X + (vWidth - ScaleX(vPicture.Width)) / 2
'Y = Y + (vHeight - ScaleY(vPicture.Height)) / 2

If Magnify Then
    If vWidth <= 16 Then
        If vHeight <= 13 Then
'prop = 22 / vHeight
            vHeight = vHeight * 2
            vWidth = vWidth * 2
        End If
    End If
End If


X = X + (m_Def_RowHeight - vWidth) / 2
Y = Y + (m_Def_RowHeight - vHeight) / 2

PaintPicture vPicture, X, Y, vWidth, vHeight, Opcode:=vbNotSrcCopy      ', X, Y, ScaleY(vPicture.Width), ScaleY(vPicture.Height)

'If ScaleX(vPicture.Width) > vWidth Then
'
'    X = X + (vWidth - ScaleX(vPicture.Width)) / 2
'    Y = Y + (vHeight - ScaleY(vPicture.Height)) / 2
'    PaintPicture vPicture, X, Y
'Else
'    PaintPicture vPicture, X, Y, vWidth, vHeight
'End If

End Sub

Private Sub InitializeSubClassing()
On Error GoTo handle

If Not m_hWnd = 0 Then Exit Sub

' Subclass in runtime
If Ambient.UserMode Then

    bTrack = True
    bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

    If Not bTrackUser32 Then
        If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
            bTrack = False
        End If
    End If

    If Not bTrack Then Exit Sub

    With UserControl

' Start subclassing our calendar
        Call Subclass_Start(.hWnd)

' Adding the messages we need to track
        Call Subclass_AddMsg(.hWnd, WM_MOUSEWHEEL, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_MOUSEMOVE, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_MOUSELEAVE, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_MOUSEHOVER, MSG_AFTER)
        Call Subclass_AddMsg(.hWnd, WM_VSCROLL, MSG_AFTER)

    End With

End If

'Create Standard Scrollbars
If m_hWnd = 0 Then SBCreate UserControl.hWnd
CreateItemComplete

handle:
End Sub

Private Sub Selection_Clear()
Dim X As Long
Dim xMax As Long
Dim lCurrentSelected As Long

lCurrentSelected = m_SelItem

xMax = m_Selected.count
For X = 1 To xMax
    m_Selected.Remove (1)
Next X

End Sub


Private Function Selection_Find(ByVal vIdex As Long) As Long
Dim X As Long
Dim xMax As Long

xMax = m_Selected.count

For X = 1 To xMax
    If m_Selected(X) = vIdex Then Selection_Find = X: Exit Function
Next X

End Function


Private Sub Selection_Sort()
Dim X As Long
Dim vPos As Long
Dim vCount As Long
Dim vStart As Long
Dim vNewCount As Long
Dim vNew As New Collection

vStart = 1
vCount = m_Selected.count

For X = vStart To vCount
    vNewCount = vNew.count
    For vPos = 1 To vNewCount
        If vNew(vPos) < m_Selected(X) Then Exit For
    Next vPos

    If vPos = 1 Then
        vNew.Add m_Selected(X)
    Else
        vNew.Add m_Selected(X), , vPos - 1
    End If

Next X

Set m_Selected = vNew

End Sub


Private Sub CreateItemComplete()

' Create a new object
Set m_PicComplete = UserControl.Controls.Add("vb.PictureBox", "PicComplete")
m_PicComplete.AutoRedraw = True
m_PicComplete.BorderStyle = 1
m_PicComplete.Appearance = 0
m_PicComplete.BackColor = UserControl.FillColor
m_PicComplete.ForeColor = vbBlack
m_PicComplete.Enabled = False

' Hide window from TaskBar
SetParent m_PicComplete.hWnd, GetDesktopWindow
SetWindowLong m_PicComplete.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW

End Sub


'Private Sub ShowItemComplete(ByVal lItem As Long)
'Dim mArray() As String
'Dim mTmp As Long
'Dim Rct As RECT
'Dim wRect As RECT
'Dim ScrX As Long
'Dim ScrY As Long
'Dim lWidth As Long
'Dim sText As String
'Dim mRowHeight As Long    'Double
'
''On Error GoTo Handle
'If m_Items.Count = 0 Then Exit Sub
'If m_Items.Count < SBValue(efsVertical) + lItem Then Exit Sub
'mArray = Split(m_Items(SBValue(efsVertical) + lItem), SPLITER)
'
'sText = mArray(0)
'UserControl.FontBold = mArray(2)
'lWidth = TextWidth(sText)
'If lWidth < ScaleWidth - m_Left Then GoTo Handle
'If m_LastComplete = lItem Then Exit Sub
'mRowHeight = ScaleHeight \ m_iCount    'sur /
'ScrX = Screen.TwipsPerPixelX
'ScrY = Screen.TwipsPerPixelY
'
'' Set the Rect
'Rct.Left = 5
'Rct.Top = (mRowHeight - TextHeight("A")) / 2 - 2
'Set m_PicComplete.Font = m_Font
'm_PicComplete.Cls
'
'' Move it
'GetWindowRect hwnd, wRect
'mTmp = Screen.Width \ ScrX - (wRect.Left + m_Left)    'sur
'If mTmp < lWidth Then
'
'    lWidth = mTmp - 10
'    mArray = SplitToLines(sText, mTmp - 15)
'    sText = Join(mArray, vbCrLf)
'    mTmp = TextHeight(sText)
'    If mTmp < mRowHeight Then
'        Rct.Top = (mRowHeight - mTmp) / 2 - 2
'    Else
'        Rct.Top = 5
'    End If
'    m_PicComplete.Move (wRect.Left + m_Left + 2) * ScrX, Int((wRect.Top + 3 + (lItem - 1) * mRowHeight) * ScrY), (lWidth) * ScrX, (5 + mTmp + 5) * ScrY
'
'Else
'    m_PicComplete.Move (wRect.Left + m_Left + 2) * ScrX, Int((wRect.Top + 3 + (lItem - 1) * mRowHeight) * ScrY), (lWidth + 12) * ScrX, (mRowHeight - 1) * ScrY
'End If
'
'Rct.Right = m_PicComplete.ScaleWidth
'Rct.Bottom = m_PicComplete.ScaleHeight
'
'' Draw the text
'If IsNT Then
'    DrawTextW m_PicComplete.hdc, StrPtr(sText), -1, Rct, vbLeftJustify
'Else
'    DrawTextA m_PicComplete.hdc, sText, -1, Rct, vbLeftJustify
'End If
'
'' Show
'm_PicComplete.ZOrder (0)
'm_PicComplete.Visible = True
'm_LastComplete = lItem
'
'Exit Sub
'Handle:
'm_PicComplete.Visible = False
'm_LastComplete = -100
'End Sub
'
'

' By Gary Noble
Private Function IsValidSelection(ByVal lSelectedItemIndex As Long) As Boolean
On Error Resume Next

If lSelectedItemIndex > ListCount Then
    m_SelItem = ListCount - 1
    IsValidSelection = False
ElseIf lSelectedItemIndex <= -1 Then
    m_SelItem = 0
    IsValidSelection = False
Else
    m_SelItem = lSelectedItemIndex
    IsValidSelection = True
End If

On Error GoTo 0
End Function

'------------------------------------------------------------------------------------------
' Procedure  : UserControl_WriteProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Write design time propery changes
'------------------------------------------------------------------------------------------

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    'Debug.Print "Writing Properties..."
Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
Call PropBag.WriteProperty("SelColor", m_SelColor, m_def_SelColor)
Call PropBag.WriteProperty("SelForeColor", m_SelForeColor, m_def_SelForeColor)
Call PropBag.WriteProperty("StrechIcon", m_StrechIcon, m_def_StrechIcon)
Call PropBag.WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
Call PropBag.WriteProperty("FullRowSelect", m_FullRowSel, m_def_FullRowSel)
Call PropBag.WriteProperty("SortOrder", m_SortOrder, m_def_SortOrder)
Call PropBag.WriteProperty("IconFocus", m_IconFocus, m_def_IconFocus)
Call PropBag.WriteProperty("TextAlignment", m_TextAllineMent, m_def_TextAllignMent)
Call PropBag.WriteProperty("RowHeight", m_RowHeight, m_Def_RowHeight)
Call PropBag.WriteProperty("GridLines", m_GridLines, m_def_GridLines)
Call PropBag.WriteProperty("GridColor", m_GridColor, m_Def_GridColor)
Call PropBag.WriteProperty("BackGradient", m_BackGradient, m_def_BackGradient)
Call PropBag.WriteProperty("SelGradient", m_SelGradient, m_def_SelGradient)
Call PropBag.WriteProperty("BackGradientCol", m_BackGradientCol, m_def_BackGradientCol)
Call PropBag.WriteProperty("SelGradientCol", m_SelGradientCol, m_def_SelGradientCol)
Call PropBag.WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
Call PropBag.WriteProperty("FocusRectangle", m_FocusRectangle, m_def_FocusRectangle)
Call PropBag.WriteProperty("ShowIcon", m_ShowIcon, m_def_ShowIcon)
Call PropBag.WriteProperty("SelectionStyle", m_SelectionStyle, m_def_SelectionStyle)
Call PropBag.WriteProperty("FlatScrollBar", m_FlatScrollBar, m_def_FlatScrollBar)
Call PropBag.WriteProperty("AutoHideScrollBars", m_AutoHideScrollBars, m_def_AutoHideScrollBars)
'    'Debug.Print "Completed Reading properties!!"

End Sub

'------------------------------------------------------------------------------------------
' Procedure  : UserControl_ReadProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Read design time propery changes
'------------------------------------------------------------------------------------------

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    'Debug.Print "Reading Properties..."
m_SelItem = -1
Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
m_SelColor = PropBag.ReadProperty("SelColor", m_def_SelColor)
m_SelForeColor = PropBag.ReadProperty("SelForeColor", m_def_SelForeColor)
m_StrechIcon = PropBag.ReadProperty("StrechIcon", m_def_StrechIcon)
Me.Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
Me.BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
m_FullRowSel = PropBag.ReadProperty("FullRowSelect", m_def_FullRowSel)
m_SortOrder = PropBag.ReadProperty("SortOrder", m_def_SortOrder)
m_IconFocus = PropBag.ReadProperty("IconFocus", m_def_IconFocus)
m_TextAllineMent = PropBag.ReadProperty("TextAlignment", m_def_TextAllignMent)
m_RowHeight = PropBag.ReadProperty("RowHeight", m_Def_RowHeight)
Set UserControl.Font = m_Font
If m_RowHeight < TextHeight("A") Then m_RowHeight = TextHeight("A")
m_GridLines = PropBag.ReadProperty("GridLines", m_def_GridLines)
m_GridColor = PropBag.ReadProperty("GridColor", m_Def_GridColor)
m_BackGradient = PropBag.ReadProperty("BackGradient", m_def_BackGradient)
m_SelGradient = PropBag.ReadProperty("SelGradient", m_def_SelGradient)
m_BackGradientCol = PropBag.ReadProperty("BackGradientCol", m_def_BackGradientCol)
m_SelGradientCol = PropBag.ReadProperty("SelGradientCol", m_def_SelGradientCol)
m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
UserControl.BackColor = m_BackColor
m_FocusRectangle = PropBag.ReadProperty("FocusRectangle", m_def_FocusRectangle)
m_ShowIcon = PropBag.ReadProperty("ShowIcon", m_def_ShowIcon)
m_SelectionStyle = PropBag.ReadProperty("SelectionStyle", m_def_SelectionStyle)
m_AutoHideScrollBars = PropBag.ReadProperty("AutoHideScrollBars", m_def_AutoHideScrollBars)
m_FlatScrollBar = PropBag.ReadProperty("FlatScrollBar", m_def_FlatScrollBar)
m_bNoFlatScrollBars = Not m_FlatScrollBar

InitializeSubClassing

'    'Debug.Print "Completed Reading Properties!!"
Me.Refresh

End Sub


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : SplitToLines
' Auther    : Jim Jose
' Input     : Object, Text to split an parameters
' OutPut    : Splitted text array
' Purpose   : Split a string into lines by length!
'------------------------------------------------------------------------------------------------------------------------------------------

'Public Function SplitToLines(ByVal sText As String, ByVal lLength As Long, _
  '                             Optional ByVal bFilterLines As Boolean = True) As String()
'Dim mArray() As String
'Dim mChar As String
'Dim mLine As String
'Dim lnCount As Long
'Dim xMax As String
'Dim mPos As Long
'Dim X As Long
'Dim lDone As Long
'
'If bFilterLines Then sText = Replace(sText, vbNewLine, vbNullString)
'xMax = Len(sText)
'
'For X = 1 To xMax
'
'    mChar = Mid(sText, X, 1)
'
'    If IsDelim(mChar) Then mPos = X - (lDone + 1)
'    If TextWidth(mLine & mChar) >= lLength Or X = xMax Then
'        If mPos = 0 Then mPos = X - (lDone + 1)
'        ReDim Preserve mArray(lnCount)
'        mArray(lnCount) = RTrim(LTrim(Mid(mLine, 1, mPos)))
'        mLine = Mid(mLine, mPos + 1, Len(mLine) - mPos)
'        lDone = lDone + mPos: mPos = 0
'        lnCount = lnCount + 1
'    End If
'
'    mLine = mLine & mChar
'
'Next X
'
'mArray(lnCount - 1) = mArray(lnCount - 1) & mChar
'SplitToLines = mArray
'
'End Function


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : IsDelim
' Auther    : Rde
' Input     : Char
' OutPut    : IsDelim?
' Purpose   : Check if the input char is a Delimiter or not!
'------------------------------------------------------------------------------------------------------------------------------------------

Public Function IsDelim(Char As String) As Boolean
Select Case Asc(Char)    ' Upper/Lowercase letters,Underscore Not delimiters
Case 65 To 90, 95, 97 To 122
    IsDelim = False
Case Else: IsDelim = True    ' Another Character Is delimiter
End Select
End Function


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : IsNT
' Auther    : Dana Seaman
' Input     : None
' OutPut    : NT?
' Purpose   : Check for the NT Platform
'------------------------------------------------------------------------------------------------------------------------------------------

Private Function IsNT() As Boolean
Static m_bInit As Boolean
Dim udtVer As OSVERSIONINFO
On Error Resume Next

'Cache m_bIsNT on first execution
If Not m_bInit Then
    m_bInit = True
    udtVer.dwOSVersionInfoSize = Len(udtVer)
    If GetVersionEx(udtVer) Then
        If udtVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            m_bIsNT = True
        End If
    End If
End If
IsNT = m_bIsNT

On Error GoTo 0
End Function

'------------------------------------------------------------------------------------------
' Procedure  : SortCollection
' Auther     : Jim Jose
' Input      : Collection to Sort + Order
' OutPut     : None
' Purpose    : To sort the Data-Collection Ascending/Descending
'------------------------------------------------------------------------------------------

Private Sub SortCollection(ByRef vCollection As Collection, _
                           ByVal vSortOrder As SortOrderEnum)
Dim X As Long
Dim vPos As Long
Dim vRtn As Long
Dim vCount As Long
Dim vStart As Long
Dim vNewCount As Long
Dim vNew As New Collection

' Check Sort?
If vSortOrder = Sort_None Then Exit Sub

' Get current Count
vStart = 1
vCount = vCollection.count

' Loop through Current collection
For X = vStart To vCount

' Get new collection count
    vNewCount = vNew.count

' Loop through new collection
    For vPos = 1 To vNewCount

' Compair each item in new collection
        vRtn = StrComp(vCollection(X), vNew(vPos), vbTextCompare)
' Escape with purpose
        If vRtn = vSortOrder Then Exit For

    Next vPos

    If X = vStart Or vPos = vNewCount + 1 Then
' New item at last
        vNew.Add vCollection(X), "K " & X
    Else
' New item somewhere b/w
        vNew.Add vCollection(X), "K " & X, vPos
    End If

Next X

' Return Sorted Collection
Set vCollection = vNew

End Sub

' -------------------------------------------------------------------------------------
' Procedure : BlendColor
' Type      : Property
' DateTime  : 03/02/2005
' Author    : Gary Noble
' Purpose   : Blends Two Colours Together
' Returns   : Long
' -------------------------------------------------------------------------------------

Public Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR, _
                               Optional ByVal Alpha As Long = 128) As Long

Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
Dim lCFrom As Long
Dim lCTo As Long

lCFrom = TranslateColor(oColorFrom)
lCTo = TranslateColor(oColorTo)
lSrcR = lCFrom And &HFF
lSrcG = (lCFrom And &HFF00&) \ &H100&
lSrcB = (lCFrom And &HFF0000) \ &H10000
lDstR = lCTo And &HFF
lDstG = (lCTo And &HFF00&) \ &H100&
lDstB = (lCTo And &HFF0000) \ &H10000
BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property

' -------------------------------------------------------------------------------------
' Procedure : TranslateColor
' Type      : Function
' DateTime  : 03/02/2005
' Author    : Gary Noble
' Purpose   : Convert Automation color to Windows color
' Returns   : Long
' -------------------------------------------------------------------------------------

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                               Optional hPal As Long = 0) As Long

If OleTranslateColor(oClr, hPal, TranslateColor) Then
    TranslateColor = -1
End If

End Function


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : FillGradient
' Auther    : Jim Jose
' Input     : Hdc + Parameters
' OutPut    : None
' Purpose   : Middleout Gradients with Carls's DIB solution
'------------------------------------------------------------------------------------------------------------------------------------------

Private Sub FillGradient(ByVal hdc As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As GradientDirectionEnum, _
                         Optional Right2Left As Boolean = True)

Dim tmpCol As Long

' Exit if needed
If GradientDirection = Fill_None Then Exit Sub

' Right-To-Left
If Right2Left Then
    tmpCol = Col1
    Col1 = Col2
    Col2 = tmpCol
End If

' Translate system colors
If Col1 < 0 Then Col1 = TranslateColor(Col1)
If Col2 < 0 Then Col2 = TranslateColor(Col2)

Select Case GradientDirection
Case Fill_HorizontalMiddleOut
    DIBGradient hdc, X, Y, Width / 2, Height, Col1, Col2, Fill_Horizontal
    DIBGradient hdc, X + Width / 2 - 1, Y, Width / 2, Height, Col2, Col1, Fill_Horizontal

Case Fill_VerticalMiddleOut
    DIBGradient hdc, X, Y, Width, Height / 2, Col1, Col2, Fill_Vertical
    DIBGradient hdc, X, Y + Height / 2 - 1, Width, Height / 2, Col2, Col1, Fill_Vertical

Case Else
    DIBGradient hdc, X, Y, Width, Height, Col1, Col2, GradientDirection
End Select

End Sub

'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : DIBGradient
' Auther    : Carls P.V.
' Input     : Hdc + Parameters
' OutPut    : None
' Purpose   : DIB solution for fast gradients
'------------------------------------------------------------------------------------------------------------------------------------------

Private Sub DIBGradient(ByVal hdc As Long, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        ByVal Width As Long, _
                        ByVal Height As Long, _
                        ByVal Col1 As Long, _
                        ByVal Col2 As Long, _
                        ByVal GradientDirection As GradientDirectionEnum)

Dim uBIH As BITMAPINFOHEADER
Dim lBits() As Long
Dim lGrad() As Long

Dim R1 As Long
Dim G1 As Long
Dim B1 As Long
Dim R2 As Long
Dim G2 As Long
Dim B2 As Long
Dim dR As Long
Dim dG As Long
Dim dB As Long

Dim Scan As Long
Dim i As Long
Dim iEnd As Long
Dim iOffset As Long
Dim j As Long
Dim jEnd As Long
Dim iGrad As Long

'-- A minor check
If (Width < 1 Or Height < 1) Then Exit Sub

'-- Decompose Cols
Col1 = Col1 And &HFFFFFF
R1 = Col1 Mod &H100&
Col1 = Col1 \ &H100&
G1 = Col1 Mod &H100&
Col1 = Col1 \ &H100&
B1 = Col1 Mod &H100&
Col2 = Col2 And &HFFFFFF
R2 = Col2 Mod &H100&
Col2 = Col2 \ &H100&
G2 = Col2 Mod &H100&
Col2 = Col2 \ &H100&
B2 = Col2 Mod &H100&

'-- Get Col distances
dR = R2 - R1
dG = G2 - G1
dB = B2 - B1

'-- Size gradient-Cols array
Select Case GradientDirection
Case [Fill_Horizontal]
    ReDim lGrad(0 To Width - 1)
Case [Fill_Vertical]
    ReDim lGrad(0 To Height - 1)
Case Else
    ReDim lGrad(0 To Width + Height - 2)
End Select

'-- Calculate gradient-Cols
iEnd = UBound(lGrad())
If (iEnd = 0) Then
'-- Special case (1-pixel wide gradient)
    lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
Else
    For i = 0 To iEnd
        lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
    Next i
End If

'-- Size DIB array
ReDim lBits(Width * Height - 1) As Long
iEnd = Width - 1
jEnd = Height - 1
Scan = Width

'-- Render gradient DIB
Select Case GradientDirection

Case [Fill_Horizontal]

    For j = 0 To jEnd
        For i = iOffset To iEnd + iOffset
            lBits(i) = lGrad(i - iOffset)
        Next i
        iOffset = iOffset + Scan
    Next j

Case [Fill_Vertical]

    For j = jEnd To 0 Step -1
        For i = iOffset To iEnd + iOffset
            lBits(i) = lGrad(j)
        Next i
        iOffset = iOffset + Scan
    Next j

Case [Fill_DownwardDiagonal]

    iOffset = jEnd * Scan
    For j = 1 To jEnd + 1
        For i = iOffset To iEnd + iOffset
            lBits(i) = lGrad(iGrad)
            iGrad = iGrad + 1
        Next i
        iOffset = iOffset - Scan
        iGrad = j
    Next j

Case [Fill_UpwardDiagonal]

    iOffset = 0
    For j = 1 To jEnd + 1
        For i = iOffset To iEnd + iOffset
            lBits(i) = lGrad(iGrad)
            iGrad = iGrad + 1
        Next i
        iOffset = iOffset + Scan
        iGrad = j
    Next j
End Select

'-- Define DIB header
With uBIH
    .biSize = 40
    .biPlanes = 1
    .biBitCount = 32
    .biWidth = Width
    .biHeight = Height
End With

'-- Paint it!
Call StretchDIBits(hdc, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)

End Sub

'-----------------------------------------------------------------------------------------------------------
' Following is the API ScrollBar code by 'Gary Noble'. All the properties
' and functions are handles here.
'-----------------------------------------------------------------------------------------------------------
' Auther    :   Gary Noble
' Purpose   :   API ScrollBars
' About     :   This code uses Paul Carton's self subclaser. Create the ScrollBar(horz/vert),
'               on the event 'ReadProperties'. This code also enable 'MouseWheel'.
'               Gary, thanks a lot for providing this best code for API scrollBars...
'-----------------------------------------------------------------------------------------------------------

Friend Property Get SBVisible(ByVal eBar As EFSScrollBarConstants) As Boolean
If (eBar = efsHorizontal) Then
    SBVisible = m_bVisibleHorz
Else
    SBVisible = m_bVisibleVert
End If
End Property

Friend Property Let SBVisible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
If (eBar = efsHorizontal) Then
    m_bVisibleHorz = bState
Else
    m_bVisibleVert = bState
End If
'   If (m_bNoFlatScrollBars) Then
ShowScrollBar m_hWnd, eBar, Abs(bState)
'   Else
'      FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
'  End If
End Property

Friend Property Get Orientation() As ScrollBarOrienationEnum
SBOrientation = m_eOrientation
End Property

Friend Property Let SBOrientation(ByVal eOrientation As ScrollBarOrienationEnum)
m_eOrientation = eOrientation
pSBSetOrientation
End Property

Private Sub pSBSetOrientation()
ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Horizontal))
ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Vertical))
End Sub

'Private Sub SBRefresh()
'EnableScrollBar m_hWnd, SB_VERT, ESB_ENABLE_BOTH
'End Sub

Private Sub pSBGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim Lo As Long

Lo = eBar
tSI.fMask = fMask
tSI.cbSize = LenB(tSI)

'  If (m_bNoFlatScrollBars) Then
GetScrollInfo m_hWnd, Lo, tSI
'    Else
'       FlatSB_GetScrollInfo m_hWnd, Lo, tSI
'   End If

End Sub

Private Sub pSBLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim Lo As Long

Lo = eBar
tSI.fMask = fMask
tSI.cbSize = LenB(tSI)

'    If (m_bNoFlatScrollBars) Then
SetScrollInfo m_hWnd, Lo, tSI, True
'   Else
'        FlatSB_SetScrollInfo m_hWnd, Lo, tSI, True
'   End If

End Sub

Friend Property Get SBStyle() As ScrollBarStyleEnum
SBStyle = m_eStyle
End Property

Friend Property Let SBStyle(ByVal eStyle As ScrollBarStyleEnum)
'Dim lR As Long
'If (m_bNoFlatScrollBars) Then
'    ' can't do it..
'    ''Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
Exit Property
'Else
'    If (m_eOrientation = Scroll_Horizontal) Or (m_eOrientation = Scroll_Both) Then
'        lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
'    End If
'    If (m_eOrientation = Scroll_Vertical) Or (m_eOrientation = Scroll_Both) Then
'        lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
'    End If
'    ''Debug.Print lR
'    m_eStyle = eStyle
'End If
End Property

Friend Property Get SBSmallChange(ByVal eBar As EFSScrollBarConstants) As Long
If (eBar = efsHorizontal) Then
    SBSmallChange = m_lSmallChangeHorz
Else
    SBSmallChange = m_lSmallChangeVert
End If
End Property

Friend Property Let SBSmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
If (eBar = efsHorizontal) Then
    m_lSmallChangeHorz = lSmallChange
Else
    m_lSmallChangeVert = lSmallChange
End If
End Property

Friend Property Get SBEnabled(ByVal eBar As EFSScrollBarConstants) As Boolean
If (eBar = efsHorizontal) Then
    SBEnabled = m_bEnabledHorz
Else
    SBEnabled = m_bEnabledVert
End If
End Property

Friend Property Let SBEnabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
Dim Lo As Long
Dim lF As Long

Lo = eBar
If (bEnabled) Then
    lF = ESB_ENABLE_BOTH
Else
    lF = ESB_DISABLE_BOTH
End If
'   If (m_bNoFlatScrollBars) Then
EnableScrollBar m_hWnd, Lo, lF
'    Else
'        FlatSB_EnableScrollBar m_hWnd, Lo, lF
'    End If

End Property

Friend Property Get SBMin(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
pSBGetSI eBar, tSI, SIF_RANGE
SBMin = tSI.nMin
End Property

Friend Property Get SBMax(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
pSBGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
SBMax = tSI.nMax                                  ' - tSI.nPage
End Property

Friend Property Get SBValue(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
pSBGetSI eBar, tSI, SIF_POS
SBValue = tSI.nPos
End Property

Friend Property Get SBLargeChange(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
pSBGetSI eBar, tSI, SIF_PAGE
SBLargeChange = tSI.nPage
End Property

Friend Property Let SBMin(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
Dim tSI As SCROLLINFO
tSI.nMin = iMin
tSI.nMax = SBMax(eBar) + SBLargeChange(eBar)
pSBLetSI eBar, tSI, SIF_RANGE
End Property

Friend Property Let SBMax(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
Dim tSI As SCROLLINFO
tSI.nMax = iMax + SBLargeChange(eBar)
tSI.nMin = SBMin(eBar)
pSBLetSI eBar, tSI, SIF_RANGE
End Property

Friend Property Let SBValue(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
Dim tSI As SCROLLINFO
If (iValue <> SBValue(eBar)) Then
    tSI.nPos = iValue
    pSBLetSI eBar, tSI, SIF_POS
'ReDrawList
End If
End Property

Friend Property Let SBLargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
Dim tSI As SCROLLINFO

pSBGetSI eBar, tSI, SIF_ALL
tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
tSI.nPage = iLargeChange
pSBLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property

Friend Property Get SBCanBeFlat() As Boolean
SBCanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub pSBCreateScrollBar()
On Error Resume Next
m_bNoFlatScrollBars = True
SBStyle = m_eStyle
On Error GoTo 0
End Sub

Friend Sub SBCreate(ByVal hWndA As Long)
pSBClearUp
m_hWnd = hWndA
pSBCreateScrollBar
End Sub

Private Sub pSBClearUp()
On Error Resume Next
m_hWnd = 0
m_bInitialised = False
On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------------------------------------------------------------
' The following bytes are donated exclusively for Paul Caton's Subclassing
' We need this to track the movement information of the m_picCalendar and
' sizing/positioning of parent form
'---------------------------------------------------------------------------------------------------------------------------------------------
' Auther    : Paul Caton
' Purpose   : Advanced subclassing for UserControls (Self subclasser)
' Comment   : Thanks a Billion for this ever green piece of code on subclassing!!!
'---------------------------------------------------------------------------------------------------------------------------------------------

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
        Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
        Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
End With
End Sub


'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
'lng_hWnd  - The handle of the window to be subclassed
'Returns;
'The sc_aSubData() index
Const CODE_LEN As Long = 200                                'Length of the machine code in bytes
Const FUNC_CWP As String = "CallWindowProcA"                'We use CallWindowProc to call the original WndProc
Const FUNC_EBM As String = "EbMode"                         'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWL As String = "SetWindowLongA"                 'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER As String = "user32"                         'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5 As String = "vba5"                           'Location of the EbMode function if running VB5
Const MOD_VBA6 As String = "vba6"                           'Location of the EbMode function if running VB6
Const PATCH_01 As Long = 18                                 'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02 As Long = 68                                 'Address of the previous WndProc
Const PATCH_03 As Long = 78                                 'Relative address of SetWindowsLong
Const PATCH_06 As Long = 116                                'Address of the previous WndProc
Const PATCH_07 As Long = 121                                'Relative address of CallWindowProc
Const PATCH_0A As Long = 186                                'Address of the owner object
Static aBuf(1 To CODE_LEN) As Byte                                                  'Static code buffer byte array
Static pCWP As Long                                   'Address of the CallWindowsProc
Static pEbMode As Long                                      'Address of the EbMode IDE break/stop/running function
Static pSWL As Long                                   'Address of the SetWindowsLong function
Dim i As Long                             'Loop index
Dim j As Long                             'Loop index
Dim nSubIdx As Long                                   'Subclass data index
Dim sHex As String                              'Hex code string

'If it's the first time through here..
If aBuf(1) = 0 Then

'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
        j = j + 1
        aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
        i = i + 2
    Loop                                                                                'Next pair of hex characters

'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
        aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
        aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
        pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
        If pEbMode = 0 Then                                                               'Found?
            pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
        End If
    End If

    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
        nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
        ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If

    Subclass_Start = nSubIdx
End If

With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
Dim i As Long
On Error Resume Next
i = UBound(sc_aSubData())      'Get the upper bound of the subclass data array
If Err Then Exit Sub

Do While i >= 0                'Iterate through each element
    With sc_aSubData(i)
        If .hWnd <> 0 Then             'If not previously Subclass_Stop'd
            Call Subclass_Stop(.hWnd)     'Subclass_Stop
        End If
    End With

    i = i - 1           'Next element
Loop

On Error GoTo 0
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
'lng_hWnd  - The handle of the window to stop being subclassed
With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
End With
End Sub

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
Dim nEntry As Long                                                                  'Message table entry index
Dim nOff1 As Long                                                                 'Machine code buffer offset 1
Dim nOff2 As Long                                                                 'Machine code buffer offset 2

If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
        nEntry = nEntry + 1

        If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
            aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
            Exit Sub                                                                        'Bail
        ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
            Exit Sub                                                                        'Bail
        End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
End If

If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
End If

If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
End If
Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function


'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
zIdx = UBound(sc_aSubData)
Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
        If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
            If Not bAdd Then                                                                'If we're searching not adding
                Exit Function                                                                 'Found
            End If
        ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
            If bAdd Then                                                                    'If we're adding
                Exit Function                                                                 'Re-use it
            End If
        End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
Loop

If Not bAdd Then
Debug.Assert False                                                                  'hWnd not found, programmer error
End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
zSetTrue = True
bValue = True
End Function

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
Dim hmod As Long
Dim bLibLoaded As Boolean

hmod = GetModuleHandleA(sModule)

If hmod = 0 Then
    hmod = LoadLibraryA(sModule)
    If hmod Then
        bLibLoaded = True
    End If
End If

If hmod Then
    If GetProcAddress(hmod, sFunction) Then
        IsFunctionExported = True
    End If
End If

If bLibLoaded Then
    Call FreeLibrary(hmod)
End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
Dim tme As TRACKMOUSEEVENT_STRUCT

If bTrack Then
    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE
        .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
        Call TrackMouseEvent(tme)
    Else
        Call TrackMouseEventComCtl(tme)
    End If
End If
End Sub

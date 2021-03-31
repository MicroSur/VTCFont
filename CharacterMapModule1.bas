Attribute VB_Name = "Module1"
Option Explicit
'From MSDN article
'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64
Public Type RECT   '  16  Bytes
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
'Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
'Private Declare Function InvertRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As POINTAPI) As Long
'Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Public Declare Function ExtCreatePen Lib "gdi32.dll" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, ByRef lplb As LOGBRUSH, ByVal dwStyleCount As Long, ByRef lpStyle As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Public Const PS_GEOMETRIC As Long = &H10000
Public Const PS_ENDCAP_SQUARE As Long = &H100
Public Const PS_SOLID As Long = 0

Declare Function DrawFocusRect& Lib "user32" (ByVal hdc As Long, _
                                              lprect As RECT)



Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

Declare Function EnumFontFamilies Lib "gdi32" Alias _
                                  "EnumFontFamiliesA" _
                                  (ByVal hdc As Long, ByVal lpszFamily As String, _
                                   ByVal lpEnumFontFamProc As Long, LParam As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                         ByVal hdc As Long) As Long
                                         
'Private m_selectedsquare


Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
                         ByVal FontType As Long, LParam As ListBox) As Long
Dim FaceName As String
'Dim FullName As String

FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
LParam.AddItem left$(FaceName, InStr(FaceName, vbNullChar) - 1)
EnumFontFamProc = 1
End Function

Sub FillListWithFonts(LB As ComboBox)    'ListBox)
Dim hdc As Long
LB.Clear
hdc = GetDC(LB.hWnd)
EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, LB
ReleaseDC LB.hWnd, hdc
End Sub


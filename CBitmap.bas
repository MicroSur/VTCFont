Attribute VB_Name = "CBitmap"
'Autor: ALKO
'e-mail: alfred.koppold@freenet.de

Option Explicit
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

' Constants
Private Const SRCCOPY = &HCC0020
Private Const BI_RGB = 0&
Private Const CBM_INIT = &H4
Private Const DIB_RGB_COLORS = 0
' Types
Public Type RGBTriple
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

'Private Type Bitmap
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type
Private Type BITMAPINFOHEADER   '40 bytes
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

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type

Private Type BITMAPINFO_1
    bmiHeader As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type
Private Type BITMAPINFO_2
    bmiHeader As BITMAPINFOHEADER
    bmiColors(3) As RGBQUAD
End Type
Private Type BITMAPINFO_4
    bmiHeader As BITMAPINFOHEADER
    bmiColors(15) As RGBQUAD
End Type
Private Type BITMAPINFO_8
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type
Private Type BITMAPINFO_16
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type BITMAPINFO_24
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type BITMAPINFO_24a
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTriple
End Type

' Functions

'Private Declare Function GetDIBits1 Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_1, ByVal wUsage As Long) As Long

Public Declare Function SetPixel Lib "gdi32" _
                                 (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDIBitmap_1 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_1, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_2 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_2, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_4 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_4, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_8 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_8, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_16 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_16, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24a Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24a, ByVal wUsage As Long) As Long
'Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Public Declare Function RectangleX Lib "gdi32" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


'header
Private bm1 As BITMAPINFO_1
Private bm2 As BITMAPINFO_2
Private bm4 As BITMAPINFO_4
Private bm8 As BITMAPINFO_8
Private bm16 As BITMAPINFO_16
Private bm24 As BITMAPINFO_24
Private bm24a As BITMAPINFO_24a
'bitmap handle.
Private hBmp As Long

Private Type ScTw
    Width As Long
    Height As Long
End Type

'Jim Deutch
'MS Dev MVP
'picturebox must have AutoRedraw = True, ScaleMode = vbPixel

'Private Type BITMAPFILEHEADER    '14 bytes
'    bfType As Integer
'    bfSize As Long
'    bfReserved1 As Integer
'    bfReserved2 As Integer
'    bfOffBits As Long
'End Type






Public Sub InitColorTable_1(Optional Sorting As Integer = 1)
Dim Fb1 As Byte
Dim Fb2 As Byte
Select Case Sorting
Case 0
    Fb1 = 255
    Fb2 = 0
Case 1
    Fb1 = 0
    Fb2 = 255
End Select
bm1.bmiColors(0).rgbRed = Fb1
bm1.bmiColors(0).rgbGreen = Fb1
bm1.bmiColors(0).rgbBlue = Fb1
bm1.bmiColors(0).rgbReserved = 0
bm1.bmiColors(1).rgbRed = Fb2
bm1.bmiColors(1).rgbGreen = Fb2
bm1.bmiColors(1).rgbBlue = Fb2
bm1.bmiColors(1).rgbReserved = 0

End Sub
Public Sub InitColorTable_1Palette(Palettenbyte() As Byte)
If UBound(Palettenbyte) = 5 Then
    bm1.bmiColors(0).rgbRed = Palettenbyte(0)
    bm1.bmiColors(0).rgbGreen = Palettenbyte(1)
    bm1.bmiColors(0).rgbBlue = Palettenbyte(2)
    bm1.bmiColors(0).rgbReserved = 0
    bm1.bmiColors(1).rgbRed = Palettenbyte(3)
    bm1.bmiColors(1).rgbGreen = Palettenbyte(4)
    bm1.bmiColors(1).rgbBlue = Palettenbyte(5)
    bm1.bmiColors(1).rgbReserved = 0
Else
    InitColorTable_1
End If
End Sub

Public Sub InitColorTable_8(ByteArray() As Byte)
'Construct the palette
'==================================================
Dim Palette8() As RGBTriple
ReDim Palette8(255)
CopyMemory Palette8(0), ByteArray(0), UBound(ByteArray) + 1
Dim nCount As Long
On Error Resume Next
'Create Palette
For nCount = 0 To 255
    bm8.bmiColors(nCount).rgbBlue = Palette8(nCount).Blue
    bm8.bmiColors(nCount).rgbGreen = Palette8(nCount).Green
    bm8.bmiColors(nCount).rgbRed = Palette8(nCount).Red
    bm8.bmiColors(nCount).rgbReserved = 0
Next nCount
End Sub
Public Sub InitColorTable_4(ByteArray() As Byte)
Dim Palette4() As RGBTriple
ReDim Palette4(15)
CopyMemory Palette4(0), ByteArray(0), UBound(ByteArray) + 1

Dim i As Integer
' Create a color table
For i = 0 To 15
    bm4.bmiColors(i).rgbRed = Palette4(i).Red
    bm4.bmiColors(i).rgbGreen = Palette4(i).Green
    bm4.bmiColors(i).rgbBlue = Palette4(i).Blue
    bm4.bmiColors(i).rgbReserved = 0
Next i

End Sub


Public Sub CreateBitmap_1(ByteArray() As Byte, BMPWidth As Long, BMPHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
' Create a 1bit Bitmap
Dim hdc As Long
With bm1.bmiHeader
    .biSize = Len(bm1.bmiHeader)
    .biWidth = BMPWidth
    If Orientation = 0 Then
        .biHeight = BMPHeight                    'Bitmap Height, bitmap is top down.
    Else
        .biHeight = -BMPHeight
    End If
    .biPlanes = 1
    .biBitCount = 1
    .biCompression = BI_RGB
    .biSizeImage = 0
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
    .biClrUsed = Colorused
    .biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_1(hdc, bm1.bmiHeader, CBM_INIT, ByteArray(0), bm1, DIB_RGB_COLORS)
End Sub
Public Sub CreateBitmap_2(ByteArray() As Byte, BMPWidth As Long, BMPHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
' Create a 2bit Bitmap
Dim hdc As Long
With bm1.bmiHeader
    .biSize = Len(bm1.bmiHeader)
    .biWidth = BMPWidth
    If Orientation = 0 Then
        .biHeight = BMPHeight                    'Bitmap Height, bitmap is top down.
    Else
        .biHeight = -BMPHeight
    End If
    .biPlanes = 1
    .biBitCount = 2
    .biCompression = BI_RGB
    .biSizeImage = 0
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
    .biClrUsed = Colorused
    .biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_2(hdc, bm2.bmiHeader, CBM_INIT, ByteArray(0), bm2, DIB_RGB_COLORS)
End Sub

Public Sub CreateBitmap_4(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
' Creates a device independent bitmap
' from the pixel data in Data().
Dim hdc As Long
With bm4.bmiHeader
    .biSize = Len(bm1.bmiHeader)
    .biWidth = PicWidth
    If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
    Else
        .biHeight = -PicHeight
    End If
    .biPlanes = 1
    .biBitCount = 4
    .biCompression = BI_RGB
    .biSizeImage = 0
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
    .biClrUsed = Colorused
    .biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_4(hdc, bm4.bmiHeader, CBM_INIT, ByteArray(0), bm4, DIB_RGB_COLORS)
End Sub

Public Sub CreateBitmap_8(bitmaparray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
' Creates a device independent bitmap
' from the pixel data in BitmapArry().
Dim hdc As Long
With bm8.bmiHeader
    .biSize = Len(bm8.bmiHeader)
    .biWidth = PicWidth
    If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
    Else
        .biHeight = -PicHeight
    End If
    .biPlanes = 1
    .biBitCount = 8
    .biCompression = BI_RGB
    .biSizeImage = 0
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
    .biClrUsed = Colorused
    .biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_8(hdc, bm8.bmiHeader, CBM_INIT, bitmaparray(0), bm8, DIB_RGB_COLORS)
End Sub

Public Sub DrawBitmap(PicWidth As Long, PicHeight As Long, PicObject As Object, Scalierung As Boolean, Optional X As Long = 0, Optional Y As Long = 0, Optional DrawToBG As Boolean = False)
Dim cDC As Long
Dim a As Long
Dim b As Long
Dim Übergabe As ScTw
Dim realheight As Long
Dim realwidth As Long
PicObject.Cls
If TypeOf PicObject Is Form Then
'change ScaleMode direct
Else
    b = PicObject.Parent.ScaleMode
    PicObject.Parent.ScaleMode = 1
End If

a = PicObject.ScaleMode
PicObject.ScaleMode = 1
Select Case Scalierung
Case True
    Übergabe = PixelToTwips(PicWidth, PicHeight)
    If DrawToBG = False Then
        PicObject.Height = Übergabe.Height
        PicObject.Width = Übergabe.Width
    End If
Case False
End Select
If DrawToBG = False Then
    If PicObject.Height <> PicObject.ScaleHeight Then    'with Boarders
        Übergabe = Twipstopixel(PicObject.Width, PicObject.Height)
        realheight = Übergabe.Height
        realwidth = Übergabe.Width
        PicObject.Height = PicObject.Height + (PicObject.Height - PicObject.ScaleHeight)
        PicObject.Width = PicObject.Width + (PicObject.Width - PicObject.ScaleWidth)
    Else
        PicObject.ScaleMode = 3
        realheight = PicObject.ScaleHeight
        realwidth = PicObject.ScaleWidth
    End If
Else
    realheight = Übergabe.Height
    realwidth = Übergabe.Width
    PicHeight = realheight
    PicWidth = realwidth
End If
If hBmp Then
    cDC = CreateCompatibleDC(PicObject.hdc)
    SelectObject cDC, hBmp
    Call StretchBlt(PicObject.hdc, X, Y, realwidth, realheight, cDC, 0, 0, PicWidth, PicHeight, SRCCOPY)
    DeleteDC cDC
    DeleteObject hBmp
    hBmp = 0
End If
If TypeOf PicObject Is Form Then
'change ScaleMode direct
Else
    PicObject.Parent.ScaleMode = b
End If
PicObject.ScaleMode = a
PicObject.Picture = PicObject.Image
End Sub






Public Sub CreateBitmap_24(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional ThreeToOrToFour As Integer = 0)
' Creates a device independent bitmap
' from the pixel data in BitmapArray().

Dim hdc As Long
Dim Bits() As RGBQUAD
Dim BitsA() As RGBTriple
Select Case ThreeToOrToFour
Case 0
    ReDim Bits((UBound(ByteArray) / 4) - 1)
    CopyMemory Bits(0), ByteArray(0), UBound(ByteArray)
    With bm24.bmiHeader
        .biSize = Len(bm24.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth        'Bitmap Width
        If Orientation = 0 Then
            .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -PicHeight
        End If
        .biBitCount = 32                        '32 bit alignment
        .biPlanes = 1                           'Single plane
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With

Case 1
    ReDim BitsA((UBound(ByteArray) / 3) - 1)
    CopyMemory BitsA(0), ByteArray(0), UBound(ByteArray)

    With bm24a.bmiHeader
        .biSize = Len(bm24.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth        'Bitmap Width
        If Orientation = 0 Then
            .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -PicHeight
        End If
        .biBitCount = 24                        '24 bit alignment
        .biPlanes = 1                           'Single plane
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With
End Select
' Get the DC.
hdc = GetDC(0)
Select Case ThreeToOrToFour
Case 0
    hBmp = CreateDIBitmap_24(hdc, bm24.bmiHeader, CBM_INIT, Bits(0), bm24, DIB_RGB_COLORS)
Case 1
    hBmp = CreateDIBitmap_24a(hdc, bm24a.bmiHeader, CBM_INIT, BitsA(0), bm24a, DIB_RGB_COLORS)
End Select
End Sub
Public Sub CreateBitmap_16(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
' Creates a device independent bitmap
' from the pixel data in BitmapArray().
Dim hdc As Long

With bm16.bmiHeader
    .biSize = Len(bm16.bmiHeader)        'SizeOf Struct
    .biWidth = PicWidth                       'Bitmap Width
    If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
    Else
        .biHeight = -PicHeight
    End If
    .biPlanes = 1                           'Single plane
    .biBitCount = 16                        '32 bit alignment
    .biCompression = BI_RGB                 'No Compression
    .biSizeImage = 0                        'Default
    .biXPelsPerMeter = 0                    'Default
    .biYPelsPerMeter = 0                    'Default
    .biClrUsed = 0                          'Default
    .biClrImportant = 0                     'Default
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_16(hdc, bm16.bmiHeader, CBM_INIT, ByteArray(0), bm16, DIB_RGB_COLORS)
End Sub

Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
'Dim XWert1 As Long
'Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
PixelToTwips.Width = xwert * ux
uy = Screen.TwipsPerPixelY
PixelToTwips.Height = ywert * uy
End Function



Public Function Twipstopixel(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
'Dim XWert1 As Long
'Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
Twipstopixel.Width = xwert / ux
uy = Screen.TwipsPerPixelY
Twipstopixel.Height = ywert / uy
End Function

Public Function InitColorTable_Grey(BitDepth As Integer, Optional To8Bit As Boolean = False) As Byte()
Dim CurLevel As Integer
Dim Übergabe() As Byte
Dim n As Long
Dim LevelDiff As Byte
Dim Tbl() As RGBQUAD
Dim Table3() As RGBTriple
Erase bm8.bmiColors
If BitDepth <> 16 Then
    ReDim Tbl(2 ^ BitDepth - 1)
    ReDim Table3(2 ^ BitDepth - 1)
Else
    ReDim Tbl(255)
    ReDim Table3(255)
End If
LevelDiff = 255 / UBound(Tbl)

For n = 0 To UBound(Tbl)
    With Tbl(n)
        .rgbRed = CurLevel
        .rgbGreen = CurLevel
        .rgbBlue = CurLevel
    End With
    With Table3(n)
        .Red = CurLevel
        .Green = CurLevel
        .Blue = CurLevel
    End With
    CurLevel = CurLevel + LevelDiff

Next n
Select Case BitDepth
Case 1
    If To8Bit = True Then
        CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 8
    End If
Case 2
    CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 16
Case 4
    If To8Bit = True Then
        CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 64
    Else
        CopyMemory ByVal VarPtr(bm4.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 64
    End If
Case 8
    CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 1024
End Select
ReDim Übergabe(((UBound(Table3) + 1) * 3) - 1)
CopyMemory Übergabe(0), ByVal VarPtr(Table3(0).Red), ((UBound(Table3) + 1) * 3)
InitColorTable_Grey = Übergabe
End Function


'Public Sub SaveBMP1bit(pic As PictureBox, SaveFileName As String)
'Dim SaveBitmapInfo_1 As BITMAPINFO_1
'Dim SaveFileHeader As BITMAPFILEHEADER
'Dim SaveBits() As Byte
'Dim BufferSize As Long
'Dim fNum As Long
'Dim Retval As Long
'Dim nLen As Long
'Const BitsPixel = 1
'
''size a buffer for the pixel data
'BufferSize = ((pic.ScaleWidth / 8 + 3) And &HFFFC) * pic.ScaleHeight
'ReDim SaveBits(0 To BufferSize - 1)
''fill the header info for the save copy
'With SaveBitmapInfo_1.bmiHeader
'    .biSize = 40
'    .biWidth = pic.ScaleWidth
'    .biHeight = pic.ScaleHeight
'    .biPlanes = 1
'    .biBitCount = BitsPixel
'    .biCompression = 0
'    .biClrUsed = 0
'    .biClrImportant = 0
'    .biSizeImage = BufferSize
'End With
'nLen = Len(SaveBitmapInfo_1)
''get the bitmap from the picturebox
'Retval = GetDIBits1(pic.hdc, pic.Image, 0, _
'                    SaveBitmapInfo_1.bmiHeader.biHeight, SaveBits(0), SaveBitmapInfo_1, DIB_RGB_COLORS)
'
'' create a header for the save file
'With SaveFileHeader
'    .bfType = &H4D42
'    .bfSize = Len(SaveFileHeader) + nLen + BufferSize
'    .bfOffBits = Len(SaveFileHeader) + nLen
'End With
'
'' save it to disk
'fNum = FreeFile
'
'Open SaveFileName For Binary As fNum
'Put fNum, , SaveFileHeader
'Put fNum, , SaveBitmapInfo_1
'Put fNum, , SaveBits()
'Close fNum
'
'End Sub


Public Function SetBitmapData(ByVal hdc As Long, _
                              ByVal Width As Long, _
                              ByVal Height As Long, _
                              ByVal Value As Long, _
                              Optional ByVal BitCount As Integer = 32, _
                              Optional ByVal ReSize As Double = 1) As Boolean

Dim bi As BITMAPINFO, mhDC As Long, bitsPtr As Long, hDIB As Long
Dim old_bmp As Long, Ret As Long

mhDC = CreateCompatibleDC(0)
If mhDC <> 0 Then
    With bi.bmiHeader
        .biSize = Len(bi.bmiHeader)
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biBitCount = BitCount
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine(.biWidth, .biBitCount) * .biHeight
    End With

    hDIB = CreateDIBSection(mhDC, bi, DIB_RGB_COLORS, bitsPtr, 0, 0)

    If hDIB <> 0 Then
        old_bmp = SelectObject(mhDC, hDIB)
        Ret = SetDIBits(mhDC, hDIB, 0, bi.bmiHeader.biHeight, ByVal Value, bi, DIB_RGB_COLORS)

        If ReSize <> 1 Then
            Ret = StretchBlt(hdc, 0, 0, Width * ReSize, Height * ReSize, mhDC, 0, 0, Width, Height, SRCCOPY)
        Else
            Ret = BitBlt(hdc, 0, 0, Width, Height, mhDC, 0, 0, SRCCOPY)
        End If
    Else
        DeleteDC mhDC
        Exit Function
    End If
End If

DeleteObject hDIB
SelectObject mhDC, old_bmp
DeleteDC mhDC

SetBitmapData = Ret > 0
End Function

Public Function BytesPerScanLine(Width As Long, BitCount As Integer) As Long
BytesPerScanLine = (Width * BitCount)
If (BytesPerScanLine Mod 32 > 0) Then BytesPerScanLine = BytesPerScanLine + 32 - (BytesPerScanLine Mod 32)
BytesPerScanLine = BytesPerScanLine \ 8
End Function


'Extract the red, green, or blue value from an RGB() Long
Public Function ExtractR(ByVal currentColor As Long) As Integer
ExtractR = currentColor And 255
End Function

Public Function ExtractG(ByVal currentColor As Long) As Integer
ExtractG = (currentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal currentColor As Long) As Integer
ExtractB = (currentColor \ 65536) And 255
End Function
'This function will return the luminance value of an RGB triplet.  Note that the value will be in the [0,255] range instead
' of the usual [0,1.0] one.
Public Function getLuminance(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
Dim Max As Long, Min As Long
Max = Max3Int(r, g, b)
Min = Min3Int(r, g, b)
getLuminance = (Max + Min) \ 2
End Function
'Return the maximum of three integer values
Public Function Max3Int(rR As Long, rG As Long, rB As Long) As Long
If (rR > rG) Then
    If (rR > rB) Then
        Max3Int = rR
    Else
        Max3Int = rB
    End If
Else
    If (rB > rG) Then
        Max3Int = rB
    Else
        Max3Int = rG
    End If
End If
End Function
'Return the minimum of three integer values
Public Function Min3Int(rR As Long, rG As Long, rB As Long) As Long
If (rR < rG) Then
    If (rR < rB) Then
        Min3Int = rR
    Else
        Min3Int = rB
    End If
Else
    If (rB < rG) Then
        Min3Int = rB
    Else
        Min3Int = rG
    End If
End If
End Function


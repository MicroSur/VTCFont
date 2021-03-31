Attribute VB_Name = "mod_SaveBMP"
Option Explicit

Private Const DIB_RGB_COLORS As Long = 0
Private Const SRCCOPY As Long = &HCC0020
Private Const BI_RGB As Long = 0&

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

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

'Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, _
 ByVal Y As Long, _
 ByVal nWidth As Long, _
 ByVal nHeight As Long, _
 ByVal hSrcDC As Long, _
 ByVal XSrc As Long, _
 ByVal YSrc As Long, _
 ByVal dwRop As Long) As Long

Private Declare Function GetDIBits Lib "gdi32.dll" ( _
 ByVal aHDC As Long, _
 ByVal hBitmap As Long, _
 ByVal nStartScan As Long, _
 ByVal nNumScans As Long, _
 ByRef lpBits As Any, _
 ByRef lpbi As BITMAPINFO, _
 ByVal wUsage As Long) As Long

'Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function GetBitmapObject Lib "gdi32" Alias "GetObjectA" ( _
 ByVal hBitmap As Long, _
 ByVal cbBuffer As Long, _
 ByRef destBmp As Any) As Long

Public Sub SavePictureBW(ByVal ctrl As PictureBox, ByVal destfile As String)
Dim hdcMono As Long, hbmpMono As Long, hbmpOld As Long, dxBlt As Long, dyBlt As Long, success As Long
Dim numscans As Long, byteswide As Long, totalbytes As Long, lfilesize As Long
Dim bmpsrc As Bitmap ', bmpdst As Bitmap
Dim bInfo As BITMAPINFO
Dim bitmaparray() As Byte, fileheader() As Byte
Dim ff As Integer, by8 As Long

On Error GoTo modErr

'Object's scalemode must be Pixel.
dxBlt = ctrl.ScaleWidth
dyBlt = ctrl.ScaleHeight

'Create monochrome bitmap from control.
hdcMono = CreateCompatibleDC(0)
hbmpMono = CreateCompatibleBitmap(hdcMono, dxBlt, dyBlt)
success = GetBitmapObject(hbmpMono, Len(bmpsrc), bmpsrc)
hbmpOld = SelectObject(hdcMono, hbmpMono)
success = BitBlt(hdcMono, 0, 0, dxBlt, dyBlt, ctrl.hdc, 0, 0, SRCCOPY)

'Calculate array size needed for bitmap bits (dword aligned)
numscans = dyBlt
by8 = dxBlt / 8
If (dxBlt Mod 8) = 0 And (by8 Mod 4) = 0 Then
    byteswide = by8
Else
    byteswide = (Int(by8) + 4) - (Int(by8) Mod 4)
End If
totalbytes = numscans * byteswide
ReDim bitmaparray(1 To totalbytes)

'Set BITMAPINFO values to pass to GetDIBits function.
With bInfo
    .bmiHeader.biSize = Len(.bmiHeader)
    .bmiHeader.biWidth = bmpsrc.bmWidth
    .bmiHeader.biHeight = bmpsrc.bmHeight
    .bmiHeader.biPlanes = bmpsrc.bmPlanes
    .bmiHeader.biBitCount = bmpsrc.bmBitsPixel
    .bmiHeader.biCompression = BI_RGB
End With

success = GetDIBits(hdcMono, ctrl.Image, 0, numscans, bitmaparray(1), bInfo, DIB_RGB_COLORS)

'bitmaparray should now contain bitmap bit data. Now create bitmap file header.
ReDim fileheader(1 To &H3E)
fileheader(1) = &H42    'B
fileheader(2) = &H4D    'M
lfilesize = UBound(fileheader) + UBound(bitmaparray)
fileheader(3) = lfilesize And 255
fileheader(4) = (lfilesize \ 256) And 255
fileheader(5) = (lfilesize \ 65536) And 255
fileheader(6) = (lfilesize \ 16777216) And 255
fileheader(11) = &H3E    'offset
fileheader(15) = &H28    'size of bitmapinfoheader
fileheader(19) = dxBlt And 255
fileheader(20) = (dxBlt \ 256) And 255
fileheader(21) = (dxBlt \ 65536) And 255
fileheader(22) = (dxBlt \ 16777216) And 255
fileheader(23) = dyBlt And 255
fileheader(24) = (dyBlt \ 256) And 255
fileheader(25) = (dyBlt \ 65536) And 255
fileheader(26) = (dyBlt \ 16777216) And 255
fileheader(27) = 1
fileheader(29) = 1
fileheader(35) = UBound(bitmaparray) And 255
fileheader(36) = (UBound(bitmaparray) \ 256) And 255
fileheader(37) = (UBound(bitmaparray) \ 65536) And 255
fileheader(38) = (UBound(bitmaparray) \ 16777216) And 255
fileheader(47) = 2
fileheader(51) = 2
fileheader(59) = &HFF
fileheader(60) = &HFF
fileheader(61) = &HFF


If FileExists(destfile) Then Kill destfile
ff = FreeFile
Open destfile For Binary Access Write As #ff
Put #ff, , fileheader
Put #ff, , bitmaparray
Close #ff

' Clean up
Call SelectObject(hdcMono, hbmpOld)
Call DeleteDC(hdcMono)
Call DeleteObject(hbmpMono)

Exit Sub

modErr:
MsgBox Err.Description & ": SavePictureBW"
End Sub



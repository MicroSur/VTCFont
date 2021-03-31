Attribute VB_Name = "modBmp"
'Jim Deutch
'MS Dev MVP
'picturebox must have AutoRedraw = True, ScaleMode = vbPixel

Private Type BITMAPFILEHEADER    '14 bytes
   bfType As Integer
   bfSize As Long
   bfReserved1 As Integer
   bfReserved2 As Integer
   bfOffBits As Long
End Type

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

Private Type BITMAPINFO_1
   bmiHeader As BITMAPINFOHEADER
   bmiColors(1) As RGBQUAD
End Type

Private Const PIXEL As Integer = 3
Private Const DIB_RGB_COLORS As Long = 0
Private Const PALVERSION = &H300

Private Declare Function GetDIBits1 Lib "gdi32" Alias "GetDIBits" _
  (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO_1, ByVal wUsage As Long) As Long


Public Sub SaveBMP1bit(pic As PictureBox, SaveFileName As String)
Dim SaveBitmapInfo_1 As BITMAPINFO_1
Dim SaveFileHeader As BITMAPFILEHEADER
Dim SaveBits() As Byte
Dim BufferSize As Long
Dim fNum As Long
Dim Retval As Long
Dim nLen As Long
Const BitsPixel = 1

    'size a buffer for the pixel data
    BufferSize = ((pic.ScaleWidth / 8 + 3) And &HFFFC) * pic.ScaleHeight
    ReDim SaveBits(0 To BufferSize - 1)
    'fill the header info for the save copy
    With SaveBitmapInfo_1.bmiHeader
        .biSize = 40
        .biWidth = pic.ScaleWidth
        .biHeight = pic.ScaleHeight
        .biPlanes = 1
        .biBitCount = BitsPixel
        .biCompression = 0
        .biClrUsed = 0
        .biClrImportant = 0
        .biSizeImage = BufferSize
    End With
    nLen = Len(SaveBitmapInfo_1)
    'get the bitmap from the picturebox
    Retval = GetDIBits1(pic.hdc, pic.Image, 0, _
        SaveBitmapInfo_1.bmiHeader.biHeight, SaveBits(0), SaveBitmapInfo_1, DIB_RGB_COLORS)

    ' create a header for the save file
    With SaveFileHeader
       .bfType = &H4D42
       .bfSize = Len(SaveFileHeader) + nLen + BufferSize
       .bfOffBits = Len(SaveFileHeader) + nLen
    End With

    ' save it to disk
    fNum = FreeFile

    Open SaveFileName For Binary As fNum
    Put fNum, , SaveFileHeader
    Put fNum, , SaveBitmapInfo_1
    Put fNum, , SaveBits()
    Close fNum

End Sub

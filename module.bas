Attribute VB_Name = "module"
Option Explicit

Public uMagic As Long '408376 &H63B38 'old

Public BMPFileName As String 'save name of pic when import for export
Public QwickResizeWidth As Integer
'''''''''Params in patches for Pather
Public OldPatchIndex As Integer

Public arrByteCount() As Integer    'сколько байт у параметра
Public arrMulByte() As Integer    'степень множителя 10 у параметра
Public arrParamDescr() As String    'название параметра, в зависимости от языка
Public arrBytesAll() As String    'in hex
Public arrAddrAll() As String    ' in hex
Public arrParamsCount() As Integer
Public arrBytePos() As Integer    'позиция байта 1 Lo -> Hi

Public carrParamBytes() As Variant    'collection
Public carrParamAddrs() As Variant
Public carrParamBytePos() As Variant
Public iParamNum As Integer

''''''''''''''''''''''''''''
Public arrPtcAdr() As Long    'hex address of changes
Public arrPtcNew() As Byte
Public arrPtcFileName() As String
Public arrPtcName() As String
Public arrPatchFormat() As Byte    '1 vtc, 2 nfe, 3 crk
Public arrPtcComment() As String 'in line comments
'''''''

Public MouseUpBug As Boolean

Public fFileOpen As Boolean
Public Hardtext As String    'store fo current file open
Public Language As String
Public lngFileName As String    'with path
Public OldlngFileName As String    'with path

Public old_SBValue_efsVertical As Long 'last vert slider position
Public reloadFW_flag As Boolean 'for button reload

Public lngFileNameOnly As String
Public ArrMsg() As String    'for msgbox
Public bFileIn As Integer    '#file fw
Public lngBytes As Long     'len of file
Public FileNameFW As String    'current open FW file
Public LastOpenedFW As String    'fo reload on first start
Public FileTitle As String

Public Magnify As Boolean    'x3 realpic
Public MagnifyBy As Integer

Public FindFilePath As String
Public Const mySpace = " "
Public ChangesIndArr() As Boolean    'list changes indexes
Public NoIniFlag As Boolean
Public iniFileName As String
Public InvertMouseB As Integer
Public fPicDithered As Integer    'pic load grayscale method
Public CheckCharSizeFlag As Boolean    'check font size when paste and write
Public AllWordsInLineFlag As Boolean    'if true - in vert line, else in square

Public ImArr() As StdPicture    'load from file


Public LastPath As String
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Const LOCALE_USER_DEFAULT& = &H400
Private Const LOCALE_SDECIMAL& = &HE
'Private Const LOCALE_STHOUSAND& = &HF
Private Declare Function GetLocaleInfo& Lib "kernel32" Alias _
                                        "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
                                                          ByVal lpLCData As String, ByVal cchData As Long)

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function MoveWindow& Lib "user32" (ByVal hWnd As Long, _
                                                  ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, ByVal bRepaint As Long)

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_READONLY = &H1

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type
Public Declare Function FindFirstFile _
                         Lib "kernel32" Alias "FindFirstFileA" ( _
                             ByVal lpFileName As String, _
                             lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile _
                         Lib "kernel32" Alias "FindNextFileA" ( _
                             ByVal hFindFile As Long, _
                             lpFindFileData As WIN32_FIND_DATA) As Long


Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

Public Declare Function MultiByteToWideCharA Lib "kernel32.dll" Alias "MultiByteToWideChar" ( _
                                             ByVal CodePage As Long, ByVal dwFlags As Long, _
                                             ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, _
                                             ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Const CP_UTF8 = 65001

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hWnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long
'Private Const SW_HIDE As Long = 0
Public Const SW_SHOWNORMAL As Long = 1
'Private Const SW_SHOWMAXIMIZED As Long = 3
'Private Const SW_SHOWMINIMIZED As Long = 2
'Private Const SW_HIDE = 0
'Private Const SW_MAXIMIZE = 3
'Private Const SW_MINIMIZE = 6
'Private Const SW_RESTORE = 9
'Private Const SW_SHOW = 5
'Private Const SW_SHOWDEFAULT = 10
'Private Const SW_SHOWMAXIMIZED = 3
'Private Const SW_SHOWMINIMIZED = 2
'Private Const SW_SHOWMINNOACTIVE = 7
'Private Const SW_SHOWNA = 8
'Private Const SW_SHOWNOACTIVATE = 4
'Private Const SW_SHOWNORMAL = 1
Public Declare Function FindWindow Lib "user32" _
Alias "FindWindowA" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long
'Public Declare Function ShowWindow Lib "user32" _
(ByVal hwnd As Long, ByVal nCmdShow As Integer) As Integer
Public Declare Function SetForegroundWindow Lib "user32" ( _
ByVal hWnd As Long) As Long

Public Declare Sub CopyMemoryString Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, Source As Any, ByVal length As Long)



Public Function GetChangesCount() As Integer
Dim i As Integer
On Error GoTo modErr

For i = 0 To UBound(ChangesIndArr)
    If ChangesIndArr(i) Then GetChangesCount = GetChangesCount + 1
Next i

'''
Exit Function
modErr:
MsgBox Err.Description & ": GetChangesCount()"
End Function


Public Function Bin2Dec(Num As String) As Double
Dim n As Long, a As Integer, X As String
n = Len(Num) - 1
a = n
Do While n > -1
    X = Mid$(Num, ((a + 1) - n), 1)
    Bin2Dec = IIf((X = "1"), Bin2Dec + (2 ^ (n)), Bin2Dec)
    n = n - 1
Loop
End Function

Public Function DecimalSeparator() As String
Dim r As Long, s As String
s = String(10, "a")
r = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, s, 10)
DecimalSeparator = left$(s, r)
End Function
    
Public Function Bin2Hex(ByVal BinaryString As String) As String
Dim X As Integer
Const BinValues = "*0000*0001*0010*0011" & _
      "*0100*0101*0110*0111" & _
      "*1000*1001*1010*1011" & _
      "*1100*1101*1110*1111*"
Const HexValues = "0123456789ABCDEF"
If BinaryString Like "*[!01]*" Then
    Bin2Hex = vbNullString
Else
    BinaryString = String$((4 - Len(BinaryString) _
                            Mod 4) Mod 4, "0") & BinaryString
    For X = 1 To Len(BinaryString) - 3 Step 4
        Bin2Hex = Bin2Hex & Mid$(HexValues, _
                                 (4 + InStr(BinValues, "*" & _
                                                       Mid$(BinaryString, X, 4) & "*")) \ 5, 1)
    Next
End If
End Function

Public Function RemoveXtraSpaces(strVal As String) As String
Do While InStr(1, strVal, vbCrLf)
    strVal = Replace(strVal, vbCrLf, "")
Loop
Do While InStr(1, strVal, " ")
    strVal = Replace(strVal, " ", "")
Loop
RemoveXtraSpaces = StrConv(strVal, vbUpperCase)
End Function

Public Function Hex2Bin(ByVal sHex As String, Optional ByteSep As String) As String
Dim arNibs() As String, i As Long, j As Long

arNibs = Split("0000,0001,0010,0011,0100,0101,0110,0111," & _
               "1000,1001,1010,1011,1100,1101,1110,1111", ",")
'-- convert each Hex digit to nibble and add byte separator if required
For i = 1 To Len(sHex)
    If sHex Like "*[!(0-9)ABCDEFabcdef]*" Then
        Hex2Bin = "Error input"
        Exit Function
    End If
    j = CLng("&H" & Mid$(sHex, i, 1))
    Hex2Bin = Hex2Bin & ByteSep & arNibs(j)
Next
'    If Len(ByteSep) Then '-- remove extra leading separator
'        Hex2Bin = Mid$(Hex2Bin, Len(ByteSep) + 1)
'    Else '-- remove leading 0's
'        i = InStr(Hex2Bin, "1")
'        If i > 1 Then Hex2Bin = Mid$(Hex2Bin, i)
'    End If
End Function
Public Function dec2binByte(ByVal DecNum As Long) As String

Select Case DecNum
Case Is > 1
    dec2binByte = dec2binByte(DecNum \ 2) & CStr(DecNum And 1)
Case Is > 0
    dec2binByte = "1"
Case Else
    dec2binByte = "0"
End Select

dec2binByte = right$("00000000" & dec2binByte, 8)
End Function
Public Function dec2bin(mynum As Long, Optional NBytes As Integer = 8) As String
Dim s As String
Dim loopcounter As Integer

Do
    If (mynum And 2 ^ loopcounter) = 2 ^ loopcounter Then
        dec2bin = "1" & dec2bin
    Else
        dec2bin = "0" & dec2bin
    End If
    loopcounter = loopcounter + 1
Loop Until 2 ^ loopcounter > mynum

s = Space$(NBytes)
s = Replace(s, " ", "0")
dec2bin = right$(s & dec2bin, NBytes)

End Function
'Public Function Dec2Hex(ByVal strDec As String) As String
'Dim mybyte(0 To 19) As Byte
'Dim lp As Long
'
'If strDec = 0 Then Dec2Hex = "00": Exit Function
'CopyMemory mybyte(0), ByVal VarPtr(CDec(strDec)), 16
'
'' Quick reorganise so we can then just step through the entire thing in one loop
'For lp = 7 To 4 Step -1
'    mybyte(12 + lp) = mybyte(lp)
'Next
'
'' Build the hex string
'For lp = 19 To 8 Step -1
'    If (Not Len(Dec2Hex) And mybyte(lp) <> 0) Or Len(Dec2Hex) Then
'        Dec2Hex = Dec2Hex & right("0" & Hex(mybyte(lp)), 2)    'IIf(Len(Dec2Hex), "00", "0"))
'    End If
'Next
'End Function

Public Function GetNameExt(ByVal sThePathAndName As String) As String
' Return a string containing the file name from a fully qualified file name. Larry
' If no path then return the file name anyhow
Dim sTemp1 As String                        'temporary string
Dim sTemp2 As String
sTemp1 = Trim$(sThePathAndName)             'save it
sTemp2 = GetPathFromPathAndName(sTemp1)     'get path
If InStrB(sThePathAndName, sTemp2) <> 0 Then
    sTemp1 = Mid$(sThePathAndName, Len(sTemp2) + 1)
End If
GetNameExt = sTemp1             'now contains just file name
End Function
Public Function GetName(ByVal sThePathAndName As String) As String
' Return a string containing the file name from a fully qualified file name. Larry
' If no path then return the file name anyhow
Dim Ret As Integer
Dim sTemp1 As String                        'temporary string
Dim sTemp2 As String
sTemp1 = Trim$(sThePathAndName)             'save it
sTemp2 = GetNameExt(sTemp1)     'get name+ext
Ret = InStrRev(sTemp2, ".")
sTemp1 = left$(sTemp2, Ret - 1)
GetName = sTemp1             'now contains just file name
End Function
Public Function FileExists(sfile As String) As Boolean
Dim tFnd As WIN32_FIND_DATA
If Len(sfile) = 0 Then Exit Function
'sfile = "C:\EvicVTCFont\BIN\Wismec_30.06.16\Reuleaux_RX200S_V4.10.bin"
FileExists = (FindFirstFile(sfile, tFnd) <> -1)
'MsgBox FindFirstFile(sfile, tFnd)
End Function
Public Function TrimNull(ByVal Item As String) As String
Dim Pos As Integer

Pos = InStr(Item, vbNullChar)
If Pos = 1 Then
    Item = vbNullString
ElseIf Pos > 1 Then
    Item = left$(Item, Pos - 1)
End If

TrimNull = Item
End Function
Public Sub BubbleSort(ByRef pvarArray() As Long)
Dim i As Long
Dim iMin As Long
Dim iMax As Long
Dim varSwap As Long
Dim blnSwapped As Boolean

iMin = LBound(pvarArray)
iMax = UBound(pvarArray) - 1
Do
    blnSwapped = False
    For i = iMin To iMax
        If pvarArray(i) > pvarArray(i + 1) Then
            varSwap = pvarArray(i)
            pvarArray(i) = pvarArray(i + 1)
            pvarArray(i + 1) = varSwap
            blnSwapped = True
        End If
    Next
    iMax = iMax - 1
Loop Until Not blnSwapped
End Sub
Public Sub searchForFile(ByVal startPath As String, ByVal match As String)
'FindFilePath= путь к match(файлу)
Dim fPath As String, fname As String, fPathName As String
Dim hfind As Long, nameLen As Integer, matchLen As Integer
Dim WFD As WIN32_FIND_DATA
Dim found As Boolean


Const Dot1 = "."
Const Dot2 = ".."
FindFilePath = vbNullString

fPath = LCase$(startPath)
If right$(fPath, 1) <> "\" Then fPath = fPath & "\"

matchLen = Len(match)
match = LCase$(match)

'The first API call is to FindFirstFile.
'  Note that we get all files with a "*"
'  and not specify just the file extension
'  because we need to get the directories too.
hfind = FindFirstFile(fPath & "*", WFD)
found = (hfind > 0)

Do While found

'    DoEvents
'         If GetAsyncKeyState(vbKeyEscape) And &H1 = &H1 Then
''         If GetKeyState(vbKeyEscape) < 0 Then
'            StopSearching = True
'         End If

    fname = TrimNull(WFD.cFileName)
    fname = LCase$(fname)
    nameLen = Len(fname)
    fPathName = fPath & fname
    If fname = Dot1 Or fname = Dot2 Then

    ElseIf WFD.dwFileAttributes And _
           FILE_ATTRIBUTE_DIRECTORY Then

'frmmain.TextItemHid = fPathName

        Call searchForFile(fPathName, match)

    ElseIf matchLen = nameLen Then

        If InStrB(fname, match) <> 0 Then
            FindFilePath = fPathName

        End If

'Don't do anything if found is too short
'        ElseIf LCase$(Right$(fName, matchLen)) _
         '          = match Then
'            'We have an extension match
'            AddAnItem fPathName
'        Else
'            DoEvents
    End If

    If FindFilePath <> vbNullString Then Exit Do
'Subsequent API calls are to FindNextFile.
'        If StopSearching Then
'            Exit Function
'        Else
    found = FindNextFile(hfind, WFD)
'        End If

Loop

'Then close the findfile operation
FindClose hfind
'frmmain.TextItemHid = FindFilePath

End Sub

Public Function pLoadDialog(Optional DTitle As String, Optional ByRef FileTitle As String) As String
Dim cd As cCommonDialog
Dim sfile As String
Dim indir As String

If Len(LastPath) = 0 Then
    indir = App.Path
Else
    indir = LastPath
End If
'DoEvents
Set cd = New cCommonDialog
'sfile = FileName

If (cd.VBGetOpenFileName( _
    sfile, _
    FileTitle, _
    ReadOnly:=True, _
    HideReadOnly:=True, _
    Filter:="Binary files |*.bin|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    InitDir:=indir, _
    DlgTitle:=DTitle, _
    DefaultExt:="", _
    Owner:=frmmain.hWnd)) Then
    pLoadDialog = sfile
End If

Set cd = Nothing
End Function
Public Function fLoadDialog(Optional DTitle As String, Optional ByRef FileTitle As String) As String
Dim cd As cCommonDialog
Dim sfile As String
Dim indir As String

If Len(LastPath) = 0 Then
    indir = App.Path
Else
    indir = LastPath
End If
'DoEvents
Set cd = New cCommonDialog
'sfile = FileName

If (cd.VBGetOpenFileName( _
    sfile, _
    FileTitle, _
    ReadOnly:=True, _
    HideReadOnly:=True, _
    Filter:="Font files |*.ttf|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    InitDir:=indir, _
    DlgTitle:=DTitle, _
    DefaultExt:="", _
    Owner:=frmmain.hWnd)) Then
    fLoadDialog = sfile
End If

Set cd = Nothing
End Function
Public Function FWSaveDialog(sfile As String, Optional DTitle As String) As String
Dim cd As cCommonDialog
'Dim sfile As String
Dim indir As String

If Len(LastPath) = 0 Then
    indir = App.Path
Else
    indir = LastPath
End If

Set cd = New cCommonDialog
'sfile = FileName

If (cd.VBGetSaveFileName( _
    sfile, _
    Filter:="Binary files |*.bin|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    DlgTitle:=DTitle, _
    DefaultExt:="bin", _
    InitDir:=indir, _
    Owner:=frmmain.hWnd)) Then

    FWSaveDialog = sfile
End If

Set cd = Nothing
End Function

Public Function BMPSaveDialog(indir As String, Optional DTitle As String) As String
Dim cd As New cCommonDialog
Dim sfile As String

If Len(BMPFileName) <> 0 Then sfile = BMPFileName
If (cd.VBGetSaveFileName( _
    sfile, _
    Filter:="bmp (*.bmp)|*.bmp|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    DlgTitle:=DTitle, _
    DefaultExt:="bmp", _
    InitDir:=indir, _
    Owner:=frmmain.hWnd)) Then

    BMPSaveDialog = sfile
End If

Set cd = Nothing
End Function
Public Function ExportSaveDialog(indir As String, Optional DTitle As String) As String
Dim cd As New cCommonDialog
Dim sfile As String

'sFile = myFileName
If (cd.VBGetSaveFileName( _
    sfile, _
    Filter:="txt (*.txt)|*.txt|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    DlgTitle:=DTitle, _
    DefaultExt:="txt", _
    InitDir:=indir, _
    Owner:=frmmain.hWnd)) Then

    ExportSaveDialog = sfile
End If

Set cd = Nothing
End Function
Public Function BMPLoadDialog(Optional DTitle As String, Optional ByRef FileTitle As String) As String
Dim cd As cCommonDialog
Dim sfile As String
Dim indir As String

If Len(LastPath) = 0 Then
    indir = App.Path
Else
    indir = LastPath
End If
'DoEvents
Set cd = New cCommonDialog
'sfile = FileName

If (cd.VBGetOpenFileName( _
    sfile, _
    FileTitle, _
    ReadOnly:=True, _
    HideReadOnly:=True, _
    Filter:="Picture files |*.bmp;*.gif;*.ico;*.jpg;*.jpeg;*.png;*.cur;*.rle;*.wmf;*.emf|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    InitDir:=indir, _
    DlgTitle:=DTitle, _
    DefaultExt:="", _
    Owner:=frmmain.hWnd)) Then
    BMPLoadDialog = sfile
End If

Set cd = Nothing
End Function
Public Function ImportLoadDialog(indir As String, Optional DTitle As String, Optional ByRef FileTitle As String) As String
Dim cd As cCommonDialog
Dim sfile As String
'Dim indir As String

'If Len(LastPath) = 0 Then
'    indir = App.Path
'Else
'    indir = LastPath
'End If

Set cd = New cCommonDialog

If (cd.VBGetOpenFileName( _
    sfile, _
    FileTitle, _
    ReadOnly:=True, _
    HideReadOnly:=True, _
    Filter:="TXT files or preview bmp |*.txt;*.respack;*.bmp|All Files (*.*)|*.*", _
    FilterIndex:=1, _
    InitDir:=indir, _
    DlgTitle:=DTitle, _
    DefaultExt:="", _
    Owner:=frmmain.hWnd)) Then
    ImportLoadDialog = sfile
End If

Set cd = Nothing
End Function
Public Function GetPathFromPathAndName(ByVal sThePathAndName As String) As String
' Return a string containing the file's path from a fully qualified file name. Larry
' Return "" if no path
Dim i As Integer                    'used in for/next loops
Dim sTemp As String

sTemp = Trim$(sThePathAndName)      'trim it
If InStrB(sTemp, "\") = 0 Then       'any backslash?
    Exit Function
End If
For i = Len(sTemp) To 1 Step -1     'find the right most one
    If Mid$(sTemp, i, 1) = "\" Then
        GetPathFromPathAndName = Mid$(sTemp, 1, i)    'now have just path
        Exit Function
    End If
Next
End Function

'Public Function pSaveDialog(Optional DTitle As String, Optional myFileName As String) As String
'
'Dim cd As cCommonDialog
'Dim sfile As String
'Dim fTitle As String
'
'Set cd = New cCommonDialog
''sfile = FileName
''DoEvents
'If (cd.VBGetSaveFileName( _
  '    sfile, _
  '    fTitle, _
  '    Filter:="Binary files |*.bin|All Files (*.*)|*.*", _
  '    FilterIndex:=1, _
  '    DlgTitle:=DTitle, _
  '    DefaultExt:="bin", _
  '    Owner:=frmmain.hwnd)) Then
'    pSaveDialog = sfile
'End If
'
'Set cd = Nothing
'End Function
'

Public Function intBytes(nBits As Long) As Integer
'ret 8,16,24...
intBytes = 8
Do While intBytes / nBits < 1
    intBytes = intBytes + 8
Loop

End Function

'Public Function b4toLong_LE(startAdr As Long, arr() As Byte) As Long
'On Error Resume Next
'b4toLong_LE = arr(startAdr + 3) + arr(startAdr + 2) * 2 ^ 8 + arr(startAdr + 1) * 2 ^ 16 + arr(startAdr) * 2 ^ 24
'End Function

Public Function b4toLong_BE(startAdr As Long, arr() As Byte) As Long
On Error Resume Next
b4toLong_BE = arr(startAdr) + arr(startAdr + 1) * 2 ^ 8 + arr(startAdr + 2) * 2 ^ 16 + arr(startAdr + 3) * 2 ^ 24
End Function
Private Sub Main()
Dim iccex As InitCommonControlsExStruct, hmod As Long
' constant descriptions: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
'Const ICC_ANIMATE_CLASS As Long = &H80&
'Const ICC_BAR_CLASSES As Long = &H4&
'Const ICC_COOL_CLASSES As Long = &H400&
'Const ICC_DATE_CLASSES As Long = &H100&
'Const ICC_HOTKEY_CLASS As Long = &H40&
'Const ICC_INTERNET_CLASSES As Long = &H800&
'Const ICC_LINK_CLASS As Long = &H8000&
'Const ICC_LISTVIEW_CLASSES As Long = &H1&
'Const ICC_NATIVEFNTCTL_CLASS As Long = &H2000&
'Const ICC_PAGESCROLLER_CLASS As Long = &H1000&
'Const ICC_PROGRESS_CLASS As Long = &H20&
'Const ICC_TAB_CLASSES As Long = &H8&
'Const ICC_TREEVIEW_CLASSES As Long = &H2&
'Const ICC_UPDOWN_CLASS As Long = &H10&
'Const ICC_USEREX_CLASSES As Long = &H200&
Const ICC_STANDARD_CLASSES As Long = &H4000&
'Const ICC_WIN95_CLASSES As Long = &HFF&
'Const ICC_ALL_CLASSES As Long = &HFDFF&    ' combination of all values above

With iccex
    .lngSize = LenB(iccex)
    .lngICC = ICC_STANDARD_CLASSES    ' vb intrinsic controls (buttons, textbox, etc)
    ' if using Common Controls; add appropriate ICC_ constants for type of control you are using
    ' example if using CommonControls v5.0 Progress bar:
    ' .lngICC = ICC_STANDARD_CLASSES Or ICC_PROGRESS_CLASS
End With

On Error Resume Next    ' error? InitCommonControlsEx requires IEv3 or above

hmod = LoadLibraryA("shell32.dll")    ' patch to prevent XP crashes when VB usercontrols present
InitCommonControlsEx iccex
If Err Then
    InitCommonControls    ' try Win9x version
    Err.Clear
End If

'On Error GoTo 0

'... show your main form next (i.e., Form1.Show)
frmmain.Show

'On Error Resume Next
Dim Tmp As String
If Len(Command) <> 0 Then
    DoEvents
    Tmp = Replace(Command, """", vbNullString)
    Tmp = Replace(Tmp, """", vbNullString)
    If FileExists(Tmp) Then
        FileNameFW = Tmp
        FileTitle = GetNameExt(FileNameFW)

        'encrypt if decrypted
        If Not frmmain.EncryptFW Then
            MsgBoxEx ArrMsg(19), , , CenterOwner, vbCritical
            'Close #bFileIn
            fFileOpen = False
            Exit Sub    'unknown FW
        End If

        Call frmmain.LoadFWfile
    End If
End If
On Error GoTo 0

If hmod Then FreeLibrary hmod


'** Tip 1: Avoid using VB Frames when applying XP/Vista themes
'          In place of VB Frames, use pictureboxes instead.
'          'bug' may no longer apply to Win7+
'** Tip 2: Avoid using Graphical Style property of buttons, checkboxes and option buttons
'          Doing so will prevent them from being themed.
End Sub

'Public Function SplitToLines(hdcObject As Object, ByVal sText As String, _
 '                             ByVal lLength As Long, Optional ByVal bFilterLines As Boolean = True) As String()
'
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
'    If hdcObject.TextWidth(mLine & mChar) >= lLength Or X = xMax Then
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

Public Function IsDelim(Char As String) As Boolean
Select Case Asc(Char)    ' Upper/Lowercase letters,Underscore Not delimiters
Case 65 To 90, 95, 97 To 122
    IsDelim = False
Case Else: IsDelim = True    ' Another Character Is delimiter
End Select
End Function

Public Function GetBlockFromText(strText As String, StartAt As String, Optional FinAt As String) As String
'get txt block after StartAt and before FinAt
'or from begin string to
'or from start to end string
Dim StartPos As Long, FinPos As Long    ', BeginPos As Long
On Error GoTo modErr

StartPos = InStr(strText, StartAt)
If StartPos > 0 Then
    StartPos = StartPos + Len(StartAt)

    FinPos = InStr(StartPos, strText, FinAt)
    If FinPos > StartPos Then
        GetBlockFromText = Mid$(strText, StartPos, FinPos - StartPos)
    Else
        If Len(strText) > StartPos + 1 Then
            GetBlockFromText = right$(strText, Len(strText) - StartPos + 1)
        End If
    End If
'   End If
Else
    FinPos = InStr(1, strText, FinAt)
    If FinPos > 0 Then
        GetBlockFromText = Mid$(strText, 1, FinPos - 1)
    Else
        GetBlockFromText = strText
    End If
End If


'''
Exit Function
modErr:
MsgBox Err.Description & ": GetBlockFromText"
End Function
Public Function DecodeUTF8(ByVal sInput As String) As String
Dim iStrSize As Long, lMaxSize As Long, str1 As String
'Dim p As Long
'Dim str2 As String
On Error GoTo modErr

If Len(sInput) = 0 Then Exit Function

lMaxSize = Len(sInput)
str1 = String$(lMaxSize, 0&)
iStrSize = MultiByteToWideCharA(CP_UTF8, 0&, sInput, &HFFFF, StrPtr(str1), lMaxSize)
If iStrSize > 0 Then
    DecodeUTF8 = left$(str1, iStrSize - 1)
Else
    DecodeUTF8 = sInput
End If
'''
Exit Function
modErr:
MsgBox Err.Description & ": DecodeUTF8"
End Function
Public Function FillByteArray(ByVal HexValues As String, Data() As Byte) As Boolean
   
   Dim i As Long, n As Long
   n = Len(HexValues)
   ' Input string must be a multiple of two chars.
   If (n > 0) And (n Mod 2 = 0) Then
      ReDim Data(0 To n \ 2 - 1) As Byte
      For i = 0 To UBound(Data)
         Data(i) = Val("&h" & Mid$(HexValues, i * 2 + 1, 2))
      Next i
      FillByteArray = True
   End If
End Function

Public Sub GetLanguage(frm As Integer)
Dim Contrl As Control
Dim i As Integer

'Dim n As Integer
'Dim keysArr() As String
On Error GoTo modErr

If Len(lngFileName) = 0 Then Exit Sub

Select Case frm
Case 1
    With frmmain
        For Each Contrl In .Controls

            If TypeOf Contrl Is CommandButton Then
                If Contrl.Name = "cmdToolBar" Then

                    For i = 0 To .cmdToolBar.UBound    '13    'toolbar num
                        .cmdToolBar(i).ToolTipText = ReadLang("Main", Contrl.Name & i & ".tt", .cmdToolBar(i).ToolTipText)
                    Next i
                ElseIf Contrl.Name = "cmdUndoRedo" Then
                    For i = 0 To 1    'undo-redo
                        .cmdUndoRedo(i).ToolTipText = ReadLang("Main", Contrl.Name & i & ".tt", .cmdUndoRedo(i).ToolTipText)
                    Next i
                ElseIf Contrl.Name = "cmdINIShow" Then
                    For i = 0 To 1    'undo-redo
                        .cmdINIShow(i).ToolTipText = ReadLang("Main", Contrl.Name & i & ".tt", .cmdINIShow(i).ToolTipText)
                    Next i
                ElseIf Contrl.Name = "cmdVocabSL" Then
                    For i = 0 To 1    'undo-redo
                        .cmdVocabSL(i).ToolTipText = ReadLang("Main", Contrl.Name & i & ".tt", .cmdVocabSL(i).ToolTipText)
                    Next i

                Else
                    Contrl.Caption = ReadLang("Main", Contrl.Name, Contrl.Caption)
                    Contrl.ToolTipText = ReadLang("Main", Contrl.Name & ".tt", Contrl.ToolTipText)
                End If

            End If

            If TypeOf Contrl Is CheckBox Then
                Contrl.Caption = ReadLang("Main", Contrl.Name, Contrl.Caption)
                Contrl.ToolTipText = ReadLang("Main", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If
            If TypeOf Contrl Is Label Then
                Contrl.Caption = ReadLang("Main", Contrl.Name, Contrl.Caption)
                Contrl.ToolTipText = ReadLang("Main", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If
            If TypeOf Contrl Is TextBox Then
                'Contrl.Caption = ReadLang("Main", Contrl.Name, Contrl.Caption)
                Contrl.ToolTipText = ReadLang("Main", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If

            If TypeOf Contrl Is ComboBox Then
                Contrl.ToolTipText = ReadLang("Main", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If
            If TypeOf Contrl Is OptionButton Then
                If Contrl.Name = "optBlock" Then
                    For i = 0 To 1
                        .optBlock(i).Caption = ReadLang("Main", Contrl.Name & i, .optBlock(i).Caption)
                        .optBlock(i).ToolTipText = ReadLang("Main", Contrl.Name & i & ".tt", .optBlock(i).ToolTipText)
                    Next i
                End If

            End If

            If TypeOf Contrl Is PictureBox Then
                Contrl.ToolTipText = ReadLang("Main", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If

            If TypeOf Contrl Is Menu Then
                If Contrl.Name = "mnu_SaveBMP" Then
                    For i = 0 To 1
                        .mnu_SaveBMP(i).Caption = ReadLang("Main", Contrl.Name & i, .mnu_SaveBMP(i).Caption)
                    Next i

                ElseIf Contrl.Name = "mnu_CopyPic" Then
                    For i = 0 To 1
                        .mnu_CopyPic(i).Caption = ReadLang("Main", Contrl.Name & i, .mnu_CopyPic(i).Caption)
                    Next i

                ElseIf Contrl.Name = "mnu_Export" Then
                    For i = 0 To .mnu_Export.count - 1
                        .mnu_Export(i).Caption = ReadLang("Main", Contrl.Name & i, .mnu_Export(i).Caption)
                    Next i
                Else

                    Contrl.Caption = ReadLang("Main", Contrl.Name, Contrl.Caption)
                End If

            End If
        Next
    End With

Case 2
    With frmPatch
        For Each Contrl In .Controls

            If TypeOf Contrl Is CommandButton Then
                Contrl.Caption = ReadLang("Patcher", Contrl.Name, Contrl.Caption)
                Contrl.ToolTipText = ReadLang("Patcher", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If

            If TypeOf Contrl Is Label Then
                Contrl.Caption = ReadLang("Patcher", Contrl.Name, Contrl.Caption)
            End If
            If TypeOf Contrl Is CheckBox Then
                Contrl.Caption = ReadLang("Patcher", Contrl.Name, Contrl.Caption)
                Contrl.ToolTipText = ReadLang("Patcher", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If

        Next
    End With
Case 3
    With frmOptions
        .Caption = ArrMsg(35)

        For Each Contrl In .Controls

            If TypeOf Contrl Is CommandButton Then
                Contrl.Caption = ReadLang("Options", Contrl.Name, Contrl.Caption)
                Contrl.ToolTipText = ReadLang("Options", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If

            If TypeOf Contrl Is Label Then
                Contrl.Caption = ReadLang("Options", Contrl.Name, Contrl.Caption)
            End If
            '            If TypeOf Contrl Is CheckBox Then
            '                Contrl.Caption = ReadLang("Options", Contrl.Name, Contrl.Caption)
            '                Contrl.ToolTipText = ReadLang("Options", Contrl.Name & ".tt", Contrl.ToolTipText)
            '            End If
        Next
    End With
Case 4
    With frmParameters
        For Each Contrl In .Controls
            If TypeOf Contrl Is CommandButton Then
                Contrl.Caption = ReadLang("Param", Contrl.Name, Contrl.Caption)
                Contrl.ToolTipText = ReadLang("Param", Contrl.Name & ".tt", Contrl.ToolTipText)
            End If

        Next
    End With
End Select

'n = GetKeyNames("Messages", lngFileName, keysArr)
'ReDim ArrMsg(n)

If OldlngFileName <> lngFileName Then
    For i = 0 To UBound(ArrMsg)
        ArrMsg(i) = VBGetPrivateProfileString("Messages", CStr(i), lngFileName, ArrMsg(i))
    Next i
    OldlngFileName = lngFileName
End If

'''
Exit Sub
modErr:
MsgBox Err.Description & ": GetLanguage"
End Sub
Public Function ReadLang(sect As String, Itm As String, ss As String) As String

ReadLang = VBGetPrivateProfileString(sect, Itm, lngFileName, ss)

End Function

Public Function OpenFW_read() As Boolean
On Error GoTo modErr

If fFileOpen Then
    Close #bFileIn
    fFileOpen = False
'bFileIn = FreeFile
End If
If Len(FileNameFW) = 0 Then Exit Function

'MsgBox "OpenFW_read " & vbCrLf & FileNameFW & vbCrLf & "C:\EvicVTCFont\BIN\Wismec_30.06.16\Reuleaux_RX200S_V4.10.bin"
'FileNameFW = "C:\EvicVTCFont\BIN\Wismec_30.06.16\Reuleaux_RX200S_V4.10.bin"
'Open FileNameFW For Binary Access Read Shared As #bFileIn
'FileNameFW = Replace(FileNameFW, vbCrLf, vbNullString)
Open FileNameFW For Binary Access Read As #bFileIn
OpenFW_read = True
'''
Exit Function
modErr:
MsgBox Err.Description & ": OpenFW_read"
End Function

Public Function OpenFW_write() As Boolean
On Error GoTo modErr
If fFileOpen Then
    Close #bFileIn
    fFileOpen = False
End If
Open FileNameFW For Binary Access Read Write As #bFileIn
OpenFW_write = True

'''
Exit Function
modErr:
MsgBox Err.Description & ": OpenFW_write"

End Function

Public Sub SetMessages()
'go Sub GetLanguage
ReDim ArrMsg(48)
ArrMsg(0) = "Open firmware file first!"
ArrMsg(1) = "No block 1 specified in INI file."
ArrMsg(2) = "No block 2 specified in INI file."
ArrMsg(3) = "Wrong font size in buffer! "
ArrMsg(4) = "Wrong size of character in file! "
ArrMsg(5) = "Wrong size of character in other block! "
ArrMsg(6) = "Warning: wrong size. Must be "
ArrMsg(7) = " space separated bytes"
ArrMsg(8) = "Error: VTCFont.ini not found!"
ArrMsg(9) = "All done."
ArrMsg(10) = "Check one or more patches to apply!"
ArrMsg(11) = "Warning: Language file not found!"
ArrMsg(12) = "Save &All"
ArrMsg(13) = "&Copy"
ArrMsg(14) = "Open Bitmap picture"
ArrMsg(15) = "Open firmware file"
ArrMsg(16) = "Choose TTfont"
ArrMsg(17) = "Open text file with font data"
ArrMsg(18) = "Save font data to text file"
ArrMsg(19) = "Unknown firmware, proceed?"
ArrMsg(20) = "Quit program without save to firmware?"
ArrMsg(21) = "Save to BMP file"
ArrMsg(22) = "Image size too big for use with opened firmware (256 pixels max)."
ArrMsg(23) = "Unsaved edits will be lost! Proceed?"
ArrMsg(24) = "Change character size?"
ArrMsg(25) = "Record symbol?"
ArrMsg(26) = "Yes"
ArrMsg(27) = "No"
ArrMsg(28) = "Normal"
ArrMsg(29) = "Swap L-R"
ArrMsg(30) = "Inverse"
ArrMsg(31) = "CustomShades"
ArrMsg(32) = "AtkinsonGS"
ArrMsg(33) = "AtkinsonBW"
ArrMsg(34) = "ShadesDithered"
ArrMsg(35) = "Options"
ArrMsg(36) = "In one collumn"
ArrMsg(37) = "In box"
ArrMsg(38) = "Change glyph size"
ArrMsg(39) = "Set width to:"
ArrMsg(40) = "Set height to:"
ArrMsg(41) = "Show installed"
ArrMsg(42) = "Hide installed"
ArrMsg(43) = "Value is greater than allowable!"
ArrMsg(44) = "Wrong parameter!"
ArrMsg(45) = "FWUpdater.exe not found."
ArrMsg(46) = "Shure convert patch to VTCFont format?"
ArrMsg(47) = "Save Encrypted FW file"
ArrMsg(48) = "Save Decrypted FW file"
End Sub

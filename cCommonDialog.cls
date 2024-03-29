VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
'Private Declare Sub CopyMemoryStr Lib "KERNEL32" Alias "RtlMoveMemory" ( _
  lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    Flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
'Private Declare Function GetFileTitle Lib "COMDLG32" Alias "GetFileTitleA" (ByVal szFile As String, ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private m_lApiReturn As Long
Private m_lExtendedError As Long
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long


Function VBGetOpenFileName(filename As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional Flags As Long = 0) As Boolean

Dim opfile As OPENFILENAME, s As String    'mzt , afFlags As Long
Dim ch As String, i As Integer
m_lApiReturn = 0
m_lExtendedError = 0

MouseUpBug = True

With opfile
    .lStructSize = Len(opfile)

' Add in specific flags and strip out non-VB flags

    .Flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
             (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle

' To make Windows-style filter, replace | and : with nulls

    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

' Pad file and file title buffers to maximum path
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
' All other fields set to zero

    m_lApiReturn = GetOpenFileName(opfile)
    Select Case m_lApiReturn
    Case 1
' Success
        VBGetOpenFileName = True
        filename = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        Flags = .Flags
' Return the filter index
        FilterIndex = .nFilterIndex
' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
        If (.Flags And OFN_READONLY) Then ReadOnly = True
    Case 0
' Cancelled
        VBGetOpenFileName = False
        filename = vbNullString
        FileTitle = vbNullString
        Flags = 0
        FilterIndex = -1
        Filter = vbNullString
    Case Else
' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBGetOpenFileName = False
        filename = vbNullString
        FileTitle = vbNullString
        Flags = 0
        FilterIndex = -1
        Filter = vbNullString
    End Select
End With
End Function
Private Function StrZToStr(s As String) As String
StrZToStr = left$(s, lstrlen(s))
End Function

Function VBGetSaveFileName(filename As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional Flags As Long) As Boolean

Dim opfile As OPENFILENAME, s As String

m_lApiReturn = 0
m_lExtendedError = 0
MouseUpBug = True

With opfile
    .lStructSize = Len(opfile)

' Add in specific flags and strip out non-VB flags
    .Flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (Flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle

' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

' Pad file and file title buffers to maximum path
    s = filename & String$(MAX_PATH - Len(filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
' All other fields zero

    m_lApiReturn = GetSaveFileName(opfile)
    Select Case m_lApiReturn
    Case 1
        VBGetSaveFileName = True
        filename = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        Flags = .Flags
' Return the filter index
        FilterIndex = .nFilterIndex
' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
    Case 0
' Cancelled:
        VBGetSaveFileName = False
        filename = vbNullString
        FileTitle = vbNullString
        Flags = 0
        FilterIndex = 0
        Filter = vbNullString
    Case Else
' Extended error:
        VBGetSaveFileName = False
        m_lExtendedError = CommDlgExtendedError()
        filename = vbNullString
        FileTitle = vbNullString
        Flags = 0
        FilterIndex = 0
        Filter = vbNullString
    End Select
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
Dim iStart As Long, iEnd As Long, s As String
iStart = 1
If sFilters = vbNullString Then Exit Function
Do
' Cut out both parts marked by null character
    iEnd = InStr(iStart, sFilters, vbNullChar)
    If iEnd = 0 Then Exit Function
    iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
    If iEnd Then
        s = Mid$(sFilters, iStart, iEnd - iStart)
    Else
        s = Mid$(sFilters, iStart)
    End If
    iStart = iEnd + 1
    If iCur = 1 Then
        FilterLookup = s
        Exit Function
    End If
    iCur = iCur - 1
Loop While iCur
End Function


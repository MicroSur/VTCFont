Attribute VB_Name = "modMoveMsgBox"
Option Explicit

Public Enum StartupPos
    CenterScreen
    CenterOwner
    Custom
End Enum

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type CWPSTRUCT
    LParam As Long
    wParam As Long
    message As Long
    hWnd As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, LParam As Any) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lprect As RECT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const GWL_WNDPROC = (-4)
Private Const WH_CALLWNDPROC = 4
Private Const WM_CREATE = &H1
Private Const WM_INITDIALOG = &H110

Private MSGBOXEX_X As Integer
Private MSGBOXEX_Y As Integer
Private MSGBOXEX_STARTUP As StartupPos
Private lPrevWnd As Long
Private lHook As Long

Private Function SubMsgBox(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Dim tRECT As RECT
Dim tOWNER As RECT
If Msg = WM_INITDIALOG Then
'Reposition the MsgBox is Neccessary..
    If MSGBOXEX_STARTUP = CenterOwner Then
        Call GetWindowRect(GetParent(hWnd), tOWNER)
    Else
        Call GetWindowRect(0, tOWNER)
    End If
    Call GetWindowRect(hWnd, tRECT)
    Select Case MSGBOXEX_STARTUP
    Case Custom
        If MSGBOXEX_X = -1 Then
'Center Horz
            tRECT.left = ((Screen.Width / Screen.TwipsPerPixelX) - (tRECT.right - tRECT.left)) / 2
        Else
'Position Horz
            tRECT.left = MSGBOXEX_X
        End If
        If MSGBOXEX_Y = -1 Then
'Center Vert
            tRECT.top = ((Screen.Height / Screen.TwipsPerPixelY) - (tRECT.bottom - tRECT.top)) / 2
        Else
'Position Vert
            tRECT.top = MSGBOXEX_Y
        End If
        Call SetWindowPos(hWnd, 0, tRECT.left, tRECT.top, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)

    Case CenterOwner
        tRECT.left = tOWNER.left + (((tOWNER.right - tOWNER.left) - (tRECT.right - tRECT.left)) / 2)
        tRECT.top = tOWNER.top + (((tOWNER.bottom - tOWNER.top) - (tRECT.bottom - tRECT.top)) / 2)
        Call SetWindowPos(hWnd, 0, tRECT.left, tRECT.top, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)

    End Select
'Remove the Messagebox Subclassing
    Call SetWindowLong(hWnd, GWL_WNDPROC, lPrevWnd)
End If
SubMsgBox = CallWindowProc(lPrevWnd, hWnd, Msg, wParam, ByVal LParam)
End Function

Private Function HookWindow(ByVal nCode As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Dim tCWP As CWPSTRUCT
Dim sClass As String
'This is where you need to Hook the Messagebox
CopyMemory tCWP, ByVal LParam, Len(tCWP)
Select Case tCWP.message
Case WM_CREATE
    sClass = Space(255)
    sClass = left$(sClass, GetClassName(tCWP.hWnd, ByVal sClass, 255))
    If sClass = "#32770" Then
'Subclass the Messagebox as it's created
        lPrevWnd = SetWindowLong(tCWP.hWnd, GWL_WNDPROC, AddressOf SubMsgBox)
    End If
End Select
HookWindow = CallNextHookEx(lHook, nCode, wParam, ByVal LParam)
End Function

Public Function MsgBoxEx(Prompt, Optional X = -1, Optional Y = -1, Optional StartupPosition As StartupPos = CenterScreen, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title, Optional HelpFile, Optional Context) As VbMsgBoxResult
Dim tOWNER As RECT

MSGBOXEX_X = X
MSGBOXEX_Y = Y
MSGBOXEX_STARTUP = StartupPosition

'sur. shure visible
Call GetWindowRect(frmmain.hWnd, tOWNER)
If tOWNER.right * Screen.TwipsPerPixelX > Screen.Width Then
    MSGBOXEX_STARTUP = CenterScreen
ElseIf tOWNER.left * Screen.TwipsPerPixelX < 0 Then
    MSGBOXEX_STARTUP = CenterScreen
ElseIf tOWNER.bottom * Screen.TwipsPerPixelY > Screen.Height Then
    MSGBOXEX_STARTUP = CenterScreen
End If


'set a Thread Message Hook..
lHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf HookWindow, App.hInstance, App.ThreadID)
MsgBoxEx = MsgBox(Prompt, Buttons, Title, HelpFile, Context)
'Remove the Hook
Call UnhookWindowsHookEx(lHook)
End Function


Attribute VB_Name = "mInputBoxEx"
Option Explicit

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNTEXT = 18
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const WH_CALLWNDPROC = 4
Private Const GWL_WNDPROC = (-4)
Private Const WM_GETFONT = &H31
Private Const WM_CREATE = &H1
Private Const WM_CTLCOLORBTN = &H135
Private Const WM_CTLCOLORDLG = &H136
Private Const WM_CTLCOLORSTATIC = &H138
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_DESTROY = &H2
Private Const WM_SHOWWINDOW = &H18
Private Const WM_COMMAND = &H111
Private Const BN_CLICKED = 0
Private Const IDOK = 1
Private Const EM_SETPASSWORDCHAR = &HCC
Private INPUTBOX_HOOK As Long
Private INPUTBOX_HWND As Long
Private INPUTBOX_PASSCHAR As String
Private INPUTBOX_BACKCOLOR As Long
Private INPUTBOX_FORECOLOR As Long
Private INPUTBOX_FONT As String
Private INPUTBOX_FONTSIZE As Integer
Private INPUTBOX_SHOWING As Boolean
Private INPUTBOX_CENTERV As Boolean
Private INPUTBOX_CENTERH As Boolean
Private INPUTBOX_OK As Boolean

Private Function InputBoxProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tLB As LOGBRUSH
Dim lFont As Long
Dim tRECT As RECT
Dim lNotify As Long
Dim lID As Long
Select Case Msg
    Case WM_COMMAND
        lNotify = Val("&H" & Left$(Right$("00000000" & Hex$(wParam), 8), 4))
        lID = Val("&H" & Right$(Right$("00000000" & Hex$(wParam), 8), 4))
        If lNotify = BN_CLICKED Then
            INPUTBOX_OK = (lID = IDOK)
        End If
    Case WM_SHOWWINDOW
        Call GetWindowRect(hWnd, tRECT)
        If INPUTBOX_CENTERH Then tRECT.Left = ((Screen.Width / Screen.TwipsPerPixelX) - (tRECT.Right - tRECT.Left)) / 2
        If INPUTBOX_CENTERV Then tRECT.Top = ((Screen.Height / Screen.TwipsPerPixelY) - (tRECT.Bottom - tRECT.Top)) / 2
        Call SetWindowPos(hWnd, 0, tRECT.Left, tRECT.Top, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
    Case WM_CTLCOLORDLG, WM_CTLCOLORSTATIC, WM_CTLCOLORBTN, WM_CTLCOLOREDIT
        If Msg = WM_CTLCOLOREDIT Then
            If Len(INPUTBOX_PASSCHAR) Then
                Call SendMessage(lParam, EM_SETPASSWORDCHAR, Asc(INPUTBOX_PASSCHAR), ByVal 0&)
            End If
        Else
            Call SetTextColor(wParam, INPUTBOX_FORECOLOR)
            Call SetBkColor(wParam, INPUTBOX_BACKCOLOR)
            If Msg = WM_CTLCOLORSTATIC Then
                lFont = CreateFont(-((INPUTBOX_FONTSIZE / 72) * 96), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, INPUTBOX_FONT)
                Call SelectObject(wParam, lFont)
            End If
            tLB.lbColor = INPUTBOX_BACKCOLOR
            InputBoxProc = CreateBrushIndirect(tLB)
            Exit Function
        End If
    Case WM_DESTROY
        Call SetWindowLong(hWnd, GWL_WNDPROC, INPUTBOX_HWND)
End Select
InputBoxProc = CallWindowProc(INPUTBOX_HWND, hWnd, Msg, wParam, ByVal lParam)
End Function

Private Function HookWindow(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tCWP As CWPSTRUCT
Dim sClass As String
CopyMemory tCWP, ByVal lParam, Len(tCWP)
If tCWP.message = WM_CREATE Then
    sClass = Space(255)
    sClass = Left(sClass, GetClassName(tCWP.hWnd, ByVal sClass, 255))
    If sClass = "#32770" Then
        If INPUTBOX_SHOWING Then
            INPUTBOX_HWND = SetWindowLong(tCWP.hWnd, GWL_WNDPROC, AddressOf InputBoxProc)
        End If
    End If
End If
HookWindow = CallNextHookEx(INPUTBOX_HOOK, nCode, wParam, ByVal lParam)
End Function

Public Function InputBoxEx(ByVal Prompt As String, Optional ByVal Title As String, Optional ByVal Default As String, Optional ByVal XPos As Single = -1, Optional ByVal YPos As Single = -1, Optional ByVal HelpFile As String, Optional ByVal Context As Long, Optional ByVal ForeColor As ColorConstants, Optional ByVal BackColor As ColorConstants, Optional ByVal FontName As String, Optional ByVal FontSize As Long, Optional ByVal PasswordChar As String, Optional ByVal CancelError As Boolean = False) As String
If Len(Title) = 0 Then Title = App.Title
INPUTBOX_FONT = "MS Sans Serif"
INPUTBOX_FONTSIZE = 8
INPUTBOX_FORECOLOR = GetSysColor(COLOR_BTNTEXT)
INPUTBOX_BACKCOLOR = GetSysColor(COLOR_BTNFACE)
INPUTBOX_CENTERH = (XPos = -1)
INPUTBOX_CENTERV = (YPos = -1)
INPUTBOX_PASSCHAR = PasswordChar
If Len(FontName) Then INPUTBOX_FONT = FontName
If FontSize > 0 Then INPUTBOX_FONTSIZE = FontSize
If ForeColor > 0 Then INPUTBOX_FORECOLOR = ForeColor
If BackColor > 0 Then INPUTBOX_BACKCOLOR = BackColor
INPUTBOX_SHOWING = True
INPUTBOX_HOOK = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf HookWindow, App.hInstance, App.ThreadID)
InputBoxEx = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
INPUTBOX_SHOWING = False
Call UnhookWindowsHookEx(INPUTBOX_HOOK)
If Not INPUTBOX_OK And CancelError Then Err.Raise vbObjectError + 1, , "User Pressed " & Chr(34) & "Cancel" & Chr(34)
End Function

Attribute VB_Name = "Module1"
Option Explicit

' ***********************************
' Author : Imran Zaheer
' Email  : imraanz@mail.com
' Web    : www.imraanz.com
' Y2K
' Module : Contains declarations and functions for
'          vHotKeys.

Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, _
    ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
    
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, _
    ByVal ID As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = -4

Public Const MOD_CTRL = &H2 'This example uses CTRL
Public Const MOD_SHFT = &H4
Public Const MOD_ALT = &H1

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'
' and others are listed below

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79 ' This example uses F10
Public Const VK_F11 = &H7A ' This example uses F11
Public Const VK_F12 = &H7B ' This example uses F12
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public glWinRet As Long


' Function : CallbackMsgs
' This functions is used as a parameter in the
' API SetWindowLong(), by AddresOf operator, so as to
' Subclass the form to get the Windows Callback msgs...
Public Function CallbackMsgs(ByVal wHwnd As Long, ByVal wmsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
    If wmsg = WM_HOTKEY Then
        Call DoFunctions(wp_id)
        CallbackMsgs = 1
        Exit Function
    End If
    CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wmsg, wp_id, lp_id)
End Function


' Sub : DoFunction
' Activated by the Function "CallbackMsgs()" whenever
' a hotkey is pressed.
Public Sub DoFunctions(ByVal vKeyID As Byte)
    ' Important Notes :
    ' Do not include any msgboxes or Modal forms in
    ' this procedure, else if you include then by
    ' pressing the Hotkey twice/thrice the application
    ' will be terminated abnormally.
    '
    ' But if it is a requirement for you to include the
    ' Modal forms or msgbox in this procedure then put
    ' the RegisterHotKey() API before hiding the Form
    ' and put the UnRegisterHotKey() API before Showing
    ' the form.
    
    Form1.Show
    Form1.WindowState = 0
    DoEvents
    ' When the Hotkey is pressed once
    ' check if the Dofunctions() has completed
    ' before the CallbackMsgs().
    ' This check is not required if the form is
    ' minimized in the SysTray ...
    If Form1.Visible = False Then
        Form1.Show
        Form1.WindowState = 0
    End If
    
    Form1.Cls
    
    If vKeyID = 0 Then
        Form1.Label1.Caption = "1st HotKey Pressed !"
    Else
        If vKeyID = 1 Then
            Form1.Label1.Caption = "2nd HotKey Pressed !"
        Else
            Form1.Label1.Caption = "3rd HotKey Pressed !"
        End If
    End If
End Sub

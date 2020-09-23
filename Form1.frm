VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VHotKeys"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   495
      TabIndex        =   0
      Top             =   1200
      Width           =   3675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ***********************************
' Author : Imran Zaheer
' Email  : imraanz@mail.com
' Web    : www.imraanz.com
' Y2K


Dim retVal0 As Boolean, retVal1 As Boolean, retVal2 As Boolean

Private Sub Form_Load()

    MsgBox "Registering CTRL+F10, CTRL+F11, CTRL+F12 as hot keys..."
    
    retVal0 = RegisterHotKey(Me.hwnd, 0, MOD_CTRL, VK_F10)
    If Not retVal0 Then
        MsgBox "Can not register all or one of the hotkeys CTRL+F10 ... Try other keys this key is already registered by some other running applications.", vbCritical
    End If
    
    retVal1 = RegisterHotKey(Me.hwnd, 1, MOD_CTRL, VK_F11)
    If Not retVal1 Then
        MsgBox "Can not register all or one of the hotkeys CTRL+F11 ... Try other keys this key is already registered by some other running applications.", vbCritical
    End If
    
    retVal2 = RegisterHotKey(Me.hwnd, 2, MOD_CTRL, VK_F12)
    If Not retVal2 Then
        MsgBox "Can not register all or one of the hotkeys CTRL+F12 ... Try other keys this key is already registered by some other running applications.", vbCritical
    End If
    
    If (retVal0 = False And retVal1 = False And retVal2 = False) Then
        MsgBox "No Hotkey could be registered ...!", vbCritical
        End
    End If
    
    ' Subclassing the form to get the Windows callback msgs.
    glWinRet = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf CallbackMsgs)
    Me.Hide
    
End Sub


Private Sub Form_Resize()

    If Me.WindowState = 1 Then
        Me.Hide
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    ' If first hotkey is registered then
    ' unregister it.
    If retVal0 Then
        UnregisterHotKey Me.hwnd, 0
    End If
    
    ' If second hotkey is registered then
    ' unregister it.
    If retVal1 Then
        UnregisterHotKey Me.hwnd, 1
    End If
    
    ' If third hotkey is registered then
    ' unregister it.
    If retVal2 Then
        UnregisterHotKey Me.hwnd, 2
    End If
    
End Sub

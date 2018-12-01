VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   12315
   ClientTop       =   9630
   ClientWidth     =   5910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5910
   Visible         =   0   'False
   Begin VB.Menu trMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu trMenuOptions 
         Caption         =   "Настройки"
      End
      Begin VB.Menu trMenuAbout 
         Caption         =   "О программе"
      End
      Begin VB.Menu trMenuExit 
         Caption         =   "Выход"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Key_Shift = &H4
Private bSubclass As Boolean
Private bFailed As Boolean
Public WithEvents cTray As TrayIconAndBalloon
Attribute cTray.VB_VarHelpID = -1

Private Sub Form_Load()
    Dim lRet As Long
    Dim I As Integer
    bSubclass = True
    bFailed = False
    lRet = 0

    App.TaskVisible = False
    
    Set cTray = New TrayIconAndBalloon
    
    cTray.hwnd = hwnd
    cTray.Icon = Icon
    cTray.ToolTipText = App.ProductName
    cTray.Add
    
    loadConf
    
    For I = 65 To 90
        lRet = RegisterHotKey(Me.hwnd, I, 0, I)
        bFailed = (lRet = 0)
    Next
    For I = 65 To 90
        lRet = RegisterHotKey(Me.hwnd, I + 100, Key_Shift, I)
        bFailed = (lRet = 0)
    Next
    lRet = RegisterHotKey(Me.hwnd, 201, 0, &HBC)
    bFailed = (lRet = 0)

    lRet = RegisterHotKey(Me.hwnd, 201 + 100, Key_Shift, &HBC)
    bFailed = (lRet = 0)
    
    lRet = RegisterHotKey(Me.hwnd, 202, 0, &HBE)
    bFailed = (lRet = 0)

    lRet = RegisterHotKey(Me.hwnd, 202 + 100, Key_Shift, &HBE)
    bFailed = (lRet = 0)
    
    lRet = RegisterHotKey(Me.hwnd, 203, 0, &HDB)
    bFailed = (lRet = 0)

    lRet = RegisterHotKey(Me.hwnd, 203 + 100, Key_Shift, &HDB)
    bFailed = (lRet = 0)
    
    lRet = RegisterHotKey(Me.hwnd, 204, 0, &HDD)
    bFailed = (lRet = 0)

    lRet = RegisterHotKey(Me.hwnd, 204 + 100, Key_Shift, &HDD)
    bFailed = (lRet = 0)
    
    lRet = RegisterHotKey(Me.hwnd, 205, 0, &HBA)
    bFailed = (lRet = 0)

    lRet = RegisterHotKey(Me.hwnd, 205 + 100, Key_Shift, &HBA)
    bFailed = (lRet = 0)
    
    lRet = RegisterHotKey(Me.hwnd, 206, 0, &HDE)
    bFailed = (lRet = 0)

    lRet = RegisterHotKey(Me.hwnd, 206 + 100, Key_Shift, &HDE)
    bFailed = (lRet = 0)
    
    If bFailed Then
        MsgBox "Ошибка!"
        bSubclass = False
    End If
    
    If bSubclass Then
        lOrigin = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf pHcSubclass)
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call pHcUnsubclass(Me.hwnd)
    Dim I As Integer
    For I = 65 To 90
        Call UnregisterHotKey(Me.hwnd, I)
    Next
    For I = 65 To 90
        Call UnregisterHotKey(Me.hwnd, I + 100)
    Next
    Call UnregisterHotKey(Me.hwnd, 201)
    Call UnregisterHotKey(Me.hwnd, 201 + 100)
    Call UnregisterHotKey(Me.hwnd, 202)
    Call UnregisterHotKey(Me.hwnd, 202 + 100)
    Call UnregisterHotKey(Me.hwnd, 203)
    Call UnregisterHotKey(Me.hwnd, 203 + 100)
    Call UnregisterHotKey(Me.hwnd, 204)
    Call UnregisterHotKey(Me.hwnd, 204 + 100)
    Call UnregisterHotKey(Me.hwnd, 205)
    Call UnregisterHotKey(Me.hwnd, 205 + 100)
    Call UnregisterHotKey(Me.hwnd, 206)
    Call UnregisterHotKey(Me.hwnd, 206 + 100)
    cTray.Delete
    Set cTray = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      cTray.CallEvent X, Y
End Sub

Private Sub cTray_OnIcon(MouseButton As Integer, X As Single)
    If MouseButton = TRAYICON_MOUSE_LEFTDBLCLICK Then Load Form2: Form2.Show
    If MouseButton = TRAYICON_MOUSE_RIGHTUP And X = 7755 Then cTray.CallPopupMenu Me, trMenu, 2, , , trMenuOptions
    If MouseButton = TRAYICON_MOUSE_LEFTDBLCLICK Then Load Form3: Form3.Show
    If MouseButton = TRAYICON_MOUSE_RIGHTUP And X = 7754 Then cTray.CallPopupMenu Me, trMenu, 3, , , trMenuAbout
End Sub

Private Sub trMenuExit_Click()
finish
End Sub

Private Sub trMenuOptions_Click()
    Load Form2
    Form2.Show
End Sub
Private Sub trMenuAbout_Click()
    Load Form3
    Form3.Show
End Sub

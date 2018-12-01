VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrayIconAndBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Константы и структуры
'-------------------------------------------------
Private Const NOTIFYICON_VERSION = &H3
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_NULL = &H0

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

   
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type
'-----------------------------------------------------
' Объявление API функций
'-----------------------------------------------------
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" _
  (ByVal dwMessage As Long, _
   lpData As NOTIFYICONDATA) As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Private Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

Private Declare Function PostMessage Lib "user32" _
    Alias "PostMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

'---------------------------------------------------------------------

'переменные и структуры модуля
   Private NID              As NOTIFYICONDATA
   Private m_hWnd           As Long
   Private m_Icon           As Long
   Private m_ToolTipText    As String
   Public Enum TrayIcon_InfoIcon
        NIIF_NONE = &H0
        NIIF_INFO = &H1
        NIIF_WARNING = &H2
        NIIF_ERROR = &H3
        NIIF_GUID = &H5         ' Не используется
        NIIF_ICON_MASK = &HF    ' Не используется
        NIIF_NOSOUND = &H10
   End Enum
   Public Enum TrayIcon_Mouse
        TRAYICON_MOUSE_NONE = 0
        TRAYICON_MOUSE_LEFTDOWN = 1
        TRAYICON_MOUSE_LEFTUP = 2
        TRAYICON_MOUSE_LEFTDBLCLICK = 3
        TRAYICON_MOUSE_RIGHTDOWN = 4
        TRAYICON_MOUSE_RIGHTUP = 5
        TRAYICON_MOUSE_RIGHTDBLCLICK = 6
   End Enum

'ивенты
   Public Event OnIcon(MouseButton As Integer, X As Single)

'---------------------------------------------------------------------
'  Свойства

'получить свойство hWnd
Public Property Get hwnd() As Long
      hwnd = m_hWnd
End Property

'присвоить свойство hWnd
Public Property Let hwnd(ByVal NewValue As Long)
      m_hWnd = NewValue
End Property


'получить свойство Icon
Public Property Get Icon() As Long
      Icon = m_Icon
End Property

'присвоить свойство Icon
Public Property Let Icon(ByVal NewValue As Long)
      m_Icon = NewValue
End Property


'получить свойство ToolTipText
Public Property Get ToolTipText() As String
      ToolTipText = m_ToolTipText
End Property

'присвоить свойство ToolTipText
Public Property Let ToolTipText(ByVal NewValue As String)
      m_ToolTipText = NewValue
End Property

'---------------------------------------------------------------------
'  Методы

'добавление иконки в Tray
Public Sub Add()
   Dim ret As Long
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   'заполняем структуру
   With NID
      .uID = m_hWnd
      .hwnd = m_hWnd
      .hIcon = m_Icon
      .szTip = Left$(m_ToolTipText, 63) & Chr$(0)
      .uFlags = NIF_TIP + NIF_MESSAGE + NIF_ICON
      .uCallbackMessage = WM_RBUTTONDOWN
      .cbSize = NOTIFYICONDATA_SIZE
      .uTimeoutAndVersion = NOTIFYICON_VERSION
    End With
   'вызываем функцию API
      ret = Shell_NotifyIcon(NIM_ADD, NID)
   'Иформируем систему о новой иконке
      ret = Shell_NotifyIcon(NIM_SETVERSION, NID)
      
End Sub

'изменение иконки и ToolTip в Tray
Public Sub Modify()
   Dim ret As Long
   
   'заполняем структурку
    With NID
      .hIcon = m_Icon
      .szTip = Left$(m_ToolTipText, 63) & Chr$(0)
      .uFlags = NIF_TIP + NIF_ICON
      .cbSize = Len(NID)
    End With
      
   'вызываем функцию API
      ret = Shell_NotifyIcon(NIM_MODIFY, NID)
      
End Sub

' вывод всплывающего сообщения
Public Sub DisplayBalloon(ByVal sTitle As String, ByVal sText As String, Optional ByVal InfoFlags As TrayIcon_InfoIcon)
Attribute DisplayBalloon.VB_UserMemId = 0
Dim ret As Long
If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion

   With NID
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = m_hWnd
      .uID = m_hWnd
      .uFlags = NIF_INFO
      .dwInfoFlags = InfoFlags
      .szInfoTitle = sTitle & vbNullChar
      .szInfo = sText & vbNullChar
   End With

   ret = Shell_NotifyIcon(NIM_MODIFY, NID)
End Sub

'удаление иконки из Tray
Public Sub Delete()
   Dim ret As Long
   
   'заполняем структурку
      NID.cbSize = Len(NID)
   
   'вызываем функцию API
      ret = Shell_NotifyIcon(NIM_DELETE, NID)
      
End Sub


'вызвать ивент
Public Sub CallEvent(X As Single, Y As Single)
   'проверка, было ли действие по окну или по иконке в трее
      If Y Then Exit Sub
      
   'получаем тип действия
      Dim MouseButton As Integer
      MouseButton = (X And &HFF) / Screen.TwipsPerPixelX
   
   'генерируем событые
      RaiseEvent OnIcon(MouseButton, X)
      
End Sub

'вызов всплывающего меню
Public Sub CallPopupMenu(Window As Form, PopMenu As Menu, Optional Flags, Optional X, Optional Y, Optional DefMenu)
      
   'устанавливаем окно
      SetForegroundWindow m_hWnd
      
   'выводим всплывающее меню
      Window.PopupMenu PopMenu, Flags, X, Y, DefMenu
      
   'посылаем сообщение окну
      PostMessage m_hWnd, WM_NULL, 0, 0

End Sub

'Устанавливаем размер переменной NOTIFYICONDATA_SIZE
'в зависимости от версии оболочки
Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub

'Определяем версию оболочки Shell32.dll
Private Function IsShellVersion(ByVal version As Long) As Boolean

  'функция возращает Истину если версия оболочки
  '(shell32.dll) равна или больше запрашиваемой
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If
    
   End If
  
End Function

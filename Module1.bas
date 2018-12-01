Attribute VB_Name = "Module1"
Option Explicit

Public Const GWL_WNDPROC = (-4)
Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public lOrigin As Long

Dim lastTime As Long
Public options_remind As Boolean
Dim tmp As String

Public Function pHcSubclass(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hwnd = 0 Then Exit Function
    
    If uMsg = WM_HOTKEY Then
        If options_remind Then
            If Timer - lastTime >= 10 Then Form1.cTray.DisplayBalloon "StrangeLeet", "Do not forget that letters are changed to symbols!", NIIF_INFO ' + NIIF_NOSOUND
        End If
        lastTime = Timer
        If GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow, vbNull)) = 68748313 Then 'если русская раскладка
            'Дальше очень нудно
            If wParam = Asc("Q") Then SendKeys ("|i|")
            If wParam = Asc("W") Then SendKeys ("C")
            If wParam = Asc("E") Then SendKeys ("Y")
            If wParam = Asc("R") Then SendKeys ("K ")
            If wParam = Asc("T") Then SendKeys ("E")
            If wParam = Asc("Y") Then SendKeys ("#")
            If wParam = Asc("U") Then SendKeys ("G")
            If wParam = Asc("I") Then SendKeys ("|_|_|")
            If wParam = Asc("O") Then SendKeys ("|_|_|,")
            If wParam = Asc("P") Then SendKeys ("3")
            If wParam = 203 Then SendKeys ("H")
            If wParam = 204 Then SendKeys ("J")
            If wParam = Asc("A") Then SendKeys ("<|>")
            If wParam = Asc("S") Then SendKeys ("I0")
            If wParam = Asc("D") Then SendKeys ("]3")
            If wParam = Asc("F") Then SendKeys ("/-\")
            If wParam = Asc("G") Then SendKeys ("P")
            If wParam = Asc("H") Then SendKeys ("R")
            If wParam = Asc("J") Then SendKeys ("0")
            If wParam = Asc("K") Then SendKeys ("/\")
            If wParam = Asc("L") Then SendKeys ("D")
            If wParam = 205 Then SendKeys (">|<")
            If wParam = 206 Then SendKeys ("E")
            If wParam = Asc("Z") Then SendKeys ("R")
            If wParam = Asc("X") Then SendKeys ("4")
            If wParam = Asc("C") Then SendKeys ("S")
            If wParam = Asc("V") Then SendKeys ("/\/\")
            If wParam = Asc("B") Then SendKeys ("|/|")
            If wParam = Asc("N") Then SendKeys ("7")
            If wParam = Asc("M") Then SendKeys ("L")
            If wParam = 201 Then SendKeys ("6")
            If wParam = 202 Then SendKeys ("U")
            If wParam = Asc("Q") + 100 Then SendKeys ("|i|")
            If wParam = Asc("W") + 100 Then SendKeys ("C")
            If wParam = Asc("E") + 100 Then SendKeys ("Y")
            If wParam = Asc("R") + 100 Then SendKeys ("K ")
            If wParam = Asc("T") + 100 Then SendKeys ("E")
            If wParam = Asc("Y") + 100 Then SendKeys ("#")
            If wParam = Asc("U") + 100 Then SendKeys ("G")
            If wParam = Asc("I") + 100 Then SendKeys ("|_|_|")
            If wParam = Asc("O") + 100 Then SendKeys ("|_|_|,")
            If wParam = Asc("P") + 100 Then SendKeys ("3")
            If wParam = 203 + 100 Then SendKeys ("H")
            If wParam = 204 + 100 Then SendKeys ("J")
            If wParam = Asc("A") + 100 Then SendKeys ("<|>")
            If wParam = Asc("S") + 100 Then SendKeys ("I0")
            If wParam = Asc("D") + 100 Then SendKeys ("]3")
            If wParam = Asc("F") + 100 Then SendKeys ("/-\")
            If wParam = Asc("G") + 100 Then SendKeys ("P")
            If wParam = Asc("H") + 100 Then SendKeys ("R")
            If wParam = Asc("J") + 100 Then SendKeys ("0")
            If wParam = Asc("K") + 100 Then SendKeys ("/\")
            If wParam = Asc("L") + 100 Then SendKeys ("D")
            If wParam = 205 + 100 Then SendKeys (">|<")
            If wParam = 206 + 100 Then SendKeys ("E")
            If wParam = Asc("Z") + 100 Then SendKeys ("R")
            If wParam = Asc("X") + 100 Then SendKeys ("4")
            If wParam = Asc("C") + 100 Then SendKeys ("S")
            If wParam = Asc("V") + 100 Then SendKeys ("/\/\")
            If wParam = Asc("B") + 100 Then SendKeys ("|/|")
            If wParam = Asc("N") + 100 Then SendKeys ("7")
            If wParam = Asc("M") + 100 Then SendKeys ("L")
            If wParam = 201 + 100 Then SendKeys ("6")
            If wParam = 202 + 100 Then SendKeys ("U")

            ' уффф
            pHcSubclass = 0
        End If
        If GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow, vbNull)) = 67699721 Then ' Если английская раскладка
            ' Дальше очень нудно
                        If wParam = Asc("Q") Then SendKeys ("0,")
            If wParam = Asc("W") Then SendKeys ("VV")
            If wParam = Asc("E") Then SendKeys ("3")
            If wParam = Asc("R") Then SendKeys ("R")
            If wParam = Asc("T") Then SendKeys ("7")
            If wParam = Asc("Y") Then SendKeys ("Y")
            If wParam = Asc("U") Then SendKeys ("|_|")
            If wParam = Asc("I") Then SendKeys ("!")
            If wParam = Asc("O") Then SendKeys ("0")
            If wParam = Asc("P") Then SendKeys ("P")
            If wParam = 203 Then SendKeys ("[")
            If wParam = 204 Then SendKeys ("]")
            If wParam = Asc("A") Then SendKeys ("4")
            If wParam = Asc("S") Then SendKeys ("5")
            If wParam = Asc("D") Then SendKeys ("D")
            If wParam = Asc("F") Then SendKeys ("ph")
            If wParam = Asc("G") Then SendKeys ("6")
            If wParam = Asc("H") Then SendKeys ("|-|")
            If wParam = Asc("J") Then SendKeys ("J")
            If wParam = Asc("K") Then SendKeys (" |<")
            If wParam = Asc("L") Then SendKeys ("|_")
            If wParam = 205 Then SendKeys (";")
            If wParam = 206 Then SendKeys ("'")
            If wParam = Asc("Z") Then SendKeys ("2")
            If wParam = Asc("X") Then SendKeys ("X")
            If wParam = Asc("C") Then SendKeys ("€")
            If wParam = Asc("V") Then SendKeys ("\/")
            If wParam = Asc("B") Then SendKeys ("8")
            If wParam = Asc("N") Then SendKeys ("|\|")
            If wParam = Asc("M") Then SendKeys ("M")
            If wParam = 201 Then SendKeys (",")
            If wParam = 202 Then SendKeys (".")

            If wParam = Asc("Q") + 100 Then SendKeys ("0,")
            If wParam = Asc("W") + 100 Then SendKeys ("VV")
            If wParam = Asc("E") + 100 Then SendKeys ("3")
            If wParam = Asc("R") + 100 Then SendKeys ("R")
            If wParam = Asc("T") + 100 Then SendKeys ("7")
            If wParam = Asc("Y") + 100 Then SendKeys ("Y")
            If wParam = Asc("U") + 100 Then SendKeys ("|_|")
            If wParam = Asc("I") + 100 Then SendKeys ("!")
            If wParam = Asc("O") + 100 Then SendKeys ("0")
            If wParam = Asc("P") + 100 Then SendKeys ("P")
            If wParam = 203 + 100 Then SendKeys ("[")
            If wParam = 204 + 100 Then SendKeys ("]")
            If wParam = Asc("A") + 100 Then SendKeys ("4")
            If wParam = Asc("S") + 100 Then SendKeys ("5")
            If wParam = Asc("D") + 100 Then SendKeys ("D")
            If wParam = Asc("F") + 100 Then SendKeys ("ph")
            If wParam = Asc("G") + 100 Then SendKeys ("6")
            If wParam = Asc("H") + 100 Then SendKeys ("|-|")
            If wParam = Asc("J") + 100 Then SendKeys ("J")
            If wParam = Asc("K") + 100 Then SendKeys (" |<")
            If wParam = Asc("L") + 100 Then SendKeys ("|_")
            If wParam = 205 + 100 Then SendKeys (";")
            If wParam = 206 + 100 Then SendKeys ("'")
            If wParam = Asc("Z") + 100 Then SendKeys ("2")
            If wParam = Asc("X") + 100 Then SendKeys ("X")
            If wParam = Asc("C") + 100 Then SendKeys ("€")
            If wParam = Asc("V") + 100 Then SendKeys ("\/")
            If wParam = Asc("B") + 100 Then SendKeys ("8")
            If wParam = Asc("N") + 100 Then SendKeys ("|\|")
            If wParam = Asc("M") + 100 Then SendKeys ("M")
            If wParam = 201 + 100 Then SendKeys (",")
            If wParam = 202 + 100 Then SendKeys (".")

            ' уфф
            pHcSubclass = 0
        End If
    Else
        pHcSubclass = CallWindowProc(lOrigin, hwnd, uMsg, wParam, lParam)
    End If
    
End Function

Public Function pHcUnsubclass(ByVal hwnd As Long)
    SetWindowLong hwnd, GWL_WNDPROC, lOrigin
End Function

Public Function loadConf()
On Error GoTo err
options_remind = True
Open "SLconfig.ini" For Input As #1
Line Input #1, tmp
If Left(tmp, 6) = "remind" Then
    options_remind = Right(tmp, 1)
End If
Close #1
If Not options_remind Then Form1.cTray.DisplayBalloon "StrangeLeet", "The reminder is turned off.", NIIF_WARNING ' + NIIF_NOSOUND

err:
    If err.Number = 53 Then
        Close #1
        saveConf
        loadConf
    ElseIf err.Number = 13 Or err.Number = 62 Then
        Close #1
        MsgBox "Error. Please,wait!", vbCritical
        Kill "SLconfig.ini"
        loadConf
    End If
End Function

Public Function saveConf()
Open "SLconfig.ini" For Output As #1
Print #1, "remind " & CInt(options_remind) ^ 2
Close #1
End Function

Public Function finish()
Unload Form3
Unload Form2
Unload Form1
End Function

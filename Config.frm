VERSION 5.00
Begin VB.Form Config 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buttons Configure"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2340
      TabIndex        =   3
      Top             =   2565
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3915
      TabIndex        =   2
      Top             =   2565
      Width           =   1485
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   390
      Top             =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "??"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1860
      TabIndex        =   1
      Top             =   1350
      Width           =   270
   End
   Begin VB.Label Label0 
      AutoSize        =   -1  'True
      Caption         =   "Which key do you want to be match to keyboard:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   375
      TabIndex        =   0
      Top             =   315
      Width           =   5235
   End
End
Attribute VB_Name = "Config"
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


Dim inf As JOYINFO
Dim pubjoy As Boolean
Dim StartBTNcount As Byte


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
While inf.x <> NoxAxis Or inf.y <> NoyAxis
joyGetPos 0, inf
Wend
End Sub

Sub delaybtn()
While inf.btn <> NoKeys
joyGetPos 0, inf
Wend
End Sub



Private Sub Command1_Click()
Unload Config
End Sub

Private Sub Command2_Click()

Dim iFileNo As Integer
iFileNo = FreeFile
'open the file for writing
Open (App.Path + "\KeyData.txt") For Output As #iFileNo


Print #iFileNo, btn1
Print #iFileNo, btn2
Print #iFileNo, btn3
Print #iFileNo, btn4
Print #iFileNo, L1
Print #iFileNo, L2
Print #iFileNo, R1
Print #iFileNo, R2
Print #iFileNo, start
Print #iFileNo, slc
Print #iFileNo, lleft
Print #iFileNo, uup
Print #iFileNo, ddown
Print #iFileNo, rright

Close #iFileNo

Unload Config

End Sub

Private Sub Form_Load()
joyGetPos 0, inf
NoKeys = inf.btn
NoxAxis = inf.x
NoyAxis = inf.y
pubjoy = True
StartBTNcount = 0
End Sub


Private Sub Timer_Timer()
joyGetPos 0, inf
'Config.Caption = (Str(inf.x) + "/" + Str(inf.y))
'SendKeys ("{UP}")


If StartBTNcount = 0 Then
Label1.Caption = "Backspace"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 1 And inf.btn <> NoKeys Then
btn1 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 2 Then
Label1.Caption = "Alt+F4"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 3 And inf.btn <> NoKeys Then
btn2 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 4 Then
Label1.Caption = "Enter"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 5 And inf.btn <> NoKeys Then
btn3 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 6 Then
Label1.Caption = "End process list"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 7 And inf.btn <> NoKeys Then
btn4 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 8 Then
Label1.Caption = "Mouse upper left"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 9 And inf.btn <> NoKeys Then
L1 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 10 Then
Label1.Caption = "Function menu"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 11 And inf.btn <> NoKeys Then
L2 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 12 Then
Label1.Caption = "Center screen click"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 13 And inf.btn <> NoKeys Then
R1 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 14 Then
Label1.Caption = "Escape"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 15 And inf.btn <> NoKeys Then
R2 = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 16 Then
Label1.Caption = "Active & Deactive"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 17 And inf.btn <> NoKeys Then
start = inf.btn
delaybtn
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 18 Then
Label1.Caption = "Tab"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 19 And inf.btn <> NoKeys Then
slc = inf.btn
StartBTNcount = StartBTNcount + 1
delaybtn

ElseIf StartBTNcount = 20 Then
Label1.Caption = "Left arrow"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 21 And inf.x <> NoxAxis Then
lleft = inf.x
delay
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 22 Then
Label1.Caption = "Up arrow"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 23 And inf.y <> NoyAxis Then
uup = inf.y
delay
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 24 Then
Label1.Caption = "Down arrow"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 25 And inf.y <> NoyAxis Then
ddown = inf.y
delay
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 26 Then
Label1.Caption = "Right arrow"
StartBTNcount = StartBTNcount + 1
ElseIf StartBTNcount = 27 And inf.x <> NoxAxis Then
rright = inf.x
delay
StartBTNcount = StartBTNcount + 1
Command2.Enabled = True
Label1.Caption = ""

End If

End Sub











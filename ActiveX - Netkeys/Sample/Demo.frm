VERSION 5.00
Begin VB.Form Demo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sample"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2985
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Demo 
      Caption         =   "R E S E T"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1425
      Width           =   2970
   End
   Begin VB.CommandButton Demo 
      Caption         =   "Start Client ( Connect )"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   2970
   End
   Begin VB.CommandButton Demo 
      Caption         =   "Start Server ( Listen )"
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   735
      Width           =   2970
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Written by"
      Height          =   210
      Left            =   1140
      TabIndex        =   5
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label cptAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Michael A. Schmidt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   10
      Left            =   780
      MouseIcon       =   "Demo.frx":058A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2130
      Width           =   1410
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000001&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   15
      Top             =   1800
      Width           =   2955
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Sample"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   810
      TabIndex        =   3
      Top             =   210
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "Demo.frx":0894
      Top             =   75
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000001&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   720
      Left            =   15
      Top             =   0
      Width           =   2955
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cptAbout_Click(Index As Integer)
Shell "rundll32.exe url.dll,FileProtocolHandler mailto:mschmidt@mtdmarketing.com", vbMaximizedFocus
End Sub

Private Sub Demo_Click(Index As Integer)
    Select Case Index
    Case 0: Demo(1).Enabled = True
            Demo(0).Enabled = False
            GoListen
    Case 1: Demo(1).Enabled = False
            Demo(2).Enabled = True
            GoConnect
    Case 2: Demo(0).Enabled = True
            Demo(2).Enabled = False
            UnloadAll
    End Select
End Sub


Private Sub GoListen()

    Unload frmKeysOut
    frmKeysOut.KeysOut1.LocalPort = 10083
    frmKeysOut.KeysOut1.Listen
    frmKeysOut.Show

End Sub


Private Sub GoConnect()

    Unload frmKeysIn
    frmKeysIn.KeysIn1.RemotePort = 10083
    frmKeysIn.KeysIn1.RemoteIP = "127.0.0.1"
    frmKeysIn.KeysIn1.Connect
    frmKeysIn.Show

End Sub


Private Sub UnloadAll()
    Unload frmKeysIn
    Unload frmKeysOut
End Sub

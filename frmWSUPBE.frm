VERSION 5.00
Begin VB.Form frmExample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Start-Up Progress Bar Example"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   2400
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.OptionButton optDirection 
      Caption         =   "Right"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton optDirection 
      Caption         =   "Left"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.PictureBox picGrad 
      AutoSize        =   -1  'True
      Height          =   285
      Left            =   195
      Picture         =   "frmWSUPBE.frx":0000
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   347
      TabIndex        =   1
      Top             =   2400
      Width           =   5265
   End
   Begin VB.TextBox txtExplanation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmWSUPBE.frx":3D6E
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label lblIncrement 
      Alignment       =   1  'Right Justify
      Caption         =   "Pixel Increment Value:"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private picBar As clsPicBar

Private Sub cmdSet_Click()
   picBar.Incrementor = CInt(Text1.Text)
End Sub

Private Sub cmdStart_Click()
   Timer1.Enabled = True
   cmdStart.Enabled = False
   cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
   Timer1.Enabled = False
   cmdStart.Enabled = True
   cmdStop.Enabled = False
End Sub

Private Sub Form_Load()
   Set picBar = New clsPicBar
   With picBar
      Set .PicBox = picGrad
      .Incrementor = 5
   End With
   Text1.Text = "5"
   optDirection(0).Value = False
   optDirection(1).Value = True
   Timer1.Interval = 10
End Sub

Private Sub optDirection_Click(Index As Integer)
   If Index = 0 Then
      picBar.PicDir = GOLEFT
   Else
      picBar.PicDir = GORIGHT
   End If
End Sub

Private Sub Timer1_Timer()
   DoEvents
   picBar.BGScroll
End Sub

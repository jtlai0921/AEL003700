VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.Image Image5 
      Height          =   735
      Left            =   1800
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   1215
      Left            =   1560
      Picture         =   "Form1.frx":A822
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   1815
      Left            =   1200
      Picture         =   "Form1.frx":15044
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   2295
      Left            =   840
      Picture         =   "Form1.frx":1F866
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   480
      Picture         =   "Form1.frx":2A088
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
   If Image1.Visible Then
      Image1.Visible = False
      Image2.Visible = True
   ElseIf Image2.Visible Then
      Image2.Visible = False
      Image3.Visible = True
   ElseIf Image3.Visible Then
      Image3.Visible = False
      Image4.Visible = True
   ElseIf Image4.Visible Then
      Image4.Visible = False
      Image5.Visible = True
   Else
      Image5.Visible = False
      Image1.Visible = True
   End If
End Sub

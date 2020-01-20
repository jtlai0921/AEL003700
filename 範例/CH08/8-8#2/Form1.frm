VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   600
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Timer1.Interval = 100
End Sub

Private Sub Timer1_Timer()
   Image1.Visible = False
   gap = 20
   If Image1.Width >= 1000 Then
      Image1.Top = Image1.Top + gap
      Image1.Left = Image1.Left + gap
      Image1.Width = Image1.Width - gap * 2
      Image1.Height = Image1.Height - gap * 2
   End If
   Image1.Visible = True
End Sub

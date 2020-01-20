VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   2775
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Private Sub Timer1_Timer()
   Form1.BackColor = QBColor(n)
   n = n + 1
   If n > 15 Then n = 0
End Sub

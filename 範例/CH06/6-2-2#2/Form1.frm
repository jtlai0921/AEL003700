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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   a = 10
   b = 30
   c = a + b
   d = a - b
   Print "1234567890" + "1234567890"; "123456"
   Print "A+B="; c, "A-B="; d
   Print a; "+"; b; "="; a + b
   Print a; "-"; b; "="; a - b
End Sub

